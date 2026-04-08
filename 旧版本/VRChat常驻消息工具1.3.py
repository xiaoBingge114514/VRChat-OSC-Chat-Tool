import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from pythonosc import udp_client
from datetime import datetime, timedelta
import win32gui
import re
import asyncio
import time
import win32api
import psutil
import struct
import threading
from queue import Queue
import json
import os
import glob
import socket
import subprocess
import requests
import websockets
from websockets import protocol
from pydantic import BaseModel
import contextlib

from bleak import BleakScanner, BleakClient
from GPUtil import GPUtil 
from winsdk.windows.media.control import (
    GlobalSystemMediaTransportControlsSessionManager as MediaManager,
    GlobalSystemMediaTransportControlsSessionPlaybackStatus
)

HRS_UUID = "0000180d-0000-1000-8000-00805f9b34fb"
HRM_UUID = "00002a37-0000-1000-8000-00805f9b34fb"

# 配置文件路径
CONFIG_FILE = "vrchat_config.json"


class Config(BaseModel):
    osc_ip: str = "127.0.0.1"
    osc_port: int = 9000
    ncm_port: int = 9222
    ncm_path: str = ""
    refresh_interval: float = 3.0
    bar_width: int = 8
    bar_filled: str = "▓"
    bar_empty: str = "░"
    bar_thumb: str = "◘"
    template: str = "🎵 {song} - {artist}\n{bar} {time}\n{lyric1}\n{lyric2}"


class SongState:
    def __init__(
        self, song="", artist="", cur=0, dur=0, play=False, lyric1="", lyric2=""
    ):
        self.song = song
        self.artist = artist
        self.cur = cur
        self.dur = dur
        self.play = play
        self.lyric1 = lyric1
        self.lyric2 = lyric2

    def update(self, d: dict):
        for k in ["song", "artist", "cur", "dur", "play", "lyric1", "lyric2"]:
            if k in d:
                setattr(self, k, d[k])

    def copy(self):
        return SongState(
            song=self.song,
            artist=self.artist,
            cur=self.cur,
            dur=self.dur,
            play=self.play,
            lyric1=self.lyric1,
            lyric2=self.lyric2,
        )


class SharedState:
    def __init__(self):
        self.lock = threading.Lock()
        self.data = SongState()
        self.song_key = ""
        self.lyrics = []
        self.last_update = 0  # 上次获取 JS 数据的时间戳


def get_lyric(lyrics, pos):
    if not lyrics:
        return "", ""
    left, r, idx = 0, len(lyrics) - 1, -1
    while left <= r:
        m = (left + r) // 2
        if lyrics[m][0] <= pos:
            idx, left = m, m + 1
        else:
            r = m - 1
    if idx < 0:
        return lyrics[0][1], lyrics[1][1] if len(lyrics) > 1 else ""
    return lyrics[idx][1], lyrics[idx + 1][1] if idx + 1 < len(lyrics) else ""


def format_output(cfg: Config, state, lyrics, song_key, title_limit=20, artist_limit=25):
    c, d, w = state.cur, state.dur, cfg.bar_width
    pos = int(w * c / d) if d else 0
    thumb = cfg.bar_thumb
    if thumb:
        bar = cfg.bar_filled * pos + thumb + cfg.bar_empty * (w - pos)
    else:
        bar = cfg.bar_filled * pos + cfg.bar_empty * (w - pos)
    
    # 优先使用软件歌词
    l1, l2 = state.lyric1, state.lyric2
    if not l1 and lyrics and song_key == f"{state.song}-{state.artist}":
        l1, l2 = get_lyric(lyrics, c)
    l1, l2 = l1 or "纯音乐，请欣赏", l2 or ""
    
    # 应用字符限制
    limited_song = state.song[:title_limit]
    limited_artist = state.artist[:artist_limit]
    
    try:
        return cfg.template.format(
            song=limited_song,
            artist=limited_artist,
            bar=bar,
            time=f"{c // 60}:{c % 60:02d}/{d // 60}:{d % 60:02d}",
            lyric1=l1,
            lyric2=l2,
        )
    except Exception:
        return f"🎵 {limited_song} - {limited_artist}\n{bar}\n{l1}"


HEADERS = {"User-Agent": "Mozilla/5.0", "Referer": "https://music.163.com/"}


def fetch_lyrics(song, artist):
    with contextlib.suppress(Exception):
        r = requests.post(
            "https://music.163.com/api/search/get",
            data={"s": f"{song} {artist}", "type": 1, "limit": 1},
            headers=HEADERS,
            timeout=3,
        ).json()
        if r.get("result", {}).get("songs"):
            lrc = (
                requests.get(
                    f"https://music.163.com/api/song/lyric?id={r['result']['songs'][0]['id']}&lv=1",
                    headers=HEADERS,
                    timeout=3,
                )
                .json()
                .get("lrc", {})
                .get("lyric", "")
            )
            return sorted(
                [
                    (
                        int(m[1]) * 60
                        + int(m[2])
                        + float(m[3]) * (0.01 if len(m[3]) == 2 else 0.001),
                        m[4].strip(),
                    )
                    for m in re.finditer(r"\[(\d{2}):(\d{2})\.(\d{2,3})\](.*)", lrc)
                    if m[4].strip()
                ],
                key=lambda x: x[0],
            )
    return []


def find_netease() -> str | None:
    patterns = [
        r"%APPDATA%\Microsoft\Windows\Start Menu\Programs\**\*网易云*.lnk",
        r"%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\**\*网易云*.lnk",
    ]
    for pat_idx, pattern in enumerate(patterns):
        for lnk in glob.glob(os.path.expandvars(pattern), recursive=True):
            try:
                with open(lnk, "rb") as f:
                    m = re.search(
                        rb"([A-Za-z]:\\[^\x00]+?cloudmusic.exe)",
                        f.read(),
                        re.IGNORECASE,
                    )
                    if m:
                        p = m.group(1).decode("utf-8", errors="ignore")
                        if os.path.exists(p):
                            return p
            except Exception:
                continue
    return None


def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def launch_netease(port=None, path=None):
    exe = path if path and os.path.exists(path) else find_netease()
    if not exe:
        return False, "未找到网易云", None
    if port is None:
        port = find_free_port()
    proc = subprocess.Popen([
        exe,
        f"--remote-debugging-port={port}",
        "--disable-background-timer-throttling",
        "--disable-backgrounding-occluded-windows",
        "--disable-renderer-backgrounding",
    ])

    return True, exe, port


# 兼容 VIP 界面
JS_GET_STATE = r"""(() => {
    try {
        let r = { song: '', artist: '', cur: 0, dur: 0, play: false, lyric1: '', lyric2: '' };

        // 获取歌曲名
        let songEl = document.querySelector('.cmd-space.title span') 
            || document.querySelector('.main-title')
            || document.querySelector('.two-line .title')
            || document.querySelector('[class*="title"] span');
        r.song = songEl?.innerText?.trim() || songEl?.textContent?.trim() || '';

        // .author, .info.artist
        let artist = document.querySelector('.author');
        r.artist = artist?.innerText?.trim() || '';
        if (!r.artist) { artist = document.querySelector('.info.artist'); r.artist = (artist?.innerText || '').replace(/^歌手[：:]/, '').trim(); }

        // 进度遍历
        let walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
        while (walker.nextNode()) {
            let text = walker.currentNode.textContent.trim();
            let m = text.match(/^(\d+):(\d+)\s*\/\s*(\d+):(\d+)$/);
            if (m) {
                r.cur = +m[1] * 60 + +m[2];
                r.dur = +m[3] * 60 + +m[4];
                break;
            }
        }
        
        // 备用 .curtime-thumb
        if (!r.dur) {
            let timeEl = document.querySelector('.curtime-thumb');
            if (timeEl?.innerText) { 
                let m = timeEl.innerText.match(/(\d+):(\d+)\s*\/\s*(\d+):(\d+)/); 
                if (m) { r.cur = +m[1] * 60 + +m[2]; r.dur = +m[3] * 60 + +m[4]; } 
            }
        }

        // cmd-icon-pause, title
        r.play = !!document.querySelector('[class*="cmd-icon-pause"]') || !!document.querySelector('[title*="暂停（Ctrl"]');

        // .line.current
        let curLine = document.querySelector('.line.current');
        if (curLine) {
            r.lyric1 = curLine.innerText?.trim() || '';
            let next = curLine.nextElementSibling;
            if (next && next.classList?.contains('line')) {
                r.lyric2 = next.innerText?.trim() || '';
            }
        }

        return r;
    } catch (e) { return null; }
})()"""


class CallbackProtocol:
    def cb_status(self, t: str) -> None: ...
    def cb_song(self, t: str) -> None: ...
    def cb_output(self, t: str) -> None: ...


def netease_thread(
    cfg: Config, shared: SharedState, stop_event: threading.Event, cb: CallbackProtocol
):
    class NeteaseSync:
        def __init__(self):
            self.ws = None
            self.msg_id = 0

        async def connect(self):
            pages = requests.get(
                f"http://127.0.0.1:{cfg.ncm_port}/json", timeout=2
            ).json()
            self.ws = await websockets.connect(
                pages[0]["webSocketDebuggerUrl"], ping_interval=30, ping_timeout=15
            )

        async def eval_js(self, code, timeout=1):
            if not self.ws:
                return None
            
            # 激活页面
            self.msg_id += 1
            bring_id = self.msg_id
            await self.ws.send(json.dumps({"id": bring_id, "method": "Page.bringToFront"}))
            
            # 等待 bringToFront 响应
            try:
                while True:
                    msg = await asyncio.wait_for(self.ws.recv(), timeout=0.5)
                    d = json.loads(msg)
                    if d.get("id") == bring_id:
                        break
            except asyncio.TimeoutError:
                pass
            
            # 执行主代码
            self.msg_id += 1
            msg_id = self.msg_id
            await self.ws.send(json.dumps({
                "id": msg_id, "method": "Runtime.evaluate",
                "params": {"expression": code, "returnByValue": True},
            }))
            
            result = None
            try:
                end_time = asyncio.get_event_loop().time() + timeout
                while True:
                    remaining = end_time - asyncio.get_event_loop().time()
                    if remaining <= 0:
                        break
                    msg = await asyncio.wait_for(self.ws.recv(), timeout=remaining)
                    d = json.loads(msg)
                    if d.get("id") == msg_id:
                        result = d.get("result", {}).get("result", {}).get("value")
                        break
            except asyncio.TimeoutError:
                pass
            except Exception:
                pass
            
            return result

    async def run():
        cb.cb_status("连接中...")
        sync = NeteaseSync()
        retry_count = 0
        max_retries = 3
        while not sync.ws and retry_count < max_retries:
            try:
                await sync.connect()
            except Exception as e:
                retry_count += 1
                cb.cb_status(f"连接失败，重试中... {retry_count}/{max_retries} \n {e}")
                await asyncio.sleep(2)
        if not sync.ws:
            cb.cb_status("连接失败，已达最大重试次数")
            return
        cb.cb_status("已连接")
        
        timeout_count = 0
        
        while not stop_event.is_set():
            try:
                s = await sync.eval_js(JS_GET_STATE)
                
                if s is None:
                    timeout_count += 1
                    if timeout_count >= 3:
                        cb.cb_status("响应超时，重连中...")
                        try:
                            await sync.connect()
                            cb.cb_status("已重连")
                            timeout_count = 0
                        except Exception:
                            pass
                    await asyncio.sleep(0.5)
                    continue
                
                timeout_count = 0
                
                if s.get("song"):
                    need_fetch = False
                    song_to_fetch = ""
                    artist_to_fetch = ""
                    key = ""
                    
                    with shared.lock:
                        # 切歌过渡期
                        if s.get("cur", 0) == 0 and s.get("dur", 0) == 0:
                            if shared.data.dur > 0:
                                if s.get("song") != shared.data.song:
                                    s["cur"] = 0
                                    s["dur"] = shared.data.dur
                                else:
                                    s["cur"] = shared.data.cur
                                    s["dur"] = shared.data.dur
                        
                        shared.data.update(s)
                        shared.last_update = time.time()
                        
                        if shared.data.song:
                            key = f"{shared.data.song}-{shared.data.artist}"
                            if key != shared.song_key:
                                shared.song_key = key
                                shared.lyrics = []
                                need_fetch = True
                                song_to_fetch = shared.data.song
                                artist_to_fetch = shared.data.artist
                    
                    # 启动歌词获取线程
                    if need_fetch:
                        def fetch_task(k, song, artist):
                            try:
                                new_lyrics = fetch_lyrics(song, artist)
                                with shared.lock:
                                    if shared.song_key == k:
                                        shared.lyrics = new_lyrics
                            except Exception:
                                pass
                        threading.Thread(
                            target=fetch_task,
                            args=(key, song_to_fetch, artist_to_fetch),
                            daemon=True
                        ).start()
                
                await asyncio.sleep(0.3)
            except websockets.exceptions.ConnectionClosed:
                await asyncio.sleep(1)
                try:
                    await sync.connect()
                    cb.cb_status("已重连")
                except Exception:
                    cb.cb_status("重连失败，继续尝试...")
            except Exception:
                pass
        
        if sync.ws and sync.ws.state == protocol.State.OPEN:
            await sync.ws.close()

    asyncio.run(run())


def osc_thread(
    cfg: Config, shared: SharedState, stop_event: threading.Event, cb: CallbackProtocol
):
    osc = None
    last_osc = 0
    while not stop_event.is_set():
        with shared.lock:
            state = shared.data.copy()
            lyrics = list(shared.lyrics)
            song_key = shared.song_key
            last_update = shared.last_update
        
        # 推算进度
        if state.play and state.song and last_update > 0:
            elapsed = time.time() - last_update
            state.cur = min(int(state.cur + elapsed), state.dur) if state.dur else state.cur
        
        if state.play and state.song:
            out = format_output(cfg, state, lyrics, song_key)
            now = time.time()
            if now - last_osc >= cfg.refresh_interval:
                if osc is None:
                    osc = udp_client.SimpleUDPClient(cfg.osc_ip, cfg.osc_port)
                osc.send_message("/chatbox/input", [out, True, False])
                last_osc = now
                cb.cb_output(out)
                cb.cb_song(f"播放：{state.song} - {state.artist}")
        else:
            if state.song:
                cb.cb_song(f"暂停：{state.song}")
        time.sleep(0.3)


class HeartRateMonitor:
    """心率监测器类 - 独立运行不阻塞GUI"""
    def __init__(self):
        self.current_hr = 0
        self.is_connected = False
        self.device_name = None
        self._running = False
        self._thread = None
        self._queue = Queue()
        
    def start(self):
        """启动心率监测线程"""
        if not self._running:
            self._running = True
            self._thread = threading.Thread(target=self._run_async_loop, daemon=True)
            self._thread.start()
            
    def stop(self):
        """停止心率监测"""
        self._running = False
        self.is_connected = False
        self.current_hr = 0
        
    def _run_async_loop(self):
        """在新线程中运行异步事件循环"""
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(self._main_loop())
        
    async def _main_loop(self):
        """主监测循环"""
        while self._running:
            try:
                await self._scan_and_connect()
            except Exception as e:
                print(f"心率监测错误: {e}")
                await asyncio.sleep(3)
                
    async def _scan_and_connect(self):
        """扫描并连接设备"""
        self.is_connected = False
        self.device_name = None
        
        device = await BleakScanner.find_device_by_filter(
            lambda d, ad: ad.service_uuids and HRS_UUID.lower() in [u.lower() for u in ad.service_uuids],
            timeout=10.0
        )
        
        if not device or not self._running:
            await asyncio.sleep(3)
            return
            
        self.device_name = device.name or "Unknown"
        
        try:
            async with BleakClient(device) as client:
                self.is_connected = True
                
                hrs = next((s for s in client.services if s.uuid.lower() == HRS_UUID), None)
                if not hrs:
                    raise RuntimeError("无心率服务")
                
                hrm = next((c for c in hrs.characteristics if c.uuid.lower() == HRM_UUID), None)
                if not hrm:
                    raise RuntimeError("无心率特征")
                
                def on_notify(sender, data):
                    self.current_hr = self._parse_heart_rate(data)
                
                await client.start_notify(hrm, on_notify)
                
                while client.is_connected and self._running:
                    await asyncio.sleep(2)
                    
        except Exception as e:
            print(f"心率设备连接失败: {e}")
            self.is_connected = False
            await asyncio.sleep(3)
            
    @staticmethod
    def _parse_heart_rate(data: bytearray) -> int:
        """解析心率值"""
        flags = data[0]
        if flags & 0x01:
            return struct.unpack_from('<H', data, 1)[0]
        return data[1]


class VRChatAutoChat:
    def __init__(self, root):
        self.root = root 
        self.osc_ip = tk.StringVar(value="127.0.0.1")
        self.osc_port = tk.IntVar(value=9000)
        self.osc_client = udp_client.SimpleUDPClient(self.osc_ip.get(), self.osc_port.get())
        self.is_sending = False
        self.scheduled_event = None
        self.history_max_items = 20
        self.max_message_length = 144
        self.history_list = []
        self.root.minsize(900, 625)
        
        self.auto_time = tk.BooleanVar(value=False)
        self.auto_window = tk.BooleanVar(value=False)
        self.auto_wrap = tk.BooleanVar(value=False)
        self.auto_music = tk.BooleanVar(value=False)
        self.auto_idle = tk.BooleanVar(value=False)
        self.idle_threshold = tk.IntVar(value=30) 

        self.auto_hardware = tk.BooleanVar(value=False) 
        self.auto_cpu = tk.BooleanVar(value=False)
        self.auto_ram = tk.BooleanVar(value=False) 
        self.auto_gpu = tk.BooleanVar(value=False) 
        
        # 心率检测相关变量
        self.auto_heart_rate = tk.BooleanVar(value=False)
        self.heart_rate_monitor = HeartRateMonitor()
        
        # 高级音乐信息相关变量
        self.advanced_music_enabled = tk.BooleanVar(value=False)
        self.ncm_path = tk.StringVar(value="")
        self.ncm_port = tk.IntVar(value=9222)
        self.ncm_sync_running = False
        self.ncm_sync_thread = None
        self.ncm_stop_event = None
        self.ncm_shared_state = SharedState()
        self.ncm_config = Config(ncm_path="", ncm_port=9222)
        self.ncm_launch_thread = None
        self.ncm_launch_event = None

        # 模板字符串模式相关变量
        self.use_template_mode = tk.BooleanVar(value=False)
        self.template_string = tk.StringVar(value="{time}❀{message}❀{window}\n{music}{hardware}{heart_rate}{idle}")

        self.debug_update_interval = 1000
        self.debug_update_job = None
        self.debug_labels = {}
        self.last_send_time = 0

        self.original_wrap_state = False

        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)

        self.order_options = ["时间", "消息内容", "挂机状态", "窗口标题", "心率", "硬件监测" , "音乐信息"]
        
        self.cpu_custom_label = tk.StringVar(value="")
        self.ram_custom_label = tk.StringVar(value="")
        self.gpu_custom_label = tk.StringVar(value="")
        
        # 字符数限制
        self.window_title_limit = tk.IntVar(value=15)
        self.music_title_limit = tk.IntVar(value=20)
        self.music_artist_limit = tk.IntVar(value=25)

        # 初始化顺序变量，从配置加载或使用默认顺序
        self.load_config()

        self.root.title("VRChat常驻消息工具")
        self.root.geometry("850x600")
        self.create_widgets()
        self.create_watermark()
        # 修改居中逻辑
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (800 // 2)
        y = (self.root.winfo_screenheight() // 2) - (600 // 2)
        self.root.geometry(f"800x600+{x}+{y}")
        self.root.protocol('WM_DELETE_WINDOW', self.on_close)
        self.update_status()
        
        # 监听消息内容变化，控制开始按钮状态
        self.text_input.bind("<KeyRelease>", self.check_start_button_state)
        self.text_input.bind("<<Paste>>", self.check_start_button_state)

        # 添加对窗口焦点变化的监听
        self.root.bind("<FocusIn>", self.on_focus_in)
        self.root.bind("<FocusOut>", self.on_focus_out)

    def on_focus_in(self, event=None):
        """当主窗口获得焦点时的处理"""
        pass

    def on_focus_out(self, event=None):
        """当主窗口失去焦点时的处理"""
        pass

    def check_start_button_state(self, event=None):
        """检查开始按钮状态"""
        message_content = self.text_input.get("1.0", "end-1c").strip()
        has_message = bool(message_content)
        
        # 检查是否有启用的功能
        has_enabled_feature = (
            self.auto_time.get() or
            self.auto_window.get() or
            self.auto_music.get() or
            self.auto_idle.get() or
            self.auto_hardware.get() or
            self.auto_heart_rate.get()
        )
        
        # 如果有消息内容或启用了功能，则按钮可用
        if has_message or has_enabled_feature:
            self.start_btn.config(state=tk.NORMAL)
        else:
            self.start_btn.config(state=tk.DISABLED)

    def load_config(self):
        """加载配置文件"""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # 加载排序设置
                saved_order = config.get('order', [])
                self.order_vars = {option: tk.StringVar(value=str(idx)) for idx, option in enumerate(self.order_options)}
                
                # 如果保存了排序，则应用它
                if saved_order:
                    for i, option in enumerate(saved_order):
                        if option in self.order_options:
                            self.order_vars[option].set(str(i))
                
                # 加载其他设置
                self.cpu_custom_label.set(config.get('cpu_custom_label', ''))
                self.ram_custom_label.set(config.get('ram_custom_label', ''))
                self.gpu_custom_label.set(config.get('gpu_custom_label', ''))
                self.window_title_limit.set(config.get('window_title_limit', 20))
                self.music_title_limit.set(config.get('music_title_limit', 30))
                self.music_artist_limit.set(config.get('music_artist_limit', 30))
                self.osc_ip.set(config.get('osc_ip', '127.0.0.1'))
                self.osc_port.set(config.get('osc_port', 9000))
                
                # 加载自动附加项状态
                self.auto_time.set(config.get('auto_time', False))
                self.auto_window.set(config.get('auto_window', False))
                self.auto_wrap.set(config.get('auto_wrap', False))
                self.auto_music.set(config.get('auto_music', False))
                self.auto_idle.set(config.get('auto_idle', False))
                self.auto_hardware.set(config.get('auto_hardware', False))
                self.auto_cpu.set(config.get('auto_cpu', False))
                self.auto_ram.set(config.get('auto_ram', False))
                self.auto_gpu.set(config.get('auto_gpu', False))
                self.auto_heart_rate.set(config.get('auto_heart_rate', False))
                self.idle_threshold.set(config.get('idle_threshold', 30))
                
                # 加载高级音乐信息设置
                self.advanced_music_enabled.set(config.get('advanced_music_enabled', False))
                self.ncm_path.set(config.get('ncm_path', ''))
                self.ncm_port.set(config.get('ncm_port', 9222))
                self.ncm_config.bar_width = config.get('ncm_bar_width', 9)
                self.ncm_config.bar_filled = config.get('ncm_bar_filled', '▓')
                self.ncm_config.bar_thumb = config.get('ncm_bar_thumb', '◘')
                self.ncm_config.bar_empty = config.get('ncm_bar_empty', '░')
                self.ncm_config.template = config.get('ncm_template', '🎵 {song} - {artist}\n{bar} {time}\n{lyric1}\n{lyric2}')
                
                # 加载模板模式相关设置
                self.use_template_mode.set(config.get('use_template_mode', False))
                self.template_string.set(config.get('template_string', '{message}{time}{window}{idle}{music}{hardware}{heart_rate}'))
                
            else:
                # 默认配置
                self.order_vars = {option: tk.StringVar(value=str(idx)) for idx, option in enumerate(self.order_options)}
        except Exception as e:
            print(f"加载配置失败: {e}")
            # 使用默认值
            self.order_vars = {option: tk.StringVar(value=str(idx)) for idx, option in enumerate(self.order_options)}

    def save_config(self):
        """保存配置到文件"""
        try:
            # 构建配置字典，第一个键专门用于声明文件用途
            config = {
                "_file_description": "VRChat 常驻消息工具配置文件 (v1.3) - 请勿手动修改此处，以免格式错误",
                "_version": "1.3", 
                "order": self.get_current_order(),
                'cpu_custom_label': self.cpu_custom_label.get(),
                'ram_custom_label': self.ram_custom_label.get(),
                'gpu_custom_label': self.gpu_custom_label.get(),
                'window_title_limit': self.window_title_limit.get(),
                'music_title_limit': self.music_title_limit.get(),
                'music_artist_limit': self.music_artist_limit.get(),
                'osc_ip': self.osc_ip.get(),
                'osc_port': self.osc_port.get(),
                'auto_time': self.auto_time.get(),
                'auto_window': self.auto_window.get(),
                'auto_wrap': self.auto_wrap.get(),
                'auto_music': self.auto_music.get(),
                'auto_idle': self.auto_idle.get(),
                'auto_hardware': self.auto_hardware.get(),
                'auto_cpu': self.auto_cpu.get(),
                'auto_ram': self.auto_ram.get(),
                'auto_gpu': self.auto_gpu.get(),
                'auto_heart_rate': self.auto_heart_rate.get(),
                'idle_threshold': self.idle_threshold.get(),
                'advanced_music_enabled': self.advanced_music_enabled.get(),
                'ncm_path': self.ncm_path.get(),
                'ncm_port': self.ncm_port.get(),
                'ncm_bar_width': self.ncm_config.bar_width,
                'ncm_bar_filled': self.ncm_config.bar_filled,
                'ncm_bar_thumb': self.ncm_config.bar_thumb,
                'ncm_bar_empty': self.ncm_config.bar_empty,
                'ncm_template': self.ncm_config.template,
                'use_template_mode': self.use_template_mode.get(),
                'template_string': self.template_string.get()
            }

            with open('vrchat_config.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            
            print("配置已保存")
        except Exception as e:
            print(f"保存配置失败: {e}")
            messagebox.showerror("错误", f"无法保存配置文件:\n{e}")

    def get_current_order(self):
        sorted_items = sorted(self.order_vars.items(), key=lambda x: int(x[1].get()))
        return [item[0] for item in sorted_items]

    def create_watermark(self):
        bg_color = self.root.cget('bg')
        self.watermark = ttk.Label(
            self.root,
            text="From B_小槟",
            font=('微软雅黑', 10),
            foreground='#AAAAAA',
            background=bg_color,
            cursor="hand2"
        )
        self.watermark.place(relx=1.0, rely=1.0, anchor='se', x=-10, y=-5)
        self.watermark.bind("<Button-1>", self.show_about_window)

    def show_about_window(self, event=None):
        about_win = tk.Toplevel(self.root)
        about_win.title("关于")
        about_win.resizable(False, False)
        # 移除 grab_set() 以允许主窗口接收焦点
        about_win.transient(self.root)
        # 移除 grab_set()

        content = [
            ("VRChat常驻消息工具", 14, "bold"),
            ("版本：1.3", 12),
            ("作者：(VRC)B_小槟", 12),
            ("", 10),
            ("功能特性：", 12, "bold"),
            ("- OSC协议消息发送", 10),
            ("- 历史记录功能", 10),
            ("- 智能附加项（时间/窗口/音乐/心率）", 10),
            ("- 挂机检测机制", 10),
            ("- 硬件性能监测 (CPU/RAM/GPU)", 10),
            ("- 蓝牙心率监测（基于HRS协议）", 10),
            ("- 硬件监测标签自定义显示", 10),
            ("- 窗口标题/音乐信息字符数限制", 10),
            ("- 发送顺序自定义与配置持久保存", 10),
            ("- 模板字符串模式（高级）", 10),
            ("- 高级音乐信息（网易云歌词同步）", 10),
            ("", 10),
            ("提示：点击右下角水印显示本窗口", 9, "italic")
        ]
        for idx, item in enumerate(content):
            text = item[0]
            size = item[1]
            weight = "normal" if len(item) < 3 else item[2]
            fg = "#333333" if size > 10 else "#666666"
            
            label = ttk.Label(
                about_win,
                text=text,
                font=('微软雅黑', size, weight),
                foreground=fg,
                anchor="w"
            )
            label.pack(padx=20, pady=2 if size <12 else 5, fill=tk.X)

        about_win.update_idletasks()
        width = about_win.winfo_width()
        height = about_win.winfo_height()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (width // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (height // 2)
        about_win.geometry(f"+{x}+{y}")

        close_btn = ttk.Button(
            about_win,
            text="关闭",
            command=about_win.destroy
        )
        close_btn.pack(pady=10, padx=20, ipadx=20)

        about_win.bind("<Return>", lambda e: about_win.destroy())

    def create_widgets(self):
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=0)

        left_frame = ttk.Frame(self.root)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        right_frame = ttk.Frame(self.root)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)

        self.create_left_panel(left_frame)
        self.create_right_panel(right_frame)
        self.create_debug_panel(right_frame)

    def update_osc_client(self):
        ip = self.osc_ip.get()
        port = self.osc_port.get()
        self.osc_client = udp_client.SimpleUDPClient(ip, port)

    def create_left_panel(self, parent):
        input_frame = ttk.LabelFrame(parent, text="消息输入（最多144字符）")
        input_frame.pack(pady=5, padx=10, fill=tk.X, expand=True)

        self.chars_remaining_var = tk.StringVar()
        self.chars_remaining_label = ttk.Label(
            input_frame, 
            textvariable=self.chars_remaining_var,
            foreground="gray"
        )
        self.chars_remaining_label.pack(anchor=tk.E)

        self.text_input = scrolledtext.ScrolledText(
            input_frame,
            wrap=tk.WORD,
            height=7,
            font=('Microsoft YaHei', 11)
        )
        self.text_input.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)
        self.text_input.bind("<KeyRelease>", self.update_char_count)
        self.text_input.bind("<<Paste>>", self.on_paste)
        self.update_char_count()

        control_frame = ttk.Frame(parent)
        control_frame.pack(pady=5, fill=tk.X)

        ttk.Label(control_frame, text="间隔时间（秒）:").pack(side=tk.LEFT)
        self.interval_var = tk.IntVar(value=3)
        ttk.Spinbox(
            control_frame,
            from_=1,
            to=300,
            textvariable=self.interval_var,
            width=5
        ).pack(side=tk.LEFT, padx=5)

        self.start_btn = ttk.Button(
            control_frame,
            text="开始发送",
            command=self.toggle_sending
        )
        self.start_btn.pack(side=tk.LEFT, padx=5)
        # 初始状态下按钮不可用
        self.start_btn.config(state=tk.DISABLED)

        self.countdown_var = tk.StringVar()
        ttk.Label(
            control_frame,
            textvariable=self.countdown_var,
            foreground="#666666"
        ).pack(side=tk.LEFT, padx=5)

        self.status_var = tk.StringVar()
        ttk.Label(parent, textvariable=self.status_var, foreground="gray").pack()

        history_frame = ttk.LabelFrame(parent, text="发送历史（最近20条）")
        history_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
        
        self.history_text = tk.Text(
            history_frame,
            height=11,
            wrap=tk.WORD,
            font=('Arial', 10)
        )
        scrollbar = ttk.Scrollbar(history_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.history_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.history_text.yview)

    def create_right_panel(self, parent):
        tool_frame = ttk.LabelFrame(parent, text="自动附加项")
        tool_frame.pack(pady=10, padx=5, fill=tk.X)

        template_mode_frame = ttk.Frame(tool_frame)
        template_mode_frame.grid(row=0, column=0, sticky="w", pady=2)
        ttk.Checkbutton(
            template_mode_frame,
            text="启用模板字符串模式",
            variable=self.use_template_mode,
            command=self.update_template_mode  # 修复：添加回调函数
        ).pack(side=tk.LEFT)

        # 附加项自动换行复选框（可能被禁用）
        self.wrap_checkbox = ttk.Checkbutton(
            tool_frame,
            text="附加项自动换行",
            variable=self.auto_wrap,
            command=self.update_status
        )
        self.wrap_checkbox.grid(row=1, column=0, sticky="w", pady=2)

        ttk.Checkbutton(
            tool_frame,
            text="添加当前时间",
            variable=self.auto_time,
            command=self.update_status
        ).grid(row=2, column=0, sticky="w", pady=2)

        ttk.Checkbutton(
            tool_frame,
            text="添加窗口标题",
            variable=self.auto_window,
            command=self.update_status
        ).grid(row=3, column=0, sticky="w", pady=2)

        ttk.Checkbutton(
            tool_frame,
            text="添加音乐信息 (SMTC)",
            variable=self.auto_music,
            command=self.toggle_advanced_music_state
        ).grid(row=4, column=0, sticky="w", pady=2)

        # 高级音乐信息复选框
        self.advanced_music_checkbox = ttk.Checkbutton(
            tool_frame,
            text="启用高级音乐信息（网易云音乐）",
            variable=self.advanced_music_enabled,
            command=self.update_status
        )
        self.advanced_music_checkbox.grid(row=5, column=0, sticky="w", pady=2)
        
        # 根据是否启用普通音乐信息设置高级音乐信息的初始状态
        if not self.auto_music.get():
            self.advanced_music_checkbox.config(state=tk.DISABLED)

        idle_frame = ttk.Frame(tool_frame)
        idle_frame.grid(row=6, column=0, sticky="w", pady=2)
        
        ttk.Checkbutton(
            idle_frame,
            text="检测挂机",
            variable=self.auto_idle,
            command=self.update_status
        ).pack(side=tk.LEFT)
        
        ttk.Label(idle_frame, text="阈值（秒）:").pack(side=tk.LEFT, padx=5)
        spinbox = ttk.Spinbox(
            idle_frame,
            from_=5,
            to=21600,
            textvariable=self.idle_threshold,
            width=7
        )
        spinbox.pack(side=tk.LEFT)
        self.idle_threshold.trace_add("write", lambda *args: self.update_status())

        hardware_frame = ttk.Frame(tool_frame)
        hardware_frame.grid(row=7, column=0, sticky="w", pady=2)
        ttk.Checkbutton(
            hardware_frame,
            text="硬件监测",
            variable=self.auto_hardware,
            command=self.update_status
        ).pack(side=tk.LEFT)
        ttk.Checkbutton(
            hardware_frame,
            text="CPU",
            variable=self.auto_cpu,
            command=self.update_status
        ).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(
            hardware_frame,
            text="RAM",
            variable=self.auto_ram,
            command=self.update_status
        ).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(
            hardware_frame,
            text="GPU",
            variable=self.auto_gpu,
            command=self.update_status
        ).pack(side=tk.LEFT, padx=5)
        
        heart_rate_frame = ttk.Frame(tool_frame)
        heart_rate_frame.grid(row=8, column=0, sticky="w", pady=2)
        ttk.Checkbutton(
            heart_rate_frame,
            text="心率检测",
            variable=self.auto_heart_rate,
            command=self.toggle_heart_rate
        ).pack(side=tk.LEFT)
        ttk.Label(heart_rate_frame, text="(需蓝牙手环/手表)", foreground="gray").pack(side=tk.LEFT, padx=5)

        ttk.Separator(tool_frame).grid(row=9, column=0, sticky="ew", pady=5)
        
        button_frame = ttk.Frame(tool_frame)
        button_frame.grid(row=10, column=0, pady=5)
        
        ttk.Button(
            button_frame,
            text="清空历史记录",
            command=self.clear_history,
            width=15
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            button_frame,
            text="设置",
            command=self.open_settings,
            width=15
        ).pack(side=tk.LEFT)

    def toggle_advanced_music_state(self):
        """切换高级音乐信息的可用性"""
        if self.auto_music.get():
            self.advanced_music_checkbox.config(state=tk.NORMAL)
        else:
            self.advanced_music_checkbox.config(state=tk.DISABLED)
            self.advanced_music_enabled.set(False)  # 如果普通音乐被禁用，也禁用高级音乐
        self.update_status()

    def update_template_mode(self):
        """更新模板模式状态"""
        if self.use_template_mode.get():
            self.original_wrap_state = self.auto_wrap.get()
            self.auto_wrap.set(False) 
            self.wrap_checkbox.config(state=tk.DISABLED)
            self.wrap_checkbox.config(text="附加项自动换行（已启用高级模式）")
        else:
            self.auto_wrap.set(self.original_wrap_state)
            self.wrap_checkbox.config(state=tk.NORMAL)
            self.wrap_checkbox.config(text="附加项自动换行")
        self.update_status()

    def open_settings(self):
        """打开设置窗口"""
        settings_win = tk.Toplevel(self.root)
        settings_win.title("设置")
        settings_win.geometry("440x500")
        settings_win.resizable(False, False)
        settings_win.transient(self.root)
        # 移除 grab_set() 以允许主窗口接收焦点
        # settings_win.grab_set()

        notebook = ttk.Notebook(settings_win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 排序设置标签页
        order_frame = ttk.Frame(notebook)
        notebook.add(order_frame, text="发送顺序")

        ttk.Label(order_frame, text="调整附加项与消息内容的发送顺序:", font=("Arial", 10, "bold")).pack(pady=5)
        
        # 添加禁用提示（仅在模板模式启用时显示）
        if self.use_template_mode.get():
            disabled_label = ttk.Label(order_frame, text="高级模式已启用，列表排序已禁用", foreground="red")
            disabled_label.pack(pady=5)
            # 将禁用提示存储为属性，以便后续更新
            self.disabled_sorting_label = disabled_label
        else:
            # 如果不在高级模式，不显示禁用提示
            self.disabled_sorting_label = None
        
        # 创建排序列表
        order_listbox = tk.Listbox(order_frame, height=len(self.order_options))
        order_listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 填充列表框 - 按当前保存的顺序
        current_order = self.get_current_order()
        for option in current_order:
            order_listbox.insert(tk.END, option)
        # 如果还有未在列表中的选项（例如刚添加的新选项），则追加到末尾
        for option in self.order_options:
            if option not in current_order:
                order_listbox.insert(tk.END, option)

        move_up_btn = ttk.Button(order_frame, text="上移", command=lambda: self.move_item(order_listbox, -1))
        move_up_btn.pack(side=tk.LEFT, padx=5)

        move_down_btn = ttk.Button(order_frame, text="下移", command=lambda: self.move_item(order_listbox, 1))
        move_down_btn.pack(side=tk.LEFT, padx=5)

        save_order_btn = ttk.Button(order_frame, text="保存排序", command=lambda: self.save_order(order_listbox))
        save_order_btn.pack(side=tk.LEFT, padx=5)

        reset_order_btn = ttk.Button(order_frame, text="恢复默认", command=lambda: self.reset_order(order_listbox))
        reset_order_btn.pack(side=tk.LEFT, padx=5)

        hardware_frame = ttk.Frame(notebook)
        notebook.add(hardware_frame, text="硬件标签")

        ttk.Label(hardware_frame, text="自定义硬件监测标签 (留空使用默认):", font=("Arial", 10, "bold")).pack(pady=5)
        
        # CPU标签设置
        cpu_frame = ttk.Frame(hardware_frame)
        cpu_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(cpu_frame, text="CPU标签:").pack(side=tk.LEFT)
        cpu_entry = ttk.Entry(cpu_frame, textvariable=self.cpu_custom_label, width=20)
        cpu_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(cpu_frame, text="示例: CPU(%)").pack(side=tk.LEFT)

        # RAM标签设置
        ram_frame = ttk.Frame(hardware_frame)
        ram_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(ram_frame, text="RAM标签:").pack(side=tk.LEFT)
        ram_entry = ttk.Entry(ram_frame, textvariable=self.ram_custom_label, width=20)
        ram_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(ram_frame, text="示例: RAM(%)").pack(side=tk.LEFT)

        # GPU标签设置
        gpu_frame = ttk.Frame(hardware_frame)
        gpu_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(gpu_frame, text="GPU标签:").pack(side=tk.LEFT)
        gpu_entry = ttk.Entry(gpu_frame, textvariable=self.gpu_custom_label, width=20)
        gpu_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(gpu_frame, text="示例: GPU(%)").pack(side=tk.LEFT)

        # 字符数限制设置标签页
        limit_frame = ttk.Frame(notebook)
        notebook.add(limit_frame, text="字符限制")

        ttk.Label(limit_frame, text="设置获取信息的最大字符数:", font=("Arial", 10, "bold")).pack(pady=5)
        
        # 窗口标题字符限制
        window_limit_frame = ttk.Frame(limit_frame)
        window_limit_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(window_limit_frame, text="窗口标题最大字符数:").pack(side=tk.LEFT)
        window_limit_spin = ttk.Spinbox(window_limit_frame, from_=1, to=100, textvariable=self.window_title_limit, width=10)
        window_limit_spin.pack(side=tk.LEFT, padx=5)

        # 音乐标题字符限制
        music_title_limit_frame = ttk.Frame(limit_frame)
        music_title_limit_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(music_title_limit_frame, text="音乐标题最大字符数:").pack(side=tk.LEFT)
        music_title_limit_spin = ttk.Spinbox(music_title_limit_frame, from_=1, to=100, textvariable=self.music_title_limit, width=10)
        music_title_limit_spin.pack(side=tk.LEFT, padx=5)

        # 音乐艺术家字符限制
        music_artist_limit_frame = ttk.Frame(limit_frame)
        music_artist_limit_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(music_artist_limit_frame, text="音乐艺术家最大字符数:").pack(side=tk.LEFT)
        music_artist_limit_spin = ttk.Spinbox(music_artist_limit_frame, from_=1, to=100, textvariable=self.music_artist_limit, width=10)
        music_artist_limit_spin.pack(side=tk.LEFT, padx=5)

        # 高级音乐信息设置标签页
        advanced_music_frame = ttk.Frame(notebook)
        notebook.add(advanced_music_frame, text="进阶音乐信息")

        ttk.Label(advanced_music_frame, text="高级音乐信息设置:", font=("Arial", 10, "bold")).pack(pady=5)
        
        # 启用高级音乐信息复选框
        ttk.Checkbutton(
            advanced_music_frame,
            text="启用高级音乐信息（替换普通音乐信息）",
            variable=self.advanced_music_enabled,
            command=self.update_status
        ).pack(anchor="w", pady=5)
        
        # 控制按钮框架
        control_frame = ttk.Frame(advanced_music_frame)
        control_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(control_frame, text="启动网易云", command=self.launch_netease).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="选择网易云路径", command=self.browse_ncm_path).pack(side=tk.LEFT, padx=5)
        
        # 保存对开始同步歌词按钮的引用
        self.start_ncm_button = ttk.Button(control_frame, text="开始同步歌词", command=self.start_ncm_sync)
        self.start_ncm_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_ncm_button = ttk.Button(control_frame, text="停止同步歌词", command=self.stop_ncm_sync)
        self.stop_ncm_button.pack(side=tk.LEFT, padx=5)
        # 根据当前同步状态初始化按钮状态
        if self.ncm_sync_running:
            self.start_ncm_button.config(state=tk.DISABLED)
            self.stop_ncm_button.config(state=tk.NORMAL)
        else:
            self.start_ncm_button.config(state=tk.NORMAL)
            self.stop_ncm_button.config(state=tk.DISABLED)

        # 路径设置
        path_frame = ttk.Frame(advanced_music_frame)
        path_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(path_frame, text="网易云路径:").pack(side=tk.LEFT)
        path_entry = ttk.Entry(path_frame, textvariable=self.ncm_path, width=25)
        path_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(path_frame, text="(自动检测)").pack(side=tk.LEFT)
        
        # 端口设置
        port_frame = ttk.Frame(advanced_music_frame)
        port_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(port_frame, text="调试端口:").pack(side=tk.LEFT)
        port_spinbox = ttk.Spinbox(port_frame, from_=1000, to=65535, textvariable=self.ncm_port, width=10)
        port_spinbox.pack(side=tk.LEFT, padx=5)
        ttk.Label(port_frame, text="(默认: 9222)").pack(side=tk.LEFT)
        
        # 进度条设置 - 使用StringVar存储值
        bar_frame = ttk.LabelFrame(advanced_music_frame, text="进度条设置")
        bar_frame.pack(fill=tk.X, padx=10, pady=5)
        
        width_frame = ttk.Frame(bar_frame)
        width_frame.pack(fill=tk.X, pady=2)
        ttk.Label(width_frame, text="宽度:").pack(side=tk.LEFT)
        self.width_var = tk.StringVar(value=str(self.ncm_config.bar_width))
        self.width_entry = ttk.Entry(width_frame, textvariable=self.width_var, width=5)
        self.width_entry.pack(side=tk.LEFT, padx=5)
        self.width_entry.bind("<FocusOut>", lambda e: self.update_ncm_config(self.width_entry, "bar_width"))
        
        ttk.Label(width_frame, text="已播放:").pack(side=tk.LEFT, padx=(10, 0))
        self.filled_var = tk.StringVar(value=self.ncm_config.bar_filled)
        self.filled_entry = ttk.Entry(width_frame, textvariable=self.filled_var, width=5)
        self.filled_entry.pack(side=tk.LEFT, padx=5)
        self.filled_entry.bind("<FocusOut>", lambda e: self.update_ncm_config(self.filled_entry, "bar_filled"))
        
        ttk.Label(width_frame, text="滑块:").pack(side=tk.LEFT, padx=(10, 0))
        self.thumb_var = tk.StringVar(value=self.ncm_config.bar_thumb)
        self.thumb_entry = ttk.Entry(width_frame, textvariable=self.thumb_var, width=5)
        self.thumb_entry.pack(side=tk.LEFT, padx=5)
        self.thumb_entry.bind("<FocusOut>", lambda e: self.update_ncm_config(self.thumb_entry, "bar_thumb"))
        
        ttk.Label(width_frame, text="未播放:").pack(side=tk.LEFT, padx=(10, 0))
        self.empty_var = tk.StringVar(value=self.ncm_config.bar_empty)
        self.empty_entry = ttk.Entry(width_frame, textvariable=self.empty_var, width=5)
        self.empty_entry.pack(side=tk.LEFT, padx=5)
        self.empty_entry.bind("<FocusOut>", lambda e: self.update_ncm_config(self.empty_entry, "bar_empty"))
        
        # 添加"当前播放"标签
        playing_frame = ttk.Frame(advanced_music_frame)
        playing_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 从测试端口获取当前播放信息的标签
        self.current_playing_label = ttk.Label(playing_frame, text="播放中: 未连接", foreground="blue")
        self.current_playing_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 定时更新当前播放信息
        self.update_current_playing_info(settings_win)

        # 模板设置
        template_frame = ttk.LabelFrame(advanced_music_frame, text="音乐信息模板")
        template_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        ttk.Label(template_frame, text="可用变量：{song} {artist} {bar} {time} {lyric1} {lyric2}", foreground="gray").pack(anchor="w", pady=2)
        template_text = scrolledtext.ScrolledText(template_frame, height=4, font=('Arial', 10))
        template_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        template_text.insert(tk.END, self.ncm_config.template)
        template_text.bind("<FocusOut>", lambda e: self.update_ncm_config(template_text, "template"))

        send_frame = ttk.Frame(notebook)
        notebook.add(send_frame, text="发送设置")

        ttk.Label(send_frame, text="OSC发送设置:", font=("Arial", 10, "bold")).pack(pady=5)
        
        ip_frame = ttk.Frame(send_frame)
        ip_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(ip_frame, text="OSC IP地址:").pack(side=tk.LEFT)
        ip_entry = ttk.Entry(ip_frame, textvariable=self.osc_ip, width=20)
        ip_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(ip_frame, text="示例: 127.0.0.1").pack(side=tk.LEFT)

        port_frame = ttk.Frame(send_frame)
        port_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(port_frame, text="OSC端口:").pack(side=tk.LEFT)
        port_spinbox = ttk.Spinbox(
            port_frame,
            from_=1000,
            to=65535,
            textvariable=self.osc_port,
            width=10,
            command=self.update_osc_client
        )
        port_spinbox.pack(side=tk.LEFT, padx=5)
        ttk.Label(port_frame, text="示例: 9000").pack(side=tk.LEFT)

        save_config_btn = ttk.Button(send_frame, text="保存配置", command=self.save_config)
        save_config_btn.pack(pady=10)

        advanced_frame = ttk.Frame(notebook)
        notebook.add(advanced_frame, text="(高级)发送顺序")

        template_mode_check = ttk.Checkbutton(
            advanced_frame,
            text="启用模板字符串模式",
            variable=self.use_template_mode,
            command=lambda: self.update_template_mode_in_settings(template_mode_check, order_frame, order_listbox)
        )
        template_mode_check.pack(pady=5, anchor="w")

        ttk.Label(advanced_frame, text="模板字符串（支持UTF-8）:", font=("Arial", 10, "bold")).pack(pady=5)
        
        template_frame = ttk.Frame(advanced_frame)
        template_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.template_text = scrolledtext.ScrolledText(
            template_frame,
            wrap=tk.WORD,
            height=6,
            font=('Arial', 10)
        )
        self.template_text.pack(fill=tk.BOTH, expand=True)
        self.template_text.insert(tk.END, self.template_string.get())
        
        save_template_btn = ttk.Button(
            advanced_frame,
            text="保存模板",
            command=self.save_template_string
        )
        save_template_btn.pack(pady=5)

        # 附加项快捷按钮区域
        quick_buttons_frame = ttk.LabelFrame(advanced_frame, text="快速插入附加项")
        quick_buttons_frame.pack(fill=tk.X, padx=10, pady=5)

        # 定义附加项及其对应的模板变量
        quick_buttons = [
            ("消息内容", "{message}"),
            ("时间", "{time}"),
            ("窗口标题", "{window}"),
            ("挂机状态", "{idle}"),
            ("音乐信息", "{music}"),
            ("硬件监测", "{hardware}"),
            ("心率", "{heart_rate}"),
            ("换行", "\\n")
        ]

        for i, (label, value) in enumerate(quick_buttons):
            row = i // 4
            col = i % 4
            btn = ttk.Button(
                quick_buttons_frame,
                text=label,
                width=12,
                command=lambda v=value: self.insert_template_variable(v)
            )
            btn.grid(row=row, column=col, padx=2, pady=2, sticky="ew")

        for i in range(4):
            quick_buttons_frame.columnconfigure(i, weight=1)

        close_btn = ttk.Button(settings_win, text="关闭", command=lambda: self.close_settings_window(settings_win))
        close_btn.pack(pady=10)

        settings_win.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (settings_win.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (settings_win.winfo_height() // 2)
        settings_win.geometry(f"+{x}+{y}")

    def update_current_playing_info(self, settings_window=None):
        """定时更新当前播放信息"""
        if settings_window and not settings_window.winfo_exists():  # 检查窗口是否存在
            return
        if self.ncm_sync_running:
            with self.ncm_shared_state.lock:
                state = self.ncm_shared_state.data
                if state.song and state.artist:
                    duration_min = state.dur // 60
                    duration_sec = state.dur % 60
                    if settings_window and self.current_playing_label.winfo_exists():  # 检查标签是否存在
                        self.current_playing_label.config(
                            text=f"播放中: {state.song[:self.music_title_limit.get()]} - {state.artist[:self.music_artist_limit.get()]} - {duration_min}:{duration_sec:02d}",
                            foreground="green"
                        )
                else:
                    if settings_window and self.current_playing_label.winfo_exists():  # 检查标签是否存在
                        self.current_playing_label.config(text="播放中: 无音乐", foreground="red")
        else:
            if settings_window and self.current_playing_label.winfo_exists():  # 检查标签是否存在
                self.current_playing_label.config(text="播放中: 未连接", foreground="red")
        
        # 每秒更新一次
        if settings_window and settings_window.winfo_exists():  # 只有窗口存在时才安排下次更新
            self.root.after(1000, lambda: self.update_current_playing_info(settings_window))

    def close_settings_window(self, settings_window):
        """关闭设置窗口并停止更新播放信息"""
        settings_window.destroy()

    def update_ncm_config(self, widget, field):
        """更新网易云配置"""
        try:
            if field == "bar_width":
                self.ncm_config.bar_width = int(widget.get())
            elif field == "bar_filled":
                self.ncm_config.bar_filled = widget.get()
            elif field == "bar_thumb":
                self.ncm_config.bar_thumb = widget.get()
            elif field == "bar_empty":
                self.ncm_config.bar_empty = widget.get()
            elif field == "template":
                self.ncm_config.template = widget.get("1.0", "end-1c")
                # 立即应用模板变化
                self.update_status()
        except ValueError:
            if field == "bar_width":
                messagebox.showerror("错误", "进度条宽度必须是数字")

    def launch_netease(self):
        """启动网易云音乐"""
        path = self.ncm_path.get() if self.ncm_path.get() else None
        port = self.ncm_port.get()
        self.ncm_config.ncm_port = port
        self.ncm_config.ncm_path = path if path else ""
        
        if self.ncm_launch_thread and self.ncm_launch_thread.is_alive():
            messagebox.showwarning("警告", "网易云已在启动中，请稍候")
            return
            
        def launch_task():
            ok, result, detected_port = launch_netease(port, path)
            if ok:
                self.ncm_config.ncm_port = detected_port
                self.ncm_port.set(detected_port)
                self.root.after(0, lambda: messagebox.showinfo("成功", f"网易云已启动，调试端口: {detected_port}"))
            else:
                self.root.after(0, lambda: messagebox.showerror("错误", f"启动失败: {result}"))
                
        self.ncm_launch_thread = threading.Thread(target=launch_task, daemon=True)
        self.ncm_launch_thread.start()

    def browse_ncm_path(self):
        """浏览并选择网易云路径"""
        from tkinter import filedialog
        path = filedialog.askopenfilename(title="选择网易云音乐程序", filetypes=[("Executable files", "*.exe"), ("All files", "*.*")])
        if path:
            self.ncm_path.set(path)
            self.ncm_config.ncm_path = path

    def start_ncm_sync(self):
        """开始同步网易云音乐信息"""
        if self.ncm_sync_running:
            # 不再弹出警告窗口，而是直接返回
            return

        self.ncm_stop_event = threading.Event()
        self.ncm_sync_thread = threading.Thread(
            target=netease_thread,
            args=(self.ncm_config, self.ncm_shared_state, self.ncm_stop_event, self),
            daemon=True
        )
        self.ncm_sync_thread.start()
        self.ncm_sync_running = True
        
        # 更新按钮状态
        self.start_ncm_button.config(state=tk.DISABLED)
        self.stop_ncm_button.config(state=tk.NORMAL)

    def stop_ncm_sync(self):
        """停止同步网易云音乐信息"""
        if not self.ncm_sync_running:
            # 不再弹出警告窗口，而是直接返回
            return

        if self.ncm_stop_event:
            self.ncm_stop_event.set()
        if self.ncm_sync_thread and self.ncm_sync_thread.is_alive():
            self.ncm_sync_thread.join(timeout=2)
        self.ncm_sync_running = False
        
        # 更新按钮状态
        self.start_ncm_button.config(state=tk.NORMAL)
        self.stop_ncm_button.config(state=tk.DISABLED)

    def cb_status(self, t: str):
        """回调函数，用于更新状态（模拟接口）"""
        pass

    def cb_song(self, t: str):
        """回调函数，用于更新歌曲信息（模拟接口）"""
        pass

    def cb_output(self, t: str):
        """回调函数，用于输出内容（模拟接口）"""
        pass

    def update_template_mode_in_settings(self, checkbox_widget, order_frame, order_listbox):
        """在设置窗口中更新模板模式状态，并实时更新UI"""
        self.update_template_mode()
        self.save_config()
        
        # 清除现有的提示标签（通过查找所有子部件并检查类型）
        for widget in order_frame.winfo_children():
            if isinstance(widget, ttk.Label) and str(widget) not in [str(child) for child in order_frame.winfo_children() if not isinstance(child, ttk.Label)]:
                # 检查文本内容来确认是否是禁用提示标签
                try:
                    if widget.cget('text') == "高级模式已启用，列表排序已禁用":
                        widget.destroy()
                except tk.TclError:
                    # 如果这个widget没有text属性，跳过
                    continue
        
        # 根据当前模板模式状态添加新的提示标签
        if self.use_template_mode.get():
            disabled_label = ttk.Label(order_frame, text="高级模式已启用，列表排序已禁用", foreground="red")
            disabled_label.pack(pady=5)
            self.disabled_sorting_label = disabled_label
        else:
            # 如果不在高级模式，不显示禁用提示
            self.disabled_sorting_label = None

    def save_template_string(self):
        template = self.template_text.get("1.0", "end-1c")
        self.template_string.set(template)
        self.save_config()
        messagebox.showinfo("提示", "模板字符串已保存!")

    def insert_template_variable(self, variable):
        self.template_text.insert(tk.INSERT, variable)

    def move_item(self, listbox, direction):
        selection = listbox.curselection()
        if not selection:
            return

        index = selection[0]
        if (direction == -1 and index == 0) or (direction == 1 and index == listbox.size() - 1):
            return

        # 获取当前选项
        item = listbox.get(index)
        # 删除当前项
        listbox.delete(index)
        # 插入到新位置
        new_index = index + direction
        listbox.insert(new_index, item)
        # 重新选择该项
        listbox.selection_set(new_index)

    def save_order(self, listbox):
        """保存排序设置"""
        # 更新order_vars字典 - 按照列表框中的顺序
        for i, option in enumerate(listbox.get(0, tk.END)):
            if option in self.order_vars:
                self.order_vars[option].set(str(i))
        
        # 保存到配置文件
        self.save_config()
        messagebox.showinfo("提示", "排序已保存!")

    def reset_order(self, listbox):
        listbox.delete(0, tk.END)
        default_order = ["时间" , "消息内容" , "挂机状态", "窗口标题", "心率", "硬件监测" , "音乐信息"]
        for option in default_order:
            listbox.insert(tk.END, option)
        
        # 更新order_vars为默认顺序
        for i, option in enumerate(default_order):
            if option in self.order_vars:
                self.order_vars[option].set(str(i))
        
        # 保存到配置文件
        self.save_config()
        messagebox.showinfo("提示", "已恢复默认排序!")

    def toggle_heart_rate(self):
        """切换心率检测状态"""
        if self.auto_heart_rate.get():
            self.heart_rate_monitor.start()
        else:
            self.heart_rate_monitor.stop()
        self.update_status()

    def create_debug_panel(self, parent):
        self.debug_frame = ttk.LabelFrame(parent, text="调试信息")
        debug_items = [
            ("window", "当前窗口标题:"),
            ("music_title", "音乐标题:"),
            ("music_artist", "音乐艺术家:"),
            ("idle_time", "挂机时间:"),
            ("next_send", "下次发送:"),
            ("heart_rate", "手环/心率:"), 
            ("cpu_usage", "CPU 使用率:"), 
            ("ram_usage", "RAM 使用率:"), 
            ("gpu_usage", "GPU 使用率:") 
        ]
        
        for i, (key, text) in enumerate(debug_items):
            frame = ttk.Frame(self.debug_frame)
            frame.grid(row=i, column=0, sticky="ew", padx=5, pady=2)
            
            ttk.Label(frame, text=text, width=12, anchor="e").pack(side=tk.LEFT)
            self.debug_labels[key] = ttk.Label(frame, text="", foreground="#666666", width=25)
            self.debug_labels[key].pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.debug_frame.pack_forget()

    async def _get_media_info_async(self):
        try:
            sessions = await MediaManager.request_async()
            current_session = sessions.get_current_session()
            
            if current_session and current_session.get_playback_info().playback_status == \
                GlobalSystemMediaTransportControlsSessionPlaybackStatus.PLAYING:
                
                media_properties = await current_session.try_get_media_properties_async()
                return {
                    "title": media_properties.title if media_properties.title else "未知曲目",
                    "artist": media_properties.artist if media_properties.artist else "未知艺术家"
                }
        except Exception as e:
            print(f"SMTC 获取失败: {str(e)}")
        return None

    def get_formatted_music_info(self):
        # 如果启用了高级音乐信息，则返回高级信息
        if self.advanced_music_enabled.get() and self.ncm_sync_running:
            with self.ncm_shared_state.lock:
                state = self.ncm_shared_state.data.copy()
                lyrics = list(self.ncm_shared_state.lyrics)
                song_key = self.ncm_shared_state.song_key
            formatted_output = format_output(self.ncm_config, state, lyrics, song_key, self.music_title_limit.get(), self.music_artist_limit.get())
            return formatted_output
        else:
            try:
                music_info = self.loop.run_until_complete(self._get_media_info_async())
                if music_info:
                    title = music_info['title'][:self.music_title_limit.get()]
                    artist = music_info['artist'][:self.music_artist_limit.get()]
                    return f"[在听: {title} - {artist}]"
            except:
                pass
            
            try:
                hwnd = win32gui.GetForegroundWindow()
                title = win32gui.GetWindowText(hwnd)
                if "网易云音乐" in title:
                    match = re.match(r"(.+?)\s*-\s*(.+?)\s*-\s*.+?\s*网易云音乐", title)
                    if match:
                        title = match.group(1)[:self.music_title_limit.get()]
                        artist = match.group(2)[:self.music_artist_limit.get()]
                        return f"[在听: {title} - {artist}]"
            except:
                pass
            return ""

    def get_idle_duration(self):
        try:
            last_input = win32api.GetLastInputInfo()
            current_tick = win32api.GetTickCount()
            return (current_tick - last_input) // 1000
        except Exception as e:
            print(f"获取空闲时间失败: {str(e)}")
            return 0

    def format_duration(self, seconds):
        if seconds < 60:
            return f"{seconds}秒"
        minutes, sec = divmod(seconds, 60)
        return f"{minutes}分{sec}秒"

    def update_char_count(self, event=None):
        current = len(self.text_input.get("1.0", "end-1c"))
        remaining = self.max_message_length - current - self.calculate_additional_length()
        color = "green" if remaining >= 20 else "orange" if remaining >= 0 else "red"
        self.chars_remaining_label.configure(foreground=color)
        self.chars_remaining_var.set(f"可用字符: {remaining}/{self.max_message_length}")

    def on_paste(self, event):
        try:
            text = self.root.clipboard_get()
            current = self.text_input.get("1.0", "end-1c")
            available = self.max_message_length - len(current) - self.calculate_additional_length()
            
            if available <= 0:
                return "break"
            
            self.text_input.insert(tk.INSERT, text[:available])
            self.update_char_count()
        except tk.TclError:
            messagebox.showwarning("粘贴错误", "无法读取剪贴板内容")
        return "break"

    def get_formatted_time(self):
        return (datetime.utcnow() + timedelta(hours=8)).strftime("[时间:%H:%M]")

    def get_formatted_window_title(self):
        try:
            title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            return f"[在看:{title[:self.window_title_limit.get()]}]"
        except:
            return ""

    def calculate_additional_length(self):
        total = 0
        
        if self.use_template_mode.get():
            # 在模板模式下，计算模板字符串的长度
            template = self.template_string.get()
            # 替换模板变量为实际内容
            replacements = {
                '{message}': self.text_input.get("1.0", "end-1c").rstrip('\n'),
                '{time}': self.get_formatted_time() if self.auto_time.get() else '',
                '{window}': self.get_formatted_window_title() if self.auto_window.get() else '',
                '{idle}': f"[已挂机: {self.format_duration(self.get_idle_duration())}]" if self.auto_idle.get() and self.get_idle_duration() >= self.idle_threshold.get() else '',
                '{music}': self.get_formatted_music_info() if self.auto_music.get() else '',
                '{heart_rate}': f"[❤️:{self.heart_rate_monitor.current_hr} BPM]" if self.auto_heart_rate.get() and self.heart_rate_monitor.is_connected and self.heart_rate_monitor.current_hr > 0 else '',
            }
            
            # 处理硬件监测
            hardware_parts = []
            if self.auto_hardware.get():
                if self.auto_cpu.get():
                    cpu_usage = self.get_cpu_usage()
                    if cpu_usage != "N/A":
                        if self.cpu_custom_label.get():
                            custom_label = self.cpu_custom_label.get()
                            hardware_parts.append(f"CPU({custom_label}): {cpu_usage:.0f}%")
                        else:
                            hardware_parts.append(f"CPU: {cpu_usage:.0f}%")
                if self.auto_ram.get():
                    ram_usage = self.get_ram_usage()
                    if ram_usage != "N/A":
                        if self.ram_custom_label.get():
                            custom_label = self.ram_custom_label.get()
                            hardware_parts.append(f"RAM({custom_label}): {ram_usage:.0f}%")
                        else:
                            hardware_parts.append(f"RAM: {ram_usage:.0f}%")
                if self.auto_gpu.get():
                    gpu_usage = self.get_gpu_usage()
                    if gpu_usage and gpu_usage != "无法获取GPU数据":
                        if self.gpu_custom_label.get():
                            custom_label = self.gpu_custom_label.get()
                            hardware_parts.append(f"GPU({custom_label}): {gpu_usage}")
                        else:
                            hardware_parts.append(f"GPU: {gpu_usage}")
            
            if hardware_parts:
                hardware_str = f"[{', '.join(hardware_parts)}]"
            else:
                hardware_str = ''
                
            replacements['{hardware}'] = hardware_str
            
            # 替换模板变量
            result = template
            for var, replacement in replacements.items():
                result = result.replace(var, replacement)
            
            # 替换换行符
            result = result.replace('\\n', '\n')
            
            total = len(result)
        else:
            # 传统模式下的计算
            if self.auto_idle.get() and self.get_idle_duration() >= self.idle_threshold.get():
                total += len("[已挂机: 999分99秒]")
            if self.auto_time.get():
                total += len(self.get_formatted_time())
            if self.auto_window.get():
                window_str = self.get_formatted_window_title()
                if window_str:
                    total += len(window_str)
            if self.auto_music.get():
                music_str = self.get_formatted_music_info()
                if music_str:
                    total += len(music_str)
            if self.auto_hardware.get():
                hardware_parts = []
                if self.auto_cpu.get():
                    cpu_usage = self.get_cpu_usage()
                    if cpu_usage != "N/A":
                        if self.cpu_custom_label.get():  # 如果有自定义标签
                            # 格式: CPU(标签): 占用率%
                            custom_label = self.cpu_custom_label.get()
                            hardware_parts.append(f"CPU({custom_label}): {cpu_usage:.0f}%")
                        else:  # 使用默认格式
                            hardware_parts.append(f"CPU: {cpu_usage:.0f}%")
                if self.auto_ram.get():
                    ram_usage = self.get_ram_usage()
                    if ram_usage != "N/A":
                        if self.ram_custom_label.get():
                            custom_label = self.ram_custom_label.get()
                            hardware_parts.append(f"RAM({custom_label}): {ram_usage:.0f}%")
                        else:
                            hardware_parts.append(f"RAM: {ram_usage:.0f}%")
                if self.auto_gpu.get():
                    gpu_usage = self.get_gpu_usage()
                    if gpu_usage and gpu_usage != "无法获取GPU数据":
                        if self.gpu_custom_label.get():
                            custom_label = self.gpu_custom_label.get()
                            hardware_parts.append(f"GPU({custom_label}): {gpu_usage}")
                        else:
                            hardware_parts.append(f"GPU: {gpu_usage}")
                if hardware_parts:
                    total += len(f"[{', '.join(hardware_parts)}]")
            if self.auto_heart_rate.get() and self.heart_rate_monitor.is_connected:
                total += len("[❤️:120 BPM]")
        
        return total

    def process_message(self, raw_message):
        if self.use_template_mode.get():
            # 使用模板字符串模式
            template = self.template_string.get()
            
            # 准备替换内容
            replacements = {
                '{message}': raw_message.rstrip('\n'),
                '{time}': self.get_formatted_time() if self.auto_time.get() else '',
                '{window}': self.get_formatted_window_title() if self.auto_window.get() else '',
                '{idle}': f"[已挂机: {self.format_duration(self.get_idle_duration())}]" if self.auto_idle.get() and self.get_idle_duration() >= self.idle_threshold.get() else '',
                '{music}': self.get_formatted_music_info() if self.auto_music.get() else '',
                '{heart_rate}': f"[❤️:{self.heart_rate_monitor.current_hr} BPM]" if self.auto_heart_rate.get() and self.heart_rate_monitor.is_connected and self.heart_rate_monitor.current_hr > 0 else '',
            }
            
            # 处理硬件监测
            hardware_parts = []
            if self.auto_hardware.get():
                if self.auto_cpu.get():
                    cpu_usage = self.get_cpu_usage()
                    if cpu_usage != "N/A":
                        if self.cpu_custom_label.get():
                            custom_label = self.cpu_custom_label.get()
                            hardware_parts.append(f"CPU({custom_label}): {cpu_usage:.0f}%")
                        else:
                            hardware_parts.append(f"CPU: {cpu_usage:.0f}%")
                if self.auto_ram.get():
                    ram_usage = self.get_ram_usage()
                    if ram_usage != "N/A":
                        if self.ram_custom_label.get():
                            custom_label = self.ram_custom_label.get()
                            hardware_parts.append(f"RAM({custom_label}): {ram_usage:.0f}%")
                        else:
                            hardware_parts.append(f"RAM: {ram_usage:.0f}%")
                if self.auto_gpu.get():
                    gpu_usage = self.get_gpu_usage()
                    if gpu_usage and gpu_usage != "无法获取GPU数据":
                        if self.gpu_custom_label.get():
                            custom_label = self.gpu_custom_label.get()
                            hardware_parts.append(f"GPU({custom_label}): {gpu_usage}")
                        else:
                            hardware_parts.append(f"GPU: {gpu_usage}")
            
            if hardware_parts:
                hardware_str = f"[{', '.join(hardware_parts)}]"
            else:
                hardware_str = ''
                
            replacements['{hardware}'] = hardware_str
            
            # 替换模板变量
            result = template
            for var, replacement in replacements.items():
                result = result.replace(var, replacement)
            
            # 替换换行符
            result = result.replace('\\n', '\n')
            
            return result
        else:
            # 传统模式
            additions = {}
            wrap_char = "\n" if self.auto_wrap.get() else " "
        
            # 收集所有附加项
            if self.auto_idle.get():
                idle_sec = self.get_idle_duration()
                if idle_sec >= self.idle_threshold.get():
                    additions["挂机状态"] = f"[已挂机: {self.format_duration(idle_sec)}]"
        
            if self.auto_time.get():
                additions["时间"] = self.get_formatted_time()
        
            if self.auto_window.get():
                window_title = self.get_formatted_window_title()
                if window_title:
                    additions["窗口标题"] = window_title

            if self.auto_heart_rate.get() and self.heart_rate_monitor.is_connected:
                hr = self.heart_rate_monitor.current_hr
                if hr > 0:
                    additions["心率"] = f"[❤️:{hr} BPM]"

            if self.auto_music.get():
                music_info = self.get_formatted_music_info()
                if music_info:
                    additions["音乐信息"] = music_info
        
            hardware_parts = []
            if self.auto_hardware.get():
                if self.auto_cpu.get():
                    cpu_usage = self.get_cpu_usage()
                    if cpu_usage != "N/A":
                        # 修复后的硬件标签逻辑
                        if self.cpu_custom_label.get():  # 如果有自定义标签
                            # 格式: CPU(标签): 占用率%
                            custom_label = self.cpu_custom_label.get()
                            hardware_parts.append(f"CPU({custom_label}): {cpu_usage:.0f}%")
                        else:  # 使用默认格式
                            hardware_parts.append(f"CPU: {cpu_usage:.0f}%")
                if self.auto_ram.get():
                    ram_usage = self.get_ram_usage()
                    if ram_usage != "N/A":
                        if self.ram_custom_label.get():
                            custom_label = self.ram_custom_label.get()
                            hardware_parts.append(f"RAM({custom_label}): {ram_usage:.0f}%")
                        else:
                            hardware_parts.append(f"RAM: {ram_usage:.0f}%")
                if self.auto_gpu.get():
                    gpu_usage = self.get_gpu_usage()
                    if gpu_usage and gpu_usage != "无法获取GPU数据":
                        if self.gpu_custom_label.get():
                            custom_label = self.gpu_custom_label.get()
                            hardware_parts.append(f"GPU({custom_label}): {gpu_usage}")
                        else:
                            hardware_parts.append(f"GPU: {gpu_usage}")
            
            if hardware_parts:
                additions["硬件监测"] = f"[{', '.join(hardware_parts)}]"
            
            # 按照设置的顺序组织附加项和消息内容
            ordered_parts = []
            sorted_items = sorted(self.order_vars.items(), key=lambda x: int(x[1].get()))
            
            for item_name, _ in sorted_items:
                if item_name == "消息内容":
                    # 添加用户输入的消息内容
                    ordered_parts.append(raw_message.rstrip('\n'))
                elif item_name in additions:
                    ordered_parts.append(additions[item_name])
        
            # 组合最终消息
            if ordered_parts:
                full_message = wrap_char.join(ordered_parts)
            else:
                full_message = raw_message.rstrip('\n')
        
            return full_message.strip()

    def send_message(self):
        raw_message = self.text_input.get("1.0", "end-1c").rstrip('\n')
        
        # 检查是否有启用的功能或消息内容
        has_enabled_feature = (
            self.auto_time.get() or
            self.auto_window.get() or
            self.auto_music.get() or
            self.auto_idle.get() or
            self.auto_hardware.get() or
            self.auto_heart_rate.get()
        )
        
        if not raw_message.strip() and not has_enabled_feature:
            messagebox.showwarning("提示", "请先输入消息内容或启用至少一个功能！")
            return False

        try:
            final_message = self.process_message(raw_message)
            self.osc_client.send_message("/chatbox/input", [final_message, True])
            
            # 记录到历史记录的消息（按设置的顺序）
            history_msg = self.process_history_message(raw_message)
            
            self.send_to_history(history_msg)
            self.update_char_count()
            self.last_send_time = time.time()
            return True
        except Exception as e:
            messagebox.showerror("错误", f"消息发送失败: {str(e)}")
            return False

    def process_history_message(self, raw_message):
        """处理历史记录消息，按照设置的顺序排列"""
        if self.use_template_mode.get():
            # 使用模板字符串模式
            template = self.template_string.get()
            
            # 准备替换内容
            replacements = {
                '{message}': raw_message.rstrip('\n'),
                '{time}': self.get_formatted_time() if self.auto_time.get() else '',
                '{window}': self.get_formatted_window_title() if self.auto_window.get() else '',
                '{idle}': f"[已挂机: {self.format_duration(self.get_idle_duration())}]" if self.auto_idle.get() and self.get_idle_duration() >= self.idle_threshold.get() else '',
                '{music}': self.get_formatted_music_info() if self.auto_music.get() else '',
                '{heart_rate}': f"[❤️:{self.heart_rate_monitor.current_hr} BPM]" if self.auto_heart_rate.get() and self.heart_rate_monitor.is_connected and self.heart_rate_monitor.current_hr > 0 else '',
            }
            
            # 处理硬件监测
            hardware_parts = []
            if self.auto_hardware.get():
                if self.auto_cpu.get():
                    cpu_usage = self.get_cpu_usage()
                    if cpu_usage != "N/A":
                        if self.cpu_custom_label.get():
                            custom_label = self.cpu_custom_label.get()
                            hardware_parts.append(f"CPU({custom_label}): {cpu_usage:.0f}%")
                        else:
                            hardware_parts.append(f"CPU: {cpu_usage:.0f}%")
                if self.auto_ram.get():
                    ram_usage = self.get_ram_usage()
                    if ram_usage != "N/A":
                        if self.ram_custom_label.get():
                            custom_label = self.ram_custom_label.get()
                            hardware_parts.append(f"RAM({custom_label}): {ram_usage:.0f}%")
                        else:
                            hardware_parts.append(f"RAM: {ram_usage:.0f}%")
                if self.auto_gpu.get():
                    gpu_usage = self.get_gpu_usage()
                    if gpu_usage and gpu_usage != "无法获取GPU数据":
                        if self.gpu_custom_label.get():
                            custom_label = self.gpu_custom_label.get()
                            hardware_parts.append(f"GPU({custom_label}): {gpu_usage}")
                        else:
                            hardware_parts.append(f"GPU: {gpu_usage}")
            
            if hardware_parts:
                hardware_str = f"[{', '.join(hardware_parts)}]"
            else:
                hardware_str = ''
                
            replacements['{hardware}'] = hardware_str
            
            # 替换模板变量
            result = template
            for var, replacement in replacements.items():
                result = result.replace(var, replacement)
            
            # 替换换行符
            result = result.replace('\\n', '\n')
            
            return result
        else:
            # 传统模式
            additions = {}
            wrap_char = "\n" if self.auto_wrap.get() else " "
        
            # 收集所有附加项
            if self.auto_idle.get():
                idle_sec = self.get_idle_duration()
                if idle_sec >= self.idle_threshold.get():
                    additions["挂机状态"] = f"[已挂机: {self.format_duration(idle_sec)}]"
        
            if self.auto_time.get():
                additions["时间"] = self.get_formatted_time()
        
            if self.auto_window.get():
                window_title = self.get_formatted_window_title()
                if window_title:
                    additions["窗口标题"] = window_title

            if self.auto_heart_rate.get() and self.heart_rate_monitor.is_connected:
                hr = self.heart_rate_monitor.current_hr
                if hr > 0:
                    additions["心率"] = f"[❤️:{hr} BPM]"

            if self.auto_music.get():
                music_info = self.get_formatted_music_info()
                if music_info:
                    additions["音乐信息"] = music_info
        
            hardware_parts = []
            if self.auto_hardware.get():
                if self.auto_cpu.get():
                    cpu_usage = self.get_cpu_usage()
                    if cpu_usage != "N/A":
                        # 修复后的硬件标签逻辑
                        if self.cpu_custom_label.get():  # 如果有自定义标签
                            # 格式: CPU(标签): 占用率%
                            custom_label = self.cpu_custom_label.get()
                            hardware_parts.append(f"CPU({custom_label}): {cpu_usage:.0f}%")
                        else:  # 使用默认格式
                            hardware_parts.append(f"CPU: {cpu_usage:.0f}%")
                if self.auto_ram.get():
                    ram_usage = self.get_ram_usage()
                    if ram_usage != "N/A":
                        if self.ram_custom_label.get():
                            custom_label = self.ram_custom_label.get()
                            hardware_parts.append(f"RAM({custom_label}): {ram_usage:.0f}%")
                        else:
                            hardware_parts.append(f"RAM: {ram_usage:.0f}%")
                if self.auto_gpu.get():
                    gpu_usage = self.get_gpu_usage()
                    if gpu_usage and gpu_usage != "无法获取GPU数据":
                        if self.gpu_custom_label.get():
                            custom_label = self.gpu_custom_label.get()
                            hardware_parts.append(f"GPU({custom_label}): {gpu_usage}")
                        else:
                            hardware_parts.append(f"GPU: {gpu_usage}")
            
            if hardware_parts:
                additions["硬件监测"] = f"[{', '.join(hardware_parts)}]"
            
            # 按照设置的顺序组织附加项和消息内容
            ordered_parts = []
            sorted_items = sorted(self.order_vars.items(), key=lambda x: int(x[1].get()))
            
            for item_name, _ in sorted_items:
                if item_name == "消息内容":
                    # 添加用户输入的消息内容
                    ordered_parts.append(raw_message.rstrip('\n'))
                elif item_name in additions:
                    ordered_parts.append(additions[item_name])
        
            # 组合最终消息
            if ordered_parts:
                full_message = wrap_char.join(ordered_parts)
            else:
                full_message = raw_message.rstrip('\n')
        
            return full_message.strip()

    def send_to_history(self, message):
        current_time = datetime.now().strftime('%H:%M:%S')
        formatted_message = f"[{len(self.history_list)+1}] ({current_time}):\n"
        
        if self.use_template_mode.get():
            # 模板模式下，直接使用处理后的消息，保留模板中的换行
            formatted_message += message
        else:
            # 传统模式下，根据自动换行设置决定换行方式
            if self.auto_wrap.get():
                formatted_message += message.replace("\n", "\n")
            else:
                formatted_message += message.replace("\n", " ")
                
        formatted_message += "\n"
        self.history_list.append(formatted_message)
        
        if len(self.history_list) > self.history_max_items:
            self.history_list.pop(0)
        
        self.history_text.delete(1.0, tk.END)
        self.history_text.insert(tk.END, ''.join(self.history_list))
        self.history_text.see(tk.END)

    def toggle_sending(self):
        if not self.is_sending:
            raw_message = self.text_input.get("1.0", "end-1c").strip()
            has_enabled_feature = (
                self.auto_time.get() or
                self.auto_window.get() or
                self.auto_music.get() or
                self.auto_idle.get() or
                self.auto_hardware.get() or
                self.auto_heart_rate.get()
            )
            
            if not raw_message and not has_enabled_feature:
                messagebox.showwarning("提示", "请先输入消息内容或启用至少一个功能！")
                return
                
            self.start_sending()
            self.start_btn.config(text="停止发送")
        else:
            self.stop_sending()
            self.start_btn.config(text="开始发送")

    def start_sending(self):
        self.is_sending = True
        self.status_var.set("正在自动发送消息...")
        interval = self.interval_var.get()
        success = self.send_message()
        if success:
            self.scheduled_event = self.root.after(interval * 1000, self.scheduled_send_status)
            self.update_countdown(interval)
            self.start_debug_update()
        else:
            # 如果第一次发送就失败了，停止发送并更新UI
            self.stop_sending()

    def start_debug_update(self):
        if not self.debug_update_job:
            self.update_debug_info()

    def stop_debug_update(self):
        if self.debug_update_job:
            self.root.after_cancel(self.debug_update_job)
            self.debug_update_job = None

    def update_debug_info(self):
        try:
            title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            self.debug_labels['window'].config(text=title[:self.window_title_limit.get()] if title else "无")
            
            music_info = self.get_raw_music_info()
            self.debug_labels['music_title'].config(text=music_info['title'][:self.music_title_limit.get()] if music_info else "无")
            self.debug_labels['music_artist'].config(text=music_info['artist'][:self.music_artist_limit.get()] if music_info else "无")
            
            idle_sec = self.get_idle_duration()
            self.debug_labels['idle_time'].config(text=self.format_duration(idle_sec))
            
            if self.is_sending:
                remaining = self.interval_var.get() - (time.time() - self.last_send_time)
                self.debug_labels['next_send'].config(text=f"{max(0, int(remaining))}秒")
            else:
                self.debug_labels['next_send'].config(text="未启用")

            if self.auto_heart_rate.get():
                if self.heart_rate_monitor.is_connected:
                    hr = self.heart_rate_monitor.current_hr
                    device = self.heart_rate_monitor.device_name or "未知设备"
                    self.debug_labels['heart_rate'].config(
                        text=f"{device[:18]} | {hr} BPM", 
                    )
                else:
                    self.debug_labels['heart_rate'].config(
                        text="扫描中...", 
                        foreground="#999999"
                    )
            else:
                self.debug_labels['heart_rate'].config(text="未启用", foreground="#666666")

            cpu_message = "未启用功能"
            ram_message = "未启用功能"
            gpu_message = "未启用功能"

            if self.auto_hardware.get(): 
                cpu_usage = self.get_cpu_usage()
                if cpu_usage != "N/A":
                    if self.cpu_custom_label.get():
                        cpu_message = f"{cpu_usage:.0f}%"
                    else:
                        cpu_message = f"{cpu_usage:.0f}%"
                else:
                    cpu_message = cpu_usage

                ram_usage = self.get_ram_usage()
                if ram_usage != "N/A":
                    if self.ram_custom_label.get():
                        ram_message = f"{ram_usage:.0f}%"
                    else:
                        ram_message = f"{ram_usage:.0f}%"
                else:
                    ram_message = ram_usage

                gpu_usage = self.get_gpu_usage()
                if gpu_usage != "无法获取GPU数据":
                    if self.gpu_custom_label.get():
                        gpu_message = f"{gpu_usage}"
                    else:
                        gpu_message = f"{gpu_usage}"
                else:
                    gpu_message = f"未启用功能，{gpu_usage}"

            self.debug_labels['cpu_usage'].config(text=cpu_message)
            self.debug_labels['ram_usage'].config(text=ram_message)
            self.debug_labels['gpu_usage'].config(text=gpu_message)
            
        except Exception as e:
            print(f"调试更新错误: {str(e)}")
        
        self.debug_update_job = self.root.after(
            self.debug_update_interval, 
            self.update_debug_info
        )

    def get_raw_music_info(self):
        try:
            music_info = self.loop.run_until_complete(self._get_media_info_async())
            if music_info:
                return {
                    "title": music_info["title"],
                    "artist": music_info["artist"]
                }
            
            hwnd = win32gui.GetForegroundWindow()
            title = win32gui.GetWindowText(hwnd)
            if "网易云音乐" in title:
                match = re.match(r"(.+?)\s*-\s*(.+?)\s*-\s*.+?\s*网易云音乐", title)
                if match:
                    return {"title": match.group(1), "artist": match.group(2)}
        except:
            pass
        return None

    def stop_sending(self):
        if self.scheduled_event:
            self.root.after_cancel(self.scheduled_event)
        self.is_sending = False
        self.status_var.set("已停止发送")
        self.countdown_var.set("")
        self.stop_debug_update()
        # 更新按钮文本为"开始发送"
        self.start_btn.config(text="开始发送")

    def update_countdown(self, remaining_seconds):
        if remaining_seconds >= 0 and self.is_sending:
            self.countdown_var.set(f"下次发送还剩 {remaining_seconds} 秒")
            if remaining_seconds > 0:
                self.root.after(1000, self.update_countdown, remaining_seconds - 1)
        else:
            self.countdown_var.set("")

    def scheduled_send_status(self):
        if self.is_sending:
            interval = self.interval_var.get()
            success = self.send_message()
            if success:
                self.scheduled_event = self.root.after(interval * 1000, self.scheduled_send_status)
                self.update_countdown(interval)
            else:
                # 如果发送失败，停止发送并更新UI状态
                self.stop_sending()

    def update_status(self):
        status = []
        if self.auto_time.get():
            status.append("时间")
        if self.auto_window.get():
            status.append("窗口标题")
        if self.auto_wrap.get() and not self.use_template_mode.get():
            status.append("换行")
        if self.auto_music.get():
            if self.advanced_music_enabled.get():
                status.append("高级音乐信息")
            else:
                status.append("音乐信息")
        if self.auto_idle.get():
            idle_threshold = self.idle_threshold.get() if self.idle_threshold.get() is not None else 60
            status.append(f"挂机检测({idle_threshold}秒)")
        if self.auto_hardware.get():
            hardware_parts = []
            if self.auto_cpu.get():
                hardware_parts.append("CPU")
            if self.auto_ram.get():
                hardware_parts.append("RAM")
            if self.auto_gpu.get():
                hardware_parts.append("GPU")
            if hardware_parts:
                status.append(f"硬件监测（{', '.join(hardware_parts)}）")
        if self.auto_heart_rate.get():
            status.append("心率监测")
        if self.use_template_mode.get():
            status.append("模板模式")

        any_enabled = any([
            self.auto_time.get(),
            self.auto_window.get(),
            self.auto_music.get(),
            self.auto_idle.get(),
            self.auto_hardware.get(),
            self.auto_heart_rate.get()
        ])
        
        if any_enabled:
            self.debug_frame.pack(pady=10, padx=5, fill=tk.X)
            self.start_debug_update()
        else:
            self.debug_frame.pack_forget()
            self.stop_debug_update()

        # 更新开始按钮状态
        self.check_start_button_state()
        
        self.status_var.set(f"已启用：{', '.join(status)}" if status else "未启用附加功能")

    def clear_history(self):
        self.history_list.clear()
        self.history_text.delete(1.0, tk.END)

    def on_close(self):
        self.stop_sending()
        self.stop_debug_update()
        self.heart_rate_monitor.stop()  # 停止心率监测
        
        # 检查并停止网易云同步，防止访问不存在的UI元素
        if self.ncm_sync_running:
            if hasattr(self, 'ncm_stop_event') and self.ncm_stop_event:
                self.ncm_stop_event.set()
            if hasattr(self, 'ncm_sync_thread') and self.ncm_sync_thread and self.ncm_sync_thread.is_alive():
                self.ncm_sync_thread.join(timeout=2)
            self.ncm_sync_running = False
            
            # 只有当设置窗口存在时才更新按钮状态
            try:
                if hasattr(self, 'start_ncm_button'):
                    self.start_ncm_button.config(state=tk.NORMAL)
                if hasattr(self, 'stop_ncm_button'):
                    self.stop_ncm_button.config(state=tk.DISABLED)
            except tk.TclError:
                pass
        
        # 保存配置
        self.save_config()
        
        try:
            self.osc_client.close()
            self.loop.close()
        except:
            pass
        
        self.root.destroy()

    def get_cpu_usage(self):
        try:
            return psutil.cpu_percent(interval=None)
        except:
            return "N/A"

    def get_ram_usage(self):
        try:
            memory = psutil.virtual_memory()
            return memory.percent
        except:
            return "N/A"

    def get_gpu_usage(self):
     try:
         gpus = GPUtil.getGPUs()
         if gpus:
             return f"{gpus[0].load*100:.0f}%"
         else:
             return "无GPU"
     except:
         return "无法获取GPU数据"

if __name__ == "__main__":
    root = tk.Tk()
    app = VRChatAutoChat(root)
    root.mainloop()