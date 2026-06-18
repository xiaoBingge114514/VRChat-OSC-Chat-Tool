"""网易云音乐 WebSocket/调试端口同步逻辑。"""

import asyncio
import contextlib
import glob
import json
import os
import re
import socket
import subprocess
import threading
import time

import requests
import websockets
from websockets import protocol

from config import Config, SharedState


HEADERS = {"User-Agent": "Mozilla/5.0", "Referer": "https://music.163.com/"}


def _hide_window():
    """返回 startupinfo，隐藏子进程的控制台窗口（PyInstaller 无控制台模式必需）。"""
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    return startupinfo


def fetch_lyrics(song, artist):
    # 先搜歌，再取歌词并解析成时间轴列表。
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
    for pattern in patterns:
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
    # 以远程调试模式启动网易云客户端。
    exe = path if path and os.path.exists(path) else find_netease()
    if not exe:
        return False, "未找到网易云", None
    if port is None:
        port = find_free_port()
    subprocess.Popen(
        [
            exe,
            f"--remote-debugging-port={port}",
            "--disable-background-timer-throttling",
            "--disable-backgrounding-occluded-windows",
            "--disable-renderer-backgrounding",
        ],
        startupinfo=_hide_window(),
    )
    return True, exe, port


JS_GET_STATE = r"""(() => {
    try {
        let r = { song: '', artist: '', cur: 0, dur: 0, play: false, lyric1: '', lyric2: '' };

        let songEl = document.querySelector('.cmd-space.title span')
            || document.querySelector('.main-title')
            || document.querySelector('.two-line .title')
            || document.querySelector('[class*="title"] span');
        r.song = songEl?.innerText?.trim() || songEl?.textContent?.trim() || '';

        let artist = document.querySelector('.author');
        r.artist = artist?.innerText?.trim() || '';
        if (!r.artist) {
            artist = document.querySelector('.info.artist');
            r.artist = (artist?.innerText || '').replace(/^歌手[：:]/, '').trim();
        }

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

        if (!r.dur) {
            let timeEl = document.querySelector('.curtime-thumb');
            if (timeEl?.innerText) {
                let m = timeEl.innerText.match(/(\d+):(\d+)\s*\/\s*(\d+):(\d+)/);
                if (m) { r.cur = +m[1] * 60 + +m[2]; r.dur = +m[3] * 60 + +m[4]; }
            }
        }

        r.play = !!document.querySelector('[class*="cmd-icon-pause"]')
            || !!document.querySelector('[title*="暂停"]');

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


class NeteaseSync:
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.ws = None
        self.msg_id = 0

    async def connect(self):
        pages = requests.get(f"http://127.0.0.1:{self.cfg.ncm_port}/json", timeout=2).json()
        self.ws = await websockets.connect(
            pages[0]["webSocketDebuggerUrl"], ping_interval=30, ping_timeout=15
        )

    async def eval_js(self, code, timeout=1):
        # 通过 CDP 让网页在当前播放页执行 JS 并返回结果。
        if not self.ws:
            return None

        self.msg_id += 1
        bring_id = self.msg_id
        await self.ws.send(json.dumps({"id": bring_id, "method": "Page.bringToFront"}))

        try:
            while True:
                msg = await asyncio.wait_for(self.ws.recv(), timeout=0.5)
                d = json.loads(msg)
                if d.get("id") == bring_id:
                    break
        except asyncio.TimeoutError:
            pass

        self.msg_id += 1
        msg_id = self.msg_id
        await self.ws.send(
            json.dumps(
                {
                    "id": msg_id,
                    "method": "Runtime.evaluate",
                    "params": {"expression": code, "returnByValue": True},
                }
            )
        )

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


def netease_thread(
    cfg: Config, shared: SharedState, stop_event: threading.Event, cb: CallbackProtocol
):
    async def run():
        # 保持连接、抓取歌曲状态，并在新歌时异步拉歌词。
        cb.cb_status("连接中...")
        sync = NeteaseSync(cfg)
        retry_count = 0
        max_retries = 3
        while not sync.ws and retry_count < max_retries:
            try:
                await sync.connect()
            except Exception as e:
                retry_count += 1
                cb.cb_status(f"连接失败，重试中... {retry_count}/{max_retries}\n{e}")
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
                            daemon=True,
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
