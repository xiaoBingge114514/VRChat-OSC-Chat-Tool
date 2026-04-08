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
        self.root.minsize(880, 570)
        
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

        self.debug_update_interval = 1000
        self.debug_update_job = None
        self.debug_labels = {}
        self.last_send_time = 0


        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)

        # 添加设置相关变量
        self.order_options = ["消息内容", "时间", "挂机状态", "窗口标题", "心率", "音乐信息", "硬件监测"]
        
        self.cpu_custom_label = tk.StringVar(value="")
        self.ram_custom_label = tk.StringVar(value="")
        self.gpu_custom_label = tk.StringVar(value="")
        
        # 字符数限制
        self.window_title_limit = tk.IntVar(value=20)
        self.music_title_limit = tk.IntVar(value=30)
        self.music_artist_limit = tk.IntVar(value=30)

        # 初始化顺序变量，从配置加载或使用默认顺序
        self.load_config()

        self.root.title("VRChat常驻消息工具")
        self.root.geometry("800x600")
        self.create_widgets()
        self.create_watermark()
        # 修改居中逻辑
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (800 // 2)
        y = (self.root.winfo_screenheight() // 2) - (600 // 2)
        self.root.geometry(f"800x600+{x}+{y}")
        self.root.protocol('WM_DELETE_WINDOW', self.on_close)
        self.update_status()

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
                "_file_description": "VRChat 常驻消息工具配置文件 (v1.2) - 请勿手动修改此处，以免格式错误",
                "_version": "1.2", 
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
                'idle_threshold': self.idle_threshold.get()
            }

            with open('vrchat_config.json', 'w', encoding='utf-8') as f:
                # indent=4 让文件排版好看，ensure_ascii=False 支持中文
                json.dump(config, f, indent=4, ensure_ascii=False)
            
            print("配置已保存")
        except Exception as e:
            print(f"保存配置失败: {e}")
            messagebox.showerror("错误", f"无法保存配置文件:\n{e}")

    def get_current_order(self):
        """获取当前排序"""
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
        """显示关于窗口"""
        about_win = tk.Toplevel(self.root)
        about_win.title("关于")
        about_win.resizable(False, False)
        about_win.transient(self.root)
        about_win.grab_set()

        content = [
            ("VRChat常驻消息工具", 14, "bold"),
            ("版本：1.2", 12),
            ("作者：(VRC)B_小槟", 12),
            ("", 10),
            ("功能特性：", 12, "bold"),
            ("- OSC协议消息发送", 10),
            ("- 历史记录功能", 10),
            ("- 智能附加项（时间/窗口/音乐/心率）", 10),
            ("- 挂机检测机制", 10),
            ("- 硬件性能监测 (CPU/RAM/GPU)", 10),
            ("- 蓝牙心率监测（基于HRS协议）", 10),
            # 新增的功能
            ("- 硬件监测标签自定义显示", 10),
            ("- 窗口标题/音乐信息字符数限制", 10),
            ("- 发送顺序自定义与配置持久保存", 10),
            ("- 一键恢复默认排序设置", 10),
            ("- 自定义OSC发送IP地址和端口", 10),
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
            font=('Arial', 10)
        )
        self.text_input.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)
        self.text_input.bind("<KeyRelease>", self.update_char_count)
        self.text_input.bind("<<Paste>>", self.on_paste)
        self.update_char_count()

        control_frame = ttk.Frame(parent)
        control_frame.pack(pady=5, fill=tk.X)

        # OSC端口配置移到设置窗口
        ttk.Label(control_frame, text="间隔时间（秒）:").pack(side=tk.LEFT)
        self.interval_var = tk.IntVar(value=10)
        ttk.Spinbox(
            control_frame,
            from_=5,
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
            height=10,
            wrap=tk.WORD,
            font=('Arial', 10)
        )
        scrollbar = ttk.Scrollbar(history_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.history_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.history_text.yview)

    def update_osc_client(self):
        """更新OSC客户端端口"""
        ip = self.osc_ip.get()
        port = self.osc_port.get()
        self.osc_client = udp_client.SimpleUDPClient(ip, port)

    def create_right_panel(self, parent):
        tool_frame = ttk.LabelFrame(parent, text="自动附加项")
        tool_frame.pack(pady=10, padx=5, fill=tk.X)

        ttk.Checkbutton(
            tool_frame,
            text="附加项自动换行",
            variable=self.auto_wrap,
            command=self.update_status
        ).grid(row=0, column=0, sticky="w", pady=2)

        ttk.Checkbutton(
            tool_frame,
            text="添加当前时间",
            variable=self.auto_time,
            command=self.update_status
        ).grid(row=1, column=0, sticky="w", pady=2)

        ttk.Checkbutton(
            tool_frame,
            text="添加窗口标题",
            variable=self.auto_window,
            command=self.update_status
        ).grid(row=2, column=0, sticky="w", pady=2)

        ttk.Checkbutton(
            tool_frame,
            text="添加音乐信息 (SMTC)",
            variable=self.auto_music,
            command=self.update_status
        ).grid(row=3, column=0, sticky="w", pady=2)

        idle_frame = ttk.Frame(tool_frame)
        idle_frame.grid(row=4, column=0, sticky="w", pady=2)
        
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
        hardware_frame.grid(row=5, column=0, sticky="w", pady=2)
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
        heart_rate_frame.grid(row=6, column=0, sticky="w", pady=2)
        ttk.Checkbutton(
            heart_rate_frame,
            text="心率检测",
            variable=self.auto_heart_rate,
            command=self.toggle_heart_rate
        ).pack(side=tk.LEFT)
        ttk.Label(heart_rate_frame, text="(需蓝牙手环/手表)", foreground="gray").pack(side=tk.LEFT, padx=5)

        ttk.Separator(tool_frame).grid(row=7, column=0, sticky="ew", pady=5)
        
        button_frame = ttk.Frame(tool_frame)
        button_frame.grid(row=8, column=0, pady=5)
        
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

    def open_settings(self):
        """打开设置窗口"""
        settings_win = tk.Toplevel(self.root)
        settings_win.title("设置")
        settings_win.geometry("400x450")
        settings_win.resizable(False, False)
        settings_win.transient(self.root)
        settings_win.grab_set()

        # 创建笔记本控件用于多标签页
        notebook = ttk.Notebook(settings_win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 排序设置标签页
        order_frame = ttk.Frame(notebook)
        notebook.add(order_frame, text="发送顺序")

        ttk.Label(order_frame, text="调整附加项与消息内容的发送顺序:", font=("Arial", 10, "bold")).pack(pady=5)
        
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

        # 上移按钮
        move_up_btn = ttk.Button(order_frame, text="上移", command=lambda: self.move_item(order_listbox, -1))
        move_up_btn.pack(side=tk.LEFT, padx=5)

        # 下移按钮
        move_down_btn = ttk.Button(order_frame, text="下移", command=lambda: self.move_item(order_listbox, 1))
        move_down_btn.pack(side=tk.LEFT, padx=5)

        # 保存排序按钮
        save_order_btn = ttk.Button(order_frame, text="保存排序", command=lambda: self.save_order(order_listbox))
        save_order_btn.pack(side=tk.LEFT, padx=5)

        # 重置按钮
        reset_order_btn = ttk.Button(order_frame, text="恢复默认", command=lambda: self.reset_order(order_listbox))
        reset_order_btn.pack(side=tk.LEFT, padx=5)

        # 硬件标签设置标签页
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

        # 发送设置标签页
        send_frame = ttk.Frame(notebook)
        notebook.add(send_frame, text="发送设置")

        ttk.Label(send_frame, text="OSC发送设置:", font=("Arial", 10, "bold")).pack(pady=5)
        
        # IP地址设置
        ip_frame = ttk.Frame(send_frame)
        ip_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(ip_frame, text="OSC IP地址:").pack(side=tk.LEFT)
        ip_entry = ttk.Entry(ip_frame, textvariable=self.osc_ip, width=20)
        ip_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(ip_frame, text="示例: 127.0.0.1").pack(side=tk.LEFT)

        # 端口设置
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

        # 保存配置按钮
        save_config_btn = ttk.Button(send_frame, text="保存配置", command=self.save_config)
        save_config_btn.pack(pady=10)

        # 关闭按钮
        close_btn = ttk.Button(settings_win, text="关闭", command=settings_win.destroy)
        close_btn.pack(pady=10)

        # 居中显示设置窗口
        settings_win.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (settings_win.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (settings_win.winfo_height() // 2)
        settings_win.geometry(f"+{x}+{y}")

    def move_item(self, listbox, direction):
        """移动列表项"""
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
        """重置排序到默认顺序"""
        # 清空列表框
        listbox.delete(0, tk.END)
        # 添加默认顺序
        default_order = ["消息内容", "时间", "挂机状态", "窗口标题", "音乐信息", "心率", "硬件监测"]
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
                # 修复后的硬件标签逻辑
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
        if not raw_message.strip():
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
            if not self.text_input.get("1.0", "end-1c").strip():
                messagebox.showwarning("提示", "消息内容不能为空！")
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
        self.send_message()
        self.scheduled_event = self.root.after(interval * 1000, self.scheduled_send_status)
        self.update_countdown(interval)
        self.start_debug_update()

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
            if self.send_message():
                self.scheduled_event = self.root.after(interval * 1000, self.scheduled_send_status)
                self.update_countdown(interval)
            else:
                self.stop_sending()

    def update_status(self):
        status = []
        if self.auto_time.get():
            status.append("时间")
        if self.auto_window.get():
            status.append("窗口标题")
        if self.auto_wrap.get():
            status.append("换行")
        if self.auto_music.get():
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

        self.status_var.set(f"已启用：{', '.join(status)}" if status else "未启用附加功能")

    def clear_history(self):
        self.history_list.clear()
        self.history_text.delete(1.0, tk.END)

    def on_close(self):
        self.stop_sending()
        self.stop_debug_update()
        self.heart_rate_monitor.stop()  # 停止心率监测
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
