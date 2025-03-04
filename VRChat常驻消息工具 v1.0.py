"""
VRChat常驻消息自动发送工具 v2.6

功能概述：
1. 通过OSC协议向VRChat发送聊天框消息
2. 支持多种自动附加信息：
   - 当前系统时间
   - 前台窗口标题
   - 正在播放的音乐信息（支持SMTC和网易云音乐）
   - 用户空闲/挂机时间检测
   - 硬件性能监测（CPU/RAM/GPU使用率）
3. 主要特性：
   - 可定期间隔发送（5-300秒）
   - 消息历史记录（最近20条）
   - 实时字符计数和限制（最大144字符）
   - 调试信息面板实时监控
   - 智能粘贴处理
   - 多线程异步处理媒体信息

依赖库说明：
- tkinter: GUI界面构建
- python-osc: OSC协议通信
- win32gui: Windows窗口信息获取
- psutil: 系统资源监控
- GPUtil: GPU信息获取
- winsdk: Windows媒体控制接口

作者：B_小槟
最后更新：2025-3-4
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from pythonosc import udp_client
from datetime import datetime, timedelta
import win32gui
import re
import asyncio
import time
import win32api
import psutil  # 新增: 用于监测CPU和RAM性能
from GPUtil import GPUtil  # 新增: 用于监测GPU性能
from winsdk.windows.media.control import (
    GlobalSystemMediaTransportControlsSessionManager as MediaManager,
    GlobalSystemMediaTransportControlsSessionPlaybackStatus
)

class VRChatAutoChat:
    def __init__(self, root):
        self.root = root
        self.osc_client = udp_client.SimpleUDPClient("127.0.0.1", 9000)
        self.is_sending = False
        self.scheduled_event = None
        self.history_max_items = 20
        self.max_message_length = 144
        self.history_list = []
        self.root.minsize(825, 525)
        self.create_watermark()
        
        # 状态控制变量
        self.auto_time = tk.BooleanVar(value=False)
        self.auto_window = tk.BooleanVar(value=False)
        self.auto_wrap = tk.BooleanVar(value=False)
        self.auto_music = tk.BooleanVar(value=False)
        self.auto_idle = tk.BooleanVar(value=False)
        self.idle_threshold = tk.IntVar(value=60)  # 默认值为60秒 

        # 新增: 硬件性能相关变量
        self.auto_hardware = tk.BooleanVar(value=False)  # 是否启用硬件监测
        self.auto_cpu = tk.BooleanVar(value=False)       # 是否启用CPU监测
        self.auto_ram = tk.BooleanVar(value=False)       # 是否启用RAM监测
        self.auto_gpu = tk.BooleanVar(value=False)       # 是否启用GPU监测

        # 调试相关
        self.debug_update_interval = 1000  # 调试信息更新间隔（毫秒）
        self.debug_update_job = None
        self.debug_labels = {}
        self.last_send_time = 0

        # 初始化异步事件循环
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)

        # GUI 初始化
        self.root.title("VRChat常驻消息工具")
        self.root.geometry("825x525")
        self.create_widgets()
        self.create_watermark()
        self.root.eval('tk::PlaceWindow . center')
        self.root.protocol('WM_DELETE_WINDOW', self.on_close)

    def create_watermark(self):
        bg_color = self.root.cget('bg')
        watermark = ttk.Label(
            self.root,
            text="From B_小槟",
            font=('微软雅黑', 10),
            foreground='#AAAAAA',
            background=bg_color
        )
        watermark.place(relx=1.0, rely=1.0, anchor='se', x=-10, y=-5)

    def create_watermark(self):
        bg_color = self.root.cget('bg')
        self.watermark = ttk.Label(
            self.root,
            text="From B_小槟",
            font=('微软雅黑', 10),
            foreground='#AAAAAA',
            background=bg_color,
            cursor="hand2"  # 添加手型光标
        )
        self.watermark.place(relx=1.0, rely=1.0, anchor='se', x=-10, y=-5)
        # 绑定点击事件
        self.watermark.bind("<Button-1>", self.show_about_window)

    def show_about_window(self, event=None):
        """显示关于窗口"""
        about_win = tk.Toplevel(self.root)
        about_win.title("关于")
        about_win.resizable(False, False)
        about_win.transient(self.root)  # 设为模态窗口
        about_win.grab_set()

        # 窗口内容
        content = [
            ("VRChat常驻消息工具", 14, "bold"),
            ("版本：1.0", 12),
            ("作者：B_小槟", 12),
            ("", 10),
            ("功能特性：", 12, "bold"),
            ("- OSC协议消息发送", 10),
            ("- 历史记录功能", 10),
            ("- 智能附加项（时间/窗口/音乐）", 10),
            ("- 挂机检测机制", 10),
            ("- 硬件性能监测", 10),
            ("", 10),
            ("提示：点击右下角水印显示本窗口", 9, "italic")
        ]

        # 构建界面元素
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

        # 窗口位置居中
        about_win.update_idletasks()
        width = about_win.winfo_width()
        height = about_win.winfo_height()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (width // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (height // 2)
        about_win.geometry(f"+{x}+{y}")

        # 关闭按钮
        close_btn = ttk.Button(
            about_win,
            text="关闭",
            command=about_win.destroy
        )
        close_btn.pack(pady=10, padx=20, ipadx=20)

        # 绑定回车关闭
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

        ttk.Label(control_frame, text="间隔时间（秒）:").pack(side=tk.LEFT)
        self.interval_var = tk.IntVar(value=15)
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
            from_=60,
            to=21600,
            textvariable=self.idle_threshold,
            width=7
        )
        spinbox.pack(side=tk.LEFT)
        self.idle_threshold.trace_add("write", lambda *args: self.update_status())

        # 新增: 硬件性能监测相关控件
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

        ttk.Separator(tool_frame).grid(row=6, column=0, sticky="ew", pady=5)
        
        ttk.Button(
            tool_frame,
            text="清空历史记录",
            command=self.clear_history,
            width=15
        ).grid(row=7, column=0, pady=5)

    def create_debug_panel(self, parent):
        self.debug_frame = ttk.LabelFrame(parent, text="调试信息")
        debug_items = [
            ("window", "当前窗口标题:"),
            ("music_title", "音乐标题:"),
            ("music_artist", "音乐艺术家:"),
            ("idle_time", "挂机时间:"),
            ("next_send", "下次发送:"),
            ("cpu_usage", "CPU 使用率:"),  # 新增
            ("ram_usage", "RAM 使用率:"),  # 新增
            ("gpu_usage", "GPU 使用率:")   # 新增
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
                return f"[在听: {music_info['title']} - {music_info['artist']}]"
        except:
            pass
        
        try:
            hwnd = win32gui.GetForegroundWindow()
            title = win32gui.GetWindowText(hwnd)
            if "网易云音乐" in title:
                match = re.match(r"(.+?)\s*-\s*(.+?)\s*-\s*.+?\s*网易云音乐", title)
                if match:
                    return f"[在听: {match.group(1)} - {match.group(2)}]"
        except:
            pass
        return ""

    def get_idle_duration(self):
        try:
            last_input = win32api.GetLastInputInfo()
            current_tick = win32api.GetTickCount()
            return (current_tick - last_input) // 1000  # 转换为秒
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
            return f"[在看:{title[:20]}]"
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
                hardware_parts.append("CPU: 00%")
            if self.auto_ram.get():
                hardware_parts.append("RAM: 00%")
            if self.auto_gpu.get():
                hardware_parts.append("GPU: 00%")
            if hardware_parts:
                total += len(f"[{', '.join(hardware_parts)}]")
        return total

    def process_message(self, raw_message):
        additions = []
        wrap_char = "\n" if self.auto_wrap.get() else " "
    
        # 挂机检测附加项
        if self.auto_idle.get():
            idle_sec = self.get_idle_duration()
            if idle_sec >= self.idle_threshold.get():
                additions.append(f"[已挂机: {self.format_duration(idle_sec)}]")
    
        # 时间附加项
        if self.auto_time.get():
            additions.append(self.get_formatted_time())
    
        # 窗口标题附加项
        if self.auto_window.get():
            window_title = self.get_formatted_window_title()
            if window_title:
                additions.append(window_title)
    
        # 音乐信息附加项
        if self.auto_music.get():
            music_info = self.get_formatted_music_info()
            if music_info:
                additions.append(music_info)
    
        # 硬件性能附加项
        hardware_parts = []
        if self.auto_hardware.get():
            if self.auto_cpu.get():
                cpu_usage = self.get_cpu_usage()
                hardware_parts.append(f"CPU: {cpu_usage:.0f}%")
            if self.auto_ram.get():
                ram_usage = self.get_ram_usage()
                hardware_parts.append(f"RAM: {ram_usage:.0f}%")
            if self.auto_gpu.get():
                gpu_usage = self.get_gpu_usage()
                hardware_parts.append(f"GPU: {gpu_usage}" if gpu_usage else "GPU: 无法获取数据")
        
            if hardware_parts:
                additions.append(f"[{', '.join(hardware_parts)}]")
    
        # 合并所有附加项内容
        combined_additions = wrap_char.join(additions)
        user_message = raw_message.rstrip('\n')  # 去除用户输入末尾的换行符
        full_message = f"{user_message}\n{combined_additions}"  # 确保用户内容与附加项之间有一个换行符
        return full_message.strip()  # 去除整体消息的首尾空白字符

    def send_message(self):
        raw_message = self.text_input.get("1.0", "end-1c").rstrip('\n')  # 移除消息末尾的换行符
        if not raw_message.strip():
            return False

        try:
            final_message = self.process_message(raw_message)
            self.osc_client.send_message("/chatbox/input", [final_message, True])
            
            history_msg = raw_message.strip()
            if self.auto_time.get():
                history_msg += f"\n{self.get_formatted_time()}"
            if self.auto_window.get():
                history_msg += f"\n{self.get_formatted_window_title()}"
            if self.auto_music.get():
                music_info = self.get_formatted_music_info()
                if music_info:
                    history_msg += f"\n{music_info}"
            if self.auto_idle.get() and self.get_idle_duration() >= self.idle_threshold.get():
                history_msg += f"\n[已挂机: {self.format_duration(self.get_idle_duration())}]"
            # 添加硬件性能信息到历史记录
            if self.auto_hardware.get():
                hardware_parts = []
                if self.auto_cpu.get():
                    cpu_usage = self.get_cpu_usage()
                    hardware_parts.append(f"CPU: {cpu_usage:.0f}%")
                if self.auto_ram.get():
                    ram_usage = self.get_ram_usage()
                    hardware_parts.append(f"RAM: {ram_usage:.0f}%")
                if self.auto_gpu.get():
                    gpu_usage = self.get_gpu_usage()
                    hardware_parts.append(f"GPU: {gpu_usage}" if gpu_usage else "GPU: 无法获取数据")
                if hardware_parts:
                    history_msg += f"\n[{', '.join(hardware_parts)}]"
            
            self.send_to_history(history_msg)
            self.update_char_count()
            self.last_send_time = time.time()
            return True
        except Exception as e:
            messagebox.showerror("错误", f"消息发送失败: {str(e)}")
            return False

    def send_to_history(self, message):
        current_time = datetime.now().strftime('%H:%M:%S')
        formatted_message = f"[{len(self.history_list)+1}] ({current_time}):\n"
        
        if self.auto_wrap.get():
            formatted_message += message.replace("\n", "\n│ ")
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
            # 更新原有调试信息
            title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            self.debug_labels['window'].config(text=title[:40] if title else "无")
            
            music_info = self.get_raw_music_info()
            self.debug_labels['music_title'].config(text=music_info['title'][:30] if music_info else "无")
            self.debug_labels['music_artist'].config(text=music_info['artist'][:30] if music_info else "无")
            
            idle_sec = self.get_idle_duration()
            self.debug_labels['idle_time'].config(text=self.format_duration(idle_sec))
            
            if self.is_sending:
                remaining = self.interval_var.get() - (time.time() - self.last_send_time)
                self.debug_labels['next_send'].config(text=f"{max(0, int(remaining))}秒")
            else:
                self.debug_labels['next_send'].config(text="未启用")

            # 更新硬件性能调试信息
            cpu_message = "未启用功能"
            ram_message = "未启用功能"
            gpu_message = "未启用功能"

            if self.auto_hardware.get():  # 只有在启用硬件监测时才获取数据
                cpu_usage = self.get_cpu_usage()
                if cpu_usage != "N/A":
                    cpu_message = f"{cpu_usage:.0f}%"
                else:
                    cpu_message = cpu_usage

                ram_usage = self.get_ram_usage()
                if ram_usage != "N/A":
                    ram_message = f"{ram_usage:.0f}%"
                else:
                    ram_message = ram_usage

                gpu_usage = self.get_gpu_usage()
                if gpu_usage != "无法获取数据":
                    gpu_message = gpu_usage
                else:
                    gpu_message = f"未启用功能，{gpu_usage}"
            else:
                # 如果硬件监测未启用，但在调试信息中仍显示其他硬件信息
                pass  # 或者根据需要显示其他提示

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

        any_enabled = any([
            self.auto_time.get(),
            self.auto_window.get(),
            self.auto_music.get(),
            self.auto_idle.get(),
            self.auto_hardware.get()
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
        try:
            self.osc_client.close()
            self.loop.close()
        except:
            pass
        self.root.destroy()

    # 新增: 获取硬件性能数据的函数
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