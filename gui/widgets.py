"""主界面GUI组件"""

import tkinter as tk
from tkinter import scrolledtext, ttk

class WidgetsMixin:

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
            ("版本：1.4", 12),
            ("作者：", 12),
            ("VRChat: B_小槟 ", 12, ),
            ("Github: Xiaobingge114514 ", 12, ),
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
            ("- 启动软件自动发送功能", 10),
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
            label.pack(padx=20, pady=2 if size < 12 else 5, fill=tk.X)

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
        # 左侧放输入和历史，右侧放自动附加项与调试信息。
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
        # 消息输入区、发送按钮和历史记录区。
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
            width=5,
            validate="key",
            validatecommand=(self.root.register(self._validate_digits), '%P')
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

        # 开机自动启动倒计时提示
        ttk.Label(
            control_frame,
            textvariable=self.auto_start_countdown_var,
            foreground="#0066CC"
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
        # 自动附加项区：控制时间、窗口、音乐、硬件等拼接内容。
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
            width=7,
            validate="key",
            validatecommand=(self.root.register(self._validate_digits), '%P')
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
    def toggle_heart_rate(self):
        """切换心率检测状态"""
        if self.auto_heart_rate.get():
            self.heart_rate_monitor.start()
        else:
            self.heart_rate_monitor.stop()
        self.update_status()
    def create_debug_panel(self, parent):
        # 调试面板只读展示当前状态，方便排查拼接结果。
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
