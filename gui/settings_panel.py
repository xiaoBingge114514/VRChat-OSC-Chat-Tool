"""设置窗口GUI"""

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from netease_sync import launch_netease, netease_thread


class SettingsMixin:
    """Manage settings dialogs and NetEase synchronization controls."""

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
        window_limit_spin = ttk.Spinbox(window_limit_frame, from_=1, to=100, textvariable=self.window_title_limit,
                                        width=10, validate="key",
                                        validatecommand=(self.root.register(self._validate_digits), '%P'))
        window_limit_spin.pack(side=tk.LEFT, padx=5)

        # 音乐标题字符限制
        music_title_limit_frame = ttk.Frame(limit_frame)
        music_title_limit_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(music_title_limit_frame, text="音乐标题最大字符数:").pack(side=tk.LEFT)
        music_title_limit_spin = ttk.Spinbox(music_title_limit_frame, from_=1, to=100,
                                             textvariable=self.music_title_limit, width=10, validate="key",
                                             validatecommand=(self.root.register(self._validate_digits), '%P'))
        music_title_limit_spin.pack(side=tk.LEFT, padx=5)

        # 音乐艺术家字符限制
        music_artist_limit_frame = ttk.Frame(limit_frame)
        music_artist_limit_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(music_artist_limit_frame, text="音乐艺术家最大字符数:").pack(side=tk.LEFT)
        music_artist_limit_spin = ttk.Spinbox(music_artist_limit_frame, from_=1, to=100,
                                              textvariable=self.music_artist_limit, width=10, validate="key",
                                              validatecommand=(self.root.register(self._validate_digits), '%P'))
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
        port_spinbox = ttk.Spinbox(port_frame, from_=1000, to=65535, textvariable=self.ncm_port, width=10,
                                   validate="key", validatecommand=(self.root.register(self._validate_digits), '%P'))
        port_spinbox.pack(side=tk.LEFT, padx=5)
        ttk.Label(port_frame, text="(默认: 9222)").pack(side=tk.LEFT)

        # 进度条设置 - 使用StringVar存储值
        bar_frame = ttk.LabelFrame(advanced_music_frame, text="进度条设置")
        bar_frame.pack(fill=tk.X, padx=10, pady=5)

        width_frame = ttk.Frame(bar_frame)
        width_frame.pack(fill=tk.X, pady=2)
        ttk.Label(width_frame, text="宽度:").pack(side=tk.LEFT)
        self.width_var = tk.StringVar(value=str(self.ncm_config.bar_width))
        self.width_entry = ttk.Entry(width_frame, textvariable=self.width_var, width=5,
                                     validate="key", validatecommand=(self.root.register(self._validate_digits), '%P'))
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
        ttk.Label(template_frame, text="可用变量：{song} {artist} {bar} {time} {lyric1} {lyric2}",
                  foreground="gray").pack(anchor="w", pady=2)
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
            command=self.update_osc_client,
            validate="key",
            validatecommand=(self.root.register(self._validate_digits), '%P')
        )
        port_spinbox.pack(side=tk.LEFT, padx=5)
        ttk.Label(port_frame, text="示例: 9000").pack(side=tk.LEFT)

        save_config_btn = ttk.Button(send_frame, text="保存配置", command=self.save_config)
        save_config_btn.pack(pady=10)

        # 分隔线
        ttk.Separator(send_frame).pack(fill=tk.X, pady=5)

        # 开机自启发送功能
        auto_start_check = ttk.Checkbutton(
            send_frame,
            text="开启软件后自动开启消息发送",
            variable=self.auto_start_enabled,
        )
        auto_start_check.pack(anchor="w", pady=5)

        delay_frame = ttk.Frame(send_frame)
        delay_frame.pack(fill=tk.X, pady=5)
        ttk.Label(delay_frame, text="启动延时(秒):").pack(side=tk.LEFT)
        ttk.Spinbox(
            delay_frame,
            from_=0,
            to=300,
            textvariable=self.auto_start_delay,
            width=10,
            validate="key",
            validatecommand=(self.root.register(self._validate_digits), '%P')
        ).pack(side=tk.LEFT, padx=5)
        ttk.Label(delay_frame, text="(0 = 立即启动)").pack(side=tk.LEFT)

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
                            text=f"播放中: {state.song[:self._safe_int_get(self.music_title_limit, 'music_title_limit', 20)]} - {state.artist[:self._safe_int_get(self.music_artist_limit, 'music_artist_limit', 25)]} - {duration_min}:{duration_sec:02d}",
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
        port = self._safe_int_get(self.ncm_port, 'ncm_port', 9222)
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
        path = filedialog.askopenfilename(title="选择网易云音乐程序",
                                          filetypes=[("Executable files", "*.exe"), ("All files", "*.*")])
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
            if isinstance(widget, ttk.Label) and str(widget) not in [str(child) for child in
                                                                     order_frame.winfo_children() if
                                                                     not isinstance(child, ttk.Label)]:
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
        default_order = ["时间", "消息内容", "挂机状态", "窗口标题", "心率", "硬件监测", "音乐信息"]
        for option in default_order:
            listbox.insert(tk.END, option)

        # 更新order_vars为默认顺序
        for i, option in enumerate(default_order):
            if option in self.order_vars:
                self.order_vars[option].set(str(i))

        # 保存到配置文件
        self.save_config()
        messagebox.showinfo("提示", "已恢复默认排序!")
