"""图形界面入口"""

import asyncio
import tkinter as tk

from ble_heartrate import HeartRateMonitor
from config import Config, SharedState
from pythonosc import udp_client

from .config_panel import ConfigMixin
from .message_logic import MessageMixin
from .settings_panel import SettingsMixin
from .widgets import WidgetsMixin


class VRChatAutoChat(
    ConfigMixin,
    WidgetsMixin,
    SettingsMixin,
    MessageMixin,
):

    def __init__(self, root):
        # 这里集中放窗口级状态，后续各功能区直接读取。
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

        self.order_options = ["时间", "消息内容", "挂机状态", "窗口标题", "心率", "硬件监测", "音乐信息"]

        self.cpu_custom_label = tk.StringVar(value="")
        self.ram_custom_label = tk.StringVar(value="")
        self.gpu_custom_label = tk.StringVar(value="")

        # 字符数限制，保证最终消息不超出发送上限。
        self.window_title_limit = tk.IntVar(value=15)
        self.music_title_limit = tk.IntVar(value=20)
        self.music_artist_limit = tk.IntVar(value=25)

        # 先读配置，再搭界面，避免控件初始值错位。
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
    def update_osc_client(self):
        ip = self.osc_ip.get()
        port = self.osc_port.get()
        self.osc_client = udp_client.SimpleUDPClient(ip, port)
