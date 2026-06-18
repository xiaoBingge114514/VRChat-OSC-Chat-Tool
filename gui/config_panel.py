"""配置文件逻辑模块"""

import json
import os
import tkinter as tk
from tkinter import messagebox

from config import CONFIG_FILE


class ConfigMixin:
    """Load and save user-facing configuration values."""

    def load_config(self):
        """加载配置文件"""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                # 加载排序设置
                saved_order = config.get('order', [])
                self.order_vars = {option: tk.StringVar(value=str(idx)) for idx, option in
                                   enumerate(self.order_options)}

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
                self.ncm_config.template = config.get('ncm_template',
                                                      '🎵 {song} - {artist}\n{bar} {time}\n{lyric1}\n{lyric2}')

                # 加载模板模式相关设置
                self.use_template_mode.set(config.get('use_template_mode', False))
                self.template_string.set(
                    config.get('template_string', '{message}{time}{window}{idle}{music}{hardware}{heart_rate}'))

                # 加载开机自启发送设置
                self.auto_start_enabled.set(config.get('auto_start_enabled', False))
                self.auto_start_delay.set(config.get('auto_start_delay', 0))

            else:
                # 默认配置
                self.order_vars = {option: tk.StringVar(value=str(idx)) for idx, option in
                                   enumerate(self.order_options)}
        except Exception as e:
            print(f"加载配置失败: {e}")
            # 使用默认值
            self.order_vars = {option: tk.StringVar(value=str(idx)) for idx, option in enumerate(self.order_options)}
    def save_config(self):
        """保存配置到文件"""
        try:
            # 构建配置字典，第一个键专门用于声明文件用途
            config = {
                "_file_description": "VRChat 常驻消息工具配置文件 (v1.4) - 请勿手动修改此处，以免格式错误",
                "_version": "1.4",
                "order": self.get_current_order(),
                'cpu_custom_label': self.cpu_custom_label.get(),
                'ram_custom_label': self.ram_custom_label.get(),
                'gpu_custom_label': self.gpu_custom_label.get(),
                'window_title_limit': self._safe_int_get(self.window_title_limit, 'window_title_limit', 20),
                'music_title_limit': self._safe_int_get(self.music_title_limit, 'music_title_limit', 30),
                'music_artist_limit': self._safe_int_get(self.music_artist_limit, 'music_artist_limit', 30),
                'osc_ip': self.osc_ip.get(),
                'osc_port': self._safe_int_get(self.osc_port, 'osc_port', 9000),
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
                'idle_threshold': self._safe_int_get(self.idle_threshold, 'idle_threshold', 30),
                'advanced_music_enabled': self.advanced_music_enabled.get(),
                'ncm_path': self.ncm_path.get(),
                'ncm_port': self._safe_int_get(self.ncm_port, 'ncm_port', 9222),
                'ncm_bar_width': self.ncm_config.bar_width,
                'ncm_bar_filled': self.ncm_config.bar_filled,
                'ncm_bar_thumb': self.ncm_config.bar_thumb,
                'ncm_bar_empty': self.ncm_config.bar_empty,
                'ncm_template': self.ncm_config.template,
                'use_template_mode': self.use_template_mode.get(),
                'template_string': self.template_string.get(),
                'auto_start_enabled': self.auto_start_enabled.get(),
                'auto_start_delay': self._safe_int_get(self.auto_start_delay, 'auto_start_delay', 0)
            }

            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)

            print("配置已保存")
        except Exception as e:
            print(f"保存配置失败: {e}")
            messagebox.showerror("错误", f"无法保存配置文件:\n{e}")
    def get_current_order(self):
        sorted_items = sorted(self.order_vars.items(), key=lambda x: self._safe_order_int(x[1]))
        return [item[0] for item in sorted_items]
