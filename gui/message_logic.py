"""消息逻辑模块"""

import re
import time
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import messagebox

import psutil
import win32api
import win32gui
from GPUtil import GPUtil
from osc_sender import format_output
from winsdk.windows.media.control import (
    GlobalSystemMediaTransportControlsSessionManager as MediaManager,
    GlobalSystemMediaTransportControlsSessionPlaybackStatus,
)


class MessageMixin:
    """Compose outgoing messages and drive the sending lifecycle."""

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
        # 优先用高级音乐同步，失败再回退到系统媒体信息。
        # 如果启用了高级音乐信息，则返回高级信息
        if self.advanced_music_enabled.get() and self.ncm_sync_running:
            with self.ncm_shared_state.lock:
                state = self.ncm_shared_state.data.copy()
                lyrics = list(self.ncm_shared_state.lyrics)
                song_key = self.ncm_shared_state.song_key
            formatted_output = format_output(self.ncm_config, state, lyrics, song_key, self.music_title_limit.get(),
                                             self.music_artist_limit.get())
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
        # 预估附加内容长度，用于输入框右上角字数提示。
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
        # 把用户输入和自动附加项组合成最终发送文本。
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
        # 发送前先校验，再构造文本并写入历史。
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
        # 历史记录尽量保持和实际发送内容一致，方便回看。
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
        formatted_message = f"[{len(self.history_list) + 1}] ({current_time}):\n"

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
        # 首次发送成功后，切进定时发送循环。
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
        # 定时刷新调试区，显示当前拼接状态。
        try:
            title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            self.debug_labels['window'].config(text=title[:self.window_title_limit.get()] if title else "无")

            music_info = self.get_raw_music_info()
            self.debug_labels['music_title'].config(
                text=music_info['title'][:self.music_title_limit.get()] if music_info else "无")
            self.debug_labels['music_artist'].config(
                text=music_info['artist'][:self.music_artist_limit.get()] if music_info else "无")

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
        # 退出前先停线程、停定时器，再销毁窗口。
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
                return f"{gpus[0].load * 100:.0f}%"
            else:
                return "无GPU"
        except:
            return "无法获取GPU数据"
