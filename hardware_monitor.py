"""Windows 硬件、窗口和媒体信息采集工具。"""

import asyncio
import re
from datetime import datetime, timedelta

import psutil
import win32api
import win32gui
from GPUtil import GPUtil
from winsdk.windows.media.control import (
    GlobalSystemMediaTransportControlsSessionManager as MediaManager,
    GlobalSystemMediaTransportControlsSessionPlaybackStatus,
)


def get_cpu_usage():
    try:
        return psutil.cpu_percent(interval=None)
    except Exception:
        return "N/A"


def get_ram_usage():
    try:
        memory = psutil.virtual_memory()
        return memory.percent
    except Exception:
        return "N/A"


def get_gpu_usage():
    try:
        gpus = GPUtil.getGPUs()
        if gpus:
            return f"{gpus[0].load * 100:.0f}%"
        return "无GPU"
    except Exception:
        return "无法获取GPU数据"


def get_idle_duration():
    try:
        last_input = win32api.GetLastInputInfo()
        current_tick = win32api.GetTickCount()
        return (current_tick - last_input) // 1000
    except Exception as e:
        print(f"获取空闲时间失败: {e}")
        return 0


def format_duration(seconds):
    if seconds < 60:
        return f"{seconds}秒"
    minutes, sec = divmod(seconds, 60)
    return f"{minutes}分{sec}秒"


def get_formatted_time():
    return (datetime.utcnow() + timedelta(hours=8)).strftime("[时间:%H:%M]")


def get_window_title(limit=20):
    try:
        title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        return title[:limit] if title else ""
    except Exception:
        return ""


def get_formatted_window_title(limit=20):
    title = get_window_title(limit)
    return f"[在看:{title}]" if title else ""


async def get_media_info_async():
    # 优先走系统 SMTC，拿不到再回退到窗口标题解析。
    try:
        sessions = await MediaManager.request_async()
        current_session = sessions.get_current_session()

        if (
            current_session
            and current_session.get_playback_info().playback_status
            == GlobalSystemMediaTransportControlsSessionPlaybackStatus.PLAYING
        ):
            media_properties = await current_session.try_get_media_properties_async()
            return {
                "title": media_properties.title or "未知曲目",
                "artist": media_properties.artist or "未知艺术家",
            }
    except Exception as e:
        print(f"SMTC 获取失败: {e}")
    return None


def get_raw_music_info(loop=None):
    try:
        if loop is None:
            music_info = asyncio.run(get_media_info_async())
        else:
            music_info = loop.run_until_complete(get_media_info_async())
        if music_info:
            return music_info

        title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        if "网易云音乐" in title:
            match = re.match(r"(.+?)\s*-\s*(.+?)\s*-\s*.+?\s*网易云音乐", title)
            if match:
                return {"title": match.group(1), "artist": match.group(2)}
    except Exception:
        pass
    return None


def get_formatted_music_info(title_limit=30, artist_limit=30, loop=None):
    music_info = get_raw_music_info(loop=loop)
    if music_info:
        title = music_info["title"][:title_limit]
        artist = music_info["artist"][:artist_limit]
        return f"[在听: {title} - {artist}]"
    return ""


def build_hardware_parts(
    include_cpu=False,
    include_ram=False,
    include_gpu=False,
    cpu_label="",
    ram_label="",
    gpu_label="",
):
    # 按开关收集 CPU / RAM / GPU 文本片段。
    parts = []
    if include_cpu:
        cpu_usage = get_cpu_usage()
        if cpu_usage != "N/A":
            parts.append(
                f"CPU({cpu_label}): {cpu_usage:.0f}%"
                if cpu_label
                else f"CPU: {cpu_usage:.0f}%"
            )
    if include_ram:
        ram_usage = get_ram_usage()
        if ram_usage != "N/A":
            parts.append(
                f"RAM({ram_label}): {ram_usage:.0f}%"
                if ram_label
                else f"RAM: {ram_usage:.0f}%"
            )
    if include_gpu:
        gpu_usage = get_gpu_usage()
        if gpu_usage and gpu_usage != "无法获取GPU数据":
            parts.append(f"GPU({gpu_label}): {gpu_usage}" if gpu_label else f"GPU: {gpu_usage}")
    return parts


def get_formatted_hardware_info(**kwargs):
    parts = build_hardware_parts(**kwargs)
    return f"[{', '.join(parts)}]" if parts else ""
