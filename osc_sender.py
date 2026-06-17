"""OSC 发送线程与消息格式化逻辑。"""

import threading
import time

from pythonosc import udp_client

from config import Config, SharedState
from netease_sync import CallbackProtocol


def get_lyric(lyrics, pos):
    # 用二分查找定位当前时间点对应的歌词行。
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
    # 统一拼装最终发给 VRChat 的文本。
    c, d, w = state.cur, state.dur, cfg.bar_width
    pos = int(w * c / d) if d else 0
    thumb = cfg.bar_thumb
    if thumb:
        bar = cfg.bar_filled * pos + thumb + cfg.bar_empty * (w - pos)
    else:
        bar = cfg.bar_filled * pos + cfg.bar_empty * (w - pos)

    l1, l2 = state.lyric1, state.lyric2
    if not l1 and lyrics and song_key == f"{state.song}-{state.artist}":
        l1, l2 = get_lyric(lyrics, c)
    l1, l2 = l1 or "纯音乐，请欣赏", l2 or ""

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


def osc_thread(
    cfg: Config, shared: SharedState, stop_event: threading.Event, cb: CallbackProtocol
):
    # 后台轮询共享状态，并按节流间隔发送 OSC。
    osc = None
    last_osc = 0
    while not stop_event.is_set():
        with shared.lock:
            state = shared.data.copy()
            lyrics = list(shared.lyrics)
            song_key = shared.song_key
            last_update = shared.last_update

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
