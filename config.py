"""全局配置、共享状态与歌曲/歌词状态模型。"""

import threading

from pydantic import BaseModel


CONFIG_FILE = "vrchat_config.json"

# BLE 服务与特征 UUID，供心率模块直接复用。
HRS_UUID = "0000180d-0000-1000-8000-00805f9b34fb"
HRM_UUID = "00002a37-0000-1000-8000-00805f9b34fb"


class Config(BaseModel):
    osc_ip: str = "127.0.0.1"
    osc_port: int = 9000
    ncm_port: int = 9222
    ncm_path: str = ""
    refresh_interval: float = 3.0
    bar_width: int = 8
    bar_filled: str = "▓"
    bar_empty: str = "░"
    bar_thumb: str = "◘"
    template: str = "🎵 {song} - {artist}\n{bar} {time}\n{lyric1}\n{lyric2}"


class SongState:
    """当前歌曲播放状态的轻量对象。"""

    def __init__(
        self, song="", artist="", cur=0, dur=0, play=False, lyric1="", lyric2=""
    ):
        self.song = song
        self.artist = artist
        self.cur = cur
        self.dur = dur
        self.play = play
        self.lyric1 = lyric1
        self.lyric2 = lyric2

    def update(self, d: dict):
        for k in ["song", "artist", "cur", "dur", "play", "lyric1", "lyric2"]:
            if k in d:
                setattr(self, k, d[k])

    def copy(self):
        return SongState(
            song=self.song,
            artist=self.artist,
            cur=self.cur,
            dur=self.dur,
            play=self.play,
            lyric1=self.lyric1,
            lyric2=self.lyric2,
        )


class SharedState:
    """跨线程共享的数据容器，配合锁使用。"""

    def __init__(self):
        self.lock = threading.Lock()
        self.data = SongState()
        self.song_key = ""
        self.lyrics = []
        self.last_update = 0
