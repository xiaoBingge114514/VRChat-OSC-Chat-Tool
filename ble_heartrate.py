"""蓝牙 BLE 心率监测线程封装。"""

import asyncio
import struct
import threading
from queue import Queue

from bleak import BleakClient, BleakScanner

from config import HRM_UUID, HRS_UUID


class HeartRateMonitor:
    """BLE heart-rate monitor that runs outside the Tkinter UI thread."""

    def __init__(self):
        self.current_hr = 0
        self.is_connected = False
        self.device_name = None
        self._running = False
        self._thread = None
        self._queue = Queue()

    def start(self):
        if not self._running:
            self._running = True
            self._thread = threading.Thread(target=self._run_async_loop, daemon=True)
            self._thread.start()

    def stop(self):
        self._running = False
        self.is_connected = False
        self.current_hr = 0

    def _run_async_loop(self):
        # 单独开事件循环，避免阻塞 Tkinter 主线程。
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(self._main_loop())

    async def _main_loop(self):
        while self._running:
            try:
                await self._scan_and_connect()
            except Exception as e:
                print(f"心率监测错误: {e}")
                await asyncio.sleep(3)

    async def _scan_and_connect(self):
        # 先扫描设备，连接后订阅心率通知。
        self.is_connected = False
        self.device_name = None

        device = await BleakScanner.find_device_by_filter(
            lambda d, ad: ad.service_uuids
            and HRS_UUID.lower() in [u.lower() for u in ad.service_uuids],
            timeout=10.0,
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

                hrm = next(
                    (c for c in hrs.characteristics if c.uuid.lower() == HRM_UUID), None
                )
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
        # 按 HRS 协议解析 8 位或 16 位心率值。
        flags = data[0]
        if flags & 0x01:
            return struct.unpack_from("<H", data, 1)[0]
        return data[1]
