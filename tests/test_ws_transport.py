"""Проверка WebSocket-транспорта на локальном сервере, который понимает obfuscated2 + intermediate."""
import asyncio
import logging
import struct
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

try:
    import aiohttp
    from aiohttp import web
except ImportError:  # pragma: no cover
    aiohttp = None

from telethon.crypto import AESModeCTR  # noqa: E402

import tg_ws_transport as wst  # noqa: E402


async def fake_telegram(request):
    """Расшифровывает заголовок obfuscated2, принимает пакеты intermediate и отвечает эхом."""
    ws = web.WebSocketResponse(protocols=("binary",))
    await ws.prepare(request)
    buf = bytearray()
    dec = enc = None
    async for msg in ws:
        buf += msg.data
        if dec is None:
            if len(buf) < 64:
                continue
            init = bytes(buf[:64])
            del buf[:64]
            rev = init[55:7:-1]
            dec = AESModeCTR(init[8:40], init[40:56])          # ключ клиента на отправку
            enc = AESModeCTR(bytes(rev[:32]), bytes(rev[32:48]))  # ключ клиента на приём
            tag = dec.encrypt(init)[56:60]
            request.app["tags"].append(tag)
            buf = bytearray(dec.encrypt(bytes(buf)))
        else:
            buf = buf[:-len(msg.data)] + bytearray(dec.encrypt(msg.data))
        while len(buf) >= 4:
            (length,) = struct.unpack("<i", buf[:4])
            if len(buf) < 4 + length:
                break
            payload = bytes(buf[4:4 + length])
            del buf[:4 + length]
            request.app["got"].append(payload)
            await ws.send_bytes(enc.encrypt(struct.pack("<i", len(payload)) + payload[::-1]))
    return ws


@unittest.skipIf(aiohttp is None, "aiohttp не установлен")
class WebSocketTransportTests(unittest.TestCase):
    def test_urls(self):
        self.assertEqual(wst.ws_url(2), "wss://venus.web.telegram.org/apiws")
        self.assertEqual(wst.ws_url(4, test=True), "wss://vesta.web.telegram.org/apiws_test")
        self.assertIn("flora.web.telegram.org", wst.WS_HOSTS)
        with self.assertRaises(ValueError):
            wst.ws_url(9)

    def test_roundtrip_through_obfuscated_websocket(self):
        async def run():
            app = web.Application()
            app["tags"], app["got"] = [], []
            app.router.add_get("/apiws", fake_telegram)
            runner = web.AppRunner(app)
            await runner.setup()
            site = web.TCPSite(runner, "127.0.0.1", 0)
            await site.start()
            port = site._server.sockets[0].getsockname()[1]

            orig_url, orig_proxy = wst.ws_url, wst._proxy_url
            wst.ws_url = lambda dc_id, test=False: f"ws://127.0.0.1:{port}/apiws"
            wst._proxy_url = lambda: None
            try:
                log = logging.getLogger("t")
                conn = wst.ConnectionWebSocket("0.0.0.0", 443, 2,
                                               loggers={"telethon.network.connection.connection": log})
                await conn.connect(timeout=5)
                replies = []
                for payload in (b"A" * 40, b"hello-mtproto-1234", bytes(range(256)) * 5):
                    await conn.send(payload)
                    replies.append(await asyncio.wait_for(conn.recv(), 5))
                await conn.disconnect()
            finally:
                wst.ws_url, wst._proxy_url = orig_url, orig_proxy
                await runner.cleanup()
            return app, replies

        app, replies = asyncio.run(run())
        self.assertEqual(app["tags"], [b"\xee\xee\xee\xee"])  # intermediate внутри obfuscated2
        self.assertEqual(app["got"], [b"A" * 40, b"hello-mtproto-1234", bytes(range(256)) * 5])
        self.assertEqual(replies, [p[::-1] for p in app["got"]])


if __name__ == "__main__":
    unittest.main()
