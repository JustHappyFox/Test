"""
MTProto через WebSocket: wss://<dc>.web.telegram.org/apiws — так подключается Telegram Web.

Нужен там, где наружу пускают только HTTPS через прокси, а прямые TCP-подключения
к серверам Telegram по IP закрыты (облачные песочницы, корпоративные сети).
Протокол внутри тот же MTProto (obfuscated2 + intermediate), меняется только «труба».

Использование с Telethon:
    from tg_ws_transport import ConnectionWebSocket
    TelegramClient(session, api_id, api_hash, connection=ConnectionWebSocket)

Прокси берётся из HTTPS_PROXY / https_proxy, сертификаты — из SSL_CERT_FILE, если заданы.
Требуется пакет aiohttp.
"""
from __future__ import annotations

import asyncio
import os
import ssl

from telethon.network.connection.tcpintermediate import IntermediatePacketCodec
from telethon.network.connection.tcpobfuscated import ConnectionTcpObfuscated

# Имена дата-центров в веб-адресах Telegram (как в Telegram Web)
WS_DC_NAMES = {1: "pluto", 2: "venus", 3: "aurora", 4: "vesta", 5: "flora"}

# Домены, которые нужно разрешить в сетевых настройках, чтобы этот транспорт работал
WS_HOSTS = [f"{name}.web.telegram.org" for name in WS_DC_NAMES.values()]


def ws_url(dc_id: int, test: bool = False) -> str:
    name = WS_DC_NAMES.get(abs(dc_id))
    if name is None:
        raise ValueError(f"Неизвестный дата-центр Telegram: {dc_id}")
    return f"wss://{name}.web.telegram.org/{'apiws_test' if test else 'apiws'}"


def _proxy_url() -> str | None:
    return os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy") or None


def _ssl_context() -> ssl.SSLContext:
    cafile = os.environ.get("SSL_CERT_FILE") or os.environ.get("REQUESTS_CA_BUNDLE")
    return ssl.create_default_context(cafile=cafile if cafile and os.path.isfile(cafile) else None)


class _WsReader:
    """readexactly() поверх входящих бинарных кадров WebSocket."""

    def __init__(self, ws):
        self._ws = ws
        self._buf = bytearray()

    async def readexactly(self, n: int) -> bytes:
        import aiohttp

        while len(self._buf) < n:
            msg = await self._ws.receive()
            if msg.type == aiohttp.WSMsgType.BINARY:
                self._buf += msg.data
            elif msg.type in (aiohttp.WSMsgType.CLOSE, aiohttp.WSMsgType.CLOSING,
                              aiohttp.WSMsgType.CLOSED, aiohttp.WSMsgType.ERROR):
                raise asyncio.IncompleteReadError(bytes(self._buf), n)
            # TEXT/PING/PONG в этом протоколе не используются — пропускаем
        data = bytes(self._buf[:n])
        del self._buf[:n]
        return data


class _WsWriter:
    """write()/drain()/close() как у asyncio.StreamWriter: накопленное уходит одним кадром."""

    def __init__(self, ws, session):
        self._ws = ws
        self._session = session
        self._buf = bytearray()
        self._closing: asyncio.Future | None = None

    def write(self, data: bytes) -> None:
        self._buf += data

    async def drain(self) -> None:
        if self._buf:
            data = bytes(self._buf)
            self._buf.clear()
            await self._ws.send_bytes(data)

    def close(self) -> None:
        if self._closing is None:
            self._closing = asyncio.ensure_future(self._close())

    async def _close(self) -> None:
        try:
            await self._ws.close()
        finally:
            await self._session.close()

    async def wait_closed(self) -> None:
        if self._closing is not None:
            await self._closing


class ConnectionWebSocket(ConnectionTcpObfuscated):
    """Telethon-соединение: obfuscated2 + intermediate поверх wss://<dc>.web.telegram.org/apiws."""

    packet_codec = IntermediatePacketCodec
    test_servers = False

    async def _connect(self, timeout=None, ssl=None):
        import aiohttp

        url = ws_url(self._dc_id, test=self.test_servers)
        session = aiohttp.ClientSession()
        try:
            ws = await asyncio.wait_for(
                session.ws_connect(url, protocols=("binary",), proxy=_proxy_url(),
                                   ssl=_ssl_context(), max_msg_size=0, autoping=True),
                timeout=timeout,
            )
        except BaseException:
            await session.close()
            raise
        self._reader = _WsReader(ws)
        self._writer = _WsWriter(ws, session)
        self._codec = self.packet_codec(self)
        self._init_conn()
        await self._writer.drain()
