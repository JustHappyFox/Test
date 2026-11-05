"""
Быстрая предпроверка каналов через публичную веб-страницу t.me/s/<канал>.

Официальный розыгрыш (messageMediaGiveaway) на сайте t.me показывается заглушкой
«Please open Telegram to view this post» (класс message_media_not_supported): ни текста,
ни приза там нет. Заглушку получают и обычные фото/видео (защита контента, большие видео),
но у них есть подпись или блок медиа; у розыгрыша — пусто. Такие пустые заглушки редки
(замер 05.11: BetBoom 7 из 200 постов, Okko 1 из 78, агрегатор розыгрышей — все), а сами
страницы читаются без аккаунта Telegram и без его лимитов. Поэтому обход сначала смотрит веб-страницу канала и через
API открывает только посты-заглушки за нужный период; каналы без заглушек API не трогает.

Если страница недоступна (закрыт предпросмотр, это группа, t.me ограничил частоту) —
возвращаем None, и канал проверяется обычным способом через API.
"""
from __future__ import annotations

import asyncio
import logging
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Any

log = logging.getLogger("giveaway_finder")

STUB_CLASS = "message_media_not_supported"
TEXT_CLASS = "tgme_widget_message_text"
_POST_RE = re.compile(r'data-post="([^"/]+)/(\d+)"')
_DATE_RE = re.compile(r'<time datetime="([^"]+)"')
# Заглушка бывает и у обычных фото/видео (у каналов с защитой контента, у больших видео) —
# там рядом есть подпись или блок медиа. У модуля-розыгрыша нет ни текста, ни медиа.
_MEDIA_RE = re.compile(r"tgme_widget_message_(?:photo|video|document|voice|roundvideo|sticker|poll|location)")


@dataclass
class WebPost:
    id: int
    date: datetime | None
    stub: bool  # заглушка без текста и без фото/видео — так выглядит модуль-розыгрыш
    forwarded: bool = False


# Бывают каналы, у которых сайт прячет вообще все посты (одни пустые заглушки). По ним веб-проверка
# ничего не говорит: если среди собственных постов канала таких больше этой доли — канал «слепой».
BLIND_SHARE = 0.8
BLIND_MIN_POSTS = 8


def parse_page(html: str) -> list[WebPost] | None:
    """Посты со страницы t.me/s/<канал>, новые последними. None — у канала нет веб-ленты."""
    if "tgme_widget_message_wrap" not in html:
        return None
    posts: list[WebPost] = []
    for block in html.split('class="tgme_widget_message_wrap')[1:]:
        m = _POST_RE.search(block)
        if not m:
            continue
        d = _DATE_RE.search(block)
        try:
            date = datetime.fromisoformat(d.group(1)) if d else None
        except ValueError:
            date = None
        stub = STUB_CLASS in block and TEXT_CLASS not in block and not _MEDIA_RE.search(block)
        posts.append(WebPost(id=int(m.group(2)), date=date, stub=stub,
                             forwarded="tgme_widget_message_forwarded_from" in block))
    return posts


class WebPreview:
    """Чтение t.me/s/<канал> с ограничением одновременных запросов и паузой при 429."""

    BASE = "https://t.me/s/"

    def __init__(self, parallel: int = 8, timeout: float = 20, max_limited: int = 30):
        self.parallel = max(1, parallel)
        self.timeout = timeout
        self.max_limited = max_limited
        self.sem = asyncio.Semaphore(self.parallel)
        self.session: Any = None
        self.pages = 0      # прочитано страниц
        self.limited = 0    # ответов 429 (t.me просит сбавить частоту)
        self.failed = 0
        self.blind: set[str] = set()  # каналы, где сайт прячет все посты (имена в нижнем регистре)
        self.broken = False  # t.me упорно ограничивает — дальше только через API

    async def _session(self) -> Any:
        if self.session is None:
            import aiohttp
            self.session = aiohttp.ClientSession(
                trust_env=True, timeout=aiohttp.ClientTimeout(total=self.timeout),
                headers={"User-Agent": "Mozilla/5.0 (compatible; giveaway-finder)"})
        return self.session

    async def fetch(self, url: str) -> str | None:
        if self.broken:
            return None
        for attempt in range(3):
            async with self.sem:
                try:
                    session = await self._session()
                    async with session.get(url, allow_redirects=True) as r:
                        if r.status == 429:
                            self.limited += 1
                            wait = min(float(r.headers.get("Retry-After") or 5 * (attempt + 1)), 60)
                        elif r.status != 200:
                            self.failed += 1
                            return None
                        else:
                            self.pages += 1
                            return await r.text()
                except Exception as e:  # сеть, таймаут, нет доступа к t.me
                    self.failed += 1
                    log.debug("t.me: %s — %s", url, e)
                    if self.failed >= 20 and not self.pages:
                        self.broken = True
                        log.warning("Сайт t.me недоступен — проверяю каналы только через API")
                    return None
            if self.limited >= self.max_limited:
                self.broken = True
                log.warning("t.me ограничивает частоту (%d раз) — дальше проверяю только через API", self.limited)
                return None
            await asyncio.sleep(wait)
        return None

    async def candidates(self, username: str, since: datetime, max_pages: int = 20) -> list[int] | None:
        """id постов-заглушек не старше since. None — веб-страницы нет, проверять через API."""
        username = username.lstrip("@")
        url = self.BASE + username
        seen: list[WebPost] = []
        for page in range(max_pages):
            html = await self.fetch(url)
            if html is None:
                if page == 0:
                    return None
                break
            posts = parse_page(html)
            if posts is None:
                if page == 0:
                    return None
                break
            if not posts:
                break
            seen += posts
            oldest = min(posts, key=lambda p: p.id)
            if (oldest.date and oldest.date < since) or oldest.id <= 1:
                break
            url = f"{self.BASE}{username}?before={oldest.id}"
        own = [p for p in seen if not p.forwarded]
        if len(own) >= BLIND_MIN_POSTS and sum(p.stub for p in own) >= BLIND_SHARE * len(own):
            self.blind.add(username.lower())  # сайт прячет все посты — по заглушкам розыгрыш не отличить
        return sorted({p.id for p in seen if p.stub and (p.date is None or p.date >= since)}, reverse=True)

    async def close(self) -> None:
        if self.session is not None:
            await self.session.close()
            self.session = None
