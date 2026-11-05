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
    empty_service: bool = False  # служебное сообщение без текста — так сайт показывает «Канал запустил розыгрыш»


# Бывают каналы, у которых сайт прячет вообще все посты (одни пустые заглушки). По ним веб-проверка
# ничего не говорит: если среди собственных постов канала таких больше этой доли — канал «слепой».
BLIND_SHARE = 0.8
BLIND_MIN_POSTS = 8


_CLASS_RE = re.compile(r'class="tgme_widget_message ([^"]*)"')
# Служебное сообщение «Канал запустил розыгрыш» сайт выводит пустым: только имя канала и время.
# Прочие служебные («закрепил «…»», «канал создан», «фото изменено») — с текстом или картинкой.
LAUNCH_GAP_SECONDS = 600


def _empty_service(block: str) -> bool:
    """Пустое служебное сообщение. Проверено 05.11 на 6 розыгрышах (otdel34/382, DuckyBBL/2151,
    rzavayuui/43, energyplusmedia/2235, eternalilalll/13855, zwezdy_ot_dmitry/2393): прямо перед
    модулем-розыгрышем (номер на 1 меньше, та же минута) стоит именно такое сообщение."""
    m = _CLASS_RE.search(block)
    if not m or "service_message" not in m.group(1).split():
        return False
    return (TEXT_CLASS not in block and STUB_CLASS not in block and "service_photo" not in block
            and not _MEDIA_RE.search(block))


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
                             forwarded="tgme_widget_message_forwarded_from" in block,
                             empty_service=_empty_service(block)))
    return posts


def _launched(post: WebPost, prev: WebPost | None) -> bool:
    """Перед заглушкой — пустое служебное сообщение того же времени: так выглядит запуск розыгрыша."""
    if prev is None or not prev.empty_service:
        return False
    if post.date and prev.date:
        return abs((post.date - prev.date).total_seconds()) <= LAUNCH_GAP_SECONDS
    return True


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
        self.newest: dict[str, datetime] = {}  # дата самого свежего поста-кандидата по каналу
        self.launches: dict[str, int] = {}  # заглушек с «Канал запустил розыгрыш» перед ними, по каналу
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

    async def candidates(self, username: str, since: datetime, max_pages: int = 20,
                         strict: bool = False) -> list[int] | None:
        """id постов-заглушек не старше since. None — веб-страницы нет, проверять через API.
        strict — для каналов большого списка: свои заглушки берём, только если прямо перед ними стоит
        пустое служебное «Канал запустил розыгрыш» (так выглядит розыгрыш у организатора; без него это
        обычно репост истории или скрытое видео). Пересланные заглушки берём как раньше: у копии
        розыгрыша служебного сообщения нет. Сколько заглушек с «запуском» — в self.launches."""
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
        hits = [p for p in seen if p.stub and (p.date is None or p.date >= since)]
        by_id = {p.id: p for p in seen}
        launched = [p for p in hits if not p.forwarded and _launched(p, by_id.get(p.id - 1))]
        if launched:
            self.launches[username.lower()] = len(launched)
        if strict:
            oldest_id = min(by_id) if by_id else 0
            # предыдущее сообщение не попало на загруженные страницы — не отбрасываем, проверим через API
            ok = {p.id for p in launched}
            hits = [p for p in hits if p.forwarded or p.id in ok or p.id - 1 < oldest_id]
        dates = [p.date for p in hits if p.date]
        if dates:
            self.newest[username.lower()] = max(dates)
        return sorted({p.id for p in hits}, reverse=True)

    async def close(self) -> None:
        if self.session is not None:
            await self.session.close()
            self.session = None
