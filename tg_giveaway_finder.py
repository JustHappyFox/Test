#!/usr/bin/env python3
"""
tg_giveaway_finder — поиск ещё не завершённых ОФИЦИАЛЬНЫХ розыгрышей Telegram
(Premium / Stars), запущенных через встроенный механизм Giveaways.

Как это работает
----------------
Официальный розыгрыш — это сообщение с медиа `messageMediaGiveaway`
(в нём: каналы-участники, число победителей, дата окончания, приз).
Глобального API «дай все розыгрыши Telegram» не существует, поэтому скрипт:

  1. Собирает стартовый список каналов:
       * ваши каналы/группы (--dialogs);
       * каналы из файла/аргументов (--channels-file / -c);
       * публичные каналы по ключевым словам (contacts.search);
       * публичные посты по хэштегам #розыгрыш, #giveaway ... (channels.searchPosts);
       * (опционально) полнотекстовый поиск публичных постов (--post-search),
         только в пределах бесплатного дневного лимита.
  2. Читает свежую историю каждого канала и ищет messageMediaGiveaway
     с until_date в будущем (в т.ч. пересланные розыгрыши).
  3. «Снежный ком»: каналы-участники найденных розыгрышей и источники
     пересылок тоже добавляются в очередь обхода (--depth).
  4. (опционально, --check) Через payments.getGiveawayInfo проверяет,
     можете ли участвовать именно вы (страна, «только новые подписчики» и т.п.).

Нужен ПОЛЬЗОВАТЕЛЬСКИЙ аккаунт (не бот): api_id / api_hash с https://my.telegram.org.
"""
from __future__ import annotations

import argparse
import asyncio
import csv
import itertools
import json
import logging
import os
import random
import re
import sys
from collections import deque
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable

from telethon import TelegramClient, errors, events, functions, types, utils

log = logging.getLogger("giveaway_finder")

DEFAULT_KEYWORDS = [
    # русские
    "розыгрыш", "розыгрыши", "розыгрыш premium", "розыгрыш звезд",
    "раздача", "конкурс", "халява", "бесплатный премиум",
    # международные: в иностранных каналах тоже бывают розыгрыши «для всех стран»
    "giveaway", "giveaways", "premium giveaway", "stars giveaway", "telegram premium giveaway",
    "free telegram premium", "free stars", "telegram stars giveaway", "gift giveaway", "airdrop giveaway",
    "sorteo", "sorteio", "çekiliş", "розіграш", "قرعه کشی", "مسابقه", "سحب",
]
DEFAULT_HASHTAGS = ["розыгрыш", "giveaway", "конкурс", "раздача", "розыгрыши", "giveaways",
                    "telegrampremium", "premiumgiveaway", "starsgiveaway", "freepremium", "freestars",
                    "airdrop", "sorteo", "sorteio", "cekilis", "розіграш"]

# Публичные каналы-агрегаторы, которые пересылают официальные розыгрыши (найдены через веб-поиск
# и TGStat). Их надёжность не важна: в результат попадают только настоящие messageMediaGiveaway,
# а не текстовые «конкурсы». Отключить: --no-aggregators.
DEFAULT_AGGREGATORS = [
    # активные (проверено 05.11: среди последних 300 постов почти все — официальные розыгрыши)
    "@giveaways_database",        # 287 розыгрышей на 300 постов, 11 активных; сайт telegramgiveaways.ru
    "@razdacha_TelegramPremium",  # источник владельца
    # архивные: новых постов давно нет, но в истории сотни организаторов — проверяем их каналы
    "@Premiumlob", "@tg_premium_prize", "@giveaways_premium_telegramtg", "@tgpremiumx",
    "@giveaways", "@Giveaways_Premium_Tel", "@TgPremiumGiveaway",
    "@tgpremium_giveaways", "@ChannelGiveaways",
    # найдены веб-поиском 05.11 по заглушке «Please open Telegram to view this post» + «розыгрыш/giveaway»
    "@loot_premium", "@yprmk", "@PremiumGiveAwayTg", "@giveawaystarstelegram", "@rf_free",
    "@PremiumGiveaway_Prizes", "@giveawayprizeshub",
    # веб-поиск на 15 языках 05.11 (~50 каналов), оставлены только те, где за 31 день есть модули-розыгрыши:
    "@giveawayscryptobot",   # 14 модулей
    "@Permio_ir",            # 6, персидский
    "@premium_accounfree",   # 5, персидский
    "@rozygryshi_zvyozd",    # 2
    "@StarsGiveawayHub",     # владелец нашёл через Google 06.11: 55 постов за 31 день, 49 модулей (пересланных)
]

# Значения по умолчанию для обычного и осторожного (--careful) режимов.
# Явно указанные в командной строке опции всегда важнее пресета.
# Паузы между отдельными каналами по умолчанию нет: частоту запросов регулирует сам Telegram
# через FloodWait, а Telethon его автоматически пережидает.
PRESETS = {
    "normal": dict(delay="0", page_delay=None, max_channels=2000, per_channel_limit=400,
                   search_pages=3, max_flood_wait=120, batch_size=500, batch_pause=0.0),
    # Осторожный: пауза 30 с после каждых 500 каналов и терпеливое ожидание FloodWait до 5 минут.
    "careful": dict(delay="0", page_delay=None, max_channels=2000, per_channel_limit=400,
                    search_pages=3, max_flood_wait=300, batch_size=500, batch_pause=30.0),
}


def parse_post_hours(value: str | None) -> tuple[int, int] | None:
    """'7-22' -> (7, 22): публиковать с 07:00 до 22:00 по местному времени (22:00 уже нельзя)."""
    if not value:
        return None
    a, b = (int(x) for x in str(value).split("-", 1))
    if not 0 <= a < b <= 24:
        raise ValueError(f"--post-hours: нужно «начало-конец», например 7-22, а не {value!r}")
    return a, b


def in_post_window(t: datetime, window: tuple[int, int] | None, tz_hours: float = 3.0) -> datetime:
    """Ближайший момент не раньше t внутри окна публикаций (по местному времени UTC+tz_hours).
    Внутри окна — сам t; ночью — начало окна (утром того же или следующего дня)."""
    if not window:
        return t
    start, end = window
    tz = timezone(timedelta(hours=tz_hours))
    loc = t.astimezone(tz)
    if start <= loc.hour < end:
        return t
    day = loc.date() if loc.hour < start else loc.date() + timedelta(days=1)
    return datetime(day.year, day.month, day.day, start, tzinfo=tz).astimezone(timezone.utc)


def parse_range(value: str | float | tuple | None) -> tuple[float, float] | None:
    """'5' -> (5, 5), '5-15' -> (5, 15). Пауза выбирается случайно в этом диапазоне."""
    if value is None or isinstance(value, tuple):
        return value
    lo, _, hi = str(value).partition("-")
    try:
        lo_f = float(lo)
        hi_f = float(hi) if hi else lo_f
    except ValueError:
        raise argparse.ArgumentTypeError(f"неверный диапазон пауз: {value!r} (пример: 5 или 5-15)") from None
    if lo_f < 0 or hi_f < lo_f:
        raise argparse.ArgumentTypeError(f"неверный диапазон пауз: {value!r}")
    return lo_f, hi_f


def apply_preset(args: argparse.Namespace) -> argparse.Namespace:
    preset = PRESETS["careful" if getattr(args, "careful", False) else "normal"]
    for name, default in preset.items():
        if getattr(args, name, None) is None:
            setattr(args, name, default)
    args.delay = parse_range(args.delay)
    args.page_delay = parse_range(args.page_delay)
    return args


def rand_pause(rng: tuple[float, float] | None) -> float:
    return random.uniform(*rng) if rng else 0.0


# --------------------------------------------------------------------------- #
# Конфигурация
# --------------------------------------------------------------------------- #

def load_dotenv(path: Path) -> None:
    """Мини-парсер .env (KEY=VALUE), чтобы не тянуть лишние зависимости."""
    if not path.is_file():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def make_client() -> TelegramClient:
    here = Path(__file__).resolve().parent
    load_dotenv(Path.cwd() / ".env")
    load_dotenv(here / ".env")
    # .strip(): при копировании с телефона в секреты часто прилипают пробелы и переносы строк
    api_id = (os.environ.get("TG_API_ID") or "").strip()
    api_hash = (os.environ.get("TG_API_HASH") or "").strip()
    if not api_id or not api_hash:
        sys.exit(
            "Не заданы TG_API_ID / TG_API_HASH.\n"
            "Получите их на https://my.telegram.org -> API development tools\n"
            "и положите в файл .env рядом со скриптом (см. .env.example)."
        )
    string_session = "".join((os.environ.get("TG_STRING_SESSION") or "").split())
    session_file = os.environ.get("TG_SESSION", str(here / "giveaway_finder"))
    if string_session:  # удобно для облака (Colab / GitHub Actions): сессия хранится в секрете
        from telethon.sessions import StringSession
        try:
            session: Any = StringSession(string_session)
        except Exception as e:  # битая/обрезанная строка — понятная ошибка без вывода самой строки
            sys.exit(f"TG_STRING_SESSION повреждена ({type(e).__name__}): длина {len(string_session)} "
                     f"символов, начинается с {string_session[:1]!r}. Обычно это 353 символа и начинается "
                     f"с '1'. Скопируйте строку заново целиком.")
    else:
        session = session_file
    kwargs: dict[str, Any] = {}
    if use_websocket():
        # Только HTTPS через прокси (облачные среды): подключаемся как Telegram Web
        from tg_ws_transport import ConnectionWebSocket
        kwargs["connection"] = ConnectionWebSocket
        log.info("Подключение к Telegram через WebSocket (*.web.telegram.org)")
    return TelegramClient(session, int(api_id), api_hash, **kwargs)


def use_websocket() -> bool:
    """TG_TRANSPORT=ws|tcp. По умолчанию WebSocket, если задан HTTPS-прокси (через него TCP не пройдёт)."""
    mode = os.environ.get("TG_TRANSPORT", "").strip().lower()
    if mode in ("ws", "websocket", "wss"):
        return True
    if mode == "tcp":
        return False
    return bool(os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy"))


def login_state_path() -> Path:
    here = Path(__file__).resolve().parent
    return Path(os.environ.get("TG_SESSION", str(here / "giveaway_finder")) + ".login.json")


# --------------------------------------------------------------------------- #
# Модель данных
# --------------------------------------------------------------------------- #

@dataclass
class ChannelRef:
    id: int
    title: str | None = None
    username: str | None = None
    member: bool | None = None  # подписаны ли вы

    @property
    def link(self) -> str:
        return f"https://t.me/{self.username}" if self.username else f"https://t.me/c/{self.id}"

    def label(self) -> str:
        name = self.title or f"id:{self.id}"
        return f"{name} (@{self.username})" if self.username else name


@dataclass
class Giveaway:
    key: str
    channel_ids: list[int]
    quantity: int
    until_date: datetime
    posted_at: datetime | None
    months: int | None = None
    stars: int | None = None
    prize_description: str | None = None
    countries: list[str] = field(default_factory=list)
    only_new_subscribers: bool = False
    winners_are_visible: bool = False
    # (channel_id, msg_id): оригиналы — первыми, затем пересланные копии
    origins: list[tuple[int, int]] = field(default_factory=list)
    copies: list[tuple[int, int]] = field(default_factory=list)
    channels: list[ChannelRef] = field(default_factory=list)
    status: str | None = None
    eligible: bool | None = None

    @property
    def kind(self) -> str:
        if self.stars:
            return "stars"
        if self.months:
            return "premium"
        return "other"

    def prize_text(self) -> str:
        if self.stars:
            per = self.stars // max(self.quantity, 1)
            text = f"{self.stars} ⭐ Stars, победителей: {self.quantity} (~{per} ⭐ каждому)"
        elif self.months:
            text = f"Telegram Premium на {self.months} мес., победителей: {self.quantity}"
        else:
            text = f"Приз, победителей: {self.quantity}"
        if self.prize_description:
            text += f" + {self.prize_description}"
        return text

    def time_left(self, now: datetime | None = None) -> str:
        now = now or datetime.now(timezone.utc)
        sec = int((self.until_date - now).total_seconds())
        if sec <= 0:
            return "завершён"
        d, rem = divmod(sec, 86400)
        h, rem = divmod(rem, 3600)
        m = rem // 60
        if d:
            return f"{d}д {h}ч"
        if h:
            return f"{h}ч {m}м"
        return f"{m}м"


def giveaway_key(media: types.MessageMediaGiveaway) -> str:
    # Один и тот же розыгрыш публикуется во всех каналах-участниках и пересылается,
    # поэтому дедуплицируем по содержимому, а не по id сообщения.
    until = int(media.until_date.timestamp()) if media.until_date else 0
    chans = ",".join(str(c) for c in sorted(media.channels or []))
    return f"{chans}|{until}|{media.quantity}|{media.months or 0}|{media.stars or 0}"


def message_link(chan: ChannelRef | None, chan_id: int, msg_id: int) -> str:
    if chan and chan.username:
        return f"https://t.me/{chan.username}/{msg_id}"
    return f"https://t.me/c/{chan_id}/{msg_id}"


def active_username(ent: Any) -> str | None:
    if getattr(ent, "username", None):
        return ent.username
    for u in getattr(ent, "usernames", None) or []:
        if getattr(u, "active", False):
            return u.username
    return None


class WebOnly(str):
    """@username из большого списка (Common Crawl). Проверяем только по веб-странице t.me/s/…:
    нет страницы — пропускаем, через API открываем лишь каналы с постами-заглушками."""


def to_ref(ent: Any) -> ChannelRef:
    # У "min"-объектов (пришли вместе с чужими сообщениями) флаг left недостоверен.
    known = isinstance(ent, types.Channel) and not getattr(ent, "min", False)
    return ChannelRef(
        id=ent.id,
        title=getattr(ent, "title", None),
        username=active_username(ent),
        member=(not ent.left) if known else None,
    )


# --------------------------------------------------------------------------- #
# Поиск
# --------------------------------------------------------------------------- #

class GiveawayFinder:
    def __init__(self, client: TelegramClient, args: argparse.Namespace):
        self.client = client
        self.args = args
        self.now = datetime.now(timezone.utc)
        self.queue: deque[tuple[Any, int]] = deque()   # (entity | int id | str, depth)
        self.queued_ids: set[int] = set()
        self.scanned: set[int] = set()
        self.giveaways: dict[str, Giveaway] = {}
        self.refs: dict[int, ChannelRef] = {}           # кэш инфо о каналах
        self.stopped_reason: str | None = None
        # Поиск «до первой находки»: want(g) — подходит ли розыгрыш, want_count — сколько нужно
        self.want: Any = None
        self.want_count = 0
        self.wanted = 0
        # Предпроверка по веб-странице t.me/s/<канал> (--web): см. tg_web_preview.py
        self.web: Any = None
        self.web_seen: set[str] = set()
        self.web_tasks: dict[str, asyncio.Future] = {}
        self.web_checked = 0   # каналов проверено по веб-странице
        self.web_opened = 0    # из них открыто через API (были посты-заглушки)
        self.api_read = 0      # каналов прочитано через API целиком
        self.ref_tried: set[int] = set()
        self.entities: dict[int, Any] = {}
        self.deadline: float | None = None  # loop.time(), после которого обход останавливается
        # Большой список каналов (--discover-file): сколько всего, с какой позиции и сколько разобрали
        self.discover_total = 0
        self.discover_start = 0
        self.discover_consumed = 0
        self.discover_opened = 0   # открыто через API (поиск по @username у Telegram строго ограничен)
        self.discover_found = 0    # активных розыгрышей нашлось в каналах большого списка
        self.discover_paused = False
        # Посты-заглушки, которые уже открывали и это оказались не розыгрыши: {@username: {id}}.
        # Хранятся в учёте между запусками, чтобы не открывать одно и то же через API каждый поиск.
        self.seen_stubs: dict[str, set[int]] = {
            k: set(v) for k, v in (getattr(args, "seen_stubs", None) or {}).items()}
        # Каналы, которые уже находили: {@username: [id, access_hash]} — без повторного поиска по имени
        self.peer_cache: dict[str, list[int]] = dict(getattr(args, "peer_cache", None) or {})
        self.names_blocked = False  # Telegram временно запретил поиск по имени (FloodWait на ResolveUsername)

    @property
    def found_enough(self) -> bool:
        return bool(self.want_count) and self.wanted >= self.want_count

    @property
    def checked(self) -> int:
        """Сколько всего каналов проверено (по веб-странице и через API)."""
        return self.web_checked + self.api_read

    @property
    def out_of_budget(self) -> bool:
        """Обход упёрся в лимит каналов или времени — искать новые каналы уже незачем."""
        web_full = self.web is not None and self.web_checked >= (getattr(self.args, "web_max", None) or 5000)
        return bool(self.stopped_reason) or self.api_read >= self.args.max_channels or web_full

    async def pause(self) -> None:
        """Случайная пауза между запросами — чтобы не долбить API с ровным интервалом."""
        await asyncio.sleep(rand_pause(self.args.delay))

    # ---------- очередь ----------

    def enqueue(self, item: Any, depth: int, front: bool = False) -> None:
        cid: int | None = None
        if isinstance(item, int):
            cid = item
        elif not isinstance(item, str):
            if isinstance(item, types.Channel):
                self.refs.setdefault(item.id, to_ref(item))
            elif not isinstance(item, (types.InputPeerChannel, types.InputChannel, types.PeerChannel)):
                return  # личные чаты, обычные группы и т.п. — пропускаем
            cid = utils.get_peer_id(item, add_mark=False)
        if cid is not None:
            if cid in self.queued_ids or cid in self.scanned:
                return
            self.queued_ids.add(cid)
        (self.queue.appendleft if front else self.queue.append)((item, depth))

    def remember_peer(self, ent: Any) -> None:
        """Запомнить канал (id и access_hash) по @username, чтобы в следующих запусках не искать его по имени."""
        u = active_username(ent)
        if u and isinstance(ent, types.Channel) and ent.access_hash and not getattr(ent, "min", False):
            self.peer_cache[u.lower()] = [ent.id, ent.access_hash]

    async def resolve(self, item: Any) -> Any | None:
        if isinstance(item, int) and item in self.entities:
            return self.entities[item]
        if isinstance(item, str):
            cached = self.peer_cache.get((self.username_of(item) or "").lower())
            if cached:  # канал уже находили раньше — поиск по имени (строго ограничен) не нужен
                return types.InputPeerChannel(int(cached[0]), int(cached[1]))
            if self.names_blocked:
                return None  # Telegram временно ограничил поиск по имени — не дёргаем его
        try:
            if isinstance(item, (int, str)):
                ent = await self.client.get_entity(types.PeerChannel(item) if isinstance(item, int) else str(item))
                if not isinstance(ent, types.Channel):
                    log.warning("%s — не канал/супергруппа, пропускаю", item)
                    return None
                self.refs[ent.id] = to_ref(ent)
                self.remember_peer(ent)
                return ent
            return item
        except errors.FloodWaitError as e:
            if isinstance(item, str):
                self.names_blocked = True
                log.warning("Telegram ограничил поиск каналов по имени на %d мин — в этом запуске "
                            "проверяю только уже знакомые каналы", e.seconds // 60)
                return None
            raise
        except (ValueError, TypeError, errors.RPCError) as e:
            log.debug("Не удалось получить %r: %s", item, e)
            return None

    # ---------- источники ----------

    async def add_seeds(self) -> None:
        seeds: list[str] = list(self.args.channel or [])
        if not getattr(self.args, "no_aggregators", True):
            seeds += DEFAULT_AGGREGATORS
        for path in self.args.channels_file or []:
            if not Path(path).is_file():
                log.warning("Файл каналов %s не найден — пропускаю", path)
                continue
            for line in Path(path).read_text(encoding="utf-8").splitlines():
                line = line.split("#", 1)[0].strip()
                if line:
                    seeds.append(line)
        # Сначала каналы большого списка: там только подтверждённые по веб-странице розыгрыши
        # («Канал запустил розыгрыш»), их открыть быстро; стартовые каналы читаются дольше.
        self.add_discover_file(skip={str(x).strip().lstrip("@").lower() for x in seeds})
        for s in seeds:
            self.enqueue(s, 0)
        if seeds:
            log.info("Стартовых каналов из списка: %d", len(seeds))

    def add_discover_file(self, skip: set[str] = frozenset()) -> None:
        """Большой список каналов (Common Crawl): только веб-проверка, с позиции --discover-offset по кругу.
        skip — стартовые каналы: их читаем полностью, а не только по веб-странице."""
        path = getattr(self.args, "discover_file", None)
        if not path:
            return
        if not Path(path).is_file():
            log.warning("Список каналов %s не найден — пропускаю", path)
            return
        names = [ln.split("#", 1)[0].strip() for ln in Path(path).read_text(encoding="utf-8").splitlines()]
        names = [n for n in names if n]
        self.discover_total = len(names)
        if not names:
            return
        self.discover_start = (getattr(self.args, "discover_offset", 0) or 0) % len(names)
        for n in names[self.discover_start:] + names[:self.discover_start]:
            if n.lstrip("@").lower() not in skip:
                self.enqueue(WebOnly(n), 1)
        log.info("Каналов из большого списка: %d, начинаю с №%d", len(names), self.discover_start + 1)

    async def add_dialogs(self) -> None:
        n = 0
        async for d in self.client.iter_dialogs():
            if isinstance(d.entity, types.Channel):
                self.enqueue(d.entity, 0)
                n += 1
        log.info("Ваших каналов/супергрупп: %d", n)

    async def discover_keywords(self, keywords: Iterable[str]) -> None:
        before = len(self.queue)
        for kw in keywords:
            try:
                found = await self.client(functions.contacts.SearchRequest(q=kw, limit=100))
            except errors.RPCError as e:
                log.warning("Поиск каналов «%s» не удался: %s", kw, e)
                continue
            for ch in found.chats:
                if isinstance(ch, types.Channel):
                    self.enqueue(ch, 1)
            await self.pause()
        log.info("Каналов найдено поиском по названию: %d", len(self.queue) - before)

    async def _search_posts(self, label: str, pages: int, **kw: Any) -> None:
        offset_rate, offset_peer, offset_id = 0, types.InputPeerEmpty(), 0
        hits = 0
        for _ in range(pages):
            res = await self.client(functions.channels.SearchPostsRequest(
                offset_rate=offset_rate, offset_peer=offset_peer, offset_id=offset_id,
                limit=100, **kw,
            ))
            chats = {c.id: c for c in getattr(res, "chats", []) if isinstance(c, types.Channel)}
            for ch in chats.values():
                self.refs.setdefault(ch.id, to_ref(ch))
                self.remember_peer(ch)
            msgs = [m for m in getattr(res, "messages", []) if isinstance(m, types.Message)]
            for m in msgs:
                cid = utils.get_peer_id(m.peer_id, add_mark=False)
                ch = chats.get(cid)
                if ch is None:
                    continue
                hits += 1
                self.handle_message(m, cid, depth=1)  # вдруг сам пост — розыгрыш
                self.enqueue(ch, 1)
            next_rate = getattr(res, "next_rate", None)
            if not msgs or not next_rate:
                break
            last = msgs[-1]
            last_chat = chats.get(utils.get_peer_id(last.peer_id, add_mark=False))
            if last_chat is None:
                break
            offset_rate, offset_peer, offset_id = next_rate, utils.get_input_peer(last_chat), last.id
            await self.pause()
        log.info("%s: постов %d, каналов в очереди: %d", label, hits, len(self.queue))

    async def discover_hashtags(self, hashtags: Iterable[str]) -> None:
        for tag in hashtags:
            tag = tag.lstrip("#")
            try:
                await self._search_posts(f"#{tag}", self.args.search_pages, hashtag=tag)
            except errors.FloodWaitError:
                raise
            except errors.RPCError as e:
                log.warning("Поиск по хэштегу #%s не удался: %s", tag, e)

    async def discover_post_query(self, query: str) -> None:
        """Полнотекстовый поиск публичных постов. Только бесплатные запросы — Stars не тратим."""
        try:
            flood = await self.client(functions.channels.CheckSearchPostsFloodRequest(query=query))
        except errors.RPCError as e:
            log.warning("Полнотекстовый поиск недоступен: %s", e)
            return
        if not flood.query_is_free and flood.remains <= 0:
            log.warning(
                "Бесплатный лимит поиска постов на сегодня исчерпан (%s/%s); «%s» пропущен. "
                "Платный поиск за Stars скрипт специально не использует.",
                flood.remains, flood.total_daily, query,
            )
            return
        log.info("Поиск постов «%s» (осталось бесплатных: %s из %s)", query, flood.remains, flood.total_daily)
        try:
            await self._search_posts(f"«{query}»", 1, query=query)
        except errors.RPCError as e:
            log.warning("Поиск постов «%s» не удался: %s", query, e)

    # ---------- разбор сообщений ----------

    def handle_message(self, msg: types.Message, chat_id: int, depth: int) -> Giveaway | None:
        media = getattr(msg, "media", None)
        if isinstance(media, types.MessageMediaGiveawayResults):
            # итоги розыгрыша: сам розыгрыш кончился, но его организатор мог запустить новый
            if depth < self.args.depth:
                self.enqueue(media.channel_id, depth + 1)
            return None
        if not isinstance(media, types.MessageMediaGiveaway):
            return None

        fwd = getattr(msg, "fwd_from", None)
        is_forward = bool(fwd and isinstance(fwd.from_id, types.PeerChannel) and fwd.channel_post)
        origin = (fwd.from_id.channel_id, fwd.channel_post) if is_forward else (chat_id, msg.id)

        if not media.until_date or media.until_date <= self.now:
            # завершённый розыгрыш сам не берём, но каналы его организаторов проверяем:
            # кто-то мог запустить новый и не скинуть его в агрегатор
            if depth < self.args.depth:
                for cid in [*(media.channels or []), origin[0]]:
                    self.enqueue(cid, depth + 1)
            return None

        key = giveaway_key(media)
        g = self.giveaways.get(key)
        if g is None:
            g = Giveaway(
                key=key,
                channel_ids=list(media.channels or []),
                quantity=media.quantity,
                until_date=media.until_date,
                posted_at=(fwd.date if is_forward else msg.date),
                months=media.months,
                stars=media.stars,
                prize_description=media.prize_description,
                countries=list(media.countries_iso2 or []),
                only_new_subscribers=bool(media.only_new_subscribers),
                winners_are_visible=bool(media.winners_are_visible),
            )
            self.giveaways[key] = g
            if self.want is not None and self.want(g):
                self.wanted += 1
            log.info("  🎁 найден розыгрыш: %s | до %s | страны: %s | %s", g.prize_text(),
                     g.until_date.astimezone().strftime("%Y-%m-%d %H:%M"),
                     ",".join(g.countries) or "все",
                     message_link(self.refs.get(origin[0]), *origin))
        if origin not in g.origins:
            g.origins.append(origin)
        if is_forward and (chat_id, msg.id) not in g.copies:
            g.copies.append((chat_id, msg.id))

        if depth < self.args.depth:
            for cid in [*(media.channels or []), origin[0]]:
                self.enqueue(cid, depth + 1, front=True)  # каналы с розыгрышами — в приоритет
        return g

    async def scan_channel(self, ent: Any, depth: int) -> None:
        # Стартовые каналы (агрегаторы, свой список) читаем глубже — последние --source-limit записей
        # без ограничения по дате: оттуда берём организаторов и старых розыгрышей.
        deep = depth == 0 and getattr(self.args, "source_limit", 0)
        limit = self.args.source_limit if deep else self.args.per_channel_limit
        cutoff = None if deep else self.now - timedelta(days=self.args.days)
        cid = utils.get_peer_id(ent, add_mark=False)
        # wait_time — пауза между страницами истории (по 100 сообщений) одного канала
        async for msg in self.client.iter_messages(ent, limit=limit,
                                                   wait_time=rand_pause(self.args.page_delay) or None):
            if cutoff and msg.date and msg.date < cutoff:
                break
            fwd = getattr(msg, "forward", None)
            if fwd is not None and isinstance(getattr(fwd, "chat", None), types.Channel):
                self.refs.setdefault(fwd.chat.id, to_ref(fwd.chat))
            self.handle_message(msg, cid, depth)

    # ---------- предпроверка по веб-странице t.me ----------

    @staticmethod
    def cid_of(item: Any) -> int | None:
        if isinstance(item, int):
            return item
        if isinstance(item, str):
            return None
        try:
            return utils.get_peer_id(item, add_mark=False)
        except (TypeError, ValueError):
            return None

    def username_of(self, item: Any) -> str | None:
        if isinstance(item, str):
            u = item.strip().split("?", 1)[0].rstrip("/")
            for prefix in ("https://", "http://"):
                if u.startswith(prefix):
                    u = u[len(prefix):]
            if u.startswith("t.me/"):
                u = u[len("t.me/"):]
            u = u.lstrip("@")
            return u if re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{3,31}", u) else None
        if isinstance(item, types.Channel):
            return active_username(item)
        if isinstance(item, int):
            ref = self.refs.get(item)
            return ref.username if ref else None
        return None

    @property
    def web_since(self) -> datetime:
        return self.now - timedelta(days=self.args.days)

    def web_check(self, username: str, strict: bool = False) -> asyncio.Future:
        """strict — для каналов большого списка: только заглушки с «Канал запустил розыгрыш» перед ними
        (без второго уровня с пропущенным номером: 0 розыгрышей из 31 открытого канала),
        и только за последние --discover-days дней (владелец: пост старше 7 дней скорее всего истёк)."""
        since = self.web_since
        if strict:
            since = max(since, self.now - timedelta(days=getattr(self.args, "discover_days", None) or 7))
            gaps = bool(getattr(self.args, "discover_gaps", False))  # --discover-gaps: партия второго уровня
            return asyncio.ensure_future(self.web.candidates(
                username, since, getattr(self.args, "web_pages", None) or 20, strict=True, gaps=gaps))
        return asyncio.ensure_future(
            self.web.candidates(username, since, getattr(self.args, "web_pages", None) or 20))

    async def batch_refs(self, first: int) -> None:
        """@username для веб-проверки: один запрос сразу на 100 каналов из очереди (их id — из розыгрышей)."""
        more = (it for it, _ in itertools.islice(self.queue, 0, 3000)
                if isinstance(it, int) and it not in self.refs and it not in self.ref_tried)
        ids = list(dict.fromkeys([first, *itertools.islice(more, 99)]))
        inputs = []
        for cid in ids:
            self.ref_tried.add(cid)
            try:
                inputs.append(utils.get_input_channel(await self.client.get_input_entity(types.PeerChannel(cid))))
            except (ValueError, TypeError):
                pass
        await self._get_channels(inputs)

    async def _get_channels(self, inputs: list[Any]) -> None:
        """Один битый канал отклоняет всю пачку — тогда делим её пополам, пока не отсеем его."""
        if not inputs:
            return
        try:
            res = await self.client(functions.channels.GetChannelsRequest(inputs))
        except errors.FloodWaitError:
            raise
        except errors.RPCError as e:
            if len(inputs) == 1:
                return  # этот канал проверим обычным способом через API
            log.debug("Пачка из %d каналов отклонена (%s) — делю пополам", len(inputs), e)
            half = len(inputs) // 2
            await self._get_channels(inputs[:half])
            await self._get_channels(inputs[half:])
            return
        for ch in getattr(res, "chats", []):
            if isinstance(ch, types.Channel):
                self.refs[ch.id] = to_ref(ch)
                self.entities[ch.id] = ch
                self.remember_peer(ch)
        await self.pause()

    def prefetch(self) -> None:
        """Заранее (параллельно) запустить веб-проверку следующих каналов очереди."""
        room = self.web.parallel * 2 - len(self.web_tasks)
        for item, depth in itertools.islice(self.queue, 0, 60):
            if room <= 0:
                break
            u = self.username_of(item) if depth else None
            if isinstance(item, WebOnly) and self.discover_paused:
                continue  # каналы списка, оставленные на следующий поиск, не скачиваем
            if not u or u.lower() in self.web_seen or u.lower() in self.web_tasks:
                continue
            self.web_tasks[u.lower()] = self.web_check(u, strict=isinstance(item, WebOnly))
            room -= 1

    async def try_web(self, item: Any, depth: int) -> bool:
        """Проверить канал по веб-странице. True — проверен, целиком через API читать не нужно.
        Каналы большого списка (WebOnly) через API целиком не читаются никогда."""
        web_only = isinstance(item, WebOnly)
        if web_only and (self.web is None or self.web.broken):
            return True  # оставляем на следующий поиск
        count = web_only  # засчитать ли канал списка как разобранный (место остановки — по ним)
        if web_only and self.discover_paused:
            if (self.username_of(item) or "").lower() not in self.peer_cache:
                return True  # незнакомые каналы — в следующий поиск
            count = False  # знакомый проверим (поиск по имени не нужен), но в разобранные не засчитываем
        if self.web is None or self.web.broken or depth == 0:
            return False  # стартовые каналы читаем глубоко через API
        cid = self.cid_of(item)
        if cid is not None and cid in self.scanned:
            return True
        if isinstance(item, int) and item not in self.refs and item not in self.ref_tried:
            await self.batch_refs(item)
        uname = self.username_of(item)
        if not uname:
            self.discover_consumed += count
            return web_only
        key = uname.lower()
        if key in self.web_seen:
            self.discover_consumed += count
            return True
        self.web_seen.add(key)  # второй раз этот канал не проверяем никаким способом
        task = self.web_tasks.pop(key, None) or self.web_check(uname, strict=web_only)
        self.prefetch()
        cands = await task
        if cands is None:
            self.discover_consumed += count
            return web_only  # веб-страницы нет: обычный канал проверим через API, из списка — пропустим
        self.web_checked += 1
        if cid is not None:
            self.scanned.add(cid)
        if self.web_checked % 250 == 0:
            log.info("Веб-проверка: каналов %d, открыто через API %d, в очереди %d",
                     self.web_checked, self.web_opened, len(self.queue))
        seen = self.seen_stubs.get(key, set())
        cands = [c for c in cands if c not in seen]  # уже проверенные не-розыгрыши не открываем
        if web_only and key in getattr(self.web, "blind", ()):
            cands = []  # канал списка, где сайт прячет все посты: вслепую через API не открываем
        if not cands:
            self.discover_consumed += count
            return True
        if web_only:
            limit = getattr(self.args, "web_open_max", None) or 80
            known = key in self.peer_cache  # уже находили — поиск по имени не нужен и лимит не тратит
            if not known and (self.discover_opened >= limit or self.names_blocked):
                self.discover_paused = True
                self.web_seen.discard(key)  # этот и следующие каналы списка — в следующий поиск
                log.info("Большой список: %d новых каналов найдено по имени — это предел на один поиск "
                         "(поиск по @username у Telegram строго ограничен), остальные в следующий раз", limit)
                return True
            self.discover_opened += not known
            self.discover_consumed += count
        ent = await self.resolve(item)
        if ent is None:
            return True
        cid = utils.get_peer_id(ent, add_mark=False)
        self.scanned.add(cid)
        self.web_opened += 1
        ref = self.refs.get(cid)
        log.info("[веб %d, в очереди %d] %s: постов-заглушек %d — открываю через API",
                 self.web_checked, len(self.queue), ref.label() if ref else f"@{uname}", len(cands))
        try:
            msgs = await self.client.get_messages(ent, ids=cands[:100])
        except errors.FloodWaitError:
            raise
        except (errors.RPCError, ValueError) as e:
            log.debug("  пропуск: %s", e)
            self.peer_cache.pop(key, None)  # запомненный канал мог смениться — в следующий раз найдём заново
            return True
        giveaway_media = (types.MessageMediaGiveaway, types.MessageMediaGiveawayResults)
        got = {getattr(m, "id", None): m for m in (msgs or []) if m is not None}
        for mid in cands[:100]:
            msg = got.get(mid)
            if msg is None or not isinstance(getattr(msg, "media", None), giveaway_media):
                self.seen_stubs.setdefault(key, set()).add(mid)  # не розыгрыш — больше не открываем
            if msg is None:
                continue
            fwd = getattr(msg, "forward", None)
            if fwd is not None and isinstance(getattr(fwd, "chat", None), types.Channel):
                self.refs.setdefault(fwd.chat.id, to_ref(fwd.chat))
            if self.handle_message(msg, cid, depth) is not None and web_only:
                self.discover_found += 1
        await self.pause()
        return True

    async def close_web(self) -> None:
        for t in self.web_tasks.values():
            t.cancel()
        self.web_tasks.clear()
        if self.web is not None:
            await self.web.close()

    async def crawl(self) -> None:
        loop = asyncio.get_running_loop()
        try:
            while self.queue and self.api_read < self.args.max_channels and not self.found_enough:
                if self.web is not None and self.web_checked >= (getattr(self.args, "web_max", None) or 5000):
                    break
                if self.deadline is not None and loop.time() > self.deadline:
                    self.stopped_reason = f"вышло время поиска ({self.args.time_limit:g} мин)"
                    break
                item, depth = self.queue.popleft()
                if await self.try_web(item, depth):
                    continue
                ent = await self.resolve(item)
                if ent is None:
                    continue
                cid = utils.get_peer_id(ent, add_mark=False)
                if cid in self.scanned:
                    continue
                self.scanned.add(cid)
                self.api_read += 1
                ref = self.refs.get(cid)
                log.info("[%d/%d, в очереди %d] %s", self.api_read, self.args.max_channels,
                         len(self.queue), ref.label() if ref else f"id:{cid}")
                try:
                    await self.scan_channel(ent, depth)
                except errors.FloodWaitError:
                    raise
                except (errors.RPCError, ValueError) as e:
                    log.debug("  пропуск: %s", e)
                    if isinstance(item, str):
                        self.peer_cache.pop((self.username_of(item) or "").lower(), None)
                await self.pause()
                batch_size = getattr(self.args, "batch_size", 0)
                batch_pause = getattr(self.args, "batch_pause", 0)
                if batch_pause and batch_size and self.api_read % batch_size == 0 and self.queue:
                    log.info("Через API прочитано %d каналов — пауза %g с", self.api_read, batch_pause)
                    await asyncio.sleep(batch_pause)
        except errors.FloodWaitError as e:
            self.stopped_reason = f"Telegram ограничил частоту запросов на {e.seconds} с — обход остановлен"
        except (KeyboardInterrupt, asyncio.CancelledError):
            self.stopped_reason = "прервано пользователем"
        if self.stopped_reason:
            log.warning("%s. Показываю то, что успели найти.", self.stopped_reason)

    # ---------- пост-обработка ----------

    async def channel_ref(self, cid: int) -> ChannelRef:
        ref = self.refs.get(cid)
        if ref and ref.member is not None:
            return ref
        try:
            ent = await self.client.get_entity(types.PeerChannel(cid))
            ref = to_ref(ent)
        except (ValueError, TypeError, errors.RPCError):
            ref = ref or ChannelRef(id=cid)
        self.refs[cid] = ref
        return ref

    async def resolve_channels(self) -> None:
        for g in self.giveaways.values():
            g.channels = [await self.channel_ref(cid) for cid in g.channel_ids]

    async def check_eligibility(self) -> None:
        """payments.getGiveawayInfo: можно ли вам участвовать / не закончился ли уже."""
        for key, g in list(self.giveaways.items()):
            info = None
            for cid, mid in [*g.origins, *g.copies]:
                try:
                    peer = await self.client.get_input_entity(types.PeerChannel(cid))
                    info = await self.client(functions.payments.GetGiveawayInfoRequest(peer=peer, msg_id=mid))
                    break
                except (ValueError, TypeError, errors.RPCError) as e:
                    log.debug("getGiveawayInfo %s/%s: %s", cid, mid, e)
            await self.pause()
            if info is None:
                g.status = "не удалось проверить"
                continue
            if isinstance(info, types.payments.GiveawayInfoResults):
                del self.giveaways[key]  # уже подведены итоги
                continue
            g.status, g.eligible = describe_info(info)


def describe_info(info: types.payments.GiveawayInfo) -> tuple[str, bool | None]:
    if info.disallowed_country:
        return f"❌ недоступно для вашей страны ({info.disallowed_country})", False
    if info.admin_disallowed_chat_id:
        return "❌ вы админ одного из каналов-организаторов", False
    if info.joined_too_early_date:
        when = info.joined_too_early_date.astimezone().strftime("%Y-%m-%d")
        return f"❌ только для новых подписчиков (вы подписаны с {when})", False
    if info.preparing_results:
        return "⏳ подводятся итоги", None
    if info.participating:
        return "✅ вы уже участвуете", True
    return "➕ можно участвовать — подпишитесь на все каналы", True


# --------------------------------------------------------------------------- #
# Вывод
# --------------------------------------------------------------------------- #

def parse_countries(value: str | None) -> set[str]:
    return {c.strip().upper() for c in (value or "").split(",") if c.strip()}


def country_ok(g: Giveaway, allowed: set[str]) -> bool:
    """Без ограничений по странам — подходит всегда; иначе нужна хотя бы одна страна из списка."""
    return not g.countries or not allowed or bool(set(g.countries) & allowed)


def filter_giveaways(gs: Iterable[Giveaway], args: argparse.Namespace) -> list[Giveaway]:
    allowed = parse_countries(getattr(args, "countries", None))
    out = []
    for g in gs:
        if args.prize != "any" and g.kind != args.prize:
            continue
        if g.quantity < args.min_winners:
            continue
        if g.kind == "stars" and (g.stars or 0) < (getattr(args, "min_stars", 0) or 0):
            continue
        if g.kind == "premium" and g.quantity < (getattr(args, "min_premium", 0) or 0):
            continue
        if args.hide_ineligible and g.eligible is False:
            continue
        if not country_ok(g, allowed):
            continue
        out.append(g)
    if allowed:  # розыгрыши «для всех стран» — первыми
        return sorted(out, key=lambda g: (bool(g.countries), g.until_date))
    return sorted(out, key=lambda g: g.until_date)


def first_link(g: Giveaway, refs: dict[int, ChannelRef]) -> str:
    cid, mid = (g.origins or g.copies)[0]
    return message_link(refs.get(cid), cid, mid)


def print_report(gs: list[Giveaway], refs: dict[int, ChannelRef], scanned: int) -> None:
    print()
    print(f"Просканировано каналов: {scanned}. Активных официальных розыгрышей: {len(gs)}")
    print("=" * 78)
    for i, g in enumerate(gs, 1):
        until = g.until_date.astimezone().strftime("%Y-%m-%d %H:%M")
        print(f"[{i}] {g.prize_text()}")
        print(f"    До: {until} (осталось {g.time_left()})")
        chans = []
        for c in g.channels:
            mark = "" if c.member is None else (" ✅" if c.member else " ➕")
            chans.append(f"{c.label()}{mark}")
        print(f"    Каналы ({len(chans)}): " + ", ".join(chans))
        extras = []
        if g.only_new_subscribers:
            extras.append("только новые подписчики")
        if g.countries:
            extras.append("страны: " + ", ".join(g.countries))
        if extras:
            print("    Условия: " + "; ".join(extras))
        if g.status:
            print(f"    Статус: {g.status}")
        print(f"    Ссылка: {first_link(g, refs)}")
        print("-" * 78)
    if gs and any(c.member is not None for g in gs for c in g.channels):
        print("✅ — вы подписаны, ➕ — нужно подписаться (участвуют подписчики ВСЕХ каналов розыгрыша)")


def export(gs: list[Giveaway], refs: dict[int, ChannelRef], json_path: str | None, csv_path: str | None) -> None:
    rows = []
    for g in gs:
        d = asdict(g)
        d["until_date"] = g.until_date.isoformat()
        d["posted_at"] = g.posted_at.isoformat() if g.posted_at else None
        d["kind"] = g.kind
        d["prize"] = g.prize_text()
        d["link"] = first_link(g, refs)
        d["channels"] = [{**asdict(c), "link": c.link} for c in g.channels]
        rows.append(d)
    if json_path:
        Path(json_path).write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
        log.info("Сохранено в %s", json_path)
    if csv_path:
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["until_date", "kind", "prize", "winners", "only_new_subscribers",
                        "countries", "status", "link", "channels"])
            for r in rows:
                w.writerow([r["until_date"], r["kind"], r["prize"], r["quantity"],
                            r["only_new_subscribers"], ",".join(r["countries"]), r["status"] or "",
                            r["link"], " ".join(c["link"] for c in r["channels"])])
        log.info("Сохранено в %s", csv_path)


def save_sources(gs: list[Giveaway], path: str) -> None:
    """Дописывает каналы с найденными розыгрышами в файл — пригодится как --channels-file."""
    p = Path(path)
    known = set(p.read_text(encoding="utf-8").split()) if p.exists() else set()
    new = sorted({f"@{c.username}" for g in gs for c in g.channels if c.username} - known)
    if new:
        with p.open("a", encoding="utf-8") as f:
            f.write("\n".join(new) + "\n")
    log.info("В %s добавлено каналов: %d", path, len(new))


# --------------------------------------------------------------------------- #
# Команды
# --------------------------------------------------------------------------- #

def describe_sources(args: argparse.Namespace) -> str:
    parts = []
    if not args.no_aggregators:
        parts.append(f"каналы-агрегаторы ({len(DEFAULT_AGGREGATORS)})")
    if args.channel or args.channels_file:
        parts.append("ваш список каналов")
    if args.dialogs:
        parts.append("ваши подписки")
    if not args.no_discover:
        parts.append("поиск новых публичных каналов по словам и хэштегам (без подписки)")
    if args.post_search:
        parts.append("поиск по тексту постов")
    if args.depth:
        parts.append(f"каналы-спонсоры найденных розыгрышей (глубина {args.depth})")
    mode = "осторожный" if args.careful else "обычный"
    pauses = []
    if args.batch_pause:
        pauses.append(f"пауза {args.batch_pause:g} с каждые {args.batch_size} каналов")
    if args.delay and args.delay[1]:
        pauses.append(f"{args.delay[0]:g}–{args.delay[1]:g} с между каналами")
    pause_text = ", ".join(pauses) or "без пауз"
    return (f"Режим: {mode} ({pause_text}, до {args.max_channels} каналов). "
            f"Источники: " + "; ".join(parts))


async def run_finder(client: TelegramClient, args: argparse.Namespace,
                     want: Any = None, want_count: int = 0) -> "GiveawayFinder":
    """want/want_count: искать только до want_count подходящих розыгрышей. Тогда сначала
    обходим известные каналы (свой список, агрегаторы) и только если там не нашлось —
    ищем новые каналы по словам и хэштегам."""
    finder = GiveawayFinder(client, args)
    finder.want, finder.want_count = want, want_count
    log.info(describe_sources(args))
    if getattr(args, "web", False):
        from tg_web_preview import WebPreview
        finder.web = WebPreview(parallel=getattr(args, "web_parallel", None) or 8)
    if getattr(args, "time_limit", None):
        finder.deadline = asyncio.get_running_loop().time() + args.time_limit * 60
    try:
        return await _run_finder(finder, client, args)
    finally:
        await finder.close_web()


async def _run_finder(finder: "GiveawayFinder", client: TelegramClient,
                      args: argparse.Namespace) -> "GiveawayFinder":
    want_count = finder.want_count
    try:
        await finder.add_seeds()
        if args.dialogs:
            await finder.add_dialogs()
        if want_count and finder.queue:
            await finder.crawl()
        if not finder.found_enough and not finder.out_of_budget:
            if not args.no_discover:
                await finder.discover_keywords(args.keywords or DEFAULT_KEYWORDS)
                await finder.discover_hashtags(args.hashtags or DEFAULT_HASHTAGS)
            for q in args.post_search or []:
                await finder.discover_post_query(q)
    except errors.FloodWaitError as e:
        log.warning("Поиск каналов прерван: лимит запросов на %d с. Перехожу к обходу.", e.seconds)
    if not finder.queue and not finder.scanned:
        sys.exit("Нечего сканировать: укажите -c/--channels-file, --dialogs или не отключайте поиск.")

    if not finder.found_enough and not finder.out_of_budget:
        await finder.crawl()
    if finder.found_enough:
        log.info("Нашёл %d подходящий розыгрыш — дальше не ищу", finder.wanted)
    if finder.web is not None:
        log.info("Веб-проверка t.me: каналов %d, страниц %d, открыто через API %d, ограничений частоты (429) %d%s",
                 finder.web_checked, finder.web.pages, finder.web_opened, finder.web.limited,
                 " — t.me перестал отвечать, остальное через API" if finder.web.broken else "")
        if getattr(finder.web, "texted", 0) or getattr(finder.web, "imaged", 0):
            log.info("Заглушек без «запуска», у которых на странице поста есть текст / своя картинка "
                     "(не розыгрыш, не открывал): %d / %d", finder.web.texted, finder.web.imaged)
    if args.check:
        log.info("Проверяю возможность участия (%d розыгрышей)...", len(finder.giveaways))
        await finder.check_eligibility()
    await finder.resolve_channels()
    return finder


async def cmd_scan(client: TelegramClient, args: argparse.Namespace) -> None:
    finder = await run_finder(client, args)
    gs = filter_giveaways(finder.giveaways.values(), args)
    print_report(gs, finder.refs, finder.checked)
    export(gs, finder.refs, args.json, args.csv)
    if args.save_sources:
        save_sources(gs, args.save_sources)


async def find_root_post(client: TelegramClient, g: Giveaway) -> tuple[int, int] | None:
    """Пост розыгрыша в канале-организаторе (первый в списке каналов розыгрыша)."""
    root = g.channel_ids[0] if g.channel_ids else None
    for cid, mid in g.origins:
        if cid == root:
            return cid, mid
    if root is not None:
        try:
            ent = await client.get_entity(types.PeerChannel(root))
            async for m in client.iter_messages(ent, limit=300):
                if isinstance(m.media, types.MessageMediaGiveaway) and giveaway_key(m.media) == g.key:
                    return root, m.id
        except (ValueError, TypeError, errors.RPCError) as e:
            log.debug("Не удалось найти пост в канале-организаторе %s: %s", root, e)
    return None  # только напрямую из канала-организатора; копии и агрегаторы не берём


def load_state(path: Path) -> dict[str, Any]:
    """Состояние репоста:
    {"last_post": ISO-время | null,
     "posted": {ключ розыгрыша: дата окончания},       — что уже пересылали
     "queue": [{...}, ...]}                            — найденные, но ещё не опубликованные
    """
    st: dict[str, Any] = {"last_post": None, "posted": {}, "queue": []}
    if not path.is_file():
        return st
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except ValueError:
        return st
    if isinstance(data, dict) and isinstance(data.get("posted"), dict):
        st.update(data)  # сохраняем и служебные поля (our_ids, post_times …)
        st["queue"] = st.get("queue") or []
    elif isinstance(data, dict):  # старый формат: только posted
        st["posted"] = data
    return st


def save_state(path: Path, st: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(st, ensure_ascii=False, indent=1), encoding="utf-8")


def post_is_due(last_post: datetime | None, interval_hours: float, now: datetime,
                slack: timedelta = timedelta(minutes=10)) -> tuple[bool, datetime | None]:
    """Пора ли постить: прошло ли interval_hours с последнего поста (с небольшим допуском,
    т.к. запуски по расписанию GitHub приходят с опозданием)."""
    if not interval_hours or last_post is None:
        return True, None
    next_at = last_post + timedelta(hours=interval_hours)
    return now >= next_at - slack, next_at


def is_manual_post(m: Any) -> bool:
    """Обычный пост в канале (не служебное сообщение вроде «закреплено» или «канал создан»).
    Запуск официального розыгрыша — тоже пост."""
    action = getattr(m, "action", None)
    return action is None or isinstance(action, types.MessageActionGiveawayLaunch)


async def inspect_target(client: TelegramClient, target: Any, st: dict[str, Any],
                         limit: int = 100) -> tuple[datetime | None, dict[str, str]]:
    """Смотрит последние посты канала:
    * время последнего поста, опубликованного НЕ ботом (вручную, любого);
    * розыгрыши, которые уже есть в канале (кем бы ни были опубликованы), — чтобы не дублировать.
    Свои посты бот узнаёт по сохранённым id и времени публикации."""
    our_ids = set(st.get("our_ids") or [])
    our_times = [datetime.fromisoformat(t) for t in (st.get("post_times") or [])]
    if st.get("last_post"):
        our_times.append(datetime.fromisoformat(st["last_post"]))
    manual: datetime | None = None
    in_channel: dict[str, str] = {}
    async for m in client.iter_messages(target, limit=limit):
        media = getattr(m, "media", None)
        if isinstance(media, types.MessageMediaGiveaway) and media.until_date:
            in_channel[giveaway_key(media)] = media.until_date.isoformat()
        if manual is None and m.id not in our_ids and is_manual_post(m) \
                and not any(abs((m.date - t).total_seconds()) <= 120 for t in our_times):
            manual = m.date
    return manual, in_channel


def queue_order(e: dict[str, Any]) -> tuple:
    return (e["until"], bool(e.get("countries")))  # раньше кончится — раньше публикуем; при равенстве «для всех стран»


async def giveaway_closed_reason(client: TelegramClient, channel_id: int, access_hash: int | None,
                                 msg_id: int) -> str | None:
    """Почему в розыгрыш нельзя вступить, или None, если можно. Участник должен подписаться на ВСЕ каналы
    розыгрыша, поэтому каждый из них должен быть публичным и доступным (05.11 вышел пост, где 3-й канал
    оказался закрытым — владелец: «в нём невозможно выиграть»)."""
    try:
        msgs = await client.get_messages(types.InputPeerChannel(channel_id, access_hash or 0), ids=[msg_id])
    except errors.FloodWaitError:
        raise
    except (errors.RPCError, ValueError, TypeError) as e:
        return f"пост недоступен ({e.__class__.__name__})"
    media = getattr(msgs[0] if msgs else None, "media", None)
    if not isinstance(media, types.MessageMediaGiveaway):
        return "пост больше не розыгрыш"
    for cid in media.channels or []:
        try:
            ch = await client.get_entity(types.PeerChannel(cid))  # каналы приходят вместе с постом — без поиска по имени
        except errors.FloodWaitError:
            raise
        except (ValueError, TypeError, errors.RPCError) as e:
            return f"канал-участник {cid} недоступен ({e.__class__.__name__})"
        if not isinstance(ch, types.Channel):
            return f"канал-участник {cid} недоступен"
        if not active_username(ch):
            return f"канал-участник «{ch.title}» закрытый (без @username)"
        if getattr(ch, "restricted", False):
            return f"канал-участник @{active_username(ch)} заблокирован"
    return None


async def queue_entry(client: TelegramClient, g: Giveaway, finder: "GiveawayFinder") -> dict[str, Any] | None:
    """Готовая к пересылке запись: пост в канале-организаторе + access_hash, чтобы потом
    переслать без поиска (в т.ч. из новой сессии на другом сервере)."""
    root = await find_root_post(client, g)
    if root is None:
        return None
    cid, mid = root
    try:
        peer = await client.get_input_entity(types.PeerChannel(cid))
    except (ValueError, TypeError):
        return None
    ref = await finder.channel_ref(cid)
    if not ref.username:  # закрытый канал (без @username) — пропускаем
        log.info("Пропускаю: канал-организатор закрытый (%s)", ref.title or cid)
        return None
    why = await giveaway_closed_reason(client, cid, getattr(peer, "access_hash", None), mid)
    if why:
        log.info("Пропускаю %s: %s — в такой розыгрыш нельзя вступить", message_link(ref, cid, mid), why)
        return None
    return {"key": g.key, "until": g.until_date.isoformat(), "prize": g.prize_text(),
            "countries": g.countries, "channel_id": cid,
            "access_hash": getattr(peer, "access_hash", None), "msg_id": mid,
            "link": message_link(ref, cid, mid)}


async def get_target(client: TelegramClient, args: argparse.Namespace, st: dict[str, Any]) -> Any:
    """Наш канал: из учёта (без запросов), иначе по имени, а если поиск по имени временно ограничен
    Telegram — из списка чатов аккаунта (он админ канала). Поиск по имени у Telegram строго лимитирован."""
    want = str(args.to or "").lstrip("@").lower()
    cached = st.get("target_peer") or {}
    if cached.get("username") == want and cached.get("id") and cached.get("hash"):
        return types.InputPeerChannel(int(cached["id"]), int(cached["hash"]))
    try:
        ent = await client.get_entity(args.to)
    except errors.FloodWaitError as e:
        log.warning("Поиск по имени ограничен Telegram ещё на %d мин — беру @%s из списка чатов",
                    e.seconds // 60, want)
        ent = None
        async for d in client.iter_dialogs():
            u = active_username(d.entity)
            if isinstance(d.entity, types.Channel) and u and u.lower() == want:
                ent = d.entity
                break
        if ent is None:
            raise
    if isinstance(ent, types.Channel) and ent.access_hash:
        st["target_peer"] = {"username": want, "id": ent.id, "hash": ent.access_hash}
    return ent


async def cmd_post_scheduled(client: TelegramClient, args: argparse.Namespace) -> None:
    """Режим отложенных постов: в канале всегда стоят --schedule-ahead запланированных постов бота
    с шагом --min-interval-hours. Публикует их сам Telegram точно по времени, поэтому пропущенные
    запуски по расписанию GitHub не сбивают ритм. Ваши посты (опубликованные и отложенные)
    учитываются: посты бота сдвигаются на --manual-gap-hours после них."""
    state_path = Path(args.state)
    st = load_state(state_path)
    now = datetime.now(timezone.utc)
    interval = timedelta(hours=args.min_interval_hours or 0)
    gap = timedelta(hours=args.manual_gap_hours or 0)
    min_left = timedelta(hours=args.min_hours_left)

    target = await get_target(client, args, st)
    manual, in_channel = await inspect_target(client, target, st)

    posted = {k: v for k, v in st["posted"].items() if datetime.fromisoformat(v) > now}
    for k, v in in_channel.items():
        if datetime.fromisoformat(v) > now:
            posted.setdefault(k, v)
    scheduled: dict[str, dict[str, Any]] = {e["key"]: e for e in (st.get("scheduled") or [])}
    # розыгрыши, которые переслать нельзя (закрытый канал, запрет пересылки…) — больше не берём
    bad = {k: v for k, v in (st.get("bad") or {}).items() if datetime.fromisoformat(v) > now}
    post_times = [datetime.fromisoformat(t) for t in (st.get("post_times") or [])]
    our_ids: list[int] = list(st.get("our_ids") or [])
    queue: list[dict[str, Any]] = list(st["queue"])

    if getattr(args, "post_now", None):  # вне очереди: опубликовать конкретный розыгрыш из запаса сейчас
        entry = next((e for e in queue if e["link"].rstrip("/") == args.post_now.rstrip("/")), None)
        if entry is None:
            log.warning("Розыгрыша %s нет в запасе — нечего публиковать вне очереди", args.post_now)
            return
        why = await giveaway_closed_reason(client, entry["channel_id"], entry["access_hash"], entry["msg_id"])
        if why:
            log.warning("Не публикую %s: %s — в такой розыгрыш нельзя вступить", entry["link"], why)
            st.setdefault("bad", {})[entry["key"]] = entry["until"]
            st["queue"] = [e for e in queue if e is not entry]
            save_state(state_path, st)
            return
        fwd = await client.forward_messages(
            target, entry["msg_id"], from_peer=types.InputPeerChannel(entry["channel_id"], entry["access_hash"]))
        if getattr(fwd, "id", None) is not None:
            our_ids.append(fwd.id)  # свой пост — не принимать за ручной
        posted[entry["key"]] = entry["until"]
        queue.remove(entry)
        # расписание (last_post / post_times) не трогаем: пост вне очереди
        st.update(posted=posted, queue=queue, our_ids=our_ids[-100:])
        save_state(state_path, st)
        log.info("✅ Вне очереди: %s | %s", entry["prize"], entry["link"])
        return

    # что сейчас реально стоит в отложенных канала
    live: dict[str, Any] = {}
    for m in await client.get_messages(target, scheduled=True, limit=100):
        media = getattr(m, "media", None)
        k = giveaway_key(media) if isinstance(media, types.MessageMediaGiveaway) else None
        if k in scheduled:
            live[k] = m
        elif is_manual_post(m):  # ваш собственный отложенный пост — после него тоже пауза
            manual = max(manual, m.date) if manual else m.date
    # наши отложенные, которых больше нет: опубликованы или удалены вами — в любом случае не повторяем
    for k in [k for k in scheduled if k not in live]:
        posted[k] = scheduled.pop(k)["until"]

    # наш отложенный пост попал в паузу после вашего поста — переносим
    if manual and gap:
        moved = [k for k, m in live.items() if manual <= m.date < manual + gap]
        if moved and not args.dry_run:
            await client(functions.messages.DeleteScheduledMessagesRequest(
                peer=target, id=[live[k].id for k in moved]))
            for k in moved:
                e = scheduled.pop(k)
                at = datetime.fromisoformat(e["at"])
                post_times = [t for t in post_times if abs((t - at).total_seconds()) > 120]
                queue.insert(0, {kk: vv for kk, vv in e.items() if kk != "at"})
            log.info("Перенёс отложенных постов: %d (ваш пост от %s UTC)", len(moved), manual.strftime("%d.%m %H:%M"))

    times = [datetime.fromisoformat(e["at"]) for e in scheduled.values()] + [t for t in post_times if t <= now]
    if st.get("last_post"):
        times.append(datetime.fromisoformat(st["last_post"]))
    last_bot = max(times, default=None)
    next_at = max([now] + ([last_bot + interval] if last_bot else []) + ([manual + gap] if manual and gap else []))
    window = parse_post_hours(getattr(args, "post_hours", None))
    tz_hours = getattr(args, "tz_offset", 3.0)
    next_at = in_post_window(next_at, window, tz_hours)  # ночью не публикуем — переносим на утро

    def fits(e: dict[str, Any], when: datetime) -> bool:
        return (e["key"] not in posted and e["key"] not in scheduled and e["key"] not in bad
                and datetime.fromisoformat(e["until"]) - when >= min_left)

    scans, finder, last_published = 0, None, None

    def need_more() -> bool:
        if args.refill:
            return False
        if args.schedule_ahead > 0:  # держать N отложенных постов
            return len(scheduled) < args.schedule_ahead
        # 0 — публиковать один пост, только когда подошло время (раз в 3,5 ч / 5 ч после ваших постов)
        # допуск 10 минут: запуски по расписанию приходят неточно, пропускать из-за пары минут нельзя
        # (но не раньше начала окна публикаций: в 06:55 людей не будим)
        return (last_published is None and next_at <= now + timedelta(minutes=10)
                and in_post_window(now, window, tz_hours) == now)

    async def search() -> None:
        """Перебор каналов: набираем в запас до --reserve подходящих розыгрышей."""
        nonlocal finder
        log.info("Ищу новые розыгрыши перебором каналов (запас до %d)", args.reserve)
        known = {e["key"] for e in queue}
        want_at = next_at

        def want(g: Giveaway) -> bool:
            return (bool(filter_giveaways([g], args)) and g.key not in posted
                    and g.key not in scheduled and g.key not in known and g.key not in bad
                    and g.until_date - want_at >= min_left)
        # смена файла большого списка — начинаем его с начала
        list_path = getattr(args, "discover_file", None)
        list_id = None
        if list_path and Path(list_path).is_file():
            list_id = f"{list_path}:{sum(1 for ln in Path(list_path).read_text(encoding='utf-8').splitlines() if ln.strip())}"
        same_list = list_id is not None and st.get("discover_file") == list_id
        args.discover_offset = int(st.get("discover_offset", 0) or 0) if same_list else 0
        args.seen_stubs = st.get("seen_stubs") or {}
        args.peer_cache = st.get("peer_cache") or {}
        finder = await run_finder(client, args, want=want, want_count=args.reserve)
        st["peer_cache"] = dict(list(finder.peer_cache.items())[-30000:])
        # учёт проверенных заглушек: не больше 5000 каналов, по 60 последних постов на канал
        st["seen_stubs"] = {k: sorted(v)[-60:] for k, v in list(finder.seen_stubs.items())[-5000:]}
        if finder.discover_total:
            st["discover_file"] = list_id
            st["discover_offset"] = (finder.discover_start + finder.discover_consumed) % finder.discover_total
            log.info("Большой список: разобрано %d каналов (с №%d из %d), открыто через API %d, "
                     "активных розыгрышей в них %d; следующий поиск — с №%d",
                     finder.discover_consumed, finder.discover_start + 1, finder.discover_total,
                     finder.discover_opened, finder.discover_found, st["discover_offset"] + 1)
        allowed = parse_countries(getattr(args, "countries", None))
        why: dict[str, int] = {}
        added = []

        def skip(reason: str) -> None:
            why[reason] = why.get(reason, 0) + 1
        for g in finder.giveaways.values():
            if g.key in posted or g.key in scheduled or g.key in known or g.key in bad:
                skip("уже был / в канале")
            elif not country_ok(g, allowed):
                skip("другие страны")
            elif not filter_giveaways([g], args):
                skip("мал приз")
            elif g.until_date - now < min_left:
                skip("скоро кончится")
            else:
                e = await queue_entry(client, g, finder)
                if e:
                    queue.append(e)
                    added.append(e)
                else:  # нет публичного поста организатора — больше не берём
                    bad[g.key] = g.until_date.isoformat()
                    skip("закрытый канал")
        queue.sort(key=queue_order)
        st["last_search"] = now.isoformat()
        st["last_search_added"] = len(added)
        web = (f" (по веб-странице {finder.web_checked}, из них открыто через API {finder.web_opened}; "
               f"целиком через API {finder.api_read})") if finder.web is not None else ""
        log.info("Итог поиска: проверено каналов %d%s, активных розыгрышей %d, в запас добавлено %d; отсеяно: %s",
                 finder.checked, web, len(finder.giveaways), len(added),
                 ", ".join(f"{k} — {v}" for k, v in why.items()) or "нет")
        for e in added:
            log.info("  + в запас: %s | до %s | %s", e["prize"], e["until"][:16], e["link"])

    if args.refill:  # только пополнить запас, ничего не публикуя
        await search()
    elif args.schedule_ahead == 0 and not need_more():
        log.info("Ещё рано: следующий пост не раньше %s UTC — просто обновился", next_at.strftime("%d.%m %H:%M"))
    while need_more():
        queue = sorted((e for e in queue if e["key"] not in posted and e["key"] not in scheduled
                        and e["key"] not in bad and datetime.fromisoformat(e["until"]) > now),
                       key=queue_order)
        entry = next((e for e in queue if fits(e, next_at)), None)
        if entry is None:
            if scans >= 3:  # не больше трёх поисков за запуск
                break
            last = st.get("last_search")
            if (scans == 0 and last and not st.get("last_search_added")
                    and now - datetime.fromisoformat(last) < timedelta(hours=args.search_cooldown_hours)):
                log.info("Запас пуст, но поиск в %s UTC ничего нового не дал — следующий поиск не раньше чем "
                         "через %g ч (бережём аккаунт)", datetime.fromisoformat(last).strftime("%H:%M"),
                         args.search_cooldown_hours)
                break
            scans += 1
            await search()
            if not st.get("last_search_added"):  # ничего нового — повторять в этом запуске бессмысленно
                break
            continue
        queue.remove(entry)
        when = next_at
        immediate = when <= now + timedelta(minutes=2) or args.schedule_ahead == 0
        if args.dry_run:
            print(f"[dry-run] {'сейчас' if immediate else when.strftime('%d.%m %H:%M UTC')}: "
                  f"{entry['prize']} | {entry['link']}")
            scheduled[entry["key"]] = {**entry, "at": when.isoformat()}
            next_at = when + interval
            continue
        why = await giveaway_closed_reason(client, entry["channel_id"], entry["access_hash"], entry["msg_id"])
        if why:  # перед публикацией ещё раз: каналы-участники могли закрыться, пока розыгрыш лежал в запасе
            log.warning("Не публикую %s: %s — в такой розыгрыш нельзя вступить, беру другой", entry["link"], why)
            bad[entry["key"]] = entry["until"]
            continue
        try:
            fwd = await client.forward_messages(
                target, entry["msg_id"],
                from_peer=types.InputPeerChannel(entry["channel_id"], entry["access_hash"]),
                schedule=None if immediate else when)
        except errors.RPCError as ex:  # закрытый канал, запрет пересылки, пост удалён и т.п.
            log.warning("Не удалось переслать %s: %s — больше его не беру, ищу другой", entry["link"], ex)
            bad[entry["key"]] = entry["until"]
            continue
        if immediate:
            posted[entry["key"]] = entry["until"]
            if getattr(fwd, "id", None) is not None:
                our_ids.append(fwd.id)
            post_times.append(now)
            last_published = now
            log.info("✅ Переслано сейчас: %s | %s", entry["prize"], entry["link"])
        else:
            scheduled[entry["key"]] = {**entry, "at": when.isoformat()}
            post_times.append(when)
            log.info("🕒 Запланировано на %s UTC: %s | %s", when.strftime("%d.%m %H:%M"), entry["prize"], entry["link"])
        next_at = in_post_window(when + interval, window, tz_hours)

    # раз в сутки — поиск, даже если запас не пуст (владелец 06.11: агрегаторы проверять раз в день)
    daily = getattr(args, "daily_search_hours", 0) or 0
    last = st.get("last_search")
    left = [e for e in queue if e["key"] not in posted and e["key"] not in scheduled
            and e["key"] not in bad and datetime.fromisoformat(e["until"]) > now]
    if (daily and not args.refill and scans == 0 and len(left) < args.reserve
            and (not last or now - datetime.fromisoformat(last) >= timedelta(hours=daily))):
        log.info("Поиска не было %g ч — проверяю агрегаторы и пополняю запас (в запасе %d из %d)",
                 daily, len(left), args.reserve)
        queue = left
        await search()

    if not args.dry_run:
        st.update(posted=posted, queue=queue, scheduled=list(scheduled.values()), bad=bad,
                  post_times=[t.isoformat() for t in post_times][-50:], our_ids=our_ids[-100:],
                  last_post=(last_published.isoformat() if last_published else st.get("last_post")))
        save_state(state_path, st)
    if args.save_sources and finder is not None:
        save_sources(filter_giveaways(finder.giveaways.values(), args), args.save_sources)
    log.info("Отложено постов бота: %d, в очереди: %d", len(scheduled), len(queue))


async def cmd_post(client: TelegramClient, args: argparse.Namespace) -> None:
    if getattr(args, "schedule_ahead", None) is not None:
        return await cmd_post_scheduled(client, args)
    state_path = Path(args.state)
    st = load_state(state_path)
    now = datetime.now(timezone.utc)
    last_post = datetime.fromisoformat(st["last_post"]) if st["last_post"] else None
    due, next_at = post_is_due(last_post, args.min_interval_hours, now)
    if not due and not (args.dry_run or args.refill):
        log.info("Ещё рано: последний пост %s, следующий не раньше %s (UTC)",
                 last_post.strftime("%d.%m %H:%M"), next_at.strftime("%d.%m %H:%M"))
        return
    target = await get_target(client, args, st)
    manual, in_channel = await inspect_target(client, target, st)
    if args.manual_gap_hours and not args.refill:
        due_m, next_m = post_is_due(manual, args.manual_gap_hours, now)
        if manual and not due_m and not args.dry_run:
            log.info("В канале есть ваш пост от %s — после него ждём %g ч, следующий не раньше %s (UTC)",
                     manual.strftime("%d.%m %H:%M"), args.manual_gap_hours, next_m.strftime("%d.%m %H:%M"))
            return
    # завершившиеся забываем, чтобы файл не рос бесконечно
    posted = {k: v for k, v in st["posted"].items() if datetime.fromisoformat(v) > now}
    # уже есть в канале (в т.ч. опубликованные вами вручную) — не дублируем
    for k, v in in_channel.items():
        if datetime.fromisoformat(v) > now:
            posted.setdefault(k, v)
    min_left = timedelta(hours=args.min_hours_left)

    def usable(e: dict[str, Any]) -> bool:
        return e["key"] not in posted and datetime.fromisoformat(e["until"]) - now >= min_left

    queue = sorted((e for e in st["queue"] if usable(e)), key=queue_order)
    sent = 0
    our_ids: list[int] = list(st.get("our_ids") or [])

    async def publish(entries: list[dict[str, Any]]) -> list[dict[str, Any]]:
        nonlocal sent, target
        rest = []
        for e in entries:
            if args.refill or sent >= args.count:
                rest.append(e)
                continue
            if args.dry_run:
                print(f"[dry-run] переслал бы: {e['prize']} | {e['link']}")
                sent += 1
                rest.append(e)
                continue
            try:
                peer = types.InputPeerChannel(e["channel_id"], e["access_hash"])
                fwd = await client.forward_messages(target, e["msg_id"], from_peer=peer)
                if getattr(fwd, "id", None) is not None:
                    our_ids.append(fwd.id)  # чтобы не принять свой пост за ручной
            except errors.RPCError as ex:  # канал запретил пересылку, пост удалён и т.п.
                log.warning("Не удалось переслать %s: %s — убираю из очереди", e["link"], ex)
                continue
            posted[e["key"]] = e["until"]
            sent += 1
            log.info("✅ Переслано: %s | %s", e["prize"], e["link"])
        return rest

    if queue and not args.refill:
        log.info("Готовых розыгрышей в очереди: %d — беру оттуда, без сканирования", len(queue))
        queue = await publish(queue)

    finder = None
    if args.refill or sent < args.count:
        log.info("Очередь пуста — ищу новые розыгрыши")
        finder = await run_finder(client, args)
        known = {e["key"] for e in queue}
        new = []
        for g in filter_giveaways(finder.giveaways.values(), args):
            if g.key in posted or g.key in known or g.until_date - now < min_left:
                continue
            e = await queue_entry(client, g, finder)
            if e:
                new.append(e)
        log.info("Новых розыгрышей в очередь: %d", len(new))
        queue = await publish(sorted(queue + new, key=queue_order))

    if not args.dry_run:
        post_times = list(st.get("post_times") or [])
        if sent:
            post_times.append(now.isoformat())
        st.update(posted=posted, queue=queue, our_ids=our_ids[-100:], post_times=post_times[-50:],
                  last_post=now.isoformat() if sent else st["last_post"])
        save_state(state_path, st)
    if args.save_sources and finder is not None:
        save_sources(filter_giveaways(finder.giveaways.values(), args), args.save_sources)
    if sent == 0 and not args.refill:
        log.warning("Нечего пересылать: подходящих новых розыгрышей не найдено")
    log.info("В очереди осталось: %d", len(queue))


async def cmd_watch(client: TelegramClient, args: argparse.Namespace) -> None:
    """Слушает все ваши чаты в реальном времени и сообщает о новых розыгрышах."""
    finder = GiveawayFinder(client, argparse.Namespace(depth=0, delay=(0.5, 1.0)))

    @client.on(events.NewMessage())
    async def on_message(event: events.NewMessage.Event) -> None:
        msg = event.message
        if not isinstance(msg.media, types.MessageMediaGiveaway):
            return
        finder.now = datetime.now(timezone.utc)
        before = len(finder.giveaways)
        cid = utils.get_peer_id(msg.peer_id, add_mark=False)
        g = finder.handle_message(msg, cid, depth=0)
        if g is None or len(finder.giveaways) == before:
            return  # завершён или уже видели
        g.channels = [await finder.channel_ref(c) for c in g.channel_ids]
        try:
            ocid, omid = g.origins[0]
            info = await client(functions.payments.GetGiveawayInfoRequest(
                peer=await client.get_input_entity(types.PeerChannel(ocid)), msg_id=omid))
            if isinstance(info, types.payments.GiveawayInfo):
                g.status, g.eligible = describe_info(info)
        except (ValueError, TypeError, errors.RPCError):
            pass
        text = (f"🎁 Новый розыгрыш: {g.prize_text()}\n"
                f"До: {g.until_date.astimezone():%Y-%m-%d %H:%M} (осталось {g.time_left()})\n"
                f"Каналы: " + ", ".join(c.label() for c in g.channels) +
                (f"\nСтатус: {g.status}" if g.status else "") +
                f"\n{first_link(g, finder.refs)}")
        print(text, "\n" + "-" * 78, flush=True)
        if args.notify:
            await msg.forward_to("me")
            await client.send_message("me", text, link_preview=False)

    log.info("Слежу за новыми розыгрышами во всех ваших чатах. Ctrl+C — выход.")
    await client.run_until_disconnected()


def add_scan_args(s: argparse.ArgumentParser) -> None:
    src = s.add_argument_group("откуда брать каналы")
    src.add_argument("-c", "--channel", action="append", metavar="@USERNAME",
                     help="канал для сканирования (можно несколько раз, можно t.me-ссылку)")
    src.add_argument("--channels-file", action="append", metavar="FILE",
                     help="файл со списком каналов, по одному в строке")
    src.add_argument("--dialogs", action="store_true", help="сканировать все ваши каналы и супергруппы")
    src.add_argument("--no-discover", action="store_true",
                     help="не искать новые каналы по ключевым словам и хэштегам")
    src.add_argument("--no-aggregators", action="store_true",
                     help="не сканировать встроенный список каналов-агрегаторов розыгрышей")
    src.add_argument("--keywords", nargs="+", metavar="KW", help="свои ключевые слова для поиска каналов")
    src.add_argument("--hashtags", nargs="+", metavar="TAG", help="свои хэштеги для поиска постов")
    src.add_argument("--search-pages", type=int, help="страниц (по 100) на каждый хэштег (3; --careful: 1)")
    src.add_argument("--post-search", action="append", metavar="QUERY",
                     help="полнотекстовый поиск публичных постов (тратит бесплатный дневной лимит)")

    lim = s.add_argument_group("обход")
    lim.add_argument("--careful", action="store_true",
                     help="осторожный режим: пауза 30 с после каждых 500 каналов, "
                          "FloodWait до 5 минут пережидается, а не обрывает обход")
    lim.add_argument("--depth", type=int, default=2,
                     help="глубина «снежного кома» по каналам-участникам розыгрышей (2)")
    lim.add_argument("--max-channels", type=int, help="максимум каналов за запуск (2000)")
    lim.add_argument("--days", type=int, default=31, help="насколько старые посты смотреть, дней (31)")
    lim.add_argument("--per-channel-limit", type=int, help="макс. сообщений на канал (400)")
    lim.add_argument("--source-limit", type=int, default=0, metavar="N",
                     help="стартовые каналы (агрегаторы, свой список) читать на N последних записей "
                          "без ограничения по дате — чтобы найти организаторов и старых розыгрышей")
    lim.add_argument("--batch-size", type=int, help="размер пачки каналов для --batch-pause (500)")
    lim.add_argument("--batch-pause", type=float,
                     help="пауза в секундах после каждой пачки каналов (0; --careful: 30)")
    lim.add_argument("--delay", metavar="СЕК[-СЕК]",
                     help="пауза между отдельными каналами, напр. 1 или 1-3 (по умолчанию нет)")
    lim.add_argument("--page-delay", metavar="СЕК[-СЕК]",
                     help="пауза между страницами истории одного канала (по умолчанию нет)")
    lim.add_argument("--web", action="store_true",
                     help="сначала смотреть публичную веб-страницу канала t.me/s/…: модуль-розыгрыш там — "
                          "заглушка «Please open Telegram», через API открываются только такие посты")
    lim.add_argument("--web-max", type=int, default=5000, help="максимум каналов веб-проверкой за запуск (5000)")
    lim.add_argument("--web-pages", type=int, default=20, help="страниц (по ~20 постов) на канал при веб-проверке (20)")
    lim.add_argument("--web-parallel", type=int, default=8, help="одновременных запросов к t.me (8)")
    lim.add_argument("--discover-file", metavar="FILE",
                     help="большой список @username (напр. из Common Crawl): только веб-проверка, "
                          "без веб-страницы — пропуск; разбирается по кругу с --discover-offset")
    lim.add_argument("--discover-offset", type=int, default=0, help="с какой строки списка начать (0)")
    lim.add_argument("--discover-days", type=int, default=7,
                     help="каналы большого списка: смотреть только посты не старше N дней (7)")
    lim.add_argument("--discover-gaps", action="store_true",
                     help="каналы большого списка: открывать и заглушки без «Канал запустил розыгрыш» "
                          "(пропущен номер перед ними) — партиями по --web-open-max")
    lim.add_argument("--web-open-max", type=int, default=80,
                     help="сколько каналов большого списка можно открыть через API за запуск (80)")
    lim.add_argument("--time-limit", type=float, default=0, metavar="МИН",
                     help="остановить обход через столько минут и взять найденное (0 — без ограничения)")
    lim.add_argument("--max-flood-wait", type=int,
                     help="автоматически ждать FloodWait до N секунд, дольше — остановиться "
                          "(120; --careful: 300)")

    out = s.add_argument_group("результат")
    out.add_argument("--check", action="store_true",
                     help="проверить через API, можете ли вы участвовать (страна, новые подписчики...)")
    out.add_argument("--prize", choices=["any", "premium", "stars"], default="any", help="тип приза")
    out.add_argument("--min-winners", type=int, default=1, help="минимум победителей")
    out.add_argument("--min-stars", type=int, default=0, metavar="N",
                     help="розыгрыши Stars: минимум звёзд на весь розыгрыш (напр. 10000)")
    out.add_argument("--min-premium", type=int, default=0, metavar="N",
                     help="розыгрыши Premium: минимум подписок (победителей), напр. 5")
    out.add_argument("--countries", metavar="RU,BY,KZ",
                     help="оставить розыгрыши без ограничений по странам и для этих стран "
                          "(«для всех стран» показываются первыми)")
    out.add_argument("--hide-ineligible", action="store_true", help="скрыть те, где участвовать нельзя (с --check)")
    out.add_argument("--json", metavar="FILE", help="сохранить результат в JSON")
    out.add_argument("--csv", metavar="FILE", help="сохранить результат в CSV")
    out.add_argument("--save-sources", metavar="FILE",
                     help="дописать каналы с розыгрышами в файл (для следующего --channels-file)")

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Поиск активных официальных розыгрышей Telegram (Premium / Stars).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("-v", "--verbose", action="store_true", help="подробный лог")
    sub = p.add_subparsers(dest="command", required=True)

    s = sub.add_parser("scan", help="разовый поиск розыгрышей")
    add_scan_args(s)

    ps = sub.add_parser("post", help="найти розыгрыши и переслать лучшие (из канала-организатора) в ваш чат")
    add_scan_args(ps)
    ps.add_argument("--to", required=True, metavar="@ЧАТ",
                    help="куда пересылать: @username, ссылка t.me/... или числовой id группы/канала")
    ps.add_argument("--count", type=int, default=1, help="сколько розыгрышей переслать за запуск (1)")
    ps.add_argument("--state", default="posted.json", metavar="FILE",
                    help="файл с уже пересланными розыгрышами, чтобы не повторяться (posted.json)")
    ps.add_argument("--min-hours-left", type=float, default=3.0,
                    help="не пересылать розыгрыши, до конца которых меньше N часов (3)")
    ps.add_argument("--search-cooldown-hours", type=float, default=2.0, metavar="Ч",
                    help="если поиск ничего не дал, следующий — не раньше чем через Ч часов (2)")
    ps.add_argument("--post-hours", metavar="С-ДО",
                    help="публиковать только в эти часы по местному времени, например 7-22 (ночью людей не будим)")
    ps.add_argument("--tz-offset", type=float, default=3.0, metavar="Ч",
                    help="местное время для --post-hours: смещение от UTC в часах (3 — Минск/Москва)")
    ps.add_argument("--daily-search-hours", type=float, default=0, metavar="Ч",
                    help="искать и пополнять запас, если поиска не было Ч часов, даже когда запас не пуст "
                         "(0 — только когда запас кончился)")
    ps.add_argument("--post-now", metavar="ССЫЛКА",
                    help="опубликовать этот розыгрыш из запаса сейчас, вне очереди (расписание не сдвигается)")
    ps.add_argument("--reserve", type=int, default=10, metavar="N",
                    help="сколько подходящих розыгрышей набирать в запас за один поиск (10); "
                         "новый поиск — только когда запас кончится")
    ps.add_argument("--min-interval-hours", type=float, default=0, metavar="Ч",
                    help="не постить чаще, чем раз в Ч часов (напр. 3.5); при раннем запуске "
                         "выходит сразу, ничего не сканируя")
    ps.add_argument("--manual-gap-hours", type=float, default=0, metavar="Ч",
                    help="если в канале есть пост, опубликованный не ботом (вручную), "
                         "следующий пост бота — не раньше чем через Ч часов после него (напр. 5)")
    ps.add_argument("--schedule-ahead", type=int, default=None, metavar="N",
                    help="0 — публиковать один пост, когда подошло время (поиск одного розыгрыша, "
                         "учёт, пропуск непригодных); N>0 — держать в канале N отложенных постов")
    ps.add_argument("--dry-run", action="store_true", help="только показать, что было бы переслано")
    ps.add_argument("--refill", action="store_true",
                    help="только найти розыгрыши и пополнить очередь, ничего не публикуя")

    w = sub.add_parser("watch", help="следить за новыми розыгрышами в ваших чатах в реальном времени")
    w.add_argument("--notify", action="store_true", help="пересылать найденные розыгрыши в «Избранное»")

    lg = sub.add_parser("login", help="вход по шагам (для серверов/облака): --phone, затем --code")
    g = lg.add_mutually_exclusive_group()
    g.add_argument("--phone", help="номер телефона, на него придёт код в Telegram")
    g.add_argument("--code", help="код из Telegram")
    g.add_argument("--password", action="store_true", help="пароль 2FA из переменной TG_2FA_PASSWORD")
    sub.add_parser("logout", help="завершить сессию в Telegram и удалить её файл")
    return p


async def cmd_login(client: TelegramClient, args: argparse.Namespace) -> None:
    """Вход по шагам без интерактивного ввода: --phone, затем --code, затем (если есть 2FA) --password."""
    state = login_state_path()
    if await client.is_user_authorized():
        me = await client.get_me()
        print(f"✅ Уже выполнен вход: {me.first_name or ''} (@{me.username or '—'})")
        state.unlink(missing_ok=True)
        return
    if args.phone:
        sent = await client.send_code_request(args.phone)
        state.write_text(json.dumps({"phone": args.phone, "hash": sent.phone_code_hash}), encoding="utf-8")
        print("📨 Telegram отправил код входа в приложение. Дальше: login --code КОД")
        return
    if not state.is_file():
        sys.exit("Сначала: login --phone +79991234567")
    st = json.loads(state.read_text(encoding="utf-8"))
    try:
        if args.code:
            await client.sign_in(phone=st["phone"], code=args.code.strip(), phone_code_hash=st["hash"])
        elif args.password:
            password = os.environ.get("TG_2FA_PASSWORD")
            if not password:
                sys.exit("Пароль 2FA передаётся через переменную окружения TG_2FA_PASSWORD")
            await client.sign_in(password=password)
        else:
            sys.exit("Укажите --phone, --code или --password")
    except errors.SessionPasswordNeededError:
        print("🔐 Включена двухэтапная проверка. Дальше: TG_2FA_PASSWORD=... login --password")
        return
    except (errors.PhoneCodeInvalidError, errors.PhoneCodeExpiredError) as e:
        sys.exit(f"Код не подошёл ({type(e).__name__}). Запросите новый: login --phone ...")
    me = await client.get_me()
    state.unlink(missing_ok=True)
    print(f"✅ Вход выполнен: {me.first_name or ''} (@{me.username or '—'})")


async def cmd_logout(client: TelegramClient, args: argparse.Namespace) -> None:
    """Завершает сессию на стороне Telegram (пропадёт из «Устройств») и удаляет её файл."""
    if await client.is_user_authorized():
        await client.log_out()
    login_state_path().unlink(missing_ok=True)
    print("✅ Сессия завершена")


async def amain(args: argparse.Namespace) -> None:
    client = make_client()
    if hasattr(args, "max_flood_wait"):
        client.flood_sleep_threshold = args.max_flood_wait
    if args.command in ("login", "logout"):
        await client.connect()
        try:
            await (cmd_login if args.command == "login" else cmd_logout)(client, args)
        finally:
            await client.disconnect()
        return
    await client.start()
    try:
        if args.command == "scan":
            await cmd_scan(client, args)
        elif args.command == "post":
            await cmd_post(client, args)
        else:
            await cmd_watch(client, args)
    finally:
        await client.disconnect()


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    if args.command in ("scan", "post"):
        try:
            apply_preset(args)
        except (argparse.ArgumentTypeError, ValueError) as e:
            parser.error(str(e))
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(message)s", datefmt="%H:%M:%S",
    )
    logging.getLogger("telethon").setLevel(logging.WARNING)
    try:
        asyncio.run(amain(args))
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
