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
import json
import logging
import os
import random
import sys
from collections import deque
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable

from telethon import TelegramClient, errors, events, functions, types, utils

log = logging.getLogger("giveaway_finder")

DEFAULT_KEYWORDS = [
    "розыгрыш", "розыгрыши", "розыгрыш premium", "розыгрыш звезд",
    "giveaway", "giveaways", "premium giveaway", "stars giveaway",
    "раздача", "конкурс", "халява", "бесплатный премиум",
]
DEFAULT_HASHTAGS = ["розыгрыш", "giveaway", "конкурс", "раздача", "розыгрыши", "giveaways"]

# Публичные каналы-агрегаторы, которые пересылают официальные розыгрыши (найдены через веб-поиск
# и TGStat). Их надёжность не важна: в результат попадают только настоящие messageMediaGiveaway,
# а не текстовые «конкурсы». Отключить: --no-aggregators.
DEFAULT_AGGREGATORS = [
    "@giveaways", "@Giveaways_Premium_Tel", "@TgPremiumGiveaway",
    "@tgpremium_giveaways", "@ChannelGiveaways",
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

    @property
    def found_enough(self) -> bool:
        return bool(self.want_count) and self.wanted >= self.want_count

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

    async def resolve(self, item: Any) -> Any | None:
        try:
            if isinstance(item, (int, str)):
                ent = await self.client.get_entity(types.PeerChannel(item) if isinstance(item, int) else item)
                if not isinstance(ent, types.Channel):
                    log.warning("%s — не канал/супергруппа, пропускаю", item)
                    return None
                self.refs[ent.id] = to_ref(ent)
                return ent
            return item
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
        for s in seeds:
            self.enqueue(s, 0)
        if seeds:
            log.info("Стартовых каналов из списка: %d", len(seeds))

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
        if not isinstance(media, types.MessageMediaGiveaway):
            return None
        if not media.until_date or media.until_date <= self.now:
            return None  # уже завершён

        fwd = getattr(msg, "fwd_from", None)
        is_forward = bool(fwd and isinstance(fwd.from_id, types.PeerChannel) and fwd.channel_post)
        origin = (fwd.from_id.channel_id, fwd.channel_post) if is_forward else (chat_id, msg.id)

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
        cutoff = self.now - timedelta(days=self.args.days)
        cid = utils.get_peer_id(ent, add_mark=False)
        # wait_time — пауза между страницами истории (по 100 сообщений) одного канала
        async for msg in self.client.iter_messages(ent, limit=self.args.per_channel_limit,
                                                   wait_time=rand_pause(self.args.page_delay) or None):
            if msg.date and msg.date < cutoff:
                break
            fwd = getattr(msg, "forward", None)
            if fwd is not None and isinstance(getattr(fwd, "chat", None), types.Channel):
                self.refs.setdefault(fwd.chat.id, to_ref(fwd.chat))
            self.handle_message(msg, cid, depth)

    async def crawl(self) -> None:
        try:
            while self.queue and len(self.scanned) < self.args.max_channels and not self.found_enough:
                item, depth = self.queue.popleft()
                ent = await self.resolve(item)
                if ent is None:
                    continue
                cid = utils.get_peer_id(ent, add_mark=False)
                if cid in self.scanned:
                    continue
                self.scanned.add(cid)
                ref = self.refs.get(cid)
                log.info("[%d/%d, в очереди %d] %s", len(self.scanned), self.args.max_channels,
                         len(self.queue), ref.label() if ref else f"id:{cid}")
                try:
                    await self.scan_channel(ent, depth)
                except errors.FloodWaitError:
                    raise
                except (errors.RPCError, ValueError) as e:
                    log.debug("  пропуск: %s", e)
                await self.pause()
                batch_size = getattr(self.args, "batch_size", 0)
                batch_pause = getattr(self.args, "batch_pause", 0)
                if batch_pause and batch_size and len(self.scanned) % batch_size == 0 and self.queue:
                    log.info("Просканировано %d каналов — пауза %g с", len(self.scanned), batch_pause)
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
    try:
        await finder.add_seeds()
        if args.dialogs:
            await finder.add_dialogs()
        if want_count and finder.queue:
            await finder.crawl()
        if not finder.found_enough:
            if not args.no_discover:
                await finder.discover_keywords(args.keywords or DEFAULT_KEYWORDS)
                await finder.discover_hashtags(args.hashtags or DEFAULT_HASHTAGS)
            for q in args.post_search or []:
                await finder.discover_post_query(q)
    except errors.FloodWaitError as e:
        log.warning("Поиск каналов прерван: лимит запросов на %d с. Перехожу к обходу.", e.seconds)
    if not finder.queue and not finder.scanned:
        sys.exit("Нечего сканировать: укажите -c/--channels-file, --dialogs или не отключайте поиск.")

    if not finder.found_enough:
        await finder.crawl()
    if finder.found_enough:
        log.info("Нашёл %d подходящий розыгрыш — дальше не ищу", finder.wanted)
    if args.check:
        log.info("Проверяю возможность участия (%d розыгрышей)...", len(finder.giveaways))
        await finder.check_eligibility()
    await finder.resolve_channels()
    return finder


async def cmd_scan(client: TelegramClient, args: argparse.Namespace) -> None:
    finder = await run_finder(client, args)
    gs = filter_giveaways(finder.giveaways.values(), args)
    print_report(gs, finder.refs, len(finder.scanned))
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
    return g.origins[0] if g.origins else None


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
    return (bool(e.get("countries")), e["until"])  # «для всех стран» — первыми, затем кто раньше кончится


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
    return {"key": g.key, "until": g.until_date.isoformat(), "prize": g.prize_text(),
            "countries": g.countries, "channel_id": cid,
            "access_hash": getattr(peer, "access_hash", None), "msg_id": mid,
            "link": message_link(ref, cid, mid)}


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

    target = await client.get_entity(args.to)
    manual, in_channel = await inspect_target(client, target, st)

    posted = {k: v for k, v in st["posted"].items() if datetime.fromisoformat(v) > now}
    for k, v in in_channel.items():
        if datetime.fromisoformat(v) > now:
            posted.setdefault(k, v)
    scheduled: dict[str, dict[str, Any]] = {e["key"]: e for e in (st.get("scheduled") or [])}
    post_times = [datetime.fromisoformat(t) for t in (st.get("post_times") or [])]
    our_ids: list[int] = list(st.get("our_ids") or [])
    queue: list[dict[str, Any]] = list(st["queue"])

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
        bad = [k for k, m in live.items() if manual <= m.date < manual + gap]
        if bad and not args.dry_run:
            await client(functions.messages.DeleteScheduledMessagesRequest(
                peer=target, id=[live[k].id for k in bad]))
            for k in bad:
                e = scheduled.pop(k)
                at = datetime.fromisoformat(e["at"])
                post_times = [t for t in post_times if abs((t - at).total_seconds()) > 120]
                queue.insert(0, {kk: vv for kk, vv in e.items() if kk != "at"})
            log.info("Перенёс отложенных постов: %d (ваш пост от %s UTC)", len(bad), manual.strftime("%d.%m %H:%M"))

    times = [datetime.fromisoformat(e["at"]) for e in scheduled.values()] + [t for t in post_times if t <= now]
    if st.get("last_post"):
        times.append(datetime.fromisoformat(st["last_post"]))
    last_bot = max(times, default=None)
    next_at = max([now] + ([last_bot + interval] if last_bot else []) + ([manual + gap] if manual and gap else []))

    def fits(e: dict[str, Any], when: datetime) -> bool:
        return (e["key"] not in posted and e["key"] not in scheduled
                and datetime.fromisoformat(e["until"]) - when >= min_left)

    scanned, finder, last_published = False, None, None
    while len(scheduled) < args.schedule_ahead:
        queue = sorted((e for e in queue if e["key"] not in posted and e["key"] not in scheduled
                        and datetime.fromisoformat(e["until"]) > now), key=queue_order)
        entry = next((e for e in queue if fits(e, next_at)), None)
        if entry is None:
            if scanned:
                break
            scanned = True
            log.info("Подходящих в очереди нет — ищу один новый розыгрыш")
            known = {e["key"] for e in queue}
            want_at = next_at

            def want(g: Giveaway, known=known, want_at=want_at) -> bool:
                return (bool(filter_giveaways([g], args)) and g.key not in posted
                        and g.key not in scheduled and g.key not in known
                        and g.until_date - want_at >= min_left)
            finder = await run_finder(client, args, want=want, want_count=1)
            for g in filter_giveaways(finder.giveaways.values(), args):
                if g.key in posted or g.key in scheduled or g.key in known or g.until_date - now < min_left:
                    continue
                e = await queue_entry(client, g, finder)
                if e:
                    queue.append(e)
            continue
        queue.remove(entry)
        when = next_at
        immediate = when <= now + timedelta(minutes=2)
        if args.dry_run:
            print(f"[dry-run] {'сейчас' if immediate else when.strftime('%d.%m %H:%M UTC')}: "
                  f"{entry['prize']} | {entry['link']}")
            scheduled[entry["key"]] = {**entry, "at": when.isoformat()}
            next_at = when + interval
            continue
        try:
            fwd = await client.forward_messages(
                target, entry["msg_id"],
                from_peer=types.InputPeerChannel(entry["channel_id"], entry["access_hash"]),
                schedule=None if immediate else when)
        except errors.RPCError as ex:  # канал запретил пересылку, пост удалён и т.п.
            log.warning("Не удалось переслать %s: %s — убираю из очереди", entry["link"], ex)
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
        next_at = when + interval

    if not args.dry_run:
        st.update(posted=posted, queue=queue, scheduled=list(scheduled.values()),
                  post_times=[t.isoformat() for t in post_times][-50:], our_ids=our_ids[-100:],
                  last_post=(last_published.isoformat() if last_published else st.get("last_post")))
        save_state(state_path, st)
    if args.save_sources and finder is not None:
        save_sources(filter_giveaways(finder.giveaways.values(), args), args.save_sources)
    log.info("Отложено постов бота: %d, в очереди: %d", len(scheduled), len(queue))


async def cmd_post(client: TelegramClient, args: argparse.Namespace) -> None:
    if getattr(args, "schedule_ahead", 0) and not args.refill:
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
    target = await client.get_entity(args.to)
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
    lim.add_argument("--batch-size", type=int, help="размер пачки каналов для --batch-pause (500)")
    lim.add_argument("--batch-pause", type=float,
                     help="пауза в секундах после каждой пачки каналов (0; --careful: 30)")
    lim.add_argument("--delay", metavar="СЕК[-СЕК]",
                     help="пауза между отдельными каналами, напр. 1 или 1-3 (по умолчанию нет)")
    lim.add_argument("--page-delay", metavar="СЕК[-СЕК]",
                     help="пауза между страницами истории одного канала (по умолчанию нет)")
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
    ps.add_argument("--min-interval-hours", type=float, default=0, metavar="Ч",
                    help="не постить чаще, чем раз в Ч часов (напр. 3.5); при раннем запуске "
                         "выходит сразу, ничего не сканируя")
    ps.add_argument("--manual-gap-hours", type=float, default=0, metavar="Ч",
                    help="если в канале есть пост, опубликованный не ботом (вручную), "
                         "следующий пост бота — не раньше чем через Ч часов после него (напр. 5)")
    ps.add_argument("--schedule-ahead", type=int, default=0, metavar="N",
                    help="держать в канале N отложенных постов (Telegram опубликует их сам точно "
                         "по времени, даже если запуск по расписанию пропущен)")
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
