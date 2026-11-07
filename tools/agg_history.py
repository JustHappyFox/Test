"""Каналы из всей истории агрегаторов (владелец 07.11: «собирай истории каналов агрегаторов и добавляй в списки
для проверок — у них там очень большие списки тех, кто уже делал»).

Читает через Dolli историю каждого агрегатора (обычное чтение постов, по 100 за запрос, без поиска по имени —
агрегаторы берутся из запомненных в учёте), берёт все посты-розыгрыши и итоги: каналы-участники розыгрыша
(организатор и спонсоры) и канал, откуда розыгрыш переслан. @username каналов приходят в том же ответе сервера.
Пишет список в формате большого списка: «@канал  # розыгрышей последний_розыгрыш», свежие — первыми.
Дальше эти каналы проверяются только по вебу (launch_scan.yml, «Канал запустил розыгрыш» за 7 дней);
Dolli открывает канал, только если там есть свежий «запуск».
"""
import argparse
import asyncio
import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import tg_giveaway_finder as gf  # noqa: E402
from telethon import errors, functions, types  # noqa: E402


def giveaway_channels(m: types.Message) -> list[int]:
    """Каналы, связанные с постом-розыгрышем или итогами; пусто — пост не про розыгрыш."""
    media = getattr(m, "media", None)
    ids: list[int] = []
    if isinstance(media, types.MessageMediaGiveaway):
        ids += list(media.channels or [])
    elif isinstance(media, types.MessageMediaGiveawayResults):
        ids.append(media.channel_id)
    else:
        return []
    fwd = getattr(m, "fwd_from", None)
    if fwd and isinstance(getattr(fwd, "from_id", None), types.PeerChannel):
        ids.append(fwd.from_id.channel_id)
    return list(dict.fromkeys(ids))  # один канал на пост — один раз


async def read_history(client, peer, cap: int, pause: float, names: dict[int, str],
                       seen: dict[int, list], label: str) -> int:
    offset_id, read, hits = 0, 0, 0
    while read < cap:
        r = await client(functions.messages.GetHistoryRequest(
            peer=peer, offset_id=offset_id, offset_date=None, add_offset=0, limit=100,
            max_id=0, min_id=0, hash=0))
        for c in getattr(r, "chats", []):
            if isinstance(c, types.Channel):
                u = gf.active_username(c)
                if u:
                    names[c.id] = u
        msgs = [m for m in r.messages if isinstance(m, types.Message)]
        if not r.messages:
            break
        read += len(r.messages)
        for m in msgs:
            for cid in giveaway_channels(m):
                hits += 1
                rec = seen.setdefault(cid, [0, None])
                rec[0] += 1
                if rec[1] is None or m.date > rec[1]:
                    rec[1] = m.date
        offset_id = min(m.id for m in r.messages)
        if offset_id <= 1:
            break
        await asyncio.sleep(pause)
    print(f"  {label}: постов {read}, ссылок на каналы из розыгрышей {hits}", flush=True)
    return read


async def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--state", default=".state/posted.json", help="учёт бота: запомненные каналы (peer_cache)")
    ap.add_argument("--out", default="state/agg_organizers.txt")
    ap.add_argument("--cap", type=int, default=5000, help="не больше N постов на агрегатор (5000)")
    ap.add_argument("--pause", type=float, default=1.0, help="пауза между запросами, с (1)")
    a = ap.parse_args()

    cache: dict[str, list[int]] = {}
    p = Path(a.state)
    if p.is_file():
        try:
            cache = json.loads(p.read_text(encoding="utf-8")).get("peer_cache") or {}
        except ValueError:
            pass

    client = gf.make_client()
    await client.connect()
    if not await client.is_user_authorized():
        sys.exit("Сессия Telegram не авторизована")
    names: dict[int, str] = {}
    seen: dict[int, list] = {}
    resolved = 0
    agg_ids: set[int] = set()
    try:
        for agg in gf.DEFAULT_AGGREGATORS:
            u = agg.lstrip("@")
            hit = cache.get(u.lower())
            try:
                if hit:
                    peer = types.InputPeerChannel(int(hit[0]), int(hit[1]))
                else:
                    peer = await client.get_input_entity(u)  # поиск по имени — только если не запомнен
                    resolved += 1
                agg_ids.add(peer.channel_id)
                await read_history(client, peer, a.cap, a.pause, names, seen, agg)
            except errors.FloodWaitError as e:
                if e.seconds > 120:
                    print(f"  {agg}: Telegram ограничил частоту на {e.seconds} с — останавливаюсь", flush=True)
                    break
                await asyncio.sleep(e.seconds + 1)
            except (ValueError, TypeError, errors.RPCError) as e:
                print(f"  {agg}: недоступен ({e.__class__.__name__})", flush=True)
    finally:
        await client.disconnect()

    skip = agg_ids | {0}
    rows, no_name = [], 0
    for cid, (n, last) in seen.items():
        if cid in skip:
            continue
        u = names.get(cid)
        if not u:
            no_name += 1  # закрытый канал или сервер не прислал — по вебу не проверить
            continue
        rows.append((last or datetime.min, n, u))
    rows.sort(reverse=True)
    lines = [f"@{u}  # {n} {last.date().isoformat()}" for last, n, u in rows]
    Path(a.out).write_text("".join(x + "\n" for x in lines), encoding="utf-8")
    print(f"Итог: каналов из розыгрышей {len(seen)}, с публичным @username {len(rows)}, "
          f"без него {no_name}; поисков по имени {resolved}; список — {a.out}", flush=True)


if __name__ == "__main__":
    asyncio.run(main())
