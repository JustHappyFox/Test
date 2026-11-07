"""Новые розыгрыши по сообщению «Канал запустил розыгрыш» (владелец 07.11: «ищи по этому сообщению —
реклама не может получить от Telegram такое сообщение»). Без аккаунта, по веб-странице t.me/s.

Берём только СВОИ заглушки канала, перед которыми стоит пустое служебное «Канал запустил розыгрыш»
(тот же признак, что в поиске, strict, без второго уровня). Известные каналы пропускаем.
Останавливаемся, когда найдено --stop-after новых каналов.
"""
import argparse
import asyncio
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import tg_giveaway_finder as gf  # noqa: E402
import tg_web_preview as wp  # noqa: E402


def known_names() -> set[str]:
    known = {a.lstrip("@").lower() for a in gf.DEFAULT_AGGREGATORS} | {"tgagifts"}
    for f in ("state/cc_candidates.txt", "state/results_organizers.txt", "state/sources.txt"):
        p = Path(f)
        if p.is_file():
            for ln in p.read_text(encoding="utf-8").splitlines():
                parts = ln.split("#", 1)
                n = parts[0].strip().lstrip("@").lower()
                launches = parts[1].split()[-1] if len(parts) > 1 and parts[1].split() else "1"
                if n and (f != "state/cc_candidates.txt" or launches != "0"):
                    known.add(n)
    return known


async def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("names")
    ap.add_argument("--days", type=int, default=7)
    ap.add_argument("--shard", type=int, default=0)
    ap.add_argument("--shards", type=int, default=1)
    ap.add_argument("--stop-after", type=int, default=2)
    ap.add_argument("--parallel", type=int, default=8)
    ap.add_argument("--pages", type=int, default=15,
                    help="страниц на канал, по ~20 постов (15): у активных каналов «запуск» бывает глубже 5 страниц")
    ap.add_argument("--out", help="найденные каналы — в формате большого списка (@канал  # N дата N)")
    ap.add_argument("--live-out", help="имена с веб-страницей (живые каналы) — для следующих проходов")
    a = ap.parse_args()
    known = known_names()
    names = [ln.split("#", 1)[0].strip().lstrip("@") for ln in Path(a.names).read_text(encoding="utf-8").splitlines()]
    names = [n for n in names if n and n.lower() not in known][a.shard::a.shards]
    print(f"Имён {len(names)} (известные {len(known)} пропущены)", flush=True)
    since = datetime.now(timezone.utc) - timedelta(days=a.days)
    w = wp.WebPreview(parallel=a.parallel, max_limited=10 ** 9)
    found, live, t0 = [], [], time.time()

    async def one(name):
        try:
            ids = await w.candidates(name, since, max_pages=a.pages, strict=True, gaps=False)
        except Exception:
            return
        if ids is not None:
            live.append(name)
        if ids and w.launches.get(name.lower()):
            found.append((name, ids))
            for i in ids:
                print(f"  НАЙДЕН https://t.me/{name}/{i} — перед ним «Канал запустил розыгрыш» https://t.me/{name}/{i - 1}",
                      flush=True)

    step = 200
    for i in range(0, len(names), step):
        await asyncio.gather(*(one(n) for n in names[i:i + step]))
        if (i // step) % 25 == 0:
            print(f"  проверено {min(i + step, len(names))}/{len(names)}, новых каналов с запуском: {len(found)}, "
                  f"страниц {w.pages}, 429: {w.limited}, {time.time() - t0:.0f} с", flush=True)
        if a.stop_after and len(found) >= a.stop_after:
            break
    await w.close()
    print(f"Итог: новых каналов с «Канал запустил розыгрыш» за {a.days} дн.: {len(found)}; "
          f"живых каналов {len(live)} из {len(names)}")
    if a.out:
        lines = []
        for name, ids in found:
            n = w.launches.get(name.lower(), len(ids))
            newest = w.newest.get(name.lower())
            lines.append(f"@{name}  # {n} {newest.date().isoformat() if newest else ''} {n}")
        Path(a.out).write_text("".join(x + "\n" for x in lines), encoding="utf-8")
    if a.live_out:
        Path(a.live_out).write_text("".join(f"@{x}\n" for x in live), encoding="utf-8")


if __name__ == "__main__":
    asyncio.run(main())
