#!/usr/bin/env python3
"""
Сколько настоящих розыгрышей видно по отметке «Канал запустил розыгрыш» (без аккаунта Telegram).

Берёт каналы-сборники (там почти все заглушки — пересланные розыгрыши), по ссылке «Forwarded from»
открывает оригинал на t.me/s/<канал>?before=N+1 и смотрит, стоит ли прямо перед ним пустое
служебное сообщение. Доля таких — насколько строгий отбор большого списка не теряет розыгрыши.

Пример:  python tools/launch_recall.py giveaways_database rozygrysh --days 31
"""
from __future__ import annotations

import argparse
import asyncio
import sys
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import tg_web_preview as wp  # noqa: E402


async def run(aggs: list[str], days: int, pages: int) -> None:
    w = wp.WebPreview(parallel=8, max_limited=10 ** 9)
    since = datetime.now(timezone.utc) - timedelta(days=days)
    stat: Counter = Counter()
    misses: list[str] = []
    try:
        for agg in aggs:
            url = w.BASE + agg
            sources: list[str] = []
            for _ in range(pages):
                html = await w.fetch(url)
                posts = wp.parse_page(html) if html else None
                if not posts:
                    break
                sources += [p.source for p in posts if p.stub and p.forwarded and p.source
                            and (p.date is None or p.date >= since)]
                oldest = min(posts, key=lambda p: p.id)
                if oldest.date and oldest.date < since:
                    break
                url = f"{w.BASE}{agg}?before={oldest.id}"
            sources = list(dict.fromkeys(s.lower() for s in sources))
            res = await asyncio.gather(*(w._origin_launched(s) for s in sources))
            ok = sum(res)
            stat["checked"] += len(sources)
            stat["launch"] += ok
            misses += [s for s, r in zip(sources, res) if not r]
            print(f"@{agg}: пересланных заглушек {len(sources)}, у оригинала есть «запуск» — {ok}", flush=True)
    finally:
        await w.close()
    n = stat["checked"]
    print(f"Итог: {stat['launch']} из {n} ({100 * stat['launch'] / max(n, 1):.0f}%); страниц {w.pages}, 429: {w.limited}")
    print("Без отметки (первые 40):", " ".join(f"t.me/{s}" for s in misses[:40]))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("aggregators", nargs="+")
    ap.add_argument("--days", type=int, default=31)
    ap.add_argument("--pages", type=int, default=10)
    a = ap.parse_args()
    asyncio.run(run([x.lstrip("@") for x in a.aggregators], a.days, a.pages))


if __name__ == "__main__":
    main()
