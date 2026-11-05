#!/usr/bin/env python3
"""
Веб-проверка большого списка каналов без аккаунта Telegram.

Для каждого @username читает публичную страницу t.me/s/<канал> (последние ~20–60 постов) и ищет
посты-заглушки без текста и без фото/видео — так сайт показывает модуль-розыгрыш. Пишет:
  * короткий список каналов с такими постами за последние --days дней (для поиска через API);
  * итог: сколько имён, у скольких есть веб-страница, у скольких есть кандидаты.

Пример:  python tools/web_scan.py state/cc_usernames.txt --out state/cc_candidates.txt
"""
from __future__ import annotations

import argparse
import asyncio
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import tg_web_preview as wp  # noqa: E402


async def scan(names: list[str], days: int, pages: int, parallel: int, chunk: int = 1000):
    # t.me при перегрузке отвечает 429 — ждём сколько просит и продолжаем, а не бросаем проверку
    w = wp.WebPreview(parallel=parallel, max_limited=10 ** 9)
    since = datetime.now(timezone.utc) - timedelta(days=days)
    with_page = 0
    found: list[tuple[str, int]] = []
    t0 = time.time()
    try:
        for i in range(0, len(names), chunk):
            part = names[i:i + chunk]
            res = await asyncio.gather(*(w.candidates(u, since, max_pages=pages) for u in part))
            for u, r in zip(part, res):
                if r is not None:
                    with_page += 1
                if r:
                    found.append((u, len(r)))
            done = min(i + chunk, len(names))
            print(f"  {done}/{len(names)}: веб-страница у {with_page}, с кандидатами {len(found)}; "
                  f"страниц {w.pages}, 429: {w.limited}, {time.time() - t0:.0f} с", flush=True)
    finally:
        await w.close()
    return with_page, found, w


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("names", help="файл с @username, по одному в строке")
    ap.add_argument("--out", required=True, help="куда записать каналы с кандидатами")
    ap.add_argument("--days", type=int, default=31, help="насколько старые посты смотреть (31 — дольше розыгрыш не идёт)")
    ap.add_argument("--pages", type=int, default=3, help="страниц на канал, по ~20 постов (3)")
    ap.add_argument("--parallel", type=int, default=8, help="одновременных запросов к t.me (8)")
    ap.add_argument("--limit", type=int, default=0, help="проверить только первые N имён (0 — все)")
    args = ap.parse_args()

    names = [ln.strip().lstrip("@") for ln in Path(args.names).read_text(encoding="utf-8").splitlines()]
    names = [n for n in names if n and not n.startswith("#")]
    if args.limit:
        names = names[:args.limit]
    t0 = time.time()
    with_page, found, w = asyncio.run(scan(names, args.days, args.pages, args.parallel))
    found.sort(key=lambda x: -x[1])
    Path(args.out).write_text("".join(f"@{u}  # {n}\n" for u, n in found), encoding="utf-8")
    print(f"Итог: имён {len(names)}, веб-страница есть у {with_page}, с постами-кандидатами за {args.days} дн. — "
          f"{len(found)} (постов-кандидатов {sum(n for _, n in found)}); страниц {w.pages}, 429: {w.limited}, "
          f"ошибок {w.failed}, {(time.time() - t0) / 60:.0f} мин → {args.out}")


if __name__ == "__main__":
    main()
