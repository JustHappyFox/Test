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
    found: list[tuple[str, int, datetime | None]] = []
    t0 = time.time()
    try:
        for i in range(0, len(names), chunk):
            part = names[i:i + chunk]
            res = await asyncio.gather(*(w.candidates(u, since, max_pages=pages) for u in part))
            for u, r in zip(part, res):
                if r is not None:
                    with_page += 1
                if r and u.lower() not in w.blind:  # «слепые» каналы (сайт прячет все посты) не берём
                    found.append((u, len(r), w.newest.get(u.lower())))
            done = min(i + chunk, len(names))
            print(f"  {done}/{len(names)}: веб-страница у {with_page}, «слепых» {len(w.blind)}, с кандидатами {len(found)}; "
                  f"страниц {w.pages}, 429: {w.limited}, {time.time() - t0:.0f} с", flush=True)
    finally:
        await w.close()
    return with_page, found, w


def read_candidates(path: Path) -> list[tuple[str, int, datetime | None]]:
    """Строки вида «@канал  # число дата» → (канал, число кандидатов, дата самого свежего)."""
    rows = []
    for ln in path.read_text(encoding="utf-8").splitlines():
        u, _, rest = ln.partition("#")
        u = u.strip().lstrip("@")
        if not u:
            continue
        parts = rest.split()
        n = int(parts[0]) if parts and parts[0].isdigit() else 1
        d = None
        if len(parts) > 1 and parts[1] != "?":
            try:
                d = datetime.fromisoformat(parts[1]).replace(tzinfo=timezone.utc)
            except ValueError:
                d = None
        rows.append((u, n, d))
    return rows


def write_candidates(path: Path, rows: list[tuple[str, int, datetime | None]]) -> None:
    # Порядок для поиска через API (поиск по имени у Telegram ограничен ~150–200 в сутки):
    # 1) свежие кандидаты (последний не старше 10 дней) — розыгрыш, скорее всего, ещё идёт;
    # 2) каналы с 1–3 кандидатами — обычный организатор проводит 1–2 розыгрыша в месяц,
    #    а десятки «заглушек» — чаще скрытые посты другого рода (репосты историй и т.п.).
    now = datetime.now(timezone.utc)

    def order(x):
        _, n, newest = x
        fresh = newest is not None and now - newest <= timedelta(days=10)
        return (not fresh, n > 3, n, -(newest.timestamp() if newest else 0))
    rows = sorted(rows, key=order)
    path.write_text("".join(f"@{u}  # {n} {d.strftime('%Y-%m-%d') if d else '?'}\n" for u, n, d in rows),
                    encoding="utf-8")


def combine(new: Path, base: Path, scanned_names: Path) -> int:
    """Влить результат проверки new в список base: строки base по проверенным сейчас каналам заменяются."""
    scanned = {ln.split("#", 1)[0].strip().lstrip("@").lower()
               for ln in scanned_names.read_text(encoding="utf-8").splitlines() if ln.strip()}
    rows = [r for r in (read_candidates(base) if base.is_file() else []) if r[0].lower() not in scanned]
    rows += read_candidates(new)
    write_candidates(base, rows)
    return len(rows)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("names", help="файл с @username, по одному в строке")
    ap.add_argument("--out", help="куда записать каналы с кандидатами")
    ap.add_argument("--days", type=int, default=31, help="насколько старые посты смотреть (31 — дольше розыгрыш не идёт)")
    ap.add_argument("--pages", type=int, default=3, help="страниц на канал, по ~20 постов (3)")
    ap.add_argument("--parallel", type=int, default=8, help="одновременных запросов к t.me (8)")
    ap.add_argument("--limit", type=int, default=0, help="проверить только первые N имён (0 — все)")
    ap.add_argument("--merge", action="store_true",
                    help="добавить найденное к уже имеющемуся списку --out, а не заменять его")
    ap.add_argument("--combine", nargs=2, metavar=("NEW", "BASE"),
                    help="не проверять, а влить готовый результат NEW в список BASE (names — что проверяли)")
    args = ap.parse_args()

    if args.combine:
        total = combine(Path(args.combine[0]), Path(args.combine[1]), Path(args.names))
        print(f"Список кандидатов {args.combine[1]}: {total} каналов")
        return
    if not args.out:
        ap.error("нужен --out")
    names = [ln.split("#", 1)[0].strip().lstrip("@") for ln in Path(args.names).read_text(encoding="utf-8").splitlines()]
    names = [n for n in names if n]
    if args.limit:
        names = names[:args.limit]
    t0 = time.time()
    with_page, found, w = asyncio.run(scan(names, args.days, args.pages, args.parallel))
    if args.merge and Path(args.out).is_file():  # прежние кандидаты, которых в этой проверке не было
        scanned = {u.lower() for u in names}
        found += [r for r in read_candidates(Path(args.out)) if r[0].lower() not in scanned]
    write_candidates(Path(args.out), found)
    print(f"Итог: имён {len(names)}, веб-страница есть у {with_page}, «слепых» (сайт прячет все посты) {len(w.blind)}, "
          f"с постами-кандидатами за {args.days} дн. — "
          f"{len(found)} (постов-кандидатов {sum(x[1] for x in found)}); страниц {w.pages}, 429: {w.limited}, "
          f"ошибок {w.failed}, {(time.time() - t0) / 60:.0f} мин → {args.out}")


if __name__ == "__main__":
    main()
