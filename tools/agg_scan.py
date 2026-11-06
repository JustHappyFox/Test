#!/usr/bin/env python3
"""Поиск новых агрегаторов по веб-страницам (идея владельца 06.11: строка «Forwarded from (паблик)»).

Агрегатор пересылает чужие розыгрыши: на его странице t.me/s/<канал> много постов-заглушек со строкой
«Forwarded from <канал>» и ссылкой на оригинал. Для каждого имени читаем ТОЛЬКО первую страницу
(последние ~20 постов): считаем пересланные заглушки. Где их ≥ --min-forwards, проверяем у оригиналов
«Канал запустил розыгрыш» (до 3 ссылок на канал). Аккаунт Telegram не нужен.

Пример: python tools/agg_scan.py state/cc_usernames.txt --shard 0 --shards 5 --out /tmp/agg_0.txt
"""
from __future__ import annotations

import argparse
import asyncio
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import tg_web_preview as wp  # noqa: E402

KNOWN = {"tgagifts"}  # наш канал


async def scan(names: list[str], min_fwd: int, parallel: int):
    w = wp.WebPreview(parallel=parallel, max_limited=10 ** 9)
    found = []
    live = 0
    t0 = time.time()

    async def one(name: str):
        nonlocal live
        html = await w.fetch(w.BASE + name)
        posts = wp.parse_page(html) if html else None
        if posts is None:
            return
        live += 1
        fwd = [p for p in posts if p.stub and p.forwarded and p.source]
        if len(fwd) < min_fwd:
            return
        sources = list(dict.fromkeys(p.source for p in sorted(fwd, key=lambda p: -p.id)))
        confirmed = 0
        for src in sources[:3]:
            if await w._origin_launched(src):
                confirmed += 1
        newest = max((p.date for p in posts if p.date), default=None)
        found.append((name, len(fwd), len(posts), confirmed, newest, sources[:5]))

    try:
        chunk = 1000
        for i in range(0, len(names), chunk):
            await asyncio.gather(*(one(n) for n in names[i:i + chunk]))
            print(f"  {min(i + chunk, len(names))}/{len(names)}: живых {live}, с пересланными заглушками ≥{min_fwd}: "
                  f"{len(found)} (у оригиналов есть «запуск»: {sum(1 for f in found if f[3])}); страниц {w.pages}, "
                  f"429: {w.limited}, {time.time() - t0:.0f} с", flush=True)
    finally:
        await w.close()
    return live, found


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("names")
    ap.add_argument("--out", required=True)
    ap.add_argument("--shard", type=int, default=0)
    ap.add_argument("--shards", type=int, default=1)
    ap.add_argument("--min-forwards", type=int, default=3)
    ap.add_argument("--parallel", type=int, default=8)
    a = ap.parse_args()
    names = [ln.split("#", 1)[0].strip().lstrip("@") for ln in Path(a.names).read_text(encoding="utf-8").splitlines()]
    names = [n for n in names if n and n.lower() not in KNOWN]
    names = names[a.shard::a.shards]
    live, found = asyncio.run(scan(names, a.min_forwards, a.parallel))
    found.sort(key=lambda f: (-f[3], -f[1]))
    with open(a.out, "w", encoding="utf-8") as f:
        for name, nf, np_, conf, newest, src in found:
            f.write(f"@{name}  # {nf} {conf} {newest.strftime('%Y-%m-%d') if newest else '?'} {np_} {' '.join(src)}\n")
    print(f"Итог: имён {len(names)}, живых {live}, кандидатов в агрегаторы {len(found)}, "
          f"из них с подтверждённым «запуском» у оригиналов {sum(1 for f in found if f[3])} → {a.out}")
    for name, nf, np_, conf, newest, src in found[:60]:
        print(f"  @{name}: пересланных заглушек {nf} из {np_}, подтверждено {conf}, последний {newest}, откуда: {' '.join(src[:3])}")


if __name__ == "__main__":
    main()
