"""Видны ли розыгрыши Premium в блокчейне TON (идея 06.11): Fragment продаёт Premium, в т.ч. для
розыгрышей в каналах, за TON — платежи публичные, у них есть текстовый комментарий.
Смотрим комментарии к входящим платежам кошельков Fragment: что покупают и указан ли канал.

Аккаунт Telegram не нужен. Источники: tonapi.io (поиск адреса по имени), toncenter.com (транзакции).
"""
import argparse
import json
import re
import time
import urllib.parse
import urllib.request
from collections import Counter

UA = {"User-Agent": "giveaway-finder-research/1.0"}


def get(url: str, tries: int = 5):
    for i in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=30) as r:
                return json.load(r)
        except Exception as e:  # 429 / сеть — подождать
            print(f"  повтор {url[:90]}: {e}")
            time.sleep(2 * (i + 1))
    return None


def fragment_addresses() -> list[str]:
    found = []
    d = get("https://tonapi.io/v2/accounts/search?name=fragment") or {}
    for a in d.get("addresses", []):
        print(f"  tonapi: {a.get('name')!r} → {a.get('address')}")
        found.append(a.get("address"))
    return found


def comments(addr: str, pages: int):
    lt = h = None
    for _ in range(pages):
        q = {"address": addr, "limit": 100, "archival": "true"}
        if lt:
            q.update(lt=lt, hash=h)
        d = get("https://toncenter.com/api/v2/getTransactions?" + urllib.parse.urlencode(q))
        time.sleep(1.2)  # без ключа — 1 запрос в секунду
        res = (d or {}).get("result") or []
        if not res:
            return
        for t in res:
            m = t.get("in_msg") or {}
            yield t.get("utime"), m.get("source"), int(m.get("value") or 0) / 1e9, (m.get("message") or "").strip()
        last = res[-1]["transaction_id"]
        lt, h = last["lt"], last["hash"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--address", nargs="*", default=[])
    ap.add_argument("--pages", type=int, default=10)
    a = ap.parse_args()
    addrs = a.address or fragment_addresses()
    for addr in addrs[:4]:
        print(f"\n=== {addr}")
        pat = Counter()
        samples = {}
        prem = []
        n = 0
        for ut, src, ton, msg in comments(addr, a.pages):
            n += 1
            key = re.sub(r"\d+", "N", re.sub(r"Ref#\S+", "Ref#…", msg))[:80] or "(без комментария)"
            pat[key] += 1
            samples.setdefault(key, msg[:200])
            if re.search(r"premium|giveaway|розыгр|@|t\.me/", msg, re.I):
                prem.append((ut, round(ton, 2), msg[:200]))
        print(f"входящих сообщений: {n}")
        print("виды комментариев (частые):")
        for k, c in pat.most_common(25):
            print(f"  {c:5d} × {k!r}   пример: {samples[k]!r}")
        print(f"с Premium / giveaway / @ / t.me: {len(prem)}")
        for ut, ton, msg in prem[:40]:
            print(f"  {time.strftime('%d.%m %H:%M', time.gmtime(ut))}  {ton} TON  {msg!r}")


if __name__ == "__main__":
    main()
