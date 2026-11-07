"""Новые источники @username каналов (владелец 07.11: «надо расширить список каналов, не только русские»).

  wayback  — индекс Wayback Machine (archive.org CDX) по адресам t.me/…: сохранённые за годы страницы
             каналов со всего мира. Берём страницы индекса по очереди (--pages), из адресов — имена.
  wikidata — свойство P3789 «Telegram username»: официальные аккаунты компаний, СМИ, брендов всех стран.

Пишем только имена, которых ещё нет в --known. Аккаунт Telegram не нужен.
"""
import argparse
import json
import re
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

UA = {"User-Agent": "giveaway-finder-research/1.0 (public channel discovery)"}
NAME = re.compile(r"^(?:https?://)?(?:www\.)?(?:t|telegram)\.me/(?:s/)?([A-Za-z][A-Za-z0-9_]{3,31})(?:[/?#]|$)", re.I)
SKIP = {"joinchat", "addstickers", "addemoji", "share", "proxy", "socks", "iv", "login", "addlist", "boost",
        "giftcode", "invoice", "setlanguage", "confirmphone", "contact", "addtheme", "bg", "c"}


def get(url, tries=4, timeout=120):
    for i in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=timeout) as r:
                return r.read().decode("utf-8", "replace")
        except Exception as e:
            print(f"  повтор {url[:100]}: {e}", flush=True)
            time.sleep(10 * (i + 1))
    return ""


def wayback(pages: int, start: int):
    names = set()
    for host in ("t.me", "telegram.me"):
        n = get(f"https://web.archive.org/cdx/search/cdx?url={host}/&matchType=prefix&showNumPages=true").strip()
        total = int(n) if n.isdigit() else 0
        print(f"Wayback {host}: страниц индекса {total}", flush=True)
        for p in range(start, min(total, start + pages)):
            txt = get(f"https://web.archive.org/cdx/search/cdx?url={host}/&matchType=prefix&fl=original"
                      f"&collapse=urlkey&page={p}")
            before = len(names)
            for ln in txt.splitlines():
                m = NAME.match(ln.strip())
                if m and m.group(1).lower() not in SKIP:
                    names.add(m.group(1))
            print(f"  {host} страница {p}: строк {len(txt.splitlines())}, новых имён {len(names) - before}, всего {len(names)}",
                  flush=True)
            time.sleep(3)  # бережём archive.org
    return names


def wikidata():
    q = "SELECT ?u WHERE { ?i wdt:P3789 ?u }"
    txt = get("https://query.wikidata.org/sparql?format=json&query=" + urllib.parse.quote(q), timeout=300)
    names = set()
    try:
        for b in json.loads(txt)["results"]["bindings"]:
            u = b["u"]["value"].strip().lstrip("@")
            if re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{3,31}", u):
                names.add(u)
    except Exception as e:
        print("Wikidata: не разобрал ответ:", e, txt[:300])
    print(f"Wikidata P3789: имён {len(names)}", flush=True)
    return names


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("source", choices=["wayback", "wikidata", "both"])
    ap.add_argument("--known", default="state/cc_usernames.txt")
    ap.add_argument("--pages", type=int, default=5)
    ap.add_argument("--start", type=int, default=0)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    known = {ln.strip().lstrip("@").lower() for ln in Path(a.known).read_text(encoding="utf-8").splitlines() if ln.strip()}
    names = set()
    if a.source in ("wikidata", "both"):
        names |= wikidata()
    if a.source in ("wayback", "both"):
        names |= wayback(a.pages, a.start)
    new = sorted(n for n in names if n.lower() not in known)
    Path(a.out).write_text("".join(f"@{n}\n" for n in new), encoding="utf-8")
    print(f"Итог: имён {len(names)}, новых (нет в нашем списке из {len(known)}) — {len(new)} → {a.out}")
    print("Примеры новых:", " ".join(new[:: max(1, len(new) // 40)][:40]))


if __name__ == "__main__":
    main()
