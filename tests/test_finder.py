"""Офлайн-тесты логики поиска (без подключения к Telegram): python -m unittest discover tests"""
import argparse
import asyncio
import io
import json
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from telethon.tl import types  # noqa: E402

import tg_giveaway_finder as gf  # noqa: E402

NOW = datetime.now(timezone.utc)


def args(**kw):
    base = dict(depth=2, delay=0, days=31, per_channel_limit=400, max_channels=50,
                prize="any", min_winners=1, hide_ineligible=False)
    base.update(kw)
    return argparse.Namespace(**base)


def channel(cid, username=None, left=True, title=None):
    return types.Channel(id=cid, title=title or f"Chan{cid}", photo=types.ChatPhotoEmpty(),
                         date=NOW, access_hash=cid * 7, username=username, left=left,
                         broadcast=True)


def giveaway_media(channels, days_left=2.0, quantity=10, months=None, stars=None, **kw):
    return types.MessageMediaGiveaway(channels=channels, quantity=quantity,
                                      until_date=NOW + timedelta(days=days_left),
                                      months=months, stars=stars, **kw)


def message(cid, mid, media=None, fwd_from=None, age_days=1.0):
    return types.Message(id=mid, peer_id=types.PeerChannel(cid), date=NOW - timedelta(days=age_days),
                         message="", media=media, fwd_from=fwd_from)


class FakeClient:
    """Минимальная имитация TelegramClient для GiveawayFinder."""

    def __init__(self, chans, history):
        self.chans = {c.id: c for c in chans}
        self.history = history  # cid -> [Message], новые первыми
        self.scanned = []

    async def get_entity(self, x):
        cid = x.channel_id if isinstance(x, types.PeerChannel) else None
        if isinstance(x, str):
            for c in self.chans.values():
                if c.username == x.lstrip("@"):
                    return c
        if cid in self.chans:
            return self.chans[cid]
        raise ValueError("unknown")

    async def get_input_entity(self, x):
        return await self.get_entity(x)

    async def iter_messages(self, ent, limit=None):
        self.scanned.append(ent.id)
        for m in self.history.get(ent.id, [])[:limit]:
            yield m


class HandleMessageTests(unittest.TestCase):
    def setUp(self):
        self.f = gf.GiveawayFinder(client=None, args=args())

    def test_ignores_plain_and_finished(self):
        self.assertIsNone(self.f.handle_message(message(1, 1), 1, 0))
        ended = message(1, 2, giveaway_media([1], days_left=-1, months=3))
        self.assertIsNone(self.f.handle_message(ended, 1, 0))
        self.assertEqual(self.f.giveaways, {})

    def test_dedupes_copies_and_forwards(self):
        media = giveaway_media([100, 200], months=3, quantity=5)
        self.f.handle_message(message(100, 10, media), 100, 0)
        self.f.handle_message(message(200, 77, media), 200, 0)  # опубликован и во втором канале
        fwd = types.MessageFwdHeader(date=NOW, from_id=types.PeerChannel(100), channel_post=10)
        self.f.handle_message(message(999, 5, media, fwd_from=fwd), 999, 0)  # переслан агрегатором
        self.assertEqual(len(self.f.giveaways), 1)
        g = next(iter(self.f.giveaways.values()))
        self.assertEqual(g.origins, [(100, 10), (200, 77)])
        self.assertEqual(g.copies, [(999, 5)])
        self.assertEqual(g.kind, "premium")
        self.assertIn("Premium на 3 мес.", g.prize_text())

    def test_snowball_respects_depth(self):
        media = giveaway_media([100, 200], stars=5000, quantity=10)
        self.f.handle_message(message(100, 1, media), 100, depth=1)
        self.assertEqual({item for item, _ in self.f.queue}, {100, 200})
        f2 = gf.GiveawayFinder(client=None, args=args(depth=1))
        f2.handle_message(message(100, 1, media), 100, depth=1)
        self.assertEqual(len(f2.queue), 0)

    def test_stars_prize_text(self):
        g = self.f.handle_message(message(1, 1, giveaway_media([1], stars=5000, quantity=10,
                                                               prize_description="мерч")), 1, 0)
        self.assertEqual(g.kind, "stars")
        self.assertIn("5000 ⭐", g.prize_text())
        self.assertIn("~500 ⭐ каждому", g.prize_text())
        self.assertIn("+ мерч", g.prize_text())


class DescribeInfoTests(unittest.TestCase):
    def test_statuses(self):
        P = types.payments.GiveawayInfo
        self.assertEqual(gf.describe_info(P(start_date=NOW, participating=True))[1], True)
        self.assertEqual(gf.describe_info(P(start_date=NOW, disallowed_country="RU"))[1], False)
        self.assertEqual(gf.describe_info(P(start_date=NOW, joined_too_early_date=NOW))[1], False)
        self.assertEqual(gf.describe_info(P(start_date=NOW, admin_disallowed_chat_id=1))[1], False)
        self.assertIsNone(gf.describe_info(P(start_date=NOW, preparing_results=True))[1])
        self.assertTrue(gf.describe_info(P(start_date=NOW))[0].startswith("➕"))


class CrawlTests(unittest.TestCase):
    def test_crawl_follows_giveaway_channels(self):
        agg = channel(1, "aggregator", left=False)
        host = channel(100, "host")
        sponsor = channel(200, "sponsor", left=False)
        other = channel(300, "other")
        g1 = giveaway_media([100, 200], months=6, quantity=3, only_new_subscribers=True,
                            countries_iso2=["RU", "KZ"])
        g2 = giveaway_media([300], stars=1000, quantity=2, days_left=0.5)
        history = {
            1: [message(1, 50, g1, fwd_from=types.MessageFwdHeader(
                date=NOW, from_id=types.PeerChannel(100), channel_post=9))],
            100: [message(100, 9, g1)],
            200: [message(200, 4, giveaway_media([200], days_left=-2, months=1)),  # завершён
                  message(200, 3, g2, fwd_from=types.MessageFwdHeader(
                      date=NOW, from_id=types.PeerChannel(300), channel_post=8))],
            300: [message(300, 8, g2),
                  message(300, 1, giveaway_media([300], months=12), age_days=90)],  # старше --days
        }
        client = FakeClient([agg, host, sponsor, other], history)
        f = gf.GiveawayFinder(client, args())
        f.enqueue("@aggregator", 0)

        async def run():
            await f.crawl()
            await f.resolve_channels()
        asyncio.run(run())

        self.assertEqual(client.scanned, [1, 200, 300, 100])  # каналы с розыгрышами — в приоритете
        gs = gf.filter_giveaways(f.giveaways.values(), args())
        self.assertEqual(len(gs), 2)
        self.assertEqual(gs[0].kind, "stars")  # заканчивается раньше — первым
        self.assertEqual(gf.first_link(gs[1], f.refs), "https://t.me/host/9")
        self.assertEqual([c.member for c in gs[1].channels], [False, True])
        self.assertEqual(len(gf.filter_giveaways(f.giveaways.values(), args(prize="premium"))), 1)
        self.assertEqual(len(gf.filter_giveaways(f.giveaways.values(), args(min_winners=3))), 1)

        out = io.StringIO()
        with redirect_stdout(out):
            gf.print_report(gs, f.refs, len(f.scanned))
        text = out.getvalue()
        self.assertIn("Активных официальных розыгрышей: 2", text)
        self.assertIn("только новые подписчики", text)
        self.assertIn("страны: RU, KZ", text)

        with tempfile.TemporaryDirectory() as d:
            jp, cp, sp = Path(d, "o.json"), Path(d, "o.csv"), Path(d, "src.txt")
            gf.export(gs, f.refs, str(jp), str(cp))
            data = json.loads(jp.read_text(encoding="utf-8"))
            self.assertEqual(data[1]["channels"][0]["link"], "https://t.me/host")
            self.assertIn("https://t.me/host/9", cp.read_text(encoding="utf-8-sig"))
            gf.save_sources(gs, str(sp))
            gf.save_sources(gs, str(sp))  # повторно не дублирует
            self.assertEqual(sorted(sp.read_text().split()), ["@host", "@other", "@sponsor"])

    def test_max_channels_limit(self):
        chans = [channel(i, f"c{i}") for i in range(1, 6)]
        client = FakeClient(chans, {})
        f = gf.GiveawayFinder(client, args(max_channels=2))
        for c in chans:
            f.enqueue(c, 0)
        asyncio.run(f.crawl())
        self.assertEqual(len(client.scanned), 2)


class CliTests(unittest.TestCase):
    def test_parser(self):
        a = gf.build_parser().parse_args(["scan", "--dialogs", "-c", "@x", "--check", "--prize", "stars"])
        self.assertTrue(a.dialogs and a.check)
        self.assertEqual(a.channel, ["@x"])
        self.assertEqual(a.prize, "stars")


if __name__ == "__main__":
    unittest.main()
