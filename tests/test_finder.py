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

# Тесты публикации ниже проверяют расписание и очередь; доступность каналов-участников розыгрыша
# проверяется отдельно (test_web_preview.ClosedGiveawayTests), здесь считаем все каналы открытыми.
_real_closed_reason = gf.giveaway_closed_reason


async def _all_channels_open(*a, **kw):
    return None


gf.giveaway_closed_reason = _all_channels_open

NOW = datetime.now(timezone.utc)


def args(**kw):
    base = dict(depth=2, delay=(0, 0), page_delay=None, batch_size=500, batch_pause=0, days=31, per_channel_limit=400, max_channels=50,
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
        self.wait_times = []

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

    async def iter_messages(self, ent, limit=None, wait_time=None):
        cid = getattr(ent, "id", None) or ent.channel_id  # Channel или запомненный InputPeerChannel
        self.scanned.append(cid)
        self.wait_times.append(wait_time)
        for m in self.history.get(cid, [])[:limit]:
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
    def parse(self, *argv):
        return gf.apply_preset(gf.build_parser().parse_args(["scan", *argv]))

    def test_parser(self):
        a = self.parse("--dialogs", "-c", "@x", "--check", "--prize", "stars")
        self.assertTrue(a.dialogs and a.check)
        self.assertEqual(a.channel, ["@x"])
        self.assertEqual(a.prize, "stars")

    def test_normal_preset(self):
        a = self.parse()
        self.assertEqual((a.delay, a.page_delay, a.max_channels, a.batch_pause, a.max_flood_wait),
                         ((0, 0), None, 2000, 0, 120))
        src = gf.describe_sources(a)
        self.assertIn("без пауз", src)
        self.assertIn("поиск новых публичных каналов", src)
        self.assertNotIn("ваши подписки", src)

    def test_careful_preset_and_overrides(self):
        a = self.parse("--careful")
        self.assertEqual((a.delay, a.batch_size, a.batch_pause, a.max_channels, a.max_flood_wait),
                         ((0, 0), 500, 30, 2000, 300))
        self.assertIn("пауза 30 с каждые 500 каналов", gf.describe_sources(a))
        b = self.parse("--careful", "--batch-size", "200", "--batch-pause", "60", "--delay", "1-2")
        self.assertEqual((b.batch_size, b.batch_pause, b.delay), (200, 60, (1, 2)))

    def test_batch_pause_every_n_channels(self):
        chans = [channel(i, f"c{i}") for i in range(1, 8)]
        f = gf.GiveawayFinder(FakeClient(chans, {}), args(batch_size=3, batch_pause=30))
        for c in chans:
            f.enqueue(c, 0)
        sleeps = []
        real_sleep = asyncio.sleep

        async def fake_sleep(sec):
            sleeps.append(sec)
            await real_sleep(0)
        gf.asyncio.sleep = fake_sleep
        try:
            asyncio.run(f.crawl())
        finally:
            gf.asyncio.sleep = real_sleep
        self.assertEqual([s for s in sleeps if s], [30, 30])  # после 3-го и 6-го канала

    def test_parse_range(self):
        self.assertEqual(gf.parse_range("5"), (5, 5))
        self.assertEqual(gf.parse_range("2.5-4"), (2.5, 4))
        with self.assertRaises(argparse.ArgumentTypeError):
            gf.parse_range("9-3")
        with self.assertRaises(argparse.ArgumentTypeError):
            gf.parse_range("abc")
        for _ in range(50):
            self.assertTrue(2 <= gf.rand_pause((2, 4)) <= 4)
        self.assertEqual(gf.rand_pause(None), 0)

    def test_aggregators_seeded_by_default(self):
        f = gf.GiveawayFinder(None, self.parse())
        asyncio.run(f.add_seeds())
        self.assertEqual([item for item, _ in f.queue], gf.DEFAULT_AGGREGATORS)
        self.assertIn("каналы-агрегаторы", gf.describe_sources(self.parse()))
        f2 = gf.GiveawayFinder(None, self.parse("--no-aggregators"))
        asyncio.run(f2.add_seeds())
        self.assertEqual(len(f2.queue), 0)

    def test_page_delay_passed_to_history(self):
        client = FakeClient([channel(1, "a")], {})
        f = gf.GiveawayFinder(client, args(page_delay=(2, 4)))
        f.enqueue(channel(1, "a"), 0)
        asyncio.run(f.crawl())
        self.assertTrue(2 <= client.wait_times[0] <= 4)


if __name__ == "__main__":
    unittest.main()


class FakeAuthClient:
    def __init__(self, needs_password=False):
        self.authorized = False
        self.needs_password = needs_password
        self.calls = []

    async def is_user_authorized(self):
        return self.authorized

    async def send_code_request(self, phone):
        self.calls.append(("send_code", phone))
        return argparse.Namespace(phone_code_hash="HASH")

    async def sign_in(self, phone=None, code=None, *, password=None, phone_code_hash=None):
        self.calls.append(("sign_in", phone, code, password, phone_code_hash))
        if code and self.needs_password:
            raise gf.errors.SessionPasswordNeededError(request=None)
        self.authorized = True

    async def get_me(self):
        return argparse.Namespace(first_name="Test", username="tester")

    async def log_out(self):
        self.authorized = False
        self.calls.append(("log_out",))


class LoginTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.old = gf.os.environ.get("TG_SESSION")
        gf.os.environ["TG_SESSION"] = str(Path(self.tmp.name, "s"))

    def tearDown(self):
        if self.old is None:
            gf.os.environ.pop("TG_SESSION", None)
        else:
            gf.os.environ["TG_SESSION"] = self.old
        gf.os.environ.pop("TG_2FA_PASSWORD", None)
        self.tmp.cleanup()

    def run_login(self, client, *argv):
        a = gf.build_parser().parse_args(["login", *argv])
        out = io.StringIO()
        with redirect_stdout(out):
            asyncio.run(gf.cmd_login(client, a))
        return out.getvalue()

    def test_phone_code_password_flow(self):
        c = FakeAuthClient(needs_password=True)
        self.assertIn("код входа", self.run_login(c, "--phone", "+70000000000"))
        self.assertTrue(gf.login_state_path().is_file())
        self.assertIn("двухэтапная", self.run_login(c, "--code", " 12345 "))
        self.assertEqual(c.calls[1], ("sign_in", "+70000000000", "12345", None, "HASH"))
        gf.os.environ["TG_2FA_PASSWORD"] = "secret"
        self.assertIn("Вход выполнен", self.run_login(c, "--password"))
        self.assertTrue(c.authorized)
        self.assertFalse(gf.login_state_path().exists())
        self.assertIn("Уже выполнен вход", self.run_login(c))
        with redirect_stdout(io.StringIO()):
            asyncio.run(gf.cmd_logout(c, None))
        self.assertFalse(c.authorized)

    def test_code_without_phone_exits(self):
        with self.assertRaises(SystemExit):
            self.run_login(FakeAuthClient(), "--code", "1")


class PostTests(unittest.TestCase):
    def make(self, tmp, *extra):
        argv = ["post", "--to", "@mygroup", "--no-discover", "--no-aggregators", "-c", "@aggregator",
                "--countries", "RU,BY,KZ", "--state", str(Path(tmp, "posted.json")), "--delay", "0", *extra]
        return gf.apply_preset(gf.build_parser().parse_args(argv))

    def test_forwards_from_organizer_and_remembers(self):
        agg, host, iran, target = (channel(1, "aggregator"), channel(100, "host"),
                                   channel(300, "iran"), channel(900, "mygroup", left=False))
        g_all = giveaway_media([100], months=6, quantity=5)
        g_ir = giveaway_media([300], months=3, quantity=10, countries_iso2=["IR"])
        fwd = types.MessageFwdHeader(date=NOW, from_id=types.PeerChannel(100), channel_post=9)
        history = {1: [message(1, 50, g_all, fwd_from=fwd), message(1, 51, g_ir)],
                   100: [message(100, 9, g_all)], 300: [message(300, 7, g_ir)]}
        client = FakeClient([agg, host, iran, target], history)
        sent = []

        async def forward_messages(to, mid, from_peer=None):
            sent.append((to.id, mid, from_peer.channel_id))
        client.forward_messages = forward_messages

        with tempfile.TemporaryDirectory() as tmp:
            asyncio.run(gf.cmd_post(client, self.make(tmp, "--depth", "0")))
            # переслан пост из канала-организатора @host, а не копия агрегатора; иранский отфильтрован
            self.assertEqual(sent, [(900, 9, 100)])
            state = json.loads(Path(tmp, "posted.json").read_text(encoding="utf-8"))
            self.assertEqual(len(state["posted"]), 1)
            self.assertIsNotNone(state["last_post"])
            self.assertEqual(state["queue"], [])  # иранский не попал, единственный подходящий уже опубликован
            # повторный запуск не пересылает то же самое
            asyncio.run(gf.cmd_post(client, self.make(tmp, "--depth", "0")))
            self.assertEqual(len(sent), 1)

    def test_country_filter_prefers_all_countries(self):
        f = gf.GiveawayFinder(None, args())
        f.handle_message(message(1, 1, giveaway_media([1], months=3, days_left=1, countries_iso2=["RU"])), 1, 0)
        f.handle_message(message(2, 1, giveaway_media([2], months=3, days_left=5)), 2, 0)
        f.handle_message(message(3, 1, giveaway_media([3], months=3, days_left=2, countries_iso2=["US"])), 3, 0)
        gs = gf.filter_giveaways(f.giveaways.values(), args(countries="RU,BY,KZ"))
        self.assertEqual([g.countries for g in gs], [[], ["RU"]])


class PrizeThresholdTests(unittest.TestCase):
    def test_min_stars_and_min_premium(self):
        f = gf.GiveawayFinder(None, args())
        for cid, media in [(1, giveaway_media([1], stars=25000, quantity=100)),
                           (2, giveaway_media([2], stars=5000, quantity=50)),
                           (3, giveaway_media([3], months=6, quantity=5)),
                           (4, giveaway_media([4], months=3, quantity=1))]:
            f.handle_message(message(cid, 1, media), cid, 0)
        gs = gf.filter_giveaways(f.giveaways.values(), args(min_stars=10000, min_premium=5))
        self.assertEqual(sorted((g.kind, g.quantity) for g in gs), [("premium", 5), ("stars", 100)])


class IntervalTests(unittest.TestCase):
    def test_post_is_due(self):
        last = NOW - timedelta(hours=3)
        self.assertEqual(gf.post_is_due(last, 3.5, NOW)[0], False)
        self.assertEqual(gf.post_is_due(last, 3.5, NOW + timedelta(minutes=21))[0], True)  # допуск 10 мин
        self.assertEqual(gf.post_is_due(None, 3.5, NOW)[0], True)
        self.assertEqual(gf.post_is_due(last, 0, NOW)[0], True)

    def test_state_roundtrip_and_old_format(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp, "s.json")
            p.write_text(json.dumps({"k": "2030-01-01T00:00:00+00:00"}))  # старый формат
            self.assertEqual(gf.load_state(p), {"last_post": None, "queue": [],
                                                "posted": {"k": "2030-01-01T00:00:00+00:00"}})
            st = {"last_post": NOW.isoformat(), "posted": {"k": "x"}, "queue": [{"key": "q"}]}
            gf.save_state(p, st)
            self.assertEqual(gf.load_state(p), st)

    def test_post_skips_when_too_early(self):
        with tempfile.TemporaryDirectory() as tmp:
            state = Path(tmp, "posted.json")
            earlier = (NOW - timedelta(hours=1)).isoformat()
            gf.save_state(state, {"last_post": earlier, "posted": {}, "queue": []})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@x", "--state", str(state), "--min-interval-hours", "3.5"]))
            asyncio.run(gf.cmd_post(None, a))  # клиент не нужен: выходит до любых запросов
            self.assertEqual(gf.load_state(state)["last_post"], earlier)


class QueueTests(unittest.TestCase):
    def test_uses_queue_without_scanning(self):
        target = channel(900, "mygroup", left=False)

        class NoScanClient(FakeClient):
            async def iter_messages(self, ent, limit=None, wait_time=None):
                if ent.id != 900:  # читать можно только свой канал
                    raise AssertionError("сканирование не должно запускаться")
                for m in []:
                    yield m

        client = NoScanClient([target], {})
        sent = []

        async def forward_messages(to, mid, from_peer=None):
            sent.append((mid, from_peer.channel_id, from_peer.access_hash))
        client.forward_messages = forward_messages

        def entry(key, days, countries):
            return {"key": key, "until": (NOW + timedelta(days=days)).isoformat(), "prize": key,
                    "countries": countries, "channel_id": 100, "access_hash": 7, "msg_id": days,
                    "link": "https://t.me/x/1"}
        with tempfile.TemporaryDirectory() as tmp:
            state = Path(tmp, "s.json")
            gf.save_state(state, {"last_post": None, "posted": {"old": (NOW - timedelta(days=1)).isoformat()},
                                  "queue": [entry("ru", 1, ["RU"]), entry("all", 5, []),
                                            entry("ending", 0.05, [])]})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(state), "--countries", "RU,BY,KZ"]))
            asyncio.run(gf.cmd_post(client, a))
            st = gf.load_state(state)
        self.assertEqual(sent, [(1, 100, 7)])                     # раньше кончится — раньше публикуем
        self.assertEqual([e["key"] for e in st["queue"]], ["all"])  # истекающий (<3 ч) выброшен
        self.assertEqual(list(st["posted"]), ["ru"])               # завершившийся «old» забыт
        self.assertIsNotNone(st["last_post"])


class ManualPostTests(unittest.TestCase):
    """Если владелец сам запостил розыгрыш, бот ждёт 5 часов после этого поста."""

    def run_post(self, target_history, last_post_hours_ago, our_ids=()):
        target = channel(900, "mygroup", left=False)

        class TargetOnly(FakeClient):
            async def iter_messages(self, ent, limit=None, wait_time=None):
                assert ent.id == 900, "сканировать чужие каналы не нужно — есть очередь"
                for m in target_history:
                    yield m

        client = TargetOnly([target], {})
        sent = []

        async def forward_messages(to, mid, from_peer=None):
            sent.append(mid)
            return types.Message(id=777, peer_id=types.PeerChannel(900), date=NOW, message="")
        client.forward_messages = forward_messages
        last = (NOW - timedelta(hours=last_post_hours_ago)).isoformat()
        entry = {"key": "k", "until": (NOW + timedelta(days=3)).isoformat(), "prize": "p", "countries": [],
                 "channel_id": 100, "access_hash": 7, "msg_id": 5, "link": "https://t.me/x/5"}
        with tempfile.TemporaryDirectory() as tmp:
            state = Path(tmp, "s.json")
            gf.save_state(state, {"last_post": last, "posted": {}, "queue": [entry], "our_ids": list(our_ids)})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(state),
                 "--min-interval-hours", "3.5", "--manual-gap-hours", "5"]))
            asyncio.run(gf.cmd_post(client, a))
            return sent, gf.load_state(state)

    def msg(self, mid, hours_ago, text="", media=None, fwd=None):
        return types.Message(id=mid, peer_id=types.PeerChannel(900), date=NOW - timedelta(hours=hours_ago),
                             message=text, media=media, fwd_from=fwd)

    def test_recent_manual_giveaway_blocks(self):
        history = [self.msg(50, 2, "🎁 Розыгрыш Telegram Premium! Условия…")]
        sent, _ = self.run_post(history, last_post_hours_ago=4)
        self.assertEqual(sent, [])

    def test_old_manual_giveaway_allows(self):
        history = [self.msg(50, 6, "Розыгрыш звёзд")]
        sent, st = self.run_post(history, last_post_hours_ago=4)
        self.assertEqual(sent, [5])
        self.assertIn(777, st["our_ids"])  # свой новый пост запомнен
        self.assertEqual(len(st["post_times"]), 1)

    def test_own_posts_are_not_manual(self):
        fwd = types.MessageFwdHeader(date=NOW, from_id=types.PeerChannel(100), channel_post=1)
        history = [types.MessageService(id=61, peer_id=types.PeerChannel(900), date=NOW - timedelta(minutes=30),
                                        action=types.MessageActionPinMessage()),          # служебное — не пост
                   self.msg(41, 1.5, media=giveaway_media([100], months=6, quantity=5), fwd=fwd),  # наш по id
                   self.msg(40, 4, media=giveaway_media([100], months=6, quantity=5), fwd=fwd)]    # наш по времени
        sent, _ = self.run_post(history, last_post_hours_ago=4, our_ids=[41])
        self.assertEqual(sent, [5])

    def test_any_manual_post_blocks(self):
        history = [self.msg(60, 1, "всем привет, канал работает")]  # любой ваш пост, не только про розыгрыш
        sent, _ = self.run_post(history, last_post_hours_ago=4)
        self.assertEqual(sent, [])

    def test_manual_official_giveaway_detected(self):
        history = [self.msg(70, 1, media=giveaway_media([900], stars=10000, quantity=10))]
        sent, _ = self.run_post(history, last_post_hours_ago=4)
        self.assertEqual(sent, [])


    def test_skips_giveaway_already_in_channel(self):
        media = giveaway_media([100], months=6, quantity=5, days_left=3)
        fwd = types.MessageFwdHeader(date=NOW, from_id=types.PeerChannel(100), channel_post=5)
        history = [self.msg(80, 8, media=media, fwd=fwd)]  # вы уже переслали его 8 ч назад
        target = channel(900, "mygroup", left=False)

        class TargetOnly(FakeClient):
            async def iter_messages(self, ent, limit=None, wait_time=None):
                if ent.id == 900:
                    for m in history:
                        yield m

        client = TargetOnly([target], {})
        sent = []

        async def forward_messages(to, mid, from_peer=None):
            sent.append(mid)
        client.forward_messages = forward_messages
        dup = {"key": gf.giveaway_key(media), "until": media.until_date.isoformat(), "prize": "dup",
               "countries": [], "channel_id": 100, "access_hash": 7, "msg_id": 5, "link": "l"}
        other = dict(dup, key="other", msg_id=6, prize="other")
        with tempfile.TemporaryDirectory() as tmp:
            state = Path(tmp, "s.json")
            gf.save_state(state, {"last_post": (NOW - timedelta(hours=9)).isoformat(), "posted": {},
                                  "queue": [dup, other]})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(state), "--no-discover", "--no-aggregators",
                 "--min-interval-hours", "3.5", "--manual-gap-hours", "5"]))
            asyncio.run(gf.cmd_post(client, a))
        self.assertEqual(sent, [6])  # дубликат пропущен, опубликован следующий


class ScheduledPostTests(unittest.TestCase):
    """Отложенные посты: Telegram публикует их сам, запуски GitHub могут пропускаться."""

    def client(self, history=(), sched=()):
        target = channel(900, "mygroup", left=False)

        class C(FakeClient):
            def __init__(self):
                super().__init__([target], {})
                self.sched, self.forwards, self.deleted = list(sched), [], []

            async def iter_messages(self, ent, limit=None, wait_time=None):
                if ent.id == 900:
                    for m in history:
                        yield m

            async def get_messages(self, ent, scheduled=False, limit=None):
                return list(self.sched) if scheduled else []

            async def forward_messages(self, to, mid, from_peer=None, schedule=None):
                self.forwards.append((mid, schedule))
                return types.Message(id=1000 + len(self.forwards), peer_id=types.PeerChannel(900),
                                     date=schedule or NOW, message="")

            async def __call__(self, req):
                self.deleted.extend(req.id)
        return C()

    def media(self, n):
        return giveaway_media([100 + n], months=6, quantity=5, days_left=5)

    def entry(self, n):
        return {"key": gf.giveaway_key(self.media(n)), "until": (NOW + timedelta(days=5)).isoformat(),
                "prize": f"p{n}", "countries": [], "channel_id": 100 + n, "access_hash": 7,
                "msg_id": n, "link": f"https://t.me/x/{n}"}

    def run_post(self, client, st):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp, "s.json")
            gf.save_state(path, {"posted": {}, "queue": [], **st})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(path), "--no-discover", "--no-aggregators",
                 "--min-interval-hours", "3.5", "--manual-gap-hours", "5", "--schedule-ahead", "2"]))
            asyncio.run(gf.cmd_post(client, a))
            return gf.load_state(path)

    def near(self, dt, expected):
        self.assertLess(abs((dt - expected).total_seconds()), 90, f"{dt} != {expected}")

    def test_posts_now_and_keeps_two_scheduled(self):
        c = self.client()
        st = self.run_post(c, {"last_post": (NOW - timedelta(hours=4)).isoformat(),
                               "queue": [self.entry(1), self.entry(2), self.entry(3)]})
        self.assertEqual([m for m, _ in c.forwards], [1, 2, 3])
        self.assertIsNone(c.forwards[0][1])                       # пора — сразу
        self.near(c.forwards[1][1], NOW + timedelta(hours=3.5))   # дальше — отложенные с шагом 3,5 ч
        self.near(c.forwards[2][1], NOW + timedelta(hours=7))
        self.assertEqual(len(st["scheduled"]), 2)
        self.assertIn(self.entry(1)["key"], st["posted"])

    def test_nothing_to_do_when_two_already_scheduled(self):
        sched = [types.Message(id=i, peer_id=types.PeerChannel(900), date=NOW + timedelta(hours=h),
                               message="", media=self.media(i)) for i, h in ((1, 2), (2, 5.5))]
        st_sched = [{**self.entry(i), "at": (NOW + timedelta(hours=h)).isoformat()} for i, h in ((1, 2), (2, 5.5))]
        c = self.client(sched=sched)
        st = self.run_post(c, {"scheduled": st_sched, "queue": [self.entry(3)]})
        self.assertEqual(c.forwards, [])
        self.assertEqual(len(st["scheduled"]), 2)

    def test_manual_post_reschedules_bot_posts(self):
        manual = types.Message(id=50, peer_id=types.PeerChannel(900), date=NOW - timedelta(hours=1),
                               message="мой пост")
        sched = [types.Message(id=77, peer_id=types.PeerChannel(900), date=NOW + timedelta(hours=1),
                               message="", media=self.media(1))]
        st_sched = [{**self.entry(1), "at": (NOW + timedelta(hours=1)).isoformat()}]
        c = self.client(history=[manual], sched=sched)
        self.run_post(c, {"scheduled": st_sched, "last_post": (NOW - timedelta(hours=3)).isoformat(),
                          "queue": [self.entry(2)]})
        self.assertEqual(c.deleted, [77])                         # стоял в паузе после вашего поста
        self.assertEqual(c.forwards[0][0], 1)
        self.near(c.forwards[0][1], NOW + timedelta(hours=4))     # перенесён на 5 ч после вашего поста
        self.near(c.forwards[1][1], NOW + timedelta(hours=7.5))

    def test_published_scheduled_post_counts_as_posted(self):
        st_sched = [{**self.entry(1), "at": (NOW - timedelta(hours=1)).isoformat()}]  # уже вышел
        c = self.client()
        st = self.run_post(c, {"scheduled": st_sched, "post_times": [(NOW - timedelta(hours=1)).isoformat()],
                               "queue": [self.entry(1), self.entry(2), self.entry(3)]})
        self.assertIn(self.entry(1)["key"], st["posted"])
        self.assertEqual([m for m, _ in c.forwards], [2, 3])     # вышедший не повторяется
        self.near(c.forwards[0][1], NOW + timedelta(hours=2.5))   # через 3,5 ч после вышедшего


class EarlyStopTests(unittest.TestCase):
    def test_stops_after_first_suitable(self):
        chans = [channel(i, f"c{i}") for i in (1, 2, 3)]
        history = {1: [message(1, 1, giveaway_media([1], stars=500, quantity=5))],   # мелкий — не подходит
                   2: [message(2, 1, giveaway_media([2], stars=20000, quantity=50))],  # подходит
                   3: [message(3, 1, giveaway_media([3], months=6, quantity=10))]}
        client = FakeClient(chans, history)
        a = gf.apply_preset(gf.build_parser().parse_args(
            ["scan", "--no-discover", "--no-aggregators", "--depth", "0", "--min-stars", "10000",
             "--min-premium", "5", "-c", "@c1", "-c", "@c2", "-c", "@c3"]))

        def want(g):
            return bool(gf.filter_giveaways([g], a))
        f = asyncio.run(gf.run_finder(client, a, want=want, want_count=1))
        self.assertEqual(client.scanned, [1, 2])  # третий канал уже не читали
        self.assertTrue(f.found_enough)


class UnforwardableTests(unittest.TestCase):
    def test_skips_unforwardable_and_finds_next(self):
        target, c1, c2 = channel(900, "mygroup", left=False), channel(1, "c1"), channel(2, "c2")
        a_media = giveaway_media([1], stars=20000, quantity=10)   # из закрытого канала — переслать нельзя
        b_media = giveaway_media([2], stars=30000, quantity=10)
        history = {1: [message(1, 11, a_media)], 2: [message(2, 22, b_media)]}

        class C(FakeClient):
            async def get_messages(self, ent, scheduled=False, limit=None):
                return []

            async def forward_messages(self, to, mid, from_peer=None, schedule=None):
                if from_peer.channel_id == 1:
                    raise gf.errors.ChannelPrivateError(request=None)
                sent.append((mid, schedule))
                return types.Message(id=500, peer_id=types.PeerChannel(900), date=NOW, message="")

        sent = []
        client = C([target, c1, c2], history)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp, "s.json")
            gf.save_state(path, {"posted": {}, "queue": [], "last_post": (NOW - timedelta(hours=4)).isoformat()})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(path), "--no-discover", "--no-aggregators",
                 "--depth", "0", "-c", "@c1", "-c", "@c2", "--min-stars", "10000",
                 "--min-interval-hours", "3.5", "--manual-gap-hours", "5", "--schedule-ahead", "1"]))
            asyncio.run(gf.cmd_post(client, a))
            st = gf.load_state(path)
        self.assertEqual(sent, [(22, None)])                      # опубликован второй, сразу
        self.assertIn(gf.giveaway_key(a_media), st["bad"])        # первый запомнен как непригодный
        self.assertIn(gf.giveaway_key(b_media), st["posted"])


class PostWhenDueTests(ScheduledPostTests):
    """--schedule-ahead 0: раз в 30 минут просто обновляемся, раз в 3,5 часа ищем и публикуем ОДИН."""

    def run_post(self, client, st):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp, "s.json")
            gf.save_state(path, {"posted": {}, "queue": [], **st})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(path), "--no-discover", "--no-aggregators",
                 "--min-interval-hours", "3.5", "--manual-gap-hours", "5", "--schedule-ahead", "0"]))
            asyncio.run(gf.cmd_post(client, a))
            return gf.load_state(path)

    def test_not_due_just_refreshes(self):
        c = self.client()
        st = self.run_post(c, {"last_post": (NOW - timedelta(hours=1)).isoformat(),
                               "queue": [self.entry(1), self.entry(2)]})
        self.assertEqual(c.forwards, [])
        self.assertEqual(len(st["queue"]), 2)

    def test_due_posts_exactly_one(self):
        c = self.client()
        st = self.run_post(c, {"last_post": (NOW - timedelta(hours=4)).isoformat(),
                               "queue": [self.entry(1), self.entry(2)]})
        self.assertEqual(c.forwards, [(1, None)])                 # один пост, сразу, без отложенных
        self.assertEqual(st["scheduled"], [])
        self.assertIn(self.entry(1)["key"], st["posted"])

    # унаследованные тесты режима «N отложенных» здесь не нужны
    test_posts_now_and_keeps_two_scheduled = None
    test_nothing_to_do_when_two_already_scheduled = None
    test_manual_post_reschedules_bot_posts = None
    test_published_scheduled_post_counts_as_posted = None

    def run_at_local_hour(self, hour):
        """Пост пора публиковать; местное время (через --tz-offset) — hour:xx. Окно 7-22."""
        c = self.client()
        offset = (hour - NOW.hour) % 24
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp, "s.json")
            gf.save_state(path, {"posted": {}, "queue": [self.entry(1)],
                                 "last_post": (NOW - timedelta(hours=4)).isoformat()})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(path), "--no-discover", "--no-aggregators",
                 "--min-interval-hours", "3.5", "--schedule-ahead", "0",
                 "--post-hours", "7-22", "--tz-offset", str(offset)]))
            asyncio.run(gf.cmd_post(client=c, args=a))
        return c

    def test_no_posts_at_night(self):
        self.assertEqual(self.run_at_local_hour(3).forwards, [])        # 03:xx — ночь, ждём 07:00
        self.assertEqual(self.run_at_local_hour(22).forwards, [])       # 22:xx — уже нельзя
        self.assertEqual(self.run_at_local_hour(12).forwards, [(1, None)])  # днём — публикуем

    def test_post_window_shift(self):
        tz = timezone(timedelta(hours=3))
        night = datetime(2026, 11, 6, 23, 30, tzinfo=tz)
        early = datetime(2026, 11, 7, 5, 0, tzinfo=tz)
        day = datetime(2026, 11, 7, 12, 0, tzinfo=tz)
        morning = datetime(2026, 11, 7, 7, 0, tzinfo=tz)
        self.assertEqual(gf.in_post_window(night, (7, 22)), morning)
        self.assertEqual(gf.in_post_window(early, (7, 22)), morning)
        self.assertEqual(gf.in_post_window(day, (7, 22)), day)
        self.assertEqual(gf.in_post_window(night, None), night)

    def test_manual_post_blocks_when_due(self):
        manual = types.Message(id=50, peer_id=types.PeerChannel(900), date=NOW - timedelta(hours=2),
                               message="мой пост")
        c = self.client(history=[manual])
        self.run_post(c, {"last_post": (NOW - timedelta(hours=6)).isoformat(), "queue": [self.entry(1)]})
        self.assertEqual(c.forwards, [])                          # 5 часов после вашего поста ещё не прошло


class ReserveTests(unittest.TestCase):
    def test_search_builds_reserve_and_posts_soonest_ending(self):
        target = channel(900, "mygroup", left=False)
        chans = [target] + [channel(i, f"c{i}") for i in (1, 2, 3)]
        medias = {1: giveaway_media([1], stars=8000, quantity=10, days_left=6),
                  2: giveaway_media([2], months=3, quantity=5, days_left=1),     # кончится раньше всех
                  3: giveaway_media([3], stars=9000, quantity=20, days_left=3)}
        history = {i: [message(i, 10 + i, m)] for i, m in medias.items()}

        class C(FakeClient):
            async def get_messages(self, ent, scheduled=False, limit=None):
                return []

            async def forward_messages(self, to, mid, from_peer=None, schedule=None):
                sent.append(mid)
                return types.Message(id=600, peer_id=types.PeerChannel(900), date=NOW, message="")

        sent = []
        client = C(chans, history)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp, "s.json")
            gf.save_state(path, {"posted": {}, "queue": [], "last_post": (NOW - timedelta(hours=4)).isoformat()})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(path), "--no-discover", "--no-aggregators",
                 "--depth", "0", "-c", "@c1", "-c", "@c2", "-c", "@c3", "--min-stars", "5000",
                 "--min-premium", "3", "--min-interval-hours", "3.5", "--schedule-ahead", "0"]))
            asyncio.run(gf.cmd_post(client, a))
            st = gf.load_state(path)
        self.assertEqual(sent, [12])                                   # опубликован тот, что кончится раньше
        self.assertEqual([e["msg_id"] for e in st["queue"]], [13, 11])  # остальные — в запасе, по сроку


class OrganizerHistoryTests(unittest.TestCase):
    def test_finds_relaunched_giveaway_via_old_aggregator_post(self):
        agg, org = channel(1, "agg"), channel(50, "organizer")
        old = giveaway_media([50], months=3, quantity=5, days_left=-60)        # закончился 2 месяца назад
        fwd = types.MessageFwdHeader(date=NOW, from_id=types.PeerChannel(50), channel_post=3)
        new = giveaway_media([50], stars=7000, quantity=10, days_left=4)       # организатор запустил новый
        history = {1: [message(1, 900, old, fwd_from=fwd, age_days=70)],      # старше 31 дня
                   50: [message(50, 40, new, age_days=1)]}
        client = FakeClient([agg, org], history)
        a = gf.apply_preset(gf.build_parser().parse_args(
            ["scan", "--no-discover", "--no-aggregators", "-c", "@agg", "--source-limit", "500",
             "--min-stars", "5000", "--min-premium", "3"]))
        f = gf.GiveawayFinder(client, a)

        async def run():
            await f.add_seeds()
            await f.crawl()
        asyncio.run(run())
        self.assertEqual(client.scanned, [1, 50])                 # организатор старого розыгрыша проверен
        gs = gf.filter_giveaways(f.giveaways.values(), a)
        self.assertEqual([(g.stars, g.quantity) for g in gs], [(7000, 10)])

    def test_without_source_limit_old_posts_are_ignored(self):
        agg, org = channel(1, "agg"), channel(50, "organizer")
        old = giveaway_media([50], months=3, quantity=5, days_left=-60)
        history = {1: [message(1, 900, old, age_days=70)], 50: []}
        client = FakeClient([agg, org], history)
        a = gf.apply_preset(gf.build_parser().parse_args(["scan", "--no-discover", "--no-aggregators", "-c", "@agg"]))
        f = gf.GiveawayFinder(client, a)

        async def run():
            await f.add_seeds()
            await f.crawl()
        asyncio.run(run())
        self.assertEqual(client.scanned, [1])                     # старше 31 дня — не смотрим


class PublicOrganizerOnlyTests(unittest.TestCase):
    def test_private_organizer_skipped_even_if_aggregator_has_copy(self):
        target = channel(900, "mygroup", left=False)
        agg = channel(1, "agg")
        private_org = channel(60, None)                       # закрытый: без @username
        public_org = channel(70, "public_org")
        m_priv = giveaway_media([60], stars=10000, quantity=100, days_left=3)
        m_pub = giveaway_media([70], months=3, quantity=3, days_left=5)
        fwd_priv = types.MessageFwdHeader(date=NOW, from_id=types.PeerChannel(60), channel_post=8)
        fwd_pub = types.MessageFwdHeader(date=NOW, from_id=types.PeerChannel(70), channel_post=9)
        history = {1: [message(1, 101, m_priv, fwd_from=fwd_priv), message(1, 102, m_pub, fwd_from=fwd_pub)],
                   60: [message(60, 8, m_priv)], 70: [message(70, 9, m_pub)]}

        class C(FakeClient):
            async def get_messages(self, ent, scheduled=False, limit=None):
                return []

            async def forward_messages(self, to, mid, from_peer=None, schedule=None):
                sent.append((from_peer.channel_id, mid))
                return types.Message(id=700, peer_id=types.PeerChannel(900), date=NOW, message="")

        sent = []
        client = C([target, agg, private_org, public_org], history)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp, "s.json")
            gf.save_state(path, {"posted": {}, "queue": [], "last_post": (NOW - timedelta(hours=4)).isoformat()})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(path), "--no-discover", "--no-aggregators",
                 "-c", "@agg", "--min-stars", "5000", "--min-premium", "3",
                 "--min-interval-hours", "3.5", "--schedule-ahead", "0"]))
            asyncio.run(gf.cmd_post(client, a))
            st = gf.load_state(path)
        self.assertEqual(sent, [(70, 9)])                          # только из публичного организатора
        self.assertIn(gf.giveaway_key(m_priv), st["bad"])          # закрытый — пропущен и запомнен


class RefillTests(unittest.TestCase):
    def test_refill_fills_reserve_without_posting(self):
        target = channel(900, "mygroup", left=False)
        chans = [target] + [channel(i, f"c{i}") for i in (1, 2)]
        medias = {1: giveaway_media([1], stars=8000, quantity=10, days_left=6),
                  2: giveaway_media([2], months=3, quantity=5, days_left=2)}
        history = {i: [message(i, 10 + i, m)] for i, m in medias.items()}

        class C(FakeClient):
            async def get_messages(self, ent, scheduled=False, limit=None):
                return []

            async def forward_messages(self, *a, **kw):
                raise AssertionError("при --refill ничего не публикуем")

        client = C(chans, history)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp, "s.json")
            gf.save_state(path, {"posted": {}, "queue": [], "last_post": (NOW - timedelta(hours=9)).isoformat()})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(path), "--no-discover", "--no-aggregators",
                 "--depth", "0", "-c", "@c1", "-c", "@c2", "--min-stars", "5000", "--min-premium", "3",
                 "--min-interval-hours", "3.5", "--schedule-ahead", "0", "--refill"]))
            asyncio.run(gf.cmd_post(client, a))
            st = gf.load_state(path)
        self.assertEqual([e["msg_id"] for e in st["queue"]], [12, 11])  # запас по сроку окончания


class SearchCooldownTests(unittest.TestCase):
    def run_post(self, st_extra, extra_args=(), last_post_hours=4):
        target = channel(900, "mygroup", left=False)

        class C(FakeClient):
            async def get_messages(self, ent, scheduled=False, limit=None):
                return []

            async def iter_messages(self, ent, limit=None, wait_time=None):
                if ent.id != 900:
                    calls.append(ent.id)
                for m in []:
                    yield m

        calls = []
        client = C([target, channel(1, "c1")], {})
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp, "s.json")
            gf.save_state(path, {"posted": {}, "queue": [],
                                 "last_post": (NOW - timedelta(hours=last_post_hours)).isoformat(), **st_extra})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(path), "--no-discover", "--no-aggregators",
                 "--depth", "0", "-c", "@c1", "--min-interval-hours", "3.5", "--schedule-ahead", "0",
                 *extra_args]))
            asyncio.run(gf.cmd_post(client, a))
            return calls, gf.load_state(path)

    def test_empty_search_recently_skips_new_search(self):
        calls, _ = self.run_post({"last_search": (NOW - timedelta(hours=1)).isoformat(), "last_search_added": 0})
        self.assertEqual(calls, [])                       # перебор каналов не запускался

    def test_search_allowed_after_cooldown(self):
        calls, st = self.run_post({"last_search": (NOW - timedelta(hours=3)).isoformat(), "last_search_added": 0})
        self.assertEqual(calls, [1])                      # один поиск; пусто — без повторов
        self.assertEqual(st["last_search_added"], 0)

    def test_daily_search_even_when_not_due(self):
        # пост ещё рано, но поиска не было больше суток — проверяем агрегаторы и пополняем запас
        old = {"last_search": (NOW - timedelta(hours=25)).isoformat(), "last_search_added": 3}
        calls, st = self.run_post(old, ["--daily-search-hours", "24"], last_post_hours=1)
        self.assertEqual(calls, [1])
        self.assertNotEqual(st["last_search"], old["last_search"])
        calls, _ = self.run_post(old, [], last_post_hours=1)          # без опции — как раньше
        self.assertEqual(calls, [])
        fresh = {"last_search": (NOW - timedelta(hours=5)).isoformat(), "last_search_added": 3}
        calls, _ = self.run_post(fresh, ["--daily-search-hours", "24"], last_post_hours=1)
        self.assertEqual(calls, [])                                    # искали меньше суток назад

    def test_search_early_when_reserve_is_low(self):
        # запас пуст (меньше 4), поиск был 5 ч назад — ищем заранее, не дожидаясь суток
        st = {"last_search": (NOW - timedelta(hours=5)).isoformat(), "last_search_added": 3}
        calls, _ = self.run_post(st, ["--daily-search-hours", "24", "--refill-below", "4"], last_post_hours=1)
        self.assertEqual(calls, [1])
        recent = {"last_search": (NOW - timedelta(hours=1)).isoformat(), "last_search_added": 3}
        calls, _ = self.run_post(recent, ["--daily-search-hours", "24", "--refill-below", "4"], last_post_hours=1)
        self.assertEqual(calls, [])                                    # искали час назад — не чаще раза в 3 ч


class DueSlackTests(PostWhenDueTests):
    """Запуск чуть раньше срока (в пределах 10 минут) должен публиковать, а не пропускать."""
    test_not_due_just_refreshes = None
    test_due_posts_exactly_one = None
    test_manual_post_blocks_when_due = None

    def test_run_4_minutes_early_still_posts(self):
        c = self.client()
        last = NOW - timedelta(hours=3.5) + timedelta(minutes=4)   # до срока ещё 4 минуты
        st = self.run_post(c, {"last_post": last.isoformat(), "queue": [self.entry(1)]})
        self.assertEqual(c.forwards, [(1, None)])                  # опубликовано сразу, не отложено


class PostNowTests(PostWhenDueTests):
    test_not_due_just_refreshes = None
    test_due_posts_exactly_one = None
    test_manual_post_blocks_when_due = None

    def test_post_now_out_of_turn_keeps_schedule(self):
        c = self.client()
        last = (NOW - timedelta(hours=1)).isoformat()            # по расписанию ещё рано
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp, "s.json")
            gf.save_state(path, {"posted": {}, "last_post": last, "queue": [self.entry(1), self.entry(2)]})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(path), "--min-interval-hours", "3.5",
                 "--schedule-ahead", "0", "--post-now", "https://t.me/x/2"]))
            asyncio.run(gf.cmd_post(c, a))
            st = gf.load_state(path)
        self.assertEqual(c.forwards, [(2, None)])                # опубликован именно указанный
        self.assertEqual(st["last_post"], last)                  # расписание не сдвинулось
        self.assertEqual([e["msg_id"] for e in st["queue"]], [1])
        self.assertIn(1001, st["our_ids"])                       # свой пост запомнен — не «ручной»


class OtherCountriesFallbackTests(unittest.TestCase):
    """Владелец 07.11: «если ничего не найдёт, пусть использует розыгрыши и для других стран»."""

    def run_post(self, st_extra, fallback=True, history=None):
        target = channel(900, "mygroup", left=False)
        ua = giveaway_media([1], stars=1000, quantity=5, days_left=3, countries_iso2=["UA"])
        history = {1: [message(1, 11, ua)]} if history is None else history

        class C(FakeClient):
            async def get_messages(self, ent, scheduled=False, limit=None, ids=None):
                if ids:
                    return [m for m in history.get(ent.channel_id, []) if m.id in ids]
                return []

            async def forward_messages(self, to, mid, from_peer=None, schedule=None):
                sent.append((from_peer.channel_id, mid))
                return types.Message(id=800, peer_id=types.PeerChannel(900), date=NOW, message="")

        sent = []
        client = C([target, channel(1, "c1"), channel(2, "c2")], history)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp, "s.json")
            gf.save_state(path, {"posted": {}, "queue": [], "last_post": (NOW - timedelta(hours=4)).isoformat(),
                                 **st_extra})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(path), "--no-discover", "--no-aggregators",
                 "--depth", "0", "-c", "@c1", "--countries", "RU,BY,KZ", "--min-interval-hours", "3.5",
                 "--schedule-ahead", "0", *(["--fallback-other-countries"] if fallback else [])]))
            asyncio.run(gf.cmd_post(client, a))
            return sent, gf.load_state(path), ua

    def test_posts_other_country_when_nothing_else(self):
        sent, st, ua = self.run_post({})
        self.assertEqual(sent, [(1, 11)])
        self.assertIn(gf.giveaway_key(ua), st["posted"])
        self.assertEqual(st["backup"], [])

    def test_without_option_other_countries_are_skipped(self):
        sent, st, _ = self.run_post({}, fallback=False)
        self.assertEqual(sent, [])
        self.assertEqual(st["backup"], [])

    def test_own_countries_first(self):
        ru = giveaway_media([2], months=3, quantity=3, days_left=5, countries_iso2=["RU"])
        ua = giveaway_media([1], stars=1000, quantity=5, days_left=1, countries_iso2=["UA"])  # кончится раньше
        ru_e = {"key": gf.giveaway_key(ru), "until": (NOW + timedelta(days=5)).isoformat(), "prize": "ru",
                "countries": ["RU"], "channel_id": 2, "access_hash": 7, "msg_id": 22, "link": "https://t.me/c2/22"}
        ua_e = {"key": gf.giveaway_key(ua), "until": (NOW + timedelta(days=1)).isoformat(), "prize": "ua",
                "countries": ["UA"], "channel_id": 1, "access_hash": 7, "msg_id": 11, "link": "https://t.me/c1/11"}
        history = {1: [message(1, 11, ua)], 2: [message(2, 22, ru)]}
        sent, st, _ = self.run_post({"queue": [ru_e], "backup": [ua_e]}, history=history)
        self.assertEqual(sent, [(2, 22)])                       # свой — первым, хоть и кончится позже
        self.assertEqual([e["msg_id"] for e in st["backup"]], [11])


class OnePerChannelTests(PostWhenDueTests):
    """Владелец 07.11: от одного канала — один розыгрыш; из двух — где больше приза на канал-участник."""
    test_not_due_just_refreshes = None
    test_due_posts_exactly_one = None
    test_manual_post_blocks_when_due = None
    test_no_posts_at_night = None
    test_post_window_shift = None

    @staticmethod
    def e(cid, mid, chans, stars=500, qty=1, days=10):
        until = NOW + timedelta(days=days)
        key = f"{','.join(str(c) for c in sorted(chans))}|{int(until.timestamp())}|{qty}|0|{stars}"
        return {"key": key, "until": until.isoformat(), "prize": f"{stars}", "countries": [], "channel_id": cid,
                "access_hash": 7, "msg_id": mid, "link": f"https://t.me/c{cid}/{mid}"}

    def test_keeps_simplest_of_same_channel(self):
        five = self.e(1, 761, [1, 2, 3, 4, 5])
        eight = self.e(1, 768, [1, 2, 3, 4, 5, 6, 7, 8], days=20)
        other = self.e(9, 5, [9])
        bad = {}
        out = gf.one_per_channel([eight, five, other], set(), bad)
        self.assertEqual(sorted(x["msg_id"] for x in out), [5, 761])  # 500 ⭐ на 5 каналов лучше, чем на 8
        self.assertIn(eight["key"], bad)
        big = self.e(1, 770, [1, 2, 3, 4, 5, 6, 7, 8], stars=8000)     # 1000 ⭐ на канал — жирнее
        self.assertEqual([x["msg_id"] for x in gf.one_per_channel([five, big], set(), {})], [770])
        prem = {**self.e(1, 771, [1]), "key": self.e(1, 771, [1])["key"].replace("|1|0|500", "|3|3|0")}
        self.assertEqual(gf.prize_per_channel(prem), 3000)            # 3 × Premium 3 мес. ≈ 3000 ⭐

    def test_channel_already_posted_is_skipped(self):
        posted_media = giveaway_media([101, 555], stars=500, quantity=1, days_left=10)
        fwd = types.Message(id=50, peer_id=types.PeerChannel(900), date=NOW - timedelta(hours=6), message="",
                            media=posted_media,
                            fwd_from=types.MessageFwdHeader(date=NOW, from_id=types.PeerChannel(101), channel_post=7))
        c = self.client(history=[fwd])
        st = self.run_post(c, {"last_post": (NOW - timedelta(hours=6)).isoformat(), "our_ids": [50],
                               "queue": [self.entry(1), self.entry(2)]})
        self.assertEqual(c.forwards, [(2, None)])                 # от канала 101 у нас уже идёт розыгрыш
        self.assertIn(self.entry(1)["key"], st["bad"])
        self.assertEqual(st["posted_src"][self.entry(2)["key"]], 102)


class MaxChannelsTests(OnePerChannelTests):
    """Владелец 07.11: «максимум 4 канала — награда маленькая, а требования ппц»."""
    test_keeps_simplest_of_same_channel = None
    test_channel_already_posted_is_skipped = None

    def test_more_than_four_channels_dropped(self):
        five = self.e(1, 10, [1, 2, 3, 4, 5], days=2)        # кончится раньше, но каналов 5
        four = self.e(2, 20, [2, 6, 7, 8], days=5)
        c = self.client()
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp, "s.json")
            gf.save_state(path, {"posted": {}, "queue": [five, four],
                                 "last_post": (NOW - timedelta(hours=4)).isoformat()})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(path), "--no-discover", "--no-aggregators",
                 "--min-interval-hours", "3.5", "--schedule-ahead", "0", "--max-giveaway-channels", "4"]))
            asyncio.run(gf.cmd_post(c, a))
            st = gf.load_state(path)
        self.assertEqual(c.forwards, [(20, None)])
        self.assertEqual([e["msg_id"] for e in st["over_cap"]], [10])   # не в bad — в отдельный список
        self.assertNotIn(five["key"], st.get("bad", {}))


class OneSortedListTests(OnePerChannelTests):
    """Владелец 10.11: «добавляй в список все найденные, просто сортируй»."""
    test_keeps_simplest_of_same_channel = None
    test_channel_already_posted_is_skipped = None

    def test_fallback_order(self):
        big_own = self.e(1, 10, [1, 2, 3, 4, 5], days=6)                        # свои страны, 5 каналов
        ua = {**self.e(2, 20, [2], days=2), "countries": ["UA"]}                # другая страна, кончится раньше
        c = self.client()
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp, "s.json")
            gf.save_state(path, {"posted": {}, "queue": [], "backup": [ua], "over_cap": [big_own],
                                 "last_post": (NOW - timedelta(hours=4)).isoformat(),
                                 "last_search": (NOW - timedelta(minutes=30)).isoformat(), "last_search_added": 0})
            a = gf.apply_preset(gf.build_parser().parse_args(
                ["post", "--to", "@mygroup", "--state", str(path), "--no-discover", "--no-aggregators",
                 "--min-interval-hours", "3.5", "--schedule-ahead", "0", "--countries", "RU,BY,KZ",
                 "--max-giveaway-channels", "4", "--fallback-other-countries"]))
            asyncio.run(gf.cmd_post(c, a))
            st = gf.load_state(path)
        self.assertEqual(c.forwards, [(10, None)])        # свой с 5 каналами раньше чужой страны
        self.assertEqual([e["msg_id"] for e in st["backup"]], [20])
        self.assertEqual(st["over_cap"], [])


class BlocklistTests(unittest.TestCase):
    def test_load_blocklist(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp, "b.txt")
            p.write_text("# комментарий\n@CodeNameFan   # причина\n2454948751\n\n", encoding="utf-8")
            names, ids = gf.load_blocklist(str(p))
        self.assertEqual(names, {"codenamefan"})
        self.assertEqual(ids, {2454948751})
        self.assertEqual(gf.load_blocklist(None), (set(), set()))
