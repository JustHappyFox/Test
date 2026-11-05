"""Офлайн-тесты предпроверки каналов по веб-странице t.me/s/<канал>."""
import asyncio
import sys
import unittest
from datetime import timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from telethon.tl import functions, types  # noqa: E402

import tg_giveaway_finder as gf  # noqa: E402
import tg_web_preview as wp  # noqa: E402
from test_finder import NOW, FakeClient, args, channel, giveaway_media, message  # noqa: E402


def post_html(user, pid, stub=False, text=None, age_days=1.0):
    when = (NOW - timedelta(days=age_days)).strftime("%Y-%m-%dT%H:%M:%S+00:00")
    body = ('<div class="message_media_not_supported_wrap"><div class="message_media_not_supported">'
            '<div class="message_media_not_supported_label">Please open Telegram to view this post</div>'
            '</div></div>') if stub else ""
    if text:
        body += f'<div class="tgme_widget_message_text js-message_text" dir="auto">{text}</div>'
    return (f'<div class="tgme_widget_message_wrap js-widget_message_wrap">'
            f'<div class="tgme_widget_message text_not_supported_wrap js-widget_message" data-post="{user}/{pid}">'
            f'<div class="tgme_widget_message_bubble">{body}'
            f'<a class="tgme_widget_message_date" href="https://t.me/{user}/{pid}">'
            f'<time datetime="{when}" class="time">12:00</time></a></div></div></div>')


def page(user, posts):
    return "<html><body><section>" + "".join(post_html(user, *p) for p in posts) + "</section></body></html>"


class ParsePageTests(unittest.TestCase):
    def test_stub_posts_and_dates(self):
        html = page("otdel34", [(381, False, "привет"), (382, True), (383, False, "итоги скоро")])
        posts = wp.parse_page(html)
        self.assertEqual([p.id for p in posts], [381, 382, 383])
        self.assertEqual([p.stub for p in posts], [False, True, False])
        self.assertTrue(all(p.date is not None for p in posts))

    def test_protected_photos_and_big_videos_are_not_candidates(self):
        # у каналов с защитой контента заглушкой показаны и фото с подписью, и большие видео
        photo = post_html("okko", 10, stub=True, text="Новый фильм").replace(
            '<div class="message_media_not_supported_wrap">',
            '<a class="tgme_widget_message_photo_wrap"></a><div class="message_media_not_supported_wrap">')
        video = post_html("okko", 11, stub=True).replace(
            '<div class="message_media_not_supported_wrap">',
            '<div class="tgme_widget_message_video_player"></div><div class="message_media_not_supported_wrap">')
        html = "<section>" + photo + video + post_html("okko", 12, stub=True) + "</section>"
        self.assertEqual([p.stub for p in wp.parse_page(html)], [False, False, True])

    def test_blind_channel_detected(self):
        w = wp.WebPreview()
        blind = page("hidden", [(i, True, None, 1) for i in range(1, 11)])          # все 10 своих постов — заглушки
        fwd = "<section>" + "".join(post_html("agg", i, stub=True).replace(
            'class="tgme_widget_message_bubble">', 'class="tgme_widget_message_bubble"><div class="tgme_widget_message_forwarded_from"></div>')
            for i in range(1, 11)) + "</section>"                                     # агрегатор: заглушки пересланы
        normal = page("norm", [(i, i == 5, None if i == 5 else "текст", 1) for i in range(1, 11)])
        pages = {"https://t.me/s/hidden": blind, "https://t.me/s/agg": fwd, "https://t.me/s/norm": normal}

        async def fetch(url):
            return pages.get(url.split("?")[0])
        w.fetch = fetch
        since = NOW - timedelta(days=31)
        for name in ("hidden", "agg", "norm"):
            asyncio.run(w.candidates(name, since, max_pages=1))
        self.assertEqual(w.blind, {"hidden"})

    def test_no_web_feed(self):
        # канал без веб-ленты: t.me/s/… перенаправляет на обычную страницу канала
        self.assertIsNone(wp.parse_page('<html><div class="tgme_page_title">Chan</div></html>'))


class CandidatesTests(unittest.TestCase):
    def test_pages_back_until_cutoff(self):
        pages = {
            "https://t.me/s/chan": page("chan", [(58, True, None, 2), (59, False, "x", 1), (60, False, "y", 0.5)]),
            "https://t.me/s/chan?before=58": page("chan", [(55, True, None, 20), (57, False, "z", 10)]),
            "https://t.me/s/chan?before=55": page("chan", [(50, True, None, 40), (54, False, "old", 35)]),
        }
        w = wp.WebPreview()
        seen = []

        async def fetch(url):
            seen.append(url)
            return pages.get(url)
        w.fetch = fetch
        got = asyncio.run(w.candidates("@chan", NOW - timedelta(days=31)))
        self.assertEqual(got, [58, 55])  # пост 50 старше 31 дня
        self.assertEqual(len(seen), 3)    # дальше 31 дня не листаем

    def test_unavailable_page_means_use_api(self):
        w = wp.WebPreview()

        async def fetch(url):
            return None
        w.fetch = fetch
        self.assertIsNone(asyncio.run(w.candidates("chan", NOW - timedelta(days=31))))


# Разметка с t.me/s/rzavayuui (05.11): 42 — «Канал запустил розыгрыш», 43 — сам розыгрыш, 44 — «закрепил».
def real_post(pid, inner, cls="text_not_supported_wrap", time="2026-10-30T09:12:16+00:00"):
    return (f'<div class="tgme_widget_message_wrap js-widget_message_wrap"><div class="tgme_widget_message {cls} '
            f'js-widget_message" data-post="rzavayuui/{pid}" data-view="x"> <div class="tgme_widget_message_user">'
            '<a href="https://t.me/rzavayuui"><i class="tgme_widget_message_user_photo bgcolor6" data-content="р">'
            '<img src="https://cdn4.telesco.pe/file/a.jpg"></i></a></div> <div class="tgme_widget_message_bubble"> '
            '<i class="tgme_widget_message_bubble_tail"> </i> <div class="tgme_widget_message_author accent_color">'
            '<a class="tgme_widget_message_owner_name" href="https://t.me/rzavayuui"><span dir="auto">рижая</span>'
            f'</a></div> {inner}<div class="tgme_widget_message_footer compact js-message_footer"> '
            '<div class="tgme_widget_message_info short js-message_info"> <span class="tgme_widget_message_meta">'
            f'<a class="tgme_widget_message_date" href="https://t.me/rzavayuui/{pid}"><time datetime="{time}" '
            'class="time">09:12</time></a></span> </div> </div> </div> </div></div>')


LAUNCH = real_post(42, "", cls="text_not_supported_wrap service_message")
GIVEAWAY = real_post(43, '<div class="message_media_not_supported_wrap"><div class="message_media_not_supported">'
                         '<div class="message_media_not_supported_label">Please open Telegram to view this post</div>'
                         '</div></div>')
PINNED = real_post(44, '<div class="tgme_widget_message_text js-message_text" dir="auto"><a class="tgme_widget_message_author_name" '
                       'href="https://t.me/rzavayuui"><span dir="auto">рижая</span></a> pinned «<span '
                       'class="tgme_widget_service_strong_text" dir="auto"></span>»</div> ',
                   cls="text_not_supported_wrap service_message", time="2026-10-30T09:14:42+00:00")


class LaunchMessageTests(unittest.TestCase):
    def test_real_markup(self):
        posts = wp.parse_page("<section>" + LAUNCH + GIVEAWAY + PINNED + "</section>")
        self.assertEqual([(p.id, p.stub, p.empty_service) for p in posts],
                         [(42, False, True), (43, True, False), (44, False, False)])

    def _run(self, html, strict, pages=None, max_pages=20):
        w = wp.WebPreview()
        asked = []

        async def fetch(url):
            asked.append(url)
            if pages is not None:
                return pages.get(url)
            return html if "?" not in url else None
        w.fetch = fetch
        w.asked = asked
        from datetime import datetime, timezone
        return w, asyncio.run(w.candidates("rzavayuui", datetime(2026, 10, 1, tzinfo=timezone.utc),
                                           max_pages=max_pages, strict=strict))

    def test_strict_needs_launch_message(self):
        text = real_post(41, '<div class="tgme_widget_message_text js-message_text">Новость</div>')
        story = real_post(40, GIVEAWAY.split("</a></div> ", 1)[1].split('<div class="tgme_widget_message_footer')[0])
        story = story.replace("rzavayuui/43", "rzavayuui/40")
        html = "<section>" + text.replace("/41", "/39") + story + text + LAUNCH + GIVEAWAY + PINNED + "</section>"
        w, got = self._run(html, strict=False)
        self.assertEqual(got, [43, 40])
        self.assertEqual(w.launches, {"rzavayuui": 1})
        w, got = self._run(html, strict=True)
        self.assertEqual(got, [43])     # 40 — заглушка без «запуска» перед ней (репост истории и т.п.)

    def test_launch_must_be_at_the_same_time(self):
        late = LAUNCH.replace("2026-10-30T09:12:16", "2026-10-29T09:12:16")
        html = "<section>" + real_post(41, '<div class="tgme_widget_message_text">x</div>') + late + GIVEAWAY + "</section>"
        w, got = self._run(html, strict=True)
        self.assertEqual(got, [])

    def test_forwarded_stub_checked_at_the_original(self):
        # копия розыгрыша без «запуска» перед ней; оригинал (otdel34/382) — с ним
        def fwd(pid, source):
            return real_post(pid, GIVEAWAY.split("</a></div> ", 1)[1].split('<div class="tgme_widget_message_footer')[0]).replace(
                '<div class="message_media_not_supported_wrap">',
                '<div class="tgme_widget_message_forwarded_from accent_color">Forwarded from&nbsp;<a class="tgme_widget_message_forwarded_from_name" '
                f'href="https://t.me/{source}"><span dir="auto">OTDEL</span></a></div><div class="message_media_not_supported_wrap">')
        text = '<div class="tgme_widget_message_text js-message_text">x</div>'
        feed = "<section>" + real_post(40, text) + fwd(41, "otdel34/382") + fwd(42, "newsfeed/77") + real_post(43, text) + "</section>"
        origin = "<section>" + LAUNCH.replace("rzavayuui/42", "otdel34/381") + GIVEAWAY.replace("rzavayuui/43", "otdel34/382") + "</section>"
        story = "<section>" + real_post(76, text).replace("rzavayuui", "newsfeed") + GIVEAWAY.replace("rzavayuui/43", "newsfeed/77") + "</section>"
        pages = {"https://t.me/s/rzavayuui": feed, "https://t.me/s/otdel34?before=383": origin,
                 "https://t.me/s/newsfeed?before=78": story}
        w, got = self._run(None, strict=True, pages=pages)
        self.assertEqual(got, [41])      # 42 — переслан пост без «запуска» (не розыгрыш)
        self.assertEqual(w.launches, {"rzavayuui": 1})
        w, got = self._run(None, strict=False, pages=pages)
        self.assertEqual(got, [42, 41])  # без strict пересланные берутся все, оригиналы не открываются
        self.assertEqual(len(w.asked), 2)

    def test_deleted_launch_message_is_second_tier(self):
        # перед заглушкой номер пропущен — «Канал запустил розыгрыш» удалили; берём, но не считаем подтверждённым
        text = '<div class="tgme_widget_message_text js-message_text">x</div>'
        html = "<section>" + real_post(41, text) + GIVEAWAY + "</section>"   # 42 нет на странице
        w, got = self._run(html, strict=True)
        self.assertEqual(got, [43])
        self.assertEqual((w.launches, w.gaps), ({}, {"rzavayuui": 1}))

    def test_launch_on_the_next_page(self):
        # заглушка — самый старый пост на загруженных страницах: «запуск» ищем на следующей странице
        text = '<div class="tgme_widget_message_text js-message_text">x</div>'
        first = "<section>" + GIVEAWAY + PINNED + "</section>"
        pages = {"https://t.me/s/rzavayuui": first,
                 "https://t.me/s/rzavayuui?before=44": "<section>" + real_post(41, text) + LAUNCH + GIVEAWAY + "</section>"}
        w, got = self._run(None, strict=True, pages=pages, max_pages=1)
        self.assertEqual(got, [43])
        pages["https://t.me/s/rzavayuui?before=44"] = "<section>" + real_post(42, text) + GIVEAWAY + "</section>"
        w, got = self._run(None, strict=True, pages=pages, max_pages=1)
        self.assertEqual(got, [])


class FetchTests(unittest.TestCase):
    def test_retry_after_429_and_fallback(self):
        from aiohttp import web

        hits = {"busy": 0}

        async def handler(request):
            name = request.match_info["name"]
            if name == "busy":
                hits["busy"] += 1
                if hits["busy"] == 1:
                    return web.Response(status=429, headers={"Retry-After": "0"})
                return web.Response(text=page("busy", [(7, True)]), content_type="text/html")
            return web.Response(status=404)

        async def run():
            app = web.Application()
            app.router.add_get("/s/{name}", handler)
            runner = web.AppRunner(app)
            await runner.setup()
            site = web.TCPSite(runner, "127.0.0.1", 0)
            await site.start()
            port = site._server.sockets[0].getsockname()[1]
            w = wp.WebPreview(parallel=2)
            w.BASE = f"http://127.0.0.1:{port}/s/"
            try:
                busy = await w.candidates("busy", NOW - timedelta(days=31), max_pages=1)
                missing = await w.candidates("missing", NOW - timedelta(days=31))
            finally:
                await w.close()
                await runner.cleanup()
            return w, busy, missing

        w, busy, missing = asyncio.run(run())
        self.assertEqual(busy, [7])        # после 429 повторил запрос и получил страницу
        self.assertIsNone(missing)          # 404 — значит проверять через API
        self.assertEqual(w.limited, 1)
        self.assertFalse(w.broken)


class FakeWeb:
    parallel = 2
    broken = False
    pages = limited = 0

    def __init__(self, stubs):
        self.stubs = stubs  # username -> [id] или None (нет веб-страницы)
        self.asked = []

    async def candidates(self, username, since, max_pages=20, strict=False):
        self.asked.append(username)
        return self.stubs.get(username, [])

    async def close(self):
        pass


class WebClient(FakeClient):
    def __init__(self, chans, history):
        super().__init__(chans, history)
        self.opened = []      # (cid, ids) — что открыли через API по веб-проверке
        self.batches = []

    async def get_messages(self, ent, ids=None, **kw):
        cid = getattr(ent, "id", None) or ent.channel_id  # Channel или запомненный InputPeerChannel
        self.opened.append((cid, list(ids)))
        by_id = {m.id: m for m in self.history.get(cid, [])}
        return [by_id.get(i) for i in ids]

    async def __call__(self, req):
        assert isinstance(req, functions.channels.GetChannelsRequest)
        self.batches.append([c.channel_id for c in req.id])
        return types.messages.Chats(chats=[self.chans[c.channel_id] for c in req.id if c.channel_id in self.chans])

    async def get_input_entity(self, x):
        c = await self.get_entity(x)
        return types.InputPeerChannel(c.id, c.access_hash)


class WebCrawlTests(unittest.TestCase):
    def test_api_only_for_stub_posts(self):
        agg = channel(1, "aggregator")
        quiet = channel(10, "quiet")          # заглушек нет — API не трогаем
        host = channel(20, "host")            # есть заглушка — открываем только её
        sponsor = channel(30, "sponsor")      # пришёл из розыгрыша только как id
        noweb = channel(40, "noweb")          # веб-страницы нет — обычная проверка через API
        g = giveaway_media([20, 30], stars=500, quantity=5)
        history = {
            1: [message(1, 5, giveaway_media([20], days_left=-3, months=3))],  # завершён: проверить организатора
            20: [message(20, 7), message(20, 9, g)],
            30: [message(30, 3, g)],
            40: [message(40, 1)],
        }
        client = WebClient([agg, quiet, host, sponsor, noweb], history)
        f = gf.GiveawayFinder(client, args())
        f.web = FakeWeb({"quiet": [], "host": [9], "sponsor": [3], "noweb": None})
        f.enqueue("@aggregator", 0)
        f.enqueue(quiet, 1)
        f.enqueue(noweb, 1)

        asyncio.run(f.crawl())

        self.assertEqual(sorted(client.scanned), [1, 40])        # целиком через API: стартовый и без веб-страницы
        self.assertEqual(sorted(client.opened), [(20, [9]), (30, [3])])  # через API — только посты-заглушки
        self.assertIn([30], client.batches)                       # @username для id узнали одним пакетным запросом
        self.assertEqual(f.web_checked, 3)
        self.assertEqual(f.web_opened, 2)
        self.assertEqual(f.api_read, 2)
        self.assertEqual(len(f.giveaways), 1)
        self.assertNotIn("aggregator", f.web.asked)               # стартовые каналы — не по веб-странице

    def test_bad_channel_does_not_spoil_the_batch(self):
        from telethon import errors

        class BadBatchClient(WebClient):
            async def __call__(self, req):
                if any(c.channel_id == 24 for c in req.id):
                    raise errors.ChannelInvalidError(request=req)
                return await super().__call__(req)

        good = [channel(i, f"good{i}") for i in (21, 22, 23)]
        client = BadBatchClient([*good, channel(24, "bad24")], {})
        f = gf.GiveawayFinder(client, args())
        f.web = FakeWeb({})
        for cid in (21, 22, 23, 24):
            f.enqueue(cid, 1)
        asyncio.run(f.crawl())
        self.assertEqual(sorted(f.web.asked), ["good21", "good22", "good23"])  # хорошие — по веб-странице
        self.assertEqual(client.scanned, [24])                                  # битый — через API

    def test_big_list_is_web_only_with_open_limit(self):
        import tempfile
        chans = {f"chan{i}": channel(50 + i, f"chan{i}") for i in range(6)}
        g = giveaway_media([51], stars=1000, quantity=5)
        history = {51: [message(51, 7, g)], 53: [message(53, 2)]}
        client = WebClient(list(chans.values()), history)
        with tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False, encoding="utf-8") as f:
            f.write("\n".join(f"@chan{i}" for i in range(6)))
        # cc0: нет веб-страницы, cc1 и cc3: заглушки, cc2: без заглушек; открыть через API можно только 1
        f_ = gf.GiveawayFinder(client, args(discover_file=f.name, discover_offset=0, web_open_max=1))
        f_.web = FakeWeb({"chan0": None, "chan1": [7], "chan2": [], "chan3": [2], "chan4": [], "chan5": []})
        asyncio.run(self._seed_and_crawl(f_))
        self.assertEqual(client.scanned, [])                 # целиком через API — ни одного канала списка
        self.assertEqual(client.opened, [(51, [7])])         # открыт только первый канал с заглушками
        self.assertEqual(len(f_.giveaways), 1)
        self.assertEqual(f_.discover_consumed, 3)            # cc0, cc1, cc2; cc3 упёрся в лимит
        self.assertTrue(f_.discover_paused)

        # следующий поиск начинает с cc3
        client2 = WebClient(list(chans.values()), history)
        f2 = gf.GiveawayFinder(client2, args(discover_file=f.name, discover_offset=3, web_open_max=5))
        f2.web = FakeWeb({"chan3": [2], "chan4": [], "chan5": []})
        asyncio.run(self._seed_and_crawl(f2))
        self.assertEqual(f2.web.asked[0], "chan3")
        self.assertEqual(client2.opened, [(53, [2])])

        # каналы большого списка — в начале очереди, раньше стартовых; стартовый канал из списка не дублируется
        f3 = gf.GiveawayFinder(WebClient(list(chans.values()), history),
                               args(discover_file=f.name, discover_offset=0, channel=["@chan5", "@seedchan"],
                                    channels_file=[], no_aggregators=True))
        asyncio.run(f3.add_seeds())
        order = [(str(i), isinstance(i, gf.WebOnly)) for i, _ in f3.queue]
        self.assertEqual(order[:5], [(f"@chan{i}", True) for i in range(5)])
        self.assertEqual(order[5:], [("@chan5", False), ("@seedchan", False)])
        Path(f.name).unlink()

    @staticmethod
    async def _seed_and_crawl(f):
        f.args.no_aggregators = True
        f.args.channel = None
        f.args.channels_file = None
        await f.add_seeds()
        await f.crawl()

    def test_checked_non_giveaway_stubs_are_not_reopened(self):
        okko = channel(60, "okkomovies")
        host = channel(61, "hostchan")
        g = giveaway_media([61], stars=500, quantity=5)
        history = {60: [message(60, 5)], 61: [message(61, 9, g)]}  # у Okko заглушка — не розыгрыш
        client = WebClient([okko, host], history)
        f = gf.GiveawayFinder(client, args())
        f.web = FakeWeb({"okkomovies": [5], "hostchan": [9]})
        f.enqueue(okko, 1)
        f.enqueue(host, 1)
        asyncio.run(f.crawl())
        self.assertEqual(sorted(client.opened), [(60, [5]), (61, [9])])
        self.assertEqual(f.seen_stubs, {"okkomovies": {5}})  # розыгрыш в учёт «не открывать» не попал

        # следующий поиск: заглушку Okko уже проверяли — через API не открываем, розыгрыш открываем снова
        client2 = WebClient([okko, host], history)
        f2 = gf.GiveawayFinder(client2, args(seen_stubs={k: sorted(v) for k, v in f.seen_stubs.items()}))
        f2.web = FakeWeb({"okkomovies": [5], "hostchan": [9]})
        f2.enqueue(okko, 1)
        f2.enqueue(host, 1)
        asyncio.run(f2.crawl())
        self.assertEqual(client2.opened, [(61, [9])])

    def test_found_channels_are_remembered(self):
        import tempfile
        chan = channel(70, "promochan")
        g = giveaway_media([70], stars=500, quantity=5)
        history = {70: [message(70, 4, g)]}

        class CountingClient(WebClient):
            by_name = 0

            async def get_entity(self, x):
                if isinstance(x, str):
                    CountingClient.by_name += 1
                return await super().get_entity(x)

        with tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False, encoding="utf-8") as f:
            f.write("@promochan\n")
        client = CountingClient([chan], history)
        f1 = gf.GiveawayFinder(client, args(discover_file=f.name, web_open_max=5))
        f1.web = FakeWeb({"promochan": [4]})
        asyncio.run(self._seed_and_crawl(f1))
        self.assertEqual(CountingClient.by_name, 1)
        self.assertEqual(f1.peer_cache, {"promochan": [70, 70 * 7]})
        self.assertEqual(f1.discover_opened, 1)

        # следующий запуск: канал уже знаком — по имени не ищем, лимит не тратим
        f2 = gf.GiveawayFinder(client, args(discover_file=f.name, web_open_max=5, peer_cache=f1.peer_cache))
        f2.web = FakeWeb({"promochan": [4]})
        asyncio.run(self._seed_and_crawl(f2))
        self.assertEqual(CountingClient.by_name, 1)
        self.assertEqual(f2.discover_opened, 0)
        self.assertEqual(client.opened, [(70, [4]), (70, [4])])
        self.assertEqual(len(f2.giveaways), 1)
        Path(f.name).unlink()

    def test_seen_channel_not_checked_twice(self):
        host = channel(20, "host")
        client = WebClient([host], {})
        f = gf.GiveawayFinder(client, args())
        f.web = FakeWeb({"host": []})
        f.enqueue(host, 1)
        f.enqueue("@host", 1)
        asyncio.run(f.crawl())
        self.assertEqual(f.web.asked, ["host"])
        self.assertEqual(f.web_checked, 1)

    def test_username_of(self):
        f = gf.GiveawayFinder(client=None, args=args())
        self.assertEqual(f.username_of("https://t.me/otdel34/"), "otdel34")
        self.assertEqual(f.username_of("@Premiumlob"), "Premiumlob")
        self.assertIsNone(f.username_of("https://t.me/+AbCdEf123"))
        self.assertEqual(f.username_of(channel(5, "chan5")), "chan5")
        self.assertIsNone(f.username_of(5))


class NameLookupLimitTests(unittest.TestCase):
    """Telegram строго лимитирует поиск каналов по имени (ResolveUsername) — 05.11 Dolli получил FloodWait на 6 ч."""

    def test_target_from_dialogs_and_then_from_state(self):
        import argparse
        from telethon import errors
        target = channel(99, "TGaGifts", left=False)

        class Limited:
            async def get_entity(self, x):
                raise errors.FloodWaitError(request=None, capture=21767)

            async def iter_dialogs(self):
                class D:
                    entity = target
                yield D()

        st = {}
        self.assertIs(asyncio.run(gf.get_target(Limited(), argparse.Namespace(to="TGaGifts"), st)), target)
        self.assertEqual(st["target_peer"], {"username": "tgagifts", "id": 99, "hash": 99 * 7})
        again = asyncio.run(gf.get_target(None, argparse.Namespace(to="TGaGifts"), st))  # без запросов
        self.assertEqual((again.channel_id, again.access_hash), (99, 99 * 7))

    def test_stop_name_lookups_after_flood_wait(self):
        from telethon import errors
        known = channel(80, "knownchan")

        class Limited(WebClient):
            calls = 0

            async def get_entity(self, x):
                if isinstance(x, str):
                    Limited.calls += 1
                    raise errors.FloodWaitError(request=None, capture=21767)
                return await super().get_entity(x)

        client = Limited([known], {80: [message(80, 3)]})
        f = gf.GiveawayFinder(client, args(peer_cache={"knownchan": [80, 560]}))
        f.web = FakeWeb({"newchan1": [5], "newchan2": [6], "knownchan": [3]})
        for name in ("@newchan1", "@newchan2", "@knownchan"):
            f.enqueue(gf.WebOnly(name), 1)
        asyncio.run(f.crawl())
        self.assertEqual(Limited.calls, 1)              # после первого отказа по имени больше не ищем
        self.assertTrue(f.names_blocked)
        self.assertEqual(client.opened, [(80, [3])])    # знакомый канал проверили без поиска по имени


class ClosedGiveawayTests(unittest.TestCase):
    """05.11 вышел розыгрыш, где 3-й канал-участник закрытый — вступить в такой розыгрыш нельзя."""

    def _client(self, private_ids):
        from telethon import errors
        open_chans = {10: channel(10, "orgchan"), 11: channel(11, "sponsor1"),
                      12: channel(12, None, title="NoName")}

        class C:
            async def get_messages(self, peer, ids=None):
                media = giveaway_media(self.channels, months=3, quantity=1)
                return [message(peer.channel_id, ids[0], media)]

            async def get_entity(self, peer):
                if peer.channel_id in private_ids:
                    raise errors.ChannelPrivateError(request=None)
                return open_chans[peer.channel_id]
        return C()

    def test_reasons(self):
        from test_finder import _real_closed_reason as check
        c = self._client(private_ids={13})
        c.channels = [10, 11]
        self.assertIsNone(asyncio.run(check(c, 10, 70, 5)))
        c.channels = [10, 11, 13]
        self.assertIn("недоступен", asyncio.run(check(c, 10, 70, 5)))
        c.channels = [10, 12]
        self.assertIn("закрытый", asyncio.run(check(c, 10, 70, 5)))

    def test_closed_giveaway_is_not_published(self):
        import json
        import tempfile
        import test_finder as tf
        calls = []

        async def closed_for_channel_7(client, channel_id, access_hash, msg_id):
            return "канал-участник «Shum Gifts» недоступен" if channel_id == 7 else None
        gf.giveaway_closed_reason = closed_for_channel_7
        try:
            with tempfile.TemporaryDirectory() as d:
                state = Path(d, "st.json")
                until = (NOW + timedelta(days=5)).isoformat()
                q = [{"key": "a", "until": until, "prize": "A", "countries": [], "channel_id": 7, "access_hash": 1,
                      "msg_id": 3, "link": "https://t.me/frishniu/16"},
                     {"key": "b", "until": (NOW + timedelta(days=6)).isoformat(), "prize": "B", "countries": [],
                      "channel_id": 8, "access_hash": 1, "msg_id": 4, "link": "https://t.me/ok/4"}]
                state.write_text(json.dumps({"posted": {}, "queue": q, "last_post": None}), encoding="utf-8")

                class Client:
                    async def get_entity(self, x):
                        return channel(99, "TGaGifts", left=False)

                    async def iter_messages(self, *a, **kw):
                        if False:
                            yield None

                    async def get_messages(self, *a, **kw):
                        return []

                    async def forward_messages(self, to, mid, from_peer=None, schedule=None):
                        calls.append(from_peer.channel_id)
                        return types.Message(id=500, peer_id=types.PeerChannel(99), date=NOW, message="")

                a = gf.build_parser().parse_args(["post", "--to", "TGaGifts", "--state", str(state),
                                                  "--schedule-ahead", "0", "--min-interval-hours", "3.5"])
                asyncio.run(gf.cmd_post(Client(), gf.apply_preset(a)))
                st = json.loads(state.read_text(encoding="utf-8"))
        finally:
            gf.giveaway_closed_reason = tf._all_channels_open
        self.assertEqual(calls, [8])              # закрытый розыгрыш не переслан, взят следующий
        self.assertIn("a", st["bad"])             # и больше его не берём


if __name__ == "__main__":
    unittest.main()
