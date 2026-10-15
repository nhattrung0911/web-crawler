import asyncio
import json
import pathlib

import httpx

import trungtamthietbi as t

FX = pathlib.Path(__file__).parent / "fixtures"
U_SPEC = "https://www.trungtamthietbi.com/0-0-16mm-dong-ho-so-teclock-tm-1200-6336.html"
U_NOPRICE = "https://www.trungtamthietbi.com/0-0-375mm-panme-do-ghep-mi-lon-starret-208z-6777.html"


def rd(name):
    return (FX / name).read_text("utf-8")


# ---------------------------------------------------------------- parse_product
def test_parse_product_with_spec():
    (r,) = t.parse_product(rd("product_spec.html"), U_SPEC, "TCK-120-035")
    assert r["Item No"] == "TCK-120-035" and r["Input"] == "TCK-120-035" and r["Status"] == "OK"
    assert r["Page"] == "0 - 0.16mm Đồng hồ so Teclock TM-1200"
    assert r["Thuong hieu"] == "Teclock"
    assert r["Gia"] == "1,986,000 đ" and r["Gia VAT"].startswith("2,144,880")
    assert (r["Category 1"], r["Category 2"], r["Category 3"], r["Category 4"]) == (
        "Dụng cụ đo chính xác", "Dụng cụ so", "Đồng hồ so", "")
    assert r["Thong so"].splitlines() == [
        "Phạm vi đo: 0 - 0.16mm", "Độ chia: 0.001mm", "Độ chính xác: ±5µm",
        "Trọng lượng: 170g", "Nhà sản xuất: Teclock - Nhật Bản"]
    assert r["Image"].startswith("https://www.trungtamthietbi.com/products/")
    assert r["URL"] == U_SPEC
    assert all(isinstance(v, str) for v in r.values())


def test_parse_product_spec_cut_at_nbox():
    """BAY #5: carousel SP lien quan (n-box) khong duoc lot vao spec/mo ta."""
    html = rd("product_spec.html")
    assert 'class="n-box"' in html
    (r,) = t.parse_product(html, U_SPEC)
    blob = r["Thong so"] + r["Mo ta"]
    assert "VND" not in blob and "Giá:" not in blob


def test_parse_product_no_price():
    (r,) = t.parse_product(rd("product_noprice.html"), U_NOPRICE)
    assert r["Item No"] == "STR-208-053" and r["Gia"] == "Xin lien he" and r["Gia VAT"] == ""
    assert r["Thuong hieu"] == "Starret" and r["Category 3"] == "Panme đo ngoài"
    assert r["Thong so"].startswith('Phạm vị đo: 0-0.375"')


def test_parse_product_rejects_non_product():
    assert t.parse_product(rd("product_spec.html"), "https://www.trungtamthietbi.com/viewall.html") == []
    assert t.parse_product("", U_SPEC) == []


# ---------------------------------------------------------------- list / search
def test_list_items():
    items = t.Engine()._list_items(rd("list_page.html"))
    assert len(items) == 11
    assert items[0] == ("/1kg-binh-chua-chay-dang-bot-mfz1-5778.html", "ANV-000-237")
    assert all(t.is_product_path(p) and m for p, m in items)
    assert len({p for p, _ in items}) == 11


def test_search_items():
    items = t.Engine()._search_items(rd("search_teclock.html"))
    assert len(items) == 15
    assert items[0] == ("/0-0-8mm-dong-ho-so-teclock-lt-352-6340.html", "TCK-352-039")
    assert all(t.is_product_path(p) and m.startswith("TCK") for p, m in items)


# ---------------------------------------------------------------- paths
def test_clean_path_and_is_product_path():
    assert t.clean_path("../x-1.html") == "/x-1.html"
    assert t.clean_path("https://www.trungtamthietbi.com/x-1.html/junk/?a=1") == "/x-1.html"
    assert t.is_product_path("/x-1.html")
    assert not t.is_product_path("/lien-he,x-1.html")                 # form lien he (BAY #2)
    assert not t.is_product_path(t.clean_path("/lien-he,x-1.html"))
    assert not t.is_product_path("/danh-sach/x-1.html")               # co them thu muc
    assert not t.is_product_path("/viewall.html")
    assert t.product_url("/x-1.html/junk/") == t.BASE + "/x-1.html"   # duoi rac bi bo


def test_norm_code_keeps_url_case():
    assert t.norm_code(" tck-120 035 ") == "TCK-120035"
    assert t.norm_code("https://x/Ab-1.html") == "https://x/Ab-1.html"


# ---------------------------------------------------------------- checkpoint / migration
def test_checkpoint_under_out_dir(tmp_path, monkeypatch):
    monkeypatch.setattr(t, "HERE", tmp_path)
    monkeypatch.setattr(t, "OUT_DIR", tmp_path / "out_trungtamthietbi")
    monkeypatch.setattr(t, "LOG_FILE", tmp_path / "out_trungtamthietbi" / "x.log")
    e = t.Engine(log=lambda m: None)
    e.key = "all___"
    assert e.ckpt().parent == t.OUT_DIR and e.ckpt().name == "_raw_all__ee638824.jsonl"
    assert e.ckpt(t.HERE).parent == tmp_path                             # ten legacy cung ten
    e.log("[ERR] test")
    assert t.LOG_FILE.exists() and t.LOG_FILE.parent == t.OUT_DIR


def test_migration_moves_legacy_and_resumes(tmp_path, monkeypatch):
    monkeypatch.setattr(t, "HERE", tmp_path)
    monkeypatch.setattr(t, "OUT_DIR", tmp_path / "out_trungtamthietbi")
    msgs = []
    e = t.Engine(log=msgs.append)
    e.key = "all___"
    legacy = e.ckpt(tmp_path)
    recs = [{"Input": "", "Item No": f"A{i}", "Status": "OK", "URL": f"{t.BASE}/p{i}-{i}.html"}
            for i in range(3)]
    legacy.write_text("".join(json.dumps(r) + "\n" for r in recs) + '{"dang do', encoding="utf-8")
    raw = legacy.read_bytes()
    e.keep = False
    seen = e._load_ckpt()
    assert not legacy.exists()                                           # MOVE, khong copy
    assert e.ckpt().read_bytes() == raw                                  # byte-for-byte
    assert seen == {f"/p{i}-{i}.html" for i in range(3)} and e.rows == 3
    assert any("migrate checkpoint" in m for m in msgs)


def test_migration_never_overwrites_existing(tmp_path):
    legacy, new = tmp_path / "old.jsonl", tmp_path / "out" / "new.jsonl"
    legacy.write_text("L")
    new.parent.mkdir()
    new.write_text("N")
    assert t.migrate_legacy(legacy, new) is False
    assert legacy.read_text() == "L" and new.read_text() == "N"
    assert t.migrate_legacy(tmp_path / "none.jsonl", tmp_path / "out" / "z.jsonl") is False


def test_cleanup_keeps_checkpoint_when_limit_reached(tmp_path, monkeypatch):
    monkeypatch.setattr(t, "OUT_DIR", tmp_path)
    e = t.Engine()
    e.key, e.limit, e.rows = "all___", 5, 5
    e.ckpt().write_text("x")
    e.cleanup()
    assert e.ckpt().exists()                                             # crawl dang do -> giu
    e.limit = 0
    e.cleanup()
    assert not e.ckpt().exists()                                         # xong sach -> xoa


# ---------------------------------------------------------------- pool: 1 client / 1 request
def test_pool_never_lends_one_client_twice():
    busy: set[int] = set()
    seen: list[int] = []
    peak = 0

    async def handler(request):
        nonlocal peak
        cid = int(request.headers["x-cid"])
        assert cid not in busy, "client dang bi 2 request dung dong thoi"
        busy.add(cid)
        peak = max(peak, len(busy))
        await asyncio.sleep(0.02)
        busy.discard(cid)
        seen.append(cid)
        return httpx.Response(200, text="ok")

    async def main():
        pool = t.ClientPool(3)
        assert len({id(c.cookies.jar) for c in pool.clients}) == 3        # cookie jar rieng
        for i, c in enumerate(pool.clients):
            c._transport = httpx.MockTransport(handler)
            c.headers["x-cid"] = str(i)
        rs = await asyncio.gather(*(pool.get("https://x/a") for _ in range(30)))
        assert all(r.status_code == 200 for r in rs)
        assert pool.q.qsize() == 3                                       # tra du client
        await pool.aclose()

    asyncio.run(main())
    assert len(seen) == 30 and peak == 3 and set(seen) == {0, 1, 2}


def test_pool_returns_client_on_error():
    async def handler(request):
        raise httpx.ConnectError("boom")

    async def main():
        pool = t.ClientPool(1)
        pool.clients[0]._transport = httpx.MockTransport(handler)
        for _ in range(2):                                               # lan 2 khong bi treo
            try:
                await asyncio.wait_for(pool.get("https://x/a"), 2)
            except httpx.ConnectError:
                pass
        assert pool.q.qsize() == 1
        await pool.aclose()

    asyncio.run(main())


# ---------------------------------------------------------------- cross-sell trong Thong so
def test_strip_other_products_cuts_cross_sell():
    txt = 'Phạm vị đo: 0-0.375"\nĐộ chia: 0.001"\nCác sản phẩm khác:\n- Panme 208DZ\n- Đế từ 3657AA'
    assert t.strip_other_products(txt) == 'Phạm vị đo: 0-0.375"\nĐộ chia: 0.001"'
    for head in ("CAC SAN PHAM KHAC", "- Sản phẩm khác :", "Sản phẩm liên quan:", "san pham tuong tu",
                 "Các sản phẩm khác;", "Các sản phẩm khác: -", "Sản phẩm khác: Máy phay gỗ Bosch"):
        assert t.strip_other_products(f"A: 1\n{head}\nB: 2") == "A: 1"


def test_strip_other_products_keeps_real_spec():
    txt = "Màu sắc: khác\nGhi chú: xem sản phẩm khác để so sánh\nTrọng lượng: 1kg"
    assert t.strip_other_products(txt) == txt
    kb = "Sản phẩm khác biệt: có\nB: 1"
    assert t.strip_other_products(kb) == kb


def test_product_noprice_fixture_has_no_cross_sell():
    (r,) = t.parse_product(rd("product_noprice.html"), U_NOPRICE)
    assert r["Thong so"] == 'Phạm vị đo: 0-0.375"\nĐộ chia: 0.001"'
    (r,) = t.parse_product(rd("product_spec.html"), U_SPEC)
    assert r["Thong so"].count("\n") == 4


def test_iter_rows_cleans_old_checkpoint_without_rewriting(tmp_path, monkeypatch):
    monkeypatch.setattr(t, "OUT_DIR", tmp_path)
    e = t.Engine(log=lambda m: None)
    e.key, e.keep = "all___", False
    raw = json.dumps({"Thong so": "A: 1\nCác sản phẩm khác:\n- X"}) + "\n"
    e.ckpt().write_text(raw, encoding="utf-8")
    assert [r["Thong so"] for r in e.iter_rows()] == ["A: 1"]
    assert e.ckpt().read_text(encoding="utf-8") == raw


# ---------------------------------------------------------------- cat MAX_ITEMS phai co LOG
def test_max_items_truncation_is_logged_and_enforced(monkeypatch):
    monkeypatch.setattr(t, "MAX_ITEMS", 3)
    msgs, got = [], []
    e = t.Engine(conc=2, delay=0, log=msgs.append)

    async def fake_fetch(cli, url, tries=4):
        got.append(url)
        return None

    e.fetch = fake_fetch
    asyncio.run(e._crawl_paths(None, [f"/p{i}-{i}.html" for i in range(5)], set()))
    assert len(got) == 3 and e.total == 3
    assert any(m.startswith("[WARN] work-list 5 > MAX_ITEMS=3 -> cat con 3") for m in msgs)
    msgs.clear()
    asyncio.run(e._crawl_paths(None, ["/a-1.html", "/b-2.html"], set()))
    assert not any("MAX_ITEMS" in m for m in msgs)                       # khong cat -> khong WARN


# ---------------------------------------------------------------- discovery progress
def test_discovery_counters_advance_in_paged():
    e = t.Engine(delay=0, log=lambda m: None)
    calls = []

    async def fake_fetch(cli, url, tries=4):
        calls.append(url)
        return rd("search_teclock.html"), url       # trang 2 trung trang 1 -> stuck-page, dung

    e.fetch = fake_fetch
    seen_counts = []

    async def main():
        async for part in e._paged(None, t.BASE + "/search.html?keywords=teclock", "search"):
            seen_counts.append((e.pages_seen, e.found, len(part)))

    asyncio.run(main())
    assert seen_counts == [(1, 15, 15)]                 # cap nhat TRUOC khi yield, trang 2 trung -> dung
    assert len(calls) == 2


def test_stat_text_discovery_then_progress():
    e = t.Engine(log=lambda m: None)
    assert t.stat_text(e) == "0/0 | 0 dong"
    e.phase, e.pages_seen, e.found = "BFS danh sach", 7, 112
    assert t.stat_text(e) == "Dang tim: BFS danh sach | trang 7 | 112 SP tim thay"
    e.total, e.done, e.rows = 100, 40, 40
    assert t.stat_text(e) == "40/100 SP | 40 dong"
    e.phase = ""                                       # sau khi chay xong van hien done/total
    assert t.stat_text(e) == "40/100 SP | 40 dong"


# ---------------------------------------------------------------- cai moi: out dir CHUA TON TAI
def test_run_creates_missing_out_dir(tmp_path, monkeypatch):
    out = tmp_path / "chua_co" / "out_trungtamthietbi"
    monkeypatch.setattr(t, "OUT_DIR", out)
    monkeypatch.setattr(t, "LOG_FILE", out / "trungtamthietbi.log")
    assert not out.exists()

    class FakePool:
        def __init__(self, n):
            pass

        async def get(self, url):
            return httpx.Response(200, text=rd("product_spec.html"), request=httpx.Request("GET", url))

        async def aclose(self):
            pass

    monkeypatch.setattr(t, "ClientPool", FakePool)
    msgs = []
    e = t.Engine(conc=2, delay=0, log=msgs.append)
    asyncio.run(e.run("link", U_SPEC))
    assert e.rows >= 1 and e.ckpt().exists() and e.ckpt().parent == out
    assert not any("FATAL" in m for m in msgs)


# ---------------------------------------------------------------- Stop: khong bao "xong"
def test_run_all_stop_during_discovery_skips_crawl(monkeypatch):
    msgs, crawled = [], []
    e = t.Engine(conc=2, delay=0, log=msgs.append)

    async def fake_sitemap(cli):
        return {"/s-1.html"}, set()

    async def fake_fetch(cli, url, tries=4):
        if url.endswith("/viewall.html"):
            return '<a href="/danh-muc/a-1.html">', url
        if "danh-muc/" in url:
            return '<a href="/danh-sach/b-2.html">', url
        e.stop_evt.set()                              # Stop bam khi dang quet danh sach
        return rd("list_page.html"), url

    async def fake_crawl(*a, **k):
        crawled.append(1)

    e._sitemap, e.fetch, e._crawl_paths = fake_sitemap, fake_fetch, fake_crawl
    asyncio.run(e._run_all(None, set(), 0))
    assert not crawled
    assert not any("xong" in m.lower() or "Tong work-list" in m for m in msgs)
    assert any(m.startswith("[STOP] Da dung o pha BFS danh sach") for m in msgs)


def test_final_msg_stop_vs_done():
    e = t.Engine(log=lambda m: None)
    e.rows = 40
    assert t.final_msg(e).startswith("[SYS] XONG — 40 dong")
    e.stop_evt.set()
    m = t.final_msg(e)
    assert m.startswith("[STOP] DA DUNG — 40 dong (chua xong)") and "XONG" not in m
