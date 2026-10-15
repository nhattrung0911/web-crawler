"""
Universal product crawl sink — ingest layer for the `product` table (schema v7).

Every crawler builds a normalized record and calls `store(...)`. Identity is
(source, external_id) when external_id exists, else (source, url), collapsed into
one `dedup_key` so upsert is unambiguous. A row is only rewritten when its
content_hash actually changes, so `updated_at` reflects real data changes.

Usage:
    conn = connect("products.db")
    result = store(
        conn, source="phuthaitech", external_id="ABC123",
        url="https://.../product/ABC123", product_name="...",
        product={"brand": "...", "category": None},
        commercial={"price": {"value": 1000, "currency": "VND"}},
        gross_weight_kg=2.5, pack_length_cm=30, pack_width_cm=20, pack_height_cm=10,
        crawl={"status": "success", "score": 100, "crawler": "phuthai_crawl"},
        raw=raw_api_json,
    )
    # result in {"inserted", "updated", "unchanged"}
"""
from __future__ import annotations

import hashlib
import json
import sqlite3
from pathlib import Path

SCHEMA_VERSION = "v7"
SCHEMA_SQL = Path(__file__).with_name("schema.sql")


def connect(db_path: str | Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(db_path))
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))
    return conn


def _dumps(obj) -> str:
    # sort_keys so equal content always hashes identically
    return json.dumps(obj, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _sha(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def store(
    conn: sqlite3.Connection,
    *,
    source: str,
    product_name: str,
    product: dict,
    crawl: dict,
    url: str | None = None,
    external_id: str | None = None,
    specification: dict | None = None,
    commercial: dict | None = None,
    media: dict | None = None,
    net_weight_kg: float | None = None,
    gross_weight_kg: float | None = None,
    pack_length_cm: float | None = None,
    pack_width_cm: float | None = None,
    pack_height_cm: float | None = None,
    raw=None,
) -> str:
    """Insert or update one product. Returns inserted / updated / unchanged.

    Weight + packaging dims are real writable columns — pass them when the source
    has them, else leave None and back-fill later with a plain UPDATE.
    An UPDATE here does NOT overwrite existing weight/dims with None (COALESCE),
    so a manual back-fill survives re-crawls.

    Identity rule: pass external_id or url (or both) — a row must be addressable.
    Canonical units: weight in kg, packaging dims in cm, price with explicit
    currency. Convert at the crawler before calling store().
    """
    if external_id is None and url is None:
        raise ValueError("store() needs external_id or url (identity rule)")

    product_j = _dumps(product)
    spec_j = _dumps(specification) if specification is not None else None
    comm_j = _dumps(commercial) if commercial is not None else None
    media_j = _dumps(media) if media is not None else None
    crawl_j = _dumps(crawl)

    weights = (net_weight_kg, gross_weight_kg, pack_length_cm, pack_width_cm, pack_height_cm)

    # content_hash: business fields only (crawl metadata excluded so a re-crawl
    # with identical data but new duration/timestamp does not count as a change).
    content_hash = _sha(
        "|".join(
            [product_name, product_j, spec_j or "", comm_j or "", media_j or "",
             *(str(w) for w in weights)]
        )
    )
    raw_hash = _sha(_dumps(raw) if raw is not None else "")

    dedup_key = f"{source}|{external_id if external_id is not None else url}"
    existed = conn.execute(
        "SELECT 1 FROM product WHERE dedup_key=?", (dedup_key,)
    ).fetchone() is not None

    cur = conn.execute(
        """
        INSERT INTO product (
            source, external_id, url, product_name,
            product, specification, commercial, media, crawl,
            net_weight_kg, gross_weight_kg, pack_length_cm, pack_width_cm, pack_height_cm,
            content_hash, raw_hash, schema_version
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(dedup_key) DO UPDATE SET
            external_id    = excluded.external_id,
            url            = excluded.url,
            product_name   = excluded.product_name,
            product        = excluded.product,
            specification  = excluded.specification,
            commercial     = excluded.commercial,
            media          = excluded.media,
            crawl          = excluded.crawl,
            net_weight_kg   = coalesce(excluded.net_weight_kg,   product.net_weight_kg),
            gross_weight_kg = coalesce(excluded.gross_weight_kg, product.gross_weight_kg),
            pack_length_cm  = coalesce(excluded.pack_length_cm,  product.pack_length_cm),
            pack_width_cm   = coalesce(excluded.pack_width_cm,   product.pack_width_cm),
            pack_height_cm  = coalesce(excluded.pack_height_cm,  product.pack_height_cm),
            content_hash   = excluded.content_hash,
            raw_hash       = excluded.raw_hash,
            schema_version = excluded.schema_version
        WHERE product.content_hash <> excluded.content_hash
        """,
        (
            source, external_id, url, product_name,
            product_j, spec_j, comm_j, media_j, crawl_j,
            *weights,
            content_hash, raw_hash, SCHEMA_VERSION,
        ),
    )
    _log_price(conn, dedup_key, source, external_id, commercial, crawl)
    conn.commit()

    if not existed:
        return "inserted"
    return "updated" if cur.rowcount == 1 else "unchanged"


def _log_price(conn, dedup_key, source, external_id, commercial, crawl) -> None:
    """Append to price_history when the price differs from the last observation.

    Every distinct price over time is kept (time-based discounts included) — never
    overwritten. Identical consecutive prices are not re-logged (no new info).
    """
    price = (commercial or {}).get("price") or {}
    value = price.get("value")
    if value is None:
        return
    currency = price.get("currency")
    last = conn.execute(
        "SELECT price_value, currency FROM price_history "
        "WHERE dedup_key=? ORDER BY id DESC LIMIT 1",
        (dedup_key,),
    ).fetchone()
    if last is not None and last[0] == value and last[1] == currency:
        return
    conn.execute(
        "INSERT INTO price_history "
        "(dedup_key, source, external_id, price_value, currency, stock_status, account_tier) "
        "VALUES (?,?,?,?,?,?,?)",
        (
            dedup_key, source, external_id, value, currency,
            ((commercial or {}).get("stock") or {}).get("status"),
            (crawl or {}).get("account_tier"),
        ),
    )
