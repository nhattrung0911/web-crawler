-- Universal product crawl sink — schema v7 (SQLite).
-- One table, every crawler dumps into it (source + external_id/url = identity).
-- JSON blobs hold flexible per-source data; flat columns expose the hot fields
-- for plain SQL (price, stock, weight, packaging dims, shipping cost).
--
-- Weight + packaging dimensions are REAL writable columns (not generated), so
-- they can be sorted/filtered independently AND back-filled later via UPDATE
-- (crawlers that don't return them leave NULL). Only volumetric/chargeable
-- weight are generated, computed from those writable columns.
--
-- Volumetric weight divisor = 5000 (cm3/kg, standard air-courier convention).
--
-- CANONICAL UNITS (crawlers must convert before ingest): weight = kg,
-- packaging dims = cm, price = numeric with explicit currency in commercial.price.
-- Identity rule: every row needs external_id or url (CHECK ck_identity).

CREATE TABLE IF NOT EXISTS product (
  -- ---- written directly by the ingest layer ----
  id             INTEGER PRIMARY KEY AUTOINCREMENT,
  source         TEXT NOT NULL,          -- e.g. "phuthaitech"
  external_id    TEXT,                   -- source product code / id (nullable)
  url            TEXT,                   -- canonical per-product URL (nullable; API sources may lack one)
  product_name   TEXT NOT NULL,          -- sole source of truth for name
  product        TEXT NOT NULL,          -- JSON: brand/sku/mpn/category/manufacturer...
  specification  TEXT,                   -- JSON: {raw_text, items:{attr:{raw,value,unit}}}
  commercial     TEXT,                   -- JSON: {price,stock,unit,min_order_qty}
  media          TEXT,                   -- JSON: {images,documents,videos}
  crawl          TEXT NOT NULL,          -- JSON: {status,version,crawler,score,account_tier...}

  -- ---- logistics: real writable columns (sort/filter/UPDATE freely) ----
  net_weight_kg    NUMERIC,              -- bare product weight
  gross_weight_kg  NUMERIC,              -- packed weight (what carriers weigh)
  pack_length_cm   NUMERIC,              -- packaging box L
  pack_width_cm    NUMERIC,              -- packaging box W
  pack_height_cm   NUMERIC,              -- packaging box H

  content_hash   TEXT NOT NULL,          -- sha256 over parsed business fields
  raw_hash       TEXT NOT NULL,          -- sha256 over raw payload
  schema_version TEXT NOT NULL,          -- "v7"
  created_at     TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
  updated_at     TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),

  -- ---- generated: identity ----
  dedup_key TEXT GENERATED ALWAYS AS
    (source || '|' || coalesce(external_id, url)) STORED,

  -- ---- generated: hot business fields (from JSON) ----
  sku            TEXT    GENERATED ALWAYS AS (json_extract(product,'$.sku'))              STORED,
  mpn            TEXT    GENERATED ALWAYS AS (json_extract(product,'$.mpn'))              STORED,
  brand          TEXT    GENERATED ALWAYS AS (json_extract(product,'$.brand'))            STORED,
  category_path  TEXT    GENERATED ALWAYS AS (json_extract(product,'$.category'))         STORED,
  price_numeric  NUMERIC GENERATED ALWAYS AS (json_extract(commercial,'$.price.value'))   STORED,
  availability   TEXT    GENERATED ALWAYS AS (json_extract(commercial,'$.stock.status'))  STORED,
  crawl_status   TEXT    GENERATED ALWAYS AS (json_extract(crawl,'$.status'))             STORED,
  parse_score    INTEGER GENERATED ALWAYS AS (json_extract(crawl,'$.score'))              STORED,

  -- ---- generated: shipping cost helpers (from the writable columns above) ----
  volumetric_weight_kg NUMERIC GENERATED ALWAYS AS (
    (pack_length_cm * pack_width_cm * pack_height_cm) / 5000.0
  ) STORED,
  -- chargeable = heavier of actual gross vs volumetric (max() ignores NULLs)
  chargeable_weight_kg NUMERIC GENERATED ALWAYS AS (
    max( gross_weight_kg,
         (pack_length_cm * pack_width_cm * pack_height_cm) / 5000.0 )
  ) STORED,

  -- identity rule: a row must be addressable by external_id or url (else dedup_key is NULL)
  CONSTRAINT ck_identity CHECK (external_id IS NOT NULL OR url IS NOT NULL)
);

CREATE UNIQUE INDEX IF NOT EXISTS ux_product_dedup     ON product(dedup_key);
CREATE INDEX IF NOT EXISTS ix_product_source           ON product(source);
CREATE INDEX IF NOT EXISTS ix_product_sku              ON product(sku);
CREATE INDEX IF NOT EXISTS ix_product_mpn              ON product(mpn);
CREATE INDEX IF NOT EXISTS ix_product_brand            ON product(brand);
CREATE INDEX IF NOT EXISTS ix_product_category         ON product(category_path);
CREATE INDEX IF NOT EXISTS ix_product_price            ON product(price_numeric);
CREATE INDEX IF NOT EXISTS ix_product_avail            ON product(availability);
CREATE INDEX IF NOT EXISTS ix_product_name             ON product(product_name);
CREATE INDEX IF NOT EXISTS ix_product_net_weight       ON product(net_weight_kg);
CREATE INDEX IF NOT EXISTS ix_product_gross_weight     ON product(gross_weight_kg);
CREATE INDEX IF NOT EXISTS ix_product_chargeable       ON product(chargeable_weight_kg);

-- Keep updated_at fresh on real changes.
CREATE TRIGGER IF NOT EXISTS trg_product_touch
AFTER UPDATE ON product FOR EACH ROW
BEGIN
  UPDATE product SET updated_at = strftime('%Y-%m-%dT%H:%M:%fZ','now')
  WHERE id = NEW.id;
END;

-- Price history — append-only. `product` holds the latest snapshot; this keeps
-- every distinct price observed over time (time-based discounts included).
-- Never overwritten: the ingest layer appends a row whenever the price differs
-- from the last observation for that product+source. Same across all sources;
-- account_tier disambiguates dealer-tier prices from public retail.
CREATE TABLE IF NOT EXISTS price_history (
  id           INTEGER PRIMARY KEY AUTOINCREMENT,
  dedup_key    TEXT NOT NULL,          -- links to product.dedup_key
  source       TEXT NOT NULL,
  external_id  TEXT,
  price_value  NUMERIC,
  currency     TEXT,
  stock_status TEXT,
  account_tier TEXT,                   -- which account/tier this price reflects
  observed_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
);
CREATE INDEX IF NOT EXISTS ix_ph_key    ON price_history(dedup_key, observed_at);
CREATE INDEX IF NOT EXISTS ix_ph_source ON price_history(source);
