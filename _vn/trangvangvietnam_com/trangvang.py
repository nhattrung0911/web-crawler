#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
==============================================================================
TRANG VANG VIETNAM CRAWLER v1.0
Crawl danh ba doanh nghiep tu trangvangvietnam.com
==============================================================================

Pipeline:
  P0: Thu thap danh muc (findex.asp / gindex/)
  P1: Crawl danh sach cong ty (listing pages / search)
  P2: Crawl chi tiet cong ty (detail pages)
  P3: Download assets (logos / product images)

Cai dat:
  pip install httpx selectolax sqlalchemy pydantic rich openpyxl

Vi du:
  python trangvang_crawler.py --pipe 0
  python trangvang_crawler.py --pipe 1 --category 26360
  python trangvang_crawler.py --pipe 1 --search "vong bi" --where "Ha Noi"
  python trangvang_crawler.py --pipe 1 --group 41
  python trangvang_crawler.py --pipe 2 --limit 50
  python trangvang_crawler.py --pipe 3
  python trangvang_crawler.py --export csv
  python trangvang_crawler.py --stats
==============================================================================
"""

from __future__ import annotations

import argparse
import asyncio
import csv
import json
import logging
import os
import random
import re
import sys
import time
from datetime import datetime
from html import unescape
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
from urllib.parse import quote, urljoin, urlparse

import aiohttp
from pydantic import BaseModel, Field
from selectolax.parser import HTMLParser
from sqlalchemy import (
    Boolean, Column, DateTime, Float, Integer, String, Text,
    create_engine, func, inspect,
)
from sqlalchemy.orm import Session, declarative_base, sessionmaker

# =====================================================
# LOGGING
# =====================================================

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")

logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("trangvang")


# =====================================================
# PROGRESS BAR (lightweight, no Rich dependency)
# =====================================================


class ProgressBar:
    """Single-line progress bar using \\r overwrite - throttled to 4 updates/sec"""

    def __init__(self, total: int, prefix: str = ""):
        self.total = max(total, 1)
        self.done = 0
        self.ok = 0
        self.fail = 0
        self.prefix = prefix
        self._start = time.monotonic()
        self._last_render = 0.0  # throttle: max 4 redraws/sec

    def update(self, success: bool = True, count: int = 1):
        self.done += count
        if success:
            self.ok += count
        else:
            self.fail += count
        now = time.monotonic()
        # Only redraw every 250ms OR on last item
        if now - self._last_render >= 0.25 or self.done >= self.total:
            self._last_render = now
            self._print()

    def set_total(self, total: int):
        """Update total dynamically (e.g. when discovering more pages)"""
        self.total = max(total, 1)

    def _print(self):
        elapsed = time.monotonic() - self._start
        pct = min(self.done * 100 // self.total, 100)
        bar_len = 30
        filled = min(bar_len * self.done // self.total, bar_len)
        bar = "█" * filled + "░" * (bar_len - filled)

        if self.done > 0 and self.done < self.total:
            eta = elapsed * (self.total - self.done) / self.done
            eta_str = f"ETA {eta:.0f}s"
        elif self.done >= self.total:
            eta_str = f"Done {elapsed:.1f}s"
        else:
            eta_str = "..."

        speed = self.done / elapsed if elapsed > 0 else 0
        line = f"\r  {self.prefix} {bar} {self.done}/{self.total} ({pct}%) OK:{self.ok} FAIL:{self.fail} | {speed:.1f}/s {eta_str}  "
        sys.stdout.write(line)
        sys.stdout.flush()

    def finish(self):
        self.done = min(self.done, self.total)
        self._print()
        sys.stdout.write("\n")
        sys.stdout.flush()

# =====================================================
# CONSTANTS
# =====================================================

BASE_URL = "https://trangvangvietnam.com"
FINDEX_URL = f"{BASE_URL}/findex.asp"
GINDEX_URL = f"{BASE_URL}/gindex"
SEARCH_URL = f"{BASE_URL}/search.asp"
AUTOCOMPLETE_URL = f"{BASE_URL}/auto/autocomplete3.asp"
PROVINCES_URL = f"{BASE_URL}/auto/provinces.asp"

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
]

GROUP_IDS = {
    1: "nong-nghiep", 7: "hoa-chat", 8: "may-mac", 9: "bao-bi",
    10: "co-khi", 11: "quang-cao", 13: "van-tai-giao-nhan",
    14: "du-lich", 17: "do-gia-dung", 18: "y-te", 21: "dien-lanh",
    24: "go-va-do-go", 27: "an-ninh-bao-ve", 30: "o-to-xe-may",
    31: "giay", 32: "in-an", 34: "nhua", 35: "doanh-nghiep-can-dung",
    36: "moi-truong", 38: "may-moc", 39: "dien-va-thiet-bi-dien",
    41: "cong-nghiep-va-phu-tro", 42: "noi-that", 46: "bao-ho-lao-dong",
    47: "qua-tang", 48: "cao-su",
}

LETTERS = list("ABCDEFGHIKLMNOPQRSTUVWXYZ")  # no J

# =====================================================
# CONFIG
# =====================================================


class CrawlerConfig(BaseModel):
    concurrency: int = Field(default=15, ge=1, le=50)
    timeout: int = Field(default=30, ge=5)
    max_retries: int = Field(default=3, ge=1)
    output_dir: str = "output"
    db_name: str = "trangvang.db"


# =====================================================
# DATABASE MODELS
# =====================================================

Base = declarative_base()


class Category(Base):
    __tablename__ = "categories"

    id = Column(Integer, primary_key=True, autoincrement=False)
    name = Column(String(500))
    slug = Column(String(500))
    url = Column(String(1000), unique=True)
    total_companies = Column(Integer, default=0)
    parent_id = Column(Integer, nullable=True)
    group_id = Column(Integer, nullable=True)
    is_sub = Column(Boolean, default=False)
    crawled = Column(Boolean, default=False, index=True)
    created_at = Column(DateTime, default=datetime.utcnow)


class Company(Base):
    __tablename__ = "companies"

    id = Column(Integer, primary_key=True, autoincrement=False)
    name = Column(String(500))
    slug = Column(String(500))
    url = Column(String(1000), unique=True, index=True)
    category_name = Column(String(500))      # Nganh nghe (parent)
    sub_category_name = Column(String(500))  # Nhom nganh (child)
    category_id = Column(Integer, nullable=True)
    address = Column(Text)
    province = Column(String(200))
    phones = Column(Text)          # JSON array
    hotline = Column(String(100))
    zalo = Column(String(100))
    email = Column(String(300))
    website = Column(String(500))
    logo_url = Column(String(500))
    is_sponsor = Column(Boolean, default=False)
    sponsor_level = Column(String(100))
    is_verified = Column(Boolean, default=False)
    is_outdated = Column(Boolean, default=False)
    ad_text = Column(Text)
    last_updated = Column(String(100))
    source_page = Column(String(50))  # "category", "search", "province"
    detail_crawled = Column(Boolean, default=False, index=True)
    created_at = Column(DateTime, default=datetime.utcnow)


class CompanyDetail(Base):
    __tablename__ = "company_details"

    company_id = Column(Integer, primary_key=True, autoincrement=False)
    full_name = Column(String(500))
    business_type = Column(String(300))
    industry = Column(Text)             # Nganh nghe kinh doanh
    main_market = Column(String(300))
    main_customers = Column(Text)
    tax_code = Column(String(50))
    year_established = Column(String(20))
    employee_count = Column(String(100))
    fax = Column(String(100))
    all_phones = Column(Text)       # all phones found on detail page, ; separated
    all_websites = Column(Text)     # all websites found, ; separated
    description = Column(Text)
    contact_name = Column(String(200))
    contact_phone = Column(String(100))
    contact_email = Column(String(300))
    product_groups = Column(Text)   # JSON array
    products = Column(Text)         # JSON array
    product_count = Column(Integer, default=0)
    image_count = Column(Integer, default=0)
    images = Column(Text)           # JSON array of URLs
    crawled_at = Column(DateTime, default=datetime.utcnow)


# =====================================================
# DATABASE MANAGER
# =====================================================


class DatabaseManager:

    def __init__(self, db_path: str):
        self.db_path = db_path
        Path(db_path).parent.mkdir(parents=True, exist_ok=True)
        self.engine = create_engine(
            f"sqlite:///{db_path}",
            connect_args={"check_same_thread": False},
            echo=False,
        )
        Base.metadata.create_all(self.engine)
        self._Session = sessionmaker(bind=self.engine)
        self._migrate()

    def _migrate(self):
        """Add new columns to existing tables if missing"""
        from sqlalchemy import text
        insp = inspect(self.engine)
        if insp.has_table("companies"):
            existing = {c["name"] for c in insp.get_columns("companies")}
            with self.engine.connect() as conn:
                if "is_outdated" not in existing:
                    conn.execute(text(
                        "ALTER TABLE companies ADD COLUMN is_outdated BOOLEAN DEFAULT 0"
                    ))
                if "sub_category_name" not in existing:
                    conn.execute(text(
                        "ALTER TABLE companies ADD COLUMN sub_category_name VARCHAR(500)"
                    ))
                conn.commit()

    def session(self) -> Session:
        return self._Session()

    # ---------- Category ----------

    def upsert_category(self, data: dict):
        with self.session() as s:
            existing = s.get(Category, data["id"])
            if existing:
                for k, v in data.items():
                    if k != "id" and v is not None:
                        setattr(existing, k, v)
            else:
                s.add(Category(**data))
            s.commit()

    def upsert_categories(self, items: list[dict]):
        with self.session() as s:
            for data in items:
                existing = s.get(Category, data["id"])
                if existing:
                    for k, v in data.items():
                        if k != "id" and v is not None:
                            setattr(existing, k, v)
                else:
                    s.add(Category(**data))
            s.commit()

    def get_categories(self, uncrawled_only: bool = False,
                       group_id: int | None = None,
                       category_ids: list[int] | None = None,
                       letter: str | None = None) -> list[Category]:
        with self.session() as s:
            q = s.query(Category)
            if uncrawled_only:
                q = q.filter(Category.crawled == False)
            if group_id is not None:
                q = q.filter(Category.group_id == group_id)
            if category_ids:
                q = q.filter(Category.id.in_(category_ids))
            if letter:
                q = q.filter(Category.name.ilike(f"{letter}%"))
            cats = q.order_by(Category.id).all()
            s.expunge_all()
            return cats

    def mark_category_crawled(self, cat_id: int):
        with self.session() as s:
            cat = s.get(Category, cat_id)
            if cat:
                cat.crawled = True
                s.commit()

    # ---------- Company ----------

    def upsert_company(self, data: dict) -> bool:
        with self.session() as s:
            existing = s.get(Company, data["id"])
            if existing:
                for k, v in data.items():
                    if k != "id" and v is not None:
                        setattr(existing, k, v)
                s.commit()
                return False  # updated
            s.add(Company(**data))
            s.commit()
            return True  # inserted

    def upsert_companies(self, items: list[dict]) -> list[int]:
        """Upsert companies, return list of ALL IDs (both new and updated)."""
        all_ids = []
        with self.session() as s:
            for data in items:
                all_ids.append(data["id"])
                existing = s.get(Company, data["id"])
                if existing:
                    for k, v in data.items():
                        if k != "id" and v is not None:
                            setattr(existing, k, v)
                else:
                    s.add(Company(**data))
            s.commit()
        return list(set(all_ids))

    def get_companies(self, uncrawled_detail: bool = False,
                      limit: int | None = None,
                      company_ids: list[int] | None = None) -> list[Company]:
        with self.session() as s:
            q = s.query(Company)
            if uncrawled_detail:
                q = q.filter(Company.detail_crawled == False)
            if company_ids is not None:
                q = q.filter(Company.id.in_(company_ids))
            q = q.order_by(Company.id)
            if limit:
                q = q.limit(limit)
            rows = q.all()
            s.expunge_all()
            return rows

    def save_detail(self, data: dict):
        with self.session() as s:
            existing = s.get(CompanyDetail, data["company_id"])
            if existing:
                for k, v in data.items():
                    if v is not None:
                        setattr(existing, k, v)
            else:
                s.add(CompanyDetail(**data))
            # mark company detail_crawled
            comp = s.get(Company, data["company_id"])
            if comp:
                comp.detail_crawled = True
            s.commit()

    def save_details_batch(self, details: list[dict]):
        """Batch save multiple details in 1 transaction - much faster"""
        if not details:
            return
        with self.session() as s:
            for data in details:
                existing = s.get(CompanyDetail, data["company_id"])
                if existing:
                    for k, v in data.items():
                        if v is not None:
                            setattr(existing, k, v)
                else:
                    s.add(CompanyDetail(**data))
                comp = s.get(Company, data["company_id"])
                if comp:
                    comp.detail_crawled = True
            s.commit()

    # ---------- Stats ----------

    def get_stats(self) -> dict:
        with self.session() as s:
            return {
                "categories_total": s.query(Category).count(),
                "categories_crawled": s.query(Category).filter_by(crawled=True).count(),
                "companies_total": s.query(Company).count(),
                "companies_detail_crawled": s.query(Company).filter_by(detail_crawled=True).count(),
                "details_total": s.query(CompanyDetail).count(),
                "provinces": s.query(func.count(func.distinct(Company.province))).scalar() or 0,
            }

    # ---------- Reset ----------

    def reset_crawl_status(self, pipe: str | None = None):
        with self.session() as s:
            if pipe in (None, "0"):
                s.query(Category).update({"crawled": False})
            if pipe in (None, "1"):
                s.query(Category).update({"crawled": False})
            if pipe in (None, "2"):
                s.query(Company).update({"detail_crawled": False})
            s.commit()

    # ---------- Export ----------

    # Column mapping: db_field -> Vietnamese header
    EXPORT_COLUMNS = [
        # === THONG TIN CO BAN ===
        ("name",             "Ten cong ty"),
        ("province",         "Tinh/TP"),
        ("address",          "Dia chi"),
        ("all_phones",       "Dien thoai"),
        ("email",            "Email"),
        ("all_websites",     "Website"),
        ("zalo",             "Zalo"),
        ("fax",              "Fax"),
        # === LIEN HE ===
        ("contact_name",     "Nguoi lien he"),
        ("contact_phone",    "SDT lien he"),
        # === PHAN LOAI ===
        ("category_name",    "Nganh nghe"),
        ("sub_category_name","Nhom nganh"),
        ("industry",         "Linh vuc KD"),
        ("business_type",    "Loai hinh"),
        ("tax_code",         "Ma so thue"),
        ("year_established", "Nam thanh lap"),
        ("employee_count",   "So nhan vien"),
        ("main_market",      "Thi truong chinh"),
        ("main_customers",   "Khach hang chinh"),
        # === SAN PHAM ===
        ("product_count",    "So san pham"),
        ("products",         "Danh sach san pham"),
        # === DANH GIA ===
        ("sponsor_level",    "Cap do tai tro"),
        ("is_verified",      "Da xac thuc"),
        ("is_outdated",      "Het hieu luc"),
        ("image_count",      "So anh"),
        # === MO TA ===
        ("description",      "Gioi thieu"),
        # === META ===
        ("url",              "Link trang vang"),
        ("id",               "ID"),
    ]

    def _build_merged_rows(self) -> list[dict]:
        """Gop Company + CompanyDetail thanh 1 bang duy nhat, dedupe theo ID"""
        with self.session() as s:
            companies = s.query(Company).order_by(Company.id).all()
            # Load all details into a map
            details_map = {}
            for d in s.query(CompanyDetail).all():
                details_map[d.company_id] = d

            seen_ids = set()
            rows = []
            for c in companies:
                if c.id in seen_ids:
                    continue
                seen_ids.add(c.id)

                detail = details_map.get(c.id)

                # Merge phones: listing phones + detail all_phones -> dedupe
                phones_set = []
                # From listing
                if c.phones:
                    try:
                        for p in json.loads(c.phones):
                            if p and p not in phones_set:
                                phones_set.append(p)
                    except (json.JSONDecodeError, TypeError):
                        pass
                if c.hotline and c.hotline not in phones_set:
                    phones_set.append(c.hotline)
                # From detail
                if detail and detail.all_phones:
                    for p in detail.all_phones.split("; "):
                        p = p.strip()
                        if p and p not in phones_set:
                            phones_set.append(p)

                # Merge websites
                websites = []
                if c.website:
                    websites.append(c.website)
                if detail and detail.all_websites:
                    for w in detail.all_websites.split("; "):
                        w = w.strip()
                        if w and w not in websites:
                            websites.append(w)

                row = {
                    "id": c.id,
                    "name": detail.full_name if detail and detail.full_name else c.name,
                    "province": c.province or "",
                    "address": c.address or "",
                    "all_phones": "; ".join(phones_set),
                    "email": (detail.contact_email if detail and detail.contact_email
                              else c.email or ""),
                    "all_websites": "; ".join(websites),
                    "zalo": c.zalo or "",
                    "fax": detail.fax if detail else "",
                    "contact_name": detail.contact_name if detail else "",
                    "contact_phone": detail.contact_phone if detail else "",
                    "category_name": c.category_name or "",
                    "sub_category_name": c.sub_category_name or "",
                    "industry": detail.industry if detail else "",
                    "business_type": detail.business_type if detail else "",
                    "tax_code": detail.tax_code if detail else "",
                    "year_established": detail.year_established if detail else "",
                    "employee_count": detail.employee_count if detail else "",
                    "main_market": detail.main_market if detail else "",
                    "main_customers": detail.main_customers if detail else "",
                    "product_count": detail.product_count if detail else 0,
                    "products": detail.products if detail else "",
                    "sponsor_level": c.sponsor_level or "",
                    "is_verified": "Co" if c.is_verified else "",
                    "is_outdated": "Het hieu luc" if c.is_outdated else "",
                    "image_count": detail.image_count if detail else 0,
                    "description": detail.description if detail else "",
                    "url": c.url or "",
                }
                rows.append(row)

            return rows

    def export_excel(self, output_dir: str, filename: str = "") -> str:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

        Path(output_dir).mkdir(parents=True, exist_ok=True)
        rows = self._build_merged_rows()
        if not rows:
            logger.info("  Khong co du lieu de xuat")
            return ""

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Danh sach cong ty"

        # Header style
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin"),
        )
        wrap_align = Alignment(vertical="top", wrap_text=True)

        # Write headers
        headers = [col[1] for col in self.EXPORT_COLUMNS]
        fields = [col[0] for col in self.EXPORT_COLUMNS]
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

        # Write data
        for row_idx, row_data in enumerate(rows, 2):
            for col_idx, field in enumerate(fields, 1):
                val = row_data.get(field, "")
                if val is None:
                    val = ""
                if isinstance(val, int) and val == 0 and field in ("product_count", "image_count"):
                    val = ""
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.alignment = wrap_align
                cell.border = thin_border

        # Auto-width (approximate)
        col_widths = {
            "Ten cong ty": 40, "Tinh/TP": 15, "Dia chi": 45,
            "Dien thoai": 25, "Email": 30, "Website": 35,
            "Zalo": 14, "Fax": 16,
            "Nguoi lien he": 20, "SDT lien he": 16,
            "Nganh nghe": 25, "Nhom nganh": 30, "Linh vuc KD": 30, "Loai hinh": 18,
            "Ma so thue": 15, "Nam thanh lap": 12, "So nhan vien": 14,
            "Thi truong chinh": 18, "Khach hang chinh": 30,
            "So san pham": 12, "Danh sach san pham": 50,
            "Cap do tai tro": 16, "Da xac thuc": 11, "Het hieu luc": 12, "So anh": 8,
            "Gioi thieu": 60, "Link trang vang": 45, "ID": 12,
        }
        for col_idx, header in enumerate(headers, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = col_widths.get(header, 15)

        # Freeze header row
        ws.freeze_panes = "A2"
        # Auto filter
        ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(headers))}1"

        base = filename or "trangvang_data"
        filepath = os.path.join(output_dir, f"{base}.xlsx")
        saved = False
        for attempt in range(3):
            try:
                wb.save(filepath)
                saved = True
                break
            except PermissionError:
                if attempt == 0:
                    old_path = filepath + ".old"
                    try:
                        if os.path.exists(old_path):
                            os.remove(old_path)
                        os.rename(filepath, old_path)
                        continue
                    except (PermissionError, OSError):
                        pass
                ts = datetime.now().strftime("%H%M%S")
                filepath = os.path.join(output_dir, f"{base}_{ts}.xlsx")
                wb.save(filepath)
                print(f"  ! File cu dang mo, luu moi: {os.path.basename(filepath)}")
                saved = True
                break
        print(f"  Xuat {len(rows)} cong ty -> {filepath}")
        return filepath

    def export_csv(self, output_dir: str, filename: str = "") -> str:
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        rows = self._build_merged_rows()
        if not rows:
            logger.info("  Khong co du lieu de xuat")
            return ""

        headers = [col[1] for col in self.EXPORT_COLUMNS]
        fields = [col[0] for col in self.EXPORT_COLUMNS]

        base = filename or "trangvang_data"
        filepath = os.path.join(output_dir, f"{base}.csv")
        try:
            f = open(filepath, "w", newline="", encoding="utf-8-sig")
        except PermissionError:
            old_path = filepath + ".old"
            try:
                if os.path.exists(old_path):
                    os.remove(old_path)
                os.rename(filepath, old_path)
                f = open(filepath, "w", newline="", encoding="utf-8-sig")
            except (PermissionError, OSError):
                ts = datetime.now().strftime("%H%M%S")
                filepath = os.path.join(output_dir, f"{base}_{ts}.csv")
                f = open(filepath, "w", newline="", encoding="utf-8-sig")
                print(f"  ! File cu dang mo, luu moi: {os.path.basename(filepath)}")

        with f:
            writer = csv.writer(f)
            writer.writerow(headers)
            for row_data in rows:
                writer.writerow([row_data.get(field, "") or "" for field in fields])

        print(f"  Xuat {len(rows)} cong ty -> {filepath}")
        return filepath

    def export_csv_by_category(self, output_dir: str) -> list[str]:
        """Export CSV chia theo nganh nghe (category_name), moi nganh 1 file"""
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        rows = self._build_merged_rows()
        if not rows:
            print("  Khong co du lieu de xuat")
            return []

        # Group by category_name
        from collections import defaultdict
        groups = defaultdict(list)
        for row in rows:
            cat = row.get("category_name", "").strip() or "Khong_phan_loai"
            groups[cat].append(row)

        headers = [col[1] for col in self.EXPORT_COLUMNS]
        fields = [col[0] for col in self.EXPORT_COLUMNS]
        cat_dir = os.path.join(output_dir, "theo_nganh")
        Path(cat_dir).mkdir(parents=True, exist_ok=True)

        files = []
        for cat_name, cat_rows in sorted(groups.items(), key=lambda x: -len(x[1])):
            # Sanitize filename
            safe_name = re.sub(r'[<>:"/\\|?*]', '_', cat_name)
            safe_name = re.sub(r'\s+', '_', safe_name).strip('_')[:80]
            filepath = os.path.join(cat_dir, f"{safe_name}.csv")

            with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                for row_data in cat_rows:
                    writer.writerow([row_data.get(field, "") or "" for field in fields])

            files.append((cat_name, len(cat_rows), filepath))

        # Summary
        print(f"  Xuat {len(rows)} cong ty -> {len(files)} file theo nganh")
        for cat_name, count, fp in files[:10]:
            print(f"    {count:>5} CT | {cat_name}")
        if len(files) > 10:
            print(f"    ... va {len(files) - 10} nganh khac")
        print(f"  Thu muc: {os.path.abspath(cat_dir)}")
        return [fp for _, _, fp in files]


# (RateLimiter removed - server has no rate limiting, semaphore is sufficient)


# =====================================================
# PARSER HELPERS
# =====================================================


def _text(node) -> str:
    if node is None:
        return ""
    t = node.text(strip=True)
    return re.sub(r"\s+", " ", t).strip() if t else ""


def _attr(node, attr: str, default: str = "") -> str:
    if node is None:
        return default
    return node.attributes.get(attr, default) or default


def _safe_filename(name: str) -> str:
    """Chuyen ten nganh thanh ten file an toan. VD: 'Ốc, Vít, Bù Loong' -> 'oc_vit_bu_loong'"""
    import unicodedata
    # Normalize unicode
    s = unicodedata.normalize("NFD", name)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")  # bo dau
    s = s.lower().strip()
    s = re.sub(r"[^\w\s-]", " ", s)  # bo ky tu dac biet
    s = re.sub(r"[\s_]+", "_", s).strip("_")  # space -> underscore
    return s[:80] if s else "trangvang_data"


def _extract_id_from_url(url: str, segment: str) -> int | None:
    """Extract numeric ID from URL like /listings/12345/slug.html"""
    m = re.search(rf"/{segment}/(\d+)/", url)
    return int(m.group(1)) if m else None


def _extract_phones(node) -> list[str]:
    """Extract all phone numbers from tel: links within a node"""
    phones = []
    if node is None:
        return phones
    for a in node.css("a[href^='tel:']"):
        href = _attr(a, "href")
        phone = href.replace("tel:", "").strip()
        if phone and phone not in phones:
            phones.append(phone)
    return phones


# =====================================================
# TRANG VANG CRAWLER
# =====================================================


class TrangVangCrawler:

    def __init__(self, config: CrawlerConfig, db: DatabaseManager):
        self.config = config
        self.db = db
        self._session: aiohttp.ClientSession | None = None
        self._sem: asyncio.Semaphore | None = None

    async def __aenter__(self):
        conn = aiohttp.TCPConnector(
            limit=self.config.concurrency,
            limit_per_host=self.config.concurrency,
            ttl_dns_cache=300,
            ssl=False,
        )
        timeout = aiohttp.ClientTimeout(total=self.config.timeout)
        self._session = aiohttp.ClientSession(connector=conn, timeout=timeout)
        self._sem = asyncio.Semaphore(self.config.concurrency)
        return self

    async def __aexit__(self, *args):
        if self._session:
            await self._session.close()

    def _headers(self) -> dict:
        return {
            "User-Agent": random.choice(USER_AGENTS),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8",
            "Accept-Encoding": "gzip, deflate",
            "Referer": BASE_URL + "/",
        }

    async def _fetch(self, url: str) -> str | None:
        """Fetch URL with retry + semaphore (true concurrent, no serial lock)"""
        async with self._sem:
            for attempt in range(1, self.config.max_retries + 1):
                try:
                    async with self._session.get(url, headers=self._headers()) as resp:
                        resp.raise_for_status()
                        raw = await resp.read()
                        # Try utf-8 first, fallback
                        for enc in ("utf-8", "utf-8-sig"):
                            try:
                                return raw.decode(enc)
                            except UnicodeDecodeError:
                                continue
                        return raw.decode("utf-8", errors="replace")
                except Exception as e:
                    if attempt == self.config.max_retries:
                        logger.warning(f"FAIL {url}: {e}")
                        return None
                    await asyncio.sleep(1.5 ** attempt)
        return None

    async def _fetch_bytes(self, url: str) -> bytes | None:
        async with self._sem:
            try:
                async with self._session.get(url, headers=self._headers()) as resp:
                    resp.raise_for_status()
                    return await resp.read()
            except Exception:
                return None

    # =================================================
    # PIPE 0: THU THAP DANH MUC
    # =================================================

    async def pipe0_collect_categories(self):
        """Crawl findex.asp de lay tat ca danh muc"""
        logger.info("=== PIPE 0: Thu thap danh muc tu findex.asp ===")
        html = await self._fetch(FINDEX_URL)
        if not html:
            logger.error("Khong the fetch findex.asp")
            return

        tree = HTMLParser(html)
        categories = []

        # Parent categories (co class ps-2 fw500)
        for p in tree.css("p.fw500 a[href*='/categories/']"):
            url = _attr(p, "href")
            cat_id = _extract_id_from_url(url, "categories")
            if not cat_id:
                continue
            name = _text(p)
            # Remove counter from name
            counter_span = p.css_first("span.counter_number")
            count = 0
            if counter_span:
                count_text = _text(counter_span)
                m = re.search(r"\((\d+)\)", count_text)
                count = int(m.group(1)) if m else 0
                name = name.replace(count_text, "").strip()

            slug = url.split("/")[-1].replace(".html", "")
            categories.append({
                "id": cat_id, "name": name, "slug": slug, "url": url,
                "total_companies": count, "is_sub": False,
            })

        # Sub-categories (co class ps-4, inside <small>)
        for p in tree.css("p.ps-4 small a[href*='/categories/']"):
            url = _attr(p, "href")
            cat_id = _extract_id_from_url(url, "categories")
            if not cat_id:
                continue
            name = _text(p)
            counter_span = p.css_first("span.counter_number")
            count = 0
            if counter_span:
                count_text = _text(counter_span)
                m = re.search(r"\((\d+)\)", count_text)
                count = int(m.group(1)) if m else 0
                name = name.replace(count_text, "").strip()

            slug = url.split("/")[-1].replace(".html", "")
            categories.append({
                "id": cat_id, "name": name, "slug": slug, "url": url,
                "total_companies": count, "is_sub": True,
            })

        if categories:
            self.db.upsert_categories(categories)
            logger.info(f"  Da luu {len(categories)} danh muc")
        else:
            logger.warning("  Khong tim thay danh muc nao")

    async def pipe0_collect_from_group(self, group_id: int):
        """Crawl 1 group page de lay categories trong group"""
        slug = GROUP_IDS.get(group_id, str(group_id))
        url = f"{GINDEX_URL}/{group_id}/{slug}"
        logger.info(f"  Crawl group {group_id}: {url}")

        html = await self._fetch(url)
        if not html:
            return

        tree = HTMLParser(html)
        categories = []
        for a in tree.css("a[href*='/categories/']"):
            cat_url = _attr(a, "href")
            cat_id = _extract_id_from_url(cat_url, "categories")
            if not cat_id:
                continue
            name = _text(a)
            counter_span = a.css_first("span.counter_number")
            count = 0
            if counter_span:
                count_text = _text(counter_span)
                m = re.search(r"\((\d+)\)", count_text)
                count = int(m.group(1)) if m else 0
                name = name.replace(count_text, "").strip()

            cat_slug = cat_url.split("/")[-1].replace(".html", "")
            categories.append({
                "id": cat_id, "name": name, "slug": cat_slug, "url": cat_url,
                "total_companies": count, "group_id": group_id, "is_sub": False,
            })

        if categories:
            self.db.upsert_categories(categories)
            logger.info(f"    -> {len(categories)} danh muc trong group {group_id}")

    async def pipe0_collect_all_groups(self):
        """Crawl tat ca groups"""
        logger.info("=== PIPE 0: Thu thap danh muc tu tat ca groups ===")
        tasks = [self.pipe0_collect_from_group(gid) for gid in GROUP_IDS]
        await asyncio.gather(*tasks)

    # =================================================
    # PIPE 1: CRAWL LISTING
    # =================================================

    def _parse_listing_page(self, html: str) -> list[dict]:
        """Parse 1 trang listing -> list company cards"""
        tree = HTMLParser(html)
        companies = []

        for card in tree.css("div.div_list_cty > div.shadow.rounded-3"):
            # Company name + URL
            name_a = card.css_first("div.listings_center > h2 > a")
            if not name_a:
                name_a = card.css_first("h2 > a")
            if not name_a:
                continue

            comp_url = _attr(name_a, "href")
            comp_id = _extract_id_from_url(comp_url, "listings")
            if not comp_id:
                continue

            name = _text(name_a)
            slug = comp_url.split("/")[-1].replace(".html", "")

            # Category
            cat_span = card.css_first("span.nganh_listing_txt")
            category_name = _text(cat_span)

            # Address - find the div containing fa-location-dot icon
            address = ""
            province = ""

            # Strategy 1: find <i class="fa-location-dot"> anywhere in card
            loc_icon = card.css_first("i.fa-location-dot")
            if loc_icon:
                # Address is in the parent div's text
                addr_container = loc_icon.parent
                if addr_container:
                    address = _text(addr_container)
                    # Province is in span.fw500 or strong
                    prov_el = addr_container.css_first("span.fw500")
                    if not prov_el:
                        prov_el = addr_container.css_first("strong")
                    province = _text(prov_el)

            # Strategy 2: non-sponsored layout (listing_diachi_nologo)
            if not address:
                nologo_div = card.css_first("div.listing_diachi_nologo")
                if nologo_div:
                    for div in nologo_div.css("div"):
                        txt = _text(div)
                        if "," in txt and ("Nam" in txt or province):
                            address = txt
                            break
                    strong = nologo_div.css_first("strong")
                    if not province:
                        province = _text(strong)

            # Clean address: remove "Viet Nam" / "Việt Nam" and province
            for suffix in ("Viet Nam", "Việt Nam", "Vi\u1ec7t Nam"):
                address = address.replace(suffix, "")
            address = address.strip(", \t\n")
            if province:
                address = address.replace(province, "").strip(", \t\n")

            # Phones - from listing_dienthoai div AND any tel: links in card
            phone_div = card.css_first("div.listing_dienthoai")
            phones = _extract_phones(phone_div)

            # Also get phones from non-sponsored layout (listing_diachi_nologo)
            if not phones:
                nologo = card.css_first("div.listing_diachi_nologo")
                if nologo:
                    phones = _extract_phones(nologo)

            # Fallback: get ALL tel: links in the card
            if not phones:
                phones = _extract_phones(card)

            # Hotline + Zalo
            hotline = ""
            zalo = ""
            for div in card.css("div"):
                text_content = _text(div)
                if "Hotline:" in text_content or "Di dong:" in text_content:
                    hp = _extract_phones(div)
                    if hp:
                        hotline = hp[0]
                    zalo_a = div.css_first("a[href*='zalo.me']")
                    if zalo_a:
                        zalo_href = _attr(zalo_a, "href")
                        zalo = zalo_href.replace("https://zalo.me/", "").strip("/")
                    break

            # Email + Website
            email = ""
            website = ""
            email_section = card.css_first("div.email_web_section")
            if email_section:
                mail_a = email_section.css_first("a[href^='mailto:']")
                if mail_a:
                    email = _attr(mail_a, "href").replace("mailto:", "").strip()
                web_a = email_section.css_first("a[rel='nofollow']")
                if not web_a:
                    for a in email_section.css("a"):
                        href = _attr(a, "href")
                        if href and "mailto:" not in href and "trangvang" not in href:
                            web_a = a
                            break
                if web_a:
                    website = _attr(web_a, "href")

            # Logo
            logo_url = ""
            logo_img = card.css_first("img[src*='logo.trangvangvietnam']")
            if logo_img:
                logo_url = _attr(logo_img, "src")

            # Sponsor level: "3 - Diamond Sponsor" -> "3 Diamond"
            is_sponsor = False
            sponsor_level = ""
            sponsor_span = card.css_first("span.star_text")
            if sponsor_span:
                is_sponsor = True
            sponsor_img = card.css_first("img[title*='Sponsor']")
            if sponsor_img:
                is_sponsor = True
                raw_title = _attr(sponsor_img, "title")  # "3 - Diamond Sponsor"
                m = re.search(r"(\d+)\s*-?\s*(Diamond|Gold|Silver|Bronze)", raw_title, re.IGNORECASE)
                sponsor_level = f"{m.group(1)} {m.group(2)}" if m else raw_title

            # Verified
            is_verified = bool(card.css_first("img[src*='daxacthuc']") or
                               card.css_first("img[src*='xacthuc']"))

            # Outdated / inactive: "Thông tin này đã không còn chính xác"
            is_outdated = bool(card.css_first("i.fa-circle-xmark"))

            # Ad text
            ad_text = ""
            ad_div = card.css_first("div.div_textqc small.text_qc")
            if ad_div:
                ad_text = _text(ad_div)

            # Last updated (non-sponsored)
            last_updated = ""
            date_small = card.css_first("div.star_mb small")
            if date_small:
                dt_text = _text(date_small)
                m = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", dt_text)
                if m:
                    last_updated = m.group(1)

            companies.append({
                "id": comp_id,
                "name": name,
                "slug": slug,
                "url": comp_url,
                "category_name": category_name,
                "address": address,
                "province": province,
                "phones": json.dumps(phones, ensure_ascii=False) if phones else "",
                "hotline": hotline,
                "zalo": zalo,
                "email": email,
                "website": website,
                "logo_url": logo_url,
                "is_sponsor": is_sponsor,
                "sponsor_level": sponsor_level,
                "is_verified": is_verified,
                "is_outdated": is_outdated,
                "ad_text": ad_text,
                "last_updated": last_updated,
            })

        return companies

    def _get_total_pages(self, html: str) -> int:
        """Extract total pages from pagination"""
        tree = HTMLParser(html)
        paging = tree.css_first("div#paging")
        if not paging:
            return 1

        max_page = 1
        for a in paging.css("a[href*='page=']"):
            href = _attr(a, "href")
            m = re.search(r"page=(\d+)", href)
            if m:
                p = int(m.group(1))
                if p > max_page:
                    max_page = p
        return max_page

    def _get_total_count(self, html: str) -> int:
        """Extract total company count"""
        tree = HTMLParser(html)
        counter = tree.css_first("span.ketquatimkiem_counter")
        if counter:
            m = re.search(r"(\d+)", _text(counter))
            if m:
                return int(m.group(1))
        return 0

    async def _fetch_all_pages(self, base_url: str, total_pages: int,
                                prefetched_html: str | None = None,
                                source: str = "category",
                                cat_id: int | None = None) -> list[dict]:
        """Fetch tat ca pages cua 1 listing URL, return raw list (chua dedupe)."""
        all_companies = []

        # Page 1
        html = prefetched_html or await self._fetch(base_url)
        if html:
            cards = self._parse_listing_page(html)
            for c in cards:
                c["source_page"] = source
                if cat_id:
                    c["category_id"] = cat_id
            all_companies.extend(cards)

        # Pages 2..N concurrent
        if total_pages > 1:
            sep = "&" if "?" in base_url else "?"
            page_urls = [f"{base_url}{sep}page={p}" for p in range(2, total_pages + 1)]
            results = await asyncio.gather(*[self._fetch(u) for u in page_urls])

            for page_html in results:
                if page_html:
                    cards = self._parse_listing_page(page_html)
                    for c in cards:
                        c["source_page"] = source
                        if cat_id:
                            c["category_id"] = cat_id
                    all_companies.extend(cards)

        return all_companies

    def _extract_sub_categories(self, html: str, parent_id: int) -> list[tuple[int, str]]:
        """Lay danh sach sub-category URLs CHI trong section 'Nhom nganh nghe'.
        Sidebar co 3+ sections: 'Nganh nghe' (parent), 'Nhom nganh nghe' (sub),
        'Nganh xem them' (related - KHONG lay). Parse theo section header."""
        tree = HTMLParser(html)
        sidebar = tree.css_first("div.div_nganhnghe_pc")
        if not sidebar:
            return []

        subs = []
        in_sub_section = False

        for p in sidebar.css("p"):
            p_class = p.attributes.get("class", "") or ""

            # Detect section headers: <p class="... h6"> with icon
            if "h6" in p_class:
                header_text = _text(p).lower()
                # "Nhóm ngành nghề" = sub-categories we want
                import unicodedata as _ud
                header_norm = "".join(
                    c for c in _ud.normalize("NFKD", header_text) if not _ud.combining(c)
                ).lower()
                in_sub_section = "nhom" in header_norm and "nganh" in header_norm
                continue

            # Only collect links when inside "Nhom nganh nghe" section
            if not in_sub_section:
                continue

            a = p.css_first("a[href*='/categories/']")
            if not a:
                continue
            href = _attr(a, "href")
            m = re.search(r"/categories/(\d+)/", href)
            if not m:
                continue
            cid = int(m.group(1))
            if cid == parent_id:
                continue
            full_url = href if href.startswith("http") else BASE_URL + href
            subs.append((cid, full_url))

        return subs

    async def _crawl_listing_url(self, base_url: str, max_pages: int | None = None,
                                  limit: int | None = None,
                                  source: str = "category",
                                  cat_id: int | None = None,
                                  pb: ProgressBar | None = None,
                                  prefetched_html: str | None = None) -> list[dict]:
        """Crawl ALL pages + sub-categories + multi-round.
        Server xao tron ket qua moi request -> can nhieu round de bat du.
        Returns list of unique company dicts da luu vao DB."""
        html = prefetched_html or await self._fetch(base_url)
        if not html:
            return []

        total_pages = self._get_total_pages(html)
        total_count = self._get_total_count(html)
        if max_pages:
            total_pages = min(total_pages, max_pages)

        if pb:
            pb.set_total(total_pages)

        # Lay ten nganh cha tu H1
        parent_name = ""
        h1 = HTMLParser(html).css_first("h1")
        if h1:
            parent_name = _text(h1).strip()
            parent_name = re.sub(r"\s*[-–]\s*Trang\s+V[àa]ng.*$", "", parent_name, flags=re.IGNORECASE)

        seen: dict[int, dict] = {}

        def _add(comps, cat_name="", sub_name=""):
            """Add companies to seen dict, return count of new."""
            added = 0
            for c in comps:
                if c["id"] not in seen:
                    c["category_name"] = cat_name or parent_name
                    c["sub_category_name"] = sub_name
                    seen[c["id"]] = c
                    added += 1
                elif sub_name and not seen[c["id"]].get("sub_category_name"):
                    seen[c["id"]]["sub_category_name"] = sub_name
            return added

        # === PARENT: lap cho den khi het CT moi (server xao tron) ===
        round1 = await self._fetch_all_pages(
            base_url, total_pages,
            prefetched_html=html, source=source, cat_id=cat_id,
        )
        if pb:
            pb.update(True, count=total_pages)
        _add(round1)

        dup_ratio = 1 - (len(seen) / len(round1)) if round1 else 0
        if dup_ratio > 0.10 and total_pages > 1:
            # Server xao tron -> lap them cho den khi khong con CT moi
            for rd in range(2, 8):  # toi da 7 rounds
                comps = await self._fetch_all_pages(
                    base_url, total_pages, source=source, cat_id=cat_id,
                )
                new = _add(comps)
                if new:
                    print(f"  Round {rd}: +{new} moi (total: {len(seen)})")
                if new < 5:  # duoi 5 CT moi = gan het
                    break

        # === SUB-CATEGORIES: lap tung nhom cho den het ===
        if cat_id and total_pages > 1:
            sub_cats = self._extract_sub_categories(html, cat_id)
            if sub_cats:
                before = len(seen)

                async def _fetch_subcat_full(sub_cid, sub_url):
                    """Fetch 1 sub-cat, lap nhieu round neu can."""
                    h = await self._fetch(sub_url)
                    if not h:
                        return [], ""
                    sub_h1 = HTMLParser(h).css_first("h1")
                    sub_name = _text(sub_h1).strip() if sub_h1 else ""
                    sub_name = re.sub(r"\s*[-–]\s*Trang\s+V[àa]ng.*$", "",
                                      sub_name, flags=re.IGNORECASE)
                    tp = self._get_total_pages(h)
                    # Round 1
                    all_comps = await self._fetch_all_pages(
                        sub_url, tp, prefetched_html=h, source=source,
                    )
                    # Round 2 neu nhieu dup
                    raw = len(all_comps)
                    unique = len(set(c["id"] for c in all_comps))
                    if raw > 0 and (1 - unique / raw) > 0.10 and tp > 1:
                        r2 = await self._fetch_all_pages(
                            sub_url, tp, source=source,
                        )
                        all_comps.extend(r2)
                    return all_comps, sub_name

                tasks = [_fetch_subcat_full(cid, u) for cid, u in sub_cats]
                results = await asyncio.gather(*tasks)
                for comps, sub_name in results:
                    _add(comps, parent_name, sub_name)

                sub_new = len(seen) - before
                if sub_new:
                    print(f"  Sub-categories: +{sub_new} cong ty (total: {len(seen)})")

        all_companies = list(seen.values())

        if limit:
            all_companies = all_companies[:limit]

        self.db.upsert_companies(all_companies)
        return all_companies

    async def pipe1_crawl_listings(self, category_ids: list[int] | None = None,
                                    group_id: int | None = None,
                                    letter: str | None = None,
                                    max_pages: int | None = None,
                                    limit: int | None = None):
        """Crawl listing - multiple categories concurrently"""
        cats = self.db.get_categories(
            uncrawled_only=True,
            group_id=group_id,
            category_ids=category_ids,
            letter=letter,
        )

        if not cats:
            print("  Khong co danh muc nao de crawl.")
            return

        total_new = 0
        pb = ProgressBar(len(cats), prefix="Danh muc")

        # Process categories in concurrent batches
        batch_size = min(self.config.concurrency, 5)  # 5 categories at a time
        for i in range(0, len(cats), batch_size):
            batch = cats[i:i + batch_size]

            async def _crawl_cat(cat):
                comps = await self._crawl_listing_url(
                    cat.url, max_pages=max_pages, limit=limit,
                    source="category", cat_id=cat.id,
                )
                self.db.mark_category_crawled(cat.id)
                return len(comps)

            results = await asyncio.gather(*[_crawl_cat(c) for c in batch])
            for count in results:
                total_new += count
                pb.update(True)

        pb.finish()
        print(f"  Tong: {total_new} cong ty moi")

    async def pipe1_crawl_search(self, keyword: str, where: str = "",
                                  max_pages: int | None = None,
                                  limit: int | None = None):
        """Crawl theo search keyword"""
        logger.info(f"=== PIPE 1: Tim kiem '{keyword}' tai '{where}' ===")
        params = f"keyword={quote(keyword, safe='')}"
        if where:
            params += f"&where={quote(where, safe='')}"
        url = f"{SEARCH_URL}?{params}"

        await self._crawl_listing_url(
            url, max_pages=max_pages, limit=limit, source="search",
        )

    async def pipe1_crawl_province(self, category_id: int, province_slug: str,
                                    max_pages: int | None = None,
                                    limit: int | None = None):
        """Crawl theo tinh thanh"""
        url = f"{BASE_URL}/cateprovinces/{category_id}/{province_slug}.html"
        logger.info(f"=== PIPE 1: Crawl theo tinh: {url} ===")
        await self._crawl_listing_url(
            url, max_pages=max_pages, limit=limit, source="province",
        )

    # =================================================
    # PIPE 2: CRAWL CHI TIET
    # =================================================

    def _parse_detail_page(self, html: str, company_id: int) -> dict:
        """Parse trang chi tiet cong ty"""
        tree = HTMLParser(html)
        data = {"company_id": company_id}

        # Full name
        h1 = tree.css_first("h1.fs-3")
        data["full_name"] = _text(h1)

        # Business info from #section4 (div_23_txt + div_77 pairs)
        import unicodedata as _ud
        label_map = {
            "LOAI HINH": "business_type",
            "NGANH NGHE": "industry",
            "THI TRUONG": "main_market",
            "KHACH HANG": "main_customers",
            "MA SO THUE": "tax_code",
            "NAM THANH LAP": "year_established",
            "SO LUONG": "employee_count",
            "NHAN VIEN": "employee_count",
            "FAX": "fax",
        }
        for label_div in tree.css("#section4 div.div_23_txt"):
            raw = _text(label_div)
            # Strip Vietnamese diacritics: MÃ SỐ THUẾ -> MA SO THUE
            label_text = "".join(
                c for c in _ud.normalize("NFKD", raw) if not _ud.combining(c)
            ).upper()
            for keyword, field in label_map.items():
                if keyword in label_text:
                    parent = label_div.parent
                    if parent:
                        value_div = parent.css_first("div.div_77")
                        if value_div:
                            data[field] = _text(value_div)
                    break

        # All phones on detail page (merge all tel: links)
        all_phones = []
        for a in tree.css("a[href^='tel:']"):
            phone = _attr(a, "href").replace("tel:", "").strip()
            if phone and phone not in all_phones:
                all_phones.append(phone)
        data["all_phones"] = "; ".join(all_phones)

        # All websites (nofollow links, exclude trangvang/zalo/mailto)
        all_websites = []
        for a in tree.css("a[rel='nofollow']"):
            href = _attr(a, "href")
            if href and "trangvang" not in href and "zalo.me" not in href and "mailto:" not in href:
                if href not in all_websites:
                    all_websites.append(href)
        data["all_websites"] = "; ".join(all_websites)

        # Fax: try fa-fax icon first, then meta description
        if not data.get("fax"):
            fax_icon = tree.css_first("i.fa-fax")
            if fax_icon and fax_icon.parent:
                fax_text = _text(fax_icon.parent)
                # Remove "(số fax)" label, keep the number
                fax_text = re.sub(r"\(.*?fax.*?\)", "", fax_text, flags=re.IGNORECASE).strip()
                if fax_text:
                    data["fax"] = fax_text

        if not data.get("fax"):
            meta_desc = tree.css_first("meta[name='description']")
            if meta_desc:
                meta_text = _attr(meta_desc, "content")
                fax_m = re.search(r"[Ff]ax[:\s]*([0-9()\s\-\.]+)", meta_text)
                if fax_m:
                    data["fax"] = fax_m.group(1).strip()

        # Description
        desc_div = tree.css_first("div.gioithieucongty_img")
        if desc_div:
            data["description"] = _text(desc_div)

        # Contact info - normalize diacritics before matching
        contact_name = ""
        contact_phone = ""
        contact_email = ""

        for span in tree.css("span.span_lienhe"):
            span_text = _text(span)
            span_norm = "".join(
                c for c in _ud.normalize("NFKD", span_text) if not _ud.combining(c)
            ).upper()
            parent = span.parent
            if not parent:
                continue

            if "TEN LIEN HE" in span_norm or "NGUOI LIEN HE" in span_norm:
                parent_text = _text(parent)
                contact_name = parent_text.replace(span_text, "").strip(": \t")

            elif "DIEN THOAI" in span_norm or "DI DONG" in span_norm:
                phones_found = _extract_phones(parent)
                if phones_found:
                    contact_phone = phones_found[0]

            elif "EMAIL" in span_norm:
                parent_text = _text(parent)
                email_val = parent_text.replace(span_text, "").strip(": \t")
                if email_val and "@" in email_val:
                    contact_email = email_val

        # Fallback: a#email_lienhe
        if not contact_email:
            email_a = tree.css_first("a#email_lienhe")
            if email_a:
                href = _attr(email_a, "href")
                contact_email = href.replace("mailto:", "").strip()

        data["contact_name"] = contact_name
        data["contact_phone"] = contact_phone
        data["contact_email"] = contact_email

        # Product groups
        groups = []
        for a in tree.css("h3 a[href*='groupspdv']"):
            groups.append(_text(a))
        data["product_groups"] = json.dumps(groups, ensure_ascii=False) if groups else ""

        # Products - only store names (not URLs) for cleaner export
        products = []
        for a in tree.css("div.sanphamdichvu a[href*='products']"):
            name = _text(a)
            if name:
                products.append(name)
        data["products"] = "; ".join(products)
        data["product_count"] = len(products)

        # Images
        images = []
        for img in tree.css("img.img_style[src*='pic.trangvangvietnam']"):
            src = _attr(img, "src")
            if src:
                images.append(src)
        data["images"] = json.dumps(images, ensure_ascii=False) if images else ""
        data["image_count"] = len(images)

        return data

    async def _crawl_one_detail(self, company) -> dict | None:
        """Crawl chi tiet 1 cong ty, return parsed dict (not saved yet).
        company: Company object hoac dict co 'id' va 'url'.
        """
        cid = company.id if hasattr(company, "id") else company["id"]
        curl = company.url if hasattr(company, "url") else company["url"]
        html = await self._fetch(curl)
        if not html:
            return None

        try:
            return self._parse_detail_page(html, cid)
        except Exception as e:
            logger.error(f"  Parse error {curl}: {e}")
            return None

    async def pipe2_crawl_details(self, limit: int | None = None,
                                   company_ids: list[int] | None = None,
                                   force: bool = False):
        """Crawl chi tiet - concurrent fetch, batch DB save, auto-retry fails
        company_ids: chi crawl cac CT co id trong list (dung cho search mode)
        force: True = crawl tat ca (ke ca da crawl roi), dung cho search mode
        """
        if force and company_ids:
            # Search mode: crawl TAT CA CT vua tim duoc, khong can check flag
            companies = self.db.get_companies(company_ids=company_ids, limit=limit)
        else:
            companies = self.db.get_companies(uncrawled_detail=True, limit=limit,
                                              company_ids=company_ids)
        if not companies:
            print("  Tat ca cong ty da co chi tiet.")
            return

        print(f"  Lay chi tiet {len(companies)} cong ty...")
        pb = ProgressBar(len(companies), prefix="Chi tiet")
        failed_companies: list[Company] = []

        # Concurrent batches - fetch N at once, batch-save to DB
        batch_size = self.config.concurrency * 2
        for i in range(0, len(companies), batch_size):
            batch = companies[i:i + batch_size]
            tasks_list = [self._crawl_one_detail(c) for c in batch]
            results = await asyncio.gather(*tasks_list)

            # Batch save all successful results in 1 transaction
            details_to_save = [r for r in results if r is not None]
            self.db.save_details_batch(details_to_save)

            # Track failures
            for comp, result in zip(batch, results):
                if result is None:
                    failed_companies.append(comp)

            ok_count = len(details_to_save)
            fail_count = len(results) - ok_count
            pb.update(True, count=ok_count)
            if fail_count > 0:
                pb.update(False, count=fail_count)

        pb.finish()

        # Auto-retry failed companies (1 more round, sequential - more reliable)
        if failed_companies:
            print(f"  Retry {len(failed_companies)} cong ty bi loi...")
            retry_ok = 0
            still_failed: list[Company] = []
            for comp in failed_companies:
                result = await self._crawl_one_detail(comp)
                if result:
                    self.db.save_detail(result)
                    retry_ok += 1
                else:
                    still_failed.append(comp)
            if retry_ok:
                pb.ok += retry_ok
                pb.fail -= retry_ok
            print(f"  Retry: +{retry_ok} OK")

            # Report remaining failures
            if still_failed:
                print(f"\n  FAIL ({len(still_failed)} cong ty khong lay duoc):")
                for comp in still_failed:
                    print(f"    - {comp.name[:60]}  |  {comp.url}")
                print(f"  (Chay lai se tu dong retry cac cong ty nay)\n")

        print(f"  Hoan tat: OK={pb.ok}, FAIL={pb.fail}")

    async def crawl_details_direct(self, companies: list[dict]):
        """Crawl chi tiet TRUC TIEP tu danh sach companies (dict co 'id'+'url').
        Vao TAT CA link trong list, chia chunk da luong, khong dem truoc."""
        total = len(companies)
        if total == 0:
            print("  Khong co link nao.")
            return

        pb = ProgressBar(total, prefix="Chi tiet")
        failed: list[dict] = []

        # Chia chunk da luong - moi chunk = concurrency * 2
        chunk = self.config.concurrency * 2
        for i in range(0, total, chunk):
            batch = companies[i:i + chunk]
            results = await asyncio.gather(*[self._crawl_one_detail(c) for c in batch])

            # Luu ngay vao DB (crash-safe)
            ok_list = [r for r in results if r is not None]
            if ok_list:
                self.db.save_details_batch(ok_list)

            for comp, result in zip(batch, results):
                pb.update(result is not None)
                if result is None:
                    failed.append(comp)

        pb.finish()

        # Retry 1 vong cac link fail
        if failed:
            print(f"  Retry {len(failed)} link...")
            retry_ok = 0
            still_failed: list[dict] = []
            for comp in failed:
                result = await self._crawl_one_detail(comp)
                if result:
                    self.db.save_detail(result)
                    retry_ok += 1
                else:
                    still_failed.append(comp)
            print(f"  Retry: +{retry_ok} OK" + (
                f", {len(still_failed)} van fail" if still_failed else ""))
            if still_failed:
                for c in still_failed[:10]:
                    print(f"    X {c.get('name','?')[:50]}  {c['url']}")
                if len(still_failed) > 10:
                    print(f"    ... +{len(still_failed)-10} link khac")

    # =================================================
    # PIPE 3: DOWNLOAD ASSETS
    # =================================================

    async def pipe3_download_assets(self, logos: bool = True, images: bool = True):
        """Download logos va/hoac product images"""
        output = Path(self.config.output_dir)

        if logos:
            await self._download_logos(output / "logos")
        if images:
            await self._download_images(output / "images")

    async def _download_batch(self, items: list[tuple], pb: ProgressBar):
        """Download a list of (url, filepath) concurrently in batches"""
        batch_size = self.config.concurrency * 2
        for i in range(0, len(items), batch_size):
            batch = items[i:i + batch_size]

            async def _dl(url, filepath):
                filepath.parent.mkdir(parents=True, exist_ok=True)
                data = await self._fetch_bytes(url)
                if data:
                    filepath.write_bytes(data)
                return bool(data)

            results = await asyncio.gather(*[_dl(u, f) for u, f in batch])
            for ok in results:
                pb.update(ok)

    async def _download_logos(self, out_dir: Path):
        out_dir.mkdir(parents=True, exist_ok=True)
        companies = self.db.get_companies()
        to_download = [(c.logo_url, out_dir / f"L{c.id}.gif")
                       for c in companies if c.logo_url and not (out_dir / f"L{c.id}.gif").exists()]
        if not to_download:
            print("  Khong co logo nao can download")
            return

        pb = ProgressBar(len(to_download), prefix="Logos")
        await self._download_batch(to_download, pb)
        pb.finish()

    async def _download_images(self, out_dir: Path):
        out_dir.mkdir(parents=True, exist_ok=True)
        with self.db.session() as s:
            details = s.query(CompanyDetail).filter(
                CompanyDetail.images != "", CompanyDetail.images.isnot(None),
            ).all()
            s.expunge_all()

        to_download = []
        for d in details:
            try:
                imgs = json.loads(d.images)
            except (json.JSONDecodeError, TypeError):
                continue
            comp_dir = out_dir / str(d.company_id)
            for i, img_url in enumerate(imgs):
                ext = Path(urlparse(img_url).path).suffix or ".jpg"
                filepath = comp_dir / f"img_{i+1}{ext}"
                if not filepath.exists():
                    to_download.append((img_url, filepath))

        if not to_download:
            print("  Khong co anh nao can download")
            return

        pb = ProgressBar(len(to_download), prefix="Images")
        await self._download_batch(to_download, pb)
        pb.finish()

    # =================================================
    # UTILITY: AUTOCOMPLETE
    # =================================================

    async def autocomplete(self, query: str) -> list[str]:
        """Tra cuu autocomplete keyword"""
        url = f"{AUTOCOMPLETE_URL}?part={quote(query, safe='')}"
        html = await self._fetch(url)
        if not html:
            return []
        try:
            items = json.loads(html)
            return [unescape(s) for s in items if s]
        except json.JSONDecodeError:
            return []

    async def resolve_keyword_to_category(self, keyword: str) -> tuple[str, int | None] | None:
        """Resolve keyword -> category URL via search page.
        Search.asp has broken pagination (sponsored overlap), so we detect
        if search redirects/matches a category and return category URL instead.
        Returns (url, cat_id) or None if no match.
        """
        search_url = f"{SEARCH_URL}?keyword={quote(keyword, safe='')}"
        html = await self._fetch(search_url)
        if not html:
            return None

        tree = HTMLParser(html)

        # Check if the search page shows a category link matching the keyword
        # The left sidebar has category links with counters
        for a in tree.css("div.div_nganhnghe_pc a[href*='/categories/']"):
            cat_url = _attr(a, "href")
            cat_id = _extract_id_from_url(cat_url, "categories")
            if cat_id:
                return (cat_url, cat_id)

        # Also check h1 - sometimes search shows category page directly
        h1 = tree.css_first("h1.yp_red_color")
        if h1:
            h1_link = h1.css_first("a[href*='/categories/']")
            if h1_link:
                cat_url = _attr(h1_link, "href")
                cat_id = _extract_id_from_url(cat_url, "categories")
                if cat_id:
                    return (cat_url, cat_id)

        # No category match - return search URL (will have sponsored overlap)
        return (search_url, None)

    async def list_provinces(self, query: str) -> list[str]:
        """Tra cuu tinh thanh"""
        url = f"{PROVINCES_URL}?p={quote(query, safe='')}"
        html = await self._fetch(url)
        if not html:
            return []
        try:
            items = json.loads(html)
            return [unescape(s) for s in items if s]
        except json.JSONDecodeError:
            return []


# =====================================================
# CLI
# =====================================================


def show_stats(db: DatabaseManager):
    stats = db.get_stats()
    print("\n  ╔══ Thong ke ══════════════════════╗")
    labels = {
        "categories_total": "Danh muc",
        "categories_crawled": "Danh muc da crawl",
        "companies_total": "Cong ty",
        "companies_detail_crawled": "Co chi tiet",
        "details_total": "Chi tiet",
        "provinces": "Tinh/TP",
    }
    for k, v in stats.items():
        label = labels.get(k, k)
        print(f"  ║  {label:24s} {v:>6}")
    print("  ╚═══════════════════════════════════╝\n")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Trang Vang Vietnam Crawler - Crawl danh ba doanh nghiep",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Vi du su dung:
  %(prog)s --pipe 0                          # Thu thap danh muc
  %(prog)s --pipe 1                          # Crawl listings (tat ca danh muc)
  %(prog)s --pipe 1 --category 26360         # Chi 1 danh muc
  %(prog)s --pipe 1 --category 26360,15020   # Nhieu danh muc
  %(prog)s --pipe 1 --group 41               # Theo nhom nganh
  %(prog)s --pipe 1 --letter A               # Danh muc bat dau bang A
  %(prog)s --pipe 1 --search "vong bi"       # Tim kiem keyword
  %(prog)s --pipe 1 --search "bom" --where "Ha Noi"
  %(prog)s --pipe 2                          # Crawl chi tiet tat ca
  %(prog)s --pipe 2 --limit 100             # Chi tiet 100 CT dau
  %(prog)s --pipe 3                          # Download anh
  %(prog)s --pipe all                        # Chay tat ca P0-P3
  %(prog)s --export csv                      # Xuat CSV
  %(prog)s --export excel                    # Xuat Excel
  %(prog)s --export json                     # Xuat JSON
  %(prog)s --stats                           # Xem thong ke
  %(prog)s --autocomplete "vong"             # Test autocomplete
  %(prog)s --provinces "ha"                  # Liet ke tinh
        """,
    )

    # Pipeline
    p.add_argument("--pipe", type=str, choices=["0", "1", "2", "3", "all"],
                   help="Pipeline de chay (0/1/2/3/all)")

    # Filters cho Pipe 1
    p.add_argument("--category", type=str,
                   help="Category ID(s), cach nhau bang dau phay. VD: 26360,15020")
    p.add_argument("--group", type=int,
                   help="Group ID (nhom nganh). VD: 41")
    p.add_argument("--letter", type=str,
                   help="Chu cai dau cua danh muc. VD: A")
    p.add_argument("--search", type=str,
                   help="Tim kiem keyword")
    p.add_argument("--where", type=str, default="",
                   help="Tim kiem theo dia diem (dung voi --search)")

    # Gioi han
    p.add_argument("--max-pages", type=int,
                   help="So trang toi da moi danh muc")
    p.add_argument("--limit", type=int,
                   help="Gioi han tong so cong ty")

    # Hieu nang
    p.add_argument("--concurrency", type=int, default=15,
                   help="So luong request dong thoi (default: 15)")

    # Export
    p.add_argument("--export", type=str, choices=["csv", "excel", "json"],
                   help="Xuat du lieu")

    # Tien ich
    p.add_argument("--stats", action="store_true",
                   help="Hien thi thong ke")
    p.add_argument("--reset", action="store_true",
                   help="Reset trang thai crawl")
    p.add_argument("--autocomplete", type=str, metavar="QUERY",
                   help="Test autocomplete API")
    p.add_argument("--provinces", type=str, metavar="QUERY",
                   help="Tra cuu tinh thanh")

    # Database
    p.add_argument("--db", type=str, default="trangvang.db",
                   help="Duong dan database (default: trangvang.db)")
    p.add_argument("--output", type=str, default="output",
                   help="Thu muc output (default: output)")

    return p


async def async_main(args):
    config = CrawlerConfig(
        concurrency=args.concurrency,
        output_dir=args.output,
        db_name=args.db,
    )
    db = DatabaseManager(args.db)

    # Stats
    if args.stats:
        show_stats(db)
        return

    # Reset
    if args.reset:
        pipe_filter = args.pipe if args.pipe and args.pipe != "all" else None
        db.reset_crawl_status(pipe_filter)
        logger.info("Da reset trang thai crawl.")
        return

    # Export
    if args.export:
        logger.info(f"Xuat du lieu -> {args.output}/")
        if args.export in ("csv", "all"):
            db.export_csv(args.output)
        if args.export in ("excel", "all"):
            db.export_excel(args.output)
        return

    # Autocomplete test
    if args.autocomplete:
        async with TrangVangCrawler(config, db) as crawler:
            results = await crawler.autocomplete(args.autocomplete)
            print(f"Ket qua autocomplete '{args.autocomplete}':")
            for r in results:
                print(f"  - {r}")
        return

    # Provinces test
    if args.provinces:
        async with TrangVangCrawler(config, db) as crawler:
            results = await crawler.list_provinces(args.provinces)
            print(f"Tinh thanh chua '{args.provinces}':")
            for r in results:
                print(f"  - {r}")
        return

    # Pipeline execution
    if not args.pipe:
        print("Vui long chi dinh --pipe (0/1/2/3/all) hoac --help")
        return

    async with TrangVangCrawler(config, db) as crawler:
        pipes = list("0123") if args.pipe == "all" else [args.pipe]

        # Parse category IDs
        category_ids = None
        if args.category:
            category_ids = [int(x.strip()) for x in args.category.split(",")]

        for pipe in pipes:
            if pipe == "0":
                await crawler.pipe0_collect_categories()
                await crawler.pipe0_collect_all_groups()

            elif pipe == "1":
                if args.search:
                    await crawler.pipe1_crawl_search(
                        args.search, args.where,
                        max_pages=args.max_pages, limit=args.limit,
                    )
                else:
                    await crawler.pipe1_crawl_listings(
                        category_ids=category_ids,
                        group_id=args.group,
                        letter=args.letter,
                        max_pages=args.max_pages,
                        limit=args.limit,
                    )

            elif pipe == "2":
                await crawler.pipe2_crawl_details(limit=args.limit)

            elif pipe == "3":
                await crawler.pipe3_download_assets()

    # Show stats at end
    show_stats(db)


def _do_listing_helper(crawler, url, source, cat_id=None):
    """Helper: fetch ALL pages -> tra ve danh sach company dicts."""
    async def _inner():
        html = await crawler._fetch(url)
        if not html:
            print("  Khong ket noi duoc!")
            return False, [], ""
        total_pages = crawler._get_total_pages(html)
        total_count = crawler._get_total_count(html)

        # Lay ten nganh tu page title (vd: "Ốc, Vít, Bù Loong")
        cat_name = ""
        tree = HTMLParser(html)
        h1 = tree.css_first("h1")
        if h1:
            cat_name = _text(h1).strip()
            # Bo phan " - Trang Vang Viet Nam" neu co
            cat_name = re.sub(r"\s*[-–]\s*Trang\s+V[àa]ng.*$", "", cat_name, flags=re.IGNORECASE)

        print(f"  {total_count} cong ty, {total_pages} trang")
        pb = ProgressBar(total_pages, prefix="Trang")
        companies = await crawler._crawl_listing_url(
            url, source=source, cat_id=cat_id,
            pb=pb, prefetched_html=html,
        )
        pb.finish()
        dup = total_count - len(companies) if total_count > len(companies) else 0
        msg = f"  Quet xong: {len(companies)} cong ty"
        if dup > 0:
            msg += f" (bo {dup} trung lap sponsor)"
        print(msg)
        return True, companies, cat_name
    return _inner()


def _auto_export_on_exit(db: DatabaseManager, output_dir: str):
    """Export data khi dung chuong trinh giua chung (Ctrl+C)"""
    stats = db.get_stats()
    if stats["companies_total"] == 0:
        return
    print("\n\n  Dang dung... Tu dong luu du lieu da crawl duoc!")
    try:
        db.export_excel(output_dir)
        db.export_csv(output_dir)
        if stats["companies_total"] > 50:
            db.export_csv_by_category(output_dir)
    except Exception as e:
        print(f"  Loi khi xuat: {e}")
    print(f"  Da luu {stats['companies_total']} CT, {stats['details_total']} chi tiet")
    print(f"  File: {os.path.abspath(output_dir)}/")
    print("  (Chay lai de tiep tuc crawl phan con lai)\n")


def interactive_menu():
    """Interactive menu"""
    print()
    print("╔══════════════════════════════════════════════════╗")
    print("║    TRANG VANG VIETNAM - Tim kiem doanh nghiep    ║")
    print("╚══════════════════════════════════════════════════╝")
    print()

    db_path = "trangvang.db"
    output_dir = "output"
    db = DatabaseManager(db_path)
    config = CrawlerConfig(output_dir=output_dir, db_name=db_path)

    # Ctrl+C handler: auto export before exit
    import signal
    def _signal_handler(sig, frame):
        _auto_export_on_exit(db, output_dir)
        sys.exit(0)
    signal.signal(signal.SIGINT, _signal_handler)

    while True:
        stats = db.get_stats()
        if stats["companies_total"] > 0:
            print(f"  Da luu: {stats['companies_total']} cong ty | "
                  f"{stats['details_total']} chi tiet | "
                  f"{stats['provinces']} tinh/TP")
            print()

        print("Ban muon lam gi?")
        print()
        print("  1.  Crawl ALL       (tat ca nganh nghe, xuat CSV theo tung nganh)")
        print("  2.  Tim kiem        (nhap tu khoa hoac dan link)")
        print("  3.  Xuat file       (Excel / CSV)")
        print("  4.  Xem thong ke")
        print("  5.  Xoa du lieu cu")
        print("  0.  Thoat")
        print()

        choice = input("  Nhap so (0-5): ").strip()

        if choice == "0":
            print("Tam biet!")
            break

        # =========================================
        # 1. CRAWL ALL - Tat ca nganh, chia CSV theo nganh
        # =========================================
        elif choice == "1":
            print()
            print(">> CRAWL ALL - Lay toan bo danh ba trang vang")
            print()
            print("  Se crawl theo nhom nganh. Chon nhom:")
            print()

            # Show groups
            gid_list = sorted(GROUP_IDS.items(), key=lambda x: x[0])
            for gid, slug in gid_list:
                label = slug.replace("-", " ").title()
                print(f"    {gid:>3}. {label}")
            print()
            print("  Nhap so nhom (VD: 41) hoac 'all' de lay tat ca")
            print("  Nhap nhieu nhom: 10,38,41")
            print()

            group_input = input("  >> ").strip().lower()
            if not group_input:
                print("  Chua nhap gi!")
                continue

            if group_input == "all":
                selected_groups = list(GROUP_IDS.keys())
            else:
                try:
                    selected_groups = [int(x.strip()) for x in group_input.split(",") if x.strip()]
                    # Validate
                    for gid in selected_groups:
                        if gid not in GROUP_IDS:
                            print(f"  Nhom {gid} khong ton tai!")
                            selected_groups = []
                            break
                except ValueError:
                    print("  Nhap so khong hop le!")
                    continue

            if not selected_groups:
                continue

            group_names = [GROUP_IDS[g].replace("-", " ").title() for g in selected_groups]
            print(f"\n  Bat dau crawl {len(selected_groups)} nhom: {', '.join(group_names)}")
            t0 = time.monotonic()

            async def _run_all():
                async with TrangVangCrawler(config, db) as crawler:
                    print("\n  [1/3] Thu thap danh muc...")
                    for gid in selected_groups:
                        await crawler.pipe0_collect_from_group(gid)

                    cats = db.get_categories(uncrawled_only=True)
                    if not cats:
                        print("  Khong co danh muc moi nao!")
                        return

                    print(f"\n  [2/3] Crawl {len(cats)} danh muc...")
                    await crawler.pipe1_crawl_listings()

                    print(f"\n  [3/3] Lay chi tiet cong ty...")
                    await crawler.pipe2_crawl_details()

            try:
                asyncio.run(_run_all())
            except KeyboardInterrupt:
                _auto_export_on_exit(db, output_dir)
                continue

            # Export: Excel tong + CSV chia theo nganh
            print()
            db.export_excel(output_dir)
            db.export_csv(output_dir)
            db.export_csv_by_category(output_dir)

            elapsed = time.monotonic() - t0
            new_stats = db.get_stats()
            print()
            print(f"  HOAN TAT! ({elapsed:.0f}s)")
            print(f"  - {new_stats['companies_total']} cong ty")
            print(f"  - {new_stats['details_total']} chi tiet")
            print(f"  - File tong: {os.path.abspath(output_dir)}/trangvang_data.xlsx")
            print(f"  - CSV theo nganh: {os.path.abspath(output_dir)}/theo_nganh/")

        # =========================================
        # 2. Tim kiem = Listing + Detail + Export
        # =========================================
        elif choice == "2":
            print()
            print(">> Tim kiem cong ty")
            print("  Dan link hoac nhap tu khoa. VD:")
            print("    vong bi")
            print("    https://trangvangvietnam.com/categories/26360/vong-bi-bac-dan.html")
            if stats["companies_total"] > 0:
                print()
                print(f"  (DB hien co {stats['companies_total']} CT. "
                      f"Nhap 'xoa' de lam moi truoc khi crawl)")
            print()

            query = input("  >> ").strip()
            if not query:
                print("  Chua nhap gi!")
                continue

            # Allow quick reset
            if query.lower() == "xoa":
                with db.session() as s:
                    s.query(CompanyDetail).delete()
                    s.query(Company).delete()
                    s.commit()
                print("  Da xoa! Nhap tu khoa hoac link:")
                query = input("  >> ").strip()
                if not query:
                    continue

            is_url = "trangvangvietnam.com" in query
            t0 = time.monotonic()

            search_cat_name = ""  # ten nganh de dat ten file

            async def _run_search():
                nonlocal search_cat_name
                async with TrangVangCrawler(config, db) as crawler:
                    print()
                    print("  [1/3] Lay danh sach...")

                    found_companies = []
                    cat_name = ""
                    if is_url:
                        cat_id = _extract_id_from_url(query, "categories")
                        tag_id = _extract_id_from_url(query, "tagclass")
                        clean_url = query.split("?")[0]

                        if cat_id:
                            ok, comps, cat_name = await _do_listing_helper(crawler, clean_url, "category", cat_id)
                        elif tag_id:
                            ok, comps, cat_name = await _do_listing_helper(crawler, clean_url, "tag")
                        elif "/search.asp" in query:
                            ok, comps, cat_name = await _do_listing_helper(crawler, query, "search")
                        elif "/cateprovinces/" in query:
                            ok, comps, cat_name = await _do_listing_helper(crawler, clean_url, "province")
                        else:
                            print("  Link khong ho tro!")
                            return
                        if not ok:
                            return
                        found_companies = comps
                    else:
                        print(f"  Dang tim danh muc cho '{query}'...")
                        result = await crawler.resolve_keyword_to_category(query)
                        if not result:
                            print("  Khong tim thay ket qua!")
                            return
                        resolved_url, cat_id = result
                        if cat_id:
                            print(f"  -> Tim thay danh muc: {resolved_url}")
                            ok, comps, cat_name = await _do_listing_helper(crawler, resolved_url, "category", cat_id)
                        else:
                            print("  -> Dung che do tim kiem (co the co trung lap)")
                            ok, comps, cat_name = await _do_listing_helper(crawler, resolved_url, "search")
                        if not ok:
                            return
                        found_companies = comps

                    search_cat_name = cat_name or query

                    # Vao TUNG LINK -> lay chi tiet
                    print()
                    print(f"  [2/3] Vao tung link lay chi tiet...")
                    await crawler.crawl_details_direct(found_companies)

            try:
                asyncio.run(_run_search())
            except KeyboardInterrupt:
                _auto_export_on_exit(db, output_dir)
                continue

            # Dat ten file theo nganh
            safe_name = _safe_filename(search_cat_name) if search_cat_name else ""

            print()
            print("  [3/3] Xuat file...")
            filepath = db.export_excel(output_dir, filename=safe_name)
            db.export_csv(output_dir, filename=safe_name)

            elapsed = time.monotonic() - t0
            new_stats = db.get_stats()
            print()
            print(f"  HOAN TAT! ({elapsed:.0f}s)")
            print(f"  - {new_stats['companies_total']} cong ty")
            print(f"  - {new_stats['details_total']} chi tiet")
            if filepath:
                print(f"  - File: {os.path.abspath(filepath)}")

        # =========================================
        # 3. Xuat file
        # =========================================
        elif choice == "3":
            if stats["companies_total"] == 0:
                print("  Chua co du lieu. Hay crawl truoc (chon 1 hoac 2).")
                continue

            uncrawled = stats["companies_total"] - stats["companies_detail_crawled"]
            if uncrawled > 0:
                print(f"  Con {uncrawled} CT chua lay chi tiet. Dang lay...")
                async def _run_detail():
                    async with TrangVangCrawler(config, db) as crawler:
                        await crawler.pipe2_crawl_details()
                asyncio.run(_run_detail())

            print()
            print("  1. Excel (.xlsx)")
            print("  2. CSV   (.csv)")
            print("  3. CSV theo nganh (moi nganh 1 file)")
            print("  4. Tat ca")
            fmt = input("  Chon (1-4) [4]: ").strip() or "4"

            if fmt in ("1", "4"):
                db.export_excel(output_dir)
            if fmt in ("2", "4"):
                db.export_csv(output_dir)
            if fmt in ("3", "4"):
                db.export_csv_by_category(output_dir)

            print(f"  Da xuat -> {os.path.abspath(output_dir)}")

        # =========================================
        # 4. Thong ke
        # =========================================
        elif choice == "4":
            print()
            show_stats(db)

        # =========================================
        # 5. Xoa du lieu cu
        # =========================================
        elif choice == "5":
            print()
            if stats["companies_total"] == 0:
                print("  Chua co du lieu gi de xoa.")
                continue
            print(f"  Se xoa toan bo {stats['companies_total']} cong ty "
                  f"va {stats['details_total']} chi tiet")
            confirm = input("  Nhap 'xoa' de xac nhan: ").strip()
            if confirm.lower() == "xoa":
                with db.session() as s:
                    s.query(CompanyDetail).delete()
                    s.query(Company).delete()
                    s.query(Category).delete()
                    s.commit()
                print("  Da xoa sach!")
            else:
                print("  Huy.")

        else:
            print("  Nhap so tu 0-5!")

        print()
        print("──────────────────────────────────────────────────")
        print()


def main():
    parser = build_parser()
    args = parser.parse_args()

    # No args -> interactive menu
    if len(sys.argv) == 1:
        interactive_menu()
        return

    print("Trang Vang Vietnam Crawler v1.0")
    print(f"Database: {args.db} | Output: {args.output} | Concurrency: {args.concurrency}")
    print()

    asyncio.run(async_main(args))


if __name__ == "__main__":
    main()
