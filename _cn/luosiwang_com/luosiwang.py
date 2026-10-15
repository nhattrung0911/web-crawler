# -*- coding: utf-8 -*-
"""
Luosiwang 2-Pipeline Crawler (Optimized)
- Pipeline 1: lấy categories và cho user chọn (Playwright headless)
- Pipeline 2: crawl dữ liệu chi tiết (httpx direct API khi có thể, fallback Playwright)
- Resume-safe / CSV checkpoint / Dedupe / Retry / Progress realtime nhẹ
- Headless + resource blocking + API interception cho tốc độ tối đa

Yêu cầu cài:
pip install playwright aiofiles orjson rich pandas tenacity httpx
playwright install chromium

Usage:
  python lousiwang.py              # interactive mode
  python lousiwang.py --all        # auto crawl ALL categories, no prompts
  python lousiwang.py --fresh      # clear old state and start fresh
  python lousiwang.py --all --fresh
"""

from __future__ import annotations

import asyncio
import csv
import io
import os
import re
import signal
import sys
import time
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Set

import aiofiles
import httpx
import orjson
import pandas as pd
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Prompt
from rich.table import Table
from rich.progress import (
    Progress,
    SpinnerColumn,
    TextColumn,
    BarColumn,
    TimeElapsedColumn,
    TaskProgressColumn,
)
from tenacity import (
    retry,
    stop_after_attempt,
    wait_exponential,
    retry_if_exception_type,
)

from playwright.async_api import (
    async_playwright,
    Browser,
    BrowserContext,
    Page,
    TimeoutError as PlaywrightTimeoutError,
)

# Optional uvloop (Linux/macOS only)
try:
    import uvloop  # type: ignore
    uvloop.install()
except Exception:
    pass

# ── Windows UTF-8 fix ──
if sys.platform == "win32":
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name)
        if hasattr(stream, "buffer"):
            setattr(
                sys,
                stream_name,
                io.TextIOWrapper(stream.buffer, encoding="utf-8", errors="replace"),
            )

console = Console(force_terminal=True)

# ── Parse CLI args ──
AUTO_ALL = "--all" in sys.argv
FRESH_START = "--fresh" in sys.argv

# =========================
# CONFIG
# =========================

BASE_URL = (
    "https://luosiwang.com/pc/products"
    "?sid=394"
    "&categoryName=304%2520Bulong%2520l%25E1%25BB%25A5c%2520gi%25C3%25A1c%2520ngo%25C3%25A0i"
    "&categoryImage=https%253A%252F%252Fluosiwang.com%252Fuploads%252Fattach"
    "%252F2025%252F06%252F20250612%252F80b3ed5042f1b31b9c09f367b20b6d29.jpg"
)

ROOT_DIR = Path(__file__).resolve().parent
STATE_DIR = ROOT_DIR / "state"
OUTPUT_DIR = ROOT_DIR / "output"
LOG_DIR = ROOT_DIR / "logs"

STATE_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

SELECTED_CATEGORIES_FILE = STATE_DIR / "selected_categories.json"
COMPLETED_PAGES_FILE = STATE_DIR / "completed_pages.csv"
FAILED_PAGES_FILE = STATE_DIR / "failed_pages.csv"
DATA_JSONL_FILE = OUTPUT_DIR / "data_rows.jsonl"
DATA_CSV_FILE = OUTPUT_DIR / "data_rows.csv"
FINAL_CSV_FILE = OUTPUT_DIR / "luosiwang_final.csv"
RUN_LOG_FILE = LOG_DIR / "run.log"

HEADLESS = True                    # headless cho tốc độ tối đa
BROWSER_CONCURRENCY = 6            # tăng concurrency vì headless nhẹ hơn
NAV_TIMEOUT_MS = 60_000
ACTION_TIMEOUT_MS = 20_000
RETRY_ATTEMPTS = 4
BATCH_FLUSH_SIZE = 50
POLITE_DELAY = 0.15                # giảm delay vì headless không cần render

# Resource types to block in browser (save bandwidth + speed)
BLOCKED_RESOURCE_TYPES = {"image", "media", "font", "stylesheet"}
BLOCKED_URL_PATTERNS = [
    "*.png", "*.jpg", "*.jpeg", "*.gif", "*.svg", "*.ico",
    "*.woff", "*.woff2", "*.ttf", "*.eot",
    "*.mp4", "*.webm", "*.mp3",
    "*.css",
    "*google-analytics*", "*googletagmanager*", "*facebook*",
    "*hotjar*", "*sentry*",
]

# =========================
# SELECTORS
# =========================

class S:
    # language
    LANGUAGE_SWITCHER = "a.language-switcher"
    CURRENT_LANG = "span.current-lang"

    # category popup
    CATEGORY_TRIGGER = "span.clickable"
    CATEGORY_POPUP = "div.category-selector-popup"
    SECOND_LEVEL_WRAP = "div.category-selector-popup div.second-level-categories"
    THIRD_LEVEL_WRAP = "div.category-selector-popup div.third-level-categories"

    SECOND_LEVEL_ITEM = "div.category-item"
    SECOND_LEVEL_NAME = "div.category-name"
    SECOND_LEVEL_COUNT = "div.category-count"

    THIRD_LEVEL_ITEM = "div.category-list div.category-item"
    THIRD_LEVEL_NAME = "div.category-name"
    THIRD_LEVEL_IMAGE = "div.category-icon img"

    # products table
    PRODUCTS_TABLE = "table"
    TABLE_HEADER = "tr.table-header"
    PRODUCT_ROWS = "table tr:not(.table-header)"

    # product cells
    IMAGE_CELL_IMG = "td.product-image-cell img"
    PRODUCT_NAME_CELL = "td.product-name-cell"
    MATERIAL_BADGE = ".material-badge"
    PRODUCT_NAME = ".product-name"

    SPEC_NO = "td.spec-no-cell .spec-no"
    SPEC_INFO_ITEMS = "td.spec-info-cell .spec-item"
    SPEC_LABEL = ".spec-label"
    SPEC_VALUE = ".spec-value"

    WAREHOUSE_NAME = "td.warehouse-cell .warehouse-name"
    BRAND_NAME = "td.brand-cell .brand-name"

    STOCK_VALUE = "td.stock-cell .stock-value"
    STOCK_UNIT = "td.stock-cell .stock-unit"

    MOQ_VALUE = "td.package-moq-cell .moq-value"
    MOQ_UNIT = "td.package-moq-cell .moq-unit"
    PACKAGE_TEXT = "td.package-moq-cell .package-text"

    PRICE_VALUE = "td.price-cell .price-value"
    PRICE_UNIT = "td.price-cell .price-unit"

    STOCK_STATUS = ".stock-status"

    # pagination
    PAGE_SIZE_SELECT = ".ant-pagination .ant-select-selector"
    PAGE_SIZE_CURRENT = ".ant-pagination .ant-select-selection-item"
    QUICK_JUMPER_INPUT = ".ant-pagination-options-quick-jumper input"
    PAGINATION_ITEMS = "li.ant-pagination-item"
    PAGINATION_ACTIVE = "li.ant-pagination-item-active"


# =========================
# DATA MODELS
# =========================

@dataclass
class LargeCategory:
    large_name: str
    count_text: str

@dataclass
class SmallCategory:
    large_name: str
    small_name: str
    image_url: str = ""

@dataclass
class ProductRow:
    cate_lon: str
    cate_nho: str
    ten_san_pham: str
    size: str
    detail: str
    kho: str
    brand: str
    stock: str
    minimum_orders: str
    gia: str
    tinh_trang: str
    link_anh: str

    def dedupe_key(self) -> str:
        return "||".join([
            self.cate_lon.strip(),
            self.cate_nho.strip(),
            self.ten_san_pham.strip(),
            self.size.strip(),
            self.kho.strip(),
            self.brand.strip(),
        ])

PRODUCT_FIELDS = list(ProductRow.__dataclass_fields__.keys())


# =========================
# GLOBAL RUNTIME (locks created lazily inside event loop)
# =========================

@dataclass
class CapturedApi:
    """Stores an intercepted API endpoint for direct httpx calls."""
    url: str
    method: str = "GET"
    headers: Dict[str, str] = None  # type: ignore
    cookies: Dict[str, str] = None  # type: ignore

    def __post_init__(self):
        self.headers = self.headers or {}
        self.cookies = self.cookies or {}


class RuntimeState:
    def __init__(self) -> None:
        self.stop_requested = False
        self.rows_buffer: List[ProductRow] = []
        self.seen_keys: Set[str] = set()
        self.total_saved = 0
        # API interception cache: category_key -> CapturedApi
        self.captured_apis: Dict[str, CapturedApi] = {}
        # Shared httpx client (created lazily)
        self._httpx_client: Optional[httpx.AsyncClient] = None
        # Locks — created on first use inside the event loop
        self._buffer_lock: Optional[asyncio.Lock] = None
        self._completed_lock: Optional[asyncio.Lock] = None
        self._failed_lock: Optional[asyncio.Lock] = None
        self._log_lock: Optional[asyncio.Lock] = None

    async def get_httpx_client(self) -> httpx.AsyncClient:
        if self._httpx_client is None or self._httpx_client.is_closed:
            self._httpx_client = httpx.AsyncClient(
                timeout=httpx.Timeout(30.0, connect=10.0),
                follow_redirects=True,
                http2=True,
                limits=httpx.Limits(
                    max_connections=20,
                    max_keepalive_connections=10,
                ),
            )
        return self._httpx_client

    async def close_httpx(self) -> None:
        if self._httpx_client and not self._httpx_client.is_closed:
            await self._httpx_client.aclose()

    @property
    def buffer_lock(self) -> asyncio.Lock:
        if self._buffer_lock is None:
            self._buffer_lock = asyncio.Lock()
        return self._buffer_lock

    @property
    def completed_pages_lock(self) -> asyncio.Lock:
        if self._completed_lock is None:
            self._completed_lock = asyncio.Lock()
        return self._completed_lock

    @property
    def failed_pages_lock(self) -> asyncio.Lock:
        if self._failed_lock is None:
            self._failed_lock = asyncio.Lock()
        return self._failed_lock

    @property
    def log_lock(self) -> asyncio.Lock:
        if self._log_lock is None:
            self._log_lock = asyncio.Lock()
        return self._log_lock


RUNTIME = RuntimeState()


# =========================
# UTILS
# =========================

def now_str() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


async def write_log(message: str) -> None:
    line = f"[{now_str()}] {message}\n"
    async with RUNTIME.log_lock:
        async with aiofiles.open(RUN_LOG_FILE, "a", encoding="utf-8") as f:
            await f.write(line)


def clean_text(text: Optional[str]) -> str:
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip()


def clean_multiline_text(texts: List[str]) -> str:
    lines = [clean_text(x) for x in texts if clean_text(x)]
    return "\n".join(lines)


def parse_int_from_text(s: str) -> int:
    nums = re.findall(r"\d+", s or "")
    return int(nums[0]) if nums else 0


def chunked(seq: List[Any], size: int) -> List[List[Any]]:
    return [seq[i : i + size] for i in range(0, len(seq), size)]


def is_interactive() -> bool:
    """Check if stdin is interactive (not piped)."""
    try:
        return sys.stdin.isatty()
    except Exception:
        return False


def prompt_or_default(message: str, default: str) -> str:
    """Prompt user if interactive, otherwise return default."""
    if AUTO_ALL or not is_interactive():
        console.print(f"{message}: [dim](auto: {default})[/dim]")
        return default
    try:
        return Prompt.ask(message, default=default)
    except EOFError:
        return default


# ── CSV helpers using proper csv.writer ──

async def ensure_csv_header(path: Path, fieldnames: List[str]) -> None:
    if path.exists():
        return
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(fieldnames)
    async with aiofiles.open(path, "w", encoding="utf-8-sig", newline="") as f:
        await f.write(buf.getvalue())


async def append_csv_rows(
    path: Path, rows: List[Dict[str, Any]], fieldnames: List[str]
) -> None:
    if not rows:
        return
    await ensure_csv_header(path, fieldnames)
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=fieldnames, extrasaction="ignore")
    for row in rows:
        writer.writerow(row)
    async with aiofiles.open(path, "a", encoding="utf-8-sig", newline="") as f:
        await f.write(buf.getvalue())


async def load_completed_pages() -> Set[str]:
    completed: Set[str] = set()
    if not COMPLETED_PAGES_FILE.exists():
        return completed
    async with aiofiles.open(
        COMPLETED_PAGES_FILE, "r", encoding="utf-8-sig", newline=""
    ) as f:
        content = await f.read()
    reader = csv.DictReader(io.StringIO(content))
    for row in reader:
        completed.add(row["page_key"])
    return completed


async def append_completed_page(
    page_key: str, large_name: str, small_name: str, page_no: int
) -> None:
    fields = ["page_key", "large_name", "small_name", "page_no", "completed_at"]
    async with RUNTIME.completed_pages_lock:
        await append_csv_rows(
            COMPLETED_PAGES_FILE,
            [
                {
                    "page_key": page_key,
                    "large_name": large_name,
                    "small_name": small_name,
                    "page_no": page_no,
                    "completed_at": now_str(),
                }
            ],
            fields,
        )


async def append_failed_page(
    page_key: str, large_name: str, small_name: str, page_no: int, error: str
) -> None:
    fields = [
        "page_key", "large_name", "small_name", "page_no", "error", "failed_at"
    ]
    async with RUNTIME.failed_pages_lock:
        await append_csv_rows(
            FAILED_PAGES_FILE,
            [
                {
                    "page_key": page_key,
                    "large_name": large_name,
                    "small_name": small_name,
                    "page_no": page_no,
                    "error": str(error)[:200],
                    "failed_at": now_str(),
                }
            ],
            fields,
        )


async def flush_rows(force: bool = False) -> None:
    async with RUNTIME.buffer_lock:
        if not RUNTIME.rows_buffer:
            return
        if not force and len(RUNTIME.rows_buffer) < BATCH_FLUSH_SIZE:
            return

        unique_new: List[ProductRow] = []
        for row in RUNTIME.rows_buffer:
            key = row.dedupe_key()
            if key not in RUNTIME.seen_keys:
                RUNTIME.seen_keys.add(key)
                unique_new.append(row)

        if unique_new:
            # JSONL
            async with aiofiles.open(DATA_JSONL_FILE, "a", encoding="utf-8") as f:
                for row in unique_new:
                    await f.write(
                        orjson.dumps(asdict(row)).decode("utf-8") + "\n"
                    )
            # CSV
            await append_csv_rows(
                DATA_CSV_FILE,
                [asdict(r) for r in unique_new],
                PRODUCT_FIELDS,
            )
            RUNTIME.total_saved += len(unique_new)

        RUNTIME.rows_buffer.clear()


async def save_row(row: ProductRow) -> None:
    async with RUNTIME.buffer_lock:
        RUNTIME.rows_buffer.append(row)
    await flush_rows(force=False)


def install_signal_handlers() -> None:
    def _handler(sig, frame):
        RUNTIME.stop_requested = True
        console.print(
            "\n[yellow]Signal received. Flushing data and saving checkpoint...[/yellow]"
        )

    signal.signal(signal.SIGINT, _handler)
    if hasattr(signal, "SIGTERM"):
        signal.signal(signal.SIGTERM, _handler)


def clear_state_files() -> None:
    """Delete old state/output files for a fresh start."""
    for f in [
        SELECTED_CATEGORIES_FILE,
        COMPLETED_PAGES_FILE,
        FAILED_PAGES_FILE,
        DATA_JSONL_FILE,
        DATA_CSV_FILE,
        FINAL_CSV_FILE,
    ]:
        if f.exists():
            f.unlink()
            console.print(f"  [dim]Deleted {f.name}[/dim]")


# =========================
# PLAYWRIGHT HELPERS
# =========================

async def safe_text(locator) -> str:
    try:
        return clean_text(await locator.inner_text())
    except Exception:
        return ""


async def safe_attr(locator, attr: str) -> str:
    try:
        v = await locator.get_attribute(attr)
        return clean_text(v)
    except Exception:
        return ""


@retry(
    stop=stop_after_attempt(RETRY_ATTEMPTS),
    wait=wait_exponential(multiplier=2, min=2, max=15),
    retry=retry_if_exception_type((PlaywrightTimeoutError, Exception)),
    reraise=True,
)
async def goto_base(page: Page) -> None:
    await page.goto(BASE_URL, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
    # Wait for Vue app to render — wait for the category trigger to appear
    try:
        await page.wait_for_selector(
            S.CATEGORY_TRIGGER, state="visible", timeout=15_000
        )
    except Exception:
        pass
    await page.wait_for_timeout(500)


async def create_context(browser: Browser, block_resources: bool = True) -> BrowserContext:
    context = await browser.new_context(
        viewport={"width": 1600, "height": 1000},
        locale="vi-VN",
    )
    context.set_default_navigation_timeout(NAV_TIMEOUT_MS)
    context.set_default_timeout(ACTION_TIMEOUT_MS)

    if block_resources:
        # Block images, fonts, CSS, media, trackers to speed up loading
        async def _block_route(route):
            if route.request.resource_type in BLOCKED_RESOURCE_TYPES:
                await route.abort()
            else:
                await route.continue_()

        await context.route("**/*", _block_route)

    return context


async def try_switch_language(page: Page, target_lang: str = "Tiếng Việt") -> None:
    """Try to switch language. Non-blocking — failures are silently ignored.
    We use Vietnamese by default since the site content is in Vietnamese."""
    try:
        lang_el = page.locator(S.CURRENT_LANG).first
        current = clean_text(await lang_el.inner_text())
        if target_lang.lower() in current.lower():
            return
    except Exception:
        return  # No language switcher found, skip

    try:
        switcher = page.locator(S.LANGUAGE_SWITCHER).first
        await switcher.hover(timeout=5000)
        await page.wait_for_timeout(200)
        # Try clicking the target language option
        lang_option = page.get_by_text(target_lang, exact=False).first
        await lang_option.click(timeout=5000)
        await page.wait_for_timeout(500)
        await write_log(f"Switched language to {target_lang}")
    except Exception:
        pass  # Language switch is not critical


async def close_popup_if_open(page: Page) -> None:
    """Close category popup if visible."""
    for _ in range(3):
        try:
            popup = page.locator(S.CATEGORY_POPUP).first
            if not await popup.is_visible():
                return
            await page.keyboard.press("Escape")
            await page.wait_for_timeout(200)
            if not await popup.is_visible():
                return
            await page.mouse.click(10, 10)
            await page.wait_for_timeout(200)
        except Exception:
            return


async def open_category_popup(page: Page) -> None:
    """Open category popup (close first if already open)."""
    await close_popup_if_open(page)
    await page.wait_for_timeout(100)

    for attempt in range(3):
        try:
            trigger = page.locator(S.CATEGORY_TRIGGER).first
            await trigger.click(force=True)
            await page.wait_for_selector(
                S.CATEGORY_POPUP, state="visible", timeout=10_000
            )
            await page.wait_for_timeout(400)
            return
        except Exception as e:
            if attempt == 2:
                raise RuntimeError(f"Cannot open category popup: {e}")
            await page.wait_for_timeout(500)


async def wait_products_loaded(page: Page, min_rows: int = 1) -> None:
    """Wait for product table to appear and actual product rows to render."""
    # Wait for table header first
    await page.wait_for_selector(
        S.TABLE_HEADER, state="visible", timeout=NAV_TIMEOUT_MS
    )
    # Then wait for at least min_rows product rows to appear
    for _ in range(30):  # max 15 seconds
        rows = page.locator(S.PRODUCT_ROWS)
        count = await rows.count()
        if count >= min_rows:
            break
        await page.wait_for_timeout(500)
    # Extra wait for all rows to finish rendering
    await page.wait_for_timeout(300)


async def scroll_to_load_all(page: Page) -> None:
    """Scroll down the page to trigger lazy loading of all product rows."""
    for _ in range(3):
        await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        await page.wait_for_timeout(150)
    # Scroll back to top
    await page.evaluate("window.scrollTo(0, 0)")
    await page.wait_for_timeout(100)


async def set_page_size_200(page: Page) -> None:
    """Change pagination to 200/page."""
    try:
        current = page.locator(S.PAGE_SIZE_CURRENT).first
        txt = clean_text(await current.inner_text())
        if "200" in txt:
            return
    except Exception:
        pass

    for attempt in range(3):
        try:
            # Click the pagination page-size selector
            selector = page.locator(S.PAGE_SIZE_SELECT).first
            await selector.scroll_into_view_if_needed()
            await selector.click()
            await page.wait_for_timeout(400)

            # Find and click "200 / page" or "200 / trang" in the dropdown
            option = page.locator(
                "div.ant-select-item-option-content"
            ).filter(has_text="200").first
            await option.click()
            await page.wait_for_timeout(800)

            # Wait for the table to reload with new page size
            await wait_products_loaded(page, min_rows=1)
            await write_log("Page size set to 200/page")
            return
        except Exception as e:
            if attempt == 2:
                await write_log(f"[WARN] Cannot set page size 200: {e}")
            await page.wait_for_timeout(500)


async def get_total_pages(page: Page) -> int:
    """Get the last page number from pagination."""
    try:
        items = page.locator(S.PAGINATION_ITEMS)
        count = await items.count()
        if count == 0:
            return 1
        last_num = 1
        for i in range(count):
            title = await items.nth(i).get_attribute("title")
            if title and title.isdigit():
                last_num = max(last_num, int(title))
        return last_num
    except Exception:
        return 1


async def get_current_page_number(page: Page) -> int:
    """Get the currently active page number."""
    try:
        active = page.locator(S.PAGINATION_ACTIVE).first
        title = await active.get_attribute("title")
        if title and title.isdigit():
            return int(title)
    except Exception:
        pass
    return 1


async def goto_page_number(page: Page, page_no: int) -> None:
    """Navigate to a specific page using the quick jumper input.
    Verifies the page actually changed by checking the active pagination item."""
    if page_no <= 1:
        return

    # Get current page to know when it changes
    current = await get_current_page_number(page)
    if current == page_no:
        return

    # Get the first row's spec-no text to detect content change
    old_first_spec = ""
    try:
        old_first_spec = await safe_text(
            page.locator(S.PRODUCT_ROWS).first.locator(S.SPEC_NO)
        )
    except Exception:
        pass

    for attempt in range(3):
        try:
            inp = page.locator(S.QUICK_JUMPER_INPUT).first
            await inp.scroll_into_view_if_needed()
            await inp.click()
            # Triple-click to select all text, then type new value
            await inp.click(click_count=3)
            await page.wait_for_timeout(100)
            await inp.type(str(page_no), delay=30)
            await inp.press("Enter")

            # Wait for page content to actually change
            for _ in range(20):  # max ~6 seconds
                await page.wait_for_timeout(300)
                new_page = await get_current_page_number(page)
                if new_page == page_no:
                    break

            # Verify we landed on the right page
            final_page = await get_current_page_number(page)
            if final_page != page_no:
                await write_log(
                    f"[WARN] Wanted page {page_no}, got page {final_page} (attempt {attempt+1})"
                )
                if attempt < 2:
                    continue
            else:
                break

        except Exception as e:
            if attempt == 2:
                await write_log(f"[WARN] goto_page_number({page_no}) failed: {e}")
                raise
            await page.wait_for_timeout(1000)

    # Wait for new rows to render
    await wait_products_loaded(page, min_rows=1)
    await page.wait_for_timeout(200)


# =========================
# PIPELINE 1 — CATEGORY SELECTION
# =========================

async def extract_large_categories(page: Page) -> List[LargeCategory]:
    await open_category_popup(page)
    wrap = page.locator(S.SECOND_LEVEL_WRAP).first
    await wrap.wait_for(state="visible", timeout=ACTION_TIMEOUT_MS)

    items = wrap.locator(S.SECOND_LEVEL_ITEM)
    count = await items.count()

    result: List[LargeCategory] = []
    for i in range(count):
        item = items.nth(i)
        name = clean_text(await item.locator(S.SECOND_LEVEL_NAME).inner_text())
        cnt = clean_text(await item.locator(S.SECOND_LEVEL_COUNT).inner_text())
        if name:
            result.append(LargeCategory(large_name=name, count_text=cnt))

    return result


async def click_large_and_get_smalls(
    page: Page, large_name: str
) -> List[SmallCategory]:
    """Open popup, click a large category, extract its small categories."""
    await open_category_popup(page)

    # Click the large category
    wrap = page.locator(S.SECOND_LEVEL_WRAP).first
    items = wrap.locator(S.SECOND_LEVEL_ITEM)
    count = await items.count()

    found = False
    for i in range(count):
        item = items.nth(i)
        item_name = clean_text(await item.locator(S.SECOND_LEVEL_NAME).inner_text())
        if item_name == large_name:
            await item.click()
            await page.wait_for_timeout(600)
            found = True
            break

    if not found:
        raise ValueError(f"Large category not found: {large_name}")

    # Extract small categories
    third_wrap = page.locator(S.THIRD_LEVEL_WRAP).first
    await third_wrap.wait_for(state="visible", timeout=ACTION_TIMEOUT_MS)
    third_items = third_wrap.locator(S.THIRD_LEVEL_ITEM)
    third_count = await third_items.count()

    result: List[SmallCategory] = []
    for i in range(third_count):
        item = third_items.nth(i)
        small_name = clean_text(await item.locator(S.THIRD_LEVEL_NAME).inner_text())
        image_url = await safe_attr(item.locator(S.THIRD_LEVEL_IMAGE), "src")
        if small_name:
            result.append(
                SmallCategory(
                    large_name=large_name,
                    small_name=small_name,
                    image_url=image_url,
                )
            )

    return result


def render_large_categories(categories: List[LargeCategory]) -> None:
    table = Table(title="Large Categories", show_lines=False)
    table.add_column("#", style="cyan", justify="right")
    table.add_column("Category", style="green")
    table.add_column("Products", style="yellow")
    for idx, cat in enumerate(categories, 1):
        table.add_row(str(idx), cat.large_name, cat.count_text)
    console.print(table)


def render_small_categories(categories: List[SmallCategory], title: str) -> None:
    table = Table(title=title, show_lines=False)
    table.add_column("#", style="cyan", justify="right")
    table.add_column("Sub-category", style="green")
    for idx, cat in enumerate(categories, 1):
        table.add_row(str(idx), cat.small_name)
    console.print(table)


def parse_selection(selection: str, max_n: int) -> List[int]:
    selection = selection.strip().lower()
    if selection in ("all", "a", ""):
        return list(range(1, max_n + 1))

    result: Set[int] = set()
    parts = [p.strip() for p in selection.split(",") if p.strip()]
    for part in parts:
        if "-" in part:
            a, b = part.split("-", 1)
            a_i, b_i = int(a), int(b)
            for i in range(min(a_i, b_i), max(a_i, b_i) + 1):
                if 1 <= i <= max_n:
                    result.add(i)
        else:
            i = int(part)
            if 1 <= i <= max_n:
                result.add(i)
    return sorted(result)


async def pipeline_1_collect_selection(browser: Browser) -> List[SmallCategory]:
    """Pipeline 1: navigate, extract categories, let user choose."""
    context = await create_context(browser)
    page = await context.new_page()

    try:
        await goto_base(page)
        await try_switch_language(page)

        # Extract large categories
        large_categories = await extract_large_categories(page)
        if not large_categories:
            raise RuntimeError("Could not extract any large categories.")

        await close_popup_if_open(page)

        render_large_categories(large_categories)

        if AUTO_ALL:
            user_input = "all"
            console.print("[dim](Auto-selecting ALL large categories)[/dim]")
        else:
            user_input = prompt_or_default(
                "Select large categories [number / 1,3,5 / all]", "all"
            )

        selected_indexes = parse_selection(user_input, len(large_categories))
        selected_large = [large_categories[i - 1] for i in selected_indexes]

        chosen: List[SmallCategory] = []

        if len(selected_large) == 1 and not AUTO_ALL:
            # Single large category — let user pick small categories
            lg = selected_large[0]
            smalls = await click_large_and_get_smalls(page, lg.large_name)
            await close_popup_if_open(page)

            render_small_categories(smalls, f"Sub-categories of: {lg.large_name}")
            small_input = prompt_or_default(
                "Select sub-categories [number / 1,2,5 / all]", "all"
            )
            small_indexes = parse_selection(small_input, len(smalls))
            chosen = [smalls[i - 1] for i in small_indexes]
        else:
            # Multiple or all large categories — auto-select all smalls
            for lg in selected_large:
                try:
                    smalls = await click_large_and_get_smalls(page, lg.large_name)
                    chosen.extend(smalls)
                    await close_popup_if_open(page)
                    console.print(
                        f"  [dim]{lg.large_name}: {len(smalls)} sub-categories[/dim]"
                    )
                except Exception as e:
                    await write_log(
                        f"[WARN] Failed to get smalls for {lg.large_name}: {e}"
                    )
                    console.print(
                        f"  [red]Failed to get sub-categories for {lg.large_name}[/red]"
                    )

        # Save selection to disk
        payload = [asdict(x) for x in chosen]
        async with aiofiles.open(SELECTED_CATEGORIES_FILE, "wb") as f:
            await f.write(orjson.dumps(payload, option=orjson.OPT_INDENT_2))

        await write_log(f"Selected {len(chosen)} sub-categories for crawling")
        return chosen

    finally:
        await context.close()


# =========================
# NAVIGATE TO A SMALL CATEGORY
# =========================

async def navigate_to_category(
    page: Page, large_name: str, small_name: str
) -> None:
    """Go to BASE_URL, open popup, click large → small, set 200/page.
    Also sets up API interception for httpx fast path.
    Retries the full navigation on failure."""
    category_key = f"{large_name}||{small_name}"

    # Attach API interception listener (only once)
    if category_key not in RUNTIME.captured_apis:
        await setup_api_interception(page, category_key)

    for attempt in range(RETRY_ATTEMPTS):
        try:
            await goto_base(page)
            await try_switch_language(page)
            await open_category_popup(page)

            # Click large category
            wrap = page.locator(S.SECOND_LEVEL_WRAP).first
            await wrap.wait_for(state="visible", timeout=ACTION_TIMEOUT_MS)
            items = wrap.locator(S.SECOND_LEVEL_ITEM)
            count = await items.count()

            found_large = False
            for i in range(count):
                item = items.nth(i)
                name = clean_text(
                    await item.locator(S.SECOND_LEVEL_NAME).inner_text()
                )
                if name == large_name:
                    await item.click()
                    await page.wait_for_timeout(600)
                    found_large = True
                    break
            if not found_large:
                raise RuntimeError(f"Large category not found: {large_name}")

            # Click small category
            third_wrap = page.locator(S.THIRD_LEVEL_WRAP).first
            await third_wrap.wait_for(state="visible", timeout=ACTION_TIMEOUT_MS)
            third_items = third_wrap.locator(S.THIRD_LEVEL_ITEM)
            third_count = await third_items.count()

            found_small = False
            for i in range(third_count):
                item = third_items.nth(i)
                name = clean_text(
                    await item.locator(S.THIRD_LEVEL_NAME).inner_text()
                )
                if name == small_name:
                    await item.click()
                    await page.wait_for_timeout(1000)
                    found_small = True
                    break
            if not found_small:
                raise RuntimeError(f"Small category not found: {small_name}")

            # Close popup if still visible
            await close_popup_if_open(page)

            # Wait for products table to fully load
            await wait_products_loaded(page, min_rows=1)

            # Set 200/page for maximum data per page
            await set_page_size_200(page)

            # Scroll to load any lazy content
            await scroll_to_load_all(page)

            await write_log(
                f"[OK] Navigated to {large_name} > {small_name}"
            )
            return

        except Exception as e:
            await write_log(
                f"[WARN] navigate_to_category attempt {attempt+1}/{RETRY_ATTEMPTS} "
                f"failed [{large_name} > {small_name}]: {e}"
            )
            if attempt == RETRY_ATTEMPTS - 1:
                raise
            await page.wait_for_timeout(1500)


# =========================
# PIPELINE 2 — PRODUCT PARSING
# =========================

async def extract_spec_info(row) -> str:
    items = row.locator(S.SPEC_INFO_ITEMS)
    count = await items.count()
    lines = []
    for i in range(count):
        item = items.nth(i)
        label = await safe_text(item.locator(S.SPEC_LABEL))
        value = await safe_text(item.locator(S.SPEC_VALUE))
        line = clean_text(f"{label} {value}")
        if line:
            lines.append(line)
    return clean_multiline_text(lines)


async def parse_product_row(
    row, large_name: str, small_name: str
) -> ProductRow:
    # Fetch all fields concurrently with asyncio.gather for speed
    (
        material, product_name, size, detail,
        kho, brand,
        stock_val, stock_unit,
        moq_val, moq_unit, pkg,
        price_val, price_unit,
        tinh_trang, link_anh,
    ) = await asyncio.gather(
        safe_text(row.locator(f"{S.PRODUCT_NAME_CELL} {S.MATERIAL_BADGE}")),
        safe_text(row.locator(f"{S.PRODUCT_NAME_CELL} {S.PRODUCT_NAME}")),
        safe_text(row.locator(S.SPEC_NO)),
        extract_spec_info(row),
        safe_text(row.locator(S.WAREHOUSE_NAME)),
        safe_text(row.locator(S.BRAND_NAME)),
        safe_text(row.locator(S.STOCK_VALUE)),
        safe_text(row.locator(S.STOCK_UNIT)),
        safe_text(row.locator(S.MOQ_VALUE)),
        safe_text(row.locator(S.MOQ_UNIT)),
        safe_text(row.locator(S.PACKAGE_TEXT)),
        safe_text(row.locator(S.PRICE_VALUE)),
        safe_text(row.locator(S.PRICE_UNIT)),
        safe_text(row.locator(S.STOCK_STATUS)),
        safe_attr(row.locator(S.IMAGE_CELL_IMG), "src"),
    )

    full_name = (
        clean_text(f"{material} {product_name}") if material else product_name
    )
    stock = clean_text(f"{stock_val} {stock_unit}")
    minimum_orders = clean_multiline_text(
        [
            f"MOQ: {moq_val} {moq_unit}" if moq_val else "",
            f"Package: {pkg}" if pkg else "",
        ]
    )
    gia = clean_text(f"{price_val} {price_unit}")

    return ProductRow(
        cate_lon=large_name,
        cate_nho=small_name,
        ten_san_pham=full_name,
        size=size,
        detail=detail,
        kho=kho,
        brand=brand,
        stock=stock,
        minimum_orders=minimum_orders,
        gia=gia,
        tinh_trang=tinh_trang,
        link_anh=link_anh,
    )


async def parse_current_page(
    page: Page, large_name: str, small_name: str
) -> int:
    """Parse all product rows on the current page."""
    # Scroll to ensure all rows are rendered
    await scroll_to_load_all(page)

    rows = page.locator(S.PRODUCT_ROWS)
    count = await rows.count()
    saved = 0

    await write_log(
        f"[DEBUG] Parsing {count} rows on page [{large_name} > {small_name}]"
    )

    for i in range(count):
        if RUNTIME.stop_requested:
            break
        row = rows.nth(i)
        try:
            # Check if the row is actually visible/has content
            try:
                is_visible = await row.is_visible()
                if not is_visible:
                    continue
            except Exception:
                pass

            pr = await parse_product_row(row, large_name, small_name)
            if pr.ten_san_pham or pr.size or pr.gia:
                await save_row(pr)
                saved += 1
        except Exception as e:
            await write_log(
                f"[WARN] Row parse fail [{large_name}>{small_name}] row={i}: {e}"
            )

    return saved


# =========================
# API INTERCEPTION + HTTPX FAST PATH
# =========================

async def setup_api_interception(page: Page, category_key: str) -> None:
    """Attach a response listener to capture product-list API calls.
    When the Vue app fetches product data via XHR/fetch, we capture the
    URL + headers so we can replay it with httpx for subsequent pages."""

    async def _on_response(response):
        try:
            url = response.url
            # Heuristic: product list APIs typically contain these patterns
            if response.status != 200:
                return
            content_type = response.headers.get("content-type", "")
            if "json" not in content_type:
                return
            # Look for URLs that seem like product/list endpoints
            url_lower = url.lower()
            if any(kw in url_lower for kw in [
                "/product", "/goods", "/item", "/list",
                "pagesize", "page_size", "pagenum", "page_num",
                "currentpage", "current_page",
            ]):
                # Capture this API
                request = response.request
                headers = dict(request.headers) if request.headers else {}
                # Remove browser-specific headers
                for h in ["accept-encoding", "sec-", "upgrade-insecure"]:
                    headers = {k: v for k, v in headers.items()
                               if not k.lower().startswith(h)}

                api = CapturedApi(
                    url=url,
                    method=request.method,
                    headers=headers,
                )
                RUNTIME.captured_apis[category_key] = api
                await write_log(
                    f"[API-CAPTURE] Intercepted API for {category_key}: {url[:120]}"
                )
        except Exception:
            pass  # Non-critical

    page.on("response", _on_response)


async def try_httpx_fetch_page(
    category_key: str,
    page_no: int,
    large_name: str,
    small_name: str,
) -> Optional[int]:
    """Try to fetch a product page via httpx using a captured API URL.
    Returns number of saved rows, or None if httpx path is not available."""
    api = RUNTIME.captured_apis.get(category_key)
    if not api:
        return None

    try:
        from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

        parsed = urlparse(api.url)
        params = parse_qs(parsed.query, keep_blank_values=True)

        # Try to modify the page parameter
        page_param_found = False
        for key in ["pageNum", "page_num", "page", "currentPage", "current_page",
                     "pageNo", "page_no", "offset", "p"]:
            if key in params:
                params[key] = [str(page_no)]
                page_param_found = True
                break

        if not page_param_found:
            return None  # Can't determine how to set page number

        # Rebuild URL
        new_query = urlencode({k: v[0] for k, v in params.items()})
        new_url = urlunparse(parsed._replace(query=new_query))

        client = await RUNTIME.get_httpx_client()
        resp = await client.get(
            new_url,
            headers=api.headers,
        )
        resp.raise_for_status()

        data = resp.json()

        # Try to find the product list in the response
        products = _extract_products_from_json(data)
        if products is None:
            return None

        saved = 0
        for item in products:
            if RUNTIME.stop_requested:
                break
            pr = _json_to_product_row(item, large_name, small_name)
            if pr and (pr.ten_san_pham or pr.size or pr.gia):
                await save_row(pr)
                saved += 1

        await write_log(
            f"[HTTPX] Page {page_no} [{category_key}]: fetched {saved} rows via API"
        )
        return saved

    except Exception as e:
        await write_log(f"[HTTPX-FAIL] {category_key} page {page_no}: {e}")
        return None  # Fallback to browser


def _extract_products_from_json(data: Any) -> Optional[List[Dict]]:
    """Recursively search JSON response for the product list array."""
    if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
        return data
    if isinstance(data, dict):
        # Common patterns: data.list, data.records, data.items, data.rows, data.data
        for key in ["list", "records", "items", "rows", "data", "products",
                     "goodsList", "result", "content"]:
            if key in data:
                result = _extract_products_from_json(data[key])
                if result:
                    return result
    return None


def _json_to_product_row(
    item: Dict[str, Any], large_name: str, small_name: str
) -> Optional[ProductRow]:
    """Convert a JSON product item to ProductRow. Field mapping is heuristic."""
    try:
        def _get(keys: List[str], default: str = "") -> str:
            for k in keys:
                if k in item:
                    v = item[k]
                    return clean_text(str(v)) if v is not None else ""
            return default

        name = _get(["productName", "product_name", "name", "goodsName",
                      "goods_name", "title"])
        material = _get(["material", "materialName", "material_name"])
        full_name = clean_text(f"{material} {name}") if material else name

        size = _get(["specNo", "spec_no", "specCode", "spec_code", "model",
                      "specification"])
        kho = _get(["warehouseName", "warehouse_name", "warehouse", "storeName"])
        brand = _get(["brandName", "brand_name", "brand"])

        stock_val = _get(["stockNum", "stock_num", "stock", "inventory", "qty"])
        stock_unit = _get(["stockUnit", "stock_unit", "unit"])
        stock = clean_text(f"{stock_val} {stock_unit}")

        moq_val = _get(["moq", "minOrderQty", "min_order_qty", "minOrder"])
        moq_unit = _get(["moqUnit", "moq_unit"])
        pkg = _get(["packageText", "package_text", "packageInfo"])
        minimum_orders = clean_multiline_text([
            f"MOQ: {moq_val} {moq_unit}" if moq_val else "",
            f"Package: {pkg}" if pkg else "",
        ])

        price_val = _get(["price", "unitPrice", "unit_price", "salePrice"])
        price_unit = _get(["priceUnit", "price_unit", "currency"])
        gia = clean_text(f"{price_val} {price_unit}")

        tinh_trang = _get(["stockStatus", "stock_status", "status"])
        link_anh = _get(["imageUrl", "image_url", "imgUrl", "img_url", "image",
                          "pic", "picUrl"])

        # Extract spec detail items if available
        detail_items = item.get("specItems", item.get("spec_items", []))
        if isinstance(detail_items, list):
            detail_lines = []
            for si in detail_items:
                if isinstance(si, dict):
                    lbl = si.get("label", si.get("specLabel", ""))
                    val = si.get("value", si.get("specValue", ""))
                    line = clean_text(f"{lbl} {val}")
                    if line:
                        detail_lines.append(line)
            detail = clean_multiline_text(detail_lines)
        else:
            detail = _get(["specInfo", "spec_info", "detail", "description"])

        return ProductRow(
            cate_lon=large_name,
            cate_nho=small_name,
            ten_san_pham=full_name,
            size=size,
            detail=detail,
            kho=kho,
            brand=brand,
            stock=stock,
            minimum_orders=minimum_orders,
            gia=gia,
            tinh_trang=tinh_trang,
            link_anh=link_anh,
        )
    except Exception:
        return None


def make_page_key(large_name: str, small_name: str, page_no: int) -> str:
    return f"{large_name}||{small_name}||{page_no}"


async def crawl_single_page(
    page: Page,
    large_name: str,
    small_name: str,
    page_no: int,
    completed_pages: Set[str],
) -> int:
    """Crawl one pagination page. Skip if already completed.
    Tries httpx fast path first for pages > 1, falls back to browser."""
    page_key = make_page_key(large_name, small_name, page_no)
    if page_key in completed_pages:
        return 0

    try:
        # Try httpx fast path for pages > 1
        if page_no > 1:
            category_key = f"{large_name}||{small_name}"
            httpx_result = await try_httpx_fetch_page(
                category_key, page_no, large_name, small_name
            )
            if httpx_result is not None and httpx_result > 0:
                await append_completed_page(page_key, large_name, small_name, page_no)
                completed_pages.add(page_key)
                return httpx_result
            # Fall through to browser-based crawling
            await goto_page_number(page, page_no)
        else:
            # For page 1, just make sure products are loaded
            await wait_products_loaded(page, min_rows=1)

        saved = await parse_current_page(page, large_name, small_name)
        await append_completed_page(page_key, large_name, small_name, page_no)
        completed_pages.add(page_key)
        await write_log(
            f"[OK] Page {page_no} [{large_name} > {small_name}]: saved {saved} rows"
        )
        return saved
    except Exception as e:
        await append_failed_page(page_key, large_name, small_name, page_no, str(e))
        await write_log(f"[ERROR] Page crawl failed [{page_key}]: {e}")
        return 0


async def crawl_one_category(
    browser: Browser,
    sc: SmallCategory,
    completed_pages: Set[str],
    progress: Progress,
) -> None:
    """Crawl all pages of a single small category using one browser context."""
    context = await create_context(browser)
    pw_page = await context.new_page()

    try:
        await navigate_to_category(pw_page, sc.large_name, sc.small_name)
        total_pages = await get_total_pages(pw_page)
        await write_log(
            f"[INFO] {sc.large_name} > {sc.small_name}: {total_pages} page(s)"
        )

        task_id = progress.add_task(
            f"[cyan]{sc.large_name} > {sc.small_name}",
            total=total_pages,
        )

        total_saved_cat = 0
        consecutive_failures = 0

        for page_no in range(1, total_pages + 1):
            if RUNTIME.stop_requested:
                break

            page_key = make_page_key(sc.large_name, sc.small_name, page_no)
            if page_key in completed_pages:
                progress.advance(task_id, 1)
                continue

            saved = await crawl_single_page(
                pw_page, sc.large_name, sc.small_name, page_no, completed_pages
            )
            total_saved_cat += saved

            if saved == 0 and page_key not in completed_pages:
                consecutive_failures += 1
            else:
                consecutive_failures = 0

            # If too many consecutive failures, try re-navigating
            if consecutive_failures >= 3:
                await write_log(
                    f"[WARN] {consecutive_failures} consecutive empty pages, "
                    f"re-navigating [{sc.large_name} > {sc.small_name}]"
                )
                try:
                    await navigate_to_category(
                        pw_page, sc.large_name, sc.small_name
                    )
                    await set_page_size_200(pw_page)
                    consecutive_failures = 0
                except Exception as e:
                    await write_log(f"[ERROR] Re-navigation failed: {e}")
                    break

            progress.advance(task_id, 1)
            progress.update(
                task_id,
                description=(
                    f"[green]{sc.large_name} > {sc.small_name}"
                    f" | pg {page_no}/{total_pages} | +{saved}"
                ),
            )
            await flush_rows(force=False)

            # Polite delay between pages
            if page_no < total_pages:
                await asyncio.sleep(POLITE_DELAY)

        await write_log(
            f"[DONE] {sc.large_name} > {sc.small_name}: "
            f"saved {total_saved_cat} rows from {total_pages} pages"
        )

    except Exception as e:
        await write_log(
            f"[ERROR] Category failed [{sc.large_name} > {sc.small_name}]: {e}"
        )
    finally:
        await flush_rows(force=True)
        await context.close()


async def pipeline_2_crawl(
    browser: Browser, categories: List[SmallCategory]
) -> None:
    """Pipeline 2: crawl product data from selected categories with concurrency."""
    completed_pages = await load_completed_pages()
    sem = asyncio.Semaphore(BROWSER_CONCURRENCY)

    with Progress(
        SpinnerColumn(),
        TextColumn("{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        TimeElapsedColumn(),
        console=console,
        transient=False,
    ) as progress:

        async def _run(sc: SmallCategory) -> None:
            async with sem:
                await crawl_one_category(browser, sc, completed_pages, progress)

        # Process all categories concurrently (limited by semaphore)
        tasks = [asyncio.create_task(_run(sc)) for sc in categories]

        for coro in asyncio.as_completed(tasks):
            try:
                await coro
            except Exception as e:
                await write_log(f"[ERROR] Task exception: {e}")

            if RUNTIME.stop_requested:
                for t in tasks:
                    t.cancel()
                break

    await flush_rows(force=True)


# =========================
# FINALIZE
# =========================

def build_final_file() -> None:
    if not DATA_CSV_FILE.exists():
        console.print("[red]No data CSV to finalize.[/red]")
        return

    df = pd.read_csv(DATA_CSV_FILE, dtype=str).fillna("")
    dedupe_cols = ["cate_lon", "cate_nho", "ten_san_pham", "size", "kho", "brand"]
    df = df.drop_duplicates(subset=dedupe_cols, keep="first")
    df.to_csv(FINAL_CSV_FILE, index=False, encoding="utf-8-sig")

    console.print(f"[green]Final file: {FINAL_CSV_FILE}[/green]")
    console.print(f"[green]Total rows after dedupe: {len(df):,}[/green]")


# =========================
# MAIN
# =========================

async def main() -> None:
    install_signal_handlers()
    await write_log("=== START RUN ===")

    if FRESH_START:
        console.print("[yellow]Fresh start: clearing old state files...[/yellow]")
        clear_state_files()

    console.print(
        Panel.fit(
            "Luosiwang 2-Pipeline Crawler\n"
            "- Pipeline 1: Category selection\n"
            "- Pipeline 2: Product data crawling\n"
            "- Resume / CSV checkpoint / Dedupe / Progress\n"
            f"- Mode: {'AUTO ALL' if AUTO_ALL else 'Interactive'}",
            title="Crawler",
            border_style="blue",
        )
    )

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=HEADLESS,
            args=[
                "--start-maximized",
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--disable-extensions",
                "--disable-background-networking",
                "--disable-sync",
                "--disable-translate",
                "--disable-default-apps",
                "--mute-audio",
                "--no-first-run",
                "--disable-hang-monitor",
                "--disable-popup-blocking",
                "--disable-prompt-on-repost",
                "--disable-renderer-backgrounding",
                "--disable-component-update",
            ],
        )

        try:
            # Pipeline 1: category selection (or resume from saved state)
            if SELECTED_CATEGORIES_FILE.exists() and not FRESH_START:
                mode = prompt_or_default(
                    "Found saved categories. Reuse? [y/n]", "y"
                ).strip().lower()
                if mode == "y":
                    data = orjson.loads(SELECTED_CATEGORIES_FILE.read_bytes())
                    selected = [SmallCategory(**x) for x in data]
                else:
                    selected = await pipeline_1_collect_selection(browser)
            else:
                selected = await pipeline_1_collect_selection(browser)

            if not selected:
                console.print("[red]No categories selected.[/red]")
                return

            console.print(
                f"\n[green]Will crawl {len(selected)} sub-categories:[/green]"
            )
            for i, sc in enumerate(selected[:30], 1):
                console.print(f"  {i}. {sc.large_name} > {sc.small_name}")
            if len(selected) > 30:
                console.print(f"  ... and {len(selected) - 30} more")

            confirm = prompt_or_default(
                "\nStart Pipeline 2? [y/n]", "y"
            ).strip().lower()
            if confirm != "y":
                console.print("[yellow]Aborted before Pipeline 2.[/yellow]")
                return

            # Pipeline 2: crawl
            await pipeline_2_crawl(browser, selected)

        finally:
            await flush_rows(force=True)
            await RUNTIME.close_httpx()
            await browser.close()

    build_final_file()
    await write_log(f"=== END RUN | saved={RUNTIME.total_saved} ===")
    console.print(
        f"\n[green]Done. Total new records: {RUNTIME.total_saved:,}[/green]"
    )


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        console.print("\n[yellow]Stopped by user.[/yellow]")
    finally:
        try:
            if DATA_CSV_FILE.exists():
                build_final_file()
        except Exception:
            pass
