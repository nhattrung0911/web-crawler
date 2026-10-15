"""
Async Bosch product scraper.

This script crawls Bosch Vietnam category pages, discovers product URLs,
fetches product detail pages concurrently, and writes the result to Excel
or CSV.

Install dependencies:

    pip install httpx beautifulsoup4 unidecode pandas openpyxl

Example usage:

    python crawl_bosch.py
    python crawl_bosch.py --output bosch_products.csv
    python crawl_bosch.py --product-concurrency 12 --retries 4
    python crawl_bosch.py https://vn.bosch-pt.com/vn/vi/cordless-tools-131400-ocs-c/
"""

from __future__ import annotations

import argparse
import ast
import asyncio
import json
import logging
import random
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Awaitable, Callable, Iterable, Optional, Sequence, TypeVar
from urllib.parse import urljoin

import httpx
import pandas as pd
from bs4 import BeautifulSoup
from unidecode import unidecode


BASE_URL = "https://vn.bosch-pt.com"
START_URL = f"{BASE_URL}/vn/vi/professional-power-tools-131398-ocs-c/"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}
RETRYABLE_STATUS_CODES = frozenset({408, 429, 500, 502, 503, 504})
DEFAULT_RESUME_FILE = "bosch_products.resume.jsonl"
DEFAULT_FAILED_FILE = "bosch_products.failed.jsonl"
logger = logging.getLogger(__name__)

T = TypeVar("T")
R = TypeVar("R")


@dataclass(slots=True)
class ScraperConfig:
    timeout: float = 20.0
    connect_timeout: float = 10.0
    pool_timeout: float = 30.0
    retries: int = 3
    retry_backoff: float = 1.0
    max_connections: int = 20
    max_keepalive_connections: int = 10
    category_concurrency: int = 3
    product_concurrency: int = 10
    requests_per_second: float = 4.0


@dataclass(slots=True)
class ScrapeStats:
    categories_seen: int = 0
    category_pages_seen: int = 0
    product_urls_seen: int = 0
    product_pages_parsed: int = 0
    request_failures: int = 0
    product_failures: int = 0
    resumed_products: int = 0


@dataclass(slots=True)
class ProductInfo:
    url: str
    name: str
    product_type: str
    description: str
    images: list[str]
    specs: list[str]
    order_code: str


@dataclass(slots=True)
class ScrapeResult:
    dataframe: pd.DataFrame
    stats: ScrapeStats
    discovered_product_urls: list[str]
    dry_run: bool = False


class RateLimiter:
    def __init__(self, requests_per_second: float) -> None:
        self._interval = 0.0 if requests_per_second <= 0 else 1.0 / requests_per_second
        self._lock = asyncio.Lock()
        self._next_slot = 0.0

    async def acquire(self) -> None:
        if self._interval <= 0:
            return
        async with self._lock:
            loop = asyncio.get_running_loop()
            now = loop.time()
            if now < self._next_slot:
                await asyncio.sleep(self._next_slot - now)
                now = loop.time()
            self._next_slot = max(now, self._next_slot) + self._interval


class ResumeStore:
    def __init__(self, path: Optional[str]) -> None:
        self.path = Path(path) if path else None
        self._lock = asyncio.Lock()
        self._loaded_rows: dict[str, dict[str, str]] = {}

    @property
    def enabled(self) -> bool:
        return self.path is not None

    @property
    def completed_urls(self) -> set[str]:
        return set(self._loaded_rows)

    async def load_rows(self) -> dict[str, dict[str, str]]:
        if not self.path or not self.path.exists():
            self._loaded_rows = {}
            return self._loaded_rows

        def _read_lines(path: Path) -> dict[str, dict[str, str]]:
            loaded: dict[str, dict[str, str]] = {}
            with path.open("r", encoding="utf-8") as handle:
                for raw_line in handle:
                    line = raw_line.strip()
                    if not line:
                        continue
                    try:
                        row = json.loads(line)
                    except json.JSONDecodeError:
                        continue
                    if not isinstance(row, dict):
                        continue
                    product_url = str(row.get("Product URL", "")).strip()
                    if product_url:
                        loaded[product_url] = {str(key): str(value) for key, value in row.items()}
            return loaded

        self._loaded_rows = await asyncio.to_thread(_read_lines, self.path)
        return self._loaded_rows

    async def append_row(self, row: dict[str, str]) -> None:
        if not self.path:
            return

        payload = json.dumps(row, ensure_ascii=False) + "\n"

        def _append(path: Path, line: str) -> None:
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open("a", encoding="utf-8") as handle:
                handle.write(line)

        async with self._lock:
            await asyncio.to_thread(_append, self.path, payload)
            self._loaded_rows[row["Product URL"]] = row


class FailedProductStore:
    def __init__(self, path: Optional[str]) -> None:
        self.path = Path(path) if path else None
        self._lock = asyncio.Lock()

    @property
    def enabled(self) -> bool:
        return self.path is not None

    async def append_failure(self, product_url: str, stage: str, error: str) -> None:
        if not self.path:
            return

        payload = json.dumps(
            {
                "product_url": product_url,
                "stage": stage,
                "error": error,
            },
            ensure_ascii=False,
        ) + "\n"

        def _append(path: Path, line: str) -> None:
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open("a", encoding="utf-8") as handle:
                handle.write(line)

        async with self._lock:
            await asyncio.to_thread(_append, self.path, payload)


async def load_failed_product_urls(path: Optional[str]) -> list[str]:
    failed_path = Path(path) if path else None
    if not failed_path or not failed_path.exists():
        return []

    def _read(path_obj: Path) -> list[str]:
        urls: list[str] = []
        with path_obj.open("r", encoding="utf-8") as handle:
            for raw_line in handle:
                line = raw_line.strip()
                if not line:
                    continue
                try:
                    record = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if not isinstance(record, dict):
                    continue
                product_url = str(record.get("product_url", "")).strip()
                if product_url:
                    urls.append(product_url)
        return dedupe_preserve_order(urls)

    return await asyncio.to_thread(_read, failed_path)


def load_failed_records(path: Optional[str]) -> list[dict[str, str]]:
    failed_path = Path(path) if path else None
    if not failed_path or not failed_path.exists():
        return []

    records_by_url: dict[str, dict[str, str]] = {}
    with failed_path.open("r", encoding="utf-8") as handle:
        for raw_line in handle:
            line = raw_line.strip()
            if not line:
                continue
            try:
                record = json.loads(line)
            except json.JSONDecodeError:
                continue
            if not isinstance(record, dict):
                continue
            product_url = str(record.get("product_url", "")).strip()
            if not product_url:
                continue
            records_by_url[product_url] = {
                "product_url": product_url,
                "stage": str(record.get("stage", "")).strip(),
                "error": str(record.get("error", "")).strip(),
            }
    return list(records_by_url.values())


def slugify(name: str) -> str:
    text = unidecode(name).lower()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")


def dedupe_preserve_order(items: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for item in items:
        if item not in seen:
            seen.add(item)
            ordered.append(item)
    return ordered


def product_to_row(product: ProductInfo) -> dict[str, str]:
    return {
        "Product URL": product.url,
        "Tên sản phẩm": product.name,
        "Loại sản phẩm": product.product_type,
        "Mô tả": product.description,
        "Image URLs": "\n".join(product.images),
        "Spec": "\n".join(product.specs),
        "Mã đơn hàng": product.order_code,
    }


def build_absolute_url(href: str, base_url: str = BASE_URL) -> str:
    return urljoin(base_url.rstrip("/") + "/", href.strip())


def normalize_category_url(url: str) -> str:
    return build_absolute_url(url).rstrip("/") + "/"


def normalize_category_inputs(categories: Optional[Sequence[str]]) -> Optional[list[str]]:
    if not categories:
        return None
    return dedupe_preserve_order(normalize_category_url(url) for url in categories)


def parse_js_string_array(raw_values: str) -> list[str]:
    candidate = f"[{raw_values}]"
    try:
        values = json.loads(candidate)
    except json.JSONDecodeError:
        values = None
    if isinstance(values, list):
        return [str(item).strip() for item in values if str(item).strip()]
    try:
        values = ast.literal_eval(candidate)
    except (SyntaxError, ValueError):
        pattern = re.compile(r"""(['"])(.*?)(?<!\\)\1""", re.DOTALL)
        parsed = []
        for _, value in pattern.findall(raw_values):
            try:
                parsed.append(bytes(value, "utf-8").decode("unicode_escape"))
            except UnicodeDecodeError:
                parsed.append(value)
        return [item.strip() for item in parsed if item.strip()]
    if not isinstance(values, list):
        return []
    return [str(item).strip() for item in values if str(item).strip()]


def find_product_schema(node: Any) -> Optional[dict[str, Any]]:
    if isinstance(node, dict):
        node_type = node.get("@type")
        if node_type == "Product" or (
            isinstance(node_type, list) and "Product" in node_type
        ):
            return node
        for value in node.values():
            match = find_product_schema(value)
            if match:
                return match
    elif isinstance(node, list):
        for item in node:
            match = find_product_schema(item)
            if match:
                return match
    return None


def get_category_links_from_html(html: str) -> list[str]:
    soup = BeautifulSoup(html, "html.parser")
    links: list[str] = []
    for section in soup.find_all("section", class_="o-pop-category-grid"):
        for anchor in section.find_all("a", href=True):
            href = anchor["href"].strip()
            if href.endswith("-ocs-c/"):
                links.append(normalize_category_url(href))
    return dedupe_preserve_order(links)


def get_pagination_urls(category_url: str, html: str) -> list[str]:
    normalized_category = normalize_category_url(category_url)
    urls = [normalized_category]
    soup = BeautifulSoup(html, "html.parser")
    nav = soup.find("div", class_="m-ghostblock__bottom-nav")
    if not nav:
        nav = soup.find("div", class_="m-ghostblock__main-nav")
    if not nav:
        return urls

    for anchor in nav.find_all("a", href=True):
        normalized_href = normalize_category_url(anchor["href"])
        if normalized_href == normalized_category or normalized_href.startswith(
            normalized_category + "page/"
        ):
            urls.append(normalized_href)
    return dedupe_preserve_order(urls)


def parse_category_page_html(html: str) -> list[str]:
    product_links: list[str] = []
    id_match = re.search(r"utag_data\.product_id\s*=\s*\[(.*?)\]", html, re.DOTALL)
    name_match = re.search(r"utag_data\.product_name\s*=\s*\[(.*?)\]", html, re.DOTALL)
    if id_match and name_match:
        product_ids = parse_js_string_array(id_match.group(1))
        product_names = parse_js_string_array(name_match.group(1))
        if len(product_ids) == len(product_names) and product_ids:
            for name, product_id in zip(product_names, product_ids):
                product_links.append(
                    f"{BASE_URL}/vn/vi/products/{slugify(name)}-{product_id}"
                )
            return dedupe_preserve_order(product_links)

    soup = BeautifulSoup(html, "html.parser")
    anchors = soup.select("a.category-grid-tile__link-wrapper[href]")
    if not anchors:
        anchors = soup.select("a[href*='/vn/vi/products/']")
    for anchor in anchors:
        href = anchor.get("href", "").strip()
        if href:
            product_links.append(build_absolute_url(href))
    return dedupe_preserve_order(product_links)


def extract_specs(soup: BeautifulSoup) -> list[str]:
    specs: list[str] = []
    tech_section = soup.find("section", class_="o-technical_data")
    if not tech_section:
        return specs
    for row in tech_section.select("div.table__body-row"):
        cells = row.find_all(class_="table__body-cell")
        if len(cells) != 2:
            continue
        key = cells[0].get_text(" ", strip=True)
        value = cells[1].get_text(" ", strip=True)
        if key and value:
            specs.append(f"{key}: {value}")
    return specs


def extract_order_code(soup: BeautifulSoup) -> str:
    code_pattern = re.compile(
        r"ma\s+don\s+hang\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9.\-]*)",
        re.IGNORECASE,
    )
    prioritized_nodes = soup.select(
        ".o-gt-stage__intro p, .o-gt-stage__intro .bba--small, .o-gt-stage__intro li"
    )
    fallback_nodes = soup.find_all(["p", "div", "span", "li"])
    ordered_nodes = list(prioritized_nodes) + [
        node for node in fallback_nodes if node not in prioritized_nodes
    ]

    for node in ordered_nodes:
        text = node.get_text(" ", strip=True)
        if not text:
            continue
        normalized = re.sub(r"\s+", " ", unidecode(text)).replace("\u200b", "").strip()
        if "ma don hang" not in normalized.lower():
            continue
        match = code_pattern.search(normalized)
        if match:
            return match.group(1).strip()
    return ""


def extract_images(soup: BeautifulSoup, schema: Optional[dict[str, Any]]) -> list[str]:
    if schema:
        images = schema.get("image")
        if isinstance(images, list):
            return [str(image).strip() for image in images if str(image).strip()]
        if isinstance(images, str) and images.strip():
            return [images.strip()]

    meta_images = [
        tag["content"].strip()
        for tag in soup.select("meta[property='og:image'][content]")
        if tag["content"].strip()
    ]
    return dedupe_preserve_order(meta_images)


def parse_product_html(product_url: str, html: str) -> ProductInfo:
    soup = BeautifulSoup(html, "html.parser")
    schema: Optional[dict[str, Any]] = None

    for script in soup.find_all("script", attrs={"type": "application/ld+json"}):
        raw_json = script.string or script.get_text(strip=True)
        if not raw_json:
            continue
        try:
            data = json.loads(raw_json)
        except json.JSONDecodeError:
            continue
        schema = find_product_schema(data)
        if schema:
            break

    name = ""
    product_type = ""
    description = ""
    if schema:
        name = str(schema.get("name", "")).strip()
        description = str(schema.get("description", "")).strip()

    if not name:
        heading = soup.find("h1")
        if heading:
            name = heading.get_text(" ", strip=True)

    subtitle = soup.select_one(".o-gt-stage__intro h2.bba--medium")
    if subtitle:
        product_type = subtitle.get_text(" ", strip=True)
    else:
        fallback_subtitle = soup.find("h2")
        if fallback_subtitle:
            product_type = fallback_subtitle.get_text(" ", strip=True)

    if not description:
        meta_desc = soup.find("meta", attrs={"property": "og:description"})
        if meta_desc and meta_desc.get("content"):
            description = meta_desc["content"].strip()

    images = extract_images(soup, schema)
    specs = extract_specs(soup)
    order_code = extract_order_code(soup)

    return ProductInfo(
        url=product_url,
        name=name,
        product_type=product_type,
        description=description,
        images=images,
        specs=specs,
        order_code=order_code,
    )


async def map_limited(
    items: Sequence[T],
    worker: Callable[[T], Awaitable[R]],
    limit: int,
) -> list[R]:
    if not items:
        return []

    semaphore = asyncio.Semaphore(max(1, limit))
    results: list[Optional[R]] = [None] * len(items)

    async def run_one(index: int, item: T) -> None:
        async with semaphore:
            results[index] = await worker(item)

    async with asyncio.TaskGroup() as task_group:
        for index, item in enumerate(items):
            task_group.create_task(run_one(index, item))

    return [result for result in results if result is not None]


async def fetch_text(
    client: httpx.AsyncClient,
    url: str,
    config: ScraperConfig,
    stats: ScrapeStats,
    label: str,
    rate_limiter: RateLimiter,
) -> Optional[str]:
    for attempt in range(config.retries + 1):
        try:
            await rate_limiter.acquire()
            response = await client.get(url)
            if (
                response.status_code in RETRYABLE_STATUS_CODES
                and attempt < config.retries
            ):
                await asyncio.sleep(
                    config.retry_backoff * (2**attempt) + random.uniform(0.0, 0.25)
                )
                continue
            response.raise_for_status()
            return response.text
        except (httpx.TimeoutException, httpx.NetworkError, httpx.RemoteProtocolError) as exc:
            if attempt < config.retries:
                await asyncio.sleep(
                    config.retry_backoff * (2**attempt) + random.uniform(0.0, 0.25)
                )
                continue
            stats.request_failures += 1
            logger.warning("Request failed for %s %s: %s", label, url, exc)
            return None
        except httpx.HTTPStatusError as exc:
            status_code = exc.response.status_code
            if status_code in RETRYABLE_STATUS_CODES and attempt < config.retries:
                await asyncio.sleep(
                    config.retry_backoff * (2**attempt) + random.uniform(0.0, 0.25)
                )
                continue
            stats.request_failures += 1
            logger.warning("HTTP %s for %s %s", status_code, label, url)
            return None
    return None


async def get_category_links(
    client: httpx.AsyncClient,
    config: ScraperConfig,
    stats: ScrapeStats,
    rate_limiter: RateLimiter,
) -> list[str]:
    html = await fetch_text(client, START_URL, config, stats, "start page", rate_limiter)
    if not html:
        return []
    return get_category_links_from_html(html)


async def scrape_category(
    client: httpx.AsyncClient,
    category_url: str,
    config: ScraperConfig,
    stats: ScrapeStats,
    rate_limiter: RateLimiter,
) -> tuple[str, list[str]]:
    first_page_html = await fetch_text(
        client, category_url, config, stats, "category", rate_limiter
    )
    if not first_page_html:
        return category_url, []

    page_urls = get_pagination_urls(category_url, first_page_html)
    stats.category_pages_seen += len(page_urls)
    page_html_by_url: dict[str, str] = {normalize_category_url(category_url): first_page_html}

    remaining_page_urls = [
        page_url
        for page_url in page_urls
        if page_url not in page_html_by_url
    ]

    async def fetch_page(page_url: str) -> tuple[str, Optional[str]]:
        return page_url, await fetch_text(
            client, page_url, config, stats, "category page", rate_limiter
        )

    page_results = await map_limited(
        remaining_page_urls,
        fetch_page,
        limit=config.category_concurrency,
    )
    for page_url, page_html in page_results:
        if page_html:
            page_html_by_url[page_url] = page_html

    product_urls: list[str] = []
    for page_url in page_urls:
        page_html = page_html_by_url.get(page_url)
        if page_html:
            product_urls.extend(parse_category_page_html(page_html))

    deduped_product_urls = dedupe_preserve_order(product_urls)
    logger.info("%s -> %s products", category_url, len(deduped_product_urls))
    return category_url, deduped_product_urls


async def parse_product(
    client: httpx.AsyncClient,
    product_url: str,
    config: ScraperConfig,
    stats: ScrapeStats,
    rate_limiter: RateLimiter,
    failed_store: FailedProductStore,
) -> Optional[ProductInfo]:
    html = await fetch_text(client, product_url, config, stats, "product", rate_limiter)
    if not html:
        stats.product_failures += 1
        await failed_store.append_failure(product_url, "fetch_product", "fetch returned no HTML")
        return None

    try:
        product = parse_product_html(product_url, html)
    except Exception as exc:
        stats.product_failures += 1
        logger.warning("Failed to parse product %s: %s", product_url, exc)
        await failed_store.append_failure(product_url, "parse_product", str(exc))
        return None

    stats.product_pages_parsed += 1
    return product


def build_dataframe(products: Sequence[ProductInfo]) -> pd.DataFrame:
    return pd.DataFrame([product_to_row(product) for product in products])


def build_dataframe_from_rows(rows: Sequence[dict[str, str]]) -> pd.DataFrame:
    return pd.DataFrame(list(rows))


def write_output(df: pd.DataFrame, output_path: str) -> None:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    suffix = Path(output_path).suffix.lower()
    if suffix == ".csv":
        df.to_csv(output, index=False, encoding="utf-8-sig")
        return
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        df.to_excel(output, index=False)
        return
    raise ValueError("Unsupported output format. Use .xlsx, .xlsm, .xls, or .csv")


def export_failed_records_csv(failed_jsonl_path: Optional[str], csv_output_path: Optional[str]) -> int:
    if not failed_jsonl_path or not csv_output_path:
        return 0
    records = load_failed_records(failed_jsonl_path)
    if not records:
        return 0
    output = Path(csv_output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(records).to_csv(output, index=False, encoding="utf-8-sig")
    return len(records)


async def scrape_all(
    categories: Optional[Sequence[str]],
    config: ScraperConfig,
    resume_path: Optional[str] = None,
    failed_path: Optional[str] = None,
    retry_failed_from: Optional[str] = None,
    sample_products: Optional[int] = None,
    dry_run: bool = False,
) -> ScrapeResult:
    stats = ScrapeStats()
    normalized_categories = normalize_category_inputs(categories)
    rate_limiter = RateLimiter(config.requests_per_second)
    resume_store = ResumeStore(resume_path)
    failed_store = FailedProductStore(failed_path)
    existing_rows = await resume_store.load_rows()
    retry_product_urls = await load_failed_product_urls(retry_failed_from)

    timeout = httpx.Timeout(
        timeout=config.timeout,
        connect=config.connect_timeout,
        pool=config.pool_timeout,
    )
    limits = httpx.Limits(
        max_connections=config.max_connections,
        max_keepalive_connections=config.max_keepalive_connections,
    )
    transport = httpx.AsyncHTTPTransport(retries=1)

    async with httpx.AsyncClient(
        headers=HEADERS,
        timeout=timeout,
        limits=limits,
        transport=transport,
        follow_redirects=True,
    ) as client:
        if retry_product_urls:
            all_product_urls = retry_product_urls
            logger.info("Retry mode: loaded %s failed product URLs from %s", len(all_product_urls), retry_failed_from)
        else:
            if normalized_categories is None:
                normalized_categories = await get_category_links(
                    client, config, stats, rate_limiter
                )

            if not normalized_categories:
                return ScrapeResult(pd.DataFrame(), stats, [])

            stats.categories_seen = len(normalized_categories)
            logger.info("Crawling %s category pages...", len(normalized_categories))

            category_results = await map_limited(
                normalized_categories,
                lambda category_url: scrape_category(
                    client, category_url, config, stats, rate_limiter
                ),
                limit=config.category_concurrency,
            )
            all_product_urls = dedupe_preserve_order(
                product_url
                for _, product_urls in category_results
                for product_url in product_urls
            )

        if sample_products is not None:
            all_product_urls = all_product_urls[: max(0, sample_products)]
            logger.info("Sample mode: restricted to %s product URLs", len(all_product_urls))

        stats.product_urls_seen = len(all_product_urls)
        logger.info("Discovered %s unique product URLs", stats.product_urls_seen)

        if dry_run:
            preview = all_product_urls[: min(5, len(all_product_urls))]
            if preview:
                logger.info("Dry run preview URLs: %s", preview)
            return ScrapeResult(pd.DataFrame(), stats, all_product_urls, dry_run=True)

        cached_rows = [
            existing_rows[product_url]
            for product_url in all_product_urls
            if product_url in existing_rows
        ]
        pending_product_urls = [
            product_url
            for product_url in all_product_urls
            if product_url not in existing_rows
        ]
        stats.resumed_products = len(cached_rows)
        if resume_store.enabled and cached_rows:
            logger.info("Resume hit: skipping %s already completed products", len(cached_rows))

        async def parse_and_store(product_url: str) -> Optional[dict[str, str]]:
            product = await parse_product(
                client, product_url, config, stats, rate_limiter, failed_store
            )
            if product is None:
                return None
            row = product_to_row(product)
            if resume_store.enabled:
                await resume_store.append_row(row)
            return row

        new_rows = await map_limited(
            pending_product_urls,
            parse_and_store,
            limit=config.product_concurrency,
        )

    final_rows = cached_rows + [row for row in new_rows if isinstance(row, dict)]
    return ScrapeResult(
        dataframe=build_dataframe_from_rows(final_rows),
        stats=stats,
        discovered_product_urls=all_product_urls,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scrape Bosch product data asynchronously.")
    parser.add_argument(
        "categories",
        nargs="*",
        help="Optional category URLs to crawl. If omitted, all categories are discovered.",
    )
    parser.add_argument(
        "--output",
        default="bosch_products_enhanced.xlsx",
        help="Output filename. Supported extensions: .xlsx, .xlsm, .xls, .csv",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=20.0,
        help="Per-request read timeout in seconds.",
    )
    parser.add_argument(
        "--connect-timeout",
        type=float,
        default=10.0,
        help="Connection timeout in seconds.",
    )
    parser.add_argument(
        "--pool-timeout",
        type=float,
        default=30.0,
        help="Connection-pool wait timeout in seconds.",
    )
    parser.add_argument(
        "--retries",
        type=int,
        default=3,
        help="Number of retry attempts for retryable failures.",
    )
    parser.add_argument(
        "--retry-backoff",
        type=float,
        default=1.0,
        help="Base backoff in seconds. Actual sleep grows exponentially per retry.",
    )
    parser.add_argument(
        "--max-connections",
        type=int,
        default=20,
        help="Maximum total HTTP connections kept by the client.",
    )
    parser.add_argument(
        "--max-keepalive-connections",
        type=int,
        default=10,
        help="Maximum keep-alive connections retained by the client.",
    )
    parser.add_argument(
        "--category-concurrency",
        type=int,
        default=3,
        help="Number of categories processed concurrently.",
    )
    parser.add_argument(
        "--product-concurrency",
        type=int,
        default=10,
        help="Number of product pages processed concurrently.",
    )
    parser.add_argument(
        "--requests-per-second",
        type=float,
        default=4.0,
        help="Global request rate limit. Use 0 to disable rate limiting.",
    )
    parser.add_argument(
        "--resume-file",
        default=DEFAULT_RESUME_FILE,
        help="JSONL file for incremental resume/checkpoint. Use '' to disable.",
    )
    parser.add_argument(
        "--failed-file",
        default=DEFAULT_FAILED_FILE,
        help="JSONL file where failed product URLs are appended. Use '' to disable.",
    )
    parser.add_argument(
        "--export-failed-csv",
        default="",
        help="Optional CSV export path built from failed-file JSONL after the run.",
    )
    parser.add_argument(
        "--retry-failed-from",
        default="",
        help="Retry product URLs from a previous failed-file JSONL instead of recrawling categories.",
    )
    parser.add_argument(
        "--sample-products",
        type=int,
        default=None,
        help="Only crawl the first N discovered product URLs.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Discover product URLs and print a preview without fetching product detail pages.",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Logging verbosity.",
    )
    return parser.parse_args()


def build_config(args: argparse.Namespace) -> ScraperConfig:
    return ScraperConfig(
        timeout=max(1.0, args.timeout),
        connect_timeout=max(1.0, args.connect_timeout),
        pool_timeout=max(1.0, args.pool_timeout),
        retries=max(0, args.retries),
        retry_backoff=max(0.0, args.retry_backoff),
        max_connections=max(1, args.max_connections),
        max_keepalive_connections=max(1, args.max_keepalive_connections),
        category_concurrency=max(1, args.category_concurrency),
        product_concurrency=max(1, args.product_concurrency),
        requests_per_second=max(0.0, args.requests_per_second),
    )


def main() -> None:
    args = parse_args()
    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s %(levelname)s %(message)s",
    )
    config = build_config(args)
    resume_path = args.resume_file.strip() or None
    failed_path = args.failed_file.strip() or None
    failed_csv_path = args.export_failed_csv.strip() or None
    retry_failed_from = args.retry_failed_from.strip() or None
    result = asyncio.run(
        scrape_all(
            args.categories,
            config,
            resume_path=resume_path,
            failed_path=failed_path,
            retry_failed_from=retry_failed_from,
            sample_products=args.sample_products,
            dry_run=args.dry_run,
        )
    )

    if result.dry_run:
        logger.info(
            "Dry run complete. Discovered %s product URLs. No output file was written.",
            len(result.discovered_product_urls),
        )
        return

    if result.dataframe.empty:
        logger.warning("No products were scraped.")
        return

    write_output(result.dataframe, args.output)
    logger.info("Saved %s products to %s", len(result.dataframe), args.output)
    logger.info(
        "Summary: categories=%s, category_pages=%s, products_discovered=%s, "
        "products_parsed=%s, request_failures=%s, product_failures=%s",
        result.stats.categories_seen,
        result.stats.category_pages_seen,
        result.stats.product_urls_seen,
        result.stats.product_pages_parsed,
        result.stats.request_failures,
        result.stats.product_failures,
    )
    if result.stats.resumed_products:
        logger.info("Resume reused %s previously completed products", result.stats.resumed_products)
    exported_failed_count = export_failed_records_csv(failed_path, failed_csv_path)
    if exported_failed_count:
        logger.info(
            "Exported %s failed product records to %s",
            exported_failed_count,
            failed_csv_path,
        )


def launch_gui():
    import os
    import queue
    import subprocess
    import sys
    import threading
    import customtkinter as ctk
    from tkinter import messagebox

    HERE = Path(__file__).parent
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")

    class App(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title("Bosch PT Crawler (vn.bosch-pt.com)")
            self.geometry("900x600")
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(3, weight=1)
            self.proc = None
            self.log_q = queue.Queue()

            f = ctk.CTkFrame(self)
            f.grid(row=0, column=0, padx=16, pady=(16, 4), sticky="ew")
            f.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(f, text="Category URLs (moi dong 1 URL; de trong = discover tat ca)"
                         ).grid(row=0, column=0, sticky="w", padx=8, pady=(6, 0))
            self.cats = ctk.CTkTextbox(f, height=90, font=("Consolas", 11))
            self.cats.grid(row=1, column=0, padx=8, pady=6, sticky="ew")

            o = ctk.CTkFrame(self, fg_color="transparent")
            o.grid(row=1, column=0, padx=18, sticky="ew")
            ctk.CTkLabel(o, text="Output").grid(row=0, column=0, padx=(0, 4))
            self.output = ctk.CTkEntry(o, width=220)
            self.output.insert(0, "bosch_products.xlsx")
            self.output.grid(row=0, column=1, padx=4)
            ctk.CTkLabel(o, text="Sample (0=het)").grid(row=0, column=2, padx=(16, 4))
            self.sample = ctk.CTkEntry(o, width=70)
            self.sample.insert(0, "0")
            self.sample.grid(row=0, column=3, padx=4)
            self.dry = ctk.CTkCheckBox(o, text="Dry-run (chi tim URL)")
            self.dry.grid(row=0, column=4, padx=16)

            bar = ctk.CTkFrame(self, fg_color="transparent")
            bar.grid(row=2, column=0, padx=18, pady=4, sticky="ew")
            self.btn = ctk.CTkButton(bar, text="START", width=120, command=self.start)
            self.btn.grid(row=0, column=0, padx=6)
            self.btn_stop = ctk.CTkButton(bar, text="STOP", width=90, fg_color="#a33",
                                          hover_color="#c44", command=self.stop, state="disabled")
            self.btn_stop.grid(row=0, column=1, padx=6)

            self.console = ctk.CTkTextbox(self, state="disabled", font=("Consolas", 12))
            self.console.grid(row=3, column=0, padx=16, pady=12, sticky="nsew")
            self.after(100, self._drain)

        def _log(self, m):
            self.console.configure(state="normal")
            self.console.insert("end", m)
            self.console.see("end")
            self.console.configure(state="disabled")

        def _drain(self):
            try:
                while True:
                    self._log(self.log_q.get_nowait())
            except queue.Empty:
                pass
            self.after(100, self._drain)

        def start(self):
            if self.proc:
                return
            cats = [c.strip() for c in self.cats.get("1.0", "end").splitlines() if c.strip()]
            args = list(cats)
            if self.output.get().strip():
                args += ["--output", self.output.get().strip()]
            s = self.sample.get().strip()
            if s and s != "0":
                args += ["--sample-products", s]
            if self.dry.get():
                args.append("--dry-run")
            self.btn.configure(state="disabled")
            self.btn_stop.configure(state="normal")
            threading.Thread(target=self._run, args=(args,), daemon=True).start()

        def _run(self, args):
            env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"}
            self.log_q.put(f"\n===== bosch {' '.join(args)} =====\n")
            try:
                self.proc = subprocess.Popen(
                    [sys.executable, str(HERE / "bosch.py"), *args], cwd=str(HERE),
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace", env=env)
            except OSError as e:
                self.log_q.put(f"[launch error] {e}\n")
            else:
                for line in self.proc.stdout:
                    self.log_q.put(line)
                self.proc.wait()
                self.log_q.put(f"[exit {self.proc.returncode}]\n")
            self.proc = None
            self.after(0, lambda: (self.btn.configure(state="normal"),
                                   self.btn_stop.configure(state="disabled")))

        def stop(self):
            if self.proc:
                self.proc.terminate()
                self.log_q.put("\n[da dung]\n")

    App().mainloop()


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        main()
    else:
        launch_gui()
