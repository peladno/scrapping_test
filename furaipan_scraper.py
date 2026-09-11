"""Furaipan Club Scraper for E-Commerce Price Monitoring.

Extracts product titles, catalog model codes, prices, and URLs from
furaipan.com kitchen knives and cutting boards catalog for price compliance.
"""

import os
import re
import sys
import time
from typing import Dict, List, Optional, Set

from bs4 import BeautifulSoup
from curl_cffi import requests as cffi_requests
import pandas as pd

from config import (
    CATALOG_LIST_EXCEL,
    COURTESY_PAUSE_SECONDS,
    FURAIPAN_BASE_URL,
    FURAIPAN_GROUP_URL,
    FURAIPAN_ITEM_CLASS,
    FURAIPAN_ITEM_LIST_CLASS,
    FURAIPAN_PRICE_CLASS,
    FURAIPAN_PRODUCT_KIND_CLASS,
    FURAIPAN_TITLE_CLASS,
    HTTP_RETRIES,
    HTTP_TIMEOUT,
    OUTPUT_FURAIPAN_SCRAPED_EXCEL,
    TARGET_KEYWORD,
)
from utils import (
    clean_price_text,
    extract_product_code,
    load_official_product_codes,
    save_excel_with_fallback,
)

# Configure UTF-8 encoding for Windows console output
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# Realistic browser request headers for Furaipan Club
DEFAULT_FURAIPAN_HEADERS: Dict[str, str] = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/128.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;q=0.9,"
        "image/avif,image/webp,image/apng,*/*;q=0.8"
    ),
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Sec-Ch-Ua": (
        '"Chromium";v="128", "Not;A=Brand";v="24", "Google Chrome";v="128"'
    ),
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": '"Windows"',
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
}


def extract_furaipan_item_links(
    group_page_html: str,
    base_url: str = FURAIPAN_BASE_URL,
) -> List[str]:
    """Extract product detail page URLs from the Furaipan group category page.

    Args:
        group_page_html: HTML content of the group listing page.
        base_url: Base domain URL to resolve relative links.

    Returns:
        List of unique absolute detail page URLs.
    """
    if not group_page_html:
        return []

    soup = BeautifulSoup(group_page_html, "html.parser")
    item_lists = soup.find_all(class_=re.compile(FURAIPAN_ITEM_LIST_CLASS))
    if not item_lists:
        item_lists = soup.find_all("div", class_=re.compile(r"item_list"))

    detail_links: List[str] = []
    seen: Set[str] = set()

    for item_list in item_lists:
        items = item_list.find_all(class_=re.compile(FURAIPAN_ITEM_CLASS))
        for item in items:
            link = item.find("a", href=True)
            if not link:
                continue
            href = link.get("href", "").strip()
            if not href:
                continue
            if href.startswith("/"):
                href = f"{base_url.rstrip('/')}{href}"
            elif not href.startswith("http"):
                href = f"{base_url.rstrip('/')}/{href}"

            if href not in seen:
                seen.add(href)
                detail_links.append(href)

    return detail_links


def parse_furaipan_detail_page(
    html_content: str,
    page_url: str,
    official_codes: Optional[List[str]] = None,
) -> List[Dict[str, str]]:
    """Parse product kinds from a single Furaipan product detail page HTML.

    Args:
        html_content: Raw HTML content of the detail page.
        page_url: Canonical URL of the detail page.
        official_codes: Optional list of official catalog model codes.

    Returns:
        List of extracted product dictionaries.
    """
    if not html_content:
        return []

    soup = BeautifulSoup(html_content, "html.parser")
    kinds = soup.find_all(class_=re.compile(FURAIPAN_PRODUCT_KIND_CLASS))
    if not kinds:
        kinds = soup.find_all("div", class_=re.compile(r"product_kind"))

    results: List[Dict[str, str]] = []

    for kind in kinds:
        title_el = kind.find(class_=re.compile(FURAIPAN_TITLE_CLASS))
        price_el = kind.find(class_=re.compile(FURAIPAN_PRICE_CLASS))

        title = title_el.get_text(strip=True) if title_el else ""
        raw_price = price_el.get_text(strip=True) if price_el else ""

        if not title:
            continue

        # Strict Filter: Omit items if title does not contain
        # 'GLOBAL' or 'グローバル'
        has_target = (
            "GLOBAL" in title.upper()
            or "グローバル" in title
        )
        if not has_target:
            continue

        price_text = clean_price_text(raw_price)

        # Extract official catalog code
        product_code = extract_product_code(title, official_codes)
        if not product_code:
            product_code = extract_product_code(page_url, official_codes)

        results.append({
            "Store": "furaipan",
            "Company": "フライパン倶楽部 (Furaipan Club)",
            "Product": title,
            "Product_Code": (
                product_code if product_code else TARGET_KEYWORD
            ),
            "Price": price_text,
            "Product_URL": page_url,
        })

    return results


def scrape_furaipan_products(
    group_url: str = FURAIPAN_GROUP_URL,
    official_codes: Optional[List[str]] = None,
    headers: Optional[Dict[str, str]] = None,
    local_html_path: Optional[str] = None,
) -> List[Dict[str, str]]:
    """Scrape product listings from Furaipan Club group page and detail pages.

    Args:
        group_url: URL of the category group page.
        official_codes: Catalog model codes list.
        headers: Optional HTTP headers dictionary.
        local_html_path: Optional local HTML file path for testing.

    Returns:
        Consolidated list of extracted product records.
    """
    if local_html_path and os.path.exists(local_html_path):
        print(
            f"Reading from local HTML file '{local_html_path}'...",
            flush=True,
        )
        with open(
            local_html_path, "r", encoding="utf-8", errors="ignore"
        ) as f:
            content = f.read()
        return parse_furaipan_detail_page(
            content, group_url, official_codes
        )

    req_headers = headers or DEFAULT_FURAIPAN_HEADERS
    print(f"  [Furaipan Group] Requesting: {group_url}", flush=True)

    group_html = ""
    for attempt in range(HTTP_RETRIES):
        try:
            response = cffi_requests.get(
                group_url,
                headers=req_headers,
                impersonate="chrome120",
                timeout=HTTP_TIMEOUT,
            )
            if response.status_code == 200:
                group_html = response.text
                break
            elif response.status_code in [403, 503]:
                print(
                    f"  [Notice] Status {response.status_code} "
                    f"(Anti-bot check). Retrying after pause...",
                    flush=True,
                )
                time.sleep(COURTESY_PAUSE_SECONDS * (attempt + 2))
        except Exception as err:
            print(f"  [Attempt {attempt + 1}] Error: {err}", flush=True)
            time.sleep(COURTESY_PAUSE_SECONDS)

    if not group_html:
        print(
            "  [Warning] Could not retrieve Furaipan group page.",
            flush=True,
        )
        return []

    detail_links = extract_furaipan_item_links(group_html)
    print(
        f"  [Furaipan] Found {len(detail_links)} product pages to crawl.",
        flush=True,
    )

    all_results: List[Dict[str, str]] = []

    for idx, page_url in enumerate(detail_links, 1):
        print(
            f"  [{idx}/{len(detail_links)}] Fetching detail: {page_url}",
            flush=True,
        )

        detail_html = ""
        for attempt in range(HTTP_RETRIES):
            try:
                detail_resp = cffi_requests.get(
                    page_url,
                    headers=req_headers,
                    impersonate="chrome120",
                    timeout=HTTP_TIMEOUT,
                )
                if detail_resp.status_code == 200:
                    detail_html = detail_resp.text
                    break
                elif detail_resp.status_code in [403, 503]:
                    time.sleep(COURTESY_PAUSE_SECONDS * (attempt + 2))
            except Exception as err:
                print(f"    [Detail Error] {err}", flush=True)
                time.sleep(COURTESY_PAUSE_SECONDS)

        if not detail_html:
            print(
                f"    [Warning] Failed to fetch detail page: {page_url}",
                flush=True,
            )
            continue

        page_records = parse_furaipan_detail_page(
            detail_html, page_url, official_codes
        )
        all_results.extend(page_records)
        time.sleep(COURTESY_PAUSE_SECONDS)

    return all_results


def scrape_all_furaipan_products(
    output_excel: str = OUTPUT_FURAIPAN_SCRAPED_EXCEL,
    list_products_file: str = CATALOG_LIST_EXCEL,
    group_url: str = FURAIPAN_GROUP_URL,
    local_html_path: Optional[str] = None,
) -> None:
    """Scrape Furaipan Club products and save multi-sheet Excel file.

    Args:
        output_excel: Destination Excel file path.
        list_products_file: Official product catalog Excel file.
        group_url: Group page category URL.
        local_html_path: Optional path to offline HTML file.
    """
    print("\n============================================================")
    print(f"Starting Furaipan Club Scraper for URL: {group_url}")
    print("============================================================\n")

    official_codes = load_official_product_codes(list_products_file)
    print(
        f"Loaded {len(official_codes)} official product codes from "
        f"'{list_products_file}'.",
        flush=True,
    )

    results = scrape_furaipan_products(
        group_url=group_url,
        official_codes=official_codes,
        local_html_path=local_html_path,
    )

    print(
        f"\nExtracted {len(results)} total items for Furaipan Club.",
        flush=True,
    )

    store_dfs: Dict[str, pd.DataFrame] = {}
    if results:
        store_dfs["Furaipan_Club"] = pd.DataFrame(results)

    print("\n--- Exporting Furaipan Results to Excel ---", flush=True)
    save_excel_with_fallback(
        all_results=results,
        store_dfs=store_dfs,
        output_excel=output_excel,
    )


if __name__ == "__main__":
    scrape_all_furaipan_products(
        output_excel=OUTPUT_FURAIPAN_SCRAPED_EXCEL,
        list_products_file=CATALOG_LIST_EXCEL,
        group_url=FURAIPAN_GROUP_URL,
    )
