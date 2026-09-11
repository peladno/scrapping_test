"""Import Shop Aqua Scraper for E-Commerce Price Monitoring.

Extracts product titles, IDs, catalog model codes, prices, points, and URLs
from importshopaqua.com search result pages for price compliance monitoring.
"""

import os
import re
import sys
import time
import urllib.parse
from typing import Dict, List, Optional, Set

from bs4 import BeautifulSoup
from curl_cffi import requests as cffi_requests
import pandas as pd

from config import (
    AQUA_BASE_URL,
    AQUA_CARD_CLASS,
    AQUA_PRICE_CLASS,
    AQUA_SEARCH_KEYWORD,
    AQUA_SEARCH_URL,
    AQUA_TITLE_CLASS,
    CATALOG_LIST_EXCEL,
    COURTESY_PAUSE_SECONDS,
    HTTP_RETRIES,
    HTTP_TIMEOUT,
    MAX_PAGES_PER_STORE,
    OUTPUT_AQUA_SCRAPED_EXCEL,
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

# Realistic browser request headers for Import Shop Aqua
DEFAULT_AQUA_HEADERS: Dict[str, str] = {
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


def build_aqua_search_url(
    keyword: str,
    page: int = 1,
    sort: str = "keyword",
) -> str:
    """Construct Import Shop Aqua search pagination URL.

    Args:
        keyword: Search keyword (e.g. 'グローバル').
        page: Target page number (1-indexed).
        sort: Search sorting method (default: 'keyword').

    Returns:
        Full Import Shop Aqua search URL.
    """
    encoded_kw = urllib.parse.quote(keyword)
    return (
        f"{AQUA_SEARCH_URL}"
        f"?keyword={encoded_kw}&page={page}&sort={sort}"
    )


def parse_aqua_search_page(
    html_content: str,
    official_codes: Optional[List[str]] = None,
) -> List[Dict[str, str]]:
    """Parse Import Shop Aqua search results HTML and extract product records.

    Filters out items not containing 'GLOBAL' or 'グローバル'.

    Args:
        html_content: Raw HTML text of the search results page.
        official_codes: Optional list of official catalog model codes.

    Returns:
        List of extracted product dictionaries.
    """
    if not html_content:
        return []

    soup = BeautifulSoup(html_content, "html.parser")
    cards = soup.find_all("article", class_=re.compile(AQUA_CARD_CLASS))
    if not cards:
        cards = soup.find_all(
            "div", class_=re.compile(r"fs-c-productListItem")
        )

    results: List[Dict[str, str]] = []
    seen_ids: Set[str] = set()

    for card in cards:
        pid = card.get("data-product-id", "").strip()

        # Extract product title from productName element
        title_el = card.find(class_=re.compile(AQUA_TITLE_CLASS))
        if not title_el:
            pname_box = card.find(
                class_=re.compile(r"fs-c-productName|productName")
            )
            title_el = pname_box.find("a") if pname_box else card.find("a")

        title = title_el.get_text(strip=True) if title_el else ""
        if not title:
            continue

        # Strict Filter: Omit items if title does not contain
        # 'GLOBAL' or 'グローバル'
        has_target = (
            "GLOBAL" in title.upper() or "グローバル" in title
        )
        if not has_target:
            continue

        # Extract product URL
        link_el = card.find("a", href=True)
        raw_href = link_el.get("href", "") if link_el else ""
        if raw_href.startswith("/"):
            raw_href = f"{AQUA_BASE_URL.rstrip('/')}{raw_href}"
        clean_url = raw_href.split("?")[0] if "?" in raw_href else raw_href

        # Avoid duplicate product IDs on same page
        unique_key = pid if pid else clean_url
        if unique_key and unique_key in seen_ids:
            continue
        if unique_key:
            seen_ids.add(unique_key)

        # Extract price
        price_el = card.find(class_=re.compile(AQUA_PRICE_CLASS))
        raw_price = price_el.get_text(strip=True) if price_el else ""
        price_text = clean_price_text(raw_price)

        # Extract official catalog code
        product_code = extract_product_code(title, official_codes)
        if not product_code and "/c/item/" in clean_url:
            product_code = extract_product_code(clean_url, official_codes)

        results.append({
            "Store": "importshopaqua",
            "Company": "セレクトショップAQUA (Import Shop Aqua)",
            "SKU": pid,
            "Product": title,
            "Product_Code": (
                product_code if product_code else TARGET_KEYWORD
            ),
            "Price": price_text,
            "Product_URL": clean_url,
        })

    return results


def scrape_aqua_products(
    search_keyword: str = AQUA_SEARCH_KEYWORD,
    max_pages: int = MAX_PAGES_PER_STORE,
    official_codes: Optional[List[str]] = None,
    headers: Optional[Dict[str, str]] = None,
    local_html_path: Optional[str] = None,
) -> List[Dict[str, str]]:
    """Scrape product listings live from Import Shop Aqua search pagination.

    Args:
        search_keyword: Search term to query.
        max_pages: Maximum pages to paginate.
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
        return parse_aqua_search_page(content, official_codes)

    req_headers = headers or DEFAULT_AQUA_HEADERS
    all_results: List[Dict[str, str]] = []
    seen_urls: Set[str] = set()

    for page in range(1, max_pages + 1):
        target_url = build_aqua_search_url(search_keyword, page)
        print(
            f"  [Aqua Page {page}] Requesting: {target_url}",
            flush=True,
        )

        html_text = ""
        for attempt in range(HTTP_RETRIES):
            try:
                response = cffi_requests.get(
                    target_url,
                    headers=req_headers,
                    impersonate="chrome120",
                    timeout=HTTP_TIMEOUT,
                )
                if response.status_code == 200:
                    html_text = response.text
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

        if not html_text:
            print(
                f"  [Warning] Could not retrieve Aqua page {page}.",
                flush=True,
            )
            break

        page_results = parse_aqua_search_page(html_text, official_codes)
        new_items = 0
        for r in page_results:
            p_url = r.get("Product_URL", "")
            if p_url and p_url not in seen_urls:
                seen_urls.add(p_url)
                all_results.append(r)
                new_items += 1
            elif not p_url:
                all_results.append(r)
                new_items += 1

        print(
            f"  [Aqua Page {page}] Found {new_items} new items "
            f"(total: {len(all_results)}).",
            flush=True,
        )

        if new_items == 0:
            print("  No more new items found. Ending pagination.", flush=True)
            break

        # Check if next page button exists in HTML
        soup = BeautifulSoup(html_text, "html.parser")
        next_btn = soup.find(
            "a",
            class_=lambda c: c and "fs-c-pagination__item--next" in c,
        )
        if not next_btn:
            print("  Reached last page (no next link).", flush=True)
            break

        time.sleep(COURTESY_PAUSE_SECONDS)

    return all_results


def scrape_all_aqua_products(
    output_excel: str = OUTPUT_AQUA_SCRAPED_EXCEL,
    list_products_file: str = CATALOG_LIST_EXCEL,
    search_keyword: str = AQUA_SEARCH_KEYWORD,
    local_html_path: Optional[str] = None,
) -> None:
    """Scrape Import Shop Aqua products and save multi-sheet Excel file.

    Args:
        output_excel: Destination Excel file path.
        list_products_file: Official product catalog Excel file.
        search_keyword: Keyword to search on Import Shop Aqua.
        local_html_path: Optional path to offline HTML file.
    """
    print("\n============================================================")
    print(f"Starting Import Shop Aqua Scraper for keyword: '{search_keyword}'")
    print("============================================================\n")

    official_codes = load_official_product_codes(list_products_file)
    print(
        f"Loaded {len(official_codes)} official product codes from "
        f"'{list_products_file}'.",
        flush=True,
    )

    results = scrape_aqua_products(
        search_keyword=search_keyword,
        official_codes=official_codes,
        local_html_path=local_html_path,
    )

    print(
        f"\nExtracted {len(results)} total items for Import Shop Aqua.",
        flush=True,
    )

    store_dfs: Dict[str, pd.DataFrame] = {}
    if results:
        store_dfs["Import_Shop_Aqua"] = pd.DataFrame(results)

    print("\n--- Exporting Aqua Results to Excel ---", flush=True)
    save_excel_with_fallback(
        all_results=results,
        store_dfs=store_dfs,
        output_excel=output_excel,
    )


if __name__ == "__main__":
    scrape_all_aqua_products(
        output_excel=OUTPUT_AQUA_SCRAPED_EXCEL,
        list_products_file=CATALOG_LIST_EXCEL,
        search_keyword=AQUA_SEARCH_KEYWORD,
    )
