"""Yodobashi Camera Product Scraper for E-Commerce Price Monitoring.

Extracts product titles, SKUs, catalog model codes, prices, points, and URLs
from Yodobashi.com search result pages for brand price compliance monitoring.
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
    CATALOG_LIST_EXCEL,
    COURTESY_PAUSE_SECONDS,
    HTTP_RETRIES,
    HTTP_TIMEOUT,
    MAX_PAGES_PER_STORE,
    OUTPUT_YODOBASHI_SCRAPED_EXCEL,
    TARGET_KEYWORD,
    YODOBASHI_SEARCH_KEYWORD,
)
from utils import (
    clean_points_text,
    clean_price_text,
    evaluate_point_status,
    extract_product_code,
    load_official_product_codes,
    save_excel_with_fallback,
)

# Configure UTF-8 encoding for Windows console output
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# Realistic browser request headers for Yodobashi Camera
DEFAULT_YODOBASHI_HEADERS: Dict[str, str] = {
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


def build_yodobashi_search_url(keyword: str, page: int = 1) -> str:
    """Construct Yodobashi.com search pagination URL.

    Args:
        keyword: Search keyword (e.g. 'global 包丁').
        page: Target page number (1-indexed).

    Returns:
        Full Yodobashi search URL with ginput and word parameters.
    """
    encoded_kw = urllib.parse.quote_plus(keyword)
    if page <= 1:
        return (
            f"https://www.yodobashi.com/?ginput={encoded_kw}&word={encoded_kw}"
        )
    return (
        f"https://www.yodobashi.com/p{page}/"
        f"?ginput={encoded_kw}&word={encoded_kw}"
    )


def parse_yodobashi_search_page(
    html_content: str,
    official_codes: Optional[List[str]] = None,
) -> List[Dict[str, str]]:
    """Parse Yodobashi.com search results HTML and extract product records.

    Filters out competitor/unrelated items not containing 'GLOBAL'
    or 'グローバル'.

    Args:
        html_content: Raw HTML text of the Yodobashi search result page.
        official_codes: Optional list of official catalog model codes.

    Returns:
        List of extracted product dictionaries.
    """
    if not html_content:
        return []

    soup = BeautifulSoup(html_content, "html.parser")
    cards = soup.find_all(
        "div", class_=re.compile(r"srcResultItem_block|js_productBox")
    )
    if not cards:
        cards = soup.find_all("div", attrs={"data-sku": True})

    results: List[Dict[str, str]] = []
    seen_skus: Set[str] = set()

    for card in cards:
        sku = card.get("data-sku", "").strip()
        if sku and sku in seen_skus:
            continue

        # Extract title from pName div or link
        pname_el = card.find("div", class_="pName")
        if pname_el:
            title = " ".join(pname_el.stripped_strings)
        else:
            title_el = card.find("a", href=re.compile(r"/product/"))
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

        if sku:
            seen_skus.add(sku)

        # Extract direct product URL
        link_el = card.find("a", href=re.compile(r"/product/"))
        raw_href = link_el.get("href", "") if link_el else ""
        if raw_href.startswith("/"):
            raw_href = f"https://www.yodobashi.com{raw_href}"
        clean_url = raw_href.split("?")[0] if "?" in raw_href else raw_href

        # Extract Price
        price_el = card.find(
            "span", class_=re.compile(r"productPrice|pPrice|price", re.I)
        )
        raw_price = (
            price_el.get_text(strip=True) if price_el else "No disponible"
        )
        price_text = clean_price_text(raw_price)

        # Extract Points
        point_el = card.find(
            "span", class_=re.compile(r"goldPoint|pPoint|point", re.I)
        )
        raw_points = point_el.get_text(strip=True) if point_el else "N/A"
        points_text = clean_points_text(raw_points)

        # Extract catalog model code
        product_code = extract_product_code(title, official_codes)
        if not product_code and "/product/" in clean_url:
            product_code = extract_product_code(clean_url, official_codes)

        # Evaluate Point Status (Rule: <= 1% is ⭕, otherwise ❌)
        point_status = evaluate_point_status(points_text)

        results.append({
            "Store": "ヨドバシ.com",
            "Company": "ヨドバシカメラ (Yodobashi Camera)",
            "SKU": sku,
            "Product": title,
            "Product_Code": (
                product_code if product_code else TARGET_KEYWORD
            ),
            "Price": price_text,
            "Points": points_text,
            "Point Status": point_status,
            "Product_URL": clean_url,
        })

    return results


def scrape_yodobashi_products(
    search_keyword: str = YODOBASHI_SEARCH_KEYWORD,
    max_pages: int = MAX_PAGES_PER_STORE,
    official_codes: Optional[List[str]] = None,
    headers: Optional[Dict[str, str]] = None,
    local_html_path: Optional[str] = None,
) -> List[Dict[str, str]]:
    """Scrape product listings live from Yodobashi.com search pagination.

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
        return parse_yodobashi_search_page(content, official_codes)

    req_headers = headers or DEFAULT_YODOBASHI_HEADERS
    all_results: List[Dict[str, str]] = []
    seen_skus: Set[str] = set()

    for page in range(1, max_pages + 1):
        target_url = build_yodobashi_search_url(search_keyword, page)
        print(
            f"  [Yodobashi Page {page}] Requesting: {target_url}",
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
                f"  [Warning] Could not retrieve Yodobashi page {page}.",
                flush=True,
            )
            break

        page_results = parse_yodobashi_search_page(html_text, official_codes)
        new_items = 0
        for r in page_results:
            sku = r.get("SKU", "")
            if sku and sku not in seen_skus:
                seen_skus.add(sku)
                all_results.append(r)
                new_items += 1
            elif not sku:
                all_results.append(r)
                new_items += 1

        print(
            f"  [Yodobashi Page {page}] Found {new_items} new items "
            f"(total: {len(all_results)}).",
            flush=True,
        )

        if new_items == 0:
            print("  No more new items found. Ending pagination.", flush=True)
            break

        # Check if next page link exists in HTML
        soup = BeautifulSoup(html_text, "html.parser")
        next_link = soup.find("a", class_="next")
        if not next_link:
            print("  Reached last page (no next link).", flush=True)
            break

        time.sleep(COURTESY_PAUSE_SECONDS)

    return all_results


def scrape_all_yodobashi_products(
    output_excel: str = OUTPUT_YODOBASHI_SCRAPED_EXCEL,
    list_products_file: str = CATALOG_LIST_EXCEL,
    search_keyword: str = YODOBASHI_SEARCH_KEYWORD,
    local_html_path: Optional[str] = None,
) -> None:
    """Scrape Yodobashi Camera products and save multi-sheet Excel file.

    Args:
        output_excel: Destination Excel file path.
        list_products_file: Official product catalog Excel file.
        search_keyword: Keyword to search on Yodobashi.
        local_html_path: Optional path to offline HTML file.
    """
    print("\n============================================================")
    print(f"Starting Yodobashi Scraper for keyword: '{search_keyword}'")
    print("============================================================\n")

    official_codes = load_official_product_codes(list_products_file)
    print(
        f"Loaded {len(official_codes)} official product codes from "
        f"'{list_products_file}'.",
        flush=True,
    )

    results = scrape_yodobashi_products(
        search_keyword=search_keyword,
        official_codes=official_codes,
        local_html_path=local_html_path,
    )

    print(
        f"\nExtracted {len(results)} total items for Yodobashi Camera.",
        flush=True,
    )

    store_dfs: Dict[str, pd.DataFrame] = {}
    if results:
        store_dfs["Yodobashi_Camera"] = pd.DataFrame(results)

    print("\n--- Exporting Yodobashi Results to Excel ---", flush=True)
    save_excel_with_fallback(
        all_results=results,
        store_dfs=store_dfs,
        output_excel=output_excel,
    )


if __name__ == "__main__":
    scrape_all_yodobashi_products(
        output_excel=OUTPUT_YODOBASHI_SCRAPED_EXCEL,
        list_products_file=CATALOG_LIST_EXCEL,
        search_keyword=YODOBASHI_SEARCH_KEYWORD,
    )
