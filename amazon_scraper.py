"""Amazon Japan Product Scraper for E-Commerce Price Monitoring.

Extracts product titles, ASINs, catalog model codes, prices, points, and URLs
from Amazon Japan search result pages for brand price compliance monitoring.
"""

import os
import re
import sys
import time
import urllib.parse
from typing import Dict, List, Optional, Set

from bs4 import BeautifulSoup
import pandas as pd
import requests

from config import (
    AMAZON_SEARCH_KEYWORD,
    CATALOG_LIST_EXCEL,
    COURTESY_PAUSE_SECONDS,
    HTTP_RETRIES,
    HTTP_TIMEOUT,
    MAX_PAGES_PER_STORE,
    OUTPUT_AMAZON_SCRAPED_EXCEL,
    TARGET_KEYWORD,
)
from utils import (
    clean_points_text,
    clean_price_text,
    evaluate_amazon_point_status,
    extract_product_code,
    load_official_product_codes,
    save_excel_with_fallback,
)

# Configure UTF-8 encoding for Windows console output
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# Realistic browser request headers for Amazon Japan anti-bot resilience
DEFAULT_AMAZON_HEADERS: Dict[str, str] = {
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
    "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
}


def build_amazon_search_url(keyword: str, page: int = 1) -> str:
    """Construct Amazon Japan search pagination URL.

    Args:
        keyword: Search term (e.g. 'global 包丁').
        page: Target page number (1-indexed).

    Returns:
        Full Amazon Japan search URL with pagination.
    """
    encoded_kw = urllib.parse.quote_plus(keyword)
    if page <= 1:
        return f"https://www.amazon.co.jp/s?k={encoded_kw}&ref=sr_pg_1"
    return (
        f"https://www.amazon.co.jp/s?k={encoded_kw}"
        f"&page={page}&ref=sr_pg_{page}"
    )


def parse_amazon_search_page(
    html_content: str,
    official_codes: Optional[List[str]] = None,
) -> List[Dict[str, str]]:
    """Parse Amazon Japan search results page HTML and extract product records.

    Filters out competitor/sponsored ads not belonging to TARGET_KEYWORD
    (e.g., Zwilling, Kai Sekimagoroku).

    Args:
        html_content: Raw HTML text of the Amazon search result page.
        official_codes: Optional list of official catalog model codes.

    Returns:
        List of extracted product dictionaries.
    """
    if not html_content:
        return []

    soup = BeautifulSoup(html_content, "html.parser")
    items = soup.find_all(
        "div", attrs={"data-component-type": "s-search-result"}
    )
    if not items:
        items = soup.find_all("div", attrs={"data-asin": True})

    results: List[Dict[str, str]] = []
    seen_asins: Set[str] = set()

    for item in items:
        asin = item.get("data-asin", "").strip()
        if not asin or asin in seen_asins:
            continue

        # Extract title
        title_link = item.find("a", class_=re.compile(r"a-text-normal"))
        title = title_link.get_text(strip=True) if title_link else ""
        if not title:
            h2_tags = item.find_all("h2")
            if len(h2_tags) > 1:
                title = h2_tags[1].get_text(strip=True)
            elif h2_tags:
                title = h2_tags[0].get_text(strip=True)

        # Extract brand tag if present
        brand_el = item.find("h2", class_=re.compile(r"s-line-clamp-1"))
        brand = brand_el.get_text(strip=True) if brand_el else ""

        if brand and brand.lower() not in title.lower():
            full_title = f"{brand} {title}".strip()
        else:
            full_title = title

        # Extract direct URL
        raw_href = title_link.get("href", "") if title_link else ""
        if raw_href.startswith("/"):
            raw_href = f"https://www.amazon.co.jp{raw_href}"
        clean_url = raw_href.split("?")[0] if "?" in raw_href else raw_href

        # Strict Filter: Omit items if title does not contain
        # 'GLOBAL' or 'グローバル'
        has_target = (
            "GLOBAL" in full_title.upper() or "グローバル" in full_title
        )
        if not has_target:
            continue

        product_code = extract_product_code(full_title, official_codes)
        if not product_code and "/dp/" in clean_url:
            product_code = extract_product_code(clean_url, official_codes)

        seen_asins.add(asin)

        # Extract Price
        price_el = item.find("span", class_="a-price")
        raw_price = "No disponible"
        if price_el:
            offscreen = price_el.find("span", class_="a-offscreen")
            if offscreen:
                raw_price = offscreen.get_text(strip=True)
            else:
                raw_price = price_el.get_text(strip=True)
        price_text = clean_price_text(raw_price)

        # Extract Points
        points_el = item.find("span", class_=re.compile(r"a-color-price"))
        raw_points = points_el.get_text(strip=True) if points_el else "N/A"
        points_text = clean_points_text(raw_points)

        results.append({
            "Store": "Amazon.co.jp",
            "Company": "Amazon Japan",
            "ASIN": asin,
            "Product": full_title,
            "Product_Code": (
                product_code if product_code else TARGET_KEYWORD
            ),
            "Price": price_text,
            "Points": points_text,
            "points status": evaluate_amazon_point_status(points_text),
            "Product_URL": clean_url,
        })

    return results


def scrape_amazon_products(
    search_keyword: str = AMAZON_SEARCH_KEYWORD,
    max_pages: int = MAX_PAGES_PER_STORE,
    official_codes: Optional[List[str]] = None,
    headers: Optional[Dict[str, str]] = None,
    local_html_path: Optional[str] = None,
) -> List[Dict[str, str]]:
    """Scrape product listings from Amazon Japan search pages.

    Args:
        search_keyword: Search term to query on Amazon.
        max_pages: Maximum pages to paginate.
        official_codes: Catalog model codes list.
        headers: Optional HTTP headers dictionary.
        local_html_path: Optional local HTML file path for testing or
            offline execution.

    Returns:
        Consolidated list of extracted product records.
    """
    # If a local HTML file exists or is specified, parse it first
    if local_html_path and os.path.exists(local_html_path):
        print(
            f"Reading from local HTML file '{local_html_path}'...",
            flush=True,
        )
        with open(
            local_html_path, "r", encoding="utf-8", errors="ignore"
        ) as f:
            content = f.read()
        return parse_amazon_search_page(content, official_codes)

    req_headers = headers or DEFAULT_AMAZON_HEADERS
    all_results: List[Dict[str, str]] = []
    seen_asins: Set[str] = set()

    session = requests.Session()

    for page in range(1, max_pages + 1):
        target_url = build_amazon_search_url(search_keyword, page)
        print(f"  [Amazon Page {page}] Requesting: {target_url}", flush=True)

        html_text = ""
        for attempt in range(HTTP_RETRIES):
            try:
                response = session.get(
                    target_url, headers=req_headers, timeout=HTTP_TIMEOUT
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
            except requests.RequestException as err:
                print(f"  [Attempt {attempt + 1}] Error: {err}", flush=True)
                time.sleep(COURTESY_PAUSE_SECONDS)

        if not html_text:
            print(f"  [Warning] Could not retrieve page {page}.", flush=True)
            break

        # Check if Amazon served a CAPTCHA challenge page
        is_captcha = (
            "api-services-support@amazon.com" in html_text
            or "Type the characters" in html_text
        )
        if is_captcha:
            print(
                "  [Anti-bot] Amazon displayed a CAPTCHA challenge. "
                "Stopping live pagination.",
                flush=True,
            )
            break

        page_results = parse_amazon_search_page(html_text, official_codes)
        new_items = 0
        for r in page_results:
            asin = r.get("ASIN", "")
            if asin and asin not in seen_asins:
                seen_asins.add(asin)
                all_results.append(r)
                new_items += 1

        print(
            f"  [Amazon Page {page}] Found {new_items} new items "
            f"(total: {len(all_results)}).",
            flush=True,
        )

        if new_items == 0:
            print("  No more new items found. Ending pagination.", flush=True)
            break

        time.sleep(COURTESY_PAUSE_SECONDS)

    return all_results


def scrape_all_amazon_products(
    output_excel: str = OUTPUT_AMAZON_SCRAPED_EXCEL,
    list_products_file: str = CATALOG_LIST_EXCEL,
    search_keyword: str = AMAZON_SEARCH_KEYWORD,
    local_html_path: Optional[str] = None,
) -> None:
    """Scrape Amazon Japan products and save multi-sheet Excel file.

    Args:
        output_excel: Destination Excel file path.
        list_products_file: Official product catalog Excel file.
        search_keyword: Keyword to search on Amazon.
        local_html_path: Optional path to offline HTML file.
    """
    print("\n============================================================")
    print(f"Starting Amazon Japan Scraper for keyword: '{search_keyword}'")
    print("============================================================\n")

    official_codes = load_official_product_codes(list_products_file)
    print(
        f"Loaded {len(official_codes)} official product codes from "
        f"'{list_products_file}'.",
        flush=True,
    )

    results = scrape_amazon_products(
        search_keyword=search_keyword,
        official_codes=official_codes,
        local_html_path=local_html_path,
    )

    print(
        f"\nExtracted {len(results)} total items for Amazon Japan.",
        flush=True,
    )

    store_dfs: Dict[str, pd.DataFrame] = {}
    if results:
        store_dfs["Amazon_Japan"] = pd.DataFrame(results)

    print("\n--- Exporting Amazon Results to Excel ---", flush=True)
    save_excel_with_fallback(
        all_results=results,
        store_dfs=store_dfs,
        output_excel=output_excel,
    )


if __name__ == "__main__":
    scrape_all_amazon_products(
        output_excel=OUTPUT_AMAZON_SCRAPED_EXCEL,
        list_products_file=CATALOG_LIST_EXCEL,
        search_keyword=AMAZON_SEARCH_KEYWORD,
    )
