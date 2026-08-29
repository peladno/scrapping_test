"""Yahoo Store Scraper.

Reads store search URLs from master Excel configured in environment,
scrapes product info (Name, Code, Price, Points, URL) for each store
with pagination and filtering by 'GLOBAL' keyword or model codes,
and exports a multi-sheet Excel file with a consolidated sheet and
individual store tabs.
"""

import re
import sys
import time
from typing import Dict, List, Set, Tuple

from bs4 import BeautifulSoup
import pandas as pd
import requests

from config import (
    CATALOG_LIST_EXCEL,
    COURTESY_PAUSE_SECONDS,
    HTTP_RETRIES,
    HTTP_TIMEOUT,
    MAX_PAGES_PER_STORE,
    OUTPUT_YAHOO_SCRAPED_EXCEL,
    YAHOO_BRAND_CLASS,
    YAHOO_CARD_CLASS,
    YAHOO_DETAIL_LINK_CLASS,
    YAHOO_MASTER_EXCEL,
    YAHOO_POINTS_CLASS,
    YAHOO_PRICE_CLASS,
    YAHOO_SHEET_NAME,
    YAHOO_TITLE_CLASS,
)
from utils import (
    clean_points_text,
    clean_price_text,
    clean_product_url,
    extract_product_code,
    load_official_product_codes,
    save_excel_with_fallback,
)

# Configure UTF-8 encoding for Windows console output
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def load_yahoo_stores_from_excel(
    excel_path: str = YAHOO_MASTER_EXCEL,
    sheet_name: str = YAHOO_SHEET_NAME,
) -> List[Tuple[str, str, str]]:
    """Load store list and target search URLs from master Excel file.

    Args:
        excel_path: Path to master Excel file.
        sheet_name: Sheet name containing store rows.

    Returns:
        List of tuples (Store Name, Company Name, Target URL).
    """
    stores: List[Tuple[str, str, str]] = []
    try:
        df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    except Exception as exc:
        print(f"Error reading Yahoo Excel file '{excel_path}': {exc}")
        return stores

    for idx in range(len(df_raw)):
        row = df_raw.iloc[idx]
        row_vals = [str(v) for v in row.values if pd.notna(v)]

        row_text = (
            " ".join(row_vals).replace("\r", " ").replace("\n", " ")
        )
        raw_urls = re.findall(r"https?://[^\s\"'<>\\`]+", row_text)

        clean_urls: List[str] = []
        for u in raw_urls:
            u_clean = u.strip().rstrip(")\"'>,;.]}")
            if u_clean.startswith("http://"):
                u_clean = "https://" + u_clean[7:]
            if u_clean and u_clean not in clean_urls:
                clean_urls.append(u_clean)

        if clean_urls:
            store_val = row[1] if len(row) > 1 and pd.notna(row[1]) else None
            company_val = row[2] if len(row) > 2 and pd.notna(row[2]) else ""

            store_name = (
                str(store_val).strip() if store_val else f"Store_{idx}"
            )
            company_name = str(company_val).strip() if company_val else ""

            target_url = clean_urls[0]
            stores.append((store_name, company_name, target_url))

    return stores


def build_yahoo_page_url(search_url: str, page: int) -> str:
    """Build the paginated URL for Yahoo Shopping.

    Handles general search paths (/0/2/, /0/3/) and store URLs (&page=N).

    Args:
        search_url: Original store or general search URL.
        page: Page number (1-based index).

    Returns:
        Paginated URL string.
    """
    if page == 1:
        return search_url

    if "shopping.yahoo.co.jp/search" in search_url:
        if re.search(r"/0/\d*/", search_url):
            return re.sub(r"/0/\d*/", f"/0/{page}/", search_url)
        if "/0/?" in search_url:
            return search_url.replace("/0/?", f"/0/{page}/?")
        b_offset = (page - 1) * 30 + 1
        if "b=" in search_url:
            return re.sub(r"b=\d+", f"b={b_offset}", search_url)
        sep = "&" if "?" in search_url else "?"
        return f"{search_url}{sep}b={b_offset}"

    if "store.shopping.yahoo.co.jp" in search_url:
        if "page=" in search_url:
            return re.sub(r"page=\d+", f"page={page}", search_url)
        sep = "&" if "?" in search_url else "?"
        return f"{search_url}{sep}page={page}"

    return search_url


def scrape_yahoo_store_products(
    store_name: str,
    company_name: str,
    search_url: str,
    official_codes: List[str],
    headers: Dict[str, str],
    max_pages_per_store: int = MAX_PAGES_PER_STORE,
) -> List[Dict[str, str]]:
    """Scrape product catalog for a single Yahoo store with pagination.

    Iterates over detailLink tags to inspect self-contained product cards.

    Args:
        store_name: Name of the store.
        company_name: Name of the company.
        search_url: Base search URL for the Yahoo store.
        official_codes: Official product codes catalog.
        headers: HTTP headers for web requests.
        max_pages_per_store: Max pagination depth per store.

    Returns:
        List of extracted product dictionary records.
    """
    store_results: List[Dict[str, str]] = []
    seen_keys: Set[str] = set()
    page = 1

    detail_link_pattern = re.compile(YAHOO_DETAIL_LINK_CLASS)
    title_class_pattern = re.compile(YAHOO_TITLE_CLASS)
    brand_class_pattern = re.compile(YAHOO_BRAND_CLASS)
    price_class_pattern = re.compile(YAHOO_PRICE_CLASS)
    points_class_pattern = re.compile(YAHOO_POINTS_CLASS)
    card_class_pattern = re.compile(YAHOO_CARD_CLASS)

    while page <= max_pages_per_store:
        url = build_yahoo_page_url(search_url, page)

        try:
            print(
                f"  [Page {page}] Fetching... ",
                end="",
                flush=True,
            )
            start_time = time.time()

            response = None
            for attempt in range(HTTP_RETRIES):
                try:
                    response = requests.get(
                        url, headers=headers, timeout=HTTP_TIMEOUT
                    )
                    if response.status_code == 200:
                        break
                except requests.RequestException:
                    time.sleep(2 * (attempt + 1))

            elapsed = round(time.time() - start_time, 1)

            if response is None or response.status_code != 200:
                status_msg = (
                    response.status_code if response else "No response"
                )
                print(
                    f"Status {status_msg} ({elapsed}s). Stop.",
                    flush=True,
                )
                break

            print(
                f"OK ({elapsed}s). Parsing...",
                flush=True,
            )

            soup = BeautifulSoup(response.text, "html.parser")
            detail_links = soup.find_all("a", class_=detail_link_pattern)

            if not detail_links:
                print(
                    f"  [Page {page}] No item cards found. Stopping.",
                    flush=True,
                )
                break

            new_items = 0
            for link_el in detail_links:
                raw_href = link_el.get("href", "")
                product_url = clean_product_url(raw_href)

                if not product_url:
                    continue

                card = (
                    link_el.find_parent(class_=card_class_pattern)
                    or link_el.find_parent("li")
                    or link_el.parent
                )

                title_el = (
                    card.find(class_=title_class_pattern) if card else None
                )
                title_text = (
                    title_el.get_text(strip=True)
                    if title_el
                    else link_el.get_text(strip=True)
                )

                brand_text = ""
                if card:
                    brand_el = card.find(class_=brand_class_pattern)
                    if brand_el:
                        brand_text = brand_el.get_text(strip=True)

                has_global = (
                    "GLOBAL" in title_text.upper()
                    or "GLOBAL" in brand_text.upper()
                )
                product_code = extract_product_code(
                    title_text, official_codes
                )
                if not product_code and brand_text:
                    product_code = extract_product_code(
                        brand_text, official_codes
                    )

                if not has_global and not product_code:
                    continue

                dedup_key = f"{title_text}_{product_url}"
                if dedup_key in seen_keys:
                    continue
                seen_keys.add(dedup_key)

                price_el = (
                    card.find(class_=price_class_pattern) if card else None
                )
                raw_price = (
                    price_el.get_text(strip=True)
                    if price_el
                    else "No disponible"
                )
                price_text = clean_price_text(raw_price)

                points_el = (
                    card.find(class_=points_class_pattern) if card else None
                )
                raw_points = (
                    points_el.get_text(strip=True)
                    if points_el
                    else "No disponible"
                )
                points_text = clean_points_text(raw_points)

                store_results.append(
                    {
                        "Store": store_name,
                        "Company": company_name,
                        "Product": title_text,
                        "Product_Code": (
                            product_code if product_code else "GLOBAL"
                        ),
                        "Price": price_text,
                        "Points": points_text,
                        "Product_URL": product_url,
                    }
                )
                new_items += 1

            print(
                f"  [Page {page}] Found {new_items} new items "
                f"(total: {len(store_results)}).",
                flush=True,
            )

            if new_items == 0:
                print(
                    f"  [Page {page}] End of catalog. Stopping.",
                    flush=True,
                )
                break

            time.sleep(COURTESY_PAUSE_SECONDS)
            page += 1

            if "shopping.yahoo.co.jp" not in search_url:
                break

        except Exception as exc:
            print(
                f"  [Page {page}] Unexpected error: {exc}",
                flush=True,
            )
            break

    return store_results


def scrape_all_yahoo_stores(
    excel_input: str = YAHOO_MASTER_EXCEL,
    output_excel: str = OUTPUT_YAHOO_SCRAPED_EXCEL,
    list_products_file: str = CATALOG_LIST_EXCEL,
) -> None:
    """Scrape all Yahoo stores from Excel and export tab-separated Excel.

    Args:
        excel_input: Input Yahoo master Excel file path.
        output_excel: Output Excel file path.
        list_products_file: Catalog list Excel file path.
    """
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    }

    official_codes = load_official_product_codes(list_products_file)
    print(
        f"Loaded {len(official_codes)} official product codes from "
        f"'{list_products_file}'.",
        flush=True,
    )

    stores = load_yahoo_stores_from_excel(excel_input, YAHOO_SHEET_NAME)
    print(
        f"Found {len(stores)} Yahoo stores in '{excel_input}'.",
        flush=True,
    )

    all_results: List[Dict[str, str]] = []
    store_dfs: Dict[str, pd.DataFrame] = {}

    for idx, (store_name, company_name, target_url) in enumerate(
        stores, start=1
    ):
        print(
            f"\n--- Store [{idx}/{len(stores)}]: {store_name} ---",
            flush=True,
        )
        print(f"URL: {target_url}", flush=True)

        results = scrape_yahoo_store_products(
            store_name=store_name,
            company_name=company_name,
            search_url=target_url,
            official_codes=official_codes,
            headers=headers,
        )

        print(
            f"Extracted {len(results)} items for Yahoo store '{store_name}'.",
            flush=True,
        )
        all_results.extend(results)

        if results:
            df_store = pd.DataFrame(results)
            store_dfs[store_name] = df_store

    print("\n--- Exporting Yahoo Results to Excel ---", flush=True)
    save_excel_with_fallback(
        all_results=all_results,
        store_dfs=store_dfs,
        output_excel=output_excel,
        fallback_name="yahoo_prices_by_store_updated.xlsx",
    )


if __name__ == "__main__":
    scrape_all_yahoo_stores(
        excel_input=YAHOO_MASTER_EXCEL,
        output_excel=OUTPUT_YAHOO_SCRAPED_EXCEL,
        list_products_file=CATALOG_LIST_EXCEL,
    )
