"""Rakuten Store Scraper.

Reads store search URLs from Excel ('rakuten_stores.xlsx' - sheet '楽天'),
scrapes product info (Name, Code, Price, Points, URL) for each store
with pagination and filtering by 'GLOBAL' keyword or model codes,
and exports a multi-sheet Excel file ('rakuten_prices_by_store.xlsx')
with a consolidated sheet and individual store tabs.
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
    OUTPUT_SCRAPED_EXCEL,
    RAKUTEN_CARD_CLASS,
    RAKUTEN_MASTER_EXCEL,
    RAKUTEN_POINTS_CLASS,
    RAKUTEN_PRICE_CLASS,
    RAKUTEN_SHEET_NAME,
    RAKUTEN_TITLE_CLASS,
)
from utils import (
    clean_points_text,
    clean_price_text,
    extract_product_code,
    load_official_product_codes,
    save_excel_with_fallback,
)

# Configure UTF-8 encoding for Windows console output
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def load_rakuten_stores_from_excel(
    excel_path: str = RAKUTEN_MASTER_EXCEL,
    sheet_name: str = RAKUTEN_SHEET_NAME,
) -> List[Tuple[str, str, str]]:
    """Load store list and target search URLs from rakuten_stores.xlsx.

    Args:
        excel_path: Path to Rakuten master Excel file.
        sheet_name: Sheet name containing store rows.

    Returns:
        List of tuples (Store Name, Company Name, Target URL).
    """
    stores: List[Tuple[str, str, str]] = []
    try:
        df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    except Exception as exc:
        print(f"Error reading Rakuten Excel file '{excel_path}': {exc}")
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


def build_rakuten_page_url(search_url: str, page: int) -> str:
    """Build the paginated URL for Rakuten store searches.

    Handles inshop-mall path URLs and search/mall query URLs.

    Args:
        search_url: Base search URL for the Rakuten store.
        page: Page number (1-indexed).

    Returns:
        Paginated URL string.
    """
    if page == 1:
        return search_url

    m_inshop = re.search(
        r"/search/inshop-mall/([^/]+)/-/sid\.(\d+)", search_url
    )
    if m_inshop:
        keyword = m_inshop.group(1)
        sid = m_inshop.group(2)
        return (
            f"https://search.rakuten.co.jp/search/mall/{keyword}/"
            f"?p={page}&sid={sid}"
        )

    if "p=" in search_url:
        return re.sub(r"([?&])p=\d+", rf"\g<1>p={page}", search_url)

    sep = "&" if "?" in search_url else "?"
    return f"{search_url}{sep}p={page}"


def scrape_rakuten_store_products(
    store_name: str,
    company_name: str,
    search_url: str,
    official_codes: List[str],
    headers: Dict[str, str],
    max_pages_per_store: int = MAX_PAGES_PER_STORE,
) -> List[Dict[str, str]]:
    """Scrape product catalog for a single Rakuten store with pagination.

    Args:
        store_name: Name of the store.
        company_name: Name of the company.
        search_url: Base search URL for the store.
        official_codes: Official product codes catalog.
        headers: HTTP headers for web requests.
        max_pages_per_store: Max pagination depth per store.

    Returns:
        List of extracted product dictionary records.
    """
    store_results: List[Dict[str, str]] = []
    seen_urls: Set[str] = set()
    page = 1

    title_class_pattern = re.compile(RAKUTEN_TITLE_CLASS)
    price_class_pattern = re.compile(RAKUTEN_PRICE_CLASS)
    points_class_pattern = re.compile(RAKUTEN_POINTS_CLASS)
    card_class_pattern = re.compile(RAKUTEN_CARD_CLASS)

    while page <= max_pages_per_store:
        url = build_rakuten_page_url(search_url, page)

        try:
            print(
                f"  [Page {page}] Fetching... (please wait ~10s)",
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
                    time.sleep(1.5 * (attempt + 1))

            elapsed = round(time.time() - start_time, 1)

            if response is None or response.status_code != 200:
                status_msg = (
                    response.status_code if response else "No response"
                )
                print(
                    f"  [Page {page}] Status {status_msg} ({elapsed}s). Stop.",
                    flush=True,
                )
                break

            print(
                f"  [Page {page}] Response OK ({elapsed}s). Parsing...",
                flush=True,
            )

            soup = BeautifulSoup(response.text, "html.parser")
            title_elements = soup.find_all(class_=title_class_pattern)

            if not title_elements:
                print(
                    f"  [Page {page}] No items found. Stopping.",
                    flush=True,
                )
                break

            new_items = 0
            for t_el in title_elements:
                product_url = t_el.get("href", "")

                # Skip elements without href (visual image duplicates)
                if not product_url:
                    continue

                if product_url in seen_urls:
                    continue

                title_text = t_el.get_text(strip=True)

                has_global = "GLOBAL" in title_text.upper()
                product_code = extract_product_code(
                    title_text, official_codes
                )

                if not has_global and not product_code:
                    continue

                seen_urls.add(product_url)

                # Locate parent item card containing price and points
                card = (
                    t_el.find_parent(class_=card_class_pattern)
                    or t_el.find_parent("li")
                )
                if not card:
                    curr = t_el.parent
                    for _ in range(6):
                        if not curr:
                            break
                        if curr.find(class_=price_class_pattern):
                            card = curr
                            break
                        curr = curr.parent

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
                    card.find(class_=points_class_pattern)
                    if card
                    else None
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

        except Exception as exc:
            print(
                f"  [Page {page}] Unexpected error: {exc}",
                flush=True,
            )
            break

    return store_results


def scrape_all_rakuten_stores(
    excel_input: str = RAKUTEN_MASTER_EXCEL,
    output_excel: str = OUTPUT_SCRAPED_EXCEL,
    list_products_file: str = CATALOG_LIST_EXCEL,
) -> None:
    """Scrape all Rakuten stores from Excel and export tab-separated Excel.

    Args:
        excel_input: Input Rakuten master Excel file path.
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

    stores = load_rakuten_stores_from_excel(excel_input, RAKUTEN_SHEET_NAME)
    print(
        f"Found {len(stores)} Rakuten stores in '{excel_input}'.",
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

        results = scrape_rakuten_store_products(
            store_name=store_name,
            company_name=company_name,
            search_url=target_url,
            official_codes=official_codes,
            headers=headers,
        )

        print(
            f"Extracted {len(results)} items for store '{store_name}'.",
            flush=True,
        )
        all_results.extend(results)

        if results:
            df_store = pd.DataFrame(results)
            store_dfs[store_name] = df_store

    print("\n--- Exporting Rakuten Results to Excel ---", flush=True)
    save_excel_with_fallback(
        all_results=all_results,
        store_dfs=store_dfs,
        output_excel=output_excel,
    )


if __name__ == "__main__":
    scrape_all_rakuten_stores(
        excel_input=RAKUTEN_MASTER_EXCEL,
        output_excel=OUTPUT_SCRAPED_EXCEL,
        list_products_file=CATALOG_LIST_EXCEL,
    )
