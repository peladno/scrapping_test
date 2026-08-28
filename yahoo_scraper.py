"""Yahoo Store Scraper.

Reads store search URLs from Excel ('yahoo_stores.xlsx' - sheet 'Yahoo'),
scrapes product info (Name, Code, Price, Points, URL) for each store
with pagination and filtering by 'GLOBAL' keyword or model codes,
and exports a multi-sheet Excel file ('yahoo_prices_by_store.xlsx')
with a consolidated sheet and individual store tabs.

CSS Selectors (prefix regex matching without trailing dynamic hashes):
- Title: ItemTitle_SearchResultItemTitle (supports
  'ItemTitle_SearchResultItemTitle__fy4bB LineClamp')
- Price: ItemPrice_ItemPrice (supports 'ItemPrice_ItemPrice__2t7fx')
- Points: ItemPointModal / PointText / PointRate
"""

import re
import sys
import time
from typing import Dict, List, Optional, Set, Tuple

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
    YAHOO_MASTER_EXCEL,
    YAHOO_SHEET_NAME,
)

# Configure UTF-8 encoding for Windows console output
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def load_official_product_codes(
    list_products_file: str = CATALOG_LIST_EXCEL,
) -> List[str]:
    """Load official product model codes from list-products.xlsx.

    Args:
        list_products_file: Path to catalog Excel file.

    Returns:
        List of official product codes sorted by length descending.
    """
    official_codes: List[str] = []
    try:
        df_list = pd.read_excel(list_products_file, header=None)
        for idx in range(len(df_list)):
            row_vals = [
                str(v) for v in df_list.iloc[idx].values if pd.notna(v)
            ]
            if any("型番" in v for v in row_vals):
                for r in range(idx + 1, len(df_list)):
                    val = df_list.iloc[r, 1]
                    if pd.notna(val):
                        code_str = str(val).strip()
                        if code_str and code_str.lower() != "nan":
                            official_codes.append(code_str)
                break
    except Exception as exc:
        print(
            f"Note: Could not load '{list_products_file}' ({exc}). "
            "Falling back to regex prefix matching.",
            flush=True,
        )

    return sorted(official_codes, key=len, reverse=True)


def extract_product_code(
    title: str, official_codes: Optional[List[str]] = None
) -> Optional[str]:
    """Extract product code from title using catalog list or regex prefixes.

    Args:
        title: Product title text.
        official_codes: Optional list of catalog codes.

    Returns:
        Product code string if found, otherwise None.
    """
    if official_codes:
        for code in official_codes:
            pattern = r"\b" + re.escape(code) + r"\b"
            if re.search(pattern, title, re.IGNORECASE):
                return code
            if code in title:
                return code

    prefixes = (
        r"(?:GKW|GST|GSS|GCB|GTF|GTJ|SST|GS|GT|G)"
        r"-[A-Za-z0-9]+(?:/[A-Za-z0-9]+)?"
    )
    match = re.search(prefixes, title, re.IGNORECASE)
    if match:
        return match.group(0).upper()

    return None


def clean_price_text(raw_price: str) -> str:
    """Format raw price string into a clean Yen currency representation.

    Args:
        raw_price: Raw text extracted from price HTML element.

    Returns:
        Formatted price string (e.g., '¥12,100' or '¥9,900+').
    """
    if not raw_price or raw_price == "No disponible":
        return "N/A"

    price_clean = raw_price.replace("円", "").replace("〜", "+").strip()
    match = re.search(r"[\d,]+", price_clean)
    if match:
        number = match.group(0)
        has_plus = "+" in price_clean or "~" in raw_price or "〜" in raw_price
        return f"¥{number}{'+' if has_plus else ''}"

    return raw_price


def clean_points_text(raw_points: str) -> str:
    """Clean points/percentage information text extracted from Yahoo HTML.

    Args:
        raw_points: Raw points text from HTML element.

    Returns:
        Cleaned points or percentage string.
    """
    if not raw_points or raw_points == "No disponible":
        return "N/A"

    points_clean = raw_points.strip()
    return points_clean if points_clean else "N/A"


def sanitize_sheet_name(name: str) -> str:
    """Sanitize Excel worksheet tab name to max 31 valid characters.

    Args:
        name: Desired sheet name string.

    Returns:
        Sanitized valid Excel worksheet name.
    """
    clean = re.sub(r"[\\/*?:\[\]]", "_", name).strip()
    return clean[:31] if clean else "Store"


def load_yahoo_stores_from_excel(
    excel_path: str = YAHOO_MASTER_EXCEL,
    sheet_name: str = YAHOO_SHEET_NAME,
) -> List[Tuple[str, str, str]]:
    """Load store list and target search URLs from yahoo_stores.xlsx.

    Args:
        excel_path: Path to Yahoo master Excel file.
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


def scrape_yahoo_store_products(
    store_name: str,
    company_name: str,
    search_url: str,
    official_codes: List[str],
    headers: Dict[str, str],
    max_pages_per_store: int = MAX_PAGES_PER_STORE,
) -> List[Dict[str, str]]:
    """Scrape product catalog for a single Yahoo store with pagination.

    Uses prefix regex matching for CSS classes:
    - Title: ItemTitle_SearchResultItemTitle or ItemTitle
    - Price: ItemPrice_ItemPrice or ItemPrice
    - Points: ItemPointModal / PointText / PointRate

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
    seen_urls: Set[str] = set()
    page = 1

    title_class_pattern = re.compile(
        r"ItemTitle_SearchResultItemTitle|ItemTitle"
    )
    brand_class_pattern = re.compile(
        r"ItemBrand_SearchResultItemBrand"
    )
    price_class_pattern = re.compile(r"ItemPrice_ItemPrice|ItemPrice")
    points_class_pattern = re.compile(
        r"ItemPointModal|PointText|PointRate"
    )

    while page <= max_pages_per_store:
        if page == 1:
            url = search_url
        elif "shopping.yahoo.co.jp/search" in search_url:
            b_offset = (page - 1) * 30 + 1
            if "b=" in search_url:
                url = re.sub(r"b=\d+", f"b={b_offset}", search_url)
            else:
                sep = "&" if "?" in search_url else "?"
                url = f"{search_url}{sep}b={b_offset}"
        elif "store.shopping.yahoo.co.jp" in search_url:
            if "page=" in search_url:
                url = re.sub(r"page=\d+", f"page={page}", search_url)
            else:
                sep = "&" if "?" in search_url else "?"
                url = f"{search_url}{sep}page={page}"
        else:
            url = search_url

        try:
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

            if response is None or response.status_code != 200:
                status_msg = (
                    response.status_code if response else "No response"
                )
                print(
                    f"  [Page {page}] Status {status_msg}. Stop.",
                    flush=True,
                )
                break

            soup = BeautifulSoup(response.text, "html.parser")
            title_elements = soup.find_all(class_=title_class_pattern)

            if not title_elements:
                print(
                    f"  [Page {page}] No item titles found. Stopping.",
                    flush=True,
                )
                break

            new_items = 0
            for t_el in title_elements:
                title_text = t_el.get_text(strip=True)

                parent = (
                    t_el.find_parent("li")
                    or t_el.find_parent(
                        class_=re.compile(r"SearchResultItem")
                    )
                    or t_el.parent.parent.parent
                )

                brand_text = ""
                if parent:
                    brand_el = parent.find(class_=brand_class_pattern)
                    if brand_el:
                        brand_text = brand_el.get_text(strip=True)

                link_el = parent.find("a", href=True) if parent else None
                product_url = link_el["href"] if link_el else ""

                if product_url and product_url in seen_urls:
                    continue

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

                if product_url:
                    seen_urls.add(product_url)

                price_el = (
                    parent.find(class_=price_class_pattern) if parent else None
                )
                raw_price = (
                    price_el.get_text(strip=True)
                    if price_el
                    else "No disponible"
                )
                price_text = clean_price_text(raw_price)

                points_el = (
                    parent.find(class_=points_class_pattern)
                    if parent
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

            if new_items == 0:
                print(
                    f"  [Page {page}] End of catalog reached. Stopping.",
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
    print(f"Found {len(stores)} Yahoo stores in '{excel_input}'.", flush=True)

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
    try:
        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            if all_results:
                df_all = pd.DataFrame(all_results)
                df_all.to_excel(writer, sheet_name="All_Stores", index=False)

            used_sheets: Set[str] = set()
            for s_name, df_s in store_dfs.items():
                base_s = sanitize_sheet_name(s_name)
                sheet_n = base_s
                ctr = 1
                while sheet_n in used_sheets or sheet_n == "All_Stores":
                    sfx = f"_{ctr}"
                    sheet_n = f"{base_s[:31-len(sfx)]}{sfx}"
                    ctr += 1
                used_sheets.add(sheet_n)
                df_s.to_excel(writer, sheet_name=sheet_n, index=False)

        print(
            f"Yahoo Excel generated successfully at '{output_excel}'!",
            flush=True,
        )

    except PermissionError:
        fallback = "yahoo_prices_by_store_updated.xlsx"
        with pd.ExcelWriter(fallback, engine="openpyxl") as writer:
            if all_results:
                df_all = pd.DataFrame(all_results)
                df_all.to_excel(writer, sheet_name="All_Stores", index=False)

            used_sheets = set()
            for s_name, df_s in store_dfs.items():
                base_s = sanitize_sheet_name(s_name)
                sheet_n = base_s
                ctr = 1
                while sheet_n in used_sheets or sheet_n == "All_Stores":
                    sfx = f"_{ctr}"
                    sheet_n = f"{base_s[:31-len(sfx)]}{sfx}"
                    ctr += 1
                used_sheets.add(sheet_n)
                df_s.to_excel(writer, sheet_name=sheet_n, index=False)

        print(
            f"Notice: '{output_excel}' is open in Excel. "
            f"Saved successfully to '{fallback}'!",
            flush=True,
        )


if __name__ == "__main__":
    scrape_all_yahoo_stores(
        excel_input=YAHOO_MASTER_EXCEL,
        output_excel=OUTPUT_YAHOO_SCRAPED_EXCEL,
        list_products_file=CATALOG_LIST_EXCEL,
    )
