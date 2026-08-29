"""Central Utilities for Web Scraping and Excel Processing.

Provides shared functions for loading product catalogs, extracting model
codes, parsing/formatting currency prices, cleaning URLs and points text,
sanitizing sheet names, and saving Excel workbooks with fallback handling.
"""

import re
import sys
import time
from typing import Any, Dict, List, Optional, Set
import urllib.parse

import pandas as pd
import requests

from config import CATALOG_LIST_EXCEL, HTTP_RETRIES, HTTP_TIMEOUT

# Configure UTF-8 encoding for Windows console output
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


class ScrapingError(Exception):
    """Base exception class for scraping errors."""

    pass


class CatalogLoadError(ScrapingError):
    """Raised when loading official catalog Excel file fails."""

    pass


class ExportError(ScrapingError):
    """Raised when saving Excel output fails."""

    pass


def load_official_product_codes(
    list_products_file: str = CATALOG_LIST_EXCEL,
) -> List[str]:
    """Load official product model codes from catalog list Excel file.

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
            f"Note: Could not load catalog file '{list_products_file}' "
            f"({exc}). Falling back to regex prefix matching.",
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
    if not title:
        return None

    if official_codes:
        for code in official_codes:
            pattern = r"\b" + re.escape(code) + r"\b"
            if re.search(pattern, title, re.IGNORECASE):
                return code
            if code in title:
                return code

    prefixes = (
        r"(?:GKW|GST|GSS|GCB|GTF|GTJ|SST|GS|GT|G|GBX|IST|GCG|IB|"
        r"GSTC|SST|GB|GF)-[A-Za-z0-9]+(?:/[A-Za-z0-9]+)?"
    )
    match = re.search(prefixes, title, re.IGNORECASE)
    if match:
        return match.group(0).upper()

    return None


def parse_numeric_price(price_val: Any) -> Optional[float]:
    """Extract numeric float value from a price string or number.

    Args:
        price_val: Raw price string or number.

    Returns:
        Float price value or None if invalid.
    """
    if price_val is None or pd.isna(price_val):
        return None
    cleaned = re.sub(r"[^\d.]", "", str(price_val))
    try:
        return float(cleaned) if cleaned else None
    except ValueError:
        return None


def format_currency_yen(amount: Optional[float]) -> str:
    """Format float amount into Yen currency string representation.

    Args:
        amount: Numeric price amount.

    Returns:
        Formatted Yen string (e.g., '¥12,100') or 'N/A'.
    """
    if amount is None or pd.isna(amount):
        return "N/A"
    return f"¥{int(round(amount)):,}"


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
    """Clean points/percentage information text extracted from HTML.

    Args:
        raw_points: Raw points text from HTML element.

    Returns:
        Cleaned points or percentage string.
    """
    if not raw_points or raw_points == "No disponible":
        return "N/A"

    points_clean = raw_points.strip()
    return points_clean if points_clean else "N/A"


def clean_product_url(raw_url: str) -> str:
    """Extract and unquote direct product URL from tracking links.

    Args:
        raw_url: Raw href attribute from anchor element.

    Returns:
        Cleaned product URL string.
    """
    if not raw_url:
        return ""
    if "rdUrl=" in raw_url:
        match = re.search(r"rdUrl=([^&]+)", raw_url)
        if match:
            return urllib.parse.unquote(match.group(1)).split("?")[0]
    return raw_url.split("?")[0] if "?" in raw_url else raw_url


def sanitize_sheet_name(name: str) -> str:
    """Sanitize Excel worksheet tab name to max 31 valid characters.

    Args:
        name: Desired sheet name string.

    Returns:
        Sanitized valid Excel worksheet name.
    """
    clean = re.sub(r"[\\/*?:\[\]]", "_", name).strip()
    return clean[:31] if clean else "Store"


def fetch_url_with_retries(
    url: str,
    headers: Dict[str, str],
    session: Optional[requests.Session] = None,
    timeout: int = HTTP_TIMEOUT,
    retries: int = HTTP_RETRIES,
) -> Optional[requests.Response]:
    """Fetch URL over HTTP with retry mechanism and session pooling.

    Args:
        url: Target HTTP URL.
        headers: Request HTTP headers.
        session: Optional persistent requests.Session instance.
        timeout: Request timeout in seconds.
        retries: Number of retry attempts.

    Returns:
        Response object if status code is 200, otherwise None.
    """
    req_obj = session if session is not None else requests
    for attempt in range(retries):
        try:
            response = req_obj.get(url, headers=headers, timeout=timeout)
            if response.status_code == 200:
                return response
        except requests.RequestException:
            time.sleep(1.5 * (attempt + 1))
    return None


def save_excel_with_fallback(
    all_results: List[Dict[str, str]],
    store_dfs: Dict[str, pd.DataFrame],
    output_excel: str,
    fallback_name: Optional[str] = None,
) -> None:
    """Export all results DataFrame and individual store tabs into Excel.

    Handles PermissionError gracefully by saving to fallback file if original
    Excel file is open in Microsoft Excel.

    Args:
        all_results: List of all extracted item dictionaries.
        store_dfs: Dictionary of store names to DataFrames.
        output_excel: Primary target Excel file path.
        fallback_name: Optional fallback Excel file path if locked.
    """
    if fallback_name is None:
        if output_excel.endswith(".xlsx"):
            fallback_name = output_excel[:-5] + "_updated.xlsx"
        else:
            fallback_name = output_excel + "_updated.xlsx"

    try:
        _export_to_excel(all_results, store_dfs, output_excel)
        print(f"Excel generated successfully at '{output_excel}'!", flush=True)
    except PermissionError:
        _export_to_excel(all_results, store_dfs, fallback_name)
        print(
            f"Notice: '{output_excel}' is open in Excel. "
            f"Saved successfully to '{fallback_name}'!",
            flush=True,
        )


def _export_to_excel(
    all_results: List[Dict[str, str]],
    store_dfs: Dict[str, pd.DataFrame],
    target_path: str,
) -> None:
    """Internal helper to write DataFrames into multi-sheet Excel writer.

    Args:
        all_results: List of all extracted item dictionaries.
        store_dfs: Dictionary of store names to DataFrames.
        target_path: File path to save Excel workbook.
    """
    with pd.ExcelWriter(target_path, engine="openpyxl") as writer:
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
