"""Central Utilities for Web Scraping and Excel Processing.

Provides shared functions for loading product catalogs, extracting model
codes, parsing/formatting currency prices, cleaning URLs and points text,
sanitizing sheet names, and saving Excel workbooks with fallback handling.
"""

import re
import sys
import time
from typing import Any, Dict, List, Optional, Set, Tuple
import unicodedata
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


# ----------------------------------------------------------------------
# Japanese Knife Ontology & Model Mapping Table
# ----------------------------------------------------------------------
JAPANESE_KNIFE_ONTOLOGY: Dict[str, str] = {
    "三徳": "SANTOKU",
    "万能": "SANTOKU",
    "牛刀": "GYUTO",
    "シェフナイフ": "GYUTO",
    "ペティ": "PETTY",
    "ペティー": "PETTY",
    "小型": "PETTY",
    "パン切り": "BREAD",
    "ブレッド": "BREAD",
    "スライサー": "SLICER",
    "菜切り": "NAKIRI",
    "菜切": "NAKIRI",
    "出刃": "DEBA",
    "小出刃": "KODEBA",
    "刺身": "SASHIMI",
    "柳刃": "YANAGIBA",
    "筋引": "SUJIHIKI",
    "皮むき": "PEELING",
    "シャープナー": "SHARPENER",
    "研ぎ器": "SHARPENER",
}

# (Series, Category, Length_cm) -> Catalog Model Code for Single Knives
MODEL_ATTRIBUTE_MAP: Dict[Tuple[str, str, Optional[int]], str] = {
    # Standard Series
    ("STANDARD", "SANTOKU", 18): "G-46",
    ("STANDARD", "SANTOKU", 16): "G-57",
    ("STANDARD", "SANTOKU", 14): "GS-201",
    ("STANDARD", "GYUTO", 20): "G-2",
    ("STANDARD", "GYUTO", 18): "G-55",
    ("STANDARD", "PETTY", 13): "GS-3",
    ("STANDARD", "PETTY", 9): "GS-38",
    ("STANDARD", "PETTY", 10): "GS-58",
    ("STANDARD", "BREAD", 22): "G-9",
    ("STANDARD", "SLICER", 21): "G-3",
    ("STANDARD", "NAKIRI", 18): "G-5",
    ("STANDARD", "NAKIRI", 14): "GS-5",
    ("STANDARD", "SHARPENER", None): "G-91/SB",
    # IST Series
    ("IST", "SANTOKU", 19): "IST-01",
    ("IST", "PETTY", 15): "IST-02",
    ("IST", "KODEBA", 12): "IST-05",
    ("IST", "BREAD", 20): "IST-04",
    ("IST", "SHARPENER", None): "SHARPENER",
}

# (Series, Category, Length_cm, Set_Count) -> Catalog Model Code for Knife Sets
MODEL_SET_ATTRIBUTE_MAP: Dict[Tuple[str, str, Optional[int], int], str] = {
    # Standard Series - Gyuto Sets
    ("STANDARD", "GYUTO", 16, 2): "GST-A58",
    ("STANDARD", "GYUTO", 20, 2): "GST-A2",
    ("STANDARD", "GYUTO", None, 2): "GST-A2",
    ("STANDARD", "GYUTO", 18, 3): "GST-B4",
    ("STANDARD", "GYUTO", 20, 3): "GST-B2",
    ("STANDARD", "GYUTO", None, 3): "GST-B2",
    ("STANDARD", "GYUTO", 20, 4): "GST-C2",
    ("STANDARD", "GYUTO", None, 4): "GST-C2",
    # Standard Series - Santoku Sets
    ("STANDARD", "SANTOKU", 18, 2): "GST-A46",
    ("STANDARD", "SANTOKU", 16, 2): "GST-A57",
    ("STANDARD", "SANTOKU", 14, 2): "GST-AS201/SP",
    ("STANDARD", "SANTOKU", None, 2): "GST-A46",
    ("STANDARD", "SANTOKU", 18, 3): "GST-B46",
    ("STANDARD", "SANTOKU", 16, 3): "GST-B57",
    ("STANDARD", "SANTOKU", None, 3): "GST-B46",
    ("STANDARD", "SANTOKU", 18, 4): "GST-C46",
    ("STANDARD", "SANTOKU", None, 4): "GST-C46",
    # Standard Series - Petty Sets
    ("STANDARD", "PETTY", 13, 2): "GST-AS3",
    ("STANDARD", "PETTY", None, 2): "GST-AS3",
    # IST Series Sets
    ("IST", "SANTOKU", 19, 2): "IST-A01",
    ("IST", "SANTOKU", None, 2): "IST-A01",
    ("IST", "SANTOKU", None, 3): "IST-B05",
}


def normalize_japanese_text(text: str) -> str:
    """Normalize Unicode characters, full-width alphanumeric, and spaces."""
    if not text:
        return ""
    normalized = unicodedata.normalize("NFKC", text)
    return re.sub(r"\s+", " ", normalized).strip()


def extract_knife_attributes(
    title: str,
) -> Tuple[str, Optional[str], Optional[int], Optional[int]]:
    """Extract knife series, category, length, and set count from title.

    Args:
        title: Product title text in Japanese or mixed.

    Returns:
        Tuple of (series, category, length_cm, set_count).
    """
    norm_title = normalize_japanese_text(title)

    # 1. Series detection
    series = "STANDARD"
    if re.search(r"\bIST\b|イスト|GLOBAL-IST|GLOBAL IST", norm_title, re.I):
        series = "IST"
    elif re.search(r"\bPRO\b|プロ|GLOBAL-PRO", norm_title, re.I):
        series = "PRO"

    # 2. Knife Category detection
    category: Optional[str] = None
    for keyword, cat in JAPANESE_KNIFE_ONTOLOGY.items():
        if keyword in norm_title:
            category = cat
            break

    # 3. Blade Length detection in cm (e.g. 18cm, 180mm, 18 センチ)
    length_cm: Optional[int] = None
    cm_match = re.search(r"(\d+(?:\.\d+)?)\s*(?:cm|センチ)", norm_title, re.I)
    if cm_match:
        length_cm = int(round(float(cm_match.group(1))))
    else:
        mm_match = re.search(r"(\d+)\s*mm", norm_title, re.I)
        if mm_match:
            length_cm = int(round(float(mm_match.group(1)) / 10))

    # 4. Set count detection (e.g. 2点セット, 3点セット, 4点, セット)
    set_count: Optional[int] = None
    set_match = re.search(
        r"(\d+)\s*(?:点セット|点組|点set|set|Pセット|pセット|点)",
        norm_title,
        re.I,
    )
    if set_match:
        set_count = int(set_match.group(1))
    elif "セット" in norm_title or "ギフトセット" in norm_title:
        set_count = 2

    return series, category, length_cm, set_count


def match_japanese_knife_model(title: str) -> Optional[str]:
    """Map Japanese title to catalog model code based on knife attributes.

    Args:
        title: Product title text in Japanese.

    Returns:
        Official catalog model code if mapped, otherwise None.
    """
    series, category, length_cm, set_count = extract_knife_attributes(title)
    if not category:
        return None

    # Step A: Check Knife Set mapping if a set is detected
    if set_count is not None:
        set_key = (series, category, length_cm, set_count)
        if set_key in MODEL_SET_ATTRIBUTE_MAP:
            return MODEL_SET_ATTRIBUTE_MAP[set_key]

        set_key_no_length = (series, category, None, set_count)
        if set_key_no_length in MODEL_SET_ATTRIBUTE_MAP:
            return MODEL_SET_ATTRIBUTE_MAP[set_key_no_length]

    # Step B: Check Individual Knife mapping
    key = (series, category, length_cm)
    if key in MODEL_ATTRIBUTE_MAP:
        return MODEL_ATTRIBUTE_MAP[key]

    key_no_length = (series, category, None)
    if key_no_length in MODEL_ATTRIBUTE_MAP:
        return MODEL_ATTRIBUTE_MAP[key_no_length]

    return None


def extract_product_code(
    title: str, official_codes: Optional[List[str]] = None
) -> Optional[str]:
    """Extract product code from title using catalog, regex, or fuzzy matching.

    Args:
        title: Product title text.
        official_codes: Optional list of catalog codes.

    Returns:
        Product code string if found, otherwise None.
    """
    if not title:
        return None

    # Step 1: Explicit match in official codes list
    if official_codes:
        for code in official_codes:
            pattern = r"\b" + re.escape(code) + r"\b"
            if re.search(pattern, title, re.IGNORECASE):
                return code
            if code in title:
                return code

    # Step 2: Regex prefix pattern matching
    prefixes = (
        r"(?:GKW|GST|GSS|GCB|GTF|GTJ|SST|GS|GT|G|GBX|IST|GCG|IB|"
        r"GSTC|SST|GB|GF)-[A-Za-z0-9]+(?:/[A-Za-z0-9]+)?"
    )
    match = re.search(prefixes, title, re.IGNORECASE)
    if match:
        return match.group(0).upper()

    # Step 3: Japanese knife ontology fuzzy attribute matching
    fuzzy_matched_code = match_japanese_knife_model(title)
    if fuzzy_matched_code:
        return fuzzy_matched_code

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
