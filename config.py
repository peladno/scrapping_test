"""Configuration settings for Rakuten and Yahoo Scrapers and Price Comparators.

Centralizes all input/output file paths, sheet names, HTTP parameters,
scraping settings, and CSS class patterns as clean Python constants
loaded from environment variables (.env) with robust defaults.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

# Project root directory path
BASE_DIR: Path = Path(__file__).resolve().parent

# Load environment variables from .env file
load_dotenv(BASE_DIR / ".env")

# Data Directories
INPUT_DIR: Path = BASE_DIR / "data" / "inputs"
OUTPUT_DIR: Path = BASE_DIR / "data" / "outputs"

# Ensure directories exist
INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def _resolve_path(env_val: str, default_rel: Path) -> str:
    """Resolve file path from env variable or fallback relative path."""
    if env_val:
        p = Path(env_val)
        if not p.is_absolute():
            return str(BASE_DIR / p)
        return str(p)
    return str(default_rel)


# Master Input Excel Files (Loaded from .env with fallback to data/inputs/)
RAKUTEN_MASTER_EXCEL: str = _resolve_path(
    os.getenv("RAKUTEN_MASTER_EXCEL", ""),
    INPUT_DIR / "rakuten_stores.xlsx",
)
RAKUTEN_SHEET_NAME: str = os.getenv("RAKUTEN_SHEET_NAME", "楽天市場")

YAHOO_MASTER_EXCEL: str = _resolve_path(
    os.getenv("YAHOO_MASTER_EXCEL", ""),
    INPUT_DIR / "yahoo_stores.xlsx",
)
YAHOO_SHEET_NAME: str = os.getenv("YAHOO_SHEET_NAME", "Yahoo")

CATALOG_LIST_EXCEL: str = _resolve_path(
    os.getenv("CATALOG_LIST_EXCEL", ""),
    INPUT_DIR / "list-products.xlsx",
)

# Output Excel Files (Generated in data/outputs/)
OUTPUT_SCRAPED_EXCEL: str = _resolve_path(
    os.getenv("OUTPUT_SCRAPED_EXCEL", ""),
    OUTPUT_DIR / "rakuten_prices_by_store.xlsx",
)
OUTPUT_COMPARISON_EXCEL: str = _resolve_path(
    os.getenv("OUTPUT_COMPARISON_EXCEL", ""),
    OUTPUT_DIR / "rakuten_price_comparison.xlsx",
)
OUTPUT_YAHOO_SCRAPED_EXCEL: str = _resolve_path(
    os.getenv("OUTPUT_YAHOO_SCRAPED_EXCEL", ""),
    OUTPUT_DIR / "yahoo_prices_by_store.xlsx",
)
OUTPUT_YAHOO_COMPARISON_EXCEL: str = _resolve_path(
    os.getenv("OUTPUT_YAHOO_COMPARISON_EXCEL", ""),
    OUTPUT_DIR / "yahoo_price_comparison.xlsx",
)

# Scraping Settings
DEFAULT_STORE_ID: str = os.getenv("DEFAULT_STORE_ID", "211966")
MAX_PAGES_PER_STORE: int = int(os.getenv("MAX_PAGES_PER_STORE", "50"))
HTTP_TIMEOUT: int = int(os.getenv("HTTP_TIMEOUT", "15"))
HTTP_RETRIES: int = int(os.getenv("HTTP_RETRIES", "3"))
COURTESY_PAUSE_SECONDS: float = float(
    os.getenv("COURTESY_PAUSE_SECONDS", "1.5")
)

# ----------------------------------------------------------------------
# Target Brand / Product Keyword Filtering
# ----------------------------------------------------------------------
TARGET_KEYWORD: str = str(os.getenv("TARGET_KEYWORD") or "GLOBAL")

# ----------------------------------------------------------------------
# Rakuten Scraping CSS Selectors / Class Patterns
# ----------------------------------------------------------------------
RAKUTEN_TITLE_CLASS: str = r"title-link"
RAKUTEN_PRICE_CLASS: str = r"price--"
RAKUTEN_POINTS_CLASS: str = r"points--"
RAKUTEN_CARD_CLASS: str = r"searchresultitem|dui-card"

# ----------------------------------------------------------------------
# Yahoo Shopping Scraping CSS Selectors / Class Patterns
# ----------------------------------------------------------------------
YAHOO_DETAIL_LINK_CLASS: str = (
    r"SearchResult_SearchResultItem__detailLink|detailLink"
)
YAHOO_CARD_CLASS: str = r"SearchResult_SearchResultItem"
YAHOO_TITLE_CLASS: str = r"ItemTitle_SearchResultItemTitle"
YAHOO_BRAND_CLASS: str = r"ItemBrand_SearchResultItemBrand"
YAHOO_PRICE_CLASS: str = r"ItemPrice_ItemPrice|ItemPrice"
YAHOO_POINTS_CLASS: str = r"ItemPointModal|PointText|PointRate"
