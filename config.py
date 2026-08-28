"""Configuration settings for Rakuten Scraper and Price Comparator.

Centralizes all input/output file paths, sheet names, HTTP parameters,
and scraping settings as clean Python constants for quick access.
"""

from pathlib import Path

# Project root directory path
BASE_DIR: Path = Path(__file__).resolve().parent

# Master Input Excel Files
RAKUTEN_MASTER_EXCEL: str = "rakuten_stores.xlsx"
RAKUTEN_SHEET_NAME: str = "楽天市場"
YAHOO_MASTER_EXCEL: str = "yahoo_stores.xlsx"
YAHOO_SHEET_NAME: str = "Yahoo"
CATALOG_LIST_EXCEL: str = "list-products.xlsx"

# Output Excel Files
OUTPUT_SCRAPED_EXCEL: str = "rakuten_prices_by_store.xlsx"
OUTPUT_COMPARISON_EXCEL: str = "rakuten_price_comparison.xlsx"
OUTPUT_YAHOO_SCRAPED_EXCEL: str = "yahoo_prices_by_store.xlsx"
OUTPUT_YAHOO_COMPARISON_EXCEL: str = "yahoo_price_comparison.xlsx"

# Scraping Settings
DEFAULT_STORE_ID: str = "211966"
MAX_PAGES_PER_STORE: int = 50
HTTP_TIMEOUT: int = 15
HTTP_RETRIES: int = 3
COURTESY_PAUSE_SECONDS: float = 1.5
