"""Configuration settings for Rakuten and Yahoo Scrapers and Price Comparators.

Centralizes all input/output file paths, sheet names, HTTP parameters,
and scraping settings as clean Python constants for quick access.
"""

from pathlib import Path

# Project root directory path
BASE_DIR: Path = Path(__file__).resolve().parent

# Data Directories
INPUT_DIR: Path = BASE_DIR / "data" / "inputs"
OUTPUT_DIR: Path = BASE_DIR / "data" / "outputs"

# Ensure directories exist
INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Master Input Excel Files (Stored in data/inputs/)
RAKUTEN_MASTER_EXCEL: str = str(INPUT_DIR / "rakuten_stores.xlsx")
RAKUTEN_SHEET_NAME: str = "楽天市場"
YAHOO_MASTER_EXCEL: str = str(INPUT_DIR / "yahoo_stores.xlsx")
YAHOO_SHEET_NAME: str = "Yahoo"
CATALOG_LIST_EXCEL: str = str(INPUT_DIR / "list-products.xlsx")

# Output Excel Files (Generated in data/outputs/)
OUTPUT_SCRAPED_EXCEL: str = str(OUTPUT_DIR / "rakuten_prices_by_store.xlsx")
OUTPUT_COMPARISON_EXCEL: str = str(
    OUTPUT_DIR / "rakuten_price_comparison.xlsx"
)
OUTPUT_YAHOO_SCRAPED_EXCEL: str = str(
    OUTPUT_DIR / "yahoo_prices_by_store.xlsx"
)
OUTPUT_YAHOO_COMPARISON_EXCEL: str = str(
    OUTPUT_DIR / "yahoo_price_comparison.xlsx"
)

# Scraping Settings
DEFAULT_STORE_ID: str = "211966"
MAX_PAGES_PER_STORE: int = 50
HTTP_TIMEOUT: int = 15
HTTP_RETRIES: int = 3
COURTESY_PAUSE_SECONDS: float = 1.5
