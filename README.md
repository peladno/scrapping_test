# 🛒 E-Commerce Scraper & Price Comparator (Rakuten & Yahoo Shopping Japan)

An automated **Web Scraping** and **Price Comparison** engine for Japanese e-commerce stores on **Rakuten Ichiba** and **Yahoo Shopping Japan**.

The system extracts product catalogs, filters items by brand/keywords (e.g., `GLOBAL`) and official product model codes, exports consolidated multi-tab Excel workbooks, and generates color-coded price comparison reports against an official product catalog list.

---

## 📁 Project Structure

```text
scrapping/
├── config.py                 # Central configuration, .env loader, and CSS selectors
├── utils.py                  # Shared utilities (parsing, cleaning, regex, Excel export)
├── rakuten_scraper.py        # Rakuten store scraper with dynamic pagination handling
├── yahoo_scraper.py          # Yahoo Shopping store scraper with card-level parsing
├── compare_prices.py         # Rakuten price comparator vs official catalog
├── yahoo_compare_prices.py   # Yahoo Shopping price comparator vs official catalog
├── data/
│   ├── inputs/               # Master input Excel workbooks (ignored by Git)
│   └── outputs/              # Generated scraped & comparison Excel reports (ignored by Git)
├── tests/                    # Automated unit test suite with pytest
│   ├── conftest.py
│   ├── test_utils.py
│   ├── test_compare_prices.py
│   └── test_scrapers.py
├── .env.example              # Environment variables template
├── .gitignore                # Git exclusion rules (Excel files, .env, pycache, venv)
├── pyproject.toml            # Poetry configuration, dependencies, and tool settings
└── README.md
```

---

## 🚀 Requirements

- **Python 3.11+**
- **Poetry** (Dependency and virtual environment manager)

---

## ⚙️ Installation & Setup

### 1. Install Dependencies

```bash
# Install all required and development packages with Poetry
poetry install
```

### 2. Configure Environment Variables (`.env`)

Copy `.env.example` to create your local `.env` file:

```bash
cp .env.example .env
```

Configure `.env` with your input and output file paths:

```env
# Master Input Excel Files
RAKUTEN_MASTER_EXCEL="data/inputs/rakuten_stores.xlsx"
RAKUTEN_SHEET_NAME="楽天市場"

YAHOO_MASTER_EXCEL="data/inputs/yahoo_stores.xlsx"
YAHOO_SHEET_NAME="Yahoo"

CATALOG_LIST_EXCEL="data/inputs/list-products.xlsx"

# Generated Output Excel Files
OUTPUT_SCRAPED_EXCEL="data/outputs/rakuten_prices_by_store.xlsx"
OUTPUT_COMPARISON_EXCEL="data/outputs/rakuten_price_comparison.xlsx"
OUTPUT_YAHOO_SCRAPED_EXCEL="data/outputs/yahoo_prices_by_store.xlsx"
OUTPUT_YAHOO_COMPARISON_EXCEL="data/outputs/yahoo_price_comparison.xlsx"
```

> 🔒 **Security Notice:** All `.xlsx` files and `.env` files are excluded in `.gitignore` to prevent confidential data, proprietary store lists, or catalog prices from being committed to Git.

---

## 📖 Usage Guide

### 1. Unified CLI Orchestrator (`main.py`)

Run the entire pipeline (scraping + comparison) or target specific platforms with a single command:

```bash
# Run all platforms (Rakuten, Yahoo Shopping, Amazon Japan, Yodobashi Camera)
poetry run python main.py

# Run specific platform pipeline
poetry run python main.py --platform rakuten
poetry run python main.py --platform yahoo
poetry run python main.py --platform amazon
poetry run python main.py --platform yodobashi

# Scrape only (skip price comparison)
poetry run python main.py --platform all --scrape-only

# Compare only (run on existing scraped Excel files)
poetry run python main.py --platform all --compare-only
```

---

### 2. Standalone Scripts

You can also run individual scrapers and comparators independently:

```bash
# Individual Scrapers
poetry run python rakuten_scraper.py
poetry run python yahoo_scraper.py
poetry run python amazon_scraper.py
poetry run python yodobashi_scraper.py

# Individual Comparators
poetry run python compare_prices.py
poetry run python yahoo_compare_prices.py
```

#### 🎨 Report Color Highlights:

- 🟢 **Green (`#C6EFCE`)**: Scraped price matches the official catalog price exactly.
- 🔴 **Red (`#FFC7CE`)**: Product code is in the catalog, but the store's published price differs.
- 🟡 **Yellow (`#FFF2CC`)**: Product model code was not found in the official catalog or is unassigned.
- **Point Status Column**: Indicates `⭕` when point thresholds are compliant (Rakuten & Yodobashi $\le 1\%$; Amazon 1% to 2%) or `❌` if exceeded or non-compliant.

---

## 🧪 Testing & Code Quality

The codebase includes automated unit tests and strict static type checks:

```bash
# Run pytest test suite
poetry run pytest -v

# Code style and linting (Flake8)
poetry run flake8 config.py utils.py compare_prices.py yahoo_compare_prices.py rakuten_scraper.py yahoo_scraper.py tests/

# Static type checking (Mypy)
poetry run mypy config.py utils.py compare_prices.py yahoo_compare_prices.py rakuten_scraper.py yahoo_scraper.py tests/
```

---

## 🛠️ CSS Selectors & DOM Maintenance

All CSS class selectors and regex patterns are centralized in `config.py`. If Rakuten or Yahoo Shopping updates their DOM structure or CSS class names, **update the constants in `config.py` without modifying the core scraper logic**:

```python
# Rakuten Selectors
RAKUTEN_TITLE_CLASS = r"title-link"
RAKUTEN_PRICE_CLASS = r"price--"
RAKUTEN_POINTS_CLASS = r"points--"
RAKUTEN_CARD_CLASS = r"searchresultitem|dui-card"

# Yahoo Shopping Selectors
YAHOO_DETAIL_LINK_CLASS = r"SearchResult_SearchResultItem__detailLink|detailLink"
YAHOO_CARD_CLASS = r"SearchResult_SearchResultItem"
YAHOO_TITLE_CLASS = r"ItemTitle_SearchResultItemTitle"
YAHOO_BRAND_CLASS = r"ItemBrand_SearchResultItemBrand"
YAHOO_PRICE_CLASS = r"ItemPrice_ItemPrice|ItemPrice"
YAHOO_POINTS_CLASS = r"ItemPointModal|PointText|PointRate"
```
