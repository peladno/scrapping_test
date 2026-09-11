# 🛒 E-Commerce Scraper & Price Comparator (Rakuten & Yahoo Shopping Japan)

An automated **Web Scraping** and **Price Comparison** engine for Japanese e-commerce stores on **Rakuten Ichiba** and **Yahoo Shopping Japan**.

The system extracts product catalogs, filters items by brand/keywords (e.g., `GLOBAL`) and official product model codes, exports consolidated multi-tab Excel workbooks, and generates color-coded price comparison reports against an official product catalog list.

---

## 📁 Project Structure

```text
scrapping/
├── core/                     # Central domain logic, utils & configuration
│   ├── __init__.py
│   ├── config.py             # Configuration, URLs, and CSS selectors
│   ├── utils.py              # Parsers, regex, Japanese knife ontology, Excel export
│   └── comparator.py         # Price comparison engine & color highlighter
├── scrapers/                 # Dedicated web scrapers by e-commerce platform
│   ├── __init__.py
│   ├── rakuten_scraper.py    # Rakuten Ichiba store scraper
│   ├── yahoo_scraper.py      # Yahoo Shopping store scraper
│   ├── amazon_scraper.py     # Amazon Japan product scraper
│   ├── yodobashi_scraper.py  # Yodobashi Camera product scraper
│   ├── aqua_scraper.py       # Import Shop Aqua catalog scraper
│   └── furaipan_scraper.py   # Furaipan Club catalog scraper
├── data/
│   ├── inputs/               # Master input Excel workbooks (ignored by Git)
│   └── outputs/              # Generated scraped & comparison Excel reports (ignored by Git)
├── tests/                    # Automated unit test suite with pytest
│   ├── conftest.py
│   ├── test_utils.py
│   ├── test_compare_prices.py
│   ├── test_scrapers.py
│   └── test_main.py
├── main.py                   # Unified CLI entrypoint
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
OUTPUT_AQUA_SCRAPED_EXCEL="data/outputs/aqua_prices_by_store.xlsx"
OUTPUT_AQUA_COMPARISON_EXCEL="data/outputs/aqua_price_comparison.xlsx"
OUTPUT_FURAIPAN_SCRAPED_EXCEL="data/outputs/furaipan_prices_by_store.xlsx"
OUTPUT_FURAIPAN_COMPARISON_EXCEL="data/outputs/furaipan_price_comparison.xlsx"
```

> 🔒 **Security Notice:** All `.xlsx` files and `.env` files are excluded in `.gitignore` to prevent confidential data, proprietary store lists, or catalog prices from being committed to Git.

---

## 🖥️ Interactive Terminal Menu (Control Panel)

You can launch an interactive terminal menu in English without needing to type CLI flags:

```bash
# Launch interactive menu with Poetry
poetry run python main.py -i

# Or simply run main.exe (or double-click main.exe in Windows)
.\main.exe
```

```text
==================================================================
  🛒 E-COMMERCE SCRAPER & COMPARATOR - CONTROL PANEL
==================================================================
  [1] 🚀 Run FULL Pipeline (Scrape + Compare) - ALL Platforms
  [2] 🌐 Run FULL Pipeline for a Single Platform
  [3] 📥 Scrape Data Only (Skip Comparison)
  [4] 📊 Compare Prices Only (Existing Excel Files)
  [5] 🧪 Run Automated Tests (Pytest)
  [6] 📁 Check System & Excel Files Status
  [0] ❌ Exit
==================================================================
👉 Select an option [0-6]:
```

---

## 📖 Execution & Usage Methods

The project supports **two execution modes** that work identically:

1. **Direct execution with Python & Poetry** (ideal for development and regular use).
2. **Standalone Windows Executable (`main.exe`)** (ideal for running on any PC without installing Python or Poetry).

---

### Option A: Running with Python & Poetry (Standard)

Use this method when developing or running on machines with Python and Poetry installed:

#### 1. Unified CLI (`main.py`)

```bash
# Run all platforms (Rakuten, Yahoo Shopping, Amazon Japan, Yodobashi Camera, Aqua, Furaipan)
poetry run python main.py

# Run specific platform pipeline
poetry run python main.py --platform rakuten
poetry run python main.py --platform yahoo
poetry run python main.py --platform amazon
poetry run python main.py --platform yodobashi
poetry run python main.py --platform aqua
poetry run python main.py --platform furaipan

# Scrape only (skip price comparison)
poetry run python main.py --platform all --scrape-only

# Compare only (run on existing scraped Excel files)
poetry run python main.py --platform all --compare-only
```

#### 2. Individual Scraper Modules

```bash
# Run standalone scraper scripts directly
poetry run python -m scrapers.rakuten_scraper
poetry run python -m scrapers.yahoo_scraper
poetry run python -m scrapers.amazon_scraper
poetry run python -m scrapers.yodobashi_scraper
poetry run python -m scrapers.aqua_scraper
poetry run python -m scrapers.furaipan_scraper
```

---

### Option B: Standalone Executable (`main.exe`)

Use this method to distribute and execute the scraper on any Windows machine **without installing Python or dependencies**.

#### 1. Build the Executable

The project includes a pre-configured [`main.spec`] that bundles all C-extensions, browser TLS emulators (`curl_cffi`), Excel engines (`openpyxl`, `pandas`), and scraper modules:

```bash
# Compile dist/main.exe
poetry run pyinstaller --clean --noconfirm main.spec
```

The output binary is placed in `dist/main.exe`.

#### 2. Distributing and Running on Any Windows PC

To run on another machine, copy or zip the following folder layout:

```text
Scraper_App/
├── main.exe                   # From dist/main.exe
├── .env                       # (Optional) Custom URLs & environment variables
└── data/
    ├── inputs/                # Master catalog & store Excel input files
    │   ├── list-products.xlsx
    │   ├── rakuten_stores.xlsx
    │   └── yahoo_stores.xlsx
    └── outputs/               # Folder where output Excel files are written
```

Run directly from PowerShell or Command Prompt (CMD):

```powershell
# Run all platforms
.\main.exe

# Run specific platform
.\main.exe --platform rakuten
.\main.exe --platform yahoo
.\main.exe --platform amazon
.\main.exe --platform yodobashi
.\main.exe --platform aqua
.\main.exe --platform furaipan

# Compare existing Excel files only
.\main.exe --platform all --compare-only

# Scrape only
.\main.exe --platform all --scrape-only
```

---

### 🎨 Report Color Highlights:

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
poetry run flake8 core/ scrapers/ tests/ main.py

# Static type checking (Mypy)
poetry run mypy core/ scrapers/ tests/ main.py
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
