"""Yahoo Price Comparison and Excel Highlighter.

Compares scraped product prices from Yahoo store Excel files
against official catalog prices configured in environment by model code.

Color Highlights:
- RED (#FFC7CE): Code exists in official catalog, but price differs.
- YELLOW (#FFF2CC): Product code not found in catalog or unassigned.
- GREEN (#C6EFCE): Scraped price matches official catalog price.
"""

import sys

from compare_prices import compare_and_highlight_excel
from config import (
    CATALOG_LIST_EXCEL,
    OUTPUT_YAHOO_COMPARISON_EXCEL,
    OUTPUT_YAHOO_SCRAPED_EXCEL,
)

# Configure UTF-8 encoding for Windows console output
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

if __name__ == "__main__":
    compare_and_highlight_excel(
        scraped_excel_input=OUTPUT_YAHOO_SCRAPED_EXCEL,
        list_products_file=CATALOG_LIST_EXCEL,
        output_excel=OUTPUT_YAHOO_COMPARISON_EXCEL,
        check_points=False,
    )
