"""Core module for e-commerce scraping, comparison and configuration."""

from core.comparator import compare_and_highlight_excel, compare_single_item
from core.config import BASE_DIR, CATALOG_LIST_EXCEL, INPUT_DIR, OUTPUT_DIR
from core.utils import (
    clean_points_text,
    clean_price_text,
    clean_product_url,
    evaluate_amazon_point_status,
    evaluate_point_status,
    extract_product_code,
    format_currency_yen,
    load_official_product_codes,
    parse_numeric_price,
    sanitize_sheet_name,
    save_excel_with_fallback,
)

__all__ = [
    "BASE_DIR",
    "INPUT_DIR",
    "OUTPUT_DIR",
    "CATALOG_LIST_EXCEL",
    "compare_and_highlight_excel",
    "compare_single_item",
    "clean_points_text",
    "clean_price_text",
    "clean_product_url",
    "evaluate_amazon_point_status",
    "evaluate_point_status",
    "extract_product_code",
    "format_currency_yen",
    "load_official_product_codes",
    "parse_numeric_price",
    "sanitize_sheet_name",
    "save_excel_with_fallback",
]
