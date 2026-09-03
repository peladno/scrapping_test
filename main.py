"""Unified Multi-Platform E-Commerce Scraping & Price Comparison CLI.

Orchestrates web scraping and official price comparisons across multiple
e-commerce platforms (Rakuten, Yahoo Shopping, etc.).
"""

import argparse
import sys
import time
from typing import List, Optional

from amazon_scraper import scrape_all_amazon_products
from compare_prices import compare_and_highlight_excel
from config import (
    CATALOG_LIST_EXCEL,
    OUTPUT_AMAZON_COMPARISON_EXCEL,
    OUTPUT_AMAZON_SCRAPED_EXCEL,
    OUTPUT_COMPARISON_EXCEL,
    OUTPUT_SCRAPED_EXCEL,
    OUTPUT_YAHOO_COMPARISON_EXCEL,
    OUTPUT_YAHOO_SCRAPED_EXCEL,
    OUTPUT_YODOBASHI_COMPARISON_EXCEL,
    OUTPUT_YODOBASHI_SCRAPED_EXCEL,
    RAKUTEN_MASTER_EXCEL,
    YAHOO_MASTER_EXCEL,
)
from rakuten_scraper import scrape_all_rakuten_stores
from yahoo_scraper import scrape_all_yahoo_stores
from yodobashi_scraper import scrape_all_yodobashi_products

# Configure UTF-8 encoding for Windows console output
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def run_rakuten_pipeline(scrape: bool = True, compare: bool = True) -> None:
    """Execute scraping and price comparison for Rakuten stores.

    Args:
        scrape: If True, executes store web scraping and saves Excel.
        compare: If True, compares prices against official catalog.
    """
    print("\n" + "=" * 60)
    print("🚀 PIPELINE: RAKUTEN")
    print("=" * 60)
    start_time = time.time()

    if scrape:
        print("\n[Step 1/2] Scraping Rakuten stores...")
        scrape_all_rakuten_stores(
            excel_input=RAKUTEN_MASTER_EXCEL,
            output_excel=OUTPUT_SCRAPED_EXCEL,
            list_products_file=CATALOG_LIST_EXCEL,
        )

    if compare:
        print("\n[Step 2/2] Comparing Rakuten prices against catalog...")
        compare_and_highlight_excel(
            scraped_excel_input=OUTPUT_SCRAPED_EXCEL,
            list_products_file=CATALOG_LIST_EXCEL,
            output_excel=OUTPUT_COMPARISON_EXCEL,
            check_points=True,
        )

    elapsed = time.time() - start_time
    print(f"✨ Rakuten pipeline completed in {elapsed:.2f}s.")


def run_yahoo_pipeline(scrape: bool = True, compare: bool = True) -> None:
    """Execute scraping and price comparison for Yahoo Shopping stores.

    Args:
        scrape: If True, executes store web scraping and saves Excel.
        compare: If True, compares prices against official catalog.
    """
    print("\n" + "=" * 60)
    print("🚀 PIPELINE: YAHOO SHOPPING")
    print("=" * 60)
    start_time = time.time()

    if scrape:
        print("\n[Step 1/2] Scraping Yahoo Shopping stores...")
        scrape_all_yahoo_stores(
            excel_input=YAHOO_MASTER_EXCEL,
            output_excel=OUTPUT_YAHOO_SCRAPED_EXCEL,
            list_products_file=CATALOG_LIST_EXCEL,
        )

    if compare:
        print("\n[Step 2/2] Comparing Yahoo prices against catalog...")
        compare_and_highlight_excel(
            scraped_excel_input=OUTPUT_YAHOO_SCRAPED_EXCEL,
            list_products_file=CATALOG_LIST_EXCEL,
            output_excel=OUTPUT_YAHOO_COMPARISON_EXCEL,
            check_points=False,
        )

    elapsed = time.time() - start_time
    print(f"✨ Yahoo pipeline completed in {elapsed:.2f}s.")


def run_amazon_pipeline(scrape: bool = True, compare: bool = True) -> None:
    """Execute scraping and price comparison for Amazon Japan products.

    Args:
        scrape: If True, executes product scraping and saves Excel.
        compare: If True, compares prices against official catalog.
    """
    print("\n" + "=" * 60)
    print("🚀 PIPELINE: AMAZON JAPAN")
    print("=" * 60)
    start_time = time.time()

    if scrape:
        print("\n[Step 1/2] Scraping Amazon Japan products...")
        scrape_all_amazon_products(
            output_excel=OUTPUT_AMAZON_SCRAPED_EXCEL,
            list_products_file=CATALOG_LIST_EXCEL,
        )

    if compare:
        print("\n[Step 2/2] Comparing Amazon prices against catalog...")
        compare_and_highlight_excel(
            scraped_excel_input=OUTPUT_AMAZON_SCRAPED_EXCEL,
            list_products_file=CATALOG_LIST_EXCEL,
            output_excel=OUTPUT_AMAZON_COMPARISON_EXCEL,
            check_points=True,
            point_platform="amazon",
        )

    elapsed = time.time() - start_time
    print(f"✨ Amazon pipeline completed in {elapsed:.2f}s.")


def run_yodobashi_pipeline(scrape: bool = True, compare: bool = True) -> None:
    """Execute scraping and price comparison for Yodobashi Camera products.

    Args:
        scrape: If True, executes product scraping and saves Excel.
        compare: If True, compares prices against official catalog.
    """
    print("\n" + "=" * 60)
    print("🚀 PIPELINE: YODOBASHI CAMERA")
    print("=" * 60)
    start_time = time.time()

    if scrape:
        print("\n[Step 1/2] Scraping Yodobashi Camera products...")
        scrape_all_yodobashi_products(
            output_excel=OUTPUT_YODOBASHI_SCRAPED_EXCEL,
            list_products_file=CATALOG_LIST_EXCEL,
        )

    if compare:
        print("\n[Step 2/2] Comparing Yodobashi prices against catalog...")
        compare_and_highlight_excel(
            scraped_excel_input=OUTPUT_YODOBASHI_SCRAPED_EXCEL,
            list_products_file=CATALOG_LIST_EXCEL,
            output_excel=OUTPUT_YODOBASHI_COMPARISON_EXCEL,
            check_points=True,
            point_platform="yodobashi",
        )

    elapsed = time.time() - start_time
    print(f"✨ Yodobashi pipeline completed in {elapsed:.2f}s.")


def build_cli_parser() -> argparse.ArgumentParser:
    """Build command-line argument parser.

    Returns:
        Configured ArgumentParser instance.
    """
    parser = argparse.ArgumentParser(
        description="Unified E-Commerce Scraper & Price Comparator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  poetry run python main.py --platform rakuten
  poetry run python main.py --platform yahoo
  poetry run python main.py --platform amazon
  poetry run python main.py --platform yodobashi
  poetry run python main.py --platform all
  poetry run python main.py --platform yodobashi --scrape-only
  poetry run python main.py --platform all --compare-only
        """,
    )
    parser.add_argument(
        "-p",
        "--platform",
        choices=["rakuten", "yahoo", "amazon", "yodobashi", "all"],
        default="all",
        help="Target platform to process (default: all)",
    )
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--scrape-only",
        action="store_true",
        help="Only run web scraping without price comparison",
    )
    group.add_argument(
        "--compare-only",
        action="store_true",
        help="Only run price comparison on existing scraped Excel files",
    )
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    """CLI application entrypoint.

    Args:
        argv: Optional list of command-line argument strings.

    Returns:
        Exit code (0 for success, non-zero for error).
    """
    parser = build_cli_parser()
    args = parser.parse_args(argv)

    scrape = not args.compare_only
    compare = not args.scrape_only

    total_start = time.time()

    if args.platform in ["rakuten", "all"]:
        run_rakuten_pipeline(scrape=scrape, compare=compare)

    if args.platform in ["yahoo", "all"]:
        run_yahoo_pipeline(scrape=scrape, compare=compare)

    if args.platform in ["amazon", "all"]:
        run_amazon_pipeline(scrape=scrape, compare=compare)

    if args.platform in ["yodobashi", "all"]:
        run_yodobashi_pipeline(scrape=scrape, compare=compare)

    total_elapsed = time.time() - total_start
    print("\n" + "=" * 60)
    print(f"🎉 All requested pipelines finished in {total_elapsed:.2f}s.")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
