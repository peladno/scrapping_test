"""Scrapers module for e-commerce platforms."""

from scrapers.amazon_scraper import scrape_all_amazon_products
from scrapers.aqua_scraper import scrape_all_aqua_products
from scrapers.furaipan_scraper import scrape_all_furaipan_products
from scrapers.rakuten_scraper import scrape_all_rakuten_stores
from scrapers.yahoo_scraper import scrape_all_yahoo_stores
from scrapers.yodobashi_scraper import scrape_all_yodobashi_products

__all__ = [
    "scrape_all_amazon_products",
    "scrape_all_aqua_products",
    "scrape_all_furaipan_products",
    "scrape_all_rakuten_stores",
    "scrape_all_yahoo_stores",
    "scrape_all_yodobashi_products",
]
