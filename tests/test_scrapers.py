"""Unit tests for scraper modules (yahoo_scraper and raskuten_scraper)."""

import re
from bs4 import BeautifulSoup

from rakuten_scraper import build_rakuten_page_url
from yahoo_scraper import build_yahoo_page_url


def test_build_rakuten_page_url_inshop() -> None:
    """Test pagination URL construction for Rakuten inshop-mall."""
    url = (
        "https://search.rakuten.co.jp/search/inshop-mall/"
        "GLOBAL/-/sid.211966-st.A"
    )
    assert build_rakuten_page_url(url, 1) == url
    url_p2 = build_rakuten_page_url(url, 2)
    assert "p=2" in url_p2 and "sid=211966" in url_p2


def test_build_rakuten_page_url_query() -> None:
    """Test pagination URL construction for Rakuten query format."""
    url = "https://search.rakuten.co.jp/search/mall/GLOBAL/?sid=243032"
    assert build_rakuten_page_url(url, 1) == url
    url_p2 = build_rakuten_page_url(url, 2)
    assert "p=2" in url_p2 and "sid=243032" in url_p2


def test_build_yahoo_page_url_general_search() -> None:
    """Test pagination URL construction for Yahoo general search."""
    base_url = (
        "https://shopping.yahoo.co.jp/search/global+%E5%8C%85%E4%B8%81/0/?"
        "oq=GLOBAL&first=1"
    )

    url_p1 = build_yahoo_page_url(base_url, 1)
    assert url_p1 == base_url

    url_p2 = build_yahoo_page_url(base_url, 2)
    assert "/0/2/?" in url_p2

    url_p3 = build_yahoo_page_url(base_url, 3)
    assert "/0/3/?" in url_p3


def test_build_yahoo_page_url_store_search() -> None:
    """Test pagination URL construction for Yahoo store search."""
    store_url = (
        "https://store.shopping.yahoo.co.jp/eclity/search.html?p=GLOBAL"
    )

    url_p1 = build_yahoo_page_url(store_url, 1)
    assert url_p1 == store_url

    url_p2 = build_yahoo_page_url(store_url, 2)
    assert "&page=2" in url_p2


def test_yahoo_card_extraction(yahoo_item_card_html: str) -> None:
    """Test parsing a sample Yahoo product card HTML snippet."""
    soup = BeautifulSoup(yahoo_item_card_html, "html.parser")

    detail_link_pattern = re.compile(
        r"SearchResult_SearchResultItem__detailLink|detailLink"
    )
    title_pattern = re.compile(r"ItemTitle_SearchResultItemTitle")
    price_pattern = re.compile(r"ItemPrice_ItemPrice")
    points_pattern = re.compile(r"ItemPointModal|PointText")

    link_el = soup.find("a", class_=detail_link_pattern)
    assert link_el is not None

    title_el = soup.find(class_=title_pattern)
    assert title_el is not None
    assert "G-46" in title_el.get_text(strip=True)

    price_el = soup.find(class_=price_pattern)
    assert price_el is not None
    assert "12,100" in price_el.get_text(strip=True)

    points_el = soup.find(class_=points_pattern)
    assert points_el is not None
    assert "5%" in points_el.get_text(strip=True)
