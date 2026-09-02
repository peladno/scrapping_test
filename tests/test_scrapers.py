"""Unit tests for scraper modules (yahoo_scraper and rakuten_scraper)."""

import re
from bs4 import BeautifulSoup

from config import (
    YAHOO_DETAIL_LINK_CLASS,
    YAHOO_POINTS_CLASS,
    YAHOO_PRICE_CLASS,
    YAHOO_TITLE_CLASS,
)
from amazon_scraper import (
    build_amazon_search_url,
    parse_amazon_search_page,
)
from rakuten_scraper import (
    build_rakuten_page_url,
    parse_rakuten_spec_table,
)
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

    detail_link_pattern = re.compile(YAHOO_DETAIL_LINK_CLASS)
    title_pattern = re.compile(YAHOO_TITLE_CLASS)
    price_pattern = re.compile(YAHOO_PRICE_CLASS)
    points_pattern = re.compile(YAHOO_POINTS_CLASS)

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


def test_parse_rakuten_spec_table() -> None:
    """Test extracting model code from Rakuten SpecTableArea HTML."""
    sample_spec_html = """
    <html>
      <body>
        <table>
          <tr>
            <td irc="SpecTableArea">
              <table>
                <tbody>
                  <tr>
                    <td><div>ブランド名</div></td>
                    <td><div>GLOBAL / グローバル</div></td>
                  </tr>
                  <tr>
                    <td><div>メーカー型番</div></td>
                    <td><div>G-46</div></td>
                  </tr>
                  <tr>
                    <td><div>代表カラー</div></td>
                    <td><div>シルバー</div></td>
                  </tr>
                </tbody>
              </table>
            </td>
          </tr>
        </table>
      </body>
    </html>
    """
    code = parse_rakuten_spec_table(sample_spec_html, ["G-46", "G-2"])
    assert code == "G-46"

    # Test with surrounding Japanese text in model code div
    sample_spec_html_complex = """
    <td irc="SpecTableArea">
      <table>
        <tr>
          <td><div>メーカー型番</div></td>
          <td><div>型番：GST-A58（牛刀2点セット）</div></td>
        </tr>
      </table>
    </td>
    """
    code_complex = parse_rakuten_spec_table(
        sample_spec_html_complex, ["GST-A58"]
    )
    assert code_complex == "GST-A58"


def test_build_amazon_search_url() -> None:
    """Test Amazon search pagination URL construction."""
    url_p1 = build_amazon_search_url("GLOBAL 包丁", 1)
    assert "https://www.amazon.co.jp/s?k=" in url_p1
    assert "ref=sr_pg_1" in url_p1

    url_p2 = build_amazon_search_url("GLOBAL 包丁", 2)
    assert "page=2" in url_p2
    assert "ref=sr_pg_2" in url_p2


def test_parse_amazon_search_page() -> None:
    """Test parsing Amazon Japan search result HTML cards."""
    sample_amazon_card = """
    <div data-component-type="s-search-result" data-asin="B0006A03QA">
      <h2 class="s-line-clamp-1">Global</h2>
      <h2>
        <a class="a-text-normal"
           href="/Global-Santoku-Length-G-46/dp/B0006A03QA">
          Santoku Blade Length 7.1 inches (18 cm) G-46
        </a>
      </h2>
      <span class="a-price">
        <span class="a-offscreen">¥12,100</span>
      </span>
      <span class="a-color-price">242 pt (2%)</span>
    </div>
    <div data-component-type="s-search-result" data-asin="B003YUBLUQ">
      <h2 class="s-line-clamp-1">KAI</h2>
      <a class="a-text-normal" href="/KAI-AE5200/dp/B003YUBLUQ">
        KAI Sekimagoroku Damascus 165mm
      </a>
      <span class="a-price"><span class="a-offscreen">¥8,491</span></span>
    </div>
    <div data-component-type="s-search-result" data-asin="B00005OL44">
      <a class="a-text-normal" href="/dp/B00005OL44">
        グローバル 包丁 牛刀 20cm G-2
      </a>
      <span class="a-price"><span class="a-offscreen">¥12,100</span></span>
    </div>
    """
    records = parse_amazon_search_page(
        sample_amazon_card, ["G-46", "G-2"]
    )
    # Competitor KAI should be excluded; Global and グローバル included
    assert len(records) == 2
    r1 = records[0]
    assert r1["ASIN"] == "B0006A03QA"
    assert r1["Product_Code"] == "G-46"
    assert r1["points status"] == "⭕"

    r2 = records[1]
    assert r2["ASIN"] == "B00005OL44"
    assert r2["Product_Code"] == "G-2"
    assert "グローバル" in r2["Product"]
    assert r2["points status"] == "❌"
