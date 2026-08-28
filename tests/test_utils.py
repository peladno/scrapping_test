"""Unit tests for utils.py module."""

from typing import List
from utils import (
    clean_points_text,
    clean_price_text,
    clean_product_url,
    extract_product_code,
    format_currency_yen,
    parse_numeric_price,
    sanitize_sheet_name,
)


def test_parse_numeric_price() -> None:
    """Test numeric price parsing from various raw inputs."""
    assert parse_numeric_price("12,100円") == 12100.0
    assert parse_numeric_price("¥9,900") == 9900.0
    assert parse_numeric_price(23100) == 23100.0
    assert parse_numeric_price("No disponible") is None
    assert parse_numeric_price(None) is None


def test_format_currency_yen() -> None:
    """Test Yen currency formatting."""
    assert format_currency_yen(12100.0) == "¥12,100"
    assert format_currency_yen(9900.0) == "¥9,900"
    assert format_currency_yen(None) == "N/A"


def test_clean_price_text() -> None:
    """Test clean price text formatting."""
    assert clean_price_text("12,100円") == "¥12,100"
    assert clean_price_text("9,900円〜") == "¥9,900+"
    assert clean_price_text("No disponible") == "N/A"


def test_clean_points_text() -> None:
    """Test points text cleaning."""
    assert clean_points_text("5%（454pt）") == "5%（454pt）"
    assert clean_points_text("No disponible") == "N/A"
    assert clean_points_text("") == "N/A"


def test_clean_product_url() -> None:
    """Test unquoting and cleaning product URLs."""
    raw_tracking = (
        "https://shopping-item-reach.yahoo.co.jp/v1/click?uuid=123"
        "&rdUrl=https%3A%2F%2Fstore.shopping.yahoo.co.jp%2Fp-s%2Fy002.html"
        "%3Fnodeeplink%3D0"
    )
    expected = "https://store.shopping.yahoo.co.jp/p-s/y002.html"
    assert clean_product_url(raw_tracking) == expected

    simple_url = "https://store.shopping.yahoo.co.jp/eclity/g-46.html?sc_i=123"
    assert (
        clean_product_url(simple_url)
        == "https://store.shopping.yahoo.co.jp/eclity/g-46.html"
    )


def test_extract_product_code(sample_catalog_codes: List[str]) -> None:
    """Test product model code extraction."""
    title = "特典付 GLOBAL 三徳 18cm G-46 三徳包丁 万能包丁"
    assert extract_product_code(title, sample_catalog_codes) == "G-46"

    title_ist = "GLOBAL-IST 小型 15cm IST-02 両刃"
    assert extract_product_code(title_ist, sample_catalog_codes) == "IST-02"

    title_generic = "GLOBAL 包丁 カッティングボード GCB-02"
    assert extract_product_code(title_generic, None) == "GCB-02"

    assert extract_product_code("", sample_catalog_codes) is None


def test_sanitize_sheet_name() -> None:
    """Test Excel sheet name sanitization and max 31 char limit."""
    long_name = "Tienda: Con [Caracteres/Especiales] * Y Muchos Caracteres"
    sanitized = sanitize_sheet_name(long_name)
    assert len(sanitized) <= 31
    assert "/" not in sanitized
    assert "[" not in sanitized
