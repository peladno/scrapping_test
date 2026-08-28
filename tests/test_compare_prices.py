"""Unit tests for compare_prices.py module."""

from typing import Dict, Optional
from compare_prices import compare_single_item


def test_compare_single_item_match(
    sample_catalog_prices: Dict[str, Dict[str, Optional[float]]],
) -> None:
    """Test exact price match against catalog tax-included price."""
    status, p_incl, p_excl = compare_single_item(
        "G-46", "12,100円", sample_catalog_prices
    )
    assert status == "MATCH"
    assert p_incl == 12100.0
    assert p_excl == 11000.0


def test_compare_single_item_mismatch(
    sample_catalog_prices: Dict[str, Dict[str, Optional[float]]],
) -> None:
    """Test price mismatch when code exists but scraped price differs."""
    status, p_incl, p_excl = compare_single_item(
        "G-46", "15,000円", sample_catalog_prices
    )
    assert status == "PRICE_MISMATCH"
    assert p_incl == 12100.0
    assert p_excl == 11000.0


def test_compare_single_item_code_not_found(
    sample_catalog_prices: Dict[str, Dict[str, Optional[float]]],
) -> None:
    """Test code not found status for invalid or unassigned codes."""
    status, p_incl, p_excl = compare_single_item(
        "GLOBAL", "12,100円", sample_catalog_prices
    )
    assert status == "CODE_NOT_FOUND"
    assert p_incl is None

    status_missing, _, _ = compare_single_item(
        "UNKNOWN-999", "10,000円", sample_catalog_prices
    )
    assert status_missing == "CODE_NOT_FOUND"
