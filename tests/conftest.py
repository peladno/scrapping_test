"""Pytest fixtures for unit tests."""

from typing import Dict, List, Optional
import pytest


@pytest.fixture
def sample_catalog_codes() -> List[str]:
    """Sample list of official catalog product codes."""
    return [
        "GST-B46",
        "GST-A46",
        "GST-AS3",
        "IST-01",
        "IST-02",
        "G-46",
        "G-57",
        "GS-3",
        "GS-5",
        "G-2",
    ]


@pytest.fixture
def sample_catalog_prices() -> Dict[str, Dict[str, Optional[float]]]:
    """Sample catalog price mapping."""
    return {
        "G-46": {"tax_excluded": 11000.0, "tax_included": 12100.0},
        "GS-3": {"tax_excluded": 9000.0, "tax_included": 9900.0},
        "GST-B46": {"tax_excluded": 21000.0, "tax_included": 23100.0},
    }


@pytest.fixture
def yahoo_item_card_html() -> str:
    """Sample Yahoo Shopping HTML item card."""
    return """
    <div class="SearchResult_SearchResultItem__contents__WO5EH">
        <a class="SearchResult_SearchResultItem__detailLink__G4Top"
           href="https://store.shopping.yahoo.co.jp/eclity/ykk-16-g-46.html?sc_i=123">
            <h3 class="ItemTitle_SearchResultItemTitle__fy4bB LineClamp">
                包丁 GLOBAL グローバル 三徳 18cm G-46
            </h3>
        </a>
        <div class="ItemBrand_SearchResultItemBrand__text__YpTUv">
            GLOBAL
        </div>
        <div class="ItemPrice_ItemPrice__2t7fx">
            12,100円
        </div>
        <div class="ItemPointModal_PointText__contents__LqdQ5">
            5%（555pt）
        </div>
    </div>
    """
