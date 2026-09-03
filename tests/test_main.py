"""Unit tests for main.py CLI orchestrator."""

from unittest.mock import MagicMock, patch

import pytest

from main import build_cli_parser, main


def test_build_cli_parser_defaults() -> None:
    """Test default CLI parser values."""
    parser = build_cli_parser()
    args = parser.parse_args([])
    assert args.platform == "all"
    assert args.scrape_only is False
    assert args.compare_only is False


def test_build_cli_parser_custom_platform() -> None:
    """Test parsing custom platform argument."""
    parser = build_cli_parser()
    args = parser.parse_args(["-p", "rakuten", "--scrape-only"])
    assert args.platform == "rakuten"
    assert args.scrape_only is True
    assert args.compare_only is False


def test_build_cli_parser_mutually_exclusive() -> None:
    """Test mutually exclusive scrape-only and compare-only flags."""
    parser = build_cli_parser()
    with pytest.raises(SystemExit):
        parser.parse_args(["--scrape-only", "--compare-only"])


@patch("main.run_yodobashi_pipeline")
@patch("main.run_amazon_pipeline")
@patch("main.run_rakuten_pipeline")
@patch("main.run_yahoo_pipeline")
def test_main_runs_all_by_default(
    mock_yahoo: MagicMock,
    mock_rakuten: MagicMock,
    mock_amazon: MagicMock,
    mock_yodobashi: MagicMock,
) -> None:
    """Test main runs all pipelines by default with scrape and compare."""
    exit_code = main([])
    assert exit_code == 0
    assert mock_rakuten.called
    assert mock_yahoo.called
    assert mock_amazon.called
    assert mock_yodobashi.called


@patch("main.run_yodobashi_pipeline")
@patch("main.run_amazon_pipeline")
@patch("main.run_rakuten_pipeline")
@patch("main.run_yahoo_pipeline")
def test_main_runs_rakuten_only(
    mock_yahoo: MagicMock,
    mock_rakuten: MagicMock,
    mock_amazon: MagicMock,
    mock_yodobashi: MagicMock,
) -> None:
    """Test main only runs Rakuten when specified."""
    exit_code = main(["--platform", "rakuten", "--compare-only"])
    assert exit_code == 0
    mock_rakuten.assert_called_once_with(
        scrape=False, compare=True
    )
    assert not mock_yahoo.called
    assert not mock_amazon.called
    assert not mock_yodobashi.called


@patch("main.run_yodobashi_pipeline")
@patch("main.run_amazon_pipeline")
@patch("main.run_rakuten_pipeline")
@patch("main.run_yahoo_pipeline")
def test_main_runs_amazon_only(
    mock_yahoo: MagicMock,
    mock_rakuten: MagicMock,
    mock_amazon: MagicMock,
    mock_yodobashi: MagicMock,
) -> None:
    """Test main only runs Amazon when specified."""
    exit_code = main(["--platform", "amazon", "--scrape-only"])
    assert exit_code == 0
    mock_amazon.assert_called_once_with(
        scrape=True, compare=False
    )
    assert not mock_rakuten.called
    assert not mock_yahoo.called
    assert not mock_yodobashi.called


@patch("main.run_yodobashi_pipeline")
@patch("main.run_amazon_pipeline")
@patch("main.run_rakuten_pipeline")
@patch("main.run_yahoo_pipeline")
def test_main_runs_yodobashi_only(
    mock_yahoo: MagicMock,
    mock_rakuten: MagicMock,
    mock_amazon: MagicMock,
    mock_yodobashi: MagicMock,
) -> None:
    """Test main only runs Yodobashi when specified."""
    exit_code = main(["--platform", "yodobashi"])
    assert exit_code == 0
    mock_yodobashi.assert_called_once_with(
        scrape=True, compare=True
    )
    assert not mock_rakuten.called
    assert not mock_yahoo.called
    assert not mock_amazon.called
