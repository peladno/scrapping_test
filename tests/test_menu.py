"""Unit tests for interactive terminal menu."""

from unittest.mock import MagicMock, patch

from core.menu import (
    launch_interactive_menu,
    prompt_platform_choice,
    show_system_status,
)


def test_show_system_status() -> None:
    """Test show_system_status prints status without crashing."""
    show_system_status()


def test_prompt_platform_choice_exit() -> None:
    """Test prompt_platform_choice handles cancel option '0'."""
    with patch("builtins.input", return_value="0"):
        res = prompt_platform_choice("TEST")
        assert res is None


def test_prompt_platform_choice_all() -> None:
    """Test prompt_platform_choice handles '1' (all)."""
    with patch("builtins.input", return_value="1"):
        res = prompt_platform_choice("TEST")
        assert res == "all"


def test_prompt_platform_choice_specific() -> None:
    """Test prompt_platform_choice handles specific platform number."""
    with patch("builtins.input", return_value="2"):
        res = prompt_platform_choice("TEST")
        assert res == "rakuten"


def test_launch_interactive_menu_exit() -> None:
    """Test launch_interactive_menu exits cleanly on '0'."""
    mock_pipeline = MagicMock()
    with patch("builtins.input", return_value="0"):
        launch_interactive_menu(mock_pipeline)
    assert not mock_pipeline.called


def test_launch_interactive_menu_run_all() -> None:
    """Test launch_interactive_menu runs full pipeline on option '1'."""
    mock_pipeline = MagicMock()
    # 1 -> run all, then enter on pause_prompt, then 0 -> exit
    inputs = iter(["1", "", "0"])
    with patch("builtins.input", side_effect=lambda _: next(inputs)):
        launch_interactive_menu(mock_pipeline)
    mock_pipeline.assert_called_once_with("all", True, True)
