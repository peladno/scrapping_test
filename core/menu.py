"""Interactive Terminal User Interface (TUI) Menu.

Provides an easy-to-use numbered menu in the terminal so users can execute
scraping, price comparisons, automated tests, and system status checks
without remembering CLI flags.
"""

from pathlib import Path
import subprocess
import sys
from typing import Callable, Optional

from core.config import (
    BASE_DIR,
    CATALOG_LIST_EXCEL,
    INPUT_DIR,
    OUTPUT_DIR,
    RAKUTEN_MASTER_EXCEL,
    YAHOO_MASTER_EXCEL,
)

PLATFORMS = [
    ("rakuten", "🔴 Rakuten Ichiba"),
    ("yahoo", "🟣 Yahoo Shopping"),
    ("amazon", "🟡 Amazon Japan"),
    ("yodobashi", "🔵 Yodobashi Camera"),
    ("aqua", "🟢 Import Shop Aqua"),
    ("furaipan", "🍳 Furaipan Club"),
]


def print_header(title: str) -> None:
    """Print a styled section header."""
    print("\n" + "=" * 66)
    print(f"  {title}")
    print("=" * 66)


def pause_prompt() -> None:
    """Pause and wait for the user to press Enter."""
    print()
    input("👉 Press [Enter] to return to the menu...")


def show_system_status() -> None:
    """Check and display the status of directories and input Excel files."""
    print_header("📁 SYSTEM & FILES STATUS")
    print(f"• Base Directory: {BASE_DIR}")
    print(f"• Input Directory: {INPUT_DIR} "
          f"({'EXISTS' if INPUT_DIR.exists() else 'MISSING'})")
    print(f"• Output Directory: {OUTPUT_DIR} "
          f"({'EXISTS' if OUTPUT_DIR.exists() else 'MISSING'})")
    print(f"• .env Configuration: "
          f"{'FOUND' if (BASE_DIR / '.env').exists() else 'NOT FOUND'}")
    print()
    print("Master Input Files:")

    input_files = [
        ("Product Catalog List", CATALOG_LIST_EXCEL),
        ("Rakuten Stores Master", RAKUTEN_MASTER_EXCEL),
        ("Yahoo Stores Master", YAHOO_MASTER_EXCEL),
    ]

    for label, path_str in input_files:
        p = Path(path_str)
        if p.exists():
            size_kb = p.stat().st_size / 1024.0
            print(f"  [OK] {label}: {p.name} ({size_kb:.1f} KB)")
        else:
            print(f"  [MISSING] {label}: {p.name} (Path: {p})")

    print("\nOutput Files in 'data/outputs/':")
    if OUTPUT_DIR.exists():
        outputs = list(OUTPUT_DIR.glob("*.xlsx"))
        if outputs:
            for f in outputs:
                size_kb = f.stat().st_size / 1024.0
                print(f"  - {f.name} ({size_kb:.1f} KB)")
        else:
            print("  (No output Excel files generated yet)")
    else:
        print("  (Output directory does not exist yet)")


def run_unit_tests() -> None:
    """Run the test suite using pytest if available."""
    print_header("🧪 RUNNING AUTOMATED UNIT TESTS")
    try:
        import pytest
        print("Executing pytest test suite...\n")
        exit_code = pytest.main(["-v", "tests"])
        if exit_code == 0:
            print("\n✅ All automated tests passed successfully!")
        else:
            print(f"\n❌ Tests completed with exit code: {exit_code}")
    except ImportError:
        # Fallback to subprocess if pytest is installed in Python environment
        try:
            res = subprocess.run(
                [sys.executable, "-m", "pytest", "-v", "tests"],
                check=False,
            )
            if res.returncode == 0:
                print("\n✅ All automated tests passed successfully!")
            else:
                print(f"\n❌ Tests exited with code: {res.returncode}")
        except Exception as exc:
            print(f"❌ Could not run pytest: {exc}")


def prompt_platform_choice(action_title: str) -> Optional[str]:
    """Display a platform selection submenu.

    Args:
        action_title: Title describing the action.

    Returns:
        Selected platform key string or None to cancel.
    """
    while True:
        print_header(f"SELECT PLATFORM - {action_title}")
        print("  [1] 🌐 ALL Platforms")
        for idx, (_, name) in enumerate(PLATFORMS, start=2):
            print(f"  [{idx}] {name}")
        print("  [0] ⬅️ Back to Main Menu")
        print("=" * 66)

        choice = input("👉 Enter your choice [0-7]: ").strip()
        if choice == "0":
            return None
        if choice == "1":
            return "all"

        try:
            val = int(choice)
            if 2 <= val <= len(PLATFORMS) + 1:
                return PLATFORMS[val - 2][0]
        except ValueError:
            pass

        print("⚠️ Invalid selection. Please choose a valid number.")


def launch_interactive_menu(
    run_pipeline_func: Callable[[str, bool, bool], None]
) -> None:
    """Display main control panel loop.

    Args:
        run_pipeline_func: Callback function (platform, scrape, compare)
                           that executes the chosen pipeline.
    """
    while True:
        print_header("🛒 E-COMMERCE SCRAPER & COMPARATOR - CONTROL PANEL")
        print("  [1] 🚀 Run FULL Pipeline (Scrape + Compare) - ALL Platforms")
        print("  [2] 🌐 Run FULL Pipeline for a Single Platform")
        print("  [3] 📥 Scrape Data Only (Skip Comparison)")
        print("  [4] 📊 Compare Prices Only (Existing Excel Files)")
        print("  [5] 🧪 Run Automated Tests (Pytest)")
        print("  [6] 📁 Check System & Excel Files Status")
        print("  [0] ❌ Exit")
        print("=" * 66)

        choice = input("👉 Select an option [0-6]: ").strip()

        if choice == "0":
            print("\n👋 Exiting E-Commerce Scraper. Goodbye!\n")
            break

        if choice == "1":
            print("\n🚀 Starting Full Pipeline for ALL platforms...")
            run_pipeline_func("all", True, True)
            pause_prompt()

        elif choice == "2":
            target = prompt_platform_choice("FULL PIPELINE (SCRAPE + COMPARE)")
            if target:
                print(f"\n🚀 Running Full Pipeline for: {target.upper()}...")
                run_pipeline_func(target, True, True)
                pause_prompt()

        elif choice == "3":
            target = prompt_platform_choice("SCRAPE DATA ONLY")
            if target:
                print(f"\n📥 Running Scrape-Only for: {target.upper()}...")
                run_pipeline_func(target, True, False)
                pause_prompt()

        elif choice == "4":
            target = prompt_platform_choice("COMPARE PRICES ONLY")
            if target:
                print(f"\n📊 Running Compare-Only for: {target.upper()}...")
                run_pipeline_func(target, False, True)
                pause_prompt()

        elif choice == "5":
            run_unit_tests()
            pause_prompt()

        elif choice == "6":
            show_system_status()
            pause_prompt()

        else:
            print(f"⚠️ Invalid choice '{choice}'. "
                  f"Please select between 0 and 6.")
