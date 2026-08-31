---
name: excel-dashboard-generator
description: Generate professional executive dashboards and advanced multi-sheet Excel reports with openpyxl, including KPI summary cards, price distribution charts, conditional formatting, and auto-adjusted column widths.
---

# Excel Dashboard Generator Skill

This skill guides the creation of executive-ready Excel reports and analytical dashboards using `openpyxl` and `pandas` for e-commerce price monitoring.

## 🎯 When to Use This Skill

- Adding an executive summary tab (KPIs, match percentages, violation summaries) to price comparison reports.
- Generating native Excel bar charts or pie charts showing price compliance by store.
- Applying advanced conditional formatting (color scales, status highlights, data bars).
- Auto-formatting table headers, numeric currency formatting (`¥#,##0`), and auto-adjusting column widths.

---

## 📐 Standard Dashboard Layout Pattern

When creating an Executive Dashboard sheet (`Resumen Ejecutivo` / `Summary`):

1. **KPI Summary Cards (Rows 2–5):**
   - **Total Products Scraped**
   - **Price Compliance Rate (%)** (GREEN matches / Total valid)
   - **Price Discrepancy Rate (%)** (RED mismatches / Total valid)
   - **Unassigned / Unmapped Products** (YELLOW count)
2. **Store Compliance Table (Rows 7–20):**
   - Columns: `Tienda`, `Total Productos`, `Coincidencias (Verde)`, `Diferencias (Rojo)`, `Sin Código (Amarillo)`, `% Cumplimiento`.
3. **Native Excel Chart:**
   - Bar chart or donut chart visualizing compliance percentage by store.

---

## 💻 Implementation Snippet (Reusable Pattern)

```python
"""Helper for generating an executive summary dashboard in openpyxl."""

from typing import Any, Dict
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


def apply_header_styling(ws: Worksheet, max_col: int, row: int = 1) -> None:
    """Applies modern dark-navy headers with bold white text."""
    header_fill = PatternFill(
        start_color="1F4E79", end_color="1F4E79", fill_type="solid"
    )
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align


def auto_fit_columns(ws: Worksheet, max_padding: int = 4) -> None:
    """Adjusts column widths based on cell content length."""
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or "")
            # Account for double-width characters (Japanese text)
            char_len = sum(2 if ord(c) > 127 else 1 for c in val_str)
            if char_len > max_len:
                max_len = char_len
        ws.column_dimensions[col_letter].width = min(max_len + max_padding, 60)


def create_kpi_card(
    ws: Worksheet,
    top_left_cell: str,
    title: str,
    value: Any,
    bg_color: str,
    text_color: str = "FFFFFF",
) -> None:
    """Creates a styled 2x2 metric KPI card in the worksheet."""
    # Top cell: Title, Bottom cell: Big numeric value
    # Example usage: create_kpi_card(ws, "B2", "Total Match", 842, "2E7D32")
    pass
```

---

## 🛡️ Best Practices

- Always check if the output file is currently locked/opened by Excel and use a fallback timestamp filename (`utils.save_excel_with_fallback`).
- Format currency cells explicitly with `cell.number_format = '"¥"#,##0'`.
- Ensure gridlines remain visible on dashboard tabs with `ws.views.sheetView[0].showGridLines = True`.
