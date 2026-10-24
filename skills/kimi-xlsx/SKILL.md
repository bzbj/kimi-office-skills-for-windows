---
name: kimi-xlsx
description: Generate and analyze Excel files on Windows computers using pandas and openpyxl. Preferred for spreadsheet analysis, formatting, and .xlsx output on Windows.
---

# Windows Excel Workflow

Use this skill for `.xlsx` generation, analysis, and formatting on Windows computers.

## CLI

Resolve this path relative to the directory containing this SKILL.md, not the current working directory:

`scripts/xlsx_windows.py`

Commands:

- `python <script> env`
- `python <script> create-from-csv <input.csv> <output.xlsx>`
- `python <script> validate <file.xlsx>`

## Runtime

- Python 3
- `pandas` for tabular analysis
- `openpyxl` for workbook styling, formulas, and verification

## Default approach

1. Load or build data with `pandas`
2. Export to `.xlsx`
3. Re-open with `openpyxl`
4. Apply formatting, widths, number formats, and alignment
5. Save and re-open for a quick smoke test when practical

## Chinese text

- `.xlsx` handles Chinese text correctly through Unicode
- Pay attention to column widths and sheet names, not encoding

## Preferred patterns

```python
import pandas as pd
from openpyxl import load_workbook

df = pd.DataFrame([{"部门": "销售", "销售额": 1000}])
df.to_excel("output.xlsx", index=False)

wb = load_workbook("output.xlsx")
ws = wb.active
ws.column_dimensions["A"].width = 12
ws.column_dimensions["B"].width = 14
wb.save("output.xlsx")
```

## Use cases that fit this skill

- Financial summaries
- Spreadsheet cleanup
- Data exports
- Styled analysis workbooks
- Multi-sheet reports

## Limits on Windows

- Do not rely on the bundled `KimiXlsx` Linux binary on Windows computers
- Validation should be done by reopening the workbook with `openpyxl` and checking key cells
