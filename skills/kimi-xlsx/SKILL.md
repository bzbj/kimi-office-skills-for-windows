---
name: kimi-xlsx
description: Generate, edit, and analyze Excel files on Windows using pandas and openpyxl. Preserve existing cell formatting during data updates.
---

# Windows Excel Workflow

Use this skill for `.xlsx` generation, analysis, and formatting on Windows computers.

## Editing existing workbooks: preserve formatting

Apply these rules when updating data, filling columns, or extending an existing workbook, including when saving the edited workbook under a new filename. The creation and design defaults below apply to new workbooks; they do not authorize reformatting an existing workbook.

- Load and update the existing workbook. Do not rebuild it with `DataFrame.to_excel()`, replace an existing worksheet, or run blanket styling/autosizing helpers as part of a data refresh.
- For data updates and column fills, change only the target cell's `.value` to the requested numeric value or Excel formula. Keep numeric inputs numeric. Preserve the destination cell's font, fill, border, alignment, number format, and protection; do not copy source styles onto existing destination cells. Change existing formatting only when the user explicitly requests it.
- Keep existing row heights, column widths, merged ranges, conditional formatting, and worksheet layout unchanged unless a requested structural change requires adjusting the affected range. Preserve formulas outside the requested update.
- If extra rows or columns are needed, give each new cell the format of the nearest suitable cell in the same table section: use the corresponding cell in an adjacent data row for a new row, or an adjacent equivalent column for a new column. Match the role and data type; do not use a header, subtotal, or total as the template for a normal data cell.
- Copy styles only into new cells, using `copy.copy(template._style)` or copies of the individual style attributes. Match the relevant row height or column width for new rows/columns. Write the required values/formulas separately; do not duplicate template values or formulas unintentionally.
- When inserting rows/columns, update affected formula references, table ranges, conditional formatting, and data validation as needed. Do not assume row/column insertion adjusts every dependent reference automatically.
- Save and reopen the result. Verify the written values/formulas, check that existing cells retain their original formatting, and confirm that new cells match the selected neighboring templates. A successful open alone does not prove formatting was preserved.
- Use a method that retains the workbook's existing features. If an openpyxl round trip would discard unsupported objects, use Excel automation or a targeted OOXML edit instead of silently losing them.

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

## Creating a new workbook

Use this workflow only for new workbooks. For an existing workbook, follow the formatting-preserving rules above.

1. Load or build data with `pandas`
2. Export to `.xlsx`
3. Re-open with `openpyxl`
4. Apply formatting, widths, number formats, and alignment
5. Save and re-open for a quick smoke test when practical

## Chinese text

- `.xlsx` handles Chinese text correctly through Unicode
- Pay attention to column widths and sheet names, not encoding

## New workbook example

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
