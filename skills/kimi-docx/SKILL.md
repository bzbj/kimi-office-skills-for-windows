---
name: kimi-docx
description: Generate and edit Word documents (.docx) on Windows computers using Python. Use python-docx for document creation and lxml/docx XML edits when needed. Preferred for Word generation tasks on Windows.
---

# Windows Word Workflow

Use this skill for `.docx` creation or simple editing on Windows computers.

## CLI

Resolve this path relative to the directory containing this SKILL.md, not the current working directory:

`scripts/docx_windows.py`

Commands:

- `python <script> env`
- `python <script> create-report <input.json> <output.docx>`
- `python <script> validate <file.docx>`

## Runtime

- Python 3
- `python-docx` for document creation and structured edits
- `lxml` only when direct XML adjustments are needed

## Default approach

1. Create new documents with `python-docx`
2. Set East Asia fonts explicitly for Chinese text
3. Save as `.docx`
4. Re-open the file with `python-docx` for a smoke test when practical

## Chinese text requirements

- Always set an East Asia font for Chinese runs
- Prefer `Microsoft YaHei`, `SimSun`, or another installed CJK font
- Chinese content should be stored directly as Unicode text; do not use encoding workarounds

Example pattern:

```python
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

doc = Document()
p = doc.add_paragraph()
run = p.add_run("中文 English")
run.font.name = "Calibri"
run.font.size = Pt(11)
run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
doc.save("output.docx")
```

## Use cases that fit this skill

- Reports
- Meeting notes
- Analysis summaries
- Simple tables in Word
- Chinese/English mixed documents

## Limits on Windows

- Do not use the older Bash + .NET OpenXML pipeline from the previous version of this skill on this machine
- If a task requires advanced OpenXML-only features, edit the `.docx` package XML directly with `zipfile` + `lxml`
