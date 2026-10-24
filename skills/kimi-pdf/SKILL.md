---
name: kimi-pdf
description: Create and process PDF files on Windows computers using Python and verified local tools. Use for PDF creation, extraction, merge, split, and related tasks on Windows.
---

# Windows PDF Workflow

Use this skill for PDF work on Windows computers.

## CLI

Resolve this path relative to the directory containing this SKILL.md, not the current working directory:

`scripts/pdf_windows.py`

Commands:

- `python <script> env`
- `python <script> create-from-text <input.txt> <output.pdf>`
- `python <script> merge <output.pdf> <input1.pdf> <input2.pdf> [...]`
- `python <script> split <input.pdf> <output_dir>`
- `python <script> validate <file.pdf>`
- `python <script> extract-text <file.pdf>`

## Task types

### PDF creation

Preferred order:

1. Create the source content first, usually as `.docx`, HTML, or structured data output
2. Convert to PDF only with a tool that is verified to exist locally
3. Re-open or inspect the PDF when practical

### PDF processing

Prefer Python-based libraries that are available locally, such as:

- `pypdf`
- `pdfplumber`
- `pikepdf`
- `PyMuPDF`

Check imports before choosing the library.

## Environment rules

- Do not assume Bash scripts or Unix shell pipelines exist
- Do not assume `pdf.sh`, Playwright, Chromium, Node, or LaTeX are installed unless verified in the current session
- Prefer Windows-friendly Python workflows

## Recommended workflow

1. Determine whether the task is creation or processing
2. Verify the local command or Python package needed for that route
3. For creation tasks, generate the source document first
4. Convert or process with the verified local tool
5. Validate the result when feasible

## Chinese text

- Chinese text is usually safest when authored first in `.docx` or HTML, then converted
- Do not assume every local PDF converter embeds CJK fonts correctly; verify output if Chinese rendering matters

## Suitable tasks

- Merge PDFs
- Split PDFs
- Extract text from PDFs
- Build a report in Word/HTML first, then export to PDF

## Limits on Windows

- The older Unix-oriented Bash route is not the default on this machine
- Only use Node, browser, or LaTeX conversion routes after explicit local verification
