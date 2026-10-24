from __future__ import annotations

import json
import sys
from pathlib import Path

import pdfplumber
import pypdf
import reportlab
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer


def env() -> int:
    payload = {
        "python": sys.executable,
        "pypdf": getattr(pypdf, "__version__", None),
        "pdfplumber": getattr(pdfplumber, "__version__", None),
        "reportlab": getattr(reportlab, "Version", None),
        "cjk_font": "STSong-Light",
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def create_from_text(input_path: Path, output_path: Path) -> int:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    styles = getSampleStyleSheet()
    body = styles["BodyText"]
    body.fontName = "STSong-Light"
    body.fontSize = 11
    body.leading = 16

    lines = input_path.read_text(encoding="utf-8").splitlines()
    document = SimpleDocTemplate(str(output_path), pagesize=A4)
    story = []
    for line in lines:
        text = line if line.strip() else " "
        story.append(Paragraph(text, body))
        story.append(Spacer(1, 10))
    document.build(story)
    print(output_path)
    return 0


def merge(output_path: Path, input_paths: list[Path]) -> int:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    writer = PdfWriter()
    for input_path in input_paths:
        reader = PdfReader(str(input_path))
        for page in reader.pages:
            writer.add_page(page)
    with output_path.open("wb") as file_obj:
        writer.write(file_obj)
    print(output_path)
    return 0


def split(input_path: Path, output_dir: Path) -> int:
    output_dir.mkdir(parents=True, exist_ok=True)
    reader = PdfReader(str(input_path))
    written = []
    for index, page in enumerate(reader.pages, start=1):
        writer = PdfWriter()
        writer.add_page(page)
        output_path = output_dir / f"page_{index:03d}.pdf"
        with output_path.open("wb") as file_obj:
            writer.write(file_obj)
        written.append(str(output_path))
    print(json.dumps(written, ensure_ascii=False, indent=2))
    return 0


def validate(path: Path) -> int:
    reader = PdfReader(str(path))
    first_page_text = reader.pages[0].extract_text() or ""
    payload = {
        "pages": len(reader.pages),
        "first_page_preview": first_page_text[:120],
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def extract_text(path: Path) -> int:
    texts = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            texts.append(page.extract_text() or "")
    print("\n".join(texts))
    return 0


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: pdf_windows.py {env|create-from-text|merge|split|validate|extract-text} [args]")
        return 1

    command = sys.argv[1]
    if command == "env":
        return env()
    if command == "create-from-text":
        if len(sys.argv) < 4:
            print("create-from-text requires input.txt and output.pdf")
            return 1
        return create_from_text(Path(sys.argv[2]), Path(sys.argv[3]))
    if command == "merge":
        if len(sys.argv) < 5:
            print("merge requires output.pdf and at least two input PDFs")
            return 1
        return merge(Path(sys.argv[2]), [Path(arg) for arg in sys.argv[3:]])
    if command == "split":
        if len(sys.argv) < 4:
            print("split requires input.pdf and output_dir")
            return 1
        return split(Path(sys.argv[2]), Path(sys.argv[3]))
    if command == "validate":
        if len(sys.argv) < 3:
            print("validate requires file path")
            return 1
        return validate(Path(sys.argv[2]))
    if command == "extract-text":
        if len(sys.argv) < 3:
            print("extract-text requires file path")
            return 1
        return extract_text(Path(sys.argv[2]))

    print(f"Unknown command: {command}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
