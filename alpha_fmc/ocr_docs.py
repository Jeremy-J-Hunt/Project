#!/usr/bin/env python3
"""Create a searchable .txt next to every downloaded PDF.

Companies House PDFs are mostly scanned images, so this uses the embedded text
layer when there is one and falls back to OCR (pdftoppm + tesseract).

Usage: python3 alpha_fmc/ocr_docs.py [glob ...]   (default: companies/*/docs/*.pdf)
"""
import os
import subprocess
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def text_layer(pdf):
    try:
        import pypdf
        t = "\n".join((p.extract_text() or "") for p in pypdf.PdfReader(pdf).pages)
        return t if len(t.strip()) > 200 else ""
    except Exception:
        return ""


def ocr(pdf):
    with tempfile.TemporaryDirectory() as d:
        subprocess.run(["pdftoppm", "-r", "150", "-gray", "-png", str(pdf), f"{d}/p"], check=True)
        pages = []
        for img in sorted(Path(d).glob("p-*.png")):
            out = subprocess.run(["tesseract", str(img), "-", "--psm", "3"], capture_output=True,
                                 text=True, env={**os.environ, "OMP_THREAD_LIMIT": "1"}).stdout
            pages.append(f"\n\n===== page {len(pages) + 1} =====\n{out}")
        return "".join(pages)


ORDER = ["_CS01_", "_SH01_", "_NEWINC_", "_IN01_", "_AA_", "_MR01_", "_SH", "_RES", "_CERTNM", "_PSC"]


def priority(pdf):
    """Ownership evidence first (shareholder lists, allotments, accounts), articles last."""
    rank = next((i for i, k in enumerate(ORDER) if k in pdf.name), len(ORDER))
    return (rank, "actium" not in pdf.parent.parent.name, pdf.name)


def process(pdf):
    txt = pdf.with_suffix(".txt")
    if txt.exists():
        return
    t = text_layer(pdf) or ocr(pdf)
    txt.write_text(t)
    print(f"{len(t):>8}  {pdf.relative_to(ROOT)}", flush=True)


if __name__ == "__main__":
    pats = sys.argv[1:] or ["companies/*/docs/*.pdf"]
    pdfs = sorted({p for pat in pats for p in ROOT.glob(pat)}, key=priority)
    with ThreadPoolExecutor(4) as ex:
        list(ex.map(process, pdfs))
