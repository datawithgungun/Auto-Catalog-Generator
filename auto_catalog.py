""""
Automated Cataloguing System (PDF -> Excel)
-------------------------------------------
Scans a directory of PDF files and produces an Excel catalog with columns:
- Book Title
- Author
- Editor
- Year of Publishing
- Publisher
- Language
- Number of Pages (optional)
- Format (optional)  (always "PDF")

Dependencies (all open-source):
  pip install pymupdf pandas openpyxl

Optional (for scanned PDFs OCR):
  pip install pytesseract pdf2image
  + install system: Tesseract & Poppler

Run:
  python auto_catalog.py -i ./pdf -o catalog.xlsx
  python auto_catalog.py -i ./pdf -o catalog.xlsx --ocr
"""

from __future__ import annotations
import argparse, re
from pathlib import Path
from typing import Dict, List

import fitz    # PyMuPDF
import pandas as pd

# Optional OCR 
try:
    import pytesseract
    from pdf2image import convert_from_path
    OCR_AVAILABLE = True
except Exception:
    OCR_AVAILABLE = False

UNKNOWN = "Unknown"

def _clean_name(s: str) -> str:
    s = re.sub(r"[,;]\s*$", "", (s or "").strip())
    s = re.sub(r"\b(Dr\.?|Prof\.?|Professor|Ph\.?D\.?|M\.?D\.?)\b\.?", "", s, flags=re.I)
    return re.sub(r"\s{2,}", " ", s).strip() or UNKNOWN

def _front_text(doc: fitz.Document, pages: int = 8) -> str:
    out = []
    for i in range(min(pages, len(doc))):
        try:
            out.append(doc.load_page(i).get_text("text"))
        except Exception:
            pass
    return "\n".join(out).strip()

def _ocr_first_pages(pdf: Path, pages: int = 3) -> str:
    if not OCR_AVAILABLE:
        return ""
    try:
        txt = []
        for img in convert_from_path(str(pdf), first_page=1, last_page=pages):
            txt.append(pytesseract.image_to_string(img))
        return "\n".join(txt)
    except Exception:
        return ""

def _guess_title(lines: List[str]) -> str:
    cand = []
    for ln in lines[:25]:
        t = ln.strip()
        if not t or len(t) < 4: continue
        if re.search(r"\b(by|edited|editor|copyright|published|publisher|isbn|issn)\b", t, re.I): continue
        score = (2 if t.istitle() else 0) + (1 if (t.isupper() and len(t) <= 80) else 0)
        score += (1 if 10 <= len(t) <= 120 else 0)
        score += (1 if re.match(r"^[\w :;,'&\-\(\)\.\!\?]+$", t) else 0)
        if len(t) > 140: score -= 1
        if score >= 2: cand.append((score, t))
    return (sorted(cand, key=lambda x: (-x[0], len(x[1])))[0][1] if cand else UNKNOWN)

def _find_people(text: str, meta_author: str | None) -> tuple[str, str]:
    author = editor = UNKNOWN
    # author
    m = re.search(r"\b(?:By|Written by|Author(?:s)?:)\s*([^\n,]+)", text, flags=re.I)
    if m: author = _clean_name(m.group(1))
    elif meta_author and not re.search(r"unknown", meta_author or "", re.I):
        author = _clean_name(meta_author)
    # editor
    m = re.search(r"\b(?:Edited by|Editor(?:s)?:)\s*([^\n,]+)", text, flags=re.I)
    if m: editor = _clean_name(m.group(1))
    if author != UNKNOWN and editor.lower() == author.lower(): editor = UNKNOWN
    return author, editor

def _find_year(text: str) -> str:
    # Prefer years near Copyright/Published; else first plausible year 1500–2035
    year_pat = r"(?<!\d)(1[5-9]\d{2}|20[0-2]\d|203[0-5])(?!\d)"
    for term in ("©", "Copyright", "First published", "Published", "Reprinted"):
        m = re.search(term + r".{0,60}" + year_pat, text, flags=re.I | re.S)
        if m:
            y = re.search(year_pat, m.group(0))
            if y: return y.group(0)
    m = re.search(year_pat, text)
    return m.group(1) if m else UNKNOWN

def _find_publisher(text: str) -> str:
    for ln in text.splitlines():
        m = re.search(r"\b(Published by|Publisher|Imprint|Printed by)\s*[:\-]\s*(.+)", ln, flags=re.I)
        if m:
            cand = re.split(r"[;|•]|Tel:|Phone:|Fax:|Email:|www\.", m.group(2).strip())[0].strip()
            cand = re.sub(r"\s{2,}", " ", cand)
            if 2 <= len(cand) <= 120: return cand
    # fallback: a plausible single line
    for ln in text.splitlines()[:80]:
        if re.search(r"\b(Press|Publications|Publishers|University|House|Books)\b", ln, flags=re.I):
            s = ln.strip()
            if 2 <= len(s) <= 120: return s
    return UNKNOWN

def _guess_language(sample: str) -> str:
    if not sample or len(sample.strip()) < 40: return UNKNOWN
    counts = {k:0 for k in ("Devanagari","Bengali","Gurmukhi","Gujarati","Oriya","Tamil","Telugu","Kannada","Malayalam","Latin","Arabic","Cyrillic")}
    for ch in sample[:5000]:
        cp = ord(ch)
        if   0x0900 <= cp <= 0x097F: counts["Devanagari"] += 1
        elif 0x0980 <= cp <= 0x09FF: counts["Bengali"] += 1
        elif 0x0A00 <= cp <= 0x0A7F: counts["Gurmukhi"] += 1
        elif 0x0A80 <= cp <= 0x0AFF: counts["Gujarati"] += 1
        elif 0x0B00 <= cp <= 0x0B7F: counts["Oriya"] += 1
        elif 0x0B80 <= cp <= 0x0BFF: counts["Tamil"] += 1
        elif 0x0C00 <= cp <= 0x0C7F: counts["Telugu"] += 1
        elif 0x0C80 <= cp <= 0x0CFF: counts["Kannada"] += 1
        elif 0x0D00 <= cp <= 0x0D7F: counts["Malayalam"] += 1
        elif 0x0400 <= cp <= 0x04FF: counts["Cyrillic"] += 1
        elif 0x0600 <= cp <= 0x06FF: counts["Arabic"] += 1
        elif (0x0000 <= cp <= 0x024F): counts["Latin"] += 1
    script, n = max(counts.items(), key=lambda kv: kv[1])
    if n < 20: return UNKNOWN
    return {
        "Devanagari":"Hindi","Bengali":"Bengali","Gurmukhi":"Punjabi","Gujarati":"Gujarati","Oriya":"Odia",
        "Tamil":"Tamil","Telugu":"Telugu","Kannada":"Kannada","Malayalam":"Malayalam",
        "Arabic":"Arabic","Cyrillic":"Russian","Latin":"English"
    }.get(script, UNKNOWN)

# core
def parse_pdf(pdf_path: Path, use_ocr: bool) -> Dict[str, str]:
    try:
        with fitz.open(pdf_path) as doc:
            meta = doc.metadata or {}
            text = _front_text(doc, 8)
            if use_ocr and (not text or len(text) < 40):
                ocr = _ocr_first_pages(pdf_path, 3)
                if len(ocr) > len(text): text = ocr

            title = (meta.get("title") or "").strip()
            if not title or title.lower() in {"", "untitled", "unknown"}:
                title = _guess_title([ln for ln in text.splitlines() if ln.strip()]) or UNKNOWN

            author, editor = _find_people(text, meta.get("author"))
            year = _find_year(text)
            publisher = _find_publisher(text)

            lang_text = text if len(doc) <= 8 else _front_text(doc, min(12, len(doc)))
            language = _guess_language(lang_text)

            return {
                "Book Title": title or UNKNOWN,
                "Author": author or UNKNOWN,
                "Editor": editor or UNKNOWN,
                "Year of Publishing": year or UNKNOWN,
                "Publisher": publisher or UNKNOWN,
                "Language": language or UNKNOWN,
                "Number of Pages": str(doc.page_count) if doc.page_count else UNKNOWN,
                "Format": "PDF",
            }
    except Exception:
        return {
            "Book Title": UNKNOWN, "Author": UNKNOWN, "Editor": UNKNOWN,
            "Year of Publishing": UNKNOWN, "Publisher": UNKNOWN,
            "Language": UNKNOWN, "Number of Pages": UNKNOWN, "Format": "PDF",
        }

def build_catalog(input_dir: Path, output_xlsx: Path, use_ocr: bool) -> pd.DataFrame:
    pdfs = sorted(input_dir.rglob("*.pdf"), key=lambda p: p.name.lower())
    rows = []
    for p in pdfs:
        rec = parse_pdf(p, use_ocr)
        rec["Source File"] = str(p.resolve())
        rows.append(rec)
    df = pd.DataFrame(rows, columns=[
        "Book Title","Author","Editor","Year of Publishing","Publisher","Language","Number of Pages","Format","Source File"
    ])
    output_xlsx.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Catalog")
    return df

# cli
def main():
    ap = argparse.ArgumentParser(description="PDF -> Excel catalog")
    ap.add_argument("-i","--input", required=True, help="Folder containing PDFs")
    ap.add_argument("-o","--output", default="catalog.xlsx", help="Output Excel path")
    ap.add_argument("--ocr", action="store_true", help="Enable OCR fallback (requires pytesseract + pdf2image + system deps)")
    args = ap.parse_args()

    src = Path(args.input).expanduser().resolve()
    if not src.is_dir(): raise SystemExit(f"ERROR: Input directory not found: {src}")
    if args.ocr and not OCR_AVAILABLE:
        print("WARNING: --ocr requested but OCR deps not available. Proceeding without OCR.")

    out = Path(args.output).expanduser().resolve()
    df = build_catalog(src, out, args.ocr and OCR_AVAILABLE)
    print(f"Catalog saved to: {out}")
    print(f"Total PDFs: {len(df)}")

if __name__ == "__main__":
    main()
