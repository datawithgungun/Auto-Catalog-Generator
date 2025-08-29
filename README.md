# Automated Cataloguing System (PDF → Excel)

This tool scans a folder of PDF books and generates a structured Excel catalog with the following columns:

- **Book Title**
- **Author**
- **Editor**
- **Year of Publishing**
- **Publisher**
- **Language**
- **Number of Pages** (optional)
- **Format** (optional, always `PDF`)

If a field cannot be extracted confidently, the tool writes `Unknown`.

## Quick Start

1. **Install Python 3.9+** and run:

```bash
pip install -r requirements.txt
```

> OCR for scanned PDFs is **optional** and requires:
> - System packages: **Tesseract** and **Poppler**
> - Python packages: `pytesseract` and `pdf2image` (already in requirements)
> - On Windows, install Poppler from: https://github.com/oschwartz10612/poppler-windows
> - On macOS: `brew install tesseract poppler`
> - On Linux (Debian/Ubuntu): `sudo apt-get install tesseract-ocr poppler-utils`

2. **Run the script:**

```bash
python auto_catalog.py --input /path/to/your/pdf/folder --output catalog.xlsx
```

Add `--ocr` if your PDFs are scanned images:

```bash
python auto_catalog.py --input /path/to/your/pdf/folder --output catalog.xlsx --ocr
```

3. **Output:** An Excel file (default `catalog.xlsx`) with one row per book. A `Source File` column
is included to trace each entry back to its PDF path (helpful for QA).

## How It Works (Extraction Heuristics)

- **Title**: PDF metadata (if reliable). Otherwise, inferred from prominent lines on page 1 (skips lines containing `by`, `edited`, `copyright`, etc.).
- **Author / Editor**: Looks for phrases like `By <Name>`, `Author:`, `Edited by`, `Editor:` in the first pages; falls back to metadata author.
- **Year of Publishing**: Prefers years close to `©`, `Copyright`, `First published`, `Published`; otherwise first plausible 4‑digit year (1500–2035) found in front matter.
- **Publisher**: Matches lines like `Published by:`, `Publisher:`, `Imprint:`; otherwise heuristically picks a line with words like `Press`, `Publications`, `University`, `Books`.
- **Language**: Uses `langdetect` over text from up to the first 12 pages (if available). If too little text, returns `Unknown`.
- **Pages**: From the PDF itself.
- **Format**: Always `PDF`.

> Note: Publishing data inside PDFs is not standardized. These heuristics aim to be robust across many layouts but will not be perfect for all files. The script is designed to be _fail‑soft_: when unsure, it outputs `Unknown`.

##🚀 Features

Extracts text from digital and scanned PDFs
OCR support using Tesseract
Creates a clean Excel catalog (catalog.xlsx)
Command-line interface with customizable input/output paths

## Tips for Better Results

- Provide text‑based PDFs when possible. For scans, use `--ocr` (slower).
- Keep title pages unobstructed: avoid heavy watermarks on page 1–3 if you want better title/author detection.
- If you know the language in advance (e.g., the entire batch is English), you can later fill the `Language` column in Excel.

## Project Structure

```
auto-catalog/
│── auto_catalog.py       # Main script
│── requirements.txt      # Dependencies
│── README.md             # Documentation
│── pdf/                  # Folder to store PDF files
│── catalog.xlsx          # Output Excel file (generated)

```
##📸 Screenshots
![Excel View](https://github.com/user-attachments/assets/15a171bc-b9af-4737-873f-fa7c184bf880)
![Terminal command run](https://github.com/user-attachments/assets/e299cfff-201b-455f-987f-f3e6808ac860)


## License

MIT – Feel free to use and modify.
