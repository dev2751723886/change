# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Qt6 C++ desktop app (PDFtoWord.exe) + Python conversion engine for multi-format document conversion. The C++ frontend calls Python via QProcess subprocess, communicating through JSON on stdout. The C++ source is **not** in this directory — only the compiled exe and DLLs. The Python script (`python/convert.py`) is the editable source.

Supported conversions: PDF ↔ Word ↔ Markdown ↔ JSON (10 paths), plus PDF split/merge/preview/info.

## Build / Run Commands

```bash
# Install Python dependencies
cd python
pip install -r requirements.txt

# Run any conversion from CLI (no GUI needed)
python python/convert.py convert "input.pdf" "output.docx"       # PDF → Word
python python/convert.py pdf2md "input.pdf" "output.md"           # PDF → Markdown
python python/convert.py word2pdf "input.docx" "output.pdf"       # Word → PDF
python python/convert.py md2pdf "input.md" "output.pdf"           # Markdown → PDF
python python/convert.py pdf2json "input.pdf" "output.json"       # PDF → JSON
python python/convert.py json2pdf "input.json" "output.pdf"       # JSON → PDF
python python/convert.py word2md "input.docx" "output.md"         # Word → Markdown
python python/convert.py md2word "input.md" "output.docx"         # Markdown → Word
python python/convert.py word2json "input.docx" "output.json"     # Word → JSON
python python/convert.py json2word "input.json" "output.docx"     # JSON → Word
python python/convert.py split "input.pdf" "output_dir" 5         # Split every 5 pages
python python/convert.py split "input.pdf" "output_dir" "" "1-3,5,7-10"  # Split by page ranges
python python/convert.py merge "merged.pdf" "file1.pdf" "file2.pdf"   # Merge PDFs
python python/convert.py preview "input.pdf" 5                    # Preview first 5 pages
python python/convert.py info "input.pdf"                         # Get page count / file size

# Launch GUI
PDFtoWord.exe
```

There is no test suite or linter configured.

## Architecture

**IPC protocol** — Qt6 calls `python convert.py <command> <args...>` via QProcess. Python prints a single JSON line to stdout: `{"success": true/false, ...}`. All errors return `{"success": false, "error": "<message>"}`. stderr is unused (qDebug goes to Qt console).

**Command routing** — `main()` (line 608) dispatches on `sys.argv[1]` to 15 branches. Each conversion function takes `(input_path, output_path)` and returns a dict.

**Key classes / functions:**

| Function | Lines | Purpose |
|----------|-------|---------|
| `convert_pdf_to_word()` | 93–108 | pdf2docx Converter wrapper |
| `convert_pdf_to_markdown()` | 111–146 | PyMuPDF dict extraction → MD with heading/bold inference |
| `convert_pdf_to_json()` | 251–289 | Deep structural extraction: page→block→line→span, includes color+bbox |
| `convert_word_to_pdf()` | 149–182 | docx2pdf primary, fitz PdfWriter fallback |
| `convert_word_to_markdown()` | 388–436 | python-docx style parsing → MD with bold/italic/table support |
| `convert_markdown_to_pdf()` | 185–247 | MD line-by-line parser → PdfWriter, supports 7 syntax types |
| `convert_markdown_to_word()` | 440–489 | MD → python-docx, regex split for inline **bold** and `code` |
| `PdfWriter` (class) | 33–77 | Custom PDF text layout engine: CJK font fallback, auto word-wrap via `text_length()`, auto page-break at 780pt |
| `_get_cjk_font()` | 21–30 | Font detection: msyh.ttc → simsun.ttc → simhei.ttf → helv |
| `split_pdf()` | 508–550 | Split by page count, page range, or every page |
| `merge_pdfs()` | 553–575 | Multi-PDF merge with page count tracking |

**Key dependencies:**
- `pdf2docx` — PDF → DOCX layout-preserving conversion
- `fitz` (PyMuPDF) — PDF read/create/merge/split/text-extract/TextWriter (used in 6 functions)
- `python-docx` — Word .docx read/write (paragraphs, styles, tables, runs)
- `docx2pdf` — Word → PDF (requires MS Word installed; falls back to fitz)

## File Structure

```
D:/文档转换工具/
├── PDFtoWord.exe          # Qt6 C++ GUI (compiled, source not here)
├── Qt6*.dll               # Qt6 runtime (Core, Gui, Widgets, Network, Svg)
├── platforms/qwindows.dll # Qt6 Windows platform plugin
├── python/
│   ├── convert.py         # All conversion logic (733 lines, single file)
│   └── requirements.txt   # Python dependencies
├── translations/*.qm      # Qt6 i18n files
└── *.txt                  # Project documentation (resume, interview Q&A)
```

## Notes

- `convert.py` is a single-file script — all 15 conversion/tool functions plus the CLI router live in one file. No classes/modules split.
- There is no C++ source in this directory. To modify the Qt6 GUI, the original `.pro`/`.cpp`/`.ui` project files would be needed (not present here).
- CJK font handling: `_get_cjk_font()` checks Windows font directory. On non-Windows systems, it falls back to Helvetica which cannot render Chinese — this would need porting for cross-platform.
- All conversion functions independently handle `os.makedirs(output_dir, exist_ok=True)`, so output directories are auto-created.
- The `convert_word_to_pdf()` function has a two-tier strategy: try `docx2pdf` (requires MS Word), fall back to `PdfWriter` (text-only, no images/tables).
