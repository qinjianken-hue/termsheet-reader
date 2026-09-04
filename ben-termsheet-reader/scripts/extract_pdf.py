# -*- coding: utf-8 -*-
"""Extract text from a termsheet PDF with pypdf (the Read tool can't render PDFs
here - poppler/pdftoppm is missing). Prints page-by-page, or writes to a .txt.

Usage:
    python extract_pdf.py <file.pdf> [out.txt]

Run with PYTHONIOENCODING=utf-8 for clean Unicode. For long termsheets, write to a
.txt and then grep/read the basket table, schedule, coupon and KI/KO sections.
"""
import sys
from pypdf import PdfReader


def main():
    if len(sys.argv) < 2:
        sys.exit("usage: python extract_pdf.py <file.pdf> [out.txt]")
    path = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None

    r = PdfReader(path)
    parts = []
    for i, pg in enumerate(r.pages, 1):
        t = pg.extract_text() or ""
        parts.append(f"\n========== PAGE {i} (chars={len(t)}) ==========\n{t}")
    txt = "".join(parts)

    if out:
        with open(out, "w", encoding="utf-8") as f:
            f.write(txt)
        print(f"{len(r.pages)} pages -> {out} ({len(txt)} chars)")
    else:
        print(f"PAGES: {len(r.pages)}")
        print(txt)


if __name__ == "__main__":
    main()
