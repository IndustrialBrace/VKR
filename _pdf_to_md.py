#!/usr/bin/env python3
"""Convert КБС 2026 (29.01).pdf to Markdown."""
import fitz  # PyMuPDF
import re
from pathlib import Path

SRC = Path("/projects/sandbox/VKR/КБС 2026 (29.01).pdf")
DST = Path("/projects/sandbox/VKR/КБС 2026 (29.01).md")


def clean_line(s: str) -> str:
    # Collapse multiple spaces
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


def page_to_md(page, page_num: int) -> str:
    """Extract text from a page preserving rough structure."""
    blocks = page.get_text("blocks")
    # Sort by Y then X
    blocks.sort(key=lambda b: (round(b[1], 1), round(b[0], 1)))

    parts = []
    for b in blocks:
        text = b[4] if len(b) > 4 else ""
        if not text or not text.strip():
            continue
        lines = [clean_line(ln) for ln in text.splitlines() if clean_line(ln)]
        if not lines:
            continue
        # Heuristic: short line in caps -> heading-like
        block_text = "\n".join(lines)
        parts.append(block_text)

    body = "\n\n".join(parts)
    return body


def main():
    doc = fitz.open(SRC)
    out_lines = []
    out_lines.append(f"# КБС 2026 (29.01)\n")
    out_lines.append(f"_Источник: {SRC.name}_  ")
    out_lines.append(f"_Всего страниц: {doc.page_count}_\n")

    for i, page in enumerate(doc, start=1):
        out_lines.append(f"\n---\n")
        out_lines.append(f"## Слайд {i}\n")
        text = page_to_md(page, i)
        if text.strip():
            out_lines.append(text)
        else:
            out_lines.append("_(на слайде не распознан текст — возможно, изображение)_")

    DST.write_text("\n".join(out_lines), encoding="utf-8")
    print(f"OK: {DST} ({DST.stat().st_size} bytes, {doc.page_count} pages)")


if __name__ == "__main__":
    main()
