#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Конвертер Word (.docx) → Markdown для файла «вкр_новая.docx».

В исходном документе все абзацы оформлены стилем Normal, поэтому структура
восстанавливается по эвристикам форматирования:

  * заголовки  — абзацы по центру и полужирные:
      – # (H1)  : ВВЕДЕНИЕ, ГЛАВА N. …, ЗАКЛЮЧЕНИЕ, СПИСОК…, ПРИЛОЖЕНИЯ;
      – ## (H2) : нумерованные параграфы вида «1.1.», «2.3.» и т. п.;
      – ### (H3): прочие центрированные полужирные строки
                  (Приложение А/Б/В, рубрики списка литературы «IV. …»);
  * маркированные списки — абзацы, начинающиеся с тире «–»/«-»/«•»;
  * таблицы — реальные таблицы Word → таблицы Markdown;
  * инлайн-форматирование **полужирный** и *курсив* берётся из run-ов.

Запускается без аргументов: вкр_новая.docx → вкр_новая.md.
"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.table import Table
from docx.text.paragraph import Paragraph

HERE = Path(__file__).parent
SRC = HERE / "вкр_новая.docx"
DST = HERE / "вкр_новая.md"

# Маркеры пунктов списка в начале абзаца.
_BULLET_RE = re.compile(r"^\s*[\u2013\u2014\u2012•\-]\s+")
# Нумерованный параграф: «1.1.», «2.3.», «1.1.1.» …
_H2_RE = re.compile(r"^\d+(?:\.\d+)+\.?\s")
# Заголовки верхнего уровня.
_H1_RE = re.compile(
    r"^(ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|ПРИЛОЖЕНИЯ|ГЛАВА\b|СПИСОК\b|ОГЛАВЛЕНИЕ|СОДЕРЖАНИЕ)"
)


def md_escape(text: str) -> str:
    """Экранировать символы, значимые для Markdown, в обычном тексте."""
    # Экранируем только звёздочки/подчёркивания/обратные кавычки и | вне таблиц
    # минимально, чтобы не ломать формулы и ссылки.
    return text.replace("\\", "\\\\").replace("|", "\\|")


def inline_from_runs(p: Paragraph) -> str:
    """Собрать строку Markdown из run-ов абзаца с учётом bold/italic."""
    parts: list[str] = []
    for r in p.runs:
        t = r.text
        if t == "":
            continue
        # Сохраняем ведущие/замыкающие пробелы вне маркеров форматирования.
        lead = t[: len(t) - len(t.lstrip())]
        trail = t[len(t.rstrip()):]
        core = t.strip()
        esc = core.replace("|", "\\|") if core else core
        if core:
            if r.bold and r.italic:
                esc = f"***{esc}***"
            elif r.bold:
                esc = f"**{esc}**"
            elif r.italic:
                esc = f"*{esc}*"
        parts.append(f"{lead}{esc}{trail}")
    out = "".join(parts)
    # Нормализуем кратные пробелы, сохраняя содержимое.
    out = re.sub(r"[ \t]{2,}", " ", out)
    return out.strip()


def classify_heading(p: Paragraph, text: str) -> int | None:
    """Вернуть уровень заголовка (1/2/3) либо None, если это не заголовок."""
    centered = p.alignment == WD_ALIGN_PARAGRAPH.CENTER
    has_bold = any(r.bold for r in p.runs if r.text.strip())
    all_bold = has_bold and all(
        (r.bold or not r.text.strip()) for r in p.runs
    )
    if not (centered and all_bold):
        return None
    if _H1_RE.match(text):
        return 1
    if _H2_RE.match(text):
        return 2
    return 3


def para_to_md(p: Paragraph) -> str | None:
    text = p.text.strip()
    if not text:
        return None

    level = classify_heading(p, text)
    if level is not None:
        return f"{'#' * level} {text}"

    # Маркированный список: заменяем ведущее тире на «- ».
    m = _BULLET_RE.match(text)
    if m:
        body = inline_from_runs(p)
        body = _BULLET_RE.sub("", body, count=1)
        return f"- {body}"

    return inline_from_runs(p) or md_escape(text)


def cell_text(cell) -> str:
    parts = [pp.text.strip() for pp in cell.paragraphs]
    parts = [x for x in parts if x]
    txt = "<br>".join(parts)
    return txt.replace("|", "\\|")


def table_to_md(tbl: Table) -> str:
    rows = tbl.rows
    if not rows:
        return ""
    n_cols = max(len(r.cells) for r in rows)
    lines: list[str] = []

    def fmt_row(cells: list[str]) -> str:
        cells = cells + [""] * (n_cols - len(cells))
        return "| " + " | ".join(cells) + " |"

    header = [cell_text(c) for c in rows[0].cells]
    lines.append(fmt_row(header))
    lines.append("| " + " | ".join(["---"] * n_cols) + " |")
    for r in rows[1:]:
        lines.append(fmt_row([cell_text(c) for c in r.cells]))
    return "\n".join(lines)


def iter_block_items(doc):
    """Итерировать абзацы и таблицы тела документа в порядке следования."""
    from docx.oxml.ns import qn

    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag
        if tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif tag == qn("w:tbl"):
            yield Table(child, doc)


def main() -> None:
    doc = Document(str(SRC))
    out: list[str] = []
    prev_was_list = False

    for block in iter_block_items(doc):
        if isinstance(block, Table):
            out.append("")
            out.append(table_to_md(block))
            out.append("")
            prev_was_list = False
            continue

        md = para_to_md(block)
        if md is None:
            continue

        is_list = md.startswith("- ")
        is_heading = md.startswith("#")

        if is_heading:
            out.append("")
            out.append(md)
            out.append("")
        elif is_list:
            # Списочные пункты идут подряд без пустых строк между ними.
            if not prev_was_list and out and out[-1] != "":
                out.append("")
            out.append(md)
        else:
            if prev_was_list:
                out.append("")
            out.append(md)
            out.append("")

        prev_was_list = is_list

    # Схлопываем тройные+ пустые строки до одной.
    text = "\n".join(out)
    text = re.sub(r"\n{3,}", "\n\n", text).strip() + "\n"
    DST.write_text(text, encoding="utf-8")

    n_lines = text.count("\n") + 1
    print(f"OK: {DST.name} ({DST.stat().st_size:,} байт, {n_lines} строк)")


if __name__ == "__main__":
    main()
