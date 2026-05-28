#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конвертер 'ВКР ГОТОВОЕ.md' -> 'ВКР ГОТОВОЕ.docx'.

Использует только стандартную библиотеку Python (zipfile + xml).

Поддерживает:
- заголовки #/##/### -> Heading1/2/3 (16/15/14 пт, жирный);
- обычные абзацы (TNR 14 пт, межстрочный 1.5, отступ первой строки 1,25 см,
  выравнивание по ширине);
- маркированные списки '- ' / '* ';
- markdown-таблицы '| ... |' с разделителем '|---|---|';
- блочные цитаты '> ...' (курсивом, с отступом и левой полоской);
- кодовые блоки ```...``` (моноширинный, без отступа первой строки);
- inline **жирный** и `моно`;
- горизонтальный разделитель '---'.

Поля страницы: 30/10/20/20 мм (лево/право/верх/низ — как в требованиях).
"""
from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape


# ---------------------------------------------------------------------------
# Низкоуровневые помощники
# ---------------------------------------------------------------------------

def x(s: str) -> str:
    """Эскейп для XML."""
    return xml_escape(s, {'"': "&quot;", "'": "&apos;"})


INLINE_RE = re.compile(r"(\*\*[^*\n]+\*\*|`[^`\n]+`)")


def parse_inline(text: str):
    """Разбить строку на куски (текст, props={'bold','mono'})."""
    out = []
    pos = 0
    for m in INLINE_RE.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], {}))
        tok = m.group(0)
        if tok.startswith("**"):
            out.append((tok[2:-2], {"bold": True}))
        else:
            out.append((tok[1:-1], {"mono": True}))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], {}))
    if not out:
        out = [("", {})]
    return out


def make_run(text: str, *, bold=False, italic=False, mono=False,
             sz: int = 28, color: str | None = None) -> str:
    """w:r XML. sz — в полупунктах: 28 = 14 пт."""
    rpr = []
    if bold:
        rpr.append("<w:b/><w:bCs/>")
    if italic:
        rpr.append("<w:i/><w:iCs/>")
    font = "Courier New" if mono else "Times New Roman"
    rpr.append(
        f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" '
        f'w:cs="{font}" w:eastAsia="{font}"/>'
    )
    rpr.append(f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')
    if color:
        rpr.append(f'<w:color w:val="{color}"/>')
    rpr_xml = "<w:rPr>" + "".join(rpr) + "</w:rPr>"
    return f'<w:r>{rpr_xml}<w:t xml:space="preserve">{x(text)}</w:t></w:r>'


def runs_from_inline(text: str, *, base_bold=False, base_italic=False,
                     sz: int = 28) -> str:
    parts = []
    for t, props in parse_inline(text):
        parts.append(make_run(
            t,
            bold=base_bold or props.get("bold", False),
            italic=base_italic,
            mono=props.get("mono", False),
            sz=sz,
        ))
    return "".join(parts)


def normal_para(text: str, *, align: str = "both", indent_first=True,
                sz: int = 28, before: int = 0, after: int = 0,
                line: int = 360, base_bold=False, base_italic=False) -> str:
    """Обычный абзац основного текста (TNR 14, 1.5, отступ 1,25 см)."""
    ppr_parts = [f'<w:jc w:val="{align}"/>']
    if indent_first:
        ppr_parts.append('<w:ind w:firstLine="709"/>')  # 1.25 cm = 709 twips
    else:
        ppr_parts.append('<w:ind w:firstLine="0"/>')
    ppr_parts.append(
        f'<w:spacing w:before="{before}" w:after="{after}" '
        f'w:line="{line}" w:lineRule="auto"/>'
    )
    ppr = "<w:pPr>" + "".join(ppr_parts) + "</w:pPr>"
    runs = runs_from_inline(text, base_bold=base_bold,
                            base_italic=base_italic, sz=sz)
    return f"<w:p>{ppr}{runs}</w:p>"


def heading(text: str, level: int) -> str:
    sz = {1: 32, 2: 30, 3: 28}.get(level, 28)
    align = "center" if level == 1 else "left"
    ppr = (
        '<w:pPr>'
        f'<w:pStyle w:val="Heading{level}"/>'
        '<w:keepNext/><w:keepLines/>'
        '<w:spacing w:before="360" w:after="180" '
        'w:line="360" w:lineRule="auto"/>'
        f'<w:jc w:val="{align}"/>'
        '<w:ind w:firstLine="0"/>'
        f'<w:outlineLvl w:val="{level - 1}"/>'
        '</w:pPr>'
    )
    runs = make_run(text, bold=True, sz=sz)
    return f"<w:p>{ppr}{runs}</w:p>"


def hr_para() -> str:
    return (
        '<w:p><w:pPr>'
        '<w:pBdr><w:bottom w:val="single" w:sz="6" '
        'w:space="1" w:color="808080"/></w:pBdr>'
        '<w:spacing w:before="120" w:after="120"/>'
        '<w:ind w:firstLine="0"/>'
        '</w:pPr></w:p>'
    )


def code_para(line: str) -> str:
    ppr = (
        '<w:pPr>'
        '<w:spacing w:before="0" w:after="0" '
        'w:line="240" w:lineRule="auto"/>'
        '<w:ind w:firstLine="0"/>'
        '<w:jc w:val="left"/>'
        '</w:pPr>'
    )
    if line == "":
        return f"<w:p>{ppr}</w:p>"
    runs = make_run(line, mono=True, sz=20)  # 10 pt
    return f"<w:p>{ppr}{runs}</w:p>"


def quote_para(text: str) -> str:
    ppr = (
        '<w:pPr>'
        '<w:ind w:left="567" w:firstLine="0"/>'
        '<w:spacing w:before="120" w:after="120" '
        'w:line="320" w:lineRule="auto"/>'
        '<w:jc w:val="left"/>'
        '<w:pBdr><w:left w:val="single" w:sz="18" '
        'w:space="8" w:color="A0A0A0"/></w:pBdr>'
        '</w:pPr>'
    )
    runs = runs_from_inline(text, base_italic=True, sz=26)
    return f"<w:p>{ppr}{runs}</w:p>"


def bullet_para(text: str, level: int = 0) -> str:
    margin = 360 + level * 360
    ppr = (
        '<w:pPr>'
        f'<w:ind w:left="{margin}" w:hanging="284" w:firstLine="0"/>'
        '<w:spacing w:before="40" w:after="40" '
        'w:line="360" w:lineRule="auto"/>'
        '<w:jc w:val="both"/>'
        '</w:pPr>'
    )
    bullet = make_run("\u2022\u00a0", sz=28)  # • + nbsp
    runs = bullet + runs_from_inline(text, sz=28)
    return f"<w:p>{ppr}{runs}</w:p>"


def make_table(rows: list[list[str]]) -> str:
    if not rows:
        return ""
    n_cols = max(len(r) for r in rows)
    # Полезная ширина страницы A4: 210 - 30 - 10 = 170 мм = 9639 twips
    total = 9639
    col_w = total // n_cols
    grid = "<w:tblGrid>" + "".join(
        f'<w:gridCol w:w="{col_w}"/>' for _ in range(n_cols)
    ) + "</w:tblGrid>"
    tbl_pr = (
        "<w:tblPr>"
        f'<w:tblW w:w="{total}" w:type="dxa"/>'
        '<w:jc w:val="center"/>'
        "<w:tblBorders>"
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        "</w:tblBorders>"
        '<w:tblLayout w:type="fixed"/>'
        "</w:tblPr>"
    )
    rows_xml = []
    for r_idx, row in enumerate(rows):
        is_header = r_idx == 0
        cells = []
        for c_idx in range(n_cols):
            txt = row[c_idx] if c_idx < len(row) else ""
            tcpr = (
                f'<w:tcPr><w:tcW w:w="{col_w}" w:type="dxa"/>'
                '<w:vAlign w:val="center"/></w:tcPr>'
            )
            cell_ppr = (
                '<w:pPr>'
                '<w:spacing w:before="40" w:after="40" '
                'w:line="240" w:lineRule="auto"/>'
                '<w:ind w:firstLine="0"/>'
                f'<w:jc w:val="{"center" if is_header else "left"}"/>'
                '</w:pPr>'
            )
            runs = runs_from_inline(txt, base_bold=is_header, sz=24)  # 12 pt
            cell_p = f"<w:p>{cell_ppr}{runs}</w:p>"
            cells.append(f"<w:tc>{tcpr}{cell_p}</w:tc>")
        trpr = "<w:trPr>" + ("<w:tblHeader/>" if is_header else "") + "</w:trPr>"
        rows_xml.append("<w:tr>" + trpr + "".join(cells) + "</w:tr>")
    return "<w:tbl>" + tbl_pr + grid + "".join(rows_xml) + "</w:tbl>"


# ---------------------------------------------------------------------------
# Парсер markdown
# ---------------------------------------------------------------------------

BULLET_RE = re.compile(r"^(\s*)[-*]\s+(.+)$")
HEADING_RE = re.compile(r"^(#+)\s+(.*)$")
TABLE_SEP_RE = re.compile(r"^\|[\s\-:|]+\|$")


def parse_markdown(md: str):
    lines = md.splitlines()
    i, n = 0, len(lines)
    while i < n:
        line = lines[i]
        s = line.strip()

        # пустая строка
        if s == "":
            i += 1
            continue

        # горизонтальный разделитель
        if s == "---":
            yield hr_para()
            i += 1
            continue

        # кодовый блок
        if s.startswith("```"):
            j = i + 1
            block = []
            while j < n and not lines[j].strip().startswith("```"):
                block.append(lines[j])
                j += 1
            for cl in block:
                yield code_para(cl)
            i = j + 1
            continue

        # заголовок
        m = HEADING_RE.match(line)
        if m:
            level = min(len(m.group(1)), 3)
            yield heading(m.group(2).strip(), level)
            i += 1
            continue

        # блочная цитата (несколько подряд > строк)
        if s.startswith(">"):
            qparts = []
            while i < n and lines[i].strip().startswith(">"):
                qparts.append(re.sub(r"^>\s?", "", lines[i].rstrip()))
                i += 1
            # выкинуть пустые в начале/конце
            joined = "\n".join(qparts).strip()
            # Каждый параграф цитаты — отдельный quote_para.
            for piece in re.split(r"\n\s*\n", joined):
                piece = piece.replace("\n", " ").strip()
                if piece:
                    yield quote_para(piece)
            continue

        # таблица
        if s.startswith("|"):
            tbl_lines = []
            while i < n and lines[i].strip().startswith("|"):
                tbl_lines.append(lines[i].strip())
                i += 1
            rows = []
            for tl in tbl_lines:
                if TABLE_SEP_RE.match(tl):
                    continue
                cells = [c.strip() for c in tl.strip("|").split("|")]
                rows.append(cells)
            if rows:
                yield make_table(rows)
            continue

        # маркированный список
        m = BULLET_RE.match(line)
        if m:
            indent_lvl = len(m.group(1)) // 2
            yield bullet_para(m.group(2).strip(), level=indent_lvl)
            i += 1
            continue

        # обычный абзац — собираем подряд идущие «нестандартные» строки
        plines = [s]
        i += 1
        while i < n:
            ls = lines[i]
            ss = ls.strip()
            if (ss == "" or ss.startswith("#") or ss.startswith("|")
                    or ss.startswith(">") or ss.startswith("```")
                    or ss == "---" or BULLET_RE.match(ls)):
                break
            plines.append(ss)
            i += 1
        text = " ".join(plines).strip()
        if text:
            yield normal_para(text)


# ---------------------------------------------------------------------------
# Сборка docx (zip + OOXML)
# ---------------------------------------------------------------------------

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
</Relationships>
"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/>
        <w:sz w:val="28"/>
        <w:szCs w:val="28"/>
        <w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:line="360" w:lineRule="auto"/>
        <w:jc w:val="both"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:line="360" w:lineRule="auto"/>
      <w:ind w:firstLine="709"/>
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:sz w:val="28"/>
      <w:szCs w:val="28"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="360" w:after="180" w:line="360" w:lineRule="auto"/>
      <w:ind w:firstLine="0"/>
      <w:jc w:val="center"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:sz w:val="32"/>
      <w:szCs w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>
      <w:ind w:firstLine="0"/>
      <w:jc w:val="left"/>
      <w:outlineLvl w:val="1"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:sz w:val="30"/>
      <w:szCs w:val="30"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>
      <w:ind w:firstLine="0"/>
      <w:jc w:val="left"/>
      <w:outlineLvl w:val="2"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:sz w:val="28"/>
      <w:szCs w:val="28"/>
    </w:rPr>
  </w:style>
</w:styles>
"""

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="708"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
    <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
  </w:compat>
  <w:themeFontLang w:val="ru-RU"/>
</w:settings>
"""

FONT_TABLE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Times New Roman">
    <w:panose1 w:val="02020603050405020304"/>
    <w:charset w:val="CC"/>
    <w:family w:val="roman"/>
    <w:pitch w:val="variable"/>
  </w:font>
  <w:font w:name="Courier New">
    <w:panose1 w:val="02070309020205020404"/>
    <w:charset w:val="CC"/>
    <w:family w:val="modern"/>
    <w:pitch w:val="fixed"/>
  </w:font>
</w:fonts>
"""

CORE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>ВКР ГОТОВОЕ</dc:title>
  <dc:creator>Гайфуллина Эльвина Илфатовна</dc:creator>
  <cp:lastModifiedBy>Kiro</cp:lastModifiedBy>
</cp:coreProperties>
"""

APP_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>md_to_docx (Kiro)</Application>
</Properties>
"""

DOC_HEAD = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<w:body>'
)

# Поля A4: верх 20 мм, низ 20 мм, лево 30 мм, право 10 мм.
# 1 мм = 56.6929 twips. 20→1134, 30→1701, 10→567.
SECT_PR = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1134" w:right="567" w:bottom="1134" w:left="1701" '
    'w:header="708" w:footer="708" w:gutter="0"/>'
    '<w:cols w:space="708"/>'
    '<w:docGrid w:linePitch="360"/>'
    '</w:sectPr>'
)

DOC_TAIL = SECT_PR + "</w:body></w:document>"


def build_document(md_text: str) -> str:
    body_parts = list(parse_markdown(md_text))
    return DOC_HEAD + "".join(body_parts) + DOC_TAIL


def write_docx(out_path: Path, doc_xml: str) -> None:
    files = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "word/_rels/document.xml.rels": DOC_RELS,
        "word/document.xml": doc_xml,
        "word/styles.xml": STYLES,
        "word/settings.xml": SETTINGS,
        "word/fontTable.xml": FONT_TABLE,
        "docProps/core.xml": CORE_XML,
        "docProps/app.xml": APP_XML,
    }
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main() -> int:
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
        "/projects/sandbox/VKR/ВКР ГОТОВОЕ.md")
    dst = Path(sys.argv[2]) if len(sys.argv) > 2 else src.with_suffix(".docx")
    md = src.read_text(encoding="utf-8")
    doc = build_document(md)
    write_docx(dst, doc)
    print(f"Wrote: {dst}  ({dst.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
