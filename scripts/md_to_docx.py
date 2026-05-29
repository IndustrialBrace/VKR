#!/usr/bin/env python3
"""
Minimal Markdown -> .docx converter for the practice report.
Pure-stdlib (no python-docx / pandoc available in sandbox).

Supports the subset used in `Отчёт по практике.md`:
- ATX headings (# ... ######) -> Heading 1..6
- Paragraphs with inline **bold**, *italic*, `code`
- Bullet lists (- item) and numbered lists (1. item) with simple flat structure
- Blockquotes (> ...)
- Tables `| ... | ... |` with `<br>` line breaks inside cells and **bold** spans
- Fenced code blocks ``` ... ```  -> preformatted monospace paragraphs
- Horizontal rules (--- on its own line)
"""

from __future__ import annotations

import re
import sys
import zipfile
from html import escape as _xml_escape
from pathlib import Path

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
W_NS_DECLS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
)


def xe(s: str) -> str:
    return _xml_escape(s, quote=True)


# --------------------------- inline run rendering ---------------------------

INLINE_RE = re.compile(
    r"(\*\*(?P<b>[^*]+?)\*\*)"
    r"|(\*(?P<i>[^*]+?)\*)"
    r"|(`(?P<c>[^`]+?)`)"
)


def runs_from_inline(text: str, *, base_bold: bool = False, base_italic: bool = False) -> str:
    """Convert inline markdown (bold/italic/code) to Word run XML."""
    out: list[str] = []
    pos = 0
    for m in INLINE_RE.finditer(text):
        if m.start() > pos:
            out.append(_run(text[pos:m.start()], bold=base_bold, italic=base_italic))
        if m.group("b") is not None:
            out.append(_run(m.group("b"), bold=True, italic=base_italic))
        elif m.group("i") is not None:
            out.append(_run(m.group("i"), bold=base_bold, italic=True))
        elif m.group("c") is not None:
            out.append(_run(m.group("c"), bold=base_bold, italic=base_italic, mono=True))
        pos = m.end()
    if pos < len(text):
        out.append(_run(text[pos:], bold=base_bold, italic=base_italic))
    return "".join(out)


def _run(text: str, *, bold: bool = False, italic: bool = False, mono: bool = False) -> str:
    if not text:
        return ""
    rpr_parts: list[str] = []
    if bold:
        rpr_parts.append("<w:b/>")
    if italic:
        rpr_parts.append("<w:i/>")
    if mono:
        rpr_parts.append('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>')
    rpr = f"<w:rPr>{''.join(rpr_parts)}</w:rPr>" if rpr_parts else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{xe(text)}</w:t></w:r>'


# --------------------------- block element builders ---------------------------

def p_heading(text: str, level: int) -> str:
    runs = runs_from_inline(text)
    return (
        f'<w:p><w:pPr><w:pStyle w:val="Heading{level}"/>'
        f'<w:keepNext/></w:pPr>{runs}</w:p>'
    )


def p_paragraph(text: str) -> str:
    runs = runs_from_inline(text)
    return f"<w:p>{runs}</w:p>"


def p_blockquote(text: str) -> str:
    runs = runs_from_inline(text, base_italic=True)
    return f'<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr>{runs}</w:p>'


def p_listitem(text: str, *, ordered: bool, lvl: int = 0) -> str:
    runs = runs_from_inline(text)
    num_id = 2 if ordered else 1
    return (
        f'<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>'
        f'<w:numPr><w:ilvl w:val="{lvl}"/><w:numId w:val="{num_id}"/></w:numPr>'
        f'</w:pPr>{runs}</w:p>'
    )


def p_code_line(text: str) -> str:
    """One line of a fenced code block as a separate paragraph (preserves layout)."""
    body = _run(text or " ", mono=True)
    return (
        '<w:p><w:pPr><w:pStyle w:val="CodeBlock"/></w:pPr>'
        f"{body}</w:p>"
    )


def p_hr() -> str:
    return (
        '<w:p><w:pPr><w:pBdr>'
        '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="808080"/>'
        '</w:pBdr></w:pPr></w:p>'
    )


def cell_paragraphs(cell_md: str) -> str:
    """Render a single table cell: split by <br> into paragraphs with inline formatting."""
    parts = re.split(r"<br\s*/?>", cell_md, flags=re.IGNORECASE)
    paragraphs: list[str] = []
    for part in parts:
        text = part.strip()
        runs = runs_from_inline(text) if text else _run(" ")
        paragraphs.append(
            f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>{runs}</w:p>'
        )
    if not paragraphs:
        paragraphs.append('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>')
    return "".join(paragraphs)


def build_table(rows: list[list[str]], *, has_header: bool) -> str:
    if not rows:
        return ""
    n_cols = max(len(r) for r in rows)
    col_w = max(1, 9000 // n_cols)  # twips inside 9000-twip total
    grid_cols = "".join(f'<w:gridCol w:w="{col_w}"/>' for _ in range(n_cols))
    borders = (
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:color="auto"/>'
        '</w:tblBorders>'
    )
    tbl_pr = (
        '<w:tblPr>'
        '<w:tblW w:w="9000" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        f'{borders}'
        '<w:tblLook w:val="04A0"/>'
        '</w:tblPr>'
    )
    tr_xml: list[str] = []
    for ri, row in enumerate(rows):
        is_header = has_header and ri == 0
        cells = list(row) + [""] * (n_cols - len(row))
        tc_xml: list[str] = []
        for cell in cells:
            shading = (
                '<w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/>' if is_header else ""
            )
            cell_inner = cell_paragraphs(cell)
            if is_header:
                # Make header text bold by wrapping inline content - simpler: post-process runs
                cell_inner = cell_inner.replace(
                    "<w:r><w:t",
                    '<w:r><w:rPr><w:b/></w:rPr><w:t',
                )
                # If a run already had rPr, also inject <w:b/> there
                cell_inner = re.sub(
                    r"<w:r><w:rPr>(?!<w:b/>)",
                    "<w:r><w:rPr><w:b/>",
                    cell_inner,
                )
            tc_xml.append(
                f'<w:tc><w:tcPr><w:tcW w:w="{col_w}" w:type="dxa"/>{shading}</w:tcPr>'
                f"{cell_inner}</w:tc>"
            )
        header_marker = (
            '<w:trPr><w:tblHeader/></w:trPr>' if is_header else ""
        )
        tr_xml.append(f"<w:tr>{header_marker}{''.join(tc_xml)}</w:tr>")
    return f"<w:tbl>{tbl_pr}{grid_cols}{''.join(tr_xml)}</w:tbl>"


# --------------------------- markdown block parser ---------------------------

HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
HR_RE = re.compile(r"^-{3,}\s*$")
ULI_RE = re.compile(r"^\s*[-*]\s+(.*)$")
OLI_RE = re.compile(r"^\s*\d+\.\s+(.*)$")
BQ_RE = re.compile(r"^>\s?(.*)$")
TABLE_SEP_RE = re.compile(r"^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)+\|?\s*$")
TABLE_ROW_RE = re.compile(r"^\s*\|.*\|\s*$")
FENCE_RE = re.compile(r"^```")


def split_table_row(line: str) -> list[str]:
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def md_to_body(md: str) -> str:
    lines = md.splitlines()
    out: list[str] = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]

        # fenced code
        if FENCE_RE.match(line):
            i += 1
            code: list[str] = []
            while i < n and not FENCE_RE.match(lines[i]):
                code.append(lines[i])
                i += 1
            if i < n:
                i += 1  # skip closing fence
            for cl in code:
                out.append(p_code_line(cl))
            continue

        # blank line
        if not line.strip():
            i += 1
            continue

        # horizontal rule
        if HR_RE.match(line):
            out.append(p_hr())
            i += 1
            continue

        # heading
        m = HEADING_RE.match(line)
        if m:
            level = len(m.group(1))
            out.append(p_heading(m.group(2).strip(), min(level, 6)))
            i += 1
            continue

        # table: detect header row + separator on next line
        if TABLE_ROW_RE.match(line) and i + 1 < n and TABLE_SEP_RE.match(lines[i + 1]):
            header = split_table_row(line)
            i += 2
            rows: list[list[str]] = [header]
            while i < n and TABLE_ROW_RE.match(lines[i]):
                rows.append(split_table_row(lines[i]))
                i += 1
            out.append(build_table(rows, has_header=True))
            continue

        # blockquote (possibly multi-line)
        if BQ_RE.match(line):
            buf: list[str] = []
            while i < n and (BQ_RE.match(lines[i]) or (lines[i].strip() == "" and buf)):
                m = BQ_RE.match(lines[i])
                if m:
                    buf.append(m.group(1))
                else:
                    buf.append("")
                i += 1
                if i < n and not BQ_RE.match(lines[i]) and lines[i].strip() != "":
                    break
            out.append(p_blockquote(" ".join(s for s in buf if s)))
            continue

        # numbered list
        if OLI_RE.match(line):
            while i < n and OLI_RE.match(lines[i]):
                m = OLI_RE.match(lines[i])
                out.append(p_listitem(m.group(1), ordered=True))
                i += 1
            continue

        # unordered list
        if ULI_RE.match(line):
            while i < n and ULI_RE.match(lines[i]):
                m = ULI_RE.match(lines[i])
                out.append(p_listitem(m.group(1), ordered=False))
                i += 1
            continue

        # paragraph (collect consecutive non-empty, non-special lines)
        para: list[str] = [line]
        i += 1
        while i < n:
            ln = lines[i]
            if (
                not ln.strip()
                or HEADING_RE.match(ln)
                or HR_RE.match(ln)
                or BQ_RE.match(ln)
                or ULI_RE.match(ln)
                or OLI_RE.match(ln)
                or FENCE_RE.match(ln)
                or TABLE_ROW_RE.match(ln)
            ):
                break
            para.append(ln)
            i += 1
        out.append(p_paragraph(" ".join(s.strip() for s in para)))

    return "".join(out)


# --------------------------- docx package files ---------------------------

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>
"""

SETTINGS_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {W}>
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="708"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>
"""


def styles_xml() -> str:
    heading_sizes = {1: 32, 2: 28, 3: 26, 4: 24, 5: 22, 6: 22}
    headings = []
    for lvl in range(1, 7):
        sz = heading_sizes[lvl]
        headings.append(f"""
  <w:style w:type="paragraph" w:styleId="Heading{lvl}">
    <w:name w:val="heading {lvl}"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="240" w:after="120"/>
      <w:outlineLvl w:val="{lvl - 1}"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:b/>
      <w:sz w:val="{sz}"/>
      <w:szCs w:val="{sz}"/>
    </w:rPr>
  </w:style>""")
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {W}>
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="28"/>
        <w:szCs w:val="28"/>
        <w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="276" w:lineRule="auto"/>
        <w:jc w:val="both"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  {''.join(headings)}
  <w:style w:type="paragraph" w:styleId="Quote">
    <w:name w:val="Quote"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="120" w:after="120"/>
      <w:ind w:left="720" w:right="720"/>
    </w:pPr>
    <w:rPr>
      <w:i/>
      <w:color w:val="595959"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:ind w:left="720"/>
      <w:contextualSpacing/>
    </w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CodeBlock">
    <w:name w:val="Code Block"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>
      <w:sz w:val="20"/>
      <w:szCs w:val="20"/>
    </w:rPr>
  </w:style>
</w:styles>
"""


NUMBERING_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering {W}>
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>
"""


def document_xml(body_inner: str) -> str:
    page = (
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" '
        'w:header="708" w:footer="708" w:gutter="0"/>'
        '<w:cols w:space="708"/>'
        '<w:docGrid w:linePitch="360"/>'
        '</w:sectPr>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document {W_NS_DECLS}>'
        f"<w:body>{body_inner}{page}</w:body>"
        '</w:document>'
    )


def build_docx(md_path: Path, out_path: Path) -> None:
    md = md_path.read_text(encoding="utf-8")
    body_inner = md_to_body(md)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", document_xml(body_inner))
        z.writestr("word/styles.xml", styles_xml())
        z.writestr("word/numbering.xml", NUMBERING_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)


def main() -> int:
    if len(sys.argv) != 3:
        print("usage: md_to_docx.py <input.md> <output.docx>", file=sys.stderr)
        return 2
    src = Path(sys.argv[1])
    dst = Path(sys.argv[2])
    build_docx(src, dst)
    print(f"wrote {dst} ({dst.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
