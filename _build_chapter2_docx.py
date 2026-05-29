#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сборка главы 2 ВКР «Разработка стратегии продвижения инновационного продукта
Face2 Ак Барс Банка» в Word-документ (.docx) с оформлением по Положению о ВКР
ИУЭиФ КФУ:

  - формат A4, поля 30/10/20/20 мм;
  - основной текст — Times New Roman 14 пт, междустрочный 1,5,
    абзацный отступ 1,25 см, выравнивание по ширине;
  - заголовки глав — заглавными буквами, по центру, полужирный 14 пт;
  - заголовки параграфов — строчными (кроме первой), по центру, полужирный 14 пт;
  - таблицы — Times New Roman 12 пт, междустрочный одинарный;
    «Таблица N» в правом верхнем углу, название по центру,
    источник — под таблицей (12 пт);
  - рисунки — нативная векторная графика DrawingML, источник над подписью,
    подпись «Рис. N. Название» по центру, без точки в конце.

Скрипт собирает .docx как ZIP с OOXML, без зависимости от внешних библиотек
(python-docx, matplotlib и др. недоступны в песочнице).
"""

from __future__ import annotations

import io
import re
import zipfile
from html import escape as _xe
from pathlib import Path

# ---------------------------------------------------------------------------
# Константы геометрии и стиля
# ---------------------------------------------------------------------------

# Twentieths of a point ("twips"): 1 cm = 567 twips (т.к. 1 inch = 1440 twips).
TWIPS_PER_CM = 567
# English Metric Units (для DrawingML): 1 cm = 360 000 EMU.
EMU_PER_CM = 360000
EMU_PER_INCH = 914400

# Размеры страницы A4 в twips.
PAGE_W_TW = 11906   # 21,0 см
PAGE_H_TW = 16838   # 29,7 см

# Поля 30/10/20/20 мм по требованию п. 10.1 Положения.
MARGIN_LEFT_TW = 30 * TWIPS_PER_CM // 10   # 30 мм
MARGIN_RIGHT_TW = 10 * TWIPS_PER_CM // 10  # 10 мм
MARGIN_TOP_TW = 20 * TWIPS_PER_CM // 10    # 20 мм
MARGIN_BOTTOM_TW = 20 * TWIPS_PER_CM // 10 # 20 мм

# Содержательная ширина страницы.
CONTENT_W_TW = PAGE_W_TW - MARGIN_LEFT_TW - MARGIN_RIGHT_TW  # 9 638 twips ≈ 17 см
CONTENT_W_CM = CONTENT_W_TW / TWIPS_PER_CM
CONTENT_W_EMU = int(CONTENT_W_CM * EMU_PER_CM)

# Шрифт.
FONT_MAIN = "Times New Roman"
SZ_BODY_HP = 28        # 14 pt в half-points (sz = pt × 2)
SZ_TABLE_HP = 24       # 12 pt
SZ_FOOTNOTE_HP = 20    # 10 pt — для номера рисунка/«Источник»

# Цвета.
CLR_BLACK = "000000"
CLR_DARKBLUE = "1F4E79"
CLR_BLUE = "2E75B6"
CLR_LIGHT_BLUE = "D9E2F3"
CLR_ORANGE = "ED7D31"
CLR_GREEN = "70AD47"
CLR_GRAY = "7F7F7F"
CLR_LIGHT_GRAY = "F2F2F2"

# Абзацный отступ 1,25 см → twips.
INDENT_FIRST_TW = int(1.25 * TWIPS_PER_CM)

# Межстрочный 1,5 — в Word lineRule="auto" с line=360 даёт ровно 1,5×.
LINE_BODY = 360
LINE_SINGLE = 240

# Spacing before/after для заголовков (в twips: 1 pt = 20 twips).
SPC_BEFORE_HEAD = 240
SPC_AFTER_HEAD = 120


# ---------------------------------------------------------------------------
# OOXML-хелперы
# ---------------------------------------------------------------------------

def xe(text: str) -> str:
    """Эскейп для XML, с сохранением неразрывных пробелов и т.п."""
    if text is None:
        return ""
    return _xe(text, quote=False)


def run(text: str, *, sz_hp: int = SZ_BODY_HP, bold: bool = False,
        italic: bool = False, color: str = CLR_BLACK,
        font: str = FONT_MAIN, preserve_space: bool = True) -> str:
    """Inline run."""
    rpr = ['<w:rPr>']
    rpr.append(
        f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" '
        f'w:cs="{font}" w:eastAsia="{font}"/>'
    )
    if bold:
        rpr.append('<w:b/><w:bCs/>')
    if italic:
        rpr.append('<w:i/><w:iCs/>')
    rpr.append(f'<w:color w:val="{color}"/>')
    rpr.append(f'<w:sz w:val="{sz_hp}"/><w:szCs w:val="{sz_hp}"/>')
    rpr.append('<w:lang w:val="ru-RU"/>')
    rpr.append('</w:rPr>')
    space = ' xml:space="preserve"' if preserve_space else ""
    return (
        f'<w:r>{"".join(rpr)}'
        f'<w:t{space}>{xe(text)}</w:t></w:r>'
    )


def ppr(*, align: str = "both", indent_first: int = 0, indent_left: int = 0,
        line: int = LINE_BODY, line_rule: str = "auto",
        space_before: int = 0, space_after: int = 0,
        keep_next: bool = False, keep_lines: bool = False) -> str:
    """Параграфные настройки."""
    parts = ['<w:pPr>']
    if keep_next:
        parts.append('<w:keepNext/>')
    if keep_lines:
        parts.append('<w:keepLines/>')
    parts.append(
        f'<w:spacing w:before="{space_before}" w:after="{space_after}" '
        f'w:line="{line}" w:lineRule="{line_rule}"/>'
    )
    if indent_first or indent_left:
        attrs = []
        if indent_left:
            attrs.append(f'w:left="{indent_left}"')
        if indent_first:
            attrs.append(f'w:firstLine="{indent_first}"')
        parts.append(f'<w:ind {" ".join(attrs)}/>')
    parts.append(f'<w:jc w:val="{align}"/>')
    parts.append(
        '<w:rPr>'
        f'<w:rFonts w:ascii="{FONT_MAIN}" w:hAnsi="{FONT_MAIN}" '
        f'w:cs="{FONT_MAIN}" w:eastAsia="{FONT_MAIN}"/>'
        f'<w:sz w:val="{SZ_BODY_HP}"/><w:szCs w:val="{SZ_BODY_HP}"/>'
        '<w:lang w:val="ru-RU"/>'
        '</w:rPr>'
    )
    parts.append('</w:pPr>')
    return "".join(parts)


def para(content: str = "", *, align: str = "both",
         indent_first: int = INDENT_FIRST_TW, indent_left: int = 0,
         line: int = LINE_BODY, line_rule: str = "auto",
         space_before: int = 0, space_after: int = 0,
         keep_next: bool = False, keep_lines: bool = False) -> str:
    """Полный параграф из готового inline-контента (runs)."""
    return (
        '<w:p>'
        f'{ppr(align=align, indent_first=indent_first, indent_left=indent_left, line=line, line_rule=line_rule, space_before=space_before, space_after=space_after, keep_next=keep_next, keep_lines=keep_lines)}'
        f'{content}'
        '</w:p>'
    )


def empty_para(line: int = LINE_BODY) -> str:
    """Пустой абзац (для разделения блоков)."""
    return para("", indent_first=0, line=line)



# ---------------------------------------------------------------------------
# Текстовые блоки: заголовки, основной текст, ссылки
# ---------------------------------------------------------------------------

# Регулярное выражение для распознавания инлайнового форматирования.
# В исходнике используется только текст без bold/italic в основном корпусе
# главы 2, поэтому достаточно простой обработки.
_INLINE_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
_INLINE_ITALIC_RE = re.compile(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)")


def runs_from_text(text: str, *, sz_hp: int = SZ_BODY_HP) -> str:
    """Преобразует строку с **bold** и *italic* в последовательность runs."""
    out = []
    pos = 0
    # Сначала находим bold-фрагменты, затем italic в остатках.
    pieces: list[tuple[str, bool, bool]] = []  # (text, bold, italic)
    pos = 0
    text2 = text
    # Простая конечная state-машина: ищем по очереди.
    i = 0
    while i < len(text2):
        m_b = _INLINE_BOLD_RE.match(text2, i)
        if m_b:
            pieces.append((m_b.group(1), True, False))
            i = m_b.end()
            continue
        m_i = _INLINE_ITALIC_RE.match(text2, i)
        if m_i:
            pieces.append((m_i.group(1), False, True))
            i = m_i.end()
            continue
        # Найти ближайший спецсимвол.
        next_idx = len(text2)
        for needle in ("**", "*"):
            j = text2.find(needle, i)
            if j != -1 and j < next_idx:
                next_idx = j
        pieces.append((text2[i:next_idx], False, False))
        i = next_idx
    if not pieces:
        pieces = [(text, False, False)]
    for chunk, b, ital in pieces:
        if not chunk:
            continue
        out.append(run(chunk, sz_hp=sz_hp, bold=b, italic=ital))
    return "".join(out)


def chapter_heading(text: str) -> str:
    """Заголовок главы: ЗАГЛАВНЫМИ, по центру, полужирный 14 пт."""
    return para(
        run(text.upper(), bold=True),
        align="center", indent_first=0,
        space_before=SPC_BEFORE_HEAD, space_after=SPC_AFTER_HEAD,
        keep_next=True, keep_lines=True,
    )


def paragraph_heading(text: str) -> str:
    """Заголовок параграфа (#-уровня): по центру, полужирный 14 пт,
    строчные кроме первой буквы (предполагается, что заголовок уже
    оформлен правильно в исходнике)."""
    return para(
        run(text, bold=True),
        align="center", indent_first=0,
        space_before=SPC_BEFORE_HEAD, space_after=SPC_AFTER_HEAD,
        keep_next=True, keep_lines=True,
    )


def subsection_heading(text: str) -> str:
    """Подзаголовок (##-уровня): по центру, полужирный 14 пт."""
    return para(
        run(text, bold=True),
        align="center", indent_first=0,
        space_before=200, space_after=120,
        keep_next=True, keep_lines=True,
    )


def body_paragraph(text: str) -> str:
    """Основной абзац: TNR 14, 1,5 интервал, отступ 1,25 см, по ширине."""
    return para(
        runs_from_text(text),
        align="both",
        indent_first=INDENT_FIRST_TW,
        line=LINE_BODY, line_rule="auto",
    )


def page_break_para() -> str:
    """Принудительный разрыв страницы перед следующей главой."""
    return (
        '<w:p>'
        f'{ppr(indent_first=0)}'
        '<w:r><w:br w:type="page"/></w:r>'
        '</w:p>'
    )


# ---------------------------------------------------------------------------
# Таблицы (TNR 12 пт, одинарный)
# ---------------------------------------------------------------------------

def table_caption_block(number: str, title: str) -> str:
    """Шапка таблицы: «Таблица N» в правом верхнем углу + название по центру.

    Реализована как два абзаца над таблицей."""
    return (
        para(
            run(f"Таблица {number}", sz_hp=SZ_TABLE_HP),
            align="right", indent_first=0,
            space_before=120, space_after=0,
            keep_next=True, keep_lines=True,
            line=LINE_SINGLE,
        )
        +
        para(
            runs_from_text(title, sz_hp=SZ_TABLE_HP),
            align="center", indent_first=0,
            space_before=0, space_after=60,
            keep_next=True, keep_lines=True,
            line=LINE_SINGLE,
        )
    )


def table_source_block(text: str) -> str:
    """Подпись «Источник: ...» под таблицей. 12 пт, по левому краю."""
    # Курсив для слова "Источник"
    m = re.match(r"^(Источник):\s*(.*)$", text)
    if m:
        content = (
            run(f"{m.group(1)}: ", sz_hp=SZ_TABLE_HP, italic=True)
            + run(m.group(2), sz_hp=SZ_TABLE_HP)
        )
    else:
        content = run(text, sz_hp=SZ_TABLE_HP, italic=True)
    return para(
        content,
        align="left", indent_first=0,
        space_before=0, space_after=200,
        line=LINE_SINGLE,
    )


def figure_source_block(text: str) -> str:
    """Подпись «Источник: ...» над рисунком. 12 пт, по центру."""
    m = re.match(r"^(Источник):\s*(.*)$", text)
    if m:
        content = (
            run(f"{m.group(1)}: ", sz_hp=SZ_TABLE_HP, italic=True)
            + run(m.group(2), sz_hp=SZ_TABLE_HP)
        )
    else:
        content = run(text, sz_hp=SZ_TABLE_HP, italic=True)
    return para(
        content,
        align="center", indent_first=0,
        space_before=120, space_after=0,
        line=LINE_SINGLE,
        keep_next=True, keep_lines=True,
    )


def figure_caption_block(text: str) -> str:
    """Подпись рисунка по центру."""
    return para(
        run(text, bold=True, sz_hp=SZ_TABLE_HP),
        align="center", indent_first=0,
        space_before=60, space_after=200,
        line=LINE_SINGLE,
        keep_lines=True,
    )


def _cell_borders() -> str:
    """Тонкие чёрные границы у ячейки таблицы."""
    return (
        '<w:tcBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tcBorders>'
    )


def _cell_para(text: str, *, bold: bool = False, align: str = "left") -> str:
    """Параграф внутри ячейки. Поддерживает многострочный текст (через \n)."""
    lines = text.split("\n") if text else [""]
    paragraphs = []
    for ln in lines:
        # Также делим по " / " для длинных значений? — нет, оставляем как есть.
        ln = ln.strip()
        if not ln:
            ln = ""
        paragraphs.append(para(
            run(ln, sz_hp=SZ_TABLE_HP, bold=bold),
            align=align, indent_first=0,
            line=LINE_SINGLE, line_rule="auto",
            space_before=0, space_after=0,
        ))
    return "".join(paragraphs)


def table(rows: list[list[str]], *, col_widths_cm: list[float] | None = None,
          header_rows: int = 1) -> str:
    """Сгенерировать таблицу из 2D-списка строк.

    rows: [[c1,c2,...], ...]. Первые header_rows строк форматируются с заливкой
    шапки (полужирный, светло-серая заливка)."""
    n_cols = max(len(r) for r in rows) if rows else 1
    if col_widths_cm is None:
        col_widths_cm = [CONTENT_W_CM / n_cols] * n_cols
    col_widths_tw = [int(w * TWIPS_PER_CM) for w in col_widths_cm]
    total_w = sum(col_widths_tw)

    # tblGrid
    grid = ''.join(
        f'<w:gridCol w:w="{w}"/>' for w in col_widths_tw
    )

    # Свойства таблицы.
    tbl_pr = (
        '<w:tblPr>'
        f'<w:tblW w:w="{total_w}" w:type="dxa"/>'
        '<w:jc w:val="center"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tblBorders>'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" '
        'w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        '</w:tblPr>'
    )

    rows_xml = []
    for ridx, row in enumerate(rows):
        is_header = ridx < header_rows
        cells_xml = []
        # Дополним короткие строки пустыми ячейками.
        row = list(row) + [""] * (n_cols - len(row))
        for cidx, cell in enumerate(row):
            shading = (
                f'<w:shd w:val="clear" w:color="auto" w:fill="{CLR_LIGHT_GRAY}"/>'
                if is_header else ''
            )
            cells_xml.append(
                '<w:tc>'
                '<w:tcPr>'
                f'<w:tcW w:w="{col_widths_tw[cidx]}" w:type="dxa"/>'
                f'{_cell_borders()}'
                f'{shading}'
                '<w:vAlign w:val="center"/>'
                '</w:tcPr>'
                f'{_cell_para(str(cell), bold=is_header, align=("center" if is_header else "left"))}'
                '</w:tc>'
            )
        # Свойства строки: повторять шапку, не разрывать.
        row_pr = (
            '<w:trPr>'
            + ('<w:tblHeader/>' if is_header else '')
            + '<w:cantSplit/>'
            + '</w:trPr>'
        )
        rows_xml.append(f'<w:tr>{row_pr}{"".join(cells_xml)}</w:tr>')

    return (
        '<w:tbl>'
        f'{tbl_pr}'
        f'<w:tblGrid>{grid}</w:tblGrid>'
        f'{"".join(rows_xml)}'
        '</w:tbl>'
    )



# ---------------------------------------------------------------------------
# DrawingML — нативные векторные рисунки (полностью редактируемые в Word)
# ---------------------------------------------------------------------------

class DrawingIdGen:
    def __init__(self, start: int = 100):
        self._n = start

    def next(self) -> int:
        self._n += 1
        return self._n


def _txbody(paragraphs: list[tuple[str, dict]]) -> str:
    """Тело текстового блока внутри shape.
    paragraphs: [(text, {sz, bold, italic, color, align, font_size_hp})]"""
    parts = ['<wps:txbx><w:txbxContent>']
    for text, opts in paragraphs:
        sz_hp = opts.get("sz_hp", SZ_TABLE_HP)
        bold = opts.get("bold", False)
        italic = opts.get("italic", False)
        color = opts.get("color", CLR_BLACK)
        align = opts.get("align", "center")
        # multi-line: split by \n
        lines = text.split("\n")
        for ln in lines:
            parts.append(
                f'<w:p>{ppr(align=align, indent_first=0, line=LINE_SINGLE, space_before=0, space_after=0)}'
                f'{run(ln, sz_hp=sz_hp, bold=bold, italic=italic, color=color)}'
                f'</w:p>'
            )
    parts.append('</w:txbxContent></wps:txbx>')
    parts.append(
        '<wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="visible" '
        'horzOverflow="visible" wrap="square" lIns="36000" tIns="36000" '
        'rIns="36000" bIns="36000" anchor="ctr" anchorCtr="0" upright="1">'
        '<a:noAutofit/></wps:bodyPr>'
    )
    return "".join(parts)


def shape_rect(idg: DrawingIdGen, x_cm: float, y_cm: float,
               w_cm: float, h_cm: float, *,
               fill: str = "FFFFFF", line: str = CLR_BLACK,
               line_w: int = 9525, geom: str = "rect",
               text_paragraphs: list[tuple[str, dict]] | None = None,
               name: str = "Rect") -> str:
    """Прямоугольник/овал/скруглённый прямоугольник с текстом.
    geom: rect, roundRect, ellipse, plus, ..."""
    sid = idg.next()
    x = int(x_cm * EMU_PER_CM)
    y = int(y_cm * EMU_PER_CM)
    w = int(w_cm * EMU_PER_CM)
    h = int(h_cm * EMU_PER_CM)
    txbody = (
        _txbody(text_paragraphs) if text_paragraphs
        else '<wps:bodyPr/>'
    )
    return f'''<wps:wsp>
  <wps:cNvPr id="{sid}" name="{name}{sid}"/>
  <wps:cNvSpPr/>
  <wps:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
    <a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>
    <a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>
  </wps:spPr>
  {txbody}
</wps:wsp>'''


def shape_textbox(idg: DrawingIdGen, x_cm: float, y_cm: float,
                  w_cm: float, h_cm: float, *,
                  text_paragraphs: list[tuple[str, dict]],
                  name: str = "TextBox") -> str:
    """Прозрачный текстовый блок (без рамки)."""
    sid = idg.next()
    x = int(x_cm * EMU_PER_CM)
    y = int(y_cm * EMU_PER_CM)
    w = int(w_cm * EMU_PER_CM)
    h = int(h_cm * EMU_PER_CM)
    return f'''<wps:wsp>
  <wps:cNvPr id="{sid}" name="{name}{sid}"/>
  <wps:cNvSpPr txBox="1"/>
  <wps:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln><a:noFill/></a:ln>
  </wps:spPr>
  {_txbody(text_paragraphs)}
</wps:wsp>'''


def shape_line(idg: DrawingIdGen, x1_cm: float, y1_cm: float,
               x2_cm: float, y2_cm: float, *,
               color: str = CLR_BLACK, line_w: int = 12700,
               arrow_end: bool = False) -> str:
    """Прямая линия (между двумя точками)."""
    sid = idg.next()
    x = int(min(x1_cm, x2_cm) * EMU_PER_CM)
    y = int(min(y1_cm, y2_cm) * EMU_PER_CM)
    w = int(abs(x2_cm - x1_cm) * EMU_PER_CM)
    h = int(abs(y2_cm - y1_cm) * EMU_PER_CM)
    flip_attrs = []
    if x2_cm < x1_cm:
        flip_attrs.append('flipH="1"')
    if y2_cm < y1_cm:
        flip_attrs.append('flipV="1"')
    flip_str = (" " + " ".join(flip_attrs)) if flip_attrs else ""
    arrow_xml = (
        '<a:tailEnd type="triangle" w="med" len="med"/>' if arrow_end else ""
    )
    # Минимальный размер 1 EMU, чтобы Word не отбрасывал shape с нулевыми измерениями.
    if w == 0:
        w = 1
    if h == 0:
        h = 1
    return f'''<wps:wsp>
  <wps:cNvPr id="{sid}" name="Line{sid}"/>
  <wps:cNvCnPr/>
  <wps:spPr>
    <a:xfrm{flip_str}><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
    <a:ln w="{line_w}">
      <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
      {arrow_xml}
    </a:ln>
  </wps:spPr>
  <wps:bodyPr/>
</wps:wsp>'''


def drawing_canvas(width_cm: float, height_cm: float, content: str,
                   *, doc_id: int = 1) -> str:
    """Обёртка <w:drawing> с группой shapes (wpg:wgp) внутри инлайн-контекста.

    Используем wordprocessingGroup, так как Word более стабильно
    рендерит группы, чем canvas. Координаты shapes — абсолютные в EMU."""
    cx = int(width_cm * EMU_PER_CM)
    cy = int(height_cm * EMU_PER_CM)
    return f'''<w:r>
  <w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="{cx}" cy="{cy}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:docPr id="{doc_id}" name="Drawing{doc_id}"/>
      <wp:cNvGraphicFramePr/>
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
          <wpg:wgp>
            <wpg:cNvGrpSpPr/>
            <wpg:grpSpPr>
              <a:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="{cx}" cy="{cy}"/>
                <a:chOff x="0" y="0"/>
                <a:chExt cx="{cx}" cy="{cy}"/>
              </a:xfrm>
            </wpg:grpSpPr>
            {content}
          </wpg:wgp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>
</w:r>'''


def figure_drawing_para(width_cm: float, height_cm: float, content: str,
                        *, doc_id: int) -> str:
    """Параграф, содержащий рисунок (по центру, без отступов)."""
    return (
        '<w:p>'
        f'{ppr(align="center", indent_first=0, line=LINE_SINGLE, space_before=120, space_after=60, keep_next=True, keep_lines=True)}'
        f'{drawing_canvas(width_cm, height_cm, content, doc_id=doc_id)}'
        '</w:p>'
    )



# ---------------------------------------------------------------------------
# Семь рисунков главы 2 — конкретные DrawingML-генераторы
# ---------------------------------------------------------------------------

# Все рисунки рисуются в холсте шириной 16 см (≈ ширина зеркала набора).

FIG_W = 16.0  # cm


def _txt(text: str, *, sz_hp: int = 20, bold: bool = False,
         italic: bool = False, color: str = CLR_BLACK,
         align: str = "center") -> tuple[str, dict]:
    """Кратко: один параграф для текстового блока."""
    return (text, {"sz_hp": sz_hp, "bold": bold, "italic": italic,
                   "color": color, "align": align})


# ---------- Рис. 2.1.1. Структура группы банка «Ак Барс» ----------

def fig_2_1_1() -> tuple[float, str]:
    """Древовидная схема: 1 родитель + 5 дочерних блоков."""
    idg = DrawingIdGen(1100)
    parts = []

    # Размер холста.
    H = 6.0

    # Корневой блок.
    root_w, root_h = 7.0, 1.0
    root_x = (FIG_W - root_w) / 2
    root_y = 0.2
    parts.append(shape_rect(
        idg, root_x, root_y, root_w, root_h,
        fill=CLR_DARKBLUE, line=CLR_DARKBLUE,
        text_paragraphs=[_txt("Группа банка «Ак Барс»",
                              sz_hp=22, bold=True, color="FFFFFF")],
        name="Root",
    ))

    # Вертикальная линия от корня вниз к горизонтали.
    midx = FIG_W / 2
    line_y_top = root_y + root_h
    line_y_mid = 2.3
    parts.append(shape_line(
        idg, midx, line_y_top, midx, line_y_mid,
        color=CLR_BLACK, line_w=12700,
    ))

    # Дочерние блоки.
    n = 5
    pad = 0.1
    cw = (FIG_W - pad * (n - 1)) / n
    ch = 3.0
    cy = line_y_mid + 0.6
    children = [
        ('ПАО «АК БАРС»\nБАНК\n(головной)', CLR_BLUE),
        ('Инвестиционные\nкомпании', CLR_BLUE),
        ('Лизинговая\nкомпания', CLR_BLUE),
        ('Страховые\nкомпании', CLR_BLUE),
        ('Негосударственный\nпенсионный\nфонд', CLR_BLUE),
    ]
    # Горизонтальная линия по центру детей.
    first_cx = pad / 2 + cw / 2
    last_cx = FIG_W - pad / 2 - cw / 2
    parts.append(shape_line(
        idg, first_cx, line_y_mid, last_cx, line_y_mid,
        color=CLR_BLACK, line_w=12700,
    ))
    for i, (label, color) in enumerate(children):
        cx_box = (cw + pad) * i
        cx_center = cx_box + cw / 2
        # Вертикальная линия от горизонтали к ящику.
        parts.append(shape_line(
            idg, cx_center, line_y_mid, cx_center, cy,
            color=CLR_BLACK, line_w=12700,
        ))
        parts.append(shape_rect(
            idg, cx_box, cy, cw, ch,
            fill=color, line=CLR_DARKBLUE, line_w=9525,
            text_paragraphs=[_txt(label, sz_hp=20, bold=True, color="FFFFFF")],
            name=f"Child{i}",
        ))
    return H, "".join(parts)


# ---------- Рис. 2.1.2. Динамика чистых процентных доходов и прибыли ----------

def _bar_chart_pair(idg: DrawingIdGen, *, x0: float, y0: float,
                    width: float, height: float,
                    title: str, years: list[str], values: list[int],
                    color: str, max_val: int) -> str:
    """Один график: заголовок + 3 вертикальных столбца с подписями."""
    parts = []
    # Заголовок.
    parts.append(shape_textbox(
        idg, x0, y0, width, 0.8,
        text_paragraphs=[_txt(title, sz_hp=20, bold=True, color=CLR_DARKBLUE)],
        name="ChartTitle",
    ))
    # Ось.
    chart_top = y0 + 0.9
    chart_bot = y0 + height - 0.7
    chart_left = x0 + 0.5
    chart_right = x0 + width
    chart_h = chart_bot - chart_top
    # Вертикальная линия — ось Y.
    parts.append(shape_line(
        idg, chart_left, chart_top, chart_left, chart_bot,
        color=CLR_GRAY, line_w=6350,
    ))
    # Горизонтальная линия — ось X.
    parts.append(shape_line(
        idg, chart_left, chart_bot, chart_right, chart_bot,
        color=CLR_GRAY, line_w=6350,
    ))
    n = len(values)
    avail_w = chart_right - chart_left - 0.4
    bar_w = avail_w / (n * 1.6)
    gap_w = (avail_w - bar_w * n) / (n + 1) + bar_w * 0.0
    gap_w = (avail_w - bar_w * n) / (n + 1)
    for i, (yr, v) in enumerate(zip(years, values)):
        bx = chart_left + 0.2 + (i + 1) * gap_w + i * bar_w
        bh = chart_h * (v / max_val)
        by = chart_bot - bh
        parts.append(shape_rect(
            idg, bx, by, bar_w, bh,
            fill=color, line=color, line_w=0,
            name=f"Bar{i}",
        ))
        # Значение над столбцом.
        parts.append(shape_textbox(
            idg, bx - 0.3, by - 0.6, bar_w + 0.6, 0.6,
            text_paragraphs=[_txt(f"{v:,}".replace(",", " "),
                                  sz_hp=18, bold=True)],
            name="Val",
        ))
        # Год под осью.
        parts.append(shape_textbox(
            idg, bx - 0.3, chart_bot + 0.05, bar_w + 0.6, 0.5,
            text_paragraphs=[_txt(yr, sz_hp=18, bold=True)],
            name="Year",
        ))
    return "".join(parts)


def fig_2_1_2() -> tuple[float, str]:
    """Парные гистограммы: чистые процентные доходы и прибыль."""
    idg = DrawingIdGen(1200)
    H = 7.0
    half_w = (FIG_W - 0.5) / 2
    parts = []
    parts.append(_bar_chart_pair(
        idg, x0=0, y0=0.2, width=half_w, height=H - 0.2,
        title="Чистые процентные доходы, млн ₽",
        years=["2023", "2024", "2025"],
        values=[31824, 46590, 58135],
        color=CLR_BLUE,
        max_val=60000,
    ))
    parts.append(_bar_chart_pair(
        idg, x0=half_w + 0.5, y0=0.2, width=half_w, height=H - 0.2,
        title="Прибыль за период, млн ₽",
        years=["2023", "2024", "2025"],
        values=[14195, 7269, 15708],
        color=CLR_ORANGE,
        max_val=16000,
    ))
    return H, "".join(parts)


# ---------- Рис. 2.1.3. Архитектура взаимодействия Face2 с ГИС ЕБС ----------

def fig_2_1_3() -> tuple[float, str]:
    """Поток: Клиент → ГИС ЕБС / КБС Ак Барс Банк → 3 продукта."""
    idg = DrawingIdGen(1300)
    H = 9.5
    parts = []

    # Уровень 1 — Клиент (слева) и ГИС ЕБС (справа).
    box_w, box_h = 4.5, 1.6
    left_x = 0.5
    right_x = FIG_W - 0.5 - box_w
    top_y = 0.3
    parts.append(shape_rect(
        idg, left_x, top_y, box_w, box_h,
        fill=CLR_LIGHT_BLUE, line=CLR_DARKBLUE,
        text_paragraphs=[_txt("Клиент",
                              sz_hp=22, bold=True, color=CLR_DARKBLUE)],
        name="Client",
    ))
    parts.append(shape_rect(
        idg, right_x, top_y, box_w, box_h,
        fill=CLR_DARKBLUE, line=CLR_DARKBLUE,
        text_paragraphs=[_txt("ГИС ЕБС\n(Минцифры РФ)",
                              sz_hp=20, bold=True, color="FFFFFF")],
        name="EBS",
    ))
    # Стрелка с подписью между ними.
    arrow_y = top_y + box_h / 2
    parts.append(shape_line(
        idg, left_x + box_w + 0.1, arrow_y,
        right_x - 0.1, arrow_y,
        color=CLR_DARKBLUE, line_w=12700, arrow_end=True,
    ))
    parts.append(shape_textbox(
        idg, left_x + box_w + 0.1, arrow_y - 0.6,
        right_x - left_x - box_w - 0.2, 0.5,
        text_paragraphs=[_txt("упрощённая биометрия",
                              sz_hp=16, italic=True, color=CLR_GRAY)],
        name="ArrowLabel",
    ))

    # Уровень 2 — КБС Ак Барс Банк (Face2). По центру.
    kbs_w, kbs_h = 6.0, 1.8
    kbs_x = (FIG_W - kbs_w) / 2
    kbs_y = top_y + box_h + 1.4
    parts.append(shape_rect(
        idg, kbs_x, kbs_y, kbs_w, kbs_h,
        fill=CLR_BLUE, line=CLR_DARKBLUE, line_w=12700,
        text_paragraphs=[_txt("КБС Ак Барс Банк (Face2)",
                              sz_hp=24, bold=True, color="FFFFFF")],
        name="KBS",
    ))

    # Стрелки от Клиента к КБС («согласие на КБС АББ») и от ГИС ЕБС к КБС («вектор»).
    parts.append(shape_line(
        idg, left_x + box_w / 2, top_y + box_h,
        kbs_x + 0.7, kbs_y,
        color=CLR_DARKBLUE, line_w=12700, arrow_end=True,
    ))
    parts.append(shape_textbox(
        idg, left_x + 0.5, (top_y + box_h + kbs_y) / 2 - 0.25,
        2.5, 0.5,
        text_paragraphs=[_txt("согласие на КБС АББ",
                              sz_hp=16, italic=True, color=CLR_GRAY)],
        name="L1",
    ))
    parts.append(shape_line(
        idg, right_x + box_w / 2, top_y + box_h,
        kbs_x + kbs_w - 0.7, kbs_y,
        color=CLR_DARKBLUE, line_w=12700, arrow_end=True,
    ))
    parts.append(shape_textbox(
        idg, right_x - 0.5, (top_y + box_h + kbs_y) / 2 - 0.25,
        2.5, 0.5,
        text_paragraphs=[_txt("вектор",
                              sz_hp=16, italic=True, color=CLR_GRAY)],
        name="L2",
    ))

    # Уровень 3 — три продукта.
    prods = [
        ("Face2Pay",     "оплата по лицу"),
        ("Face2Pass",    "контроль доступа"),
        ("Face2Check-in","подтверждение возраста"),
    ]
    pw, ph = 4.5, 1.8
    py = kbs_y + kbs_h + 1.4
    pad = (FIG_W - len(prods) * pw) / (len(prods) + 1)
    for i, (name, desc) in enumerate(prods):
        px = pad + i * (pw + pad)
        # Стрелка от КБС к продукту.
        parts.append(shape_line(
            idg, kbs_x + kbs_w / 2, kbs_y + kbs_h,
            px + pw / 2, py,
            color=CLR_DARKBLUE, line_w=12700, arrow_end=True,
        ))
        parts.append(shape_rect(
            idg, px, py, pw, ph,
            fill=CLR_LIGHT_BLUE, line=CLR_DARKBLUE,
            text_paragraphs=[
                _txt(name, sz_hp=22, bold=True, color=CLR_DARKBLUE),
                _txt(desc, sz_hp=18, color=CLR_BLACK),
            ],
            name=f"Prod{i}",
        ))
    return H, "".join(parts)


# ---------- Рис. 2.2.1. Поквартальная динамика биометрических платежей ----------

def fig_2_2_1() -> tuple[float, str]:
    idg = DrawingIdGen(1400)
    H = 7.0
    half_w = (FIG_W - 0.5) / 2
    parts = []
    parts.append(_bar_chart_pair(
        idg, x0=0, y0=0.2, width=half_w, height=H - 0.2,
        title="Количество платежей, млн",
        years=["II'24", "I'25", "II'25"],
        values=[5, 40, 60],
        color=CLR_BLUE,
        max_val=65,
    ))
    parts.append(_bar_chart_pair(
        idg, x0=half_w + 0.5, y0=0.2, width=half_w, height=H - 0.2,
        title="Оборот, млрд ₽",
        years=["II'24", "I'25", "II'25"],
        values=[4, 30, 45],
        color=CLR_GREEN,
        max_val=50,
    ))
    return H, "".join(parts)


# ---------- Рис. 2.2.2. Структура конкурентной среды ----------

def fig_2_2_2() -> tuple[float, str]:
    idg = DrawingIdGen(1500)
    H = 9.0
    parts = []

    # Корневой блок.
    root_w, root_h = 9.0, 1.0
    root_x = (FIG_W - root_w) / 2
    root_y = 0.2
    parts.append(shape_rect(
        idg, root_x, root_y, root_w, root_h,
        fill=CLR_DARKBLUE, line=CLR_DARKBLUE,
        text_paragraphs=[_txt("Российский рынок биометрии (B2B/B2C)",
                              sz_hp=22, bold=True, color="FFFFFF")],
        name="Root",
    ))

    # 4 блока второго уровня.
    items = [
        ("Слой 1.\nБанки-разработчики\nсобственных биосервисов",
         "ПАО Сбербанк\nПАО «АК БАРС» БАНК\n(Face2)",
         CLR_BLUE),
        ("Слой 2.\nТехнологические\nвендоры алгоритмов\nи платформ",
         "VisionLabs\nNtechLab\nTevian",
         CLR_BLUE),
        ("Слой 3.\nПроизводители\nполного цикла СКУД\n(ПО + оборудование)",
         "OVISION\nBioSmart\nи др.",
         CLR_BLUE),
        ("Государственная\nинфраструктура\n(АО «ЦБТ» —\nоператор ЕБС)",
         "ЕБС / ЕСИА\n(574-ФЗ, 572-ФЗ)",
         CLR_GRAY),
    ]
    n = len(items)
    pad = 0.2
    bw = (FIG_W - pad * (n - 1)) / n
    bh = 4.5
    by = root_y + root_h + 1.0
    # Соединительная горизонталь.
    midx = FIG_W / 2
    parts.append(shape_line(
        idg, midx, root_y + root_h, midx, by - 0.5,
        color=CLR_BLACK, line_w=12700,
    ))
    parts.append(shape_line(
        idg, pad / 2 + bw / 2, by - 0.5,
        FIG_W - pad / 2 - bw / 2, by - 0.5,
        color=CLR_BLACK, line_w=12700,
    ))
    for i, (head, body, color) in enumerate(items):
        bx = (bw + pad) * i
        cx = bx + bw / 2
        parts.append(shape_line(
            idg, cx, by - 0.5, cx, by,
            color=CLR_BLACK, line_w=12700,
        ))
        # Шапка слоя.
        head_h = 2.4
        parts.append(shape_rect(
            idg, bx, by, bw, head_h,
            fill=color, line=CLR_DARKBLUE,
            text_paragraphs=[_txt(head, sz_hp=18, bold=True,
                                  color="FFFFFF")],
            name=f"Layer{i}H",
        ))
        # Игроки.
        body_y = by + head_h
        body_h = bh - head_h
        parts.append(shape_rect(
            idg, bx, body_y, bw, body_h,
            fill="FFFFFF", line=CLR_DARKBLUE,
            text_paragraphs=[_txt(body, sz_hp=18, color=CLR_BLACK)],
            name=f"Layer{i}B",
        ))
    return H, "".join(parts)


# ---------- Рис. 2.2.3. Модель пяти конкурентных сил Портера ----------

def fig_2_2_3() -> tuple[float, str]:
    idg = DrawingIdGen(1600)
    H = 11.5
    parts = []

    # Центральный блок — Внутриотраслевая конкуренция.
    cw, ch = 5.5, 2.4
    cx = (FIG_W - cw) / 2
    cy = (H - ch) / 2 - 0.2
    parts.append(shape_rect(
        idg, cx, cy, cw, ch,
        fill=CLR_DARKBLUE, line=CLR_DARKBLUE,
        text_paragraphs=[
            _txt("Внутриотраслевая\nконкуренция",
                 sz_hp=20, bold=True, color="FFFFFF"),
            _txt("[СРЕДНЕ-ВЫСОКАЯ]",
                 sz_hp=16, bold=True, color="FFD966"),
        ],
        name="Center",
    ))

    # Четыре окружающих блока.
    fw, fh = 5.0, 2.6
    # Top
    top_x = (FIG_W - fw) / 2
    top_y = 0.3
    parts.append(shape_rect(
        idg, top_x, top_y, fw, fh,
        fill=CLR_BLUE, line=CLR_DARKBLUE,
        text_paragraphs=[
            _txt("Угроза новых\nучастников",
                 sz_hp=18, bold=True, color="FFFFFF"),
            _txt("[НИЗКАЯ]",
                 sz_hp=14, bold=True, color="FFD966"),
            _txt("572-ФЗ, аккредитация\nМинцифры (КБС);\nстатус оператора\nпо № 115-ФЗ; ФСБ/ФСТЭК;\n420-ФЗ и 421-ФЗ — штрафы",
                 sz_hp=12, color="FFFFFF"),
        ],
        name="ThreatNew",
    ))
    parts.append(shape_line(
        idg, FIG_W / 2, top_y + fh, FIG_W / 2, cy,
        color=CLR_DARKBLUE, line_w=15875, arrow_end=True,
    ))

    # Bottom
    bot_x = (FIG_W - fw) / 2
    bot_y = H - fh - 0.3
    parts.append(shape_rect(
        idg, bot_x, bot_y, fw, fh,
        fill=CLR_BLUE, line=CLR_DARKBLUE,
        text_paragraphs=[
            _txt("Угроза товаров-\nзаменителей",
                 sz_hp=18, bold=True, color="FFFFFF"),
            _txt("[СРЕДНЯЯ]",
                 sz_hp=14, bold=True, color="FFD966"),
            _txt("QR-эквайринг (×16),\nNFC, SMS-OTP,\nголосовая биометрия,\nкарты и мобильные ключи\nв СКУД",
                 sz_hp=12, color="FFFFFF"),
        ],
        name="ThreatSubst",
    ))
    parts.append(shape_line(
        idg, FIG_W / 2, bot_y, FIG_W / 2, cy + ch,
        color=CLR_DARKBLUE, line_w=15875, arrow_end=True,
    ))

    # Left
    left_x = 0.1
    left_y = (H - fh) / 2
    parts.append(shape_rect(
        idg, left_x, left_y, fw, fh,
        fill=CLR_BLUE, line=CLR_DARKBLUE,
        text_paragraphs=[
            _txt("Рыночная сила\nпоставщиков",
                 sz_hp=18, bold=True, color="FFFFFF"),
            _txt("[СРЕДНЯЯ]",
                 sz_hp=14, bold=True, color="FFD966"),
            _txt("Производители\nбиотерминалов;\nАО ЦБТ — монополист\nхранения векторов;\nсобственные алгоритмы\nFace2 (99,5%)",
                 sz_hp=12, color="FFFFFF"),
        ],
        name="Suppliers",
    ))
    parts.append(shape_line(
        idg, left_x + fw, left_y + fh / 2, cx, cy + ch / 2,
        color=CLR_DARKBLUE, line_w=15875, arrow_end=True,
    ))

    # Right
    right_x = FIG_W - fw - 0.1
    right_y = (H - fh) / 2
    parts.append(shape_rect(
        idg, right_x, right_y, fw, fh,
        fill=CLR_BLUE, line=CLR_DARKBLUE,
        text_paragraphs=[
            _txt("Рыночная сила\nпокупателей",
                 sz_hp=18, bold=True, color="FFFFFF"),
            _txt("[СРЕДНЕ-ВЫСОКАЯ]",
                 sz_hp=14, bold=True, color="FFD966"),
            _txt("Крупные B2B/B2G:\nбанки, транспорт,\nритейл, госучреждения;\nнизкие издержки\nпереключения\nв техслое",
                 sz_hp=12, color="FFFFFF"),
        ],
        name="Buyers",
    ))
    parts.append(shape_line(
        idg, right_x, right_y + fh / 2, cx + cw, cy + ch / 2,
        color=CLR_DARKBLUE, line_w=15875, arrow_end=True,
    ))
    return H, "".join(parts)


# ---------- Рис. 2.3.1. Воронка Face2 vs эталон ----------

def fig_2_3_1() -> tuple[float, str]:
    idg = DrawingIdGen(1700)
    H = 9.0
    parts = []

    # Заголовки колонок.
    col1_x = 0.0
    col2_x = 4.5
    col3_x = 10.2
    col_h = 0.7
    parts.append(shape_textbox(
        idg, col1_x, 0, 4.0, col_h,
        text_paragraphs=[_txt("Этап воронки", sz_hp=18, bold=True,
                              color=CLR_DARKBLUE)],
        name="Col1Hdr",
    ))
    parts.append(shape_textbox(
        idg, col2_x, 0, 5.5, col_h,
        text_paragraphs=[_txt("Эталонная воронка", sz_hp=18, bold=True,
                              color=CLR_DARKBLUE, align="left")],
        name="Col2Hdr",
    ))
    parts.append(shape_textbox(
        idg, col3_x, 0, 5.5, col_h,
        text_paragraphs=[_txt("Воронка Face2 (факт)", sz_hp=18, bold=True,
                              color=CLR_DARKBLUE, align="left")],
        name="Col3Hdr",
    ))
    # Линия под шапкой.
    parts.append(shape_line(
        idg, 0, col_h + 0.05, FIG_W, col_h + 0.05,
        color=CLR_GRAY, line_w=6350,
    ))

    # Этапы.
    stages = [
        # (name, etalon_w, face2_main_w, face2_extra_label, face2_extra_w)
        ("Осведомлённость", 0.55, 0.15, None, 0),
        ("Интерес",         0.45, 0.10, None, 0),
        ("Доверие (ядро)",  0.85, 0.85, "массовый охват", 0.20),
        ("Пробное использование", 0.55, 0.55, "вне РТ", 0.05),
        ("Принятие",        0.45, 0.30, None, 0),
        ("Лояльность",      0.30, 0.05, None, 0),
    ]
    bar_max_cm = 5.3   # макс. длина бара
    bar_h = 0.55
    row_y = col_h + 0.4
    row_step = 1.2
    for i, (st, e, f1, extra_label, f2) in enumerate(stages):
        y = row_y + i * row_step
        # Название этапа.
        parts.append(shape_textbox(
            idg, col1_x, y, 4.0, bar_h,
            text_paragraphs=[_txt(st, sz_hp=18, bold=True, align="left")],
            name=f"Stage{i}",
        ))
        # Эталон.
        parts.append(shape_rect(
            idg, col2_x, y, bar_max_cm * e, bar_h,
            fill=CLR_GREEN, line=CLR_GREEN, line_w=0,
            name=f"Et{i}",
        ))
        # Face2 основной.
        parts.append(shape_rect(
            idg, col3_x, y, bar_max_cm * f1, bar_h,
            fill=CLR_BLUE, line=CLR_BLUE, line_w=0,
            name=f"F2m{i}",
        ))
        if extra_label and f2 > 0:
            # Доп. бар чуть ниже основного.
            extra_y = y + bar_h + 0.05
            parts.append(shape_rect(
                idg, col3_x, extra_y, bar_max_cm * f2, bar_h * 0.7,
                fill=CLR_ORANGE, line=CLR_ORANGE, line_w=0,
                name=f"F2e{i}",
            ))
            parts.append(shape_textbox(
                idg, col3_x + bar_max_cm * f2 + 0.1, extra_y - 0.05,
                3.0, bar_h * 0.7,
                text_paragraphs=[_txt(f"({extra_label})", sz_hp=14,
                                      italic=True, color=CLR_ORANGE,
                                      align="left")],
                name=f"F2elbl{i}",
            ))
        if extra_label and i == 2:
            # Для строки «Доверие» — пометка для основного бара.
            parts.append(shape_textbox(
                idg, col3_x + bar_max_cm * f1 + 0.1, y - 0.05,
                3.0, bar_h,
                text_paragraphs=[_txt("(содержательно)", sz_hp=14,
                                      italic=True, color=CLR_BLUE,
                                      align="left")],
                name=f"F2mlbl{i}",
            ))
    # Высота холста — по последней строке.
    last_y = row_y + (len(stages) - 1) * row_step + bar_h * 2 + 0.2
    return last_y, "".join(parts)


# Реестр генераторов рисунков по их номеру.
FIGURE_BUILDERS = {
    "2.1.1": fig_2_1_1,
    "2.1.2": fig_2_1_2,
    "2.1.3": fig_2_1_3,
    "2.2.1": fig_2_2_1,
    "2.2.2": fig_2_2_2,
    "2.2.3": fig_2_2_3,
    "2.3.1": fig_2_3_1,
}



# ---------------------------------------------------------------------------
# Парсер главы 2 из markdown
# ---------------------------------------------------------------------------

# Блок результата парсинга — список dict'ов.
# Типы: "chapter_h", "para_h", "subsection_h", "para", "table",
#       "figure" (со ссылкой на номер из FIGURE_BUILDERS), "fig_caption",
#       "fig_source", "tbl_caption" (number, title), "tbl_source",
#       "page_break".


def _is_table_separator(line: str) -> bool:
    """Линия-разделитель шапки markdown-таблицы вида |---|---|---|."""
    s = line.strip()
    return bool(re.match(r"^\|[\s\-\|:]+\|$", s)) and "-" in s


def _split_table_row(line: str) -> list[str]:
    """Разбить строку markdown-таблицы по '|', обрезая внешние."""
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def parse_chapter2(md_path: Path) -> list[dict]:
    """Прочитать markdown и вернуть последовательность блоков главы 2.

    Глава 2 = от строки `# 2.` до конца файла."""
    text = md_path.read_text(encoding="utf-8")
    lines = text.split("\n")

    # Найти диапазон главы 2.
    start = None
    for i, ln in enumerate(lines):
        if ln.startswith("# 2. "):
            start = i
            break
    if start is None:
        raise RuntimeError("Глава 2 не найдена")
    chapter_lines = lines[start:]

    blocks: list[dict] = []
    i = 0
    n = len(chapter_lines)

    def flush_paragraph(buf: list[str]) -> None:
        if not buf:
            return
        text = " ".join(s.strip() for s in buf if s.strip())
        if not text:
            return
        # Распознать спец-абзацы перед таблицей/рисунком.
        # 1) "Таблица N.N.N" — самостоятельный блок.
        m = re.match(r"^Таблица\s+(\d+\.\d+\.\d+)\s*$", text)
        if m:
            blocks.append({"type": "tbl_number", "value": m.group(1)})
            return
        # 2) "Рис. N.N.N. ..." — подпись рисунка.
        m = re.match(r"^Рис\.\s+(\d+\.\d+\.\d+)\.\s+(.+?)\s*$", text)
        if m:
            blocks.append({
                "type": "fig_caption",
                "number": m.group(1),
                "title": m.group(2),
            })
            return
        # 3) "Источник: ..." — подпись источника.
        if text.startswith("Источник:"):
            blocks.append({"type": "source", "text": text})
            return
        # Обычный абзац.
        blocks.append({"type": "para", "text": text})

    para_buf: list[str] = []
    while i < n:
        ln = chapter_lines[i]
        stripped = ln.strip()
        # Заголовки.
        if ln.startswith("# 2. "):
            flush_paragraph(para_buf); para_buf = []
            blocks.append({"type": "chapter_h", "text": ln[2:].strip()})
            i += 1
            continue
        if ln.startswith("# 2.") and re.match(r"^# 2\.\d+\.\s", ln):
            flush_paragraph(para_buf); para_buf = []
            blocks.append({
                "type": "para_h",
                "text": ln[2:].strip(),
            })
            i += 1
            continue
        if ln.startswith("# Вывод по главе"):
            flush_paragraph(para_buf); para_buf = []
            blocks.append({"type": "chapter_h", "text": ln[2:].strip()})
            i += 1
            continue
        if ln.startswith("## "):
            flush_paragraph(para_buf); para_buf = []
            blocks.append({"type": "subsection_h", "text": ln[3:].strip()})
            i += 1
            continue
        # Горизонтальная линия --- → разрыв страницы перед следующей главой.
        if stripped == "---":
            flush_paragraph(para_buf); para_buf = []
            blocks.append({"type": "hr"})
            i += 1
            continue
        # ASCII-блок ```...``` → пропускаем (заменим рисунком).
        if stripped == "```":
            flush_paragraph(para_buf); para_buf = []
            # Найти закрывающий ```.
            j = i + 1
            while j < n and chapter_lines[j].strip() != "```":
                j += 1
            blocks.append({"type": "ascii_block_marker"})
            i = j + 1 if j < n else j
            continue
        # Markdown-таблица: строка, начинающаяся с '|', и следующая — разделитель.
        if stripped.startswith("|") and i + 1 < n and \
                _is_table_separator(chapter_lines[i + 1]):
            flush_paragraph(para_buf); para_buf = []
            header = _split_table_row(chapter_lines[i])
            i += 2
            data_rows = []
            while i < n and chapter_lines[i].strip().startswith("|"):
                data_rows.append(_split_table_row(chapter_lines[i]))
                i += 1
            blocks.append({
                "type": "table",
                "header": header,
                "rows": data_rows,
            })
            continue
        # Пустая строка — конец абзаца.
        if not stripped:
            flush_paragraph(para_buf); para_buf = []
            i += 1
            continue
        # Обычная текстовая строка.
        para_buf.append(ln)
        i += 1
    flush_paragraph(para_buf)

    # Постобработка: подцепить figure_builder к подписи и склеить
    # (source above caption) → figure block.
    out: list[dict] = []
    j = 0
    while j < len(blocks):
        b = blocks[j]
        if b["type"] == "ascii_block_marker":
            # Перед ascii_block был параграф «...показана на рисунке N.N.N».
            # После него идут: "Источник: ..." (source) и "Рис. N.N.N. Title"
            # (fig_caption). Соберём всё в один блок figure.
            # Найти source и fig_caption в следующих 2-3 блоках.
            src = None
            cap = None
            k = j + 1
            consumed = []
            while k < len(blocks) and k - j <= 3:
                if blocks[k]["type"] == "source" and src is None:
                    src = blocks[k]["text"]
                    consumed.append(k)
                elif blocks[k]["type"] == "fig_caption" and cap is None:
                    cap = blocks[k]
                    consumed.append(k)
                else:
                    break
                k += 1
            if cap:
                out.append({
                    "type": "figure",
                    "number": cap["number"],
                    "title": cap["title"],
                    "source": src,
                })
                # Пропустить consumed-блоки.
                j = max(consumed) + 1
                continue
            j += 1
            continue
        if b["type"] == "tbl_number":
            # Шапка таблицы: «Таблица N» + «Название» (следующий para)
            # + «table» (markdown-таблица) + «Источник: ...» (source).
            number = b["value"]
            title = None
            tbl_block = None
            src = None
            k = j + 1
            consumed = []
            # title — следующий "para".
            if k < len(blocks) and blocks[k]["type"] == "para":
                title = blocks[k]["text"]
                consumed.append(k)
                k += 1
            # table.
            if k < len(blocks) and blocks[k]["type"] == "table":
                tbl_block = blocks[k]
                consumed.append(k)
                k += 1
            # source.
            if k < len(blocks) and blocks[k]["type"] == "source":
                src = blocks[k]["text"]
                consumed.append(k)
                k += 1
            out.append({
                "type": "table_full",
                "number": number,
                "title": title or "",
                "header": tbl_block["header"] if tbl_block else [],
                "rows": tbl_block["rows"] if tbl_block else [],
                "source": src or "",
            })
            j = max(consumed) + 1 if consumed else j + 1
            continue
        out.append(b)
        j += 1
    return out


# ---------------------------------------------------------------------------
# Сборка тела документа
# ---------------------------------------------------------------------------

def render_blocks(blocks: list[dict]) -> str:
    """Собрать XML-тело документа из последовательности блоков."""
    parts = []
    fig_doc_id = 1000
    for b in blocks:
        t = b["type"]
        if t == "chapter_h":
            parts.append(chapter_heading(b["text"]))
        elif t == "para_h":
            parts.append(paragraph_heading(b["text"]))
        elif t == "subsection_h":
            parts.append(subsection_heading(b["text"]))
        elif t == "para":
            parts.append(body_paragraph(b["text"]))
        elif t == "source":
            # Если идёт сразу после таблицы — это уже учтено в table_full.
            # Если после рисунка — тоже. На уровне render это редкий случай.
            parts.append(figure_source_block(b["text"]))
        elif t == "fig_caption":
            parts.append(figure_caption_block(
                f"Рис. {b['number']}. {b['title']}"
            ))
        elif t == "table_full":
            number = b["number"]
            title = b["title"]
            parts.append(table_caption_block(number, title))
            # Подобрать ширину колонок: равномерно по ширине набора.
            n_cols = max(
                len(b["header"]) if b["header"] else 0,
                max((len(r) for r in b["rows"]), default=0),
            )
            if n_cols == 0:
                continue
            col_widths_cm = _adaptive_widths(
                b["header"], b["rows"], n_cols, total_cm=CONTENT_W_CM,
            )
            rows = []
            if b["header"]:
                rows.append(b["header"])
            rows.extend(b["rows"])
            parts.append(table(
                rows, col_widths_cm=col_widths_cm,
                header_rows=1 if b["header"] else 0,
            ))
            if b["source"]:
                parts.append(table_source_block(b["source"]))
            else:
                parts.append(empty_para(LINE_SINGLE))
        elif t == "figure":
            number = b["number"]
            if b["source"]:
                parts.append(figure_source_block(b["source"]))
            builder = FIGURE_BUILDERS.get(number)
            if builder:
                fig_h_cm, fig_xml = builder()
                fig_doc_id += 1
                parts.append(figure_drawing_para(
                    FIG_W, fig_h_cm, fig_xml, doc_id=fig_doc_id,
                ))
            else:
                parts.append(body_paragraph(
                    f"[Рисунок {number} — не реализован в билдере]"
                ))
            parts.append(figure_caption_block(
                f"Рис. {number}. {b['title']}"
            ))
        elif t == "hr":
            # Между параграфами — пропустить (разрыв уже даётся через
            # spacing). Между разделами «Вывод» — небольшой отступ.
            parts.append(empty_para(LINE_SINGLE))
        else:
            # неизвестный блок — пропустим.
            pass
    return "".join(parts)


def _adaptive_widths(header: list[str], rows: list[list[str]],
                     n_cols: int, total_cm: float) -> list[float]:
    """Прикидочная ширина колонок по содержимому (в символах).

    Используем длину самого длинного слова × коэффициент, минимум 1,5 см.
    """
    # Максимальная длина текста в каждой колонке.
    max_len = [0] * n_cols
    sources = [header] + rows
    for r in sources:
        for ci in range(min(len(r), n_cols)):
            cell = r[ci] or ""
            # Учитываем максимальное слово, чтобы избежать слишком узких колонок.
            words = cell.replace("\n", " ").split()
            longest_word = max((len(w) for w in words), default=0)
            cell_total = len(cell)
            # Эффективная "ширина" — между longest_word и cell_total/2.
            effective = max(longest_word, min(cell_total // 2, 30))
            max_len[ci] = max(max_len[ci], effective)
    if sum(max_len) == 0:
        return [total_cm / n_cols] * n_cols
    # Минимальная ширина 1,2 см.
    min_cm = 1.2
    raw = [max(ml, 1) for ml in max_len]
    total_raw = sum(raw)
    widths = [max(min_cm, total_cm * r / total_raw) for r in raw]
    # Нормализуем сумму ровно к total_cm.
    s = sum(widths)
    widths = [w * total_cm / s for w in widths]
    return widths


# ---------------------------------------------------------------------------
# Сборка ZIP-контейнера .docx
# ---------------------------------------------------------------------------

CONTENT_TYPES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

ROOT_RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

DOC_RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''

CORE_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Глава 2 ВКР — оформление</dc:title>
  <dc:creator>Автор</dc:creator>
  <cp:lastModifiedBy>Автор</cp:lastModifiedBy>
</cp:coreProperties>'''

APP_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>VKR-builder</Application>
</Properties>'''

SETTINGS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="708"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
  <w:themeFontLang w:val="ru-RU"/>
</w:settings>'''

FONT_TABLE_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Times New Roman">
    <w:panose1 w:val="02020603050405020304"/>
    <w:charset w:val="CC"/>
    <w:family w:val="roman"/>
    <w:pitch w:val="variable"/>
  </w:font>
</w:fonts>'''

NUMBERING_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'''

STYLES_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="{FONT_MAIN}" w:hAnsi="{FONT_MAIN}" w:cs="{FONT_MAIN}" w:eastAsia="{FONT_MAIN}"/>
        <w:sz w:val="{SZ_BODY_HP}"/>
        <w:szCs w:val="{SZ_BODY_HP}"/>
        <w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:before="0" w:after="0" w:line="{LINE_BODY}" w:lineRule="auto"/>
        <w:jc w:val="both"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>'''


def build_document_xml(body_xml: str) -> str:
    """Сборка word/document.xml с правильными неймспейсами и page setup."""
    sect_pr = (
        '<w:sectPr>'
        f'<w:pgSz w:w="{PAGE_W_TW}" w:h="{PAGE_H_TW}"/>'
        f'<w:pgMar w:top="{MARGIN_TOP_TW}" w:right="{MARGIN_RIGHT_TW}" '
        f'w:bottom="{MARGIN_BOTTOM_TW}" w:left="{MARGIN_LEFT_TW}" '
        f'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="708"/>'
        '<w:docGrid w:linePitch="360"/>'
        '</w:sectPr>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<w:body>'
        f'{body_xml}'
        f'{sect_pr}'
        '</w:body>'
        '</w:document>'
    )


def write_docx(out_path: Path, body_xml: str) -> None:
    """Записать .docx как ZIP с минимальным набором деталей."""
    doc_xml = build_document_xml(body_xml)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        zf.writestr("_rels/.rels", ROOT_RELS_XML)
        zf.writestr("word/_rels/document.xml.rels", DOC_RELS_XML)
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("word/styles.xml", STYLES_XML)
        zf.writestr("word/settings.xml", SETTINGS_XML)
        zf.writestr("word/fontTable.xml", FONT_TABLE_XML)
        zf.writestr("word/numbering.xml", NUMBERING_XML)
        zf.writestr("docProps/core.xml", CORE_XML)
        zf.writestr("docProps/app.xml", APP_XML)


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main() -> None:
    src = Path(__file__).parent / "ВКР ГОТОВОЕ.md"
    out = Path(__file__).parent / "Глава_2_оформленная.docx"
    blocks = parse_chapter2(src)
    body_xml = render_blocks(blocks)
    write_docx(out, body_xml)
    print(f"Записан файл: {out} ({out.stat().st_size:,} байт)")
    # Краткая статистика блоков.
    by_type: dict[str, int] = {}
    for b in blocks:
        by_type[b["type"]] = by_type.get(b["type"], 0) + 1
    print("Блоки:")
    for k in sorted(by_type):
        print(f"  {k}: {by_type[k]}")


if __name__ == "__main__":
    main()
