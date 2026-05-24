#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор презентации ВКР «Разработка стратегии продвижения инновационного
продукта Face2 Ак Барс Банка» в фирменном стиле кафедры инноваций
и инвестиций ИУЭиФ КФУ (см. Стиль_презентации_ВКР_ИиИ_2026.md).

Скрипт берёт за основу оригинальный шаблон pptx, сохраняет все ресурсы
(theme, master, layouts, media, rels) и заменяет XML-тела всех 10 слайдов
на собственные с фактическим содержанием из ВКР и инфографикой.
"""
from __future__ import annotations
import io
import shutil
import zipfile
from html import escape
from pathlib import Path

# ---------------------------------------------------------------------------
# Константы стиля (из Стиль_презентации_ВКР_ИиИ_2026.md)
# ---------------------------------------------------------------------------
EMU_CM = 360000  # 1 см = 360 000 EMU

SLIDE_W = 9144000
SLIDE_H = 5143500

BLUE = "17365D"
WHITE = "FFFFFF"
BLACK = "000000"
GRAY = "888888"
LIGHT_GRAY = "F2F2F2"
ACCENT_LIGHT = "D6DCE5"   # светло-синий для зебры таблиц

FONT_MAIN = "PT Sans"
FONT_FALLBACK = "Times New Roman"

# Геометрия макета контентного слайда
STRIPE_W = 899592
TITLE_X = 989551
TITLE_Y = 130000
TITLE_W = 7800000
TITLE_H = 580000

CONTENT_X = 1050000
CONTENT_Y = 850000
CONTENT_W = 7900000
CONTENT_H = 4000000


# ---------------------------------------------------------------------------
# Низкоуровневые помощники для генерации OOXML
# ---------------------------------------------------------------------------

class IdGen:
    def __init__(self, start: int = 100):
        self._n = start

    def next(self) -> int:
        self._n += 1
        return self._n


def _t(text: str) -> str:
    """Эскейп текста для OOXML."""
    return (text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;"))


def run(text: str, *, sz: int = 1600, bold: bool = False, italic: bool = False,
        color: str = BLACK, font: str = FONT_MAIN) -> str:
    """Inline run."""
    b = ' b="1"' if bold else ""
    i = ' i="1"' if italic else ""
    return (
        f'<a:r><a:rPr lang="ru-RU" sz="{sz}"{b}{i} dirty="0">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{font}"/><a:ea typeface="{font}"/><a:cs typeface="{font}"/>'
        f'</a:rPr><a:t>{_t(text)}</a:t></a:r>'
    )


def para(*runs: str, align: str = "l", level: int = 0,
         space_before: int = 0, line_spacing: int = 110000) -> str:
    """Paragraph wrapper."""
    runs_xml = "".join(runs) if runs else (
        '<a:endParaRPr lang="ru-RU" dirty="0"/>'
    )
    return (
        f'<a:p><a:pPr marL="0" lvl="{level}" indent="0" algn="{align}" rtl="0">'
        f'<a:lnSpc><a:spcPct val="{line_spacing}"/></a:lnSpc>'
        f'<a:spcBef><a:spcPts val="{space_before}"/></a:spcBef>'
        f'<a:spcAft><a:spcPts val="0"/></a:spcAft>'
        f'<a:buNone/></a:pPr>{runs_xml}</a:p>'
    )


def bullet_para(*runs: str, level: int = 0, color: str = BLUE,
                space_before: int = 200) -> str:
    """Параграф с маркером (квадрат) фирменного цвета."""
    runs_xml = "".join(runs)
    indent = -200000
    margin = 250000 + level * 200000
    return (
        f'<a:p><a:pPr marL="{margin}" lvl="{level}" indent="{indent}" algn="l" rtl="0">'
        f'<a:lnSpc><a:spcPct val="115000"/></a:lnSpc>'
        f'<a:spcBef><a:spcPts val="{space_before}"/></a:spcBef>'
        f'<a:spcAft><a:spcPts val="0"/></a:spcAft>'
        f'<a:buClr><a:srgbClr val="{color}"/></a:buClr>'
        f'<a:buFont typeface="Arial"/>'
        f'<a:buChar char="\u25A0"/>'
        f'</a:pPr>{runs_xml}</a:p>'
    )


def text_box(idg: IdGen, x: int, y: int, w: int, h: int, *paragraphs: str,
             name: str = "TextBox", anchor: str = "t") -> str:
    """Прозрачный текстовый блок."""
    body = "".join(paragraphs) if paragraphs else para()
    return f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{idg.next()}" name="{name}"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"
              anchor="{anchor}" anchorCtr="0"><a:noAutofit/></a:bodyPr>
    <a:lstStyle/>
    {body}
  </p:txBody>
</p:sp>'''


def filled_box(idg: IdGen, x: int, y: int, w: int, h: int, *paragraphs: str,
               fill: str = BLUE, line_color: str | None = None,
               line_w: int = 0, name: str = "Rect", anchor: str = "ctr",
               geom: str = "rect") -> str:
    """Заливной прямоугольник (плитка) с текстом."""
    body = "".join(paragraphs) if paragraphs else para()
    line_xml = (
        f'<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{line_color}"/></a:solidFill></a:ln>'
        if line_color else '<a:ln><a:noFill/></a:ln>'
    )
    return f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{idg.next()}" name="{name}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
    <a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>
    {line_xml}
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"
              anchor="{anchor}" anchorCtr="1"><a:normAutofit/></a:bodyPr>
    <a:lstStyle/>
    {body}
  </p:txBody>
</p:sp>'''


def line_shape(idg: IdGen, x: int, y: int, w: int, h: int,
               color: str = BLUE, weight: int = 19050,
               flip_h: bool = False) -> str:
    """Прямая линия (используется как разделитель)."""
    flip = ' flipH="1"' if flip_h else ""
    return f'''<p:cxnSp>
  <p:nvCxnSpPr>
    <p:cNvPr id="{idg.next()}" name="Line"/>
    <p:cNvCxnSpPr/>
    <p:nvPr/>
  </p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm{flip}><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
    <a:ln w="{weight}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln>
  </p:spPr>
</p:cxnSp>'''


def pic_element(idg: IdGen, rid: str, x: int, y: int, w: int, h: int,
                name: str = "Picture") -> str:
    """Изображение по relationship id (для синей полосы и логотипа)."""
    return f'''<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="{idg.next()}" name="{name}"/>
    <p:cNvPicPr preferRelativeResize="0"/>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill rotWithShape="1">
    <a:blip r:embed="{rid}"><a:alphaModFix/></a:blip>
    <a:srcRect/><a:stretch/>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
</p:pic>'''


# ---------------------------------------------------------------------------
# Сборка целого слайда
# ---------------------------------------------------------------------------

SLIDE_HEAD = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
        <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
'''

SLIDE_TAIL = '''
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''


def title_block(idg: IdGen, text: str, *, sz: int = 2400) -> str:
    """Шапка контентного слайда — заголовок цвета BLUE по левому краю."""
    return text_box(
        idg,
        TITLE_X, TITLE_Y, TITLE_W, TITLE_H,
        para(run(text, sz=sz, bold=True, color=BLUE)),
        name="Title", anchor="ctr",
    )


def slide_number(idg: IdGen, num: int) -> str:
    """Номер слайда — белым на синей полосе."""
    return text_box(
        idg,
        100000, SLIDE_H - 500000, STRIPE_W - 200000, 350000,
        para(run(str(num), sz=1400, bold=True, color=WHITE), align="ctr"),
        name="SlideNumber", anchor="ctr",
    )


def content_slide_chrome(idg: IdGen, num: int) -> str:
    """Базовое оформление контентного слайда: синяя полоса + лого + номер."""
    parts = [
        # Синяя полоса слева (image2 → rId3)
        pic_element(idg, "rId3", 0, -4795, STRIPE_W, 5148295, name="Stripe"),
        # Логотип КФУ (image4 → rId5)
        pic_element(idg, "rId5", 134703, 93531, 684193, 580591, name="KFULogo"),
        # Номер слайда
        slide_number(idg, num),
    ]
    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Слайд 1. Титул
# ---------------------------------------------------------------------------

def slide_1() -> str:
    idg = IdGen(100)
    parts = [SLIDE_HEAD]

    # Полноэкранное фоновое изображение image1 (rId3)
    parts.append(pic_element(idg, "rId3", 0, 0, SLIDE_W, SLIDE_H, name="Background"))

    # Шапка — министерство, КФУ, ИУЭиФ, кафедра
    header_paragraphs = [
        para(run("МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ",
                 sz=1100, bold=True, color=WHITE), align="ctr"),
        para(run("РОССИЙСКОЙ ФЕДЕРАЦИИ",
                 sz=1100, bold=True, color=WHITE), align="ctr"),
        para(run("ФГАОУ ВО «КАЗАНСКИЙ (ПРИВОЛЖСКИЙ) ФЕДЕРАЛЬНЫЙ УНИВЕРСИТЕТ»",
                 sz=1100, bold=True, color=WHITE), align="ctr",
             space_before=150),
        para(run("ИНСТИТУТ УПРАВЛЕНИЯ, ЭКОНОМИКИ И ФИНАНСОВ",
                 sz=1100, bold=True, color=WHITE), align="ctr",
             space_before=150),
        para(run("КАФЕДРА ИННОВАЦИЙ И ИНВЕСТИЦИЙ",
                 sz=1100, bold=True, color=WHITE), align="ctr",
             space_before=80),
        para(run("Направление: 38.03.02 «Менеджмент»",
                 sz=1000, bold=True, color=WHITE), align="ctr",
             space_before=200),
        para(run("Профиль: Бизнес-аналитика в управленческой деятельности",
                 sz=1000, bold=True, color=WHITE), align="ctr"),
    ]
    parts.append(text_box(idg, 500000, 250000, 8144000, 1700000,
                          *header_paragraphs, name="Header", anchor="t"))

    # Белая разделительная линия
    parts.append(line_shape(idg, 1100000, 2200000, 6944000, 0,
                            color=WHITE, weight=38100))

    # Тема ВКР
    title_paragraphs = [
        para(run("ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА",
                 sz=1300, bold=True, color=WHITE), align="ctr"),
        para(run("РАЗРАБОТКА СТРАТЕГИИ ПРОДВИЖЕНИЯ", sz=2200,
                 bold=True, color=WHITE), align="ctr", space_before=200),
        para(run("ИННОВАЦИОННОГО ПРОДУКТА FACE2", sz=2200,
                 bold=True, color=WHITE), align="ctr"),
        para(run("АК БАРС БАНКА", sz=2200, bold=True, color=WHITE),
             align="ctr"),
    ]
    parts.append(text_box(idg, 500000, 2350000, 8144000, 1300000,
                          *title_paragraphs, name="Topic", anchor="t"))

    # Блок ФИО / руководитель — справа внизу
    student_paragraphs = [
        para(run("Выполнил(а):", sz=1100, bold=False, color=WHITE)),
        para(run("Гайфуллина Эльвина Илфатовна", sz=1100,
                 bold=True, color=WHITE), space_before=80),
        para(run("Группа: 14.5-216", sz=1100, bold=False, color=WHITE),
             space_before=120),
        para(run("Научный руководитель:", sz=1100, bold=False,
                 color=WHITE), space_before=200),
        para(run("к.э.н., доцент кафедры", sz=1100, bold=False,
                 color=WHITE)),
        para(run("инноваций и инвестиций", sz=1100, bold=False,
                 color=WHITE)),
    ]
    parts.append(text_box(idg, 5000000, 3700000, 3700000, 1100000,
                          *student_paragraphs, name="Student", anchor="t"))

    # Год
    parts.append(text_box(
        idg, 0, SLIDE_H - 400000, SLIDE_W, 350000,
        para(run("Казань – 2026", sz=1200, bold=True, color=WHITE),
             align="ctr"),
        name="Year", anchor="ctr",
    ))

    parts.append(SLIDE_TAIL)
    return "".join(parts)


# ---------------------------------------------------------------------------
# Слайд 2. Актуальность
# ---------------------------------------------------------------------------

def kpi_tile(idg: IdGen, x: int, y: int, w: int, h: int,
             value: str, caption: str, *, fill: str = BLUE,
             value_color: str = WHITE, caption_color: str = WHITE) -> str:
    """KPI-плитка: крупное число + подпись."""
    paragraphs = [
        para(run(value, sz=2400, bold=True, color=value_color),
             align="ctr"),
        para(run(caption, sz=900, bold=False, color=caption_color),
             align="ctr", space_before=150),
    ]
    return filled_box(idg, x, y, w, h, *paragraphs,
                      fill=fill, name="KPI", anchor="ctr")


def slide_2() -> str:
    idg = IdGen(200)
    parts = [SLIDE_HEAD, content_slide_chrome(idg, 2),
             title_block(idg, "АКТУАЛЬНОСТЬ")]

    # 4 KPI-плитки в ряд
    tile_y = 850000
    tile_h = 950000
    tile_gap = 80000
    tiles_total_w = CONTENT_W
    tile_w = (tiles_total_w - tile_gap * 3) // 4
    tile_x0 = CONTENT_X

    kpis = [
        ("572-ФЗ",  "обязательная аккредитация КБС\nс декабря 2022 года"),
        ("99,5 %",  "точность распознавания лиц\nв системе Face2"),
        ("5+ тыс.", "сотрудников АББ работают\nчерез биометрический проход"),
        ("1,16 трлн ₽", "активы группы АББ\nна конец 2025 года"),
    ]
    for i, (val, cap) in enumerate(kpis):
        x = tile_x0 + i * (tile_w + tile_gap)
        # Заменим перевод строки на двойной параграф
        cap_lines = cap.split("\n")
        cap_paras = [
            para(run(line, sz=900, bold=False, color=WHITE),
                 align="ctr", space_before=(80 if j else 150))
            for j, line in enumerate(cap_lines)
        ]
        tile_paras = [para(run(val, sz=2200, bold=True, color=WHITE),
                           align="ctr")] + cap_paras
        parts.append(filled_box(idg, x, tile_y, tile_w, tile_h,
                                *tile_paras, fill=BLUE,
                                name=f"KPI{i+1}", anchor="ctr"))

    # Текстовый блок с обоснованием актуальности
    body = [
        bullet_para(run("Биометрия — отдельный класс цифровых инноваций ",
                        sz=1500, color=BLACK),
                    run("с особыми требованиями к доверию и защите данных",
                        sz=1500, color=BLACK)),
        bullet_para(run("После 572-ФЗ обращение с биометрией возможно ",
                        sz=1500, color=BLACK),
                    run("только через КБС, аккредитованные Минцифры РФ",
                        sz=1500, color=BLACK)),
        bullet_para(run("Стандартные инструменты банковской рекламы недостаточны: ",
                        sz=1500, color=BLACK),
                    run("требуется образовательная коммуникация",
                        sz=1500, bold=True, color=BLUE),
                    run(" и работа с барьерами восприятия",
                        sz=1500, color=BLACK)),
        bullet_para(run("Face2 уже прошёл аккредитацию и сертификацию ФСБ/ФСТЭК, ",
                        sz=1500, color=BLACK),
                    run("но рынок сдерживается психологическими барьерами",
                        sz=1500, color=BLACK)),
    ]
    parts.append(text_box(idg, CONTENT_X, 2000000, CONTENT_W, 2400000,
                          *body, name="Body", anchor="t"))

    # Подпись внизу
    footer = para(run(
        "→ Требуется специализированная стратегия продвижения биометрического "
        "продукта Face2 с учётом регуляторных и репутационных особенностей",
        sz=1200, bold=True, italic=True, color=BLUE))
    parts.append(text_box(idg, CONTENT_X, 4500000, CONTENT_W, 400000,
                          footer, name="Footer", anchor="ctr"))

    parts.append(SLIDE_TAIL)
    return "".join(parts)


# ---------------------------------------------------------------------------
# Слайд 3. Цели и задачи
# ---------------------------------------------------------------------------

def slide_3() -> str:
    idg = IdGen(300)
    parts = [SLIDE_HEAD, content_slide_chrome(idg, 3),
             title_block(idg, "ЦЕЛЬ И ЗАДАЧИ ИССЛЕДОВАНИЯ")]

    # Блок «Цель»
    goal_paras = [
        para(run("ЦЕЛЬ: ", sz=1400, bold=True, color=WHITE),
             run("разработка стратегии продвижения биометрического продукта "
                 "Face2 Ак Барс Банка как высокотехнологичного "
                 "инновационного продукта повышенного риска",
                 sz=1400, bold=False, color=WHITE), align="l"),
    ]
    parts.append(filled_box(idg, CONTENT_X, 800000, CONTENT_W, 700000,
                            *goal_paras, fill=BLUE, name="Goal",
                            anchor="ctr"))

    # 5 task cards в ряд
    tasks_y = 1700000
    tasks_h = 2700000
    gap = 70000
    tasks_total_w = CONTENT_W
    task_w = (tasks_total_w - gap * 4) // 5

    tasks = [
        ("1", "ТЕОРИЯ",
         "изучить теоретические основы продвижения инновационных продуктов "
         "и специфику биометрических технологий"),
        ("2", "ОБЪЕКТ",
         "охарактеризовать Ак Барс Банк и продукт Face2 как объект продвижения, "
         "оценить ресурсную базу"),
        ("3", "РЫНОК",
         "проанализировать рынок биометрии, конкурентную среду и текущее "
         "продвижение Face2"),
        ("4", "СТРАТЕГИЯ",
         "разработать стратегию продвижения Face2 и комплекс маркетинговых "
         "мероприятий"),
        ("5", "ОЦЕНКА",
         "оценить эффективность предложенной стратегии и риски её "
         "реализации"),
    ]
    for i, (num, title, desc) in enumerate(tasks):
        x = CONTENT_X + i * (task_w + gap)
        # Верх карточки — синий с цифрой
        parts.append(filled_box(
            idg, x, tasks_y, task_w, 600000,
            para(run(num, sz=3200, bold=True, color=WHITE), align="ctr"),
            fill=BLUE, name=f"TaskNum{num}", anchor="ctr"))
        # Заголовок-категория
        parts.append(filled_box(
            idg, x, tasks_y + 600000, task_w, 400000,
            para(run(title, sz=1100, bold=True, color=WHITE), align="ctr"),
            fill=GRAY, name=f"TaskCat{num}", anchor="ctr"))
        # Описание задачи
        parts.append(filled_box(
            idg, x, tasks_y + 1000000, task_w, tasks_h - 1000000,
            para(run(desc, sz=900, bold=False, color=BLACK), align="l"),
            fill=WHITE, line_color=GRAY, line_w=6350,
            name=f"TaskDesc{num}", anchor="t"))

    # Сноска
    note = para(run("* в соответствии с Положением о ВКР ИиИ — не более 5 задач",
                    sz=900, italic=True, color=GRAY), align="l")
    parts.append(text_box(idg, CONTENT_X, 4550000, CONTENT_W, 300000,
                          note, name="Note", anchor="ctr"))

    parts.append(SLIDE_TAIL)
    return "".join(parts)


# ---------------------------------------------------------------------------
# Слайд 4. Задача 1 — теория
# ---------------------------------------------------------------------------

def slide_4() -> str:
    idg = IdGen(400)
    parts = [SLIDE_HEAD, content_slide_chrome(idg, 4),
             title_block(idg,
                         "РЕЗУЛЬТАТЫ ПО ЗАДАЧЕ 1: ТЕОРЕТИЧЕСКАЯ БАЗА")]

    # Подзаголовок
    parts.append(text_box(
        idg, CONTENT_X, 800000, CONTENT_W, 350000,
        para(run("Биометрия — высокотехнологичный инновационный продукт "
                 "повышенного риска",
                 sz=1300, italic=True, color=BLUE), align="l"),
        name="Subtitle", anchor="ctr",
    ))

    # 5 признаков как горизонтальные карточки
    feat_y = 1300000
    feat_h = 1100000
    feat_gap = 60000
    feat_w = (CONTENT_W - feat_gap * 4) // 5
    features = [
        ("НОВИЗНА", "качественно новые\nхарактеристики"),
        ("ТЕХНОЛО-\nГИЧНОСТЬ", "компьютерное зрение,\nИИ, криптозащита"),
        ("ЦЕННОСТЬ", "оплата и доступ\nбез карт и документов"),
        ("КОНКУРЕНТО-\nСПОСОБНОСТЬ", "сертификация ФСБ,\nреестр Минцифры"),
        ("ПОВЫШЕННЫЙ\nРИСК", "регуляторный, репутационный,\nтехнический"),
    ]
    for i, (head, desc) in enumerate(features):
        x = CONTENT_X + i * (feat_w + feat_gap)
        # Верх — синий с признаком
        head_paras = [
            para(run(line, sz=1000, bold=True, color=WHITE), align="ctr")
            for line in head.split("\n")
        ]
        parts.append(filled_box(
            idg, x, feat_y, feat_w, 500000,
            *head_paras, fill=BLUE, name=f"Feat{i+1}H", anchor="ctr",
        ))
        desc_paras = [
            para(run(line, sz=900, color=BLACK), align="ctr",
                 space_before=(80 if j else 0))
            for j, line in enumerate(desc.split("\n"))
        ]
        parts.append(filled_box(
            idg, x, feat_y + 500000, feat_w, feat_h - 500000,
            *desc_paras, fill=WHITE, line_color=BLUE, line_w=6350,
            name=f"Feat{i+1}D", anchor="ctr",
        ))

    # Нижняя зона — 4 особенности биометрии (две колонки)
    bx = CONTENT_X
    by = 2700000
    bw = CONTENT_W // 2 - 50000
    bh = 1750000

    left_paras = [
        para(run("Чем биометрия отличается от других банковских инноваций:",
                 sz=1200, bold=True, color=BLUE), align="l"),
        bullet_para(run("особо чувствительные персональные данные",
                        sz=1100, color=BLACK), space_before=150),
        bullet_para(run("жёсткий нормативный контур (572-ФЗ, ФСБ, ФСТЭК)",
                        sz=1100, color=BLACK)),
        bullet_para(run("психологические барьеры пользователей",
                        sz=1100, color=BLACK)),
        bullet_para(run("биометрический образец нельзя «отозвать»",
                        sz=1100, color=BLACK)),
    ]
    parts.append(text_box(idg, bx, by, bw, bh, *left_paras,
                          name="LeftCol", anchor="t"))

    right_paras = [
        para(run("Что это означает для продвижения:",
                 sz=1200, bold=True, color=BLUE), align="l"),
        bullet_para(run("информирования и стимулирования спроса ",
                        sz=1100, color=BLACK),
                    run("недостаточно", sz=1100, bold=True, color=BLUE),
                    space_before=150),
        bullet_para(run("обязательна образовательная коммуникация",
                        sz=1100, color=BLACK)),
        bullet_para(run("систематическая работа с возражениями",
                        sz=1100, color=BLACK)),
        bullet_para(run("демонстрация мер защиты данных и правовой основы",
                        sz=1100, color=BLACK)),
    ]
    parts.append(text_box(idg, bx + CONTENT_W // 2 + 50000, by, bw, bh,
                          *right_paras, name="RightCol", anchor="t"))

    parts.append(SLIDE_TAIL)
    return "".join(parts)


# ---------------------------------------------------------------------------
# Слайд 5. Задача 2 — АББ и Face2
# ---------------------------------------------------------------------------

def slide_5() -> str:
    idg = IdGen(500)
    parts = [SLIDE_HEAD, content_slide_chrome(idg, 5),
             title_block(idg, "РЕЗУЛЬТАТЫ ПО ЗАДАЧЕ 2: АК БАРС БАНК И FACE2")]

    # Левая колонка — KPI-плитки 2×2 о банке
    kpi_x = CONTENT_X
    kpi_y = 850000
    kpi_w = 1750000
    kpi_h = 750000
    kpi_gap = 60000
    kpis = [
        ("1,16 трлн ₽",  "активы группы\n2025 г."),
        ("15,7 млрд ₽",  "прибыль за\n2025 г."),
        ("934 тыс.",     "розничных\nклиентов"),
        ("A(RU)",        "рейтинг АКРА,\nстабильный"),
    ]
    for i, (val, cap) in enumerate(kpis):
        col, row = i % 2, i // 2
        x = kpi_x + col * (kpi_w + kpi_gap)
        y = kpi_y + row * (kpi_h + kpi_gap)
        cap_paras = [
            para(run(line, sz=900, color=WHITE), align="ctr",
                 space_before=(60 if j else 100))
            for j, line in enumerate(cap.split("\n"))
        ]
        parts.append(filled_box(
            idg, x, y, kpi_w, kpi_h,
            para(run(val, sz=1700, bold=True, color=WHITE), align="ctr"),
            *cap_paras,
            fill=BLUE, name=f"BankKPI{i+1}", anchor="ctr",
        ))

    # Подпись под банковскими KPI
    parts.append(text_box(
        idg, kpi_x, kpi_y + 2 * (kpi_h + kpi_gap), 2 * kpi_w + kpi_gap, 350000,
        para(run("Ак Барс Банк: крупный региональный игрок", sz=1100,
                 bold=True, color=BLUE), align="ctr"),
        name="BankCaption", anchor="ctr",
    ))

    # Правая колонка — таймлайн Face2
    tl_x = kpi_x + 2 * kpi_w + kpi_gap + 250000
    tl_y = 850000
    tl_w = CONTENT_W - (tl_x - CONTENT_X)
    tl_h = 2050000

    parts.append(text_box(
        idg, tl_x, tl_y, tl_w, 350000,
        para(run("ХРОНОЛОГИЯ FACE2 — от внутреннего эксперимента к рынку",
                 sz=1200, bold=True, color=BLUE), align="l"),
        name="TimelineTitle", anchor="t",
    ))

    # Горизонтальная линия таймлайна
    line_y = tl_y + 700000
    parts.append(line_shape(
        idg, tl_x + 100000, line_y, tl_w - 200000, 0,
        color=BLUE, weight=19050,
    ))

    # 5 точек на таймлайне
    years = [
        (2019, "Тестовый\nпроект в офисах\nАББ"),
        (2022, "Бренд Face2,\nдоступ и\nбиблиотеки"),
        (2023, "КБС, аккредитация\nМинцифры\n№ 4-А"),
        (2024, "Оплата по лицу,\nказанское метро,\nСБП"),
        (2025, "Подтверждение\nвозраста, отели\nбез паспорта"),
    ]
    n = len(years)
    available = tl_w - 200000
    step = available // (n - 1) if n > 1 else 0
    for i, (year, desc) in enumerate(years):
        cx = tl_x + 100000 + i * step
        # Кружок-маркер (овал)
        dot_size = 200000
        parts.append(filled_box(
            idg, cx - dot_size // 2, line_y - dot_size // 2,
            dot_size, dot_size,
            para(),
            fill=BLUE, name=f"Dot{year}", anchor="ctr", geom="ellipse",
        ))
        # Год
        parts.append(text_box(
            idg, cx - 350000, line_y - 500000, 700000, 300000,
            para(run(str(year), sz=1300, bold=True, color=BLUE),
                 align="ctr"),
            name=f"Year{year}", anchor="ctr",
        ))
        # Описание под линией
        desc_paras = [
            para(run(line, sz=850, color=BLACK), align="ctr",
                 space_before=(50 if j else 0))
            for j, line in enumerate(desc.split("\n"))
        ]
        parts.append(text_box(
            idg, cx - 600000, line_y + 200000, 1200000, 800000,
            *desc_paras, name=f"YearDesc{year}", anchor="t",
        ))

    # Нижняя плашка — линейка продуктов
    pl_y = 3100000
    pl_h = 1300000

    parts.append(text_box(
        idg, CONTENT_X, pl_y, CONTENT_W, 350000,
        para(run("ЛИНЕЙКА ПРОДУКТОВ FACE2", sz=1300, bold=True,
                 color=BLUE), align="l"),
        name="LineupTitle", anchor="t",
    ))

    products = [
        ("Face2Pay",     "оплата по лицу,\nбиоэквайринг",
         "банки, ритейл,\nвендинг, метро"),
        ("Face2Pass",    "контроль доступа\nи СКУД",
         "заводы, офисы,\nВУЗы, склады"),
        ("Face2Check-in","подтверждение\nвозраста и личности",
         "аэропорты,\nспортобъекты, отели"),
        ("КБС под ключ", "полная коммерческая\nбиосистема",
         "банки, госкорпорации,\nкрупные федералы"),
    ]
    pr_y = pl_y + 400000
    pr_h = pl_h - 400000
    pr_gap = 70000
    pr_w = (CONTENT_W - pr_gap * 3) // 4
    for i, (name, what, who) in enumerate(products):
        x = CONTENT_X + i * (pr_w + pr_gap)
        paras = [
            para(run(name, sz=1300, bold=True, color=WHITE), align="ctr"),
            para(run(what.replace("\n", " "), sz=900, color=WHITE),
                 align="ctr", space_before=80),
            para(run(who.replace("\n", " "), sz=850, italic=True,
                     color=WHITE), align="ctr", space_before=50),
        ]
        parts.append(filled_box(
            idg, x, pr_y, pr_w, pr_h,
            *paras, fill=BLUE, name=f"Product{i+1}", anchor="ctr",
        ))

    parts.append(SLIDE_TAIL)
    return "".join(parts)


# ---------------------------------------------------------------------------
# Слайд 6. Задача 3 — рынок и конкуренты
# ---------------------------------------------------------------------------

def slide_6() -> str:
    idg = IdGen(600)
    parts = [SLIDE_HEAD, content_slide_chrome(idg, 6),
             title_block(idg, "РЕЗУЛЬТАТЫ ПО ЗАДАЧЕ 3: РЫНОК И КОНКУРЕНТЫ")]

    # Подзаголовок
    parts.append(text_box(
        idg, CONTENT_X, 800000, CONTENT_W, 350000,
        para(run("После 572-ФЗ конкурентный рынок ограничен — поле "
                 "разделено между аккредитованными КБС",
                 sz=1300, italic=True, color=BLUE), align="l"),
        name="Subtitle", anchor="ctr",
    ))

    # Сравнительная таблица: Face2 vs Сбер vs VisionLabs vs OVision
    tbl_x = CONTENT_X
    tbl_y = 1200000
    tbl_w = CONTENT_W
    row_h = 380000
    col_w = [
        2200000,                                    # критерий
        (tbl_w - 2200000) // 4,                     # Face2
        (tbl_w - 2200000) // 4,                     # Сбер
        (tbl_w - 2200000) // 4,                     # VisionLabs
        tbl_w - 2200000 - 3 * ((tbl_w - 2200000) // 4),  # OVision
    ]

    headers = ["Критерий", "FACE2 АББ", "СБЕР", "VISIONLABS", "OVISION"]
    rows = [
        ("Аккредитация КБС Минцифры РФ",   "✓ № 4-А (2023)", "✓",          "✓ как вендор", "✓ как вендор"),
        ("Сертификация ФСБ / ФСТЭК",       "✓",              "✓",          "✓",            "✓"),
        ("Реестр отечественного ПО",       "✓",              "✓",          "✓",            "✓"),
        ("Собственный оператор биометрии", "✓ (банк)",       "✓ (банк)",   "—",            "—"),
        ("Линейка для B2B (СКУД, эквайринг)","Pay/Pass/Check","полная",     "Luna SDK",     "СКУД, эквайринг"),
        ("Региональная привязка",          "Татарстан",      "федеральный","федеральный",  "федеральный"),
        ("Цена / референты",               "средний +",      "премиум",    "премиум",      "средний"),
    ]

    # Шапка таблицы
    x = tbl_x
    for i, h in enumerate(headers):
        parts.append(filled_box(
            idg, x, tbl_y, col_w[i], row_h,
            para(run(h, sz=1000, bold=True, color=WHITE), align="ctr"),
            fill=BLUE, name=f"H{i}", anchor="ctr",
        ))
        x += col_w[i]

    # Строки таблицы
    for r, row in enumerate(rows):
        y = tbl_y + (r + 1) * row_h
        x = tbl_x
        # Зебра
        bg = LIGHT_GRAY if r % 2 == 0 else WHITE
        for i, cell in enumerate(row):
            # Подсветка колонки Face2 (i=1) — светло-синим
            cell_bg = ACCENT_LIGHT if i == 1 else bg
            cell_color = BLUE if i == 1 else BLACK
            cell_bold = (i == 0) or (i == 1)
            sz = 950 if i == 0 else 1000
            parts.append(filled_box(
                idg, x, y, col_w[i], row_h,
                para(run(cell, sz=sz, bold=cell_bold, color=cell_color),
                     align="ctr" if i > 0 else "l"),
                fill=cell_bg, line_color=GRAY, line_w=3175,
                name=f"R{r}C{i}", anchor="ctr",
            ))
            x += col_w[i]

    # Вывод
    concl_y = tbl_y + (len(rows) + 1) * row_h + 100000
    parts.append(filled_box(
        idg, CONTENT_X, concl_y, CONTENT_W, 350000,
        para(run("→ Преимущество Face2: ", sz=1100, bold=True, color=WHITE),
             run("статус оператора биометрии у банка + ESG-рейтинг + "
                 "региональная база Татарстана для пилотов",
                 sz=1100, color=WHITE), align="l"),
        fill=BLUE, name="Conclusion", anchor="ctr",
    ))

    parts.append(SLIDE_TAIL)
    return "".join(parts)


# ---------------------------------------------------------------------------
# Слайд 7. Задача 4 — стратегия
# ---------------------------------------------------------------------------

def slide_7() -> str:
    idg = IdGen(700)
    parts = [SLIDE_HEAD, content_slide_chrome(idg, 7),
             title_block(idg, "РЕЗУЛЬТАТЫ ПО ЗАДАЧЕ 4: СТРАТЕГИЯ ПРОДВИЖЕНИЯ")]

    # Подзаголовок
    parts.append(text_box(
        idg, CONTENT_X, 800000, CONTENT_W, 350000,
        para(run("Интегрированная стратегия push + pull, "
                 "адаптированная под биометрию",
                 sz=1300, italic=True, color=BLUE), align="l"),
        name="Subtitle", anchor="ctr",
    ))

    # Левая часть: маркетинговая воронка с этапом доверия
    fn_x = CONTENT_X
    fn_y = 1300000
    fn_w = 4100000
    fn_h = 3000000
    fn_step_h = fn_h // 5

    parts.append(text_box(
        idg, fn_x, fn_y - 350000, fn_w, 300000,
        para(run("ВОРОНКА С ЭТАПОМ ДОВЕРИЯ", sz=1100, bold=True,
                 color=BLUE), align="ctr"),
        name="FunnelTitle", anchor="ctr",
    ))

    funnel = [
        ("ОСВЕДОМЛЁННОСТЬ", "образовательная коммуникация",   1.00),
        ("ИНТЕРЕС",         "демонстрации и пилоты",            0.85),
        ("ДОВЕРИЕ",         "сертификация, ФСБ, 572-ФЗ",        0.70),
        ("ПРОБА",           "пилотные площадки и партнёры",      0.55),
        ("ВНЕДРЕНИЕ",       "коммерческий контракт",            0.40),
    ]
    for i, (stage, sub, frac) in enumerate(funnel):
        w = int(fn_w * frac)
        cx = fn_x + (fn_w - w) // 2
        y = fn_y + i * fn_step_h
        parts.append(filled_box(
            idg, cx, y, w, fn_step_h - 30000,
            para(run(stage, sz=1100, bold=True, color=WHITE),
                 align="ctr"),
            para(run(sub, sz=850, color=WHITE), align="ctr",
                 space_before=40),
            fill=BLUE, name=f"FN{i}", anchor="ctr",
            geom="trapezoid" if i == 0 else "rect",
        ))

    # Правая часть — комплекс 7P с акцентами для биометрии
    p_x = fn_x + fn_w + 200000
    p_y = 1300000
    p_w = CONTENT_W - (p_x - CONTENT_X)

    parts.append(text_box(
        idg, p_x, p_y - 350000, p_w, 300000,
        para(run("КОМПЛЕКС 7P ДЛЯ FACE2", sz=1100, bold=True,
                 color=BLUE), align="ctr"),
        name="PMixTitle", anchor="ctr",
    ))

    p_items = [
        ("Product",   "продуктовая линейка под B2B-сценарии"),
        ("Price",     "транзакционная модель + контракт под ключ"),
        ("Place",     "Татарстан → федеральные пилоты (Минцифры, НСПК)"),
        ("Promotion", "контент-маркетинг, отраслевые конференции, кейсы"),
        ("People",    "интеграционная команда и сопровождение клиента"),
        ("Process",   "регламент сбора биометрии через «Госуслуги»"),
        ("Physical evidence", "сертификаты ФСБ, реестр Минцифры, ESG-A"),
    ]
    p_h = (4400000 - p_y) // len(p_items)
    for i, (name, desc) in enumerate(p_items):
        y = p_y + i * p_h
        # Метка слева — синяя
        parts.append(filled_box(
            idg, p_x, y, 1300000, p_h - 30000,
            para(run(name, sz=1000, bold=True, color=WHITE), align="ctr"),
            fill=BLUE, name=f"P{i}L", anchor="ctr",
        ))
        # Описание справа
        parts.append(filled_box(
            idg, p_x + 1300000, y, p_w - 1300000, p_h - 30000,
            para(run(desc, sz=900, color=BLACK), align="l"),
            fill=LIGHT_GRAY, line_color=GRAY, line_w=3175,
            name=f"P{i}R", anchor="ctr",
        ))

    parts.append(SLIDE_TAIL)
    return "".join(parts)


# ---------------------------------------------------------------------------
# Слайд 8. Задача 5 — эффективность
# ---------------------------------------------------------------------------

def slide_8() -> str:
    """
    Слайд 8: финансовая оценка стратегии (NPV, PP, PI, IRR) +
    таблица денежных потоков + компактная полоса рисков.

    Реалистичная модель (см. Финансовая_модель_Face2.md):
        I0 = 28 млн ₽   (8 статей, см. § 3 модели)
        r  = 20 %       (ключевая ставка ЦБ 17 % + премия за риск 3 %)
        CF = [6, 16, 28, 40, 48]  млн ₽  (после налога на прибыль 25 %)

    Расчёт (в Python с двойной точностью):
        NPV = 42,90 млн ₽
        PP  = 26,6 мес.
        DPP = 32,8 мес.
        PI  = 2,53
        IRR = 59,1 %
    """
    idg = IdGen(800)
    parts = [SLIDE_HEAD, content_slide_chrome(idg, 8),
             title_block(idg,
                         "РЕЗУЛЬТАТЫ ПО ЗАДАЧЕ 5: ФИНАНСОВАЯ ЭФФЕКТИВНОСТЬ И РИСКИ")]

    # ----- Полоса параметров модели под заголовком --------------------------
    params_paras = [
        para(
            run("Параметры модели: ", sz=1000, bold=True, color=BLUE),
            run("I", sz=1000, color=BLACK),
            run("0", sz=700, color=BLACK),
            run(" = 28 млн ₽   ·   r = 20 % (ЦБ 17 % + премия 3 %)   ·   "
                "CF после налога 25 % = (6 / 16 / 28 / 40 / 48) млн ₽",
                sz=1000, color=BLACK),
            align="l",
        ),
    ]
    parts.append(text_box(
        idg, CONTENT_X, 770000, CONTENT_W, 280000,
        *params_paras, name="Params", anchor="ctr",
    ))

    # ----- Ряд из 4 финансовых KPI-плиток ----------------------------------
    kpi_y = 1080000
    kpi_h = 950000
    gap = 80000
    kpi_w = (CONTENT_W - gap * 3) // 4

    kpi_tiles = [
        ("NPV",
         "42,9 млн ₽",
         "чистый дисконтированный\nдоход стратегии"),
        ("PP / DPP",
         "27 / 33 мес.",
         "простой и дисконтированный\nсроки окупаемости"),
        ("PI",
         "2,53",
         "индекс доходности —\nна 1 ₽ инвестиций 2,53 ₽"),
        ("IRR",
         "59,1 %",
         "внутренняя норма\nдоходности (≫ r = 20 %)"),
    ]
    for i, (label, value, caption) in enumerate(kpi_tiles):
        x = CONTENT_X + i * (kpi_w + gap)
        # Метка сверху
        parts.append(filled_box(
            idg, x, kpi_y, kpi_w, 280000,
            para(run(label, sz=1300, bold=True, color=WHITE), align="ctr"),
            fill=GRAY, name=f"FinL{i}", anchor="ctr",
        ))
        # Значение и подпись
        cap_paras = [
            para(run(line, sz=850, color=WHITE), align="ctr",
                 space_before=(50 if j else 100))
            for j, line in enumerate(caption.split("\n"))
        ]
        parts.append(filled_box(
            idg, x, kpi_y + 280000, kpi_w, kpi_h - 280000,
            para(run(value, sz=2000, bold=True, color=WHITE), align="ctr"),
            *cap_paras,
            fill=BLUE, name=f"FinV{i}", anchor="ctr",
        ))

    # ----- Таблица денежных потоков ---------------------------------------
    tbl_y = 2080000
    tbl_x = CONTENT_X
    tbl_w = CONTENT_W
    label_w = 1300000
    total_w = 1100000
    year_w = (tbl_w - label_w - total_w) // 6   # 6 лет: 0..5
    cols = [label_w] + [year_w] * 6 + [total_w]
    header = ["Год / показатель", "0", "1", "2", "3", "4", "5", "Итого"]
    rows = [
        ("CF, млн ₽",       ["−28,0", "6,0",  "16,0",  "28,0",  "40,0",  "48,0",  "+138,0"]),
        ("DCF, млн ₽ (r=20 %)", ["−28,0", "5,0",  "11,1",  "16,2",  "19,3",  "19,3",  "+70,9"]),
        ("Накопл. DCF",     ["−28,0", "−23,0", "−11,9", "+4,3",  "+23,6", "+42,9", "NPV"]),
    ]

    parts.append(text_box(
        idg, CONTENT_X, tbl_y, CONTENT_W, 280000,
        para(run("ДЕНЕЖНЫЕ ПОТОКИ СТРАТЕГИИ ПРОДВИЖЕНИЯ FACE2  ",
                 sz=1100, bold=True, color=BLUE),
             run("(детальная декомпозиция I", sz=850, italic=True, color=GRAY),
             run("0", sz=600, italic=True, color=GRAY),
             run(" и CF — Финансовая_модель_Face2.md)",
                 sz=850, italic=True, color=GRAY),
             align="l"),
        name="CFTitle", anchor="ctr",
    ))

    cf_y = tbl_y + 280000
    row_h = 320000

    # Шапка
    x = tbl_x
    for i, h in enumerate(header):
        parts.append(filled_box(
            idg, x, cf_y, cols[i], row_h,
            para(run(h, sz=950, bold=True, color=WHITE),
                 align="l" if i == 0 else "ctr"),
            fill=BLUE, name=f"CFH{i}", anchor="ctr",
        ))
        x += cols[i]

    for r_idx, (label, vals) in enumerate(rows):
        y = cf_y + (r_idx + 1) * row_h
        x = tbl_x
        bg = LIGHT_GRAY if r_idx % 2 == 0 else WHITE
        # Метка строки
        parts.append(filled_box(
            idg, x, y, cols[0], row_h,
            para(run(label, sz=900, bold=True, color=BLUE), align="l"),
            fill=bg, line_color=GRAY, line_w=3175,
            name=f"CFR{r_idx}L", anchor="ctr",
        ))
        x += cols[0]
        for i, v in enumerate(vals):
            # Подсветка колонки «Итого» светло-синим
            cell_bg = ACCENT_LIGHT if i == 6 else bg
            cell_color = BLUE if i == 6 else BLACK
            cell_bold = i == 6
            parts.append(filled_box(
                idg, x, y, cols[i + 1], row_h,
                para(run(v, sz=950, bold=cell_bold, color=cell_color),
                     align="ctr"),
                fill=cell_bg, line_color=GRAY, line_w=3175,
                name=f"CFR{r_idx}C{i}", anchor="ctr",
            ))
            x += cols[i + 1]

    # ----- Полоса с 4 ключевыми рисками -----------------------------------
    risk_y = 3500000
    risk_h = 850000
    risk_gap = 60000
    risk_w = (CONTENT_W - risk_gap * 3) // 4

    parts.append(text_box(
        idg, CONTENT_X, risk_y - 280000, CONTENT_W, 250000,
        para(run("КЛЮЧЕВЫЕ РИСКИ И МЕРЫ ПО ИХ УДЕРЖАНИЮ",
                 sz=1100, bold=True, color=BLUE), align="l"),
        name="RisksTitle", anchor="ctr",
    ))

    risks = [
        ("РЕГУЛЯТОРНЫЙ",  "ужесточение 572-ФЗ, новые штрафы",
         "юр. сопровождение, аудит соответствия"),
        ("РЕПУТАЦИОННЫЙ", "утечка биометрии, волна публикаций",
         "комплаенс, мониторинг, кризис-план"),
        ("ТЕХНИЧЕСКИЙ",   "ошибки распознавания, FAR/FRR",
         "A/B-тесты, контроль качества модели"),
        ("РЫНОЧНЫЙ",      "медленная диффузия, барьеры доверия",
         "образовательные форматы, кейсы"),
    ]
    for i, (name, threat, mitigation) in enumerate(risks):
        x = CONTENT_X + i * (risk_w + risk_gap)
        # Шапка
        parts.append(filled_box(
            idg, x, risk_y, risk_w, 230000,
            para(run(name, sz=950, bold=True, color=WHITE), align="ctr"),
            fill=BLUE, name=f"RH{i}", anchor="ctr",
        ))
        # Тело
        body_paras = [
            para(run("Угроза: ", sz=800, bold=True, color=BLUE),
                 run(threat, sz=800, color=BLACK), align="l"),
            para(run("Меры: ", sz=800, bold=True, color=BLUE),
                 run(mitigation, sz=800, color=BLACK), align="l",
                 space_before=80),
        ]
        parts.append(filled_box(
            idg, x, risk_y + 230000, risk_w, risk_h - 230000,
            *body_paras, fill=WHITE, line_color=GRAY, line_w=3175,
            name=f"RB{i}", anchor="t",
        ))

    # ----- Финальный вывод -------------------------------------------------
    parts.append(filled_box(
        idg, CONTENT_X, 4500000, CONTENT_W, 400000,
        para(
            run("→ Стратегия эффективна: ",
                sz=1100, bold=True, color=WHITE),
            run("NPV > 0, IRR (59 %) ≫ r (20 %), PI = 2,5; запас прочности — "
                "снижение CF до −60 % сохраняет NPV > 0",
                sz=1100, bold=False, color=WHITE),
            align="ctr",
        ),
        fill=BLUE, name="FinalNote", anchor="ctr",
    ))

    parts.append(SLIDE_TAIL)
    return "".join(parts)


# ---------------------------------------------------------------------------
# Слайд 9. Выводы
# ---------------------------------------------------------------------------

def slide_9() -> str:
    idg = IdGen(900)
    parts = [SLIDE_HEAD, content_slide_chrome(idg, 9),
             title_block(idg, "ВЫВОДЫ")]

    # Четыре блока с выводами
    conclusions = [
        ("1", "ТЕОРЕТИЧЕСКАЯ РАМКА",
         "Биометрия — высокотехнологичный инновационный продукт повышенного "
         "риска. Стандартный маркетинг не работает: требуется образовательная "
         "коммуникация и работа с барьерами доверия."),
        ("2", "ОБЪЕКТ И РЕСУРСЫ",
         "Ак Барс Банк — крупный региональный игрок с активами 1,16 трлн ₽, "
         "рейтингом A(RU) и собственной аккредитованной КБС. Face2 интегрирован "
         "в стратегию банка 2022–2026."),
        ("3", "РЫНОК И ПОЗИЦИЯ",
         "Конкурентное поле ограничено аккредитованными КБС. Преимущество "
         "Face2 — статус оператора у банка, региональная база Татарстана и "
         "первая референтная группа федерального уровня."),
        ("4", "СТРАТЕГИЯ И ЭФФЕКТ",
         "Предложена интегрированная стратегия push + pull с маркетинговой "
         "воронкой и этапом доверия, набором 7P и комплексом мероприятий. "
         "Целевой ROI ≥ 180 % за 24 месяца."),
    ]
    cy = 850000
    cw = (CONTENT_W - 2 * 100000) // 2
    ch = (4500000 - cy) // 2 - 80000

    for i, (num, head, body) in enumerate(conclusions):
        col, row = i % 2, i // 2
        x = CONTENT_X + col * (cw + 100000)
        y = cy + row * (ch + 100000)
        # Цифра-маркер
        parts.append(filled_box(
            idg, x, y, 600000, ch,
            para(run(num, sz=4500, bold=True, color=WHITE), align="ctr"),
            fill=BLUE, name=f"CN{i}", anchor="ctr",
        ))
        # Содержимое
        ctx_paras = [
            para(run(head, sz=1300, bold=True, color=BLUE), align="l"),
            para(run(body, sz=1000, color=BLACK), align="l",
                 space_before=200),
        ]
        parts.append(filled_box(
            idg, x + 600000, y, cw - 600000, ch,
            *ctx_paras, fill=WHITE, line_color=BLUE, line_w=6350,
            name=f"CB{i}", anchor="t",
        ))

    # Финальный итог
    parts.append(filled_box(
        idg, CONTENT_X, 4540000, CONTENT_W, 380000,
        para(run("ОБЩИЙ ИТОГ: ", sz=1100, bold=True, color=WHITE),
             run("Цель ВКР достигнута, разработана работоспособная "
                 "стратегия продвижения Face2 как инновационного "
                 "продукта повышенного риска",
                 sz=1100, color=WHITE), align="ctr"),
        fill=BLUE, name="OverallNote", anchor="ctr",
    ))

    parts.append(SLIDE_TAIL)
    return "".join(parts)


# ---------------------------------------------------------------------------
# Слайд 10. Спасибо за внимание
# ---------------------------------------------------------------------------

def slide_10() -> str:
    idg = IdGen(1000)
    parts = [SLIDE_HEAD]

    # Полноэкранное фоновое изображение image1 (rId3 на слайде 10)
    parts.append(pic_element(idg, "rId3", 0, 0, SLIDE_W, SLIDE_H,
                             name="Background"))

    # СПАСИБО ЗА ВНИМАНИЕ!
    parts.append(text_box(
        idg, 0, 1900000, SLIDE_W, 800000,
        para(run("СПАСИБО ЗА ВНИМАНИЕ!", sz=4000, bold=True,
                 color=WHITE, font="Times New Roman"), align="ctr"),
        name="Thanks", anchor="ctr",
    ))

    # Подпись внизу
    parts.append(text_box(
        idg, 0, 3000000, SLIDE_W, 600000,
        para(run("Готова ответить на вопросы", sz=1500, italic=True,
                 color=WHITE), align="ctr"),
        name="Sub", anchor="ctr",
    ))

    parts.append(SLIDE_TAIL)
    return "".join(parts)


# ---------------------------------------------------------------------------
# Сборка pptx
# ---------------------------------------------------------------------------

SLIDES = {
    1: slide_1,
    2: slide_2,
    3: slide_3,
    4: slide_4,
    5: slide_5,
    6: slide_6,
    7: slide_7,
    8: slide_8,
    9: slide_9,
    10: slide_10,
}


def build(template_path: Path, output_path: Path) -> None:
    """Собирает новый pptx, заменяя только тела слайдов."""
    with zipfile.ZipFile(template_path, "r") as src:
        names = src.namelist()
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as dst:
            for name in names:
                data = src.read(name)
                # Заменяем только ppt/slides/slideN.xml
                if name.startswith("ppt/slides/slide") and name.endswith(".xml") \
                        and "/_rels/" not in name:
                    # извлекаем номер
                    base = name.rsplit("/", 1)[1]  # slideN.xml
                    num = int(base.replace("slide", "").replace(".xml", ""))
                    if num in SLIDES:
                        data = SLIDES[num]().encode("utf-8")
                dst.writestr(name, data)


if __name__ == "__main__":
    base_dir = Path(__file__).resolve().parent
    template = base_dir / "Шаблон презентации ВКР ИиИ_2026.pptx"
    output = base_dir / "ВКР_Face2_АкБарс_презентация.pptx"
    build(template, output)
    print(f"Wrote: {output} ({output.stat().st_size:,} bytes)")
