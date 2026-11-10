# -*- coding: utf-8 -*-
"""把《金融与经济》投稿版 Markdown 排成 DOCX。

版式依据该刊《投稿模板》及其 25 条“详细批注”：
A4，上下 2.54 cm、左右 3.18 cm；单倍行距、段前段后 0；
主标题小二号黑体加粗居中，副标题四号宋体，作者小四号黑体右对齐；
摘要五号楷体，正文五号宋体，一级标题四号黑体，二级标题五号楷体；
三线表，表题在上、图题在下；参考文献 [1] 编号、小五号宋体。
公式用 Word 原生公式（OMML），独立公式左缩四字、右侧连续编号。
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt
from latex2mathml.converter import convert as latex2mathml
from mathml2omml import convert as mml2omml

ROOT = Path(__file__).resolve().parents[1]
DRAFTS = ROOT / "drafts"
SRC = DRAFTS / "manuscript_jjyj_v1.md"
OUT = DRAFTS / "manuscript_jjyj_v1.docx"

_ARGS = [a for a in sys.argv[1:] if not a.startswith("--")]
if len(_ARGS) >= 2:
    SRC = Path(_ARGS[0]).resolve()
    OUT = Path(_ARGS[1]).resolve()
else:
    print("用法: python md_to_docx_journal.py <源.md> <输出.docx>", file=sys.stderr)
    sys.exit(2)

BODY_PT = 10.5          # 五号
SMALL_PT = 9.0          # 小五号
INDENT_PT = BODY_PT * 2  # 首行缩进两字
TEXT_WIDTH_CM = 14.64    # 21 - 3.18*2
TABLE_AVAIL_CM = 14.64   # 表格可用宽度，与版心一致（页面改动时同步改这里）

TABLE_COUNTER = 0
EQUATION_COUNTER = 0
FOOTNOTES = []          # 脚注正文，按出现顺序编号
LAST_CAPTION = None     # 最近一个表/图题段落，脚注引用挂在其末尾
MISSING_IMAGES = []     # 找不到的图片路径，结束时汇总报告


# ---------------------------------------------------------------- 基础排版
def set_font(run, size=BODY_PT, bold=False, ea="宋体", italic=False):
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), ea)


def _omml_xml(tex: str) -> str:
    omml = mml2omml(latex2mathml(tex))
    body = omml[len("<m:oMath>"):-len("</m:oMath>")]
    return (
        '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        + body + "</m:oMath>"
    )


def add_inline_math(paragraph, tex, size=BODY_PT):
    try:
        run = paragraph.add_run()
        run.font.size = Pt(size)
        run._r.append(parse_xml(_omml_xml(tex)))
    except Exception:
        r = paragraph.add_run(tex)
        set_font(r, size=size)


def add_equation(doc, tex, num):
    """独立公式：左缩四字，右端编号。"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.left_indent = Pt(42)
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_WIDTH_CM - 0.1), WD_TAB_ALIGNMENT.RIGHT)
    try:
        run = p.add_run()
        run.font.size = Pt(BODY_PT)
        run._r.append(parse_xml(_omml_xml(tex)))
    except Exception:
        r = p.add_run(tex)
        set_font(r, size=BODY_PT)
    p.add_run().add_tab()
    r = p.add_run(f"（{num}）")
    set_font(r, size=BODY_PT)


def add_runs(p, text, size=BODY_PT, ea="宋体", bold=False, star_pt=None):
    text = text.replace("`", "")
    tail = ""
    if star_pt and re.fullmatch(r".*\*{1,3}", text):
        # 表格单元格的显著性星号单独成段，用小字号排（沿用例文：9pt 数字配 4.5pt 星号）
        text, tail = text.rstrip("*"), text[len(text.rstrip("*")):]
    parts = re.split(r"(\$[^$]+\$|\*\*.*?\*\*|(?<!\*)\*(?!\*)[^*]+(?<!\*)\*(?!\*))", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("$") and part.endswith("$") and len(part) > 2:
            add_inline_math(p, part[1:-1], size=size)
        elif part.startswith("**") and part.endswith("**") and len(part) > 4:
            r = p.add_run(part[2:-2])
            set_font(r, size=size, bold=True, ea=ea)
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            r = p.add_run(part[1:-1])
            set_font(r, size=size, bold=bold, ea=ea)
            r.italic = True
        else:
            r = p.add_run(part)
            set_font(r, size=size, bold=bold, ea=ea)
    if tail:
        r = p.add_run(tail)
        set_font(r, size=star_pt, bold=bold, ea=ea)


def body_paragraph(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Pt(INDENT_PT)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    add_runs(p, text)
    return p


def blank_line(doc):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("")
    set_font(r, size=BODY_PT)
    return p


def set_outline_level(p, level):
    """给普通段落写大纲级别，导航窗格/目录能认，但不改变任何排版外观。"""
    pPr = p._p.get_or_add_pPr()
    lvl = OxmlElement("w:outlineLvl")
    lvl.set(qn("w:val"), str(level))
    pPr.append(lvl)
    return p


def heading1(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(28)   # 四号两字
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_font(r, size=14, bold=True, ea="黑体")
    return set_outline_level(p, 0)


def heading2(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(INDENT_PT)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_font(r, size=BODY_PT, ea="楷体")
    return set_outline_level(p, 1)


def heading3(doc, text):
    # 三级标题：模板未单列样式，沿用二级的缩进与字号，加粗区分
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(INDENT_PT)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_font(r, size=BODY_PT, bold=True, ea="楷体")
    return set_outline_level(p, 2)


def caption(doc, text):
    text = re.sub(r"^([表图])\s*(\d+)[　\s]*", r"\1\2　　", text)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.keep_with_next = True
    add_runs(p, text, size=BODY_PT, ea="黑体")
    globals()["LAST_CAPTION"] = p
    return p


def footnote_here(para, text):
    """在段末插入 Word 原生脚注引用，正文存入 FOOTNOTES 待写出。"""
    FOOTNOTES.append(text)
    run = para.add_run()
    run.font.superscript = True
    run._element.append(parse_xml(
        '<w:footnoteReference %s w:id="%d"/>' % (nsdecls("w"), len(FOOTNOTES))))
    return run


def _footnotes_part(doc):
    """生成 /word/footnotes.xml（两条分隔线 + 各条脚注）。"""
    from xml.sax.saxutils import escape

    def para(inner):
        return (
            '<w:p><w:pPr><w:snapToGrid w:val="0"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:jc w:val="left"/></w:pPr>%s</w:p>' % inner
        )

    body = [
        '<w:footnote w:type="separator" w:id="-1">'
        + para('<w:r><w:separator/></w:r>') + '</w:footnote>',
        '<w:footnote w:type="continuationSeparator" w:id="0">'
        + para('<w:r><w:continuationSeparator/></w:r>') + '</w:footnote>',
    ]
    for i, t in enumerate(FOOTNOTES, 1):
        body.append(
            '<w:footnote w:id="%d">' % i + para(
                '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                '<w:footnoteRef/></w:r>'
                '<w:r><w:rPr>'
                '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"'
                ' w:eastAsia="宋体"/>'
                '<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
                '<w:t xml:space="preserve"> %s</w:t></w:r>' % escape(t)
            ) + '</w:footnote>'
        )

    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:footnotes ' + nsdecls("w") + '>'
           + ''.join(body) + '</w:footnotes>')
    return Part(
        PackURI('/word/footnotes.xml'),
        'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
        xml.encode('utf-8'),
        doc.part.package,
    )


def attach_footnotes(doc):
    if not FOOTNOTES:
        return
    part = _footnotes_part(doc)
    doc.part.relate_to(part, RT.FOOTNOTES)
    # 带圈数字、每页重新编号（模板批注 16）
    doc.sections[0]._sectPr.insert(0, parse_xml(
        '<w:footnotePr %s><w:numFmt w:val="decimalEnclosedCircle"/>'
        '<w:numRestart w:val="eachPage"/></w:footnotePr>' % nsdecls("w")))


def note(doc, text, size=SMALL_PT, center=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0) if center else Pt(size * 2)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.keep_together = True
    p.paragraph_format.keep_with_next = True
    if re.search(r"\*{1,3}、\*{1,3}", text):
        # 显著性符号说明：星号是字面量，不能被粗体正则吞掉
        parts = re.split(r"(\$[^$]+\$)", text)
        for part in parts:
            if part.startswith("$") and part.endswith("$") and len(part) > 2:
                add_inline_math(p, part[1:-1], size=size)
            elif part:
                r = p.add_run(part)
                set_font(r, size=size)
    else:
        add_runs(p, text, size=size)
    return p


def panel_label(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.keep_with_next = True
    add_runs(p, text, size=SMALL_PT, bold=True)
    return p


def ref_heading(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("[参考文献]")
    set_font(r, size=BODY_PT, ea="黑体")


def reference(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.left_indent = Pt(0)
    p.paragraph_format.first_line_indent = Pt(21)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    add_runs(p, text, size=SMALL_PT)
    return p

# ---------------------------------------------------------------- 表格
def parse_table(lines):
    rows = []
    aligns = []
    for idx, line in enumerate(lines):
        cells = [c.strip().replace(r"\|", "|") for c in re.split(r"(?<!\\)\|", line.strip().strip("|"))]
        if idx == 0:
            rows.append(cells)
            continue
        if all(re.fullmatch(r":?-{2,}:?", c) for c in cells):
            aligns = []
            for c in cells:
                if c.startswith(":") and c.endswith(":"):
                    aligns.append(WD_ALIGN_PARAGRAPH.CENTER)
                elif c.endswith(":"):
                    aligns.append(WD_ALIGN_PARAGRAPH.RIGHT)
                else:
                    aligns.append(WD_ALIGN_PARAGRAPH.LEFT)
            continue
        rows.append(cells)
    return rows, aligns


def set_three_line(table):
    tbl = table._tbl
    for st in tbl.tblPr.findall(qn("w:tblStyle")):
        tbl.tblPr.remove(st)
    for old in tbl.tblPr.findall(qn("w:tblBorders")):
        tbl.tblPr.remove(old)
    bd = OxmlElement("w:tblBorders")
    for side, val, sz in [("top", "single", "12"), ("bottom", "single", "12"),
                          ("left", "none", "0"), ("right", "none", "0"),
                          ("insideH", "none", "0"), ("insideV", "none", "0")]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), sz)
        el.set(qn("w:szB"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        bd.append(el)
    tbl.tblPr.insert_element_before(
        bd, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
        "w:tblCaption", "w:tblDescription", "w:tblPrChange",
    )
    seen = set()
    for cell in table.rows[0].cells:
        if id(cell._tc) in seen:
            continue
        seen.add(id(cell._tc))
        tcPr = cell._tc.get_or_add_tcPr()
        for old in tcPr.findall(qn("w:tcBorders")):
            tcPr.remove(old)
        tb = OxmlElement("w:tcBorders")
        b = OxmlElement("w:bottom")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "6")
        b.set(qn("w:szB"), "6")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "000000")
        tb.append(b)
        tcPr.insert_element_before(
            tb, "w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
            "w:tcFitText", "w:vAlign", "w:hideMark", "w:headers",
            "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange",
        )


def _text_width_cm(s, font_pt=9, bold=False):
    s = re.sub(r"\$([^$]*)\$", r"\1", s).replace("**", "").replace("`", "")
    w = 0.0
    for ch in s:
        if ord(ch) >= 0x4E00:
            w += font_pt
        elif ch in "−–—":
            w += font_pt * 0.60
        elif ch in "MWmw@%":
            w += font_pt * 0.95
        elif ch.isupper():
            w += font_pt * 0.70
        elif ch.isdigit() or ch in "+-=.,()[]/<>":
            w += font_pt * 0.50
        else:
            w += font_pt * 0.55
    if bold:
        w *= 1.03
    return w / 72 * 2.54


def _is_panel_row(row):
    return len(row) > 2 and bool(row[0]) and not any(row[1:])


def _header_min_widths(rows, font_pt):
    pad = 0.30
    return [_text_width_cm(c, font_pt, bold=True) + pad for c in rows[0]]


def _body_pref_widths(rows, font_pt):
    ncols = len(rows[0])
    pad = 0.30
    body = [r for r in rows[1:] if not _is_panel_row(r)]
    return [max([_text_width_cm(r[j], font_pt) for r in body if j < len(r)] + [0.0]) + pad
            for j in range(ncols)]


def _fit_table_font(rows, ncols, avail_cm=None):
    avail_cm = TABLE_AVAIL_CM if avail_cm is None else avail_cm
    for pt in (9, 8.5, 8, 7.5, 7):
        mins = _header_min_widths(rows, pt)
        pref = _body_pref_widths(rows, pt)
        if sum(max(mins[j], pref[j]) for j in range(ncols)) <= avail_cm:
            return pt
    return 7


def set_col_widths(table, rows, aligns, font_pt=9, avail_cm=None):
    avail_cm = TABLE_AVAIL_CM if avail_cm is None else avail_cm
    ncols = len(rows[0])
    mins = [max(0.9, w) for w in _header_min_widths(rows, font_pt)]
    pref = _body_pref_widths(rows, font_pt)
    if sum(mins) > avail_cm:
        scale = avail_cm / sum(mins)
        widths = [w * scale for w in mins]
    else:
        widths = [max(mins[j], pref[j]) for j in range(ncols)]
        if sum(widths) <= avail_cm:
            leftover = avail_cm - sum(widths)
            if leftover > 1e-9:
                base = sum(widths)
                widths = [w + leftover * w / base for w in widths]
        else:
            remaining = avail_cm - sum(mins)
            desired = [max(0.0, pref[j] - mins[j]) for j in range(ncols)]
            need = sum(desired)
            widths = (list(mins) if need <= 0 else
                      [mins[j] + remaining * desired[j] / need for j in range(ncols)])
    table.autofit = False
    tbl = table._tbl
    for gridCol, w in zip(tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol")), widths):
        gridCol.set(qn("w:w"), str(int(round(w / 2.54 * 1440))))
    tblPr = tbl.tblPr
    for old in tblPr.findall(qn("w:tblLayout")):
        tblPr.remove(old)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    for old in tblPr.findall(qn("w:tblW")):
        tblPr.remove(old)
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:type"), "dxa")
    tw.set(qn("w:w"), str(int(avail_cm / 2.54 * 1440)))
    tblPr.append(tw)
    for row in table.rows:
        col = 0
        for tc in row._tr.tc_lst:
            span = tc.grid_span
            tc.width = Cm(sum(widths[col:col + span]))
            col += span
    for old in tblPr.findall(qn("w:tblCellMar")):
        tblPr.remove(old)
    mar = OxmlElement("w:tblCellMar")
    for side, val in [("left", "45"), ("right", "45"), ("top", "0"), ("bottom", "0")]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), val)
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)
    for row in table.rows[1:]:
        col = 0
        for tc in row._tr.tc_lst:
            span = tc.grid_span
            if span == 1 and aligns[col] != WD_ALIGN_PARAGRAPH.LEFT:
                tc.get_or_add_tcPr().append(OxmlElement("w:noWrap"))
            col += span


def add_table(doc, lines):
    rows, aligns = parse_table(lines)
    if not rows:
        return
    ncols = max(len(r) for r in rows)
    for r in rows:
        while len(r) < ncols:
            r.append("")
    while len(aligns) < ncols:
        aligns.append(WD_ALIGN_PARAGRAPH.CENTER)
    font_pt = _fit_table_font(rows, ncols)
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        if i > 0 and _is_panel_row(row):
            cell = table.cell(i, 0).merge(table.cell(i, ncols - 1))
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_runs(p, row[0], size=font_pt)
            continue
        for j in range(ncols):
            cell = table.cell(i, j)
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i == 0:
                add_runs(p, row[j], size=font_pt, bold=True)
            else:
                add_runs(p, row[j], size=font_pt, star_pt=round(font_pt * 0.5, 1))
    set_col_widths(table, rows, aligns, font_pt=font_pt)
    set_three_line(table)
    tr = table.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    trPr.append(trPr.makeelement(qn("w:tblHeader"), {qn("w:val"): "true"}))
    for row in table.rows:
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.keep_with_next = True


# ---------------------------------------------------------------- 页面
def add_page_number(section):
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.space_before = Pt(0)
    footer.paragraph_format.space_after = Pt(0)
    run = footer.add_run()
    set_font(run, size=SMALL_PT)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)


def add_image(doc, path):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(path), width=Cm(13.6))


# ---------------------------------------------------------------- 主流程
def main():
    global EQUATION_COUNTER
    text = SRC.read_text(encoding="utf-8")
    lines = text.splitlines()

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.18)
    section.right_margin = Cm(3.18)
    section.header_distance = Cm(1.5)
    section.footer_distance = Cm(1.75)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(BODY_PT)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    add_page_number(section)

    in_refs = False
    caption_is_figure = False
    i = 0
    while i < len(lines):
        raw = lines[i].rstrip()
        s = raw.strip()
        if not s:
            i += 1
            continue

        # 指令
        if s.startswith("[[") and "]]" in s:
            tag, val = s[2:].split("]]", 1)
            val = val.strip()
            if tag == "PAGEBREAK":
                doc.add_page_break()
            elif tag == "TITLE":
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Pt(0)
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(val)
                set_font(r, size=18, bold=True, ea="黑体")
                blank_line(doc)
            elif tag == "SUBTITLE":
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Pt(0)
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(val)
                set_font(r, size=14, bold=False, ea="宋体")
                blank_line(doc)
            elif tag == "AUTHORS":
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.first_line_indent = Pt(0)
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(val)
                set_font(r, size=12, bold=False, ea="黑体")
                blank_line(doc)
            elif tag == "ABSTRACT":
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Pt(INDENT_PT)
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                add_runs(p, val, size=BODY_PT, ea="楷体")
            elif tag == "KEYWORDS":
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.first_line_indent = Pt(21)
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run("[关键词]")
                set_font(r, size=BODY_PT, ea="黑体")
                r = p.add_run(val)
                set_font(r, size=BODY_PT, ea="楷体")
            elif tag == "CLC":
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.first_line_indent = Pt(21)
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(val)
                set_font(r, size=BODY_PT, ea="黑体")
            elif tag == "DOI":
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.first_line_indent = Pt(21)
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(val)
                set_font(r, size=BODY_PT, ea="黑体")
            elif tag == "AUTHORINFO":
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Pt(21)
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                add_runs(p, val, size=BODY_PT, ea="楷体")
            i += 1
            continue

        # 独立公式
        m = re.fullmatch(r"\$\$(.+)\$\$", s)
        if m:
            EQUATION_COUNTER += 1
            add_equation(doc, m.group(1).strip(), EQUATION_COUNTER)
            i += 1
            continue

        # 图片
        m = re.match(r"!\[[^\]]*\]\(([^)]+)\)", s)
        if m:
            rel = m.group(1)
            img = (SRC.parent / rel).resolve()
            if img.exists():
                add_image(doc, img)
            else:
                MISSING_IMAGES.append(rel)
                print(f"[WARN] 图未找到，已跳过：{rel}  (解析为 {img})", file=sys.stderr)
            i += 1
            continue

        # 表格
        if s.startswith("|"):
            tbl_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tbl_lines.append(lines[i].strip())
                i += 1
            add_table(doc, tbl_lines)
            continue

        # 标题
        if s.startswith("# "):
            h = s[2:].strip()
            if h == "参考文献":
                ref_heading(doc)
                in_refs = True
            else:
                in_refs = False
                heading1(doc, h)
            i += 1
            continue
        if s.startswith("## "):
            heading2(doc, s[3:].strip())
            i += 1
            continue
        if s.startswith("### "):
            heading3(doc, s[4:].strip())
            i += 1
            continue

        # 表/图题
        if s.startswith("**") and s.endswith("**") and len(s) < 120 \
                and re.match(r"^\*\*[表图]\s*\d+[　\s]", s):
            caption(doc, s[2:-2])
            caption_is_figure = s.startswith("**图")
            i += 1
            continue

        # 表面板标题
        if s.startswith("**Panel ") and s.endswith("**"):
            panel_label(doc, s[2:-2])
            i += 1
            continue

        # 表的数据来源：挂到表题末尾，排成页下脚注
        if s.startswith("脚注：") and LAST_CAPTION is not None:
            footnote_here(LAST_CAPTION, s[3:])
            i += 1
            continue

        # 表注 / 资料来源（图注下的资料来源按模板居中）
        if s.startswith("注：") or s.startswith("资料来源："):
            center = caption_is_figure and s.startswith("资料来源：")
            note(doc, s, center=center)
            if center:
                caption_is_figure = False
            i += 1
            continue

        # 参考文献条目
        if in_refs and re.match(r"^\[\d+\]", s):
            reference(doc, s)
            i += 1
            continue

        # 分隔线
        if s == "---":
            i += 1
            continue

        # 引用块
        if s.startswith(">"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Pt(BODY_PT * 2)
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            add_runs(p, s.lstrip(">").strip(), size=SMALL_PT)
            i += 1
            continue

        # 有序列表
        m = re.match(r"^(\d+)[.)]\s+(.*)$", s)
        if m:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Pt(36)
            p.paragraph_format.first_line_indent = Pt(-18)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            add_runs(p, m.group(1) + ". " + m.group(2))
            i += 1
            continue

        # 列表项
        if re.match(r"^-\s+", s):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Pt(36)
            p.paragraph_format.first_line_indent = Pt(-12)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            add_runs(p, "• " + re.sub(r"^-\s+", "", s))
            i += 1
            continue

        body_paragraph(doc, s)
        i += 1

    attach_footnotes(doc)
    doc.save(OUT)
    print("saved", OUT)
    print("footnotes:", len(FOOTNOTES))
    print("equations:", EQUATION_COUNTER)
    if MISSING_IMAGES:
        print(f"[WARN] {len(MISSING_IMAGES)} 张图未插入（路径不存在）：", file=sys.stderr)
        for p in MISSING_IMAGES:
            print("       " + p, file=sys.stderr)


if __name__ == "__main__":
    main()
