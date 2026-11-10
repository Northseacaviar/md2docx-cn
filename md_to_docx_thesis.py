# -*- coding: utf-8 -*-
"""把 manuscript_v1.md 转成通用硕士论文格式的 DOCX（非学校模板）。"""

from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from latex2mathml.converter import convert as latex2mathml
from mathml2omml import convert as mml2omml

ROOT = Path(__file__).resolve().parents[1]
DRAFTS = ROOT / "drafts"
SRC = DRAFTS / "manuscript_v1.md"
OUT = DRAFTS / "manuscript_v1.docx"

if len(sys.argv) >= 3:
    SRC = Path(sys.argv[1]).resolve()
    OUT = Path(sys.argv[2]).resolve()
else:
    print("用法: python md_to_docx_thesis.py <源.md> <输出.docx>", file=sys.stderr)
    sys.exit(2)

MISSING_IMAGES = []


def set_run_font(run, size=12, bold=False, color=None):
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    if color:
        run.font.color.rgb = color



def latex_inline_to_text(tex: str) -> str:
    """把论文里用到的简单 LaTeX 行内公式转成 Unicode 文本。"""
    tex = tex.strip()
    pairs = [
        (r"\hat{\alpha}", "α̂"), (r"\hat{\beta}", "β̂"),
        (r"\alpha", "α"), (r"\beta", "β"), (r"\gamma", "γ"), (r"\delta", "δ"),
        (r"\varepsilon", "ε"), (r"\theta", "θ"), (r"\mu", "μ"), (r"\lambda", "λ"),
        (r"\sum", "Σ"), (r"\ln", "ln"), (r"\left", ""), (r"\right", ""),
        (r"\,", ""), (r"\;", ""), (r"\!", ""), (r"\%", "%"),
    ]
    for a, b in pairs:
        tex = tex.replace(a, b)
    tex = re.sub(r"\\mathbf\{([^}]*)\}", r"\1", tex)
    tex = re.sub(r"\\[a-zA-Z]+", "", tex)
    tex = tex.replace("{", "").replace("}", "")
    return tex


def add_omml(doc, tex: str):
    """插入 Word 原生公式；失败时退回 Unicode 文本。"""
    try:
        omml = mml2omml(latex2mathml(tex))
        body = omml[len("<m:oMath>"):-len("</m:oMath>")]
        xml = (
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            + body + "</m:oMath>"
        )
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        p.add_run()._r.append(parse_xml(xml))
    except Exception:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_runs(p, latex_inline_to_text(tex), size=12)


def add_inline_omml(paragraph, tex: str, size: int = 12):
    """在正文中插入行内 Word 公式；失败时退回 Unicode。"""
    try:
        omml = mml2omml(latex2mathml(tex))
        body = omml[len("<m:oMath>"):-len("</m:oMath>")]
        xml = (
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            + body + "</m:oMath>"
        )
        run = paragraph.add_run()
        run.font.size = Pt(size)
        run._r.append(parse_xml(xml))
    except Exception:
        r = paragraph.add_run(latex_inline_to_text(tex))
        set_run_font(r, size=size)


SUP_RE = re.compile(r"\^(\d+)")


def add_plain_runs(paragraph, text, size, bold):
    """普通文字，把 ^1、^2 形式的作者单位角标排成上标。"""
    pos = 0
    for m in SUP_RE.finditer(text):
        if m.start() > pos:
            r = paragraph.add_run(text[pos:m.start()])
            set_run_font(r, size=size, bold=bold)
        r = paragraph.add_run(m.group(1))
        set_run_font(r, size=size, bold=bold)
        r.font.superscript = True
        pos = m.end()
    if pos < len(text):
        r = paragraph.add_run(text[pos:])
        set_run_font(r, size=size, bold=bold)


def add_runs(paragraph, text, size=12, bold=False, omml=False):
    """处理 **粗体**、*斜体*、`代码`、$公式$、^上标。"""
    text = text.replace("`", "")
    parts = re.split(r"(\$[^$]+\$|\*\*.*?\*\*|(?<!\*)\*(?!\*)[^*]+(?<!\*)\*(?!\*))", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("$") and part.endswith("$") and len(part) > 2:
            if omml:
                add_inline_omml(paragraph, part[1:-1], size=size)
            else:
                r = paragraph.add_run(latex_inline_to_text(part[1:-1]))
                set_run_font(r, size=size, bold=bold)
        elif part.startswith("**") and part.endswith("**") and len(part) > 4:
            r = paragraph.add_run(part[2:-2])
            set_run_font(r, size=size, bold=True)
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            r = paragraph.add_run(part[1:-1])
            set_run_font(r, size=size, bold=bold)
            r.italic = True
        else:
            add_plain_runs(paragraph, part, size, bold)


def body_paragraph(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(24)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(28)
    p.paragraph_format.space_after = Pt(0)
    add_runs(p, text, omml=True)
    return p


def front_paragraph(doc, text, size=10.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    """首页摘要/关键词/作者简介等：不缩进，固定 20 磅行距（同模板 Normal）。"""
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(20)
    p.paragraph_format.space_after = Pt(0)
    add_runs(p, text, size=size, omml=True)
    return p


def heading(doc, text, level):
    styles = {1: "Heading 1", 2: "Heading 2", 3: "Heading 3"}
    p = doc.add_paragraph(style=styles.get(level, "Heading 3"))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if level == 1 else WD_ALIGN_PARAGRAPH.LEFT
    # 模板 heading 3（即二级标题）用 firstLine=482，即 2 字符缩进
    p.paragraph_format.first_line_indent = Pt(24) if level > 1 else Pt(0)
    p.paragraph_format.space_before = Pt(12 if level > 1 else 18)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    sizes = {1: 16, 2: 14, 3: 12}
    r = p.add_run(text)
    set_run_font(r, size=sizes.get(level, 12), bold=True)
    r.font.color.rgb = RGBColor(0, 0, 0)
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
    return p


def caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    add_runs(p, text, size=10.5, bold=True)
    return p


def note(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_after = Pt(6)
    add_runs(p, text, size=9)
    return p


def reference(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(0)
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(28)
    p.paragraph_format.space_after = Pt(0)
    add_runs(p, text, size=10.5)
    return p


def note_item(doc, text):
    """注释区条目：同参考文献的排版，但顶格、小五号。"""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(0)
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(28)
    p.paragraph_format.space_after = Pt(0)
    add_runs(p, text, size=10.5)
    return p


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


def set_three_line(table, hdr_bottom=True, group_rows=()):
    """把表格边框设成三线表：顶线、表头下线、底线，无竖线与内部横线。"""
    tbl = table._tbl
    # 1) 摘掉 Table Grid，改用无边框底
    for st in tbl.tblPr.findall(qn("w:tblStyle")):
        tbl.tblPr.remove(st)
    # 2) 表格级边框
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
    # 3) 表头行下边框 0.75pt（逐个单元格设，合并单元格也能覆盖）
    if hdr_bottom:
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
    # 4) 可选分组线
    for idx in group_rows:
        for cell in table.rows[idx].cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tb = OxmlElement("w:tcBorders")
            b = OxmlElement("w:top")
            b.set(qn("w:val"), "single")
            b.set(qn("w:sz"), "4")
            b.set(qn("w:szB"), "4")
            b.set(qn("w:space"), "0")
            b.set(qn("w:color"), "000000")
            tb.append(b)
            tcPr.insert_element_before(
                tb, "w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
                "w:tcFitText", "w:vAlign", "w:hideMark", "w:headers",
                "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange",
            )


def _text_width_cm(s, font_pt=9):
    """按字符类型估算一行文字所需宽度（厘米）。中日韩字符按全角，其余按半角。"""
    s = re.sub(r"\$([^$]*)\$", r"\1", s).replace("**", "").replace("`", "")
    pt = font_pt
    w = 0.0
    for ch in s:
        w += pt if ord(ch) > 0x2E80 else pt * 0.52
    return w / 72 * 2.54


def set_col_widths(table, rows, aligns, avail_cm=15.4, font_pt=9):
    """按每列最长内容分配列宽，使单元格内容尽量排在一行。"""
    ncols = len(rows[0])
    pad = 0.30  # 左右单元格内边距合计（厘米）
    natural = []
    for j in range(ncols):
        w = max(_text_width_cm(r[j], font_pt) for r in rows if j < len(r)) + pad
        natural.append(w)
    total = sum(natural)
    widths = list(natural)
    if total <= avail_cm:
        # 余量按内容宽度比例分配，但每列最多扩到 max(自身内容宽, 平均列宽)。
        # 只按比例分会把“模型”“类别”这类本就长的文字列撑出大片空白；
        # 只给数字列又会把三列表的“样本数”扩到 7cm。设上限后两者都不会发生。
        avg = avail_cm / ncols
        caps = [max(w, avg) for w in natural]
        pool = list(range(ncols))
        remaining = avail_cm
        while pool:
            base = sum(natural[j] for j in pool)
            capped = [j for j in pool
                      if remaining * natural[j] / base > caps[j] + 1e-9]
            if not capped:
                for j in pool:
                    widths[j] = remaining * natural[j] / base
                break
            for j in capped:
                widths[j] = caps[j]
                remaining -= caps[j]
                pool.remove(j)
    else:
        # 超宽时按比例压缩，但给每列留最小宽度
        floor = 1.15
        widths = [max(floor, w * avail_cm / total) for w in natural]
        s = sum(widths)
        widths = [w * avail_cm / s for w in widths]

    table.autofit = False
    tbl = table._tbl
    # tblGrid 必须与单元格宽度一致，否则 WPS/LibreOffice 按等分网格排版
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
        for j, cell in enumerate(row.cells):
            cell.width = Cm(widths[j])
    # 单元格左右内边距压到 0.08cm，把空间让给文字
    for old in tblPr.findall(qn("w:tblCellMar")):
        tblPr.remove(old)
    mar = OxmlElement("w:tblCellMar")
    for side, val in [("left", "45"), ("right", "45"), ("top", "0"), ("bottom", "0")]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), val)
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)

    # 数值列不让 Word 自行断行
    for row in table.rows[1:]:
        for j, cell in enumerate(row.cells):
            if aligns[j] != WD_ALIGN_PARAGRAPH.LEFT:
                tcPr = cell._tc.get_or_add_tcPr()
                tcPr.append(OxmlElement("w:noWrap"))


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
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j in range(ncols):
            cell = table.cell(i, j)
            text = row[j]
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            # 表头与数据同列同轴：表头沿用该列的对齐方式
            p.alignment = aligns[j]
            if i == 0:
                add_runs(p, text, size=9, bold=True)
            else:
                add_runs(p, text, size=9)
    set_col_widths(table, rows, aligns)
    set_three_line(table)
    # 表头重复 + 禁止单行跨页
    tr = table.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = trPr.makeelement(qn("w:tblHeader"), {qn("w:val"): "true"})
    trPr.append(tblHeader)
    for row in table.rows:
        rp = row._tr.get_or_add_trPr()
        rp.append(OxmlElement("w:cantSplit"))



def add_page_number(section):
    section.different_first_page_header_footer = True
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = footer.add_run("第 ")
    set_run_font(r1, size=9)
    run = footer.add_run()
    set_run_font(run, size=9)
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
    r2 = footer.add_run(" 页")
    set_run_font(r2, size=9)


def main() -> None:
    text = SRC.read_text(encoding="utf-8")
    lines = text.splitlines()

    doc = Document()
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.8)

    # 默认样式
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    add_page_number(section)

    i = 0
    first_h1 = True
    first_h2 = True
    chapter_started = False
    in_notes = False
    in_refs = False
    while i < len(lines):
        line = lines[i].rstrip()
        s = line.strip()
        if not s:
            i += 1
            continue

        # 独立公式
        if s == "$$":
            i += 1
            tex_lines = []
            while i < len(lines) and lines[i].strip() != "$$":
                tex_lines.append(lines[i].strip())
                i += 1
            i += 1
            add_omml(doc, " ".join(tex_lines))
            continue

        # 图片
        m = re.match(r"!\[[^\]]*\]\(([^)]+)\)", s)
        if m:
            rel = m.group(1)
            img = (SRC.parent / rel).resolve()
            if img.exists():
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(6)
                p.add_run().add_picture(str(img), width=Inches(6.0))
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

        # 目录
        if s == "[[TOC]]":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(12)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            p.paragraph_format.line_spacing = Pt(20)
            add_runs(p, "目  录", size=16, bold=True)
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            p.paragraph_format.line_spacing = Pt(20)
            r1 = p.add_run()
            fld = OxmlElement("w:fldChar")
            fld.set(qn("w:fldCharType"), "begin")
            r1._r.append(fld)
            r2 = p.add_run()
            instr = OxmlElement("w:instrText")
            instr.set(qn("xml:space"), "preserve")
            instr.text = ' TOC \\o "1-2" \\h \\z \\u '
            r2._r.append(instr)
            r3 = p.add_run()
            fld = OxmlElement("w:fldChar")
            fld.set(qn("w:fldCharType"), "separate")
            r3._r.append(fld)
            p.add_run("右键更新目录")
            r4 = p.add_run()
            fld = OxmlElement("w:fldChar")
            fld.set(qn("w:fldCharType"), "end")
            r4._r.append(fld)
            i += 1
            continue

        # 标题
        if s.startswith("# "):
            h = s[2:].strip()
            if first_h1:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(24)
                p.paragraph_format.space_after = Pt(12)
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                p.paragraph_format.line_spacing = Pt(20)
                add_runs(p, h, size=18, bold=True)
                first_h1 = False
            else:
                if (h.startswith("第 ") or h in ("参考文献", "参考文献：", "注释")
                        or h in ("摘要", "Abstract") or h.startswith("附录")):
                    doc.add_page_break()
                elif not chapter_started and re.match(r"^[一二三四五六七八九十]+、", h):
                    doc.add_page_break()
                heading(doc, h, 1)
                chapter_started = True
                in_notes = h == "注释"
                in_refs = h.startswith("参考文献")
            i += 1
            continue
        if s.startswith("## "):
            h = s[3:].strip()
            if first_h2 and not chapter_started:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_after = Pt(18)
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                p.paragraph_format.line_spacing = Pt(20)
                add_runs(p, h, size=14, bold=True)
                first_h2 = False
            else:
                heading(doc, h, 2)
            i += 1
            continue
        if s.startswith("### "):
            heading(doc, s[4:].strip(), 3)
            i += 1
            continue

        # 分隔线
        if s == "---":
            i += 1
            continue

        # 引用说明
        if s.startswith(">"):
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.0
            add_runs(p, s.lstrip("> ").strip(), size=9)
            i += 1
            continue

        # 日期
        if re.fullmatch(r"\*\*(?:初稿|修订稿).*\*\*", s):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(24)
            add_runs(p, s.strip("*"), size=12)
            i += 1
            continue

        # 首页要素（摘要、关键词、作者简介、中图分类号）：顶格、固定行距
        # 中英文摘要各占一页，Abstract 另起一页
        if re.match(r"^\*\*(摘要：|关键词：|作者简介：|中图分类号：|Abstract:|Key words:)", s):
            if s.startswith("**Abstract"):
                doc.add_page_break()
            front_paragraph(doc, s)
            i += 1
            continue

        # 英文篇名、作者署名、作者单位：居中
        if s.startswith("**") and s.endswith("**") and s.count("**") == 2 and not chapter_started \
                and "Stock Price Reactions" in s:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(12)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            p.paragraph_format.line_spacing = Pt(20)
            add_runs(p, s.strip("*"), size=18, bold=True)
            i += 1
            continue
        if not chapter_started and (s.startswith("【") or s.startswith("（1.")):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            p.paragraph_format.line_spacing = Pt(20)
            p.paragraph_format.space_after = Pt(0)
            add_runs(p, s, size=10.5)
            i += 1
            continue

        # 表/图题（模板格式：表1 标题 / 图1 标题，编号后留一个空格）
        # 学位论文写作“表 4-4”“表 B-10”，期刊版写作“表1”，三种都要认。
        if s.startswith("**") and s.endswith("**") and len(s) < 120 \
                and re.match(r"^\*\*[表图]\s*(?:[A-Z]\s*[-－]\s*)?\d+(?:\s*[-－]\s*\d+)?\s", s):
            caption(doc, s[2:-2])
            i += 1
            continue

        # 表注
        if s.startswith("注："):
            note(doc, s)
            i += 1
            continue

        # 参考文献（无序号，模板要求由编辑部统一添加）
        if in_refs:
            reference(doc, s)
            i += 1
            continue

        # 注释条目
        if in_notes:
            note_item(doc, s)
            i += 1
            continue

        body_paragraph(doc, s)
        i += 1

    # 打开 Word 时自动更新目录页码
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        update = OxmlElement("w:updateFields")
        update.set(qn("w:val"), "true")
        settings.append(update)

    doc.save(OUT)
    print("saved", OUT)
    if MISSING_IMAGES:
        print(f"[WARN] {len(MISSING_IMAGES)} 张图未插入（路径不存在）：", file=sys.stderr)
        for p in MISSING_IMAGES:
            print("       " + p, file=sys.stderr)


if __name__ == "__main__":
    main()
