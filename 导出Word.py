# -*- coding: utf-8 -*-
"""Markdown 笔记同步为 Word（表格紧凑、尽量一页内）。"""
from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

BASE = Path(__file__).resolve().parent


def set_run_font(run, size=10.5, bold=False):
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    run.font.size = Pt(size)
    run.bold = bold


def set_cell_shading(cell, hex_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_table_keep(table):
    """尽量不分页拆开表格。"""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    if tbl.tblPr is None:
        tbl.insert(0, tblPr)
    for tr in table.rows:
        trPr = tr._tr.get_or_add_trPr()
        cant = OxmlElement("w:cantSplit")
        trPr.append(cant)


def add_para(doc, text, size=10.5, bold=False, space_after=3):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    parts = re.split(r"(\*\*[^*]+\*\*|`[^`]+`)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            r = p.add_run(part[2:-2])
            set_run_font(r, size=size, bold=True)
        elif part.startswith("`") and part.endswith("`"):
            r = p.add_run(part[1:-1])
            set_run_font(r, size=max(size - 1, 8))
        else:
            r = p.add_run(part)
            set_run_font(r, size=size, bold=bold)
    return p


def flush_table(doc, rows):
    parsed = []
    for row in rows:
        cells = [c.strip() for c in row.strip().strip("|").split("|")]
        if all(re.match(r"^:?-+:?$", c.replace(" ", "")) for c in cells):
            continue
        parsed.append(cells)
    if not parsed:
        return

    cols = max(len(r) for r in parsed)
    t = doc.add_table(rows=len(parsed), cols=cols)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = True

    # 可用宽度约 A4 正文区
    usable = Cm(17.0)
    col_w = usable / cols

    for i, row in enumerate(parsed):
        for j in range(cols):
            cell = t.cell(i, j)
            cell.width = col_w
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            txt = row[j] if j < len(row) else ""
            font_size = 8 if cols >= 4 else 9
            parts = re.split(r"(\*\*[^*]+\*\*)", txt)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    r = p.add_run(part[2:-2])
                    set_run_font(r, size=font_size, bold=True)
                else:
                    r = p.add_run(part)
                    set_run_font(r, size=font_size, bold=(i == 0))
            if i == 0:
                set_cell_shading(cell, "F0F0F0")
            # 单元格内边距压紧
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcMar = OxmlElement("w:tcMar")
            for side, val in (("top", "40"), ("bottom", "40"), ("left", "60"), ("right", "60")):
                node = OxmlElement(f"w:{side}")
                node.set(qn("w:w"), val)
                node.set(qn("w:type"), "dxa")
                tcMar.append(node)
            tcPr.append(tcMar)

    set_table_keep(t)
    # 表后少留空
    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(4)


def md_to_docx(md_path: Path, docx_path: Path):
    md = md_path.read_text(encoding="utf-8")
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(1.4)
        section.bottom_margin = Cm(1.4)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)

    style = doc.styles["Normal"]
    style.font.name = "微软雅黑"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    style.font.size = Pt(10.5)
    style.paragraph_format.space_after = Pt(3)
    style.paragraph_format.line_spacing = 1.15

    in_code = False
    code_buf = []
    in_formula = False
    formula_buf = []
    lines = md.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.strip().startswith("$$"):
            if not in_formula:
                in_formula = True
                formula_buf = []
                rest = line.strip()[2:].strip()
                if rest.endswith("$$"):
                    add_para(doc, rest[:-2].strip(), size=10)
                    in_formula = False
                elif rest:
                    formula_buf.append(rest)
            else:
                add_para(doc, " ".join(formula_buf), size=10)
                in_formula = False
                formula_buf = []
            i += 1
            continue
        if in_formula:
            if line.strip().endswith("$$"):
                formula_buf.append(line.strip()[:-2])
                add_para(doc, " ".join(formula_buf), size=10)
                in_formula = False
                formula_buf = []
            else:
                formula_buf.append(line.strip())
            i += 1
            continue
        if line.startswith("```"):
            if not in_code:
                in_code = True
                code_buf = []
            else:
                in_code = False
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(3)
                r = p.add_run("\n".join(code_buf))
                set_run_font(r, size=9)
                code_buf = []
            i += 1
            continue
        if in_code:
            code_buf.append(line)
            i += 1
            continue
        if line.strip().startswith("|"):
            table_buf = [line]
            i += 1
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_buf.append(lines[i])
                i += 1
            flush_table(doc, table_buf)
            continue
        if line.startswith("# "):
            p = doc.add_heading(line[2:].strip(), level=1)
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(4)
            for r in p.runs:
                set_run_font(r, size=14, bold=True)
        elif line.startswith("## "):
            p = doc.add_heading(line[3:].strip(), level=2)
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(3)
            for r in p.runs:
                set_run_font(r, size=12, bold=True)
        elif line.startswith("### "):
            p = doc.add_heading(line[4:].strip(), level=3)
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
            for r in p.runs:
                set_run_font(r, size=11, bold=True)
        elif line.startswith("#### "):
            add_para(doc, line[5:].strip(), size=10.5, bold=True, space_after=2)
        elif line.strip() in ("", "---"):
            pass
        elif line.strip().startswith(">"):
            add_para(doc, line.strip().lstrip("> ").strip(), size=9.5, space_after=2)
        elif (
            line.strip().startswith("*")
            and line.strip().endswith("*")
            and not line.strip().startswith("**")
        ):
            add_para(doc, line.strip()[1:-1], size=9.5)
        else:
            add_para(doc, line)
        i += 1

    try:
        doc.save(str(docx_path))
        print("OK", docx_path.name)
    except PermissionError:
        alt = docx_path.with_name(docx_path.stem + "-新.docx")
        doc.save(str(alt))
        print("LOCKED", docx_path.name, "->", alt.name)


def export_one(stem: str):
    md_path = BASE / f"{stem}.md"
    if not md_path.exists():
        raise FileNotFoundError(md_path.name)
    md_to_docx(md_path, BASE / f"{stem}.docx")
    md_to_docx(md_path, BASE / f"{stem}-更新.docx")


def main():
    if len(sys.argv) > 1:
        stems = [Path(a).stem for a in sys.argv[1:]]
    else:
        stems = sorted({p.stem for p in BASE.glob("*笔记.md")})
    if not stems:
        print("未找到 *笔记.md")
        sys.exit(1)
    for stem in stems:
        export_one(stem)


if __name__ == "__main__":
    main()
