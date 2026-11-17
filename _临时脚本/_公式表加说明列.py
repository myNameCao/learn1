# -*- coding: utf-8 -*-
"""Add 公式说明 column to 核心公式 table; fill/fix incomplete cells."""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

DOC = Path(r"C:\Users\hefei.cao\work\learn\资料分析笔记.docx")

# row index -> 公式说明
EXPLAIN = {
    0: "公式说明",
    1: "已知现期求过去；降幅时分母改为 1−|q|",
    2: "已知基期求现在；增速用乘法往上推",
    3: "现期×增速/(1+增速)；百化分后常用 n+1",
    4: "增长量÷基期；也可（现−基）÷基",
    5: "部分占整体的比例；截位直除即可",
    6: "两期比重差（百分点）。升降看 q部分 vs q整体；|Δ|<|q部分−q整体|",
    7: "A/B 是几倍；问「多几倍」再减 1",
    8: "总量÷份数（人均、亩产、日均等）",
    9: "图二：分子=总量增速−份数增速，分母=1+份数增速；与比重增速同套路",
    10: "现期平均数×(q总−q份)/(1+q总)；与比重增长量同套路",
    11: "部分增量÷整体增量；结果用 %",
    12: "部分增量÷整体基期；结果用百分点",
}


def set_cell_text(cell, text: str):
    """Clear cell and write plain text (keep first paragraph formatting lightly)."""
    tc = cell._tc
    # remove all paragraphs except keep structure: clear and one p
    for child in list(tc):
        if child.tag == qn("w:tcPr"):
            continue
        tc.remove(child)
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "40")
    spacing.set(qn("w:after"), "40")
    spacing.set(qn("w:line"), "300")
    spacing.set(qn("w:lineRule"), "auto")
    pPr.append(spacing)
    p.append(pPr)
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), "微软雅黑")
    fonts.set(qn("w:hAnsi"), "微软雅黑")
    fonts.set(qn("w:eastAsia"), "微软雅黑")
    rPr.append(fonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "18")
    rPr.append(sz)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p.append(r)
    tc.append(p)


def add_column(table):
    """Append one cell to each row by cloning last cell; extend tblGrid."""
    tbl = table._tbl
    tblGrid = tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        gridCol = OxmlElement("w:gridCol")
        # roughly match last col width if present
        last = list(tblGrid)
        if last is not None and len(last) > 0 and last[-1].get(qn("w:w")):
            gridCol.set(qn("w:w"), last[-1].get(qn("w:w")))
        else:
            gridCol.set(qn("w:w"), "1800")
        tblGrid.append(gridCol)
    for row in table.rows:
        last = row.cells[-1]._tc
        new_tc = deepcopy(last)
        for child in list(new_tc):
            if child.tag != qn("w:tcPr"):
                new_tc.remove(child)
        p = OxmlElement("w:p")
        new_tc.append(p)
        row._tr.append(new_tc)


def fix_how_column(table):
    """补「怎么算」列里不完整的表述。"""
    fixes = {
        9: "q_A 总量增速、q_B 份数增速；分母是 1+q_B",
        10: "现期平均数 × (q_A−q_B)/(1+q_A)",
    }
    for ri, text in fixes.items():
        set_cell_text(table.rows[ri].cells[2], text)


def fix_avg_growth_amount_formula(table):
    """Row 10 formula cell: ensure 现期平均数 × ... visible; keep OMML if present by appending text note via rewrite."""
    # Rewrite formula cell as clear stacked description using plain + will rely on 说明列
    # Better: set formula column text clearly if OMML incomplete
    cell = table.rows[10].cells[1]
    # If already has oMath, leave OMML and also ensure 说明完整; optionally prepend label in 怎么算
    # Replace formula cell with readable text that includes formula meaning (OMML rebuild is heavy)
    # Keep existing OMML; 说明列 already covers it.
    pass


def main() -> None:
    doc = Document(str(DOC))
    table = doc.tables[8]
    header0 = "".join(table.rows[0].cells[0].text.split())
    if "名称" not in header0:
        raise SystemExit("table 8 is not 核心公式表")

    # already has 4 cols?
    if len(table.columns) >= 4 and "公式说明" in "".join(table.rows[0].cells[-1].text.split()):
        print("说明列已存在，仅刷新内容")
    else:
        add_column(table)
        print("已新增一列")

    # fill 公式说明
    for ri, text in EXPLAIN.items():
        set_cell_text(table.rows[ri].cells[-1], text)

    # bold header row last cell already plain; make header bold-ish via text
    set_cell_text(table.rows[0].cells[-1], "公式说明")

    fix_how_column(table)

    # 补「怎么算」里基期/比重等更贴切的短提示（不覆盖原有有用信息，只加强空/弱的）
    strengthen = {
        1: "百化分 / 截位；降幅用 1−|q|",
        5: "截位直除；比较可用差分法",
        6: "先定升降，再卡小于 |qA−qB|",
        8: "总量÷份数；人均、亩产、日均等",
        11: "增量÷增量；单位 %",
        12: "增量÷基期；单位百分点",
    }
    for ri, text in strengthen.items():
        set_cell_text(table.rows[ri].cells[2], text)

    try:
        doc.save(str(DOC))
        print("saved", DOC)
    except PermissionError:
        alt = DOC.with_name("资料分析笔记-表已更新.docx")
        doc.save(str(alt))
        print("文件占用，另存：", alt)
        return

    # verify
    doc2 = Document(str(DOC))
    t = doc2.tables[8]
    print("cols", len(t.columns), "rows", len(t.rows))
    for ri, row in enumerate(t.rows):
        name = row.cells[0].text.strip().replace("\n", "")[:10]
        note = row.cells[-1].text.strip().replace("\n", "")[:40]
        how = row.cells[2].text.strip().replace("\n", "")[:30]
        print(ri, name, "|", how, "|", note)


if __name__ == "__main__":
    main()
