# -*- coding: utf-8 -*-
"""Update 平均数增长量 row: stacked formula from 图二 + explanation."""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

DOC = Path(r"C:\Users\hefei.cao\work\learn\资料分析笔记.docx")


def m_el(tag: str):
    return OxmlElement(f"m:{tag}")


def w_el(tag: str):
    return OxmlElement(f"w:{tag}")


def math_run(text: str, size: str = "18"):
    r = m_el("r")
    rPr = m_el("rPr")
    sty = m_el("sty")
    sty.set(qn("m:val"), "p")
    rPr.append(sty)
    r.append(rPr)
    wrPr = w_el("rPr")
    fonts = w_el("rFonts")
    fonts.set(qn("w:ascii"), "Cambria Math")
    fonts.set(qn("w:hAnsi"), "Cambria Math")
    fonts.set(qn("w:eastAsia"), "Cambria Math")
    wrPr.append(fonts)
    sz = w_el("sz")
    sz.set(qn("w:val"), size)
    wrPr.append(sz)
    szCs = w_el("szCs")
    szCs.set(qn("w:val"), size)
    wrPr.append(szCs)
    r.append(wrPr)
    t = m_el("t")
    if text.startswith(" ") or text.endswith(" "):
        t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def text_run(text: str, size: str = "18"):
    r = w_el("r")
    rPr = w_el("rPr")
    fonts = w_el("rFonts")
    fonts.set(qn("w:ascii"), "微软雅黑")
    fonts.set(qn("w:hAnsi"), "微软雅黑")
    fonts.set(qn("w:eastAsia"), "微软雅黑")
    rPr.append(fonts)
    sz = w_el("sz")
    sz.set(qn("w:val"), size)
    rPr.append(sz)
    r.append(rPr)
    t = w_el("t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def frac(num, den):
    f = m_el("f")
    f.append(m_el("fPr"))
    n = m_el("num")
    d = m_el("den")
    for x in (num if isinstance(num, list) else [num]):
        n.append(math_run(x) if isinstance(x, str) else x)
    for x in (den if isinstance(den, list) else [den]):
        d.append(math_run(x) if isinstance(x, str) else x)
    f.append(n)
    f.append(d)
    return f


def clear_cell(cell):
    tc = cell._tc
    for child in list(tc):
        if child.tag != qn("w:tcPr"):
            tc.remove(child)


def set_formula_cell(cell):
    """图二：总量/份数 × (总量增长率−份数增长率)/(1+总量增长率)"""
    clear_cell(cell)
    p = w_el("p")
    pPr = w_el("pPr")
    spacing = w_el("spacing")
    spacing.set(qn("w:before"), "40")
    spacing.set(qn("w:after"), "40")
    spacing.set(qn("w:line"), "300")
    spacing.set(qn("w:lineRule"), "auto")
    pPr.append(spacing)
    p.append(pPr)
    om = m_el("oMath")
    om.append(frac(["总量"], ["份数"]))
    om.append(math_run(" × "))
    om.append(frac(["总量增长率 − 份数增长率"], ["1 + 总量增长率"]))
    p.append(om)
    cell._tc.append(p)


def set_how_cell(cell):
    clear_cell(cell)
    p = w_el("p")
    pPr = w_el("pPr")
    spacing = w_el("spacing")
    spacing.set(qn("w:before"), "40")
    spacing.set(qn("w:after"), "40")
    spacing.set(qn("w:line"), "300")
    spacing.set(qn("w:lineRule"), "auto")
    pPr.append(spacing)
    p.append(pPr)
    # 现期平均数 × (qA−qB)/(1+qA)
    om = m_el("oMath")
    om.append(math_run("现期平均数 × "))
    om.append(frac(["q_A − q_B"], ["1 + q_A"]))
    p.append(om)
    cell._tc.append(p)


def set_note_cell(cell):
    clear_cell(cell)
    p = w_el("p")
    pPr = w_el("pPr")
    spacing = w_el("spacing")
    spacing.set(qn("w:before"), "40")
    spacing.set(qn("w:after"), "40")
    spacing.set(qn("w:line"), "300")
    spacing.set(qn("w:lineRule"), "auto")
    pPr.append(spacing)
    p.append(pPr)
    p.append(
        text_run(
            "图二解释：现期平均数就是「总量÷份数」。"
            "增长量 = 现期平均数 × (总量增速−份数增速)/(1+总量增速)；"
            "与比重增长量同套路（分母用 1+总量增速）。"
            "符号：q_A=总量增速，q_B=份数增速。"
        )
    )
    cell._tc.append(p)


def set_rate_note(cell):
    """Also clarify 平均数增长率 公式说明 briefly mention 对比图二增长量分母不同."""
    clear_cell(cell)
    p = w_el("p")
    pPr = w_el("pPr")
    spacing = w_el("spacing")
    spacing.set(qn("w:before"), "40")
    spacing.set(qn("w:after"), "40")
    spacing.set(qn("w:line"), "300")
    spacing.set(qn("w:lineRule"), "auto")
    pPr.append(spacing)
    p.append(pPr)
    p.append(
        text_run(
            "分子=总量增速−份数增速，分母=1+份数增速；"
            "与比重增速同套路。注意：增长量分母是 1+总量增速（见图二增长量行）。"
        )
    )
    cell._tc.append(p)


def main() -> None:
    doc = Document(str(DOC))
    t = doc.tables[8]
    # find 平均数增长量 row
    target = None
    rate_row = None
    for ri, row in enumerate(t.rows):
        name = row.cells[0].text.replace("\n", "").strip()
        if name.startswith("平均数增长量"):
            target = ri
        if name.startswith("平均数增长率"):
            rate_row = ri
    if target is None:
        raise SystemExit("row not found")

    set_formula_cell(t.rows[target].cells[1])
    set_how_cell(t.rows[target].cells[2])
    set_note_cell(t.rows[target].cells[3])
    if rate_row is not None and len(t.rows[rate_row].cells) >= 4:
        set_rate_note(t.rows[rate_row].cells[3])

    try:
        doc.save(str(DOC))
        print("saved", DOC, "row", target)
    except PermissionError:
        alt = DOC.with_name("资料分析笔记-已加图二.docx")
        doc.save(str(alt))
        print("占用中，另存", alt)


if __name__ == "__main__":
    main()
