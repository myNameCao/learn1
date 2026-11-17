# -*- coding: utf-8 -*-
"""Add 基期平均数 row; polish 比重 & 增长率 formulas in 核心公式表."""
from __future__ import annotations

from copy import deepcopy
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
    for tag, val in (("sz", size), ("szCs", size)):
        el = w_el(tag)
        el.set(qn("w:val"), val)
        wrPr.append(el)
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
    n, d = m_el("num"), m_el("den")
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


def p_wrap(*children):
    p = w_el("p")
    pPr = w_el("pPr")
    spacing = w_el("spacing")
    spacing.set(qn("w:before"), "40")
    spacing.set(qn("w:after"), "40")
    spacing.set(qn("w:line"), "300")
    spacing.set(qn("w:lineRule"), "auto")
    pPr.append(spacing)
    p.append(pPr)
    for c in children:
        p.append(c)
    return p


def set_text_cell(cell, text: str):
    clear_cell(cell)
    cell._tc.append(p_wrap(text_run(text)))


def set_math_cell(cell, *nodes):
    clear_cell(cell)
    om = m_el("oMath")
    for n in nodes:
        if isinstance(n, str):
            om.append(math_run(n))
        else:
            om.append(n)
    cell._tc.append(p_wrap(om))


def clone_row(table, after_ri: int):
    """Insert a cloned row after after_ri."""
    src_tr = table.rows[after_ri]._tr
    new_tr = deepcopy(src_tr)
    src_tr.addnext(new_tr)
    # re-bind: find new row index
    return after_ri + 1


def find_row(table, prefix: str) -> int | None:
    for ri, row in enumerate(table.rows):
        name = row.cells[0].text.replace("\n", "").strip()
        if name.startswith(prefix):
            return ri
    return None


def main() -> None:
    doc = Document(str(DOC))
    t = doc.tables[8]

    def exact_row(name: str) -> int | None:
        for i, row in enumerate(t.rows):
            if row.cells[0].text.replace("\n", "").strip() == name:
                return i
        return None

    # --- polish 增长率 ---
    ri = exact_row("增长率")
    if ri is not None:
        set_math_cell(
            t.rows[ri].cells[1],
            "q = ",
            frac(["现期 − 基期"], ["基期"]),
            " = ",
            frac(["增长量"], ["基期"]),
        )
        set_text_cell(t.rows[ri].cells[2], "截位直除；也可用增长量÷基期")
        set_text_cell(
            t.rows[ri].cells[3],
            "反映相对变化快慢。增长量÷基期；也可（现−基）÷基。注意与「增长量」区分",
        )

    # --- polish 比重 ---
    ri = exact_row("比重")
    if ri is not None:
        set_math_cell(t.rows[ri].cells[1], "比重 = ", frac(["部分"], ["整体"]))
        set_text_cell(t.rows[ri].cells[2], "截位直除；很接近时用差分法")
        set_text_cell(
            t.rows[ri].cells[3],
            "部分占整体的比例。现期比重直接除；求基期比重用 现期比重×(1+q整体)/(1+q部分)",
        )

    # --- fill empty 基期说明 ---
    ri = exact_row("基期")
    if ri is not None:
        set_math_cell(t.rows[ri].cells[1], "B = ", frac(["A"], ["1 + q"]))
        set_text_cell(t.rows[ri].cells[2], "百化分 / 截位；降幅用 1−|q|")
        set_text_cell(
            t.rows[ri].cells[3],
            "已知现期求过去。基期=现期/(1+增长率)；降幅时分母改为 1−|q|",
        )

    # --- add 基期平均数 if missing ---
    if find_row(t, "基期平均数") is None:
        avg_ri = find_row(t, "平均数")
        # insert after 平均数 (not after 平均数增长率)
        for i, row in enumerate(t.rows):
            if row.cells[0].text.replace("\n", "").strip() == "平均数":
                avg_ri = i
                break
        if avg_ri is None:
            raise SystemExit("平均数 row missing")
        new_ri = clone_row(t, avg_ri)
        # after clone, table.rows updates
        row = t.rows[new_ri]
        set_text_cell(row.cells[0], "基期平均数")
        set_math_cell(
            row.cells[1],
            frac(["现期平均数"], ["1 + 平均数增长率"]),
            " = 现期平均数 × ",
            frac(["1 + q_B"], ["1 + q_A"]),
        )
        set_text_cell(row.cells[2], "先求平均数增长率，再除；或直接×(1+q份)/(1+q总)")
        set_text_cell(
            row.cells[3],
            "基期平均=现期平均/(1+平均增速)。"
            "也可用 现期平均×(1+份数增速)/(1+总量增速)。"
            "q_A=总量增速，q_B=份数增速",
        )
        print("added 基期平均数 at row", new_ri)
    else:
        print("基期平均数 already exists")

    try:
        doc.save(str(DOC))
        print("saved", DOC)
    except PermissionError:
        alt = DOC.with_name("资料分析笔记-已补基期平均.docx")
        doc.save(str(alt))
        print("占用，另存", alt)
        return

    # verify
    doc2 = Document(str(DOC))
    for ri, row in enumerate(doc2.tables[8].rows):
        name = row.cells[0].text.replace("\n", "").strip()
        if name in ("基期", "增长率", "比重", "平均数", "基期平均数") or "基期平均" in name:
            print(ri, name, "|", row.cells[3].text.strip()[:50])


if __name__ == "__main__":
    main()
