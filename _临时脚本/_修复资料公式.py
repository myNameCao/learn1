# -*- coding: utf-8 -*-
"""Replace OMML in 资料分析笔记 with clear plain-text formulas."""
from __future__ import annotations

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

PATH = r"C:\Users\hefei.cao\work\learn\资料分析笔记.docx"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


def local(tag: str) -> str:
    return tag.split("}")[-1] if "}" in tag else tag


def add_text_run(p_el, text: str) -> None:
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p_el.append(r)


def set_paragraph_text(p, text: str) -> None:
    p_el = p._p
    for child in list(p_el):
        if child.tag != qn("w:pPr"):
            p_el.remove(child)
    add_text_run(p_el, text)


def set_cell_text(cell, text: str) -> None:
    tc = cell._tc
    tcPr = tc.find(qn("w:tcPr"))
    for child in list(tc):
        if child is not tcPr:
            tc.remove(child)
    p = OxmlElement("w:p")
    add_text_run(p, text)
    tc.append(p)


def has_math(el) -> bool:
    return bool(el.findall(".//%soMath" % M)) or bool(el.findall(".//%soMathPara" % M))


def node_to_text(node) -> str:
    """Recursively convert an OMML/w node to readable text."""
    if node is None:
        return ""
    tag = local(node.tag)

    if tag == "f":  # fraction
        num = node_to_text(node.find(M + "num"))
        den = node_to_text(node.find(M + "den"))
        if any(c in den for c in "+-"):
            den = f"({den})"
        return f"{num}/{den}"

    if tag == "sSub":  # subscript → qA
        return node_to_text(node.find(M + "e")) + node_to_text(node.find(M + "sub"))

    if tag == "sSup":  # superscript
        return node_to_text(node.find(M + "e")) + node_to_text(node.find(M + "sup"))

    if tag == "t":
        return node.text or ""

    if node.tag == qn("w:t"):
        return node.text or ""

    # skip properties
    if tag.endswith("Pr") or tag in ("rPr", "pPr", "tcPr"):
        return ""

    return "".join(node_to_text(ch) for ch in node)


def element_to_readable(el) -> str:
    """Convert a paragraph or cell element that may mix text + OMML."""
    parts: list[str] = []
    for child in el:
        tag = local(child.tag)
        if tag in ("pPr", "tcPr", "rPr"):
            continue
        if child.tag == qn("w:r"):
            parts.append(node_to_text(child))
        elif tag in ("oMath", "oMathPara"):
            parts.append(node_to_text(child))
        elif tag == "p":
            parts.append(element_to_readable(child))
        else:
            parts.append(node_to_text(child))
    return "".join(parts).strip()


def main() -> None:
    doc = Document(PATH)

    formulas = [
        ("名称", "公式", "怎么算"),
        ("基期", "B = A / (1+q)", "百化分 / 截位"),
        ("现期", "A = B × (1+q)", "百化分乘法"),
        ("增长量", "G = A × q / (1+q)", "n+1 法则"),
        ("增长率", "q = (现−基)/基 = 增长量/基 = 增长量/(现−增长量)", "截位直除"),
        ("比重", "部分 / 整体", "截位直除"),
        ("比重增长量", "Δ = (A/B) × (qA−qB) / (1+qA)", "详见「五 · 比重」两步锁定"),
        ("是几倍", "A / B", "「多几倍」再减 1"),
        ("平均数", "总量 / 个数", "人均、亩产、日均等"),
        ("平均数增长率", "(qA−qB) / (1+qB)", "qA总量增速、qB份数增速"),
        ("平均数增长量", "现期平均数 × (qA−qB) / (1+qA)", "和比重增量同一套路"),
        ("贡献率", "A增量 / B增量", "详见「贡献率 vs 拉动」"),
        ("拉动增长", "A增量 / B基期", "用百分点"),
    ]
    t8 = doc.tables[8]
    for ri, row_data in enumerate(formulas):
        for ci, val in enumerate(row_data):
            set_cell_text(t8.rows[ri].cells[ci], val)
    print("fixed formula table")

    t7 = doc.tables[7]
    set_cell_text(t7.rows[1].cells[1], "65² = 4225")
    set_cell_text(t7.rows[2].cells[1], "86×11 = 946")
    set_cell_text(t7.rows[3].cells[1], "15×16 = 30×8 = 240")
    set_cell_text(t7.rows[4].cells[1], "×35.8% ≈ ×(33.3%+2.5%)")
    print("fixed heart-calc examples")

    n = 0
    for i, p in enumerate(doc.paragraphs):
        if not has_math(p._p):
            continue
        readable = element_to_readable(p._p)
        if not readable:
            continue
        set_paragraph_text(p, readable)
        n += 1
        print(f"p[{i}] {readable[:90]}")
    print(f"converted {n} math paragraphs")

    for ti, table in enumerate(doc.tables):
        if ti in (7, 8):
            continue
        for row in table.rows:
            for cell in row.cells:
                if has_math(cell._tc):
                    readable = element_to_readable(cell._tc)
                    if readable:
                        set_cell_text(cell, readable)
                        print(f"table[{ti}] {readable[:60]}")

    doc.save(PATH)
    print("saved", PATH)

    doc2 = Document(PATH)
    left = len(doc2.element.body.findall(".//%soMath" % M))
    print("remaining oMath:", left)
    for ri in range(0, 7):
        print([c.text for c in doc2.tables[8].rows[ri].cells])


if __name__ == "__main__":
    main()
