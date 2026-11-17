# -*- coding: utf-8 -*-
"""Fix broken OMML in 资料分析笔记-新.docx."""
from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SRC = Path(r"C:\Users\hefei.cao\work\learn\资料分析笔记-新.docx")
FINAL = Path(r"C:\Users\hefei.cao\work\learn\资料分析笔记.docx")

NSMAP = {
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}


def make_run(text: str, bold: bool = False):
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), "微软雅黑")
    fonts.set(qn("w:hAnsi"), "微软雅黑")
    fonts.set(qn("w:eastAsia"), "微软雅黑")
    rPr.append(fonts)
    if bold:
        rPr.append(OxmlElement("w:b"))
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "21")
    rPr.append(sz)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def make_p(parts: list[tuple[str, bool]]):
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "160")
    spacing.set(qn("w:line"), "384")
    spacing.set(qn("w:lineRule"), "auto")
    pPr.append(spacing)
    p.append(pPr)
    for text, bold in parts:
        p.append(make_run(text, bold))
    return p


def replace_p(old_el, new_el):
    parent = old_el.getparent()
    idx = list(parent).index(old_el)
    parent.remove(old_el)
    parent.insert(idx, new_el)


def fix_empty_dens(root) -> int:
    fixed = 0
    for f in list(root.findall(".//m:f", NSMAP)):
        den = f.find("m:den", NSMAP)
        num = f.find("m:num", NSMAP)
        den_text = "".join(den.itertext()).strip() if den is not None else ""
        num_text = "".join(num.itertext()).strip() if num is not None else ""
        if den is None or den_text != "" or not num_text:
            continue
        omath = f.getparent()
        r = make_run(num_text)
        if omath is not None and omath.tag == qn("m:oMath"):
            parent = omath.getparent()
            idx = list(parent).index(omath)
            parent.remove(omath)
            parent.insert(idx, r)
            fixed += 1
    return fixed


def main() -> None:
    doc = Document(str(SRC))
    print("empty-den fixes:", fix_empty_dens(doc.element.body))

    # Targeted rewrites for n+1 block (match unique fingerprints)
    rules = [
        (
            lambda t: t.startswith("当") and "1" in t and "n" in t and "q" in t,
            [("当 ", False), ("q ≈ 1/n", True), (" 时：", False)],
        ),
        (
            lambda t: "G≈" in t.replace(" ", "") and "n+1" in t.replace(" ", ""),
            [("G ≈ A/(n+1)", True), ("（因为 G = A × q/(1+q)）", False)],
        ),
        (
            lambda t: "一般式" in t and "百化分" in t,
            [("一般式", True), ("（百化分出来不是 n，而是 n/m 时）：", False)],
        ),
        (
            lambda t: "⇒" in t or ("q≈" in t and "m+n" in t.replace(" ", "")),
            [("q ≈ n/m  ⇒  G = A × q/(1+q) = A × n/(m+n)", True)],
        ),
        (
            lambda t: "n=1" in t.replace(" ", "") and "n+1" in t,
            [
                ("- n=1：", False),
                ("G = A × 1/(m+1)", True),
                ("，就是上面的 n+1 法则（q = 1/m）；", False),
            ],
        ),
        (
            lambda t: "n>1" in t.replace(" ", "") and "举例" in t,
            [
                ("- n>1 举例：", False),
                ("q ≈ 2/7 ≈ 28.6% → G = A × 2/9", True),
                ("；", False),
                ("q ≈ 3/8 = 37.5% → G = A × 3/11", True),
                ("。", False),
            ],
        ),
        (
            lambda t: "24.8%" in t and "25%" in t and "例" in t,
            [("例：", False), ("24.8% ≈ 25% = 1/4", True)],
        ),
        (
            lambda t: "6485" in t and "24.8%" in t,
            [("6485/(1+24.8%) × 24.8% ≈ 6485/5", True)],
        ),
        (
            lambda t: t.startswith("记住") and "变大" in t,
            [
                ("记住：", False),
                ("q", True),
                (" 变大，", False),
                ("q/(1+q)", True),
                (" 变大；变小则变小。", False),
            ],
        ),
        (
            lambda t: "1−" in t or "1-" in t and "1+q" in t and t.strip().startswith("q"),
            [("q/(1+q) = 1 − 1/(1+q)", True)],
        ),
    ]

    used = set()
    for p in list(doc.paragraphs):
        t = "".join(p._p.itertext())
        for i, (pred, parts) in enumerate(rules):
            if i in used:
                continue
            try:
                ok = pred(t)
            except Exception:
                ok = False
            if ok:
                replace_p(p._p, make_p(parts))
                used.add(i)
                print("rewrote rule", i, "->", "".join(x[0] for x in parts)[:60])
                break

    doc.save(str(SRC))

    # verify
    doc2 = Document(str(SRC))
    empty = 0
    for f in doc2.element.body.findall(".//m:f", NSMAP):
        den = f.find("m:den", NSMAP)
        den_text = "".join(den.itertext()).strip() if den is not None else ""
        if den is not None and den_text == "":
            empty += 1
    print("remaining empty dens:", empty)

    print("--- n+1 area ---")
    for i, p in enumerate(doc2.paragraphs):
        t = "".join(p._p.itertext())
        if any(
            k in t
            for k in (
                "1/n",
                "A/(n+1)",
                "n/m",
                "2/7",
                "1/(m+1)",
                "24.8%",
                "6485",
                "记住",
                "1 − 1/(1+q)",
            )
        ):
            print("[%d] %s" % (i, t))

    # core formula table spot-check
    print("--- formula table ---")
    t = doc2.tables[8]
    for ri in range(min(7, len(t.rows))):
        cells = []
        for c in t.rows[ri].cells:
            cells.append("".join(c._tc.itertext()).strip())
        print(ri, " | ".join(cells))

    try:
        shutil.copy2(SRC, FINAL)
        print("OK overwritten", FINAL)
    except PermissionError:
        print("当前正确版本：", SRC)
        print("请先关闭 Word 里打开的 资料分析笔记.docx，再说一声「覆盖」。")


if __name__ == "__main__":
    main()
