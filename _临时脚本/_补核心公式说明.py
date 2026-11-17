# -*- coding: utf-8 -*-
"""Under 四、核心公式一览: add 基期/比重/平均/倍数 stacked formulas + 图二解释."""
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


def math_run(text: str, size: str = "21"):
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


def text_run(text: str, bold: bool = False, size: str = "21"):
    r = w_el("r")
    rPr = w_el("rPr")
    fonts = w_el("rFonts")
    fonts.set(qn("w:ascii"), "微软雅黑")
    fonts.set(qn("w:hAnsi"), "微软雅黑")
    fonts.set(qn("w:eastAsia"), "微软雅黑")
    rPr.append(fonts)
    if bold:
        rPr.append(w_el("b"))
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
    for x in num if isinstance(num, list) else [num]:
        n.append(math_run(x) if isinstance(x, str) else x)
    for x in den if isinstance(den, list) else [den]:
        d.append(math_run(x) if isinstance(x, str) else x)
    f.append(n)
    f.append(d)
    return f


def omath(*nodes):
    om = m_el("oMath")
    for n in nodes:
        if isinstance(n, str):
            om.append(math_run(n))
        else:
            om.append(n)
    return om


def para(parts, style_val: str | None = None):
    """parts: ('t', text, bold?) | ('m', omath)."""
    p = w_el("p")
    pPr = w_el("pPr")
    if style_val:
        ps = w_el("pStyle")
        ps.set(qn("w:val"), style_val)
        pPr.append(ps)
    spacing = w_el("spacing")
    spacing.set(qn("w:before"), "60")
    spacing.set(qn("w:after"), "80")
    spacing.set(qn("w:line"), "360")
    spacing.set(qn("w:lineRule"), "auto")
    pPr.append(spacing)
    p.append(pPr)
    for part in parts:
        if part[0] == "t":
            bold = part[2] if len(part) > 2 else False
            p.append(text_run(part[1], bold=bold))
        else:
            p.append(part[1])
    return p


def main() -> None:
    doc = Document(str(DOC))

    # Find heading 四、核心公式一览 and the following 五、题型精要
    h4 = None
    h5 = None
    for p in doc.paragraphs:
        if p.text.strip() == "四、核心公式一览":
            h4 = p._p
        if p.text.strip() == "五、题型精要":
            h5 = p._p
            break
    if h4 is None or h5 is None:
        raise SystemExit("cannot find section headings")

    # Avoid duplicate insert
    body_text = "".join(doc.element.body.itertext())
    if "图二解释" in body_text and "总量增长率 − 份数增长率" in body_text:
        print("already inserted, skip")
        return

    blocks = []
    blocks.append(para([("t", "常用四式（务必记牢）★", True)]))

    # 基期
    blocks.append(
        para(
            [
                ("t", "① 基期：", True),
                ("m", omath("基期 = ", frac(["现期"], ["1 + 增长率"]))),
                ("t", "；已知现期 A、增速 q，则 "),
                ("m", omath("B = ", frac(["A"], ["1+q"]))),
                ("t", "。降幅时分母改为 1−|q|。考场优先百化分 / 截位直除。"),
            ]
        )
    )

    # 比重
    blocks.append(
        para(
            [
                ("t", "② 比重：", True),
                ("m", omath("比重 = ", frac(["部分"], ["整体"]))),
                ("t", "。两期比重差（升降看部分增速 vs 整体增速）："),
                (
                    "m",
                    omath(
                        "Δ = ",
                        frac(["部分现期"], ["整体现期"]),
                        " × ",
                        frac(["q部分 − q整体"], ["1 + q部分"]),
                    ),
                ),
                ("t", "（单位：百分点），且 |Δ| 一定小于 |q部分−q整体|。"),
            ]
        )
    )

    # 平均
    blocks.append(
        para(
            [
                ("t", "③ 平均：", True),
                ("m", omath("平均数 = ", frac(["总量"], ["份数"]))),
                ("t", "（人均、亩产、日均等）。"),
            ]
        )
    )
    blocks.append(
        para(
            [
                ("t", "平均数增长率（图二）★：", True),
                (
                    "m",
                    omath(
                        frac(["总量增长率 − 份数增长率"], ["1 + 份数增长率"])
                    ),
                ),
                ("t", "。"),
            ]
        )
    )
    blocks.append(
        para(
            [
                ("t", "图二解释：", True),
                (
                    "t",
                    "平均数 = 总量÷份数，所以它的增速不是两个增速相减那么简单，"
                    "而要像「部分÷整体」的比重增长率一样处理——"
                    "分子用（总量增速 − 份数增速），分母用（1 + 份数增速）。"
                    "口诀：升降只看「总量增速 ? 份数增速」；算大小时分母记得加「1+份数增速」，"
                    "结果一定小于 |总量增速−份数增速|。"
                    "符号上常记 q总=q_A、q份=q_B，即 ",
                ),
                ("m", omath(frac(["q_A − q_B"], ["1 + q_B"]))),
                ("t", "。"),
            ]
        )
    )
    blocks.append(
        para(
            [
                ("t", "平均数增长量：", True),
                (
                    "m",
                    omath(
                        "现期平均数 × ",
                        frac(["q_A − q_B"], ["1 + q_A"]),
                    ),
                ),
                ("t", "，套路与比重增长量相同（分母是 1+总量增速）。"),
            ]
        )
    )

    # 倍数
    blocks.append(
        para(
            [
                ("t", "④ 倍数：", True),
                ("m", omath("是几倍 = ", frac(["A"], ["B"]))),
                ("t", "；"),
                ("m", omath("多几倍 = ", frac(["A"], ["B"]), " − 1")),
                ("t", "。问「多几倍 / 增长了几倍」务必减 1，别和「是几倍」搞混。"),
            ]
        )
    )

    # insert before 五、题型精要
    for el in blocks:
        h5.addprevious(el)

    try:
        doc.save(str(DOC))
        print("saved", DOC)
    except PermissionError:
        alt = DOC.with_name("资料分析笔记-已补公式.docx")
        doc.save(str(alt))
        print("主文件占用，已另存：", alt)

    # show inserted
    doc2 = Document(str(DOC) if DOC.exists() else str(alt))
    for p in doc2.paragraphs:
        t = "".join(x.text or "" for x in p._p.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"))
        if any(k in t for k in ("常用四式", "图二解释", "① 基期", "② 比重", "③ 平均", "④ 倍数", "平均数增长率（图二）")):
            print("-", t[:100])


if __name__ == "__main__":
    main()
