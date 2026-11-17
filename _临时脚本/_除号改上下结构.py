# -*- coding: utf-8 -*-
"""Convert slash formulas in notes to stacked OMML fractions."""
from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SRC = Path(r"C:\Users\hefei.cao\work\learn\资料分析笔记-新.docx")
OUT = Path(r"C:\Users\hefei.cao\work\learn\资料分析笔记-公式上下.docx")
FINAL = Path(r"C:\Users\hefei.cao\work\learn\资料分析笔记.docx")

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


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


def frac(num_nodes, den_nodes):
    """num_nodes/den_nodes: list of omml nodes or strings."""
    f = m_el("f")
    f.append(m_el("fPr"))
    num = m_el("num")
    den = m_el("den")
    for n in num_nodes:
        num.append(math_run(n) if isinstance(n, str) else n)
    for n in den_nodes:
        den.append(math_run(n) if isinstance(n, str) else n)
    f.append(num)
    f.append(den)
    return f


def omath(*nodes):
    om = m_el("oMath")
    for n in nodes:
        if isinstance(n, str):
            om.append(math_run(n))
        else:
            om.append(n)
    return om


def para_with_parts(parts, center: bool = False):
    """parts: list of ('t', text, bold?) or ('m', omath_element)."""
    p = w_el("p")
    pPr = w_el("pPr")
    spacing = w_el("spacing")
    spacing.set(qn("w:before"), "40")
    spacing.set(qn("w:after"), "80")
    spacing.set(qn("w:line"), "336")
    spacing.set(qn("w:lineRule"), "auto")
    pPr.append(spacing)
    if center:
        jc = w_el("jc")
        jc.set(qn("w:val"), "center")
        pPr.append(jc)
    p.append(pPr)
    for part in parts:
        if part[0] == "t":
            bold = part[2] if len(part) > 2 else False
            p.append(text_run(part[1], bold=bold))
        else:
            p.append(part[1])
    return p


def replace(old, new):
    parent = old.getparent()
    idx = list(parent).index(old)
    parent.remove(old)
    parent.insert(idx, new)


def main() -> None:
    doc = Document(str(SRC))

    # Build replacement map by unique fingerprint -> new paragraph element
    # q/(1+q) = 1 - 1/(1+q)
    p_id = para_with_parts(
        [
            ("m", omath(frac(["q"], ["1+q"]), " = 1 − ", frac(["1"], ["1+q"]))),
        ],
        center=True,
    )
    # 当 q ≈ 1/n 时
    p_when = para_with_parts(
        [
            ("t", "当 "),
            ("m", omath("q ≈ ", frac(["1"], ["n"]))),
            ("t", " 时："),
        ]
    )
    # G ≈ A/(n+1)
    p_g = para_with_parts(
        [
            ("m", omath("G ≈ ", frac(["A"], ["n+1"]))),
            ("t", "（因为 "),
            ("m", omath("G = A × ", frac(["q"], ["1+q"]))),
            ("t", "）"),
        ]
    )
    # 一般式
    p_gen = para_with_parts(
        [
            ("t", "一般式", True),
            ("t", "（百化分出来不是 n，而是 "),
            ("m", omath(frac(["n"], ["m"]))),
            ("t", " 时）："),
        ]
    )
    # q ≈ n/m ⇒ G = ...
    p_qm = para_with_parts(
        [
            (
                "m",
                omath(
                    "q ≈ ",
                    frac(["n"], ["m"]),
                    "  ⇒  G = A × ",
                    frac(["q"], ["1+q"]),
                    " = A × ",
                    frac(["n"], ["m+n"]),
                ),
            )
        ],
        center=True,
    )
    # n=1
    p_n1 = para_with_parts(
        [
            ("t", "- n=1："),
            ("m", omath("G = A × ", frac(["1"], ["m+1"]))),
            ("t", "，就是上面的 n+1 法则（"),
            ("m", omath("q = ", frac(["1"], ["m"]))),
            ("t", "）；"),
        ]
    )
    # n>1 examples
    p_n2 = para_with_parts(
        [
            ("t", "- n>1 举例："),
            (
                "m",
                omath(
                    "q ≈ ",
                    frac(["2"], ["7"]),
                    " ≈ 28.6% → G = A × ",
                    frac(["2"], ["9"]),
                ),
            ),
            ("t", "；"),
            (
                "m",
                omath(
                    "q ≈ ",
                    frac(["3"], ["8"]),
                    " = 37.5% → G = A × ",
                    frac(["3"], ["11"]),
                ),
            ),
            ("t", "。"),
        ]
    )
    # 24.8% example
    p_248 = para_with_parts(
        [
            ("t", "例："),
            ("m", omath("24.8% ≈ 25% = ", frac(["1"], ["4"]))),
        ]
    )
    # 6485 example
    p_6485 = para_with_parts(
        [
            (
                "m",
                omath(
                    frac(["6485"], ["1+24.8%"]),
                    " × 24.8% ≈ ",
                    frac(["6485"], ["5"]),
                ),
            )
        ],
        center=True,
    )
    # 记住
    p_rem = para_with_parts(
        [
            ("t", "记住："),
            ("m", omath("q")),
            ("t", " 变大，"),
            ("m", omath(frac(["q"], ["1+q"]))),
            ("t", " 变大；变小则变小。"),
        ]
    )
    # 平均数
    p_avg = para_with_parts(
        [
            ("t", "平均数 = 总量 ÷ 份数。平均数增长率 = "),
            ("m", omath(frac(["q总 − q份"], ["1 + q份"]))),
            ("t", "，与比重增长率同套路；平均数增长量 = 现期平均数 × "),
            ("m", omath(frac(["q总 − q份"], ["1 + q总"]))),
            ("t", "，与比重增长量同套路。先定升降（比 q总 与 q份），再卡小于 |q总−q份|。"),
        ]
    )
    # 倍数
    p_bei = para_with_parts(
        [
            ("t", "是几倍 = "),
            ("m", omath(frac(["A"], ["B"]))),
            ("t", "；多几倍 = "),
            ("m", omath(frac(["A"], ["B"]), " − 1")),
            ("t", "。问「多几倍 / 增长了几倍」务必减 1。"),
        ]
    )
    # 隔年增量
    p_ge = para_with_parts(
        [
            ("t", "隔年增量快捷（现期已知、两年同速 q）★：2012 比 2010 的增量 =「2011 比 2010 增量」+「2012 比 2011 增量」 = "),
            ("m", omath("A × ", frac(["q"], ["1+q"]), " + A × q")),
            ("t", "。例：A=896.31、q=30.74% → 210.7 + 275.5 ≈ 486.3。选项若出现 210.7 或 275.5，多半是只算了一年增量。"),
        ]
    )
    # 例3
    p_ex3a = para_with_parts(
        [
            ("t", "例 3（两分数相减，分母都凑 100）：求 "),
            ("m", omath(frac(["9466"], ["96"]), " − ", frac(["6519"], ["101"]))),
            ("t", "。"),
        ]
    )
    p_ex3b = para_with_parts(
        [
            ("t", "左：96→100（+4，约 4%），分子同向加 "),
            ("m", omath("9466 × ", frac(["4"], ["96"]), " ≈ 394 → ", frac(["9860"], ["100"]), " = 98.6")),
            ("t", "；右：101→100（−1，约 1%），分子同向减 "),
            ("m", omath("6519 × ", frac(["1"], ["101"]), " ≈ 65 → ", frac(["6454"], ["100"]), " = 64.5")),
            ("t", "；相减 98.6−64.5=34.1（真值约 34.06）。"),
        ]
    )

    rules = [
        (lambda t: "q/(1+q) = 1" in t or "q/(1+q)=1" in t.replace(" ", ""), p_id),
        (lambda t: t.startswith("当") and "1/n" in t, p_when),
        (lambda t: "G ≈ A/(n+1)" in t or "G≈A/(n+1)" in t.replace(" ", ""), p_g),
        (lambda t: "一般式" in t and "n/m" in t, p_gen),
        (lambda t: "⇒" in t and "m+n" in t, p_qm),
        (lambda t: "n=1" in t.replace(" ", "") and "m+1" in t, p_n1),
        (lambda t: "n>1" in t.replace(" ", "") and "2/7" in t, p_n2),
        (lambda t: "24.8%" in t and "1/4" in t and "例" in t, p_248),
        (lambda t: "6485" in t and "24.8%" in t, p_6485),
        (lambda t: t.startswith("记住") and "q/(1+q)" in t, p_rem),
        (lambda t: t.startswith("平均数 =") or t.startswith("平均数="), p_avg),
        (lambda t: t.startswith("是几倍"), p_bei),
        (lambda t: "隔年增量快捷" in t, p_ge),
        (lambda t: "例 3" in t and "9466" in t, p_ex3a),
        (lambda t: "96→100" in t and "9466" in t, p_ex3b),
    ]

    used = set()
    for p in list(doc.paragraphs):
        t = "".join(x.text or "" for x in p._p.findall(".//w:t", {"w": W}))
        # also include math text for matching already-converted? skip if already has m:f
        has_frac = p._p.find(".//m:f", {"m": M}) is not None
        for i, (pred, new_p) in enumerate(rules):
            if i in used:
                continue
            if pred(t):
                # allow replace even if has frac for text-only targets
                replace(p._p, new_p)
                used.add(i)
                print("converted", i, t[:50])
                break

    # Ensure 480 example (para with 480) uses stacked — already should; rebuild if linear only
    for p in list(doc.paragraphs):
        maths = p._p.findall(".//m:oMath", {"m": M})
        joined = "".join(p._p.itertext())
        if "480" in joined and "384" in joined:
            # force rebuild clean stacked version
            new480 = para_with_parts(
                [
                    (
                        "m",
                        omath(
                            frac(["480"], ["1+", frac(["1"], ["4"])]),
                            " = 480 × ",
                            frac(["4"], ["5"]),
                            " = 384",
                        ),
                    )
                ],
                center=True,
            )
            # fix: den can't mix str and frac easily in my frac() - rebuild carefully
            # Rebuild den as nodes: math_run('1+') + frac(1,4)
            den_inner = [math_run("1+"), frac(["1"], ["4"])]
            f1 = m_el("f")
            f1.append(m_el("fPr"))
            num1 = m_el("num")
            num1.append(math_run("480"))
            den1 = m_el("den")
            for n in den_inner:
                den1.append(n)
            f1.append(num1)
            f1.append(den1)
            om = omath(f1, " = 480 × ", frac(["4"], ["5"]), " = 384")
            omp = m_el("oMathPara")
            omp.append(om)
            new480 = w_el("p")
            pPr = w_el("pPr")
            spacing = w_el("spacing")
            spacing.set(qn("w:before"), "40")
            spacing.set(qn("w:after"), "80")
            spacing.set(qn("w:line"), "336")
            spacing.set(qn("w:lineRule"), "auto")
            pPr.append(spacing)
            jc = w_el("jc")
            jc.set(qn("w:val"), "center")
            pPr.append(jc)
            new480.append(pPr)
            new480.append(omp)
            replace(p._p, new480)
            print("rebuilt 480 example as stacked fractions")
            break

    doc.save(str(OUT))
    print("saved", OUT)
    for dest in (SRC, FINAL):
        try:
            shutil.copy2(OUT, dest)
            print("overwrote", dest)
        except PermissionError:
            print("占用中，未能写入：", dest)
    print("请打开查看：", OUT)


if __name__ == "__main__":
    main()
