# -*- coding: utf-8 -*-
"""Reorganize 资料分析笔记.docx, preserving OMML formulas (no deepcopy loss)."""
from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BACKUP = Path(r"C:\Users\hefei.cao\work\learn\资料分析笔记-整理前备份.docx")
OUT = Path(r"C:\Users\hefei.cao\work\learn\资料分析笔记-新.docx")
FINAL = Path(r"C:\Users\hefei.cao\work\learn\资料分析笔记.docx")


def make_p(text: str, style_val: str | None = None):
    p = OxmlElement("w:p")
    if style_val:
        pPr = OxmlElement("w:pPr")
        ps = OxmlElement("w:pStyle")
        ps.set(qn("w:val"), style_val)
        pPr.append(ps)
        p.append(pPr)
    r = OxmlElement("w:r")
    # match east-asia friendly run if possible - keep simple
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p.append(r)
    return p


def set_heading_text(p_el, text: str):
    for child in list(p_el):
        tag = child.tag
        if tag == qn("w:r") or tag.endswith("}oMath") or tag.endswith("}oMathPara"):
            p_el.remove(child)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p_el.append(r)
    return p_el


def set_style(p_el, style_val: str):
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_el.insert(0, pPr)
    ps = pPr.find(qn("w:pStyle"))
    if ps is None:
        ps = OxmlElement("w:pStyle")
        pPr.insert(0, ps)
    ps.set(qn("w:val"), style_val)


def main() -> None:
    shutil.copy2(BACKUP, OUT)
    doc = Document(str(OUT))
    body = doc.element.body
    sect = body.find(qn("w:sectPr"))
    kids = [c for c in list(body) if c is not sect]
    # detach all content nodes (keep identity — no deepcopy)
    for c in kids:
        body.remove(c)

    def K(i: int):
        return kids[i]

    # style ids in this file: 3=H1, 4=H2, 5=H3
    set_heading_text(K(0), "资料分析")
    set_heading_text(K(1), "一、基本概念")
    set_heading_text(K(2), "1. 量与率")
    set_heading_text(K(5), "2. 基期 · 现期 · 增长")

    set_heading_text(K(8), "二、速算方法")
    set_heading_text(K(9), "1. 有效数字保留（总纲）★")
    set_heading_text(K(22), "2. 截位直除（最常用）")
    set_heading_text(K(24), "3. 百化分（必背）")
    set_heading_text(K(29), "4. 修正法 ★")
    set_heading_text(K(67), "5. 心算乘法")

    set_heading_text(K(71), "三、比较方法 ★")
    set_heading_text(K(72), "1. 加减比较（A ± B）")
    set_heading_text(K(75), "2. 除法比较（A / B）")
    set_heading_text(K(81), "3. 增长量比较（乘除混合）")

    set_heading_text(K(86), "四、核心公式一览")

    set_heading_text(K(100), "五、题型精要")
    set_heading_text(K(101), "1. 基期 / 现期 / 增长率")
    set_heading_text(K(105), "2. 增长量（含 n+1）★")
    set_heading_text(K(111), "3. 比重（两期比重差）★")
    set_heading_text(K(119), "6. 贡献率 vs 拉动")
    set_heading_text(K(126), "7. 混合增速")
    set_heading_text(K(130), "9. 特征比例 / 替代比大小")

    # 概念积累 was H3 -> promote to H2
    set_heading_text(K(122), "六、概念积累（冷门词）★")
    set_style(K(122), "4")

    set_heading_text(K(136), "七、读材料 · 易错")

    # new paragraphs
    p_41 = make_p("4.1 凑整修正（为了好除）")
    p_ex3a = make_p(
        "例 3（两分数相减，分母都凑 100）：求 9466/96 − 6519/101。"
    )
    p_ex3b = make_p(
        "左：96→100（+4，约 4%），分子同向加 9466×4/96≈394 → 9860/100=98.6；"
        "右：101→100（−1，约 1%），分子同向减 6519×1/101≈65 → 6454/100=64.5；"
        "相减 98.6−64.5=34.1（真值约 34.06）。"
    )
    p_42 = make_p("4.2 差分法（两分数很接近时，用来比大小）★")
    p_avg_h = make_p("4. 平均数", "5")
    p_avg = make_p(
        "平均数 = 总量 ÷ 份数。平均数增长率 = (q总 − q份)/(1 + q份)，与比重增长率同套路；"
        "平均数增长量 = 现期平均数 × (q总 − q份)/(1 + q总)，与比重增长量同套路。"
        "先定升降（比 q总 与 q份），再卡小于 |q总−q份|。"
    )
    p_bei_h = make_p("5. 倍数", "5")
    p_bei = make_p(
        "是几倍 = A/B；多几倍 = A/B − 1。问「多几倍 / 增长了几倍」务必减 1。"
    )
    p_mix_tip = make_p(
        "口诀：混合后的增速一定夹在两段增速之间，且靠近「基期量大」的那边。"
    )
    p_ge_h = make_p("8. 隔年增长 ★", "5")
    p_ge1 = make_p(
        "隔年增速 Q = (1+q1)(1+q2) − 1 = q1 + q2 + q1q2（不能直接把两个增速相加）。"
    )
    p_ge2 = make_p(
        "隔年增量快捷（现期已知、两年同速 q）★："
        "2012 比 2010 的增量 =「2011 比 2010 增量」+「2012 比 2011 增量」"
        " = A×q/(1+q) + A×q。"
        "例：A=896.31、q=30.74% → 210.7 + 275.5 ≈ 486.3。"
        "选项若出现 210.7 或 275.5，多半是只算了一年增量。"
    )
    p_np1 = make_p("优先 n+1 法则（高频）：")

    # skip old plain labels that we replace
    skip = {
        26,  # old "3.1 凑整修正"
        41,  # old "3.2 差分法"
        89,  # old "三、增长量 n+1" heading — content kept, heading skipped
    }

    ordered = []

    def add(*idxs_or_els):
        for x in idxs_or_els:
            if isinstance(x, int):
                if x in skip:
                    continue
                ordered.append(K(x))
            else:
                ordered.append(x)

    # 一、基本概念
    add(0, 1, 2, 3, 5, 6)

    # 二、速算
    add(8, 9, 10, 12, 13, 14, 15, 16, 17, 19, 20, 21)
    add(22, 23)
    add(24, 25, 27, 28)
    add(29, 30, p_41, 32, 33, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46)
    add(p_ex3a, p_ex3b)
    add(p_42, 48, 49, 50, 51, 52, 53, 54, 55, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66)
    add(67, 68, 70)

    # 三、比较
    add(71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85)

    # 四、公式
    add(86, 87)

    # 五、题型
    add(100, 101, 102, 103, 104)
    add(105, p_np1)
    # n+1 body was kids 90-99 (after old heading 89)
    add(90, 91, 92, 93, 94, 95, 96, 97, 98, 99)
    add(106, 107, 108, 109, 110)
    add(111, 112, 113, 114, 115, 116, 117, 118)
    add(p_avg_h, p_avg, p_bei_h, p_bei)
    add(119, 120)
    add(126, 127, p_mix_tip)
    add(p_ge_h, 128, 129, p_ge1, p_ge2)
    add(130, 131, 132, 133, 134, 135)

    # 六、概念
    add(122, 123, 124)

    # 七、易错
    add(136, 137, 138)

    # unused empties / old headings intentionally dropped:
    # 4,7,11,14,18,28,34,49,56,61,69,79,88,114,121,125,139 empties mostly
    # 4,7 empty after tables; 116-118 used; etc.

    for el in ordered:
        if sect is not None:
            sect.addprevious(el)
        else:
            body.append(el)

    doc.save(str(OUT))

    from lxml import etree

    raw = etree.tostring(Document(str(OUT)).element.body, encoding="unicode")
    print("saved", OUT)
    print("oMath count", raw.count("oMath"))
    print("tables", len(doc.tables), "paragraphs", len(doc.paragraphs))
    for p in doc.paragraphs:
        if p.style and p.style.name.startswith("Heading"):
            print(p.style.name + ":", p.text)
    try:
        shutil.copy2(OUT, FINAL)
        print("also wrote", FINAL)
    except PermissionError:
        print("NOTE: close Word then rename 资料分析笔记-新.docx -> 资料分析笔记.docx")


if __name__ == "__main__":
    main()
