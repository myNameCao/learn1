# -*- coding: utf-8 -*-
from docx import Document
from docx.oxml.ns import qn

doc = Document(r"C:\Users\hefei.cao\work\learn\资料分析笔记-新.docx")

NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}


def omml_to_linear(el):
    parts = []

    def walk(node):
        tag = node.tag.split("}")[-1] if "}" in node.tag else node.tag
        if tag == "t" and node.text:
            parts.append(node.text)
            return
        if tag == "f":
            nums = node.findall("./m:num", NSMAP)
            dens = node.findall("./m:den", NSMAP)
            n = "".join(omml_to_linear(x) for x in nums)
            d = "".join(omml_to_linear(x) for x in dens)
            parts.append("(%s)/(%s)" % (n, d))
            return
        if tag in ("sSup", "sSub", "sSubSup"):
            e = node.find("./m:e", NSMAP)
            sup = node.find("./m:sup", NSMAP)
            sub = node.find("./m:sub", NSMAP)
            base = omml_to_linear(e) if e is not None else ""
            if sub is not None:
                base += "_" + omml_to_linear(sub)
            if sup is not None:
                base += "^" + omml_to_linear(sup)
            parts.append(base)
            return
        if tag == "d":
            e = node.find("./m:e", NSMAP)
            parts.append("(" + (omml_to_linear(e) if e is not None else "") + ")")
            return
        for c in node:
            walk(c)

    walk(el)
    return "".join(parts)


def plain_of(el):
    texts = []
    for node in el.iter():
        if node.tag == qn("w:t") and node.text:
            texts.append(node.text)
    return "".join(texts)


print("=== PARAS WITH MATH ===")
for i, p in enumerate(doc.paragraphs):
    maths = p._p.findall(".//m:oMath", NSMAP)
    if not maths:
        continue
    plain = plain_of(p._p)
    print("[%d] plain=%r" % (i, plain[:80]))
    for j, m in enumerate(maths):
        print("    math%d: %s" % (j, omml_to_linear(m)))

print("=== TABLES WITH MATH ===")
for ti, t in enumerate(doc.tables):
    for ri, row in enumerate(t.rows):
        for ci, c in enumerate(row.cells):
            maths = c._tc.findall(".//m:oMath", NSMAP)
            if not maths:
                continue
            print(
                "T%dR%dC%d plain=%r | maths=%s"
                % (
                    ti,
                    ri,
                    ci,
                    plain_of(c._tc)[:60],
                    [omml_to_linear(m) for m in maths],
                )
            )
