# -*- coding: utf-8 -*-
"""Markdown 笔记同步为 Word。"""
from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from lxml import etree
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = Path(__file__).resolve().parent
COMPACT = False
H2_COLOR = RGBColor(0x1F, 0x6B, 0xB5)
H4_COLOR = RGBColor(0xC0, 0x39, 0x2B)
H5_COLOR = RGBColor(0x1A, 0x52, 0x76)
UP_COLOR = RGBColor(0x1E, 0x84, 0x4A)
DOWN_COLOR = RGBColor(0xC0, 0x39, 0x2B)
EQ_COLOR = RGBColor(0x1A, 0x52, 0x76)
HEAD_FILL = {
    "一、增长": "D4E6F1",
    "二、倍数": "D5F5E3",
    "三、比重": "FDEBD0",
    "四、平均数": "E8DAEF",
}
BODY_FILL = {
    "一、增长": "EBF5FB",
    "二、倍数": "EAF7F0",
    "三、比重": "FEF9F0",
    "四、平均数": "F5EEF8",
}
FORMULA_LABELS = (
    "平均数增长量",
    "平均数增长率",
    "总量增长率",
    "基期平均数",
    "比重增长量",
    "基期是几倍",
    "基期比重",
    "基期值",
    "现期值",
    "增长量",
    "增长率",
    "是几倍",
    "多几倍",
    "整体值",
    "部分值",
    "平均数",
    "总量",
    "份数",
    "比重",
)
HL_WORDS = (
    ("百分点", H4_COLOR, True),
    ("上升", UP_COLOR, True),
    ("增加", UP_COLOR, True),
    ("下降", DOWN_COLOR, True),
    ("减少", DOWN_COLOR, True),
    ("相等", EQ_COLOR, True),
    ("qA", H2_COLOR, True),
    ("qB", H2_COLOR, True),
)


IMG_RE = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")


def nest_indent_cm(level: int) -> float:
    if level >= 5:
        return 0.75
    if level >= 4:
        return 0.45
    return 0.0


# 配图统一画幅（等比缩放后居中贴白底），Word 里宽高一致
IMG_BOX_W_CM = 12.0
IMG_BOX_H_CM = 5.2
IMG_DPI = 150


def add_md_image(doc: Document, md_path: Path, rel: str):
    import io

    from PIL import Image as PILImage

    img = (md_path.parent / rel.strip()).resolve()
    if not img.is_file():
        add_para(doc, f"[缺图] {rel}", size=9, color=H4_COLOR)
        return
    box_w = max(1, int(IMG_BOX_W_CM / 2.54 * IMG_DPI))
    box_h = max(1, int(IMG_BOX_H_CM / 2.54 * IMG_DPI))
    with PILImage.open(img) as im:
        im = im.convert("RGBA")
        scale = min(box_w / im.width, box_h / im.height)
        nw = max(1, int(im.width * scale))
        nh = max(1, int(im.height * scale))
        resized = im.resize((nw, nh), PILImage.Resampling.LANCZOS)
        canvas = PILImage.new("RGB", (box_w, box_h), (255, 255, 255))
        canvas.paste(
            resized,
            ((box_w - nw) // 2, (box_h - nh) // 2),
            resized,
        )
        buf = io.BytesIO()
        canvas.save(buf, format="PNG")
        buf.seek(0)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(8)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(buf, width=Cm(IMG_BOX_W_CM), height=Cm(IMG_BOX_H_CM))


def set_run_font(run, size=10.5, bold=False, color: RGBColor | None = None):
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def set_cell_shading(cell, hex_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_table_keep(table):
    for tr in table.rows:
        trPr = tr._tr.get_or_add_trPr()
        cant = OxmlElement("w:cantSplit")
        trPr.append(cant)


def set_repeat_header(table):
    """把首行设为跨页重复的表头行，避免表头被单独留在上一页。"""
    if not len(table.rows):
        return
    tr = table.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    hdr = OxmlElement("w:tblHeader")
    hdr.set(qn("w:val"), "true")
    trPr.append(hdr)


def keep_row_with_next(row):
    """让该行的段落与下一行保持在同一页（表头不落单）。"""
    for cell in row.cells:
        for p in cell.paragraphs:
            pPr = p._p.get_or_add_pPr()
            kn = OxmlElement("w:keepNext")
            pPr.append(kn)


M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# LaTeX 命令 -> Word 纯文本符号（只影响外观，不参与结构）
TEX_SYMBOLS = {
    r"\times": "×",
    r"\Delta": "Δ",
    r"\delta": "δ",
    r"\div": "÷",
    r"\pm": "±",
    r"\mp": "∓",
    r"\approx": "≈",
    r"\ne": "≠",
    r"\neq": "≠",
    r"\le": "≤",
    r"\leq": "≤",
    r"\ge": "≥",
    r"\geq": "≥",
    r"\rightarrow": "→",
    r"\to": "→",
    r"\Rightarrow": "⇒",
    r"\cdot": "·",
    r"\ldots": "…",
    r"\dots": "…",
    r"\cdots": "⋯",
    r"\infty": "∞",
    r"\prime": "′",
    r"\circ": "°",
    r"\degree": "°",
    r"\qquad": "    ",
    r"\quad": "  ",
    r"\left": "",
    r"\right": "",
    r"\, ": " ",
    r"\,": " ",
    r"\;": " ",
    r"\!": "",
    r"\ ": " ",
}

# 只包裹内容、内容本身按普通文本渲染的命令
TEX_TEXT_CMDS = {"text", "mathrm", "mathbf", "textbf", "textit", "mbox", "operatorname"}

# 支持分数的命令
TEX_FRAC_CMDS = {"frac", "dfrac", "tfrac", "cfrac"}


def _tex_el(tag):
    return etree.Element(f"{{{M}}}{tag}")


def _tex_math_run(text, size=10.5, bold=False):
    """生成 OMML 的普通文本 run。"""
    r = _tex_el("r")
    rpr = _tex_el("rPr")
    sty = _tex_el("sty")
    sty.set(f"{{{M}}}val", "p")
    rpr.append(sty)
    r.append(rpr)

    wrpr = etree.Element(f"{{{W}}}rPr")
    # 必须指定 Cambria Math，否则 Word 编辑器实时校验会报"数学公式错误"
    rfonts = etree.Element(f"{{{W}}}rFonts")
    rfonts.set(qn("w:ascii"), "Cambria Math")
    rfonts.set(qn("w:hAnsi"), "Cambria Math")
    rfonts.set(qn("w:eastAsia"), "Cambria Math")
    wrpr.append(rfonts)
    if bold:
        wrpr.append(etree.Element(f"{{{W}}}b"))
    sz = etree.Element(f"{{{W}}}sz")
    sz.set(f"{{{W}}}val", str(int(round(size * 2))))
    wrpr.append(sz)
    szcs = etree.Element(f"{{{W}}}szCs")
    szcs.set(f"{{{W}}}val", str(int(round(size * 2))))
    wrpr.append(szcs)
    r.append(wrpr)

    t = _tex_el("t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    return r


def _tex_group(container, latex, size=10.5, bold=False):
    """把一段 LaTeX 片段转成 OMML 节点，append 到 container。支持嵌套。"""
    i = 0
    n = len(latex)

    def read_group(start):
        """从 '{' 处读取到配对的 '}'，返回 (内容, 结束索引)。"""
        depth = 0
        j = start
        while j < n:
            c = latex[j]
            if c == "{":
                depth += 1
            elif c == "}":
                depth -= 1
                if depth == 0:
                    return latex[start + 1 : j], j + 1
            j += 1
        return latex[start + 1 :], n

    while i < n:
        c = latex[i]

        # 上下标运算符 ^ _（须在普通文本分支之前判断）
        if c in "^_":
            j = i + 1
            while j < n and latex[j] in " \t":
                j += 1
            if j < n and latex[j] == "{":
                script, j = read_group(j)
            elif j < n:
                script = latex[j]
                j += 1
            else:
                script = ""

            # 找到上一个已生成的节点作为底数
            if len(container):
                base = container[-1]
                tag = "sSup" if c == "^" else "sSub"
                node = _tex_el(tag)
                e = _tex_el("e")
                container.remove(base)
                e.append(base)
                sub = _tex_el("sup" if c == "^" else "sub")
                _tex_group(sub, script, size=size, bold=bold)
                node.append(e)
                node.append(sub)
                container.append(node)
            else:
                container.append(_tex_math_run(c, size=size, bold=bold))
            i = j
            continue

        if c not in "\\{}":
            j = i
            while j < n and latex[j] not in "\\{}^_":
                j += 1
            seg = latex[i:j]
            if seg.strip() or seg == " ":
                container.append(_tex_math_run(seg, size=size, bold=bold))
            i = j
            continue

        if c == "{":
            content, i = read_group(i)
            container.append(_tex_math_run(content, size=size, bold=bold))
            continue

        if c == "}":
            i += 1
            continue

        # 反斜杠命令
        m = re.match(r"\\([A-Za-z]+)", latex[i:])
        if m:
            cmd = m.group(1)
            name = "\\" + cmd
            j = i + m.end()

            # 分数（可嵌套）
            if cmd in TEX_FRAC_CMDS:
                while j < n and latex[j] in " \t":
                    j += 1
                if j < n and latex[j] == "[":
                    k = latex.find("]", j)
                    j = (k + 1) if k != -1 else j + 1
                while j < n and latex[j] in " \t":
                    j += 1
                num, j = read_group(j)
                while j < n and latex[j] in " \t":
                    j += 1
                den, j = read_group(j)

                frac = _tex_el("f")
                fpr = _tex_el("fPr")
                ftype = _tex_el("type")
                ftype.set(f"{{{M}}}val", "bar")
                fpr.append(ftype)
                frac.append(fpr)
                num_el = _tex_el("num")
                den_el = _tex_el("den")
                _tex_group(num_el, num, size=size, bold=bold)
                _tex_group(den_el, den, size=size, bold=bold)
                frac.append(num_el)
                frac.append(den_el)
                container.append(frac)
                i = j
                continue

            # 文本类命令
            if cmd in TEX_TEXT_CMDS:
                while j < n and latex[j] in " \t":
                    j += 1
                if j < n and latex[j] == "{":
                    content, j = read_group(j)
                    container.append(_tex_math_run(content, size=size, bold=bold))
                    i = j
                    continue
                container.append(_tex_math_run(name, size=size, bold=bold))
                i = j
                continue

            # 平方根
            if cmd == "sqrt":
                while j < n and latex[j] in " \t":
                    j += 1
                rad = _tex_el("rad")
                radpr = _tex_el("radPr")
                dh = _tex_el("degHide")
                dh.set(f"{{{M}}}val", "1")
                radpr.append(dh)
                rad.append(radpr)
                rad.append(_tex_el("deg"))
                e = _tex_el("e")
                if j < n and latex[j] == "{":
                    content, j = read_group(j)
                else:
                    content, j = latex[j], j + 1
                _tex_group(e, content, size=size, bold=bold)
                rad.append(e)
                container.append(rad)
                i = j
                continue

            # 符号替换
            if name in TEX_SYMBOLS:
                container.append(_tex_math_run(TEX_SYMBOLS[name], size=size, bold=bold))
                i = j
                continue

            # 未知命令：保留原样，避免信息丢失
            container.append(_tex_math_run(name, size=size, bold=bold))
            i = j
            continue

        # 上下标运算符 ^ _
        if c in "^_":
            j = i + 1
            while j < n and latex[j] in " \t":
                j += 1
            if j < n and latex[j] == "{":
                script, j = read_group(j)
            elif j < n:
                script = latex[j]
                j += 1
            else:
                script = ""

            # 找到上一个已生成的节点作为底数
            if len(container):
                base = container[-1]
                tag = "sSup" if c == "^" else "sSub"
                node = _tex_el(tag)
                e = _tex_el("e")
                # 把 base 移入 e
                container.remove(base)
                e.append(base)
                sub = _tex_el("sup" if c == "^" else "sub")
                _tex_group(sub, script, size=size, bold=bold)
                node.append(e)
                node.append(sub)
                container.append(node)
            else:
                container.append(_tex_math_run(c, size=size, bold=bold))
            i = j
            continue

        # 转义字符 \% \$ \& \# \{ \} \_
        esc = re.match(r"\\([%&#$_{}])", latex[i:])
        if esc:
            container.append(_tex_math_run(esc.group(1), size=size, bold=bold))
            i += esc.end()
            continue

        container.append(_tex_math_run("\\", size=size, bold=bold))
        i += 1

    return container


def append_omml_math(paragraph, latex, size=10.5, bold=False):
    """把一段 LaTeX 公式转成 Word 原生公式（OMML）并插入段落。"""
    latex = latex.strip()
    if not latex:
        return
    omath = _tex_el("oMath")
    _tex_group(omath, latex, size=size, bold=bold)
    paragraph._p.append(omath)


def extract_math_segments(text):
    """从一段文本里抽出所有数学片段。

    返回 [(kind, payload)]，kind:
      - 'math' -> LaTeX 源码（不含分隔符）
      - 'text' -> 普通文本
    支持 $$...$$、\\[...\\]、\\(...\\)、$...$
    """
    segments = []
    i = 0
    n = len(text)
    buf = []

    def flush():
        if buf:
            segments.append(("text", "".join(buf)))
            buf.clear()

    while i < n:
        if text.startswith("$$", i):
            end = text.find("$$", i + 2)
            if end != -1:
                flush()
                segments.append(("math", text[i + 2 : end].strip()))
                i = end + 2
                continue
        if text.startswith("\\[", i):
            end = text.find("\\]", i + 2)
            if end != -1:
                flush()
                segments.append(("math", text[i + 2 : end].strip()))
                i = end + 2
                continue
        if text.startswith("\\(", i):
            end = text.find("\\)", i + 2)
            if end != -1:
                flush()
                segments.append(("math", text[i + 2 : end].strip()))
                i = end + 2
                continue
        if text[i] == "$":
            end = text.find("$", i + 1)
            if end != -1:
                body = text[i + 1 : end]
                # 判定为公式：不能跨行，且要么含 LaTeX 命令，
                # 要么是无空格短片段（排除 "$100" 这类货币写法）
                if body.strip() and "\n" not in body:
                    has_cmd = "\\" in body
                    no_space = " " not in body.strip()
                    if has_cmd or no_space:
                        flush()
                        segments.append(("math", body.strip()))
                        i = end + 1
                        continue
        buf.append(text[i])
        i += 1

    flush()
    return segments


BARE_TEX_RE = re.compile(
    r"\\(?:frac|dfrac|tfrac|cfrac|sqrt|text|mathrm|mathbf|textbf|times|div|pm|mp|approx|ne|neq|le|leq|ge|geq|rightarrow|to|Rightarrow|cdot|ldots|dots|cdots|infty|prime|circ|degree|qquad|quad|left|right)\b"
)


def _split_bare_tex(text):
    """把混在普通文本里的裸 LaTeX 命令切出来，返回 [(is_math, payload)]。"""
    out = []
    pos = 0
    for m in BARE_TEX_RE.finditer(text):
        if m.start() > pos:
            out.append((False, text[pos : m.start()]))
        # 取出该命令对应的完整片段（含它后面可能跟的 {} 参数）
        chunk = m.group(0)
        # 分数：接管后续两个花括号组
        if m.group(0).lstrip("\\") in TEX_FRAC_CMDS:
            j = m.end()
            for _ in range(2):
                while j < len(text) and text[j] in " \t":
                    j += 1
                if j < len(text) and text[j] == "{":
                    depth = 0
                    k = j
                    while k < len(text):
                        if text[k] == "{":
                            depth += 1
                        elif text[k] == "}":
                            depth -= 1
                            if depth == 0:
                                break
                        k += 1
                    if k < len(text):
                        chunk = text[m.start() : k + 1]
                        j = k + 1
                    else:
                        break
        out.append((True, chunk))
        pos = m.start() + len(chunk)
    if pos < len(text):
        out.append((False, text[pos:]))
    return out


def append_hl_text(p, text, size=10.5, bold=False, color=None):
    if not text:
        return
    keys = [re.escape(w) for w, _, _ in HL_WORDS]
    parts = re.split(f"({'|'.join(keys)})", text) if keys else [text]
    cmap = {w: (c, b) for w, c, b in HL_WORDS}
    for part in parts:
        if not part:
            continue
        if part in cmap:
            c, b = cmap[part]
            r = p.add_run(part)
            set_run_font(r, size=size, bold=True if b else bold, color=c)
        else:
            r = p.add_run(part)
            set_run_font(r, size=size, bold=bold, color=color)


def append_text_with_frac(p, text, size=10.5, bold=False, color=None):
    """把可能含 LaTeX 公式的文本渲染进段落：公式转 OMML，其余走普通 run。"""
    if not text:
        return
    for kind, payload in extract_math_segments(text):
        if kind == "math":
            if payload:
                append_omml_math(p, payload, size=size, bold=bold)
        else:
            # 文本段里可能混有裸 LaTeX 命令（如 \frac{...}{...}）
            for is_math, sub in _split_bare_tex(payload):
                if not sub:
                    continue
                if is_math:
                    append_omml_math(p, sub, size=size, bold=bold)
                else:
                    append_hl_text(p, sub, size=size, bold=bold, color=color)


def _has_math(text):
    """判断一段文本里是否含 LaTeX 公式。"""
    if not text:
        return False
    if re.search(r"\\(frac|dfrac|tfrac|cfrac)\b", text):
        return True
    for kind, _ in extract_math_segments(text):
        if kind == "math":
            return True
    return False


def add_formula_para(doc, latex, size=10.5, space_after=4, space_before=2, indent_cm=0.0):
    """块级公式：渲染为 Word 原生公式（OMML），居中。"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.4
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if indent_cm:
        p.paragraph_format.left_indent = Cm(indent_cm)
    if latex.strip():
        append_omml_math(p, latex, size=size)
    return p


def add_para(doc, text, size=10.5, bold=False, space_after=2, indent_cm=0.0, space_before=0, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    frac_after = 1 if COMPACT else 8
    frac_ls = 1.12 if COMPACT else 1.6
    body_after = 0 if COMPACT else space_after
    p.paragraph_format.space_after = Pt(max(body_after, frac_after) if _has_math(text) else body_after)
    if _has_math(text):
        p.paragraph_format.line_spacing = frac_ls
    else:
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    if indent_cm:
        p.paragraph_format.left_indent = Cm(indent_cm)
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
            append_text_with_frac(p, part, size=size, bold=bold, color=color)
    return p


def set_table_indent(table, indent_cm: float):
    if not indent_cm:
        return
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblInd = OxmlElement("w:tblInd")
    tblInd.set(qn("w:w"), str(int(indent_cm * 567)))
    tblInd.set(qn("w:type"), "dxa")
    tblPr.append(tblInd)


def fill_rich_cell(cell, txt, font_size, header=False):
    chunks = [c.strip() for c in re.split(r"<br\s*/?>", txt, flags=re.I)]
    if not chunks:
        chunks = [""]
    for i, chunk in enumerate(chunks):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1 if COMPACT else 2)
        p.paragraph_format.line_spacing = 1.3 if _has_math(chunk) else 1.0
        labeled = False
        if not header:
            for lab in FORMULA_LABELS:
                if chunk.startswith(lab + " ") or chunk.startswith(lab + "　") or chunk.startswith(lab + "="):
                    r = p.add_run(lab)
                    set_run_font(r, size=font_size, bold=True, color=H4_COLOR)
                    chunk = chunk[len(lab) :]
                    labeled = True
                    break
        parts = re.split(r"(\*\*[^*]+\*\*)", chunk)
        for part in parts:
            if not part:
                continue
            if part.startswith("**") and part.endswith("**"):
                inner = part[2:-2]
                r = p.add_run(inner)
                if inner in (">",):
                    set_run_font(r, size=font_size, bold=True, color=UP_COLOR)
                elif inner in ("<",):
                    set_run_font(r, size=font_size, bold=True, color=DOWN_COLOR)
                elif inner in ("=",):
                    set_run_font(r, size=font_size, bold=True, color=EQ_COLOR)
                elif "求" in inner or "判断" in inner or "基本" in inner or inner.startswith("（"):
                    set_run_font(r, size=font_size, bold=True, color=H4_COLOR)
                else:
                    set_run_font(r, size=font_size, bold=True, color=H4_COLOR)
            else:
                append_text_with_frac(p, part, size=font_size, bold=header)


def flush_table(doc, rows, indent_cm=0.0):
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
    usable = Cm(17.4 if COMPACT else 17.0)
    col_w = usable / cols
    if COMPACT and cols == 2:
        font_size = 8.5
    elif cols >= 4:
        font_size = 8
    else:
        font_size = 9
    for i, row in enumerate(parsed):
        header = i == 0 or (
            COMPACT and cols == 2 and all(len(c) <= 8 and "\\" not in c for c in row)
        )
        for j in range(cols):
            cell = t.cell(i, j)
            cell.width = col_w
            cell.text = ""
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            txt = row[j] if j < len(row) else ""
            fill_rich_cell(cell, txt, font_size, header=header)
            col_name = parsed[0][j] if parsed and j < len(parsed[0]) else ""
            if header:
                set_cell_shading(cell, HEAD_FILL.get(col_name, "F0F0F0"))
            elif col_name in BODY_FILL:
                set_cell_shading(cell, BODY_FILL[col_name])
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcMar = OxmlElement("w:tcMar")
            pad = "40" if COMPACT else "20"
            for side, val in (("top", pad), ("bottom", pad), ("left", "50"), ("right", "50")):
                node = OxmlElement(f"w:{side}")
                node.set(qn("w:w"), val)
                node.set(qn("w:type"), "dxa")
                tcMar.append(node)
            tcPr.append(tcMar)
    set_table_keep(t)
    set_repeat_header(t)
    if len(t.rows) > 1:
        keep_row_with_next(t.rows[0])
    set_table_indent(t, indent_cm)
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(2)
    if indent_cm:
        sp.paragraph_format.left_indent = Cm(indent_cm)


def md_to_docx(md_path: Path, docx_path: Path):
    global COMPACT
    COMPACT = md_path.stem == "常用公式"
    md = md_path.read_text(encoding="utf-8")
    doc = Document()
    for section in doc.sections:
        m = 0.55 if COMPACT else 1.0
        side = 0.8 if COMPACT else 1.2
        section.top_margin = Cm(m)
        section.bottom_margin = Cm(m)
        section.left_margin = Cm(side)
        section.right_margin = Cm(side)
    style = doc.styles["Normal"]
    style.font.name = "微软雅黑"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    style.font.size = Pt(10.5)

    in_code = False
    code_buf = []
    in_formula = False
    formula_buf = []
    heading_level = 1
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
                    add_formula_para(doc, rest[:-2].strip(), size=10)
                    in_formula = False
                elif rest:
                    formula_buf.append(rest)
            else:
                add_formula_para(doc, " ".join(formula_buf), size=10)
                in_formula = False
                formula_buf = []
            i += 1
            continue
        if in_formula:
            if line.strip().endswith("$$"):
                formula_buf.append(line.strip()[:-2])
                add_formula_para(doc, " ".join(formula_buf), size=10)
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
                r = p.add_run("\n".join(code_buf))
                set_run_font(r, size=9)
                code_buf = []
            i += 1
            continue
        if in_code:
            code_buf.append(line)
            i += 1
            continue
        img_m = IMG_RE.fullmatch(line.strip())
        if img_m:
            add_md_image(doc, md_path, img_m.group(2))
            i += 1
            continue
        if line.strip().startswith("|"):
            table_buf = [line]
            i += 1
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_buf.append(lines[i])
                i += 1
            flush_table(doc, table_buf, indent_cm=nest_indent_cm(heading_level))
            continue
        if line.startswith("# "):
            heading_level = 1
            title = line[2:].strip()
            if COMPACT:
                is_cover = title == "常用公式"
                add_para(
                    doc,
                    title,
                    size=13 if is_cover else 11,
                    bold=True,
                    space_before=0 if is_cover else 5,
                    space_after=1,
                    color=None if is_cover else H2_COLOR,
                )
            else:
                p = doc.add_heading(title, level=1)
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(4)
                for r in p.runs:
                    set_run_font(r, size=16, bold=True)
        elif line.startswith("## "):
            heading_level = 2
            title = line[3:].strip()
            if COMPACT:
                add_para(
                    doc,
                    title,
                    size=10,
                    bold=True,
                    space_before=3,
                    space_after=0,
                    color=H4_COLOR,
                )
            else:
                p = doc.add_heading(title, level=2)
                p.paragraph_format.space_before = Pt(10)
                p.paragraph_format.space_after = Pt(4)
                for r in p.runs:
                    set_run_font(r, size=13, bold=True, color=H2_COLOR)
        elif line.startswith("### "):
            heading_level = 3
            p = doc.add_heading(line[4:].strip(), level=3)
            p.paragraph_format.space_before = Pt(4 if COMPACT else 8)
            p.paragraph_format.space_after = Pt(2 if COMPACT else 3)
            for r in p.runs:
                set_run_font(r, size=11 if COMPACT else 12, bold=True)
        elif line.startswith("#### "):
            heading_level = 4
            p = add_para(
                doc,
                line[5:].strip(),
                size=12,
                bold=True,
                space_after=3,
                space_before=8,
                indent_cm=0.2,
                color=H4_COLOR,
            )
        elif line.startswith("##### "):
            heading_level = 5
            p = add_para(
                doc,
                line[6:].strip(),
                size=11,
                bold=True,
                space_after=2,
                space_before=6,
                indent_cm=0.45,
                color=H5_COLOR,
            )
        elif line.strip() in ("", "---"):
            pass
        elif line.strip().startswith(">"):
            add_para(
                doc,
                line.strip().lstrip("> ").strip(),
                size=9.5,
                space_after=2,
                indent_cm=nest_indent_cm(heading_level),
            )
        elif (
            line.strip().startswith("*")
            and line.strip().endswith("*")
            and not line.strip().startswith("**")
        ):
            add_para(doc, line.strip()[1:-1], size=9.5, indent_cm=nest_indent_cm(heading_level))
        else:
            add_para(
                doc,
                line,
                size=9 if COMPACT else 10.5,
                indent_cm=nest_indent_cm(heading_level),
            )
        i += 1

    try:
        doc.save(str(docx_path))
        print("OK", docx_path.name)
    except PermissionError:
        alt = docx_path.with_name(docx_path.stem + "-新.docx")
        doc.save(str(alt))
        print("LOCKED", docx_path.name, "->", alt.name)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if args:
        names = args
    else:
        names = ["言语理解笔记"]
    for name in names:
        stem = name[:-3] if name.endswith(".md") else name
        md_path = BASE / f"{stem}.md"
        docx_path = BASE / f"{stem}.docx"
        if stem == "词语积累笔记":
            print("SKIP", md_path.name, "(use 导出Excel.py)")
            continue
        if not md_path.exists():
            print("MISSING", md_path.name)
            continue
        md_to_docx(md_path, docx_path)


if __name__ == "__main__":
    main()
