# -*- coding: utf-8 -*-
"""在「7. 混合增速」之后插入比值型公式（真分数结构）+ 整体比值规则"""
import docx
from lxml import etree
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = r'C:/Users/hefei.cao/work/learn/资料分析笔记.docx'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


def q(tag, ns=M):
    return '{%s}%s' % (ns, tag)


d = docx.Document(PATH)
anchor = None
for p in d.paragraphs:
    if p.text.strip().startswith('8. 隔年增长'):
        anchor = p
        break
assert anchor is not None, 'anchor 未找到'

FONT = '微软雅黑'


def new_para():
    p = anchor.insert_paragraph_before()
    pPr = p._p.get_or_add_pPr()
    return p


def set_spacing(p, before, after, line):
    pPr = p._p.get_or_add_pPr()
    for old in pPr.findall(qn('w:spacing')):
        pPr.remove(old)
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), str(before))
    sp.set(qn('w:after'), str(after))
    sp.set(qn('w:line'), str(line))
    sp.set(qn('w:lineRule'), 'auto')
    pPr.append(sp)


def add_text(p, text, bold=False):
    r = p.add_run()
    rPr = r._r.get_or_add_rPr()
    rf = OxmlElement('w:rFonts')
    rf.set(qn('w:ascii'), FONT)
    rf.set(qn('w:hAnsi'), FONT)
    rf.set(qn('w:eastAsia'), FONT)
    rPr.append(rf)
    b = OxmlElement('w:b')
    if not bold:
        b.set(qn('w:val'), '0')
    rPr.append(b)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '21')
    rPr.append(sz)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r._r.append(t)
    return r


def add_frac(p, num, den):
    om = etree.SubElement(p._p, q('oMath'))
    f = etree.SubElement(om, q('f'))
    etree.SubElement(f, q('fPr'))
    for tag, txt in (('num', num), ('den', den)):
        holder = etree.SubElement(f, q(tag))
        r = etree.SubElement(holder, q('r'))
        mrPr = etree.SubElement(r, q('rPr'))
        sty = etree.SubElement(mrPr, q('sty'))
        sty.set(q('val'), 'p')
        wrPr = etree.SubElement(r, q('rPr', W))
        rf = etree.SubElement(wrPr, q('rFonts', W))
        for a in ('ascii', 'hAnsi', 'eastAsia'):
            rf.set(q(a, W), 'Cambria Math')
        sz = etree.SubElement(wrPr, q('sz', W))
        sz.set(q('val', W), '21')
        szcs = etree.SubElement(wrPr, q('szCs', W))
        szcs.set(q('val', W), '21')
        t = etree.SubElement(r, q('t'))
        t.text = txt


# 1) 小标题
p1 = new_para()
set_spacing(p1, 240, 60, 240)
add_text(p1, '比值型公式（结构统一：分子 / 分母）', bold=True)

# 2) 公式行一：浓度 / 增长率
p2 = new_para()
set_spacing(p2, 40, 80, 336)
add_text(p2, '浓度 = ')
add_frac(p2, '溶质质量', '溶液质量')
add_text(p2, '　　　增长率 = ')
add_frac(p2, '增长量', '基期值')

# 3) 公式行二：平均数 / 比重
p3 = new_para()
set_spacing(p3, 40, 80, 336)
add_text(p3, '平均数 = ')
add_frac(p3, '总量', '份数')
add_text(p3, '　　　比重 = ')
add_frac(p3, '部分值', '整体值')

# 4) 整体比值规则
p4 = new_para()
set_spacing(p4, 100, 80, 300)
add_text(p4, '整体比值规则（通用）', bold=True)
add_text(p4, '：整体比值介于部分比值之间，且靠近')
add_text(p4, '分母较大', bold=True)
add_text(p4, '的部分（分母＝权重）——混合增长率、混合比重、混合平均数、浓度混合都适用。')

d.save(PATH)
print('saved ok')

# 校验
d2 = docx.Document(PATH)
got = [p.text.strip() for p in d2.paragraphs if p.text.strip()]
print('段落总数:', len(d2.paragraphs))
print('oMath 总数:', len(d2.element.body.findall('.//' + q('oMath'))))
idx = got.index([t for t in got if t.startswith('比值型公式')][0])
for t in got[idx:idx + 4]:
    print(' >', t[:90])
