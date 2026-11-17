# -*- coding: utf-8 -*-
import os
import re
import shutil
import tempfile
from pathlib import Path

import pymupdf as fitz

files = [
    Path(r"C:\Users\hefei.cao\work\learn\中公电子书-资料\【事业单位】公文写作一本通\写作范文60篇.pdf"),
    Path(r"C:\Users\hefei.cao\work\learn\中公电子书-资料\事业单位招聘考试-公文写作范文60篇\公文范文60篇.pdf"),
]
pat = re.compile(r"^[一二三四五六七八九十百零〇0-9]+[、.．].{2,30}$")


def clean(path: Path) -> None:
    doc = fitz.open(path)
    toc = []
    seen = set()
    for i in range(doc.page_count):
        for ln in (doc[i].get_text("text") or "").splitlines():
            s = ln.strip()
            if not pat.match(s):
                continue
            if "。" in s:
                continue
            key = (s, i + 1)
            if key in seen:
                continue
            seen.add(key)
            toc.append([1, s[:40], i + 1])
    out = []
    titles = set()
    for lvl, t, p in toc:
        if t in titles:
            continue
        titles.add(t)
        out.append([lvl, t, p])
    if not out:
        out = [[1, "全文", 1]]
    doc.set_toc(out)
    fd, tmp = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)
    doc.save(tmp, encryption=fitz.PDF_ENCRYPT_NONE, deflate=True, garbage=3, clean=True)
    doc.close()
    shutil.move(tmp, str(path))
    print(path.name, "toc", len(out), "sample", out[:3])


if __name__ == "__main__":
    for f in files:
        clean(f)
