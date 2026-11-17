# -*- coding: utf-8 -*-
"""Batch: strip any encryption and add bookmarks to Offcn ebook PDFs."""
from __future__ import annotations

import os
import re
import shutil
import tempfile
from pathlib import Path

import pymupdf as fitz

ROOT = Path(r"C:\Users\hefei.cao\work\learn\中公电子书-资料")

NOISE = {
    "内部资料 免费交流",
    "内部资料  免费交流",
    "内部资料",
    "免费交流",
    "CONTENTS",
    "目",
    "录",
    "目录",
}

HEAD_RE = re.compile(
    r"^("
    r"第[一二三四五六七八九十百零〇0-9]+[组章节部分篇讲]"
    r"|第\s*\d+\s*[组章节部分篇讲]"
    r"|专题[一二三四五六七八九十0-9]+"
    r"|附录[一二三四五六七八九十0-9]*"
    r"|[一二三四五六七八九十]+[、.．]\s*\S{1,40}"
    r"|（[一二三四五六七八九十]+）\s*\S{1,40}"
    r"|第一节|第二节|第三节|第四节|第五节|第六节|第七节|第八节|第九节|第十节"
    r"|第一章|第二章|第三章|第四章|第五章|第六章|第七章|第八章|第九章|第十章"
    r"|第一篇|第二篇|第三篇|第四篇|第五篇"
    r")"
)

TOC_LINE_RE = re.compile(
    r"^(?P<title>.+?)\s*[·．.。…\-—–_]{2,}\s*(?P<page>\d+)\s*$"
)
CHAPTER_ONLY_RE = re.compile(
    r"^(第[一二三四五六七八九十百零〇0-9]+[章节篇部分讲]"
    r"|专题[一二三四五六七八九十0-9]+"
    r"|附录[一二三四五六七八九十0-9]*"
    r"|第[一二三四五六七八九十0-9]+节)$"
)


def is_noise(line: str) -> bool:
    s = line.strip()
    if not s or s in NOISE:
        return True
    if re.fullmatch(r"\d{1,3}", s):
        return True
    if re.fullmatch(r"[IVXivx]+", s):
        return True
    return False


def page_lines(doc: fitz.Document, i: int) -> list[str]:
    return [ln.strip() for ln in (doc[i].get_text("text") or "").splitlines() if ln.strip()]


def text_page_count(doc: fitz.Document) -> int:
    n = 0
    for page in doc:
        if len((page.get_text("text") or "").strip()) > 30:
            n += 1
    return n


def parse_printed_toc(doc: fitz.Document) -> list[list]:
    """Parse early-page printed TOC: title ...... pageNo (supports split chapter lines)."""
    entries: list[tuple[str, int]] = []
    for i in range(min(6, doc.page_count)):
        lines = page_lines(doc, i)
        pending = ""
        for ln in lines:
            if CHAPTER_ONLY_RE.match(ln) and len(ln) <= 12:
                pending = ln
                continue
            m = TOC_LINE_RE.match(ln)
            if not m:
                pending = ""
                continue
            title = m.group("title").strip(" .·．。…—–-_")
            if pending:
                title = f"{pending} {title}".strip()
                pending = ""
            page = int(m.group("page"))
            if is_noise(title) or page < 1:
                continue
            entries.append((title[:60], page))
    if len(entries) < 2:
        return []
    # Heuristic: printed page numbers may need offset if body page 1 is after TOC pages
    # Prefer absolute PDF page if title appears on that page; else try +offset
    toc: list[list] = []
    used = set()
    for title, p in entries:
        pdf_page = None
        for cand in (p, p + 1, p + 2, p + 3):
            if 1 <= cand <= doc.page_count and cand not in used:
                # prefer page that contains a short form of title
                short = re.sub(r"\s+", "", title)[:8]
                body = re.sub(r"\s+", "", page_lines(doc, cand - 1) and "\n".join(page_lines(doc, cand - 1)) or "")
                if short and short in body:
                    pdf_page = cand
                    break
        if pdf_page is None:
            # fallback: if p looks like in-book page, add offset=TOC pages count estimate
            offset = 0
            for i in range(min(3, doc.page_count)):
                joined = "\n".join(page_lines(doc, i))
                if "目录" in joined or "CONTENTS" in joined or TOC_LINE_RE.search(joined or ""):
                    offset = i + 1
            cand = p + offset
            if 1 <= cand <= doc.page_count:
                pdf_page = cand
            elif 1 <= p <= doc.page_count:
                pdf_page = p
        if pdf_page and pdf_page not in used:
            used.add(pdf_page)
            level = 1
            if title.startswith(("第", "专题", "附录")) and any(
                x in title[:4] for x in ("节", "（", "(")
            ):
                level = 2
            if title.startswith(("一、", "二、", "三、", "四、", "五、", "（", "(")) or title.startswith(
                ("第一节", "第二节", "第三节", "第四节", "第五节")
            ):
                level = 2
            toc.append([level, title, pdf_page])
    return fix_levels(toc)


def collect_headings(doc: fitz.Document) -> list[list]:
    toc: list[list] = []
    seen: set[tuple[int, str]] = set()
    for i in range(doc.page_count):
        joined = "\n".join(page_lines(doc, i))
        # skip printed directory pages to avoid fake headings all pointing at TOC
        if i < 3 and ("目录" in joined.replace(" ", "") or "CONTENTS" in joined.upper()):
            # only skip if this page also has dotted leaders
            if TOC_LINE_RE.search(joined) or "····" in joined or "...." in joined:
                continue
        # size-based + regex
        page = doc[i]
        d = page.get_text("dict")
        sizes = []
        for b in d.get("blocks", []):
            if b.get("type") != 0:
                continue
            for line in b.get("lines", []):
                spans = line.get("spans") or []
                if not spans:
                    continue
                text = "".join(s.get("text", "") for s in spans).strip()
                if not text or is_noise(text) or len(text) > 60:
                    continue
                size = max(float(s.get("size", 0)) for s in spans)
                sizes.append(size)
                key = (i + 1, text)
                if key in seen:
                    continue
                matched = bool(HEAD_RE.match(text))
                big = size >= 12.0
                if matched or (big and ("第" in text[:2] or "章" in text or "节" in text or "组【" in text)):
                    level = 1
                    if text.startswith(("第一节", "第二节", "第三节", "第四节", "第五节", "（", "(")) or re.match(
                        r"^[一二三四五六七八九十]+[、.．]", text
                    ):
                        level = 2
                    if "组【" in text or re.match(r"^第.+组", text):
                        level = 1
                    toc.append([level, text, i + 1])
                    seen.add(key)
        # also plain lines regex (in case dict missed)
        for ln in page_lines(doc, i):
            if is_noise(ln) or len(ln) > 60:
                continue
            if not HEAD_RE.match(ln):
                continue
            key = (i + 1, ln)
            if key in seen:
                continue
            level = 1
            if ln.startswith(("第一节", "第二节", "第三节", "第四节", "第五节", "（", "(")) or re.match(
                r"^[一二三四五六七八九十]+[、.．]", ln
            ):
                level = 2
            toc.append([level, ln, i + 1])
            seen.add(key)
    return fix_levels(toc)


def fix_levels(toc: list[list]) -> list[list]:
    if not toc:
        return toc
    out = []
    prev = 1
    for idx, (lvl, title, page) in enumerate(toc):
        lvl = max(1, min(int(lvl), 3))
        if idx == 0:
            lvl = 1
        elif lvl > prev + 1:
            lvl = prev + 1
        out.append([lvl, title, page])
        prev = lvl
    return out


def existing_toc_ok(doc: fitz.Document) -> list[list] | None:
    toc = doc.get_toc()
    if not toc or len(toc) < 2:
        return None
    # normalize
    cleaned = []
    for item in toc:
        if len(item) < 3:
            continue
        lvl, title, page = int(item[0]), str(item[1]).strip(), int(item[2])
        if not title or page < 1 or page > doc.page_count:
            continue
        cleaned.append([lvl, title[:80], page])
    return fix_levels(cleaned) if len(cleaned) >= 2 else None


def build_toc(doc: fitz.Document) -> list[list]:
    existed = existing_toc_ok(doc)
    # Prefer richer auto headings for vocab-style books with many 第N组
    auto = collect_headings(doc)
    printed = parse_printed_toc(doc)

    # choose best
    candidates = []
    if existed:
        candidates.append(("exist", existed))
    if printed and len(printed) >= 3:
        candidates.append(("printed", printed))
    if auto and len(auto) >= 2:
        candidates.append(("auto", auto))
    if not candidates:
        # image-only fallback: bookmark every 5 pages
        if doc.page_count <= 1:
            return [[1, "全文", 1]]
        toc = []
        step = 5 if doc.page_count > 20 else max(1, doc.page_count // 6 or 1)
        for p in range(1, doc.page_count + 1, step):
            toc.append([1, f"第{p}页", p])
        if toc[-1][2] != doc.page_count:
            toc.append([1, f"第{doc.page_count}页", doc.page_count])
        return toc

    # Prefer auto if it finds many group entries (e.g. 第一组/第二组 vocab books)
    if auto and sum(1 for x in auto if "组" in x[1]) >= 5:
        return auto
    # Prefer printed TOC when it looks like a real directory
    if printed and len(printed) >= 2:
        # Avoid auto that mostly scraped the TOC page itself (many titles on same early page)
        if auto:
            pages = [x[2] for x in auto]
            early = sum(1 for p in pages if p <= 3)
            if early >= max(3, len(auto) * 0.5) and len(printed) >= 2:
                return printed
        return printed
    if existed and len(existed) >= 5 and not (auto and len(auto) > len(existed) * 1.5):
        return existed
    best = max(candidates, key=lambda x: (len(x[1]), 1 if x[0] == "exist" else 0))
    return best[1]


def process_one(path: Path) -> tuple[str, int, int]:
    doc = fitz.open(path)
    if doc.needs_pass:
        if not doc.authenticate(""):
            doc.close()
            return ("need_password", 0, 0)
    old_toc_n = len(doc.get_toc() or [])
    toc = fix_levels(build_toc(doc))
    if not toc:
        toc = [[1, "全文", 1]]
    doc.set_toc(toc)
    # save without encryption via temp then replace
    fd, tmp = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)
    tmp_path = Path(tmp)
    try:
        doc.save(tmp_path, encryption=fitz.PDF_ENCRYPT_NONE, deflate=True, garbage=3, clean=True)
        doc.close()
        shutil.move(str(tmp_path), str(path))
    except Exception:
        doc.close()
        if tmp_path.exists():
            tmp_path.unlink(missing_ok=True)
        raise
    return ("ok", old_toc_n, len(toc))


def main() -> None:
    pdfs = sorted(ROOT.rglob("*.pdf"))
    print(f"found {len(pdfs)} pdfs under {ROOT}")
    ok = fail = 0
    for p in pdfs:
        rel = p.relative_to(ROOT)
        try:
            status, old_n, new_n = process_one(p)
            if status != "ok":
                print(f"FAIL {rel}: {status}")
                fail += 1
            else:
                print(f"OK {rel}: toc {old_n} -> {new_n}")
                ok += 1
        except Exception as e:
            print(f"ERR {rel}: {type(e).__name__}: {e}")
            fail += 1
    print(f"DONE ok={ok} fail={fail}")


if __name__ == "__main__":
    main()
