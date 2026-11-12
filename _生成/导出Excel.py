# -*- coding: utf-8 -*-
"""Markdown 同步为 Excel（词语本：成语表 / 词语表）。"""
from __future__ import annotations

import re
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

TOOL_DIR = Path(__file__).resolve().parent
ROOT = TOOL_DIR.parent
BASE = ROOT
HEAD_FILL = PatternFill("solid", fgColor="1F6BB5")
HEAD_FONT = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name="微软雅黑", size=11)
KEY_FONT = Font(name="微软雅黑", size=11, bold=True, color="C0392B")
ALT_FILL = PatternFill("solid", fgColor="EBF5FB")
THIN = Border(
    left=Side(style="thin", color="D0D7DE"),
    right=Side(style="thin", color="D0D7DE"),
    top=Side(style="thin", color="D0D7DE"),
    bottom=Side(style="thin", color="D0D7DE"),
)
WRAP = Alignment(wrap_text=True, vertical="center")

IDIOM_HEADERS = ["成语", "含义", "对象 / 侧重点", "易混", "相近成语"]
IDIOM_WIDTHS = [14, 32, 30, 36, 22]
WORD_HEADERS = ["词语", "对象 / 侧重点", "搭配词"]
WORD_WIDTHS = [12, 36, 28]

STEM_MARK_RE = re.compile(
    r"<!--\s*stem:([^|>]+?)(?:\s*\|\s*out:(docx|xlsx))?\s*-->",
    re.I,
)
SOURCE_MD = TOOL_DIR / "笔记生成源.md"
# 兼容旧标记名
STEM_ALIASES = {"词语积累笔记": "词语本"}


def strip_md(text: str) -> str:
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    return text.strip()


def parse_table(lines: list[str]) -> list[list[str]]:
    rows = []
    for line in lines:
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if all(re.match(r"^:?-+:?$", c.replace(" ", "")) for c in cells):
            continue
        rows.append([strip_md(c) for c in cells])
    return rows


def iter_md_tables(md: str) -> list[tuple[list[str], list[list[str]]]]:
    tables = []
    lines = md.splitlines()
    i = 0
    while i < len(lines):
        if lines[i].strip().startswith("|"):
            buf = [lines[i]]
            i += 1
            while i < len(lines) and lines[i].strip().startswith("|"):
                buf.append(lines[i])
                i += 1
            parsed = parse_table(buf)
            if len(parsed) >= 1:
                tables.append((parsed[0], parsed[1:]))
            continue
        i += 1
    return tables


def collect_ciyu_ben(md: str) -> tuple[list[dict], list[dict]]:
    """返回（成语行, 词语行）。"""
    idioms: list[dict] = []
    words: list[dict] = []
    seen_idiom = set()
    seen_word = set()

    for heads, rows in iter_md_tables(md):
        h0 = heads[0] if heads else ""
        # 主成语表
        if h0 == "成语" and len(heads) >= 4 and "含义" in heads[1]:
            for row in rows:
                item = {
                    IDIOM_HEADERS[j]: (row[j] if j < len(row) else "")
                    for j in range(len(IDIOM_HEADERS))
                }
                key = item["成语"]
                if not key or key in seen_idiom:
                    continue
                seen_idiom.add(key)
                idioms.append(item)
            continue
        # 常考释义三列表 → 并入成语表
        if h0 == "成语" and ("意思" in heads or "易错" in heads):
            idx_mean = heads.index("意思") if "意思" in heads else 1
            idx_err = heads.index("易错") if "易错" in heads else 2
            for row in rows:
                name = row[0] if row else ""
                if not name or name in seen_idiom:
                    continue
                seen_idiom.add(name)
                idioms.append(
                    {
                        "成语": name,
                        "含义": row[idx_mean] if len(row) > idx_mean else "",
                        "对象 / 侧重点": "",
                        "易混": row[idx_err] if len(row) > idx_err else "",
                        "相近成语": "",
                    }
                )
            continue
        # 词语表：词语 | 对象 / 侧重点 | 搭配词（兼容旧五列表）
        if h0 in ("词", "词语"):
            # 新三列 or 旧：词语|含义|对象|易混|相近
            is_new = len(heads) >= 2 and "对象" in heads[1]
            for row in rows:
                name = row[0] if row else ""
                if not name or name in seen_word:
                    continue
                seen_word.add(name)
                if is_new:
                    focus = row[1] if len(row) > 1 else ""
                    collocate = row[2] if len(row) > 2 else ""
                else:
                    # 旧表：含义 + 对象 合并到侧重点；相近→搭配词
                    meaning = row[1] if len(row) > 1 else ""
                    obj = row[2] if len(row) > 2 else ""
                    focus = "；".join(x for x in [meaning, obj] if x)
                    collocate = row[4] if len(row) > 4 else (row[3] if len(row) > 3 else "")
                words.append(
                    {
                        "词语": name,
                        "对象 / 侧重点": focus,
                        "搭配词": collocate,
                    }
                )
            continue
    return idioms, words


def write_sheet(
    ws,
    rows: list[dict],
    headers: list[str],
    widths: list[float],
    key_col: str,
    table_name: str,
):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    for i, h in enumerate(headers, 1):
        cell = ws.cell(1, i, h)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.row_dimensions[1].height = 22
    last_col = get_column_letter(len(headers))
    for r, item in enumerate(rows, 2):
        for c, key in enumerate(headers, 1):
            cell = ws.cell(r, c, item.get(key, ""))
            cell.font = KEY_FONT if key == key_col else BODY_FONT
            cell.alignment = WRAP
            cell.border = THIN
            if r % 2 == 0:
                cell.fill = ALT_FILL
        ws.row_dimensions[r].height = 36
    last = max(len(rows) + 1, 2)
    if not rows:
        for c in range(1, len(headers) + 1):
            ws.cell(2, c, "").border = THIN
        last = 2
    table = Table(displayName=table_name, ref=f"A1:{last_col}{last}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"


def split_source(md_text: str) -> dict[str, tuple[str, str]]:
    parts: dict[str, tuple[str, str]] = {}
    matches = list(STEM_MARK_RE.finditer(md_text))
    for i, m in enumerate(matches):
        stem = STEM_ALIASES.get(m.group(1).strip(), m.group(1).strip())
        kind = (m.group(2) or "docx").lower()
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(md_text)
        parts[stem] = (md_text[start:end].strip() + "\n", kind)
    return parts


def resolve_md_text(stem: str) -> str | None:
    stem = STEM_ALIASES.get(stem, stem)
    if SOURCE_MD.exists():
        parts = split_source(SOURCE_MD.read_text(encoding="utf-8"))
        if stem in parts:
            body, kind = parts[stem]
            if kind != "xlsx":
                return None
            return body
    standalone = BASE / f"{stem}.md"
    if standalone.exists():
        return standalone.read_text(encoding="utf-8")
    return None


def md_to_xlsx(xlsx_path: Path, md_text: str):
    idioms, words = collect_ciyu_ben(md_text)
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "成语"
    write_sheet(ws1, idioms, IDIOM_HEADERS, IDIOM_WIDTHS, "成语", "ChengYu")
    ws2 = wb.create_sheet("词语")
    write_sheet(ws2, words, WORD_HEADERS, WORD_WIDTHS, "词语", "CiYu")
    try:
        wb.save(str(xlsx_path))
        print("OK", xlsx_path.name, f"(成语{len(idioms)} / 词语{len(words)})")
    except PermissionError:
        alt = xlsx_path.with_name(xlsx_path.stem + "-新.xlsx")
        wb.save(str(alt))
        print("LOCKED", xlsx_path.name, "->", alt.name)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if args:
        names = [STEM_ALIASES.get(a[:-3] if a.endswith(".md") else a, a[:-3] if a.endswith(".md") else a) for a in args]
    elif SOURCE_MD.exists():
        parts = split_source(SOURCE_MD.read_text(encoding="utf-8"))
        names = [s for s, (_, k) in parts.items() if k == "xlsx"]
    else:
        names = ["词语本"]
    for stem in names:
        stem = STEM_ALIASES.get(stem, stem)
        text = resolve_md_text(stem)
        if text is None:
            print("MISSING", stem)
            continue
        md_to_xlsx(BASE / f"{stem}.xlsx", text)


if __name__ == "__main__":
    main()
