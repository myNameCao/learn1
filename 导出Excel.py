# -*- coding: utf-8 -*-
"""词语积累 Markdown 同步为 Excel，不用 Word。"""
from __future__ import annotations

import re
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

BASE = Path(__file__).resolve().parent
HEADERS = ["成语", "含义", "对象 / 侧重点", "易混", "相近成语"]
COL_WIDTHS = [14, 32, 30, 36, 22]
HEAD_FILL = PatternFill("solid", fgColor="1F6BB5")
HEAD_FONT = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name="微软雅黑", size=11)
IDIOM_FONT = Font(name="微软雅黑", size=11, bold=True, color="C0392B")
ALT_FILL = PatternFill("solid", fgColor="EBF5FB")
THIN = Border(
    left=Side(style="thin", color="D0D7DE"),
    right=Side(style="thin", color="D0D7DE"),
    top=Side(style="thin", color="D0D7DE"),
    bottom=Side(style="thin", color="D0D7DE"),
)
WRAP = Alignment(wrap_text=True, vertical="center")


def strip_md(text: str) -> str:
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    return text.strip()


def is_four_chars(word: str) -> bool:
    return len(re.sub(r"\s+", "", word)) == 4


def parse_table(lines: list[str]) -> list[list[str]]:
    rows = []
    for line in lines:
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if all(re.match(r"^:?-+:?$", c.replace(" ", "")) for c in cells):
            continue
        rows.append([strip_md(c) for c in cells])
    return rows


def collect_entries(md: str) -> tuple[list[dict], list[dict]]:
    four, two = [], []
    lines = md.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.strip().startswith("|"):
            buf = [line]
            i += 1
            while i < len(lines) and lines[i].strip().startswith("|"):
                buf.append(lines[i])
                i += 1
            parsed = parse_table(buf)
            if parsed:
                heads = parsed[0]
                for row in parsed[1:]:
                    mapped = {heads[j]: row[j] if j < len(row) else "" for j in range(len(heads))}
                    word = mapped.get("成语") or mapped.get("词") or (row[0] if row else "")
                    item = {
                        "成语": word,
                        "含义": mapped.get("含义", row[1] if len(row) > 1 else ""),
                        "对象 / 侧重点": mapped.get("对象 / 侧重点")
                        or mapped.get("搭配提示")
                        or mapped.get("和「视野」")
                        or (row[2] if len(row) > 2 else ""),
                        "易混": mapped.get("易混") or (row[3] if len(row) > 3 else ""),
                        "相近成语": mapped.get("相近成语") or (row[4] if len(row) > 4 else ""),
                    }
                    (four if is_four_chars(word) else two).append(item)
            continue
        i += 1
    return four, two


def write_sheet(ws, rows: list[dict], table_name: str):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    for i, h in enumerate(HEADERS, 1):
        cell = ws.cell(1, i, h)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN
        ws.column_dimensions[get_column_letter(i)].width = COL_WIDTHS[i - 1]
    ws.row_dimensions[1].height = 22
    last_col = get_column_letter(len(HEADERS))
    for r, item in enumerate(rows, 2):
        for c, key in enumerate(HEADERS, 1):
            cell = ws.cell(r, c, item.get(key, ""))
            cell.font = IDIOM_FONT if key == "成语" else BODY_FONT
            cell.alignment = WRAP
            cell.border = THIN
            if r % 2 == 0:
                cell.fill = ALT_FILL
        ws.row_dimensions[r].height = 36
    last = max(len(rows) + 1, 2)
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


STEM_MARK_RE = re.compile(
    r"<!--\s*stem:([^|>]+?)(?:\s*\|\s*out:(docx|xlsx))?\s*-->",
    re.I,
)
SOURCE_MD = BASE / "笔记生成源.md"


def split_source(md_text: str) -> dict[str, tuple[str, str]]:
    parts: dict[str, tuple[str, str]] = {}
    matches = list(STEM_MARK_RE.finditer(md_text))
    for i, m in enumerate(matches):
        stem = m.group(1).strip()
        kind = (m.group(2) or "docx").lower()
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(md_text)
        parts[stem] = (md_text[start:end].strip() + "\n", kind)
    return parts


def resolve_md_text(stem: str) -> str | None:
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
    four, two = collect_entries(md_text)
    wb = Workbook()
    ws = wb.active
    ws.title = "四字成语"
    write_sheet(ws, four, "SiZiChengYu")
    if two:
        ws2 = wb.create_sheet("旧两字词")
        write_sheet(ws2, two, "JiuLiangZi")
    try:
        wb.save(str(xlsx_path))
        print("OK", xlsx_path.name)
    except PermissionError:
        alt = xlsx_path.with_name(xlsx_path.stem + "-新.xlsx")
        wb.save(str(alt))
        print("LOCKED", xlsx_path.name, "->", alt.name)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if args:
        names = [a[:-3] if a.endswith(".md") else a for a in args]
    elif SOURCE_MD.exists():
        parts = split_source(SOURCE_MD.read_text(encoding="utf-8"))
        names = [s for s, (_, k) in parts.items() if k == "xlsx"]
    else:
        names = ["词语积累笔记"]
    for stem in names:
        text = resolve_md_text(stem)
        if text is None:
            print("MISSING", stem)
            continue
        md_to_xlsx(BASE / f"{stem}.xlsx", text)


if __name__ == "__main__":
    main()
