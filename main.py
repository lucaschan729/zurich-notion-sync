import datetime
import hashlib
import io
import json
import math
import os
import re
import sys
import time
import unicodedata
from urllib.parse import urljoin

import pdfplumber
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties

PAGE = "https://www.zurich.com.hk/zh-hk/customer-services/life-policy-services/zlihk"
HEADING = "分紅終身人壽保險之實現率"
NA = {"不適用"}
YEARS = [str(i) for i in range(1, 11)] + ["10+"]
EXPECT_2025 = {"tcv": (5, 4), "fr": (5, 4)}     # 探查所得：資料行、數值個數
SECTION_NAME = {"tcv": "總現金價值比率", "fr": "分紅實現率"}
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
PDF_URL = re.compile(r"https?://[^\s\"'<>\\]+?\.pdf", re.I)
PCT = re.compile(r"^(\d+(?:\.\d+)?)[%％]$")
TITLE = re.compile(r"^(\d{4})年報告年度(?:之|的)?(.+)$")
YEAR_COL = re.compile(r"^保單年度(\d+\+?)$")
ISSUER = re.compile(r"^由(.+?)[*＊]?繕發的人壽保單")
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
API = "https://api.notion.com/v1"
VERSION = "2025-09-03"
AUTO = "自動生成"
FORMAT = "v1"            # 改了 Excel 格式時改這個值，下次運行就會重新上傳
HINTS = {
    401: "NOTION_TOKEN 無效：請重新貼上「AIA 自動更新」的 Internal Integration Secret",
    404: "找不到資料庫：請確認資料庫已連接「AIA 自動更新」，以及 NOTION_DATABASE_ID 正確",
}
N_FIXED = 4              # 產品名稱、貨幣、類型、繕發公司
HEAD_ROW = 3
FIRST = 4
ROW_H = 24
SEP_H = 8
YEAR_W = 13
LOW_LIMIT = 0.8
HIGH_LIMIT = 1.0
NAVY_HEX = "1F4E79"
NAME_COLOR = "FF1F2933"
CUR_COLOR = "FF7030A0"
SEP_COLOR = "FFE9ECEF"


class Stop(Exception):
    pass


# ---------- 讀取官方文件 ----------
def squash(s):
    return re.sub(r"\s+", "", s or "")


def join_lines(s):
    """把儲存格內分行的文字接回一行；只有英數字之間才補空格。"""
    out = ""
    for p in [x.strip() for x in (s or "").split("\n") if x.strip()]:
        if out and out[-1].isascii() and out[-1].isalnum() and p[0].isascii() and p[0].isalnum():
            out += " "
        out += p
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", out)).strip()


def get(url, what, pdf=False):
    err = ""
    for _ in range(3):
        try:
            r = requests.get(url, headers=UA, timeout=60)
            if r.status_code == 200 and (not pdf or r.content.startswith(b"%PDF")):
                return r
            err = f"HTTP {r.status_code}，{len(r.content)} bytes"
        except requests.RequestException as e:
            err = type(e).__name__
        time.sleep(5)
    raise Stop(f"讀取{what}失敗：{url}（{err}）")


def find_index_pdf(raw):
    for m in re.finditer(HEADING, raw):
        hit = PDF_URL.search(raw, m.end(), m.end() + 2000)
        if hit:
            return hit.group(0)
    for m in PDF_URL.finditer(raw):
        if "fulfil" in m.group(0).lower():
            return m.group(0)
    raise Stop("官方頁找不到「分紅終身人壽保險之實現率」的文件連結")


def read_index(content, base):
    """索引文件：每間繕發公司一行，旁邊有一個連結。"""
    with pdfplumber.open(io.BytesIO(content)) as doc:
        for no, page in enumerate(doc.pages, 1):
            issuers = []
            for ln in page.extract_text_lines():
                m = ISSUER.match(squash(ln["text"]))
                if m:
                    issuers.append((ln["top"], m.group(1)))
            if not issuers:
                continue
            links = []
            for a in page.annots or []:
                uri = a.get("uri")
                if isinstance(uri, bytes):
                    uri = uri.decode("utf-8", errors="replace")
                if uri and uri.strip():
                    full = urljoin(base, uri.strip())
                    if full not in [u for _, u in links]:
                        links.append((a.get("top") or 0, full))
            if len(links) != len(issuers):
                raise Stop(f"索引文件第 {no} 頁：繕發公司 {len(issuers)} 行，但連結 {len(links)} 個")
            issuers.sort()
            links.sort()
            return [(name, url) for (_, name), (_, url) in zip(issuers, links)]
    raise Stop("索引文件找不到「由……繕發的人壽保單」的繁體頁")


def load_sources():
    r = get(PAGE, "官方頁")
    raw = r.content.decode("utf-8", errors="replace").replace("\\u002F", "/").replace("\\/", "/")
    index = find_index_pdf(raw)
    r = get(index, "索引文件", pdf=True)
    print(f"索引文件：{index}（last-modified {r.headers.get('last-modified')}）")
    docs = []
    for issuer, url in read_index(r.content, index):
        d = get(url, "披露文件", pdf=True)
        modified = d.headers.get("last-modified") or ""
        print(f"披露文件：{issuer}｜{len(d.content) // 1024} KB｜last-modified {modified}\n   {url}")
        docs.append({"issuer": issuer, "url": url, "content": d.content, "modified": modified})
    return index, docs


def classify(rest):
    if "總現金價值" in rest:
        return "tcv", "總現金價值"
    if "分紅實現率" in rest:
        if "終期" in rest:
            return "fr", "終期紅利"
        if "周年" in rest or "週年" in rest:
            return "fr", "周年紅利"
    return None


def header_cols(cells, where):
    cols, names = [], []
    for i, c in enumerate(cells[2:], 2):
        s = squash(c)
        if not s:
            continue
        m = YEAR_COL.match(s)
        if not m:
            raise Stop(f"{where}：看不懂年度欄「{s}」")
        cols.append(i)
        names.append(m.group(1))
    if names != YEARS:
        raise Stop(f"{where}：年度欄不是 1 至 10、10+（讀到 {names}）")
    return cols


def parse_doc(doc, issuer):
    rows, year = [], None
    pct_text = na_text = 0
    for no, page in enumerate(doc.pages, 1):
        where = f"{issuer} 第 {no} 頁"
        titles = []
        for ln in page.extract_text_lines():
            m = TITLE.match(squash(ln["text"]))
            if m:
                kind = classify(m.group(2))
                if kind is None:
                    raise Stop(f"{where}：看不懂標題「{squash(ln['text'])}」")
                titles.append((ln["top"], int(m.group(1)), kind))
        if not titles:          # 英文頁和簡體頁沒有繁體標題，略過
            continue
        text = page.extract_text() or ""
        pct_text += len(re.findall(r"\d+(?:\.\d+)?\s*[%％]", text))
        na_text += text.count("不適用")
        prev_head = -1.0
        for table in sorted(page.find_tables(), key=lambda t: t.bbox[1]):
            cols, cur, last = None, None, None
            for row_obj, raw in zip(table.rows, table.extract()):
                cells = [c or "" for c in raw]
                top = row_obj.bbox[1]
                if len(cells) > 2 and squash(cells[0]) == "產品" and squash(cells[1]) == "貨幣":
                    cols = header_cols(cells, where)
                    above = [t for t in titles if prev_head < t[0] < top + 1]
                    if not above:
                        raise Stop(f"{where}：表格上方找不到「XXXX年報告年度之……」標題")
                    _, y, cur = max(above)
                    if year is None:
                        year = y
                    elif y != year:
                        raise Stop(f"{where}：報告年度 {y} 與前面的 {year} 不同")
                    prev_head, last = top, None
                    continue
                if cols is None or not any(squash(c) for c in cells):
                    continue
                if len(cells) <= max(cols):
                    raise Stop(f"{where}：有一行的欄數不符")
                ccy = squash(cells[1])
                vals = [squash(cells[i]) for i in cols]
                extra = [squash(c) for c in cells[2:] if squash(c)]
                if not ccy:
                    if any(PCT.match(v) or v in NA for v in extra):
                        raise Stop(f"{where}：有一行數值缺少貨幣")
                    cols = None         # 表格內混入的說明文字，數據部分到此為止
                    continue
                note = ""
                if not all(PCT.match(v) or v in NA for v in vals):
                    if len(extra) == 1 and not PCT.match(extra[0]) and extra[0] not in NA:
                        note = join_lines(next(c for c in cells[2:] if squash(c)))
                        vals = [""] * len(cols)
                    else:
                        raise Stop(f"{where}有看不懂的一行：{'｜'.join(extra)[:80]}")
                name = join_lines(cells[0]) or last
                if not name:
                    raise Stop(f"{where}：有一行找不到產品名稱")
                last = name
                rows.append({"section": cur[0], "kind": cur[1], "issuer": issuer, "name": name,
                             "ccy": ccy, "cells": vals, "note": note,
                             "vals": [float(PCT.match(v).group(1)) if PCT.match(v) else None for v in vals]})

    if year is None or not rows:
        raise Stop(f"{issuer}：讀不到任何資料行，PDF 版面可能已改")
    for sec in ("tcv", "fr"):
        if not any(r["section"] == sec for r in rows):
            raise Stop(f"{issuer}：找不到{SECTION_NAME[sec]}的表格")
    n_num = sum(1 for r in rows for v in r["vals"] if v is not None)
    n_na = sum(1 for r in rows for v in r["cells"] if v in NA)
    if n_num != pct_text:
        raise Stop(f"{issuer}：表格讀到 {n_num} 個數值，但頁面文字有 {pct_text} 個百分比，可能漏讀")
    print(f"   文字核對：百分比 {pct_text} 個 ✓｜「不適用」表格 {n_na} 個、文字 {na_text} 個（文字包括註釋）")
    return year, rows


# ---------- 保誠格式 ----------
def solid(argb):
    return PatternFill(fill_type="solid", start_color=argb, end_color=argb)


DARK = solid("FF" + NAVY_HEX)
LIGHT = solid("FFDDEBF7")
PALE = solid("FFF4F7FB")
WHITE = solid("FFFFFFFF")
SEP_FILL = solid(SEP_COLOR)
GOOD_FILL = solid("FFE2F0D9")
LOW_FILL = solid("FFFCE4E4")
WHITE_BOLD = Font(bold=True, color="FFFFFFFF")
BLACK_BOLD = Font(bold=True, color="FF000000")
NAME_FONT = Font(bold=True, color=NAME_COLOR)
CUR_FONT = Font(bold=True, color=CUR_COLOR)
KIND_FONT = Font(color="FF2F5597")
LINK_FONT = Font(bold=True, underline="single", color="FF" + NAVY_HEX)
GREY_FONT = Font(color="FFBFBFBF")
GOOD_FONT = Font(color="FF375623")
LOW_FONT = Font(color="FFC00000")
TITLE_FONT = Font(bold=True, size=16, color="FFFFFFFF")
NOTE_FONT = Font(size=10, color="FF44546A")
SECTION_FONT = Font(bold=True, size=12, color="FF" + NAVY_HEX)
LINE = Side(style="thin", color="FFD0D7E2")
GROUP = Side(style="thin", color="FF7F9BBF")
BORDER = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
CENTER_FIT = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
LEFT_FIT = Alignment(horizontal="left", vertical="center", shrink_to_fit=True, indent=1)


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def pct_format(v):
    return "0%" if abs(v * 100 - round(v * 100)) < 1e-9 else "0.0%"


def mean_or_slash(nums):
    return sum(nums) / len(nums) if nums else "/"


def numbers(rows):
    return [v for r in rows for v in r[N_FIXED:] if is_num(v)]


def to_rows(data):
    return [[r["name"], r["ccy"], r["kind"], r["issuer"]]
            + ["/" if v is None else round(v / 100, 4) for v in r["vals"]] for r in data]


def groups_by(rows, idx):
    return [(k or "（未分類）", [r for r in rows if r[idx] == k]) for k in dict.fromkeys(r[idx] for r in rows)]


def build_summary(groups):
    width = N_FIXED + len(YEARS)
    head = ["所有年份平均值"] + [f"持滿{n}年平均值" for n in YEARS[1:]]
    cum = []
    for label, sub in groups:
        if not sub:
            continue
        line = [label]
        for start in range(N_FIXED, width):
            line.append(mean_or_slash([r[c] for r in sub for c in range(start, width) if is_num(r[c])]))
        cum.append(line)
    all_label, everything = groups[0]
    year_line = [all_label] + [mean_or_slash([r[c] for r in everything if is_num(r[c])])
                               for c in range(N_FIXED, width)]
    return {"head": head, "cum": cum, "year_head": [f"保单年度{n}" for n in YEARS], "year": [year_line]}


def write_row(ws, r, col, values):
    for i, v in enumerate(values):
        ws.cell(row=r, column=col + i, value=v)


def dark(cell):
    cell.fill = DARK
    cell.font = WHITE_BOLD


def light(cell):
    cell.fill = LIGHT
    cell.font = BLACK_BOLD


def heat(cell, v, base):
    if not is_num(v):
        cell.fill = base
        cell.font = GREY_FONT
    elif v >= HIGH_LIMIT:
        cell.fill = GOOD_FILL
        cell.font = GOOD_FONT
    elif v < LOW_LIMIT:
        cell.fill = LOW_FILL
        cell.font = LOW_FONT
    else:
        cell.fill = base


def units(s):
    return sum(2 if ord(ch) > 255 else 1 for ch in str(s))


def fit_width(values, low, cap):
    longest = max([units(v) for v in values] or [0])
    return min(max(math.ceil(longest * 1.1) + 4, low), cap)


def add_sheet(wb, title, heading, notes, header, rows, summary, head_label, tab):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    last_col = len(header)
    lows = [30, 10, 14, 26]
    caps = [52, 12, 24, 40]
    widths = [fit_width([header[k]] + [row[k] for row in rows], lows[k], caps[k])
              for k in range(N_FIXED)]

    ws.cell(row=1, column=1, value=heading)
    ws.cell(row=2, column=1, value=notes)
    for rr in (1, 2):
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
        ws.cell(row=rr, column=1).alignment = LEFT
    ws.cell(row=1, column=1).font = TITLE_FONT
    ws.cell(row=1, column=1).fill = DARK
    ws.cell(row=2, column=1).font = NOTE_FONT
    ws.cell(row=2, column=1).fill = PALE
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 22 * (notes.count("\n") + 1)

    write_row(ws, HEAD_ROW, 1, header)
    for c in range(1, last_col + 1):
        cell = ws.cell(row=HEAD_ROW, column=c)
        dark(cell)
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[HEAD_ROW].height = 42

    def key(row):
        return row[0], row[3]

    rr = FIRST
    for i, row in enumerate(rows):
        first = i == 0 or key(rows[i - 1]) != key(row)
        last = i == len(rows) - 1 or key(rows[i + 1]) != key(row)
        if first and i > 0:
            ws.cell(row=rr, column=1).fill = SEP_FILL
            ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
            ws.row_dimensions[rr].height = SEP_H
            rr += 1
        border = Border(left=LINE, right=LINE,
                        top=GROUP if first else LINE,
                        bottom=GROUP if last else LINE)
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=rr, column=c, value=v)
            cell.border = border
            cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
            cell.fill = WHITE
            if c == 1:
                cell.font = NAME_FONT
            elif c == 2:
                cell.font = CUR_FONT
            elif c == 3:
                cell.font = KIND_FONT
            elif c > N_FIXED:
                if is_num(v):
                    cell.number_format = pct_format(v)
                heat(cell, v, WHITE)
        ws.row_dimensions[rr].height = ROW_H
        rr += 1
    data_end = rr - 1

    if rows:
        ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(last_col)}{data_end}"

    lc = N_FIXED
    r = data_end + 2
    ws.cell(row=r, column=lc, value="平均值摘要").font = SECTION_FONT
    ws.row_dimensions[r].height = ROW_H
    r += 1
    table_start = r
    write_row(ws, r, lc, [head_label] + summary["head"])
    r += 1
    value_rows = []
    for line in summary["cum"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    year_head_row = r
    write_row(ws, r, lc, ["當年平均值"] + summary["year_head"])
    r += 1
    for line in summary["year"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    table_end = r - 1
    for row_i in range(table_start, table_end + 1):
        for c in range(lc, last_col + 1):
            cell = ws.cell(row=row_i, column=c)
            cell.alignment = CENTER
            cell.border = BORDER
    for c in range(lc, last_col + 1):
        dark(ws.cell(row=table_start, column=c))
    dark(ws.cell(row=year_head_row, column=lc))
    for c in range(lc + 1, last_col + 1):
        light(ws.cell(row=year_head_row, column=c))
    for row_i in (table_start, year_head_row):
        ws.row_dimensions[row_i].height = 32
    for vr, line in value_rows:
        label_cell = ws.cell(row=vr, column=lc)
        light(label_cell)
        label_cell.alignment = CENTER_FIT
        ws.row_dimensions[vr].height = ROW_H
        for k, v in enumerate(line[1:]):
            cell = ws.cell(row=vr, column=lc + 1 + k)
            cell.alignment = CENTER_FIT
            cell.number_format = "0.00%"
            heat(cell, v, WHITE)
    note_row = table_end + 1
    ws.cell(row=note_row, column=lc,
            value="「持滿N年平均值」＝保單年度N及之後所有數字的平均；"
                  "「當年平均值」＝該年度一欄所有數字的平均。平均值計入所有貨幣，為固定數字，不會隨篩選變動。")
    ws.merge_cells(start_row=note_row, start_column=lc, end_row=note_row, end_column=last_col)
    ws.cell(row=note_row, column=lc).font = NOTE_FONT
    ws.cell(row=note_row, column=lc).alignment = LEFT
    ws.row_dimensions[note_row].height = 32

    for i in range(1, last_col + 1):
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1] if i <= N_FIXED else YEAR_W
    ws.freeze_panes = ws.cell(row=FIRST, column=2)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"


def add_overview(wb, year, today, sheets, fr_rows, remarks):
    ws = wb.create_sheet("概覽", 0)
    ws.sheet_properties.tabColor = "FFC000"
    ws.sheet_view.showGridLines = False
    last_col = 6
    for i, w in enumerate([36, 18, 18, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def banner(r, value, font, fill, height):
        cell = ws.cell(row=r, column=1, value=value)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        cell.font = font
        cell.fill = fill
        cell.alignment = LEFT
        ws.row_dimensions[r].height = height

    def section(r, value):
        ws.cell(row=r, column=1, value=value).font = SECTION_FONT
        ws.row_dimensions[r].height = ROW_H

    def stat_table(r, heading, first_col, groups):
        section(r, heading)
        r += 1
        for c, v in enumerate([first_col, "產品數", "行數", "有數字的格數", "平均值", "低於 80% 的比例"], 1):
            cell = ws.cell(row=r, column=c, value=v)
            dark(cell)
            cell.alignment = CENTER_FIT
            cell.border = BORDER
        ws.row_dimensions[r].height = ROW_H
        for label, sub, link in groups:
            nums = numbers(sub)
            low = sum(1 for v in nums if v < LOW_LIMIT) / len(nums) if nums else "/"
            vals = [label, len({(x[3], x[0]) for x in sub}), len(sub), len(nums), mean_or_slash(nums), low]
            r += 1
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.border = BORDER
                cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
                cell.fill = WHITE
            head = ws.cell(row=r, column=1)
            if link:
                head.font = LINK_FONT
                head.hyperlink = Hyperlink(ref=head.coordinate, location=f"'{link}'!A1")
            else:
                head.font = BLACK_BOLD
            ws.cell(row=r, column=5).number_format = "0.00%"
            heat(ws.cell(row=r, column=5), vals[4], WHITE)
            ws.cell(row=r, column=6).number_format = "0.0%"
            ws.row_dimensions[r].height = ROW_H
        return r + 2

    def lines(r, texts):
        for t in texts:
            r += 1
            cell = ws.cell(row=r, column=1, value="・" + t)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
            cell.font = NOTE_FONT
            cell.alignment = LEFT
            ws.row_dimensions[r].height = 18 * max(1, math.ceil(units(t) / 110)) + 6
        return r

    banner(1, f"蘇黎世人壽分紅表現概覽（報告年度 {year}）", TITLE_FONT, DARK, 40)
    banner(2, f"資料來源：蘇黎世香港網站｜更新日期：{today}", NOTE_FONT, PALE, ROW_H)

    r = stat_table(4, "資料概況", "工作表", [(title, rows, title) for title, rows in sheets])
    r = stat_table(r, "分紅實現率：按紅利類型", "紅利類型",
                   [(label, sub, None) for label, sub in groups_by(fr_rows, 2)])

    section(r, "顏色說明")
    samples = [(1.0, "100% 或以上：與銷售時的說明一致或更好"),
               (0.79, "低於 80%：明顯低於銷售時的說明"),
               ("/", "沒有數字：官方標示「不適用」或未能提供")]
    for v, desc in samples:
        r += 1
        cell = ws.cell(row=r, column=1, value=v)
        cell.border = BORDER
        cell.alignment = CENTER_FIT
        if is_num(v):
            cell.number_format = "0%"
        heat(cell, v, WHITE)
        d = ws.cell(row=r, column=2, value=desc)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
        d.alignment = LEFT_FIT
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "備註")
    r = lines(r, remarks)

    r += 2
    section(r, "使用提示")
    lines(r, ["點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
              "工作表第 3 行有篩選箭嘴，可按貨幣、類型或繕發公司篩選。",
              "向右或向下捲動時，標題列和產品名稱會固定不動。",
              "平均值摘要在工作表最底部。",
              "手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。"])


def build_excel(path, year, data, today, index_url, docs):
    tcv = to_rows([r for r in data if r["section"] == "tcv"])
    fr = to_rows([r for r in data if r["section"] == "fr"])
    legend = ("綠色＝100% 或以上；紅色＝低於 80%；「/」＝官方標示「不適用」或未能提供。\n"
              f"資料來源：蘇黎世香港網站（報告年度 {year}）｜更新日期：{today}")
    tcv_notes = ("總現金價值比率＝報告年度內所有相關保單的實際總現金價值 ÷ 銷售時所示的金額。"
                 "100% 即與銷售時一致，高於 100% 即比銷售時多。\n" + legend)
    fr_notes = ("分紅實現率＝實際的非保證紅利 ÷ 銷售時說明的金額。"
                "周年紅利按有效保單的累積紅利及利息計算；終期紅利按報告年度內終止的保單計算。\n" + legend)
    remarks = ["以下說明摘自 2025 報告年度的官方文件，如官方日後修改，以官方文件為準。",
               "「不適用」的原因：相關產品不再可供購買；沒有有效的保單；投保時的終期紅利和周年紅利為零；"
               "或該報告年度中並無相關保單終結（只適用於終期紅利）。",
               "蘇黎世人壽保險有限公司的保單組合已經轉移至蘇黎世人壽保險（香港）有限公司。"]
    for r in data:
        if r["note"]:
            line = f"{r['name']}（{SECTION_NAME[r['section']]}）：{r['note']}"
            if line not in remarks:
                remarks.append(line)
    remarks.append(f"官方來源（網頁）：{PAGE}")
    remarks.append(f"官方來源（索引文件）：{index_url}")
    for d in docs:
        remarks.append(f"官方來源（{d['issuer']}）：{d['url']}")

    fixed = ["产品名称", "货币", "类型", "缮发公司"]
    header = fixed + [f"保单年度{n}" for n in YEARS]
    tcv_title, fr_title = f"{year}總現金價值比率", f"{year}分紅實現率"
    tcv_groups = [("所有產品", tcv)] + [g for g in groups_by(tcv, 3) if numbers(g[1])]
    fr_groups = [("所有分紅", fr)] + [g for g in groups_by(fr, 2) if numbers(g[1])]

    wb = Workbook()
    wb.remove(wb.active)
    add_sheet(wb, tcv_title, f"蘇黎世人壽總現金價值比率（報告年度 {year}）", tcv_notes,
              header, tcv, build_summary(tcv_groups), "繕發公司", "2E75B6")
    add_sheet(wb, fr_title, f"蘇黎世人壽分紅實現率（報告年度 {year}）", fr_notes,
              header, fr, build_summary(fr_groups), "紅利類型", "548235")
    add_overview(wb, year, today, [(tcv_title, tcv), (fr_title, fr)], fr, remarks)
    wb.save(path)
    print("工作表：", "、".join(wb.sheetnames))


# ---------- Notion（與恒生版相同做法） ----------
def check(r, what):
    if r.status_code >= 300:
        hint = HINTS.get(r.status_code, "")
        raise Stop(f"{what} 失敗：{r.status_code} {r.text[:300]}" + (f"\n→ {hint}" if hint else ""))
    return r.json()


def resolve_source(raw, headers):
    m = re.search(r"[0-9a-fA-F]{32}", raw.replace("-", ""))
    ident = m.group(0) if m else raw

    def first_source(db_id):
        r = requests.get(f"{API}/databases/{db_id}", headers=headers, timeout=30)
        if r.status_code == 200:
            ds = r.json().get("data_sources") or []
            return ds[0]["id"] if ds else None
        return None

    found = first_source(ident)
    if found:
        return found
    r = requests.get(f"{API}/data_sources/{ident}", headers=headers, timeout=30)
    if r.status_code == 200:
        return ident
    status = r.status_code
    r = requests.get(f"{API}/blocks/{ident}/children?page_size=100", headers=headers, timeout=30)
    if r.status_code == 200:
        for b in r.json().get("results", []):
            if b.get("type") == "child_database":
                found = first_source(b["id"])
                if found:
                    return found
    raise Stop(f"找不到蘇黎世的年份資料庫（狀態 {status}）。401＝金鑰不正確；"
               "404＝頁面未連接「AIA 自動更新」或 NOTION_DATABASE_ID 不正確。")


def page_title(page):
    for prop in page.get("properties", {}).values():
        if prop.get("type") == "title":
            return "".join(t.get("plain_text", "") for t in prop.get("title", []))
    return ""


def file_blocks(page_id, headers):
    r = requests.get(f"{API}/blocks/{page_id}/children?page_size=100", headers=headers, timeout=30)
    out = []
    for block in check(r, "讀取頁面內容").get("results", []):
        if block["type"] == "file":
            info = block["file"]
            cap = "".join(t.get("plain_text", "") for t in info.get("caption", []))
            out.append((block["id"], info.get("name", ""), cap))
    return out


def upload(path, year, digest, today):
    token = os.environ.get("NOTION_TOKEN", "").strip()
    target = os.environ.get("NOTION_DATABASE_ID", "").strip()
    if not token or not target:
        raise Stop("未設定 NOTION_TOKEN 或 NOTION_DATABASE_ID")
    headers = {"Authorization": f"Bearer {token}", "Notion-Version": VERSION,
               "Content-Type": "application/json"}
    source = resolve_source(target, headers)
    r = requests.post(f"{API}/data_sources/{source}/query", headers=headers, json={"page_size": 100}, timeout=30)
    pages = check(r, "讀取年份資料庫").get("results", [])
    page = next((p for p in pages if str(year) in page_title(p)), None)
    if page is None:
        body = {"parent": {"type": "data_source_id", "data_source_id": source},
                "properties": {"title": {"title": [{"text": {"content": str(year)}}]}}}
        r = requests.post(f"{API}/pages", headers=headers, json=body, timeout=30)
        page = check(r, "建立年度頁面")
        print(f"已建立頁面：{year}")
    else:
        print("使用現有頁面：", page_title(page))
    page_id = page["id"]

    blocks = file_blocks(page_id, headers)
    if any(AUTO in name and digest in cap for _, name, cap in blocks):
        print("數據和格式與上次上傳相同，略過上傳")
        return
    old_auto = [bid for bid, name, _ in blocks if AUTO in name]

    filename = os.path.basename(path)
    r = requests.post(f"{API}/file_uploads", headers=headers,
                      json={"mode": "single_part", "filename": filename, "content_type": XLSX}, timeout=30)
    upload_id = check(r, "建立上傳")["id"]
    with open(path, "rb") as fh:
        r = requests.post(f"{API}/file_uploads/{upload_id}/send",
                          headers={"Authorization": f"Bearer {token}", "Notion-Version": VERSION},
                          files={"file": (filename, fh.read(), XLSX)}, timeout=120)
    check(r, "傳送檔案")
    block = {"object": "block", "type": "file",
             "file": {"type": "file_upload", "file_upload": {"id": upload_id},
                      "caption": [{"type": "text", "text": {"content": f"更新 {today}｜{digest}"}}]}}
    r = requests.patch(f"{API}/blocks/{page_id}/children", headers=headers, json={"children": [block]}, timeout=60)
    check(r, "加入檔案到頁面")
    for bid in old_auto:
        r = requests.delete(f"{API}/blocks/{bid}", headers=headers, timeout=30)
        check(r, "刪除上一次的自動生成檔案")
    print(f"已上傳：{filename}（刪除舊的自動生成檔案 {len(old_auto)} 個）")


# ---------- 主流程 ----------
def main():
    index_url, docs = load_sources()
    year, data = None, []
    for d in docs:
        with pdfplumber.open(io.BytesIO(d["content"])) as pdf:
            y, part = parse_doc(pdf, d["issuer"])
        if year is None:
            year = y
        elif y != year:
            raise Stop(f"兩份文件的報告年度不同（{year} 和 {y}）")
        data += part
    keys = [(r["section"], r["issuer"], r["name"], r["ccy"], r["kind"]) for r in data]
    dup = sorted({k for k in keys if keys.count(k) > 1})
    if dup:
        raise Stop(f"有重複的資料行：{dup[:5]}")

    today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d")
    print(f"報告年度：{year}")
    for sec in ("tcv", "fr"):
        sub = [r for r in data if r["section"] == sec]
        print(f"== {SECTION_NAME[sec]}")
        for r in sub:
            shown = " ".join("/" if v is None else f"{v:g}%" for v in r["vals"])
            print(f"   {r['name']}｜{r['ccy']}｜{r['kind']}｜{r['issuer']}｜{shown}"
                  + (f"｜註：{r['note']}" if r["note"] else ""))
        nums = [v for r in sub for v in r["vals"] if v is not None]
        blanks = sum(1 for r in sub for v in r["vals"] if v is None)
        stats = (f"｜平均 {sum(nums) / len(nums):.2f}%｜範圍 {min(nums):g}%–{max(nums):g}%"
                 f"｜低於 80%：{sum(1 for x in nums if x < 80)}｜100% 或以上：{sum(1 for x in nums if x >= 100)}"
                 if nums else "")
        print(f"   合計：產品 {len({(r['issuer'], r['name']) for r in sub})}｜資料行 {len(sub)}"
              f"｜數值 {len(nums)}｜沒有數字 {blanks}{stats}")
        if year == 2025:
            ok = (len(sub), len(nums)) == EXPECT_2025[sec]
            print("   與探查結果" + ("一致 ✓" if ok else f"不一致 ⚠（探查為 {EXPECT_2025[sec]}）"))

    path = f"蘇黎世人壽分紅實現率及總現金價值比率（{year}年）{AUTO}.xlsx"
    build_excel(path, year, data, today, index_url, docs)
    print(f"已生成 {path}（{os.path.getsize(path) // 1024} KB）")

    digest = hashlib.sha256(json.dumps(
        [FORMAT] + [[r["section"], r["issuer"], r["name"], r["ccy"], r["kind"], r["cells"], r["note"]]
                    for r in data], ensure_ascii=False).encode()).hexdigest()[:12]
    if os.environ.get("UPLOAD", "").lower() == "true":
        upload(path, year, digest, today)
        print("完成")
    else:
        print("只讀試跑：沒有上傳")


if __name__ == "__main__":
    try:
        main()
    except Stop as e:
        print(f"✗ 已停止：{e}")
        sys.exit(1)
