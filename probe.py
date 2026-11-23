import io
import re
import sys
import time
from urllib.parse import urljoin

import pdfplumber
import requests

PAGE = "https://www.zurich.com.hk/zh-hk/customer-services/life-policy-services/zlihk"
HEADING = "分紅終身人壽保險之實現率"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
PDF_URL = re.compile(r"https?://[^\s\"'<>\\]+?\.pdf", re.I)
MAX_DOCS = 8
MAX_PAGES = 40


def get(url):
    for _ in range(3):
        try:
            r = requests.get(url, headers=UA, timeout=60)
            print(f"HTTP {r.status_code}｜{len(r.content)} bytes｜{r.headers.get('content-type')}"
                  f"｜last-modified {r.headers.get('last-modified')}｜{url}")
            if r.status_code == 200:
                return r
        except requests.RequestException as e:
            print(f"{type(e).__name__}｜{url}")
        time.sleep(5)
    return None


def find_index_pdf(raw):
    for m in re.finditer(HEADING, raw):
        hit = PDF_URL.search(raw, m.end(), m.end() + 2000)
        if hit:
            return hit.group(0)
    for m in PDF_URL.finditer(raw):
        if "fulfil" in m.group(0).lower():
            return m.group(0)
    return None


def links_of(page):
    out = []
    for a in page.annots or []:
        uri = a.get("uri")
        if isinstance(uri, bytes):
            uri = uri.decode("utf-8", errors="replace")
        if uri:
            out.append((round(a.get("top") or 0), round(a.get("x0") or 0), uri.strip()))
    return out


def dump_pdf(content):
    with pdfplumber.open(io.BytesIO(content)) as doc:
        print(f"共 {len(doc.pages)} 頁")
        total_pct = total_na = 0
        for no, page in enumerate(doc.pages[:MAX_PAGES], 1):
            text = page.extract_text() or ""
            tables = page.extract_tables()
            pct = len(re.findall(r"\d+(?:\.\d+)?\s*[%％]", text))
            na = text.count("不適用") + len(re.findall(r"\bN/?A\b", text))
            total_pct += pct
            total_na += na
            print(f"\n== 第 {no} 頁｜文字 {len(text)} 字｜百分比 {pct} 個｜不適用／N/A {na} 個"
                  f"｜表格 {len(tables)} 個｜連結 {len(links_of(page))} 個")
            print(text[:1500 if no <= 2 else 500])
            for t in tables:
                print(f"-- 表格：{len(t)} 行 x {max(len(x) for x in t)} 欄")
                for row in t[:8]:
                    print("   ", [(c or "").replace("\n", "⏎")[:16] for c in row])
        print(f"\n合計：百分比 {total_pct} 個｜不適用／N/A {total_na} 個")


def dump(url):
    print(f"\n########## {url}")
    r = get(url)
    if r is None:
        print("讀不到")
        return
    if r.content.startswith(b"%PDF"):
        dump_pdf(r.content)
        return
    page = r.content.decode("utf-8", errors="replace")
    found = list(dict.fromkeys(PDF_URL.findall(page)))
    print(f"不是 PDF；頁面內有 {len(found)} 個 PDF 連結：")
    for u in found[:30]:
        print("  ", u)


def main():
    r = get(PAGE)
    if r is None:
        print("讀不到官方頁")
        sys.exit(1)
    raw = r.content.decode("utf-8", errors="replace").replace("\\u002F", "/").replace("\\/", "/")
    index = find_index_pdf(raw)
    if not index:
        print("官方頁找不到分紅實現率的 PDF 連結")
        sys.exit(1)
    print("\n索引文件：", index)
    r = get(index)
    if r is None or not r.content.startswith(b"%PDF"):
        print("讀不到索引文件")
        sys.exit(1)

    chosen, everything = [], []
    with pdfplumber.open(io.BytesIO(r.content)) as doc:
        for no, page in enumerate(doc.pages, 1):
            text = page.extract_text() or ""
            links = links_of(page)
            print(f"\n== 索引第 {no} 頁｜註解 {len(page.annots or [])} 個｜連結 {len(links)} 個")
            try:
                for ln in page.extract_text_lines():
                    print(f"   文字 y={round(ln['top'])}｜{ln['text'][:60]}")
            except Exception as e:
                print(f"   （讀不到文字位置：{type(e).__name__}）")
            for top, x0, uri in links:
                print(f"   連結 y={top} x={x0}｜{uri}")
                full = urljoin(index, uri)
                if full not in everything:
                    everything.append(full)
                if "繕發" in text and full not in chosen:
                    chosen.append(full)

    targets = chosen or everything
    if not targets:
        print("\n索引文件內找不到任何連結，請把這份日誌傳回，我再換方法")
        return
    print(f"\n準備讀取 {min(len(targets), MAX_DOCS)} 份文件（共找到 {len(targets)} 個連結）")
    for u in targets[:MAX_DOCS]:
        dump(u)


if __name__ == "__main__":
    main()
