import html as htmllib
import io
import re
import sys
import time
from urllib.parse import urljoin

import pdfplumber
import requests

PAGE = "https://www.zurich.com.hk/zh-hk/customer-services/life-policy-services/zlihk"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
KEY = re.compile(r"fulfil|ratio|dividend|bonus|participat", re.I)
MAX_PDF = 8
MAX_PAGES = 30


def get(url):
    for _ in range(3):
        try:
            r = requests.get(url, headers=UA, timeout=60)
            print(f"HTTP {r.status_code}｜{len(r.content)} bytes｜server={r.headers.get('server')}｜{url}")
            if r.status_code == 200:
                return r
        except requests.RequestException as e:
            print(f"{type(e).__name__}｜{url}")
        time.sleep(5)
    return None


def clean(s):
    return re.sub(r"\s+", " ", htmllib.unescape(re.sub(r"<[^>]+>", " ", s))).strip()


def dump(url):
    print(f"\n########## {url}")
    r = get(url)
    if r is None or not r.content.startswith(b"%PDF"):
        print("讀不到，或者不是 PDF")
        return
    print(f"last-modified：{r.headers.get('last-modified')}")
    with pdfplumber.open(io.BytesIO(r.content)) as doc:
        print(f"共 {len(doc.pages)} 頁")
        for no, page in enumerate(doc.pages[:MAX_PAGES], 1):
            text = page.extract_text() or ""
            tables = page.extract_tables()
            print(f"\n== 第 {no} 頁｜文字 {len(text)} 字｜百分比 {len(re.findall(r'[%％]', text))} 個"
                  f"｜不適用 {text.count('不適用')}｜N/A {len(re.findall(r'N/?A', text))}｜表格 {len(tables)} 個")
            print(text[:1500 if no <= 2 else 500])
            for t in tables:
                print(f"-- 表格：{len(t)} 行 x {max(len(x) for x in t)} 欄")
                for row in t[:8]:
                    print("   ", [(c or "").replace("\n", "⏎")[:20] for c in row])


def main():
    r = get(PAGE)
    if r is None:
        print("讀不到官方頁")
        sys.exit(1)
    raw = r.content.decode("utf-8", errors="replace").replace("\\u002F", "/").replace("\\/", "/")

    hits = [m.start() for m in re.finditer("實現率", raw)]
    print(f"\n「實現率」在頁面原始碼出現 {len(hits)} 次")
    for p in hits[:6]:
        print("-----")
        print(raw[max(0, p - 300):p + 900])

    all_pdf, cands = [], []
    for m in re.finditer(r"(?:https?://|/)[^\s\"'<>\\()]+?\.pdf", raw, flags=re.I):
        url = urljoin(PAGE, m.group(0))
        if url not in all_pdf:
            all_pdf.append(url)
        ctx = clean(raw[max(0, m.start() - 400):m.end() + 200])
        if (KEY.search(url) or "實現率" in ctx) and url not in cands:
            cands.append(url)
    print(f"\n頁面共有 {len(all_pdf)} 個 PDF 連結，候選 {len(cands)} 個：")
    for u in cands:
        print("  ", u)
    if not cands:
        print("找不到候選 PDF，頁面可能由 JavaScript 載入。全部 PDF 連結如下：")
        for u in all_pdf:
            print("  ", u)
    for u in cands[:MAX_PDF]:
        dump(u)


if __name__ == "__main__":
    main()
