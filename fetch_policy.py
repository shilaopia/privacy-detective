# -*- coding: utf-8 -*-
"""抓取隐私政策网页并转为纯文本。
用法: python fetch_policy.py <url> <out_path> [来源标注]
输出文件头部自动写入 来源URL + 访问日期。
PDF 仅原样保存并在头部标注（需后续人工/PyMuPDF 处理）。
"""
import sys, re, ssl, gzip, io, html, datetime
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")

def fetch(url):
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = Request(url, headers={
        "User-Agent": UA,
        "Accept": "text/html,application/xhtml+xml,application/pdf;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Accept-Encoding": "gzip",
    })
    with urlopen(req, timeout=40, context=ctx) as r:
        data = r.read()
        if r.headers.get("Content-Encoding") == "gzip":
            data = gzip.decompress(data)
        ctype = r.headers.get("Content-Type", "")
        return data, ctype, r.geturl()

def decode(data, ctype):
    m = re.search(r"charset=([\w-]+)", ctype, re.I)
    cands = []
    if m: cands.append(m.group(1))
    head = data[:4096].decode("ascii", "ignore")
    m2 = re.search(r'charset=["\']?([\w-]+)', head, re.I)
    if m2: cands.insert(0, m2.group(1))
    cands += ["utf-8", "gb18030"]
    for enc in cands:
        try:
            return data.decode(enc)
        except Exception:
            continue
    return data.decode("utf-8", "replace")

def html_to_text(doc):
    doc = re.sub(r"(?is)<(script|style|noscript|svg)[^>]*>.*?</\1>", " ", doc)
    doc = re.sub(r"(?is)<!--.*?-->", " ", doc)
    doc = re.sub(r"(?i)<br\s*/?>", "\n", doc)
    doc = re.sub(r"(?i)</(p|div|li|tr|h[1-6]|section|article|table)>", "\n", doc)
    doc = re.sub(r"(?i)</t[dh]>", "\t", doc)
    txt = re.sub(r"(?s)<[^>]+>", " ", doc)
    txt = html.unescape(txt)
    txt = re.sub(r"[ \t ​]+", " ", txt)
    txt = re.sub(r"\n\s*\n+", "\n", txt)
    lines = [ln.strip() for ln in txt.splitlines()]
    return "\n".join(ln for ln in lines if ln)

def main():
    url, out = sys.argv[1], sys.argv[2]
    note = sys.argv[3] if len(sys.argv) > 3 else ""
    try:
        data, ctype, final_url = fetch(url)
    except (HTTPError, URLError, Exception) as e:
        print(f"FETCH_FAIL {type(e).__name__}: {e}")
        sys.exit(2)
    today = datetime.date.today().isoformat()
    header = [f"来源URL: {url}"]
    if final_url != url:
        header.append(f"最终URL: {final_url}")
    header.append(f"访问日期: {today}")
    if note:
        header.append(f"标注: {note}")
    if "pdf" in ctype.lower() or data[:5] == b"%PDF-":
        with open(out, "wb") as f:
            f.write(data)
        print(f"SAVED_PDF {len(data)} bytes -> {out} (需PyMuPDF另行提取)")
        return
    doc = decode(data, ctype)
    text = html_to_text(doc)
    body = "\n".join(header) + "\n\n" + text + "\n"
    with open(out, "w", encoding="utf-8") as f:
        f.write(body)
    n = len(text)
    hit = ("个人信息" in text) or ("personal information" in text.lower()) or ("privacy" in text.lower())
    print(f"SAVED {n} chars, 关键词命中={hit} -> {out}")

if __name__ == "__main__":
    main()
