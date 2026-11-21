import html
import re
import subprocess
import sys

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi"])
from curl_cffi import requests as creq

URL = "https://www.sunlife.com.hk/zh-hant/insurance/fulfillment-ratios-of-respective-products/"


def clean(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", html.unescape(s)).strip()


r = creq.get(URL, impersonate="chrome", timeout=60)
body = r.content.decode("utf-8", "replace")
text = clean(body)
tables = re.findall(r"<table\b.*?</table>", body, flags=re.S | re.I)
print(f"狀態 {r.status_code}｜原始長度 {len(body)}｜文字長度 {len(text)}")
print(f"表格 {len(tables)} 個｜百分比 {len(re.findall(r'[0-9]+(?:[.][0-9]+)?\s*%', text))} 個")
for kw in ["分紅實現率", "總現金價值比率", "保單年度", "報告年度", "周年紅利", "終期紅利", "美元", "港元"]:
    print(f"  關鍵字「{kw}」：{text.count(kw)} 次")

print("=== 標題 ===")
for h in re.findall(r"<h[1-4]\b[^>]*>(.*?)</h[1-4]>", body, flags=re.S | re.I)[:40]:
    if clean(h):
        print("  ", clean(h)[:80])

print("=== 可能的資料連結 ===")
links = re.findall(r"""(?:href|src|action|data-[\w-]+)\s*=\s*["']([^"']+)["']""", body)
pat = re.compile(r"\.(pdf|json|xlsx?|csv)\b|fulfil|ratio|dividend|bonus|api|ajax", re.I)
for u in list(dict.fromkeys(x for x in links if pat.search(x)))[:60]:
    print("  ", u[:200])

print("=== iframe ===")
for u in re.findall(r"""<iframe\b[^>]*src\s*=\s*["']([^"']+)""", body, flags=re.I)[:10]:
    print("  ", u[:200])

print("=== 下拉選單 ===")
for i, sel in enumerate(re.findall(r"<select\b.*?</select>", body, flags=re.S | re.I)[:8], 1):
    opts = [clean(o) for o in re.findall(r"<option\b[^>]*>(.*?)</option>", sel, flags=re.S | re.I)]
    print(f"  選單 {i}：{len(opts)} 項｜", "、".join(opts[:15])[:300])

print("=== 首 6 個表格 ===")
for i, t in enumerate(tables[:6], 1):
    rows = re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I)
    print(f"  表格 {i}：{len(rows)} 行")
    for tr in rows[:4]:
        cells = [clean(c) for c in re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", tr, flags=re.S | re.I)]
        print("     ", "｜".join(cells)[:300])

print("=== 文字抽樣 ===")
m = re.search(r"[0-9]+(?:[.][0-9]+)?\s*%", text)
pos = m.start() if m else text.find("分紅實現率")
print(text[max(pos - 300, 0):pos + 500] if pos >= 0 else text[:800])
