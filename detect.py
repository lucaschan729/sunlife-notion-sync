import collections
import html
import os
import re
import subprocess
import sys

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi"])
from curl_cffi import requests as creq

URL = "https://www.sunlife.com.hk/zh-hant/insurance/fulfillment-ratios-of-respective-products/"
PCT = re.compile(r"[0-9]+(?:[.][0-9]+)?\s*%")
out = []


def say(s=""):
    print(s, flush=True)
    out.append(s)


def clean(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", html.unescape(s)).strip()


def cells(tr):
    res = []
    for attrs, inner in re.findall(r"<t[hd]\b([^>]*)>(.*?)</t[hd]>", tr, flags=re.S | re.I):
        t = clean(inner)[:40]
        for key, tag in (("colspan", "橫"), ("rowspan", "直")):
            m = re.search(key + r"\s*=\s*[\"']?(\d+)", attrs, flags=re.I)
            if m and m.group(1) != "1":
                t += "[" + tag + m.group(1) + "]"
        res.append(t)
    return res


def before(pos, raw, keep):
    return clean(body[max(pos - raw, 0):pos])[-keep:]


r = creq.get(URL, impersonate="chrome", timeout=60)
body = r.content.decode("utf-8", "replace")
say(f"狀態 {r.status_code}｜原始長度 {len(body)}")

say("=== 下拉選單 ===")
for i, m in enumerate(re.finditer(r"<select\b([^>]*)>(.*?)</select>", body, flags=re.S | re.I), 1):
    opts = re.findall(r"<option\b([^>]*)>(.*?)</option>", m.group(2), flags=re.S | re.I)
    attrs = re.sub(r"\s+", " ", m.group(1)).strip()[:150]
    say(f"選單 {i}｜{len(opts)} 項｜屬性 {attrs}｜前文 {before(m.start(), 800, 60)}")
    for a, t in opts[:8] + opts[-3:]:
        say(f"    {re.sub(r'\s+', ' ', a).strip()[:90]}｜{clean(t)[:50]}")

info = []
for i, m in enumerate(re.finditer(r"<table\b([^>]*)>(.*?)</table>", body, flags=re.S | re.I), 1):
    rows = [cells(tr) for tr in re.findall(r"<tr\b.*?</tr>", m.group(2), flags=re.S | re.I)]
    text = clean(m.group(2))
    ids = re.findall(r"""\bid\s*=\s*["']([^"']+)""", body[max(m.start() - 3000, 0):m.start()])[-2:]
    info.append({"i": i, "m": m, "rows": rows, "pct": len(PCT.findall(text)),
                 "year": "保單年度" in text, "report": "報告年度" in text, "ids": ids,
                 "cols": max([len(x) for x in rows] or [0])})

say()
say("=== 表格形狀統計（行x欄：數量）===")
shape = collections.Counter(f"{len(t['rows'])}x{t['cols']}" for t in info)
say("、".join(f"{k}：{v}" for k, v in shape.most_common(15)))
say(f"含「保單年度」{sum(t['year'] for t in info)} 個｜含「報告年度」{sum(t['report'] for t in info)} 個"
    f"｜有百分比 {sum(1 for t in info if t['pct'])} 個")

say()
say("=== 全部表格索引 ===")
for t in info:
    flag = ("年" if t["year"] else "") + ("報" if t["report"] else "")
    first = "｜".join(t["rows"][0][:4]) if t["rows"] else ""
    line = (f"{t['i']:>3}｜{len(t['rows'])}x{t['cols']}｜%{t['pct']}｜{flag}｜id {','.join(t['ids'])[:50]}"
            f"｜前文 {before(t['m'].start(), 1500, 45)}｜首行 {first}")
    say(line[:270])

data = [t for t in info if t["year"] and t["pct"] and t["i"] > 6]
blank = [t for t in info if t["year"] and not t["pct"] and t["i"] > 6]
mid = len(data) // 2
picks = data[:4] + data[mid:mid + 2] + data[-2:] + blank[:2]
seen = set()
say()
say("=== 抽樣表格完整內容 ===")
for t in picks:
    if t["i"] in seen:
        continue
    seen.add(t["i"])
    say(f"--- 表格 {t['i']}｜屬性 {re.sub(r'\s+', ' ', t['m'].group(1)).strip()[:120]}")
    say(f"    前文：{before(t['m'].start(), 2500, 220)}")
    for row in t["rows"][:20]:
        say("    " + "｜".join(row)[:330])

os.makedirs("output", exist_ok=True)
with open("output/sunlife-detect.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(out))
