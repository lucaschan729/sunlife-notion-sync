import collections
import datetime
import html as htmllib
import os
import re
import subprocess
import sys
import time

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi", "openpyxl", "requests"])
import requests
from curl_cffi import requests as creq
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties

PAGE = "https://www.sunlife.com.hk/zh-hant/insurance/fulfillment-ratios-of-respective-products/"
LABELS = ["總現金價值比率", "分紅實現率"]
DEFS = {
    "總現金價值比率": "總現金價值比率＝實際總現金價值 ÷ 銷售時說明的總現金價值。100% 即與銷售時一致，高於 100% 即比銷售時多。定義以永明網站的說明為準。",
    "分紅實現率": "分紅實現率＝實際派發的非保證紅利 ÷ 銷售時說明的金額。100% 即與預期一致，高於 100% 即比預期多。定義以永明網站的說明為準。",
}
LEGEND = "綠色＝100% 或以上；紅色＝低於 80%；「/」＝沒有數字（網站顯示「已停售」「尚未推出」或「不適用」）。"
AVG_NOTE = "「持滿N年平均值」＝保單年度N及之後所有數字的平均；「當年平均值」＝該年度一欄所有數字的平均。平均值為固定數字，不會隨篩選變動。"
PCT = re.compile(r"(\d+(?:\.\d+)?)\s*%")
TITLE = re.compile(r"(\d{4})\s*報告年度的\s*([^。，,；;]{0,30}?)\s*(之分紅實現率|總現金價值比率)")
YEAR = re.compile(r"第\s*(\d+)\s*個保單年度\s*(\+?)\s*[（(]\s*於(.+?)開始生\s*效之保單\s*[）)]")
NAVY, GREY, DARK = "1F3864", "8C8C8C", "262626"
SEP_FILL = PatternFill("solid", fgColor="EDEDED")
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ROW_HEIGHT = 30
HEAD_ROW = 5
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def get(url):
    last = ""
    for i in range(3):
        try:
            r = creq.get(url, impersonate="chrome", timeout=60)
            if r.status_code == 200:
                return r.content.decode("utf-8", "replace")
            last = f"狀態 {r.status_code}"
        except Exception as e:
            last = str(e)[:80]
        time.sleep(3 * (i + 1))
    raise RuntimeError(last)


def clean(s):
    s = re.sub(r"<(script|style|sup)\b.*?</\1>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


def product(s):
    s = re.sub(r"\s+\d(?=\s*(?:[,，、]|$))", "", s)
    return re.sub(r"\s*[,，]\s*", "、", s).strip()


def value(s):
    m = PCT.search(s)
    return float(m.group(1)) / 100 if m else "/"


def cells(tr):
    row = []
    for attrs, inner in re.findall(r"<t[hd]\b([^>]*)>(.*?)</t[hd]>", tr, flags=re.S | re.I):
        m = re.search(r"colspan\s*=\s*[\"']?(\d+)", attrs, flags=re.I)
        row += [clean(inner)] * (int(m.group(1)) if m else 1)
    return row


def parse(body, log):
    sets = {label: [] for label in LABELS}
    ycols, ysubs, end = [], {}, 0
    for m in re.finditer(r"<table\b[^>]*>(.*?)</table>", body, flags=re.S | re.I):
        gap = clean(body[max(end, m.start() - 6000):m.start()])
        end = m.end()
        rows = [cells(tr) for tr in re.findall(r"<tr\b.*?</tr>", m.group(1), flags=re.S | re.I)]
        hi = next((i for i, row in enumerate(rows) if sum("保單年度" in c for c in row) >= 2), None)
        if hi is None:
            continue
        log["tables"] += 1
        head = rows[hi]
        idx = [i for i, c in enumerate(head) if "保單年度" in c]
        first = rows[hi + 1][0][:30] if len(rows) > hi + 1 and rows[hi + 1] else ""
        titles = list(TITLE.finditer(gap))
        if not titles:
            log["untitled"].append(f"{first}｜前文：{gap[-60:]}")
            continue
        t = titles[-1]
        log["years"][t.group(1)] += 1
        label = "總現金價值比率" if t.group(3) == "總現金價值比率" else "分紅實現率"
        kind = t.group(2).strip().replace("或", "/")
        keys = []
        for i in idx:
            ym = YEAR.search(head[i])
            if ym:
                key = f"保單年度{ym.group(1)}{ym.group(2)}"
                ys = re.findall(r"\d{4}", ym.group(3))
                sub = ys[0] if len(ys) == 1 else (f"{max(ys)}或之前" if ys else "")
            else:
                key, sub = re.sub(r"\s+", "", head[i])[:20], ""
            if key not in ycols:
                ycols.append(key)
            ysubs.setdefault(key, sub)
            keys.append(key)
        for row in rows[hi + 1:]:
            if len(row) <= idx[-1] or not row[0]:
                continue
            sets[label].append({"name": product(row[0]), "cat": row[1] if idx[0] > 1 else "", "kind": kind,
                                "vals": {k: value(row[i]) for k, i in zip(keys, idx)}})
    return sets, ycols, ysubs


def pct_format(v):
    return "0%" if abs(v * 100 - round(v * 100)) < 1e-6 else "0.0%"


def numbers(rows):
    return [v for r in rows for v in r["vals"].values() if isinstance(v, float)]


def mean(nums):
    return sum(nums) / len(nums) if nums else None


def averages(rows, ycols):
    return [mean([r["vals"][k] for r in rows if isinstance(r["vals"].get(k), float)]) for k in ycols]


def cumulative(rows, ycols):
    return [mean([r["vals"][k] for r in rows for k in ycols[i:] if isinstance(r["vals"].get(k), float)])
            for i in range(len(ycols))]


def hold_label(i, ycols):
    if i == 0:
        return "所有年份平均值"
    n = ycols[i].replace("保單年度", "")
    if n.endswith("+"):
        return f"持滿{n[:-1]}年+平均值"
    if i + 1 < len(ycols) and ycols[i + 1].endswith("+"):
        return f"第{n}年平均值"
    return f"持滿{n}年平均值"


def head_cell(c):
    c.font = Font(bold=True, color="FFFFFF", size=10)
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = CENTER
    c.border = BOX


def write_averages(ws, start, label, rows, ycols, ysubs, nlead):
    last = nlead + len(ycols)
    kinds = list(dict.fromkeys(x["kind"] for x in rows))
    multi = len(kinds) > 1
    all_name = "所有分红" if multi else "所有產品"
    groups = [(all_name, rows)]
    if multi:
        groups += [(k, [x for x in rows if x["kind"] == k]) for k in kinds]
    groups = [(t, s) for t, s in groups if numbers(s)]

    def line(r, title, vals):
        c = ws.cell(row=r, column=nlead, value=title)
        c.font = Font(bold=True, size=10, color=DARK)
        c.alignment = CENTER
        c.border = BOX
        for j, v in enumerate(vals):
            c = ws.cell(row=r, column=nlead + 1 + j, value=v if v is not None else "/")
            c.font = Font(bold=True, size=10)
            c.alignment = CENTER
            c.border = BOX
            if v is not None:
                c.number_format = "0.00%"
        ws.row_dimensions[r].height = 24

    def heading(r, title, heads):
        head_cell(ws.cell(row=r, column=nlead, value=title))
        for j, h in enumerate(heads):
            head_cell(ws.cell(row=r, column=nlead + 1 + j, value=h))
        ws.row_dimensions[r].height = 40

    r = start
    c = ws.cell(row=r, column=nlead, value="平均值摘要")
    c.font = Font(bold=True, size=11, color=NAVY)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[r].height = 24
    r += 1
    heading(r, "分紅類型" if multi else label, [hold_label(i, ycols) for i in range(len(ycols))])
    for title, sub in groups:
        r += 1
        line(r, title, cumulative(sub, ycols))
    r += 1
    heading(r, "當年平均值", [ysubs.get(k, "") or k for k in ycols])
    r += 1
    line(r, all_name, averages(rows, ycols))
    r += 1
    c = ws.cell(row=r, column=nlead, value=AVG_NOTE)
    c.font = Font(size=9, color=GREY)
    c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.merge_cells(start_row=r, start_column=nlead, end_row=r, end_column=last)
    ws.row_dimensions[r].height = 30


def write_sheet(ws, label, rows, ycols, ysubs, year, today):
    cols = [("产品名称", 34, "name"), ("产品种类", 14, "cat")]
    if len(set(r["kind"] for r in rows)) > 1:
        cols.append(("红利类型", 18, "kind"))
    heads = [c[0] for c in cols] + [(k.replace("保單年度", "保单年度") + "\n" + ysubs.get(k, "")).strip() for k in ycols]
    widths = [c[1] for c in cols] + [13] * len(ycols)
    separate = len(rows) > len(set(r["name"] for r in rows))
    tops = [(f"永明{label}（報告年度 {year}）", Font(bold=True, size=14, color=NAVY), 28),
            (DEFS[label], Font(size=10), 30),
            (LEGEND, Font(size=9, color=GREY), 18),
            (f"資料來源：香港永明金融網站（報告年度 {year}）｜更新日期：{today}", Font(size=9, color=GREY), 18)]
    for i, (s, font, height) in enumerate(tops, 1):
        c = ws.cell(row=i, column=1, value=s)
        c.font = font
        c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=len(heads))
        ws.row_dimensions[i].height = height
    for i, (h, w) in enumerate(zip(heads, widths), 1):
        head_cell(ws.cell(row=HEAD_ROW, column=i, value=h))
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[HEAD_ROW].height = 40
    ri, prev = HEAD_ROW, None
    for r in rows:
        if r["name"] != prev:
            if prev is not None and separate:
                ri += 1
                ws.cell(row=ri, column=1).fill = SEP_FILL
                ws.merge_cells(start_row=ri, start_column=1, end_row=ri, end_column=len(heads))
                ws.row_dimensions[ri].height = 9
            prev = r["name"]
        ri += 1
        for ci, (_, w, key) in enumerate(cols, 1):
            c = ws.cell(row=ri, column=ci, value=r[key])
            c.font = Font(bold=key == "name", size=10, color=DARK if key == "name" else "000000")
            c.alignment = Alignment(horizontal="left" if key == "name" else "center",
                                    vertical="center", wrap_text=True)
            c.border = BOX
        for j, k in enumerate(ycols):
            v = r["vals"].get(k, "/")
            c = ws.cell(row=ri, column=len(cols) + 1 + j, value=v)
            c.alignment = CENTER
            c.border = BOX
            if isinstance(v, float):
                c.number_format = pct_format(v)
                if v >= 0.9995:
                    c.fill = GREEN_FILL
                    c.font = Font(size=10, color="006100")
                elif v < 0.7995:
                    c.fill = RED_FILL
                    c.font = Font(size=10, bold=True, color="9C0006")
                else:
                    c.font = Font(size=10)
            else:
                c.font = Font(size=10, color=GREY)
        ws.row_dimensions[ri].height = ROW_HEIGHT
    ws.freeze_panes = f"B{HEAD_ROW + 1}"
    ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(len(heads))}{max(ri, HEAD_ROW + 1)}"
    write_averages(ws, ri + 2, label, rows, ycols, ysubs, len(cols))
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"


def write_overview(ws, sheets, year, today):
    ncol = 6
    ws.sheet_view.showGridLines = False
    for i, w in enumerate([30, 16, 12, 16, 16, 20], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def text(r, s, bold=False, size=10, height=20, color="000000"):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        c = ws.cell(row=r, column=1, value=s)
        c.font = Font(bold=bold, size=size, color=color)
        c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = height

    def body(r, i, v, fmt="0.00%"):
        c = ws.cell(row=r, column=i, value=v if v is not None else "/")
        c.border = BOX
        c.font = Font(size=10, bold=i == 1)
        c.alignment = Alignment(horizontal="left" if i == 1 else "center", vertical="center")
        if isinstance(v, float):
            c.number_format = fmt
        ws.row_dimensions[r].height = 22
        return c

    text(1, f"永明分紅表現概覽（報告年度 {year}）", True, 14, 28, NAVY)
    text(2, f"資料來源：香港永明金融網站｜更新日期：{today}", color=GREY)
    text(4, "資料概況", True, 11, 22, NAVY)
    r = 5
    for i, h in enumerate(["工作表", "產品數", "行數", "有數字的格數", "整體平均值", "低於 80% 的比例"], 1):
        head_cell(ws.cell(row=r, column=i, value=h))
    ws.row_dimensions[r].height = 30
    for label, rows in sheets:
        r += 1
        nums = numbers(rows)
        low = len([v for v in nums if v < 0.7995]) / len(nums) if nums else None
        name = f"{year}{label}"
        line = [name, len(set(x["name"] for x in rows)), len(rows), len(nums), mean(nums), low]
        for i, v in enumerate(line, 1):
            c = body(r, i, v, "0.0%" if i == 6 else "0.00%")
        link = ws.cell(row=r, column=1)
        link.hyperlink = Hyperlink(ref="", location=f"'{name}'!A1", display=name)
        link.value = name
        link.font = Font(size=10, bold=True, color="0563C1", underline="single")
    r += 2
    cats = list(dict.fromkeys(x["cat"] for _, rows in sheets for x in rows if x["cat"]))
    if cats:
        text(r, "按產品種類的整體平均值", True, 11, 22, NAVY)
        r += 1
        for i, h in enumerate(["產品種類"] + [label for label, _ in sheets], 1):
            head_cell(ws.cell(row=r, column=i, value=h))
        ws.row_dimensions[r].height = 30
        for cat in cats:
            r += 1
            line = [cat] + [mean(numbers([x for x in rows if x["cat"] == cat])) for _, rows in sheets]
            for i, v in enumerate(line, 1):
                body(r, i, v)
        r += 2
    text(r, "顏色說明", True, 11, 22, NAVY)
    samples = [(1.0, GREEN_FILL, "006100", "100% 或以上：與銷售時的說明一致或更好"),
               (0.79, RED_FILL, "9C0006", "低於 80%：明顯低於銷售時的說明"),
               ("/", None, GREY, "沒有數字：網站顯示「已停售」「尚未推出」或「不適用」")]
    for v, fill, color, desc in samples:
        r += 1
        c = ws.cell(row=r, column=1, value=v)
        c.font = Font(size=10, bold=True, color=color)
        c.alignment = CENTER
        c.border = BOX
        if fill:
            c.fill = fill
        if isinstance(v, float):
            c.number_format = "0%"
        d = ws.cell(row=r, column=2, value=desc)
        d.font = Font(size=10)
        d.alignment = Alignment(vertical="center")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncol)
        ws.row_dimensions[r].height = 22
    r += 2
    text(r, "使用提示", True, 11, 22, NAVY)
    tips = [
        "・點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
        f"・每個工作表第 {HEAD_ROW} 行有篩選箭嘴，可按產品種類或紅利類型篩選。",
        "・向右或向下捲動時，標題列和產品名稱會固定不動。",
        "・平均值摘要在每個工作表最底部。",
        "・手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。",
    ]
    for s in tips:
        r += 1
        text(r, s)


def notion_upload(path, year):
    token = os.environ.get("NOTION_TOKEN", "").strip()
    raw = os.environ.get("NOTION_DATABASE_ID", "").strip()
    if not token or not raw:
        print("未設定 NOTION_TOKEN 或 NOTION_DATABASE_ID，略過 Notion 上傳")
        return True
    m = re.search(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", raw)
    ids = [m.group(0)] if m else re.findall(r"[0-9a-fA-F]{32}", raw.split("?")[0])
    if not ids:
        print("Notion 上傳失敗：NOTION_DATABASE_ID 內找不到資料庫 ID")
        return False
    dbid = ids[-1]
    head = {"Authorization": f"Bearer {token}", "Notion-Version": "2025-09-03"}

    def api(step, method, path_, **kw):
        r = requests.request(method, "https://api.notion.com/v1/" + path_, headers=head, timeout=120, **kw)
        if r.status_code >= 300:
            raise RuntimeError(f"{step}：狀態 {r.status_code}｜{r.text[:200]}")
        return r.json()

    try:
        try:
            ds_id = api("讀取資料庫", "GET", f"databases/{dbid}")["data_sources"][0]["id"]
        except Exception:
            ds_id = dbid
        props = api("讀取資料來源", "GET", f"data_sources/{ds_id}")["properties"]
        title_prop = next(n for n, p in props.items() if p["type"] == "title")
        file_prop = next((n for n, p in props.items() if p["type"] == "files"), None)

        def page_title(p):
            return "".join(t.get("plain_text", "") for t in p["properties"][title_prop]["title"]).strip()

        found = api("尋找年度頁面", "POST", f"data_sources/{ds_id}/query",
                    json={"filter": {"property": title_prop, "title": {"starts_with": year}}})["results"]
        page = (next((p for p in found if page_title(p) == f"{year}年"), None)
                or next((p for p in found if page_title(p) == year), None)
                or (found[0] if found else None))
        if page:
            print(f"Notion：找到「{page_title(page)}」頁面")
        else:
            page = api("建立年度頁面", "POST", "pages",
                       json={"parent": {"type": "data_source_id", "data_source_id": ds_id},
                             "properties": {title_prop: {"title": [{"text": {"content": f"{year}年"}}]}}})
            print(f"Notion：找不到年度頁面，已建立「{year}年」")
        fname = os.path.basename(path)
        ascii_name = f"SunLife_{year}.xlsx"

        def is_auto(n):
            return n == fname or n.startswith("SunLife_")

        up = api("建立上傳", "POST", "file_uploads",
                 json={"mode": "single_part", "filename": fname, "content_type": XLSX})
        with open(path, "rb") as f:
            api("傳送檔案", "POST", f"file_uploads/{up['id']}/send", files={"file": (ascii_name, f, XLSX)})
        new = {"name": fname, "type": "file_upload", "file_upload": {"id": up["id"]}}
        if file_prop:
            keep = []
            for x in page.get("properties", {}).get(file_prop, {}).get("files", []):
                if is_auto(x.get("name", "")):
                    continue
                if x.get("type") == "file":
                    keep.append({"name": x["name"], "type": "file", "file": {"url": x["file"]["url"]}})
                elif x.get("type") == "external":
                    keep.append({"name": x["name"], "type": "external", "external": {"url": x["external"]["url"]}})
            done = api("更新檔案欄", "PATCH", f"pages/{page['id']}",
                       json={"properties": {file_prop: {"files": keep + [new]}}})
            count = len(done.get("properties", {}).get(file_prop, {}).get("files", []))
            print(f"Notion：已放入「{file_prop}」欄｜保留原有檔案 {len(keep)} 個｜現有檔案 {count} 個")
            if count < len(keep) + 1:
                print("注意：檔案數目比預期少，請到 Notion 檢查原有檔案")
        else:
            kids = api("讀取頁面內容", "GET", f"blocks/{page['id']}/children?page_size=100")["results"]
            removed = 0
            for b in kids:
                if b.get("type") == "file" and is_auto(b["file"].get("name", "")):
                    api("刪除舊檔", "DELETE", f"blocks/{b['id']}")
                    removed += 1
            api("加入檔案", "PATCH", f"blocks/{page['id']}/children",
                json={"children": [{"type": "file", "file": new}]})
            print(f"Notion：已把檔案放在頁面內容最底｜替換舊自動檔 {removed} 個｜其他內容不變")
        return True
    except Exception as e:
        print(f"Notion 上傳失敗｜{str(e)[:300]}")
        return False


def main():
    body = get(PAGE)
    log = {"untitled": [], "years": collections.Counter(), "tables": 0}
    sets, ycols, ysubs = parse(body, log)
    sheets = [(label, sets[label]) for label in LABELS]
    year = log["years"].most_common(1)[0][0] if log["years"] else str(datetime.date.today().year)
    today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d")

    def show(vals):
        return " ".join(f"{v:.2%}" if v is not None else "/" for v in vals)

    print("=== 總結 ===")
    print(f"資料表格：{log['tables']} 個｜報告年度：{dict(log['years'])}")
    print(f"保單年度欄 {len(ycols)} 個：", "、".join(f"{k}（{ysubs.get(k, '')}）" for k in ycols))
    for label, rows in sheets:
        kinds = dict(collections.Counter(r["kind"] or "—" for r in rows))
        print(f"{label}：產品 {len(set(r['name'] for r in rows))} 個｜行數 {len(rows)}"
              f"｜有數字格數 {len(numbers(rows))}｜類型 {kinds}")
        print(f"    當年平均值：{show(averages(rows, ycols))}")
        print(f"    持滿平均值：{show(cumulative(rows, ycols))}")
        if len(kinds) > 1:
            for k in kinds:
                print(f"    {k} 持滿平均值：{show(cumulative([x for x in rows if x['kind'] == k], ycols))}")
    print(f"無法識別標題的表格：{len(log['untitled'])} 個")
    for s in log["untitled"][:20]:
        print("   ", s)

    if not any(rows for _, rows in sheets):
        print("完全讀不到資料，沒有產生 Excel")
        sys.exit(1)
    wb = Workbook()
    write_overview(wb.active, sheets, year, today)
    wb.active.title = "概覽"
    for label, rows in sheets:
        write_sheet(wb.create_sheet(f"{year}{label}"), label, rows, ycols, ysubs, year, today)
    os.makedirs("output", exist_ok=True)
    path = f"output/永明總現金價值分紅實現率（{year}年）自動生成.xlsx"
    wb.save(path)
    print(f"已產生 {path}")
    if any(not rows for _, rows in sheets) or log["untitled"]:
        print("其中一類沒有資料，或有表格無法識別，不上傳 Notion")
        sys.exit(1)
    if not notion_upload(path, year):
        sys.exit(1)


main()
