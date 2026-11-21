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
from openpyxl.worksheet.properties import PageSetupProperties

PAGE = "https://www.sunlife.com.hk/zh-hant/insurance/fulfillment-ratios-of-respective-products/"
LABELS = ["總現金價值比率", "分紅實現率"]
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
HEAD_ROW = 3
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
    if m:
        return float(m.group(1)) / 100
    return s if s and s not in ("-", "–") else "/"


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
                sub = re.sub(r"\s+", "", ym.group(3)) + "生效"
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


def hold_label(i, key):
    if i == 0:
        return "所有年份平均值"
    n = key.replace("保單年度", "")
    return f"持滿 {n[:-1]} 年以上平均值" if n.endswith("+") else f"持滿 {n} 年平均值"


def head_cell(c):
    c.font = Font(bold=True, color="FFFFFF", size=10)
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = CENTER


def write_averages(ws, start, rows, ycols, nlead):
    last = nlead + len(ycols)
    kinds = list(dict.fromkeys(x["kind"] for x in rows))
    groups = [("所有分紅" if len(kinds) > 1 else "全部", rows)]
    if len(kinds) > 1:
        groups += [(k, [x for x in rows if x["kind"] == k]) for k in kinds]
    groups = [(t, s) for t, s in groups if numbers(s)]
    r = start
    c = ws.cell(row=r, column=1, value="平均值（固定數字，不隨篩選變化；只計算有百分比的格）")
    c.font = Font(bold=True, size=11, color=NAVY)
    c.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last)
    ws.row_dimensions[r].height = 24
    r += 1
    tables = (("持滿年期平均值（該年度及之後合併計算）", [hold_label(i, k) for i, k in enumerate(ycols)], cumulative),
              ("當年平均值（各保單年度）", ycols, averages))
    for heading, heads, fn in tables:
        head_cell(ws.cell(row=r, column=1, value=heading))
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=nlead)
        for j, h in enumerate(heads):
            head_cell(ws.cell(row=r, column=nlead + 1 + j, value=h))
        ws.row_dimensions[r].height = 46
        r += 1
        for title, sub in groups:
            c = ws.cell(row=r, column=1, value=title)
            c.font = Font(bold=True, size=10, color=DARK)
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border = BOX
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=nlead)
            for j, v in enumerate(fn(sub, ycols)):
                c = ws.cell(row=r, column=nlead + 1 + j, value=v if v is not None else "/")
                c.font = Font(bold=True, size=10)
                c.alignment = CENTER
                c.border = BOX
                if v is not None:
                    c.number_format = "0.00%"
            ws.row_dimensions[r].height = 24
            r += 1
        r += 1


def write_sheet(ws, title, rows, ycols, ysubs):
    cols = [("產品名稱", 34, "name"), ("產品種類", 14, "cat")]
    if len(set(r["kind"] for r in rows)) > 1:
        cols.append(("紅利類型", 18, "kind"))
    heads = [c[0] for c in cols] + [(k + "\n" + ysubs.get(k, "")).strip() for k in ycols]
    widths = [c[1] for c in cols] + [13] * len(ycols)
    separate = len(rows) > len(set(r["name"] for r in rows))
    c = ws.cell(row=1, column=1, value=title)
    c.font = Font(bold=True, size=14, color=NAVY)
    c.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(heads))
    ws.row_dimensions[1].height = 28
    c = ws.cell(row=2, column=1,
                value="100% 或以上綠底，低於 80% 紅底｜「已停售」「尚未推出」「不適用」為永明網站原文｜固定平均值在表格底部")
    c.font = Font(size=9, color=GREY)
    c.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(heads))
    ws.row_dimensions[2].height = 18
    for i, (h, w) in enumerate(zip(heads, widths), 1):
        head_cell(ws.cell(row=HEAD_ROW, column=i, value=h))
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[HEAD_ROW].height = 62
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
            v = r["vals"].get(k, "")
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
                c.font = Font(size=9, color=GREY)
        ws.row_dimensions[ri].height = ROW_HEIGHT
    ws.freeze_panes = f"B{HEAD_ROW + 1}"
    ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(len(heads))}{max(ri, HEAD_ROW + 1)}"
    write_averages(ws, ri + 2, rows, ycols, len(cols))
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"


def write_overview(ws, sheets, year, now):
    ncol = 6
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 38
    for i in range(2, ncol + 1):
        ws.column_dimensions[get_column_letter(i)].width = 17

    def text(r, s, bold=False, size=10, height=None, color="000000"):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        c = ws.cell(row=r, column=1, value=s)
        c.font = Font(bold=bold, size=size, color=color)
        c.alignment = Alignment(vertical="center", wrap_text=True)
        if height:
            ws.row_dimensions[r].height = height

    def table(r, heads, lines, fmts):
        for i, h in enumerate(heads, 1):
            head_cell(ws.cell(row=r, column=i, value=h))
        ws.row_dimensions[r].height = 30
        for line in lines:
            r += 1
            for i, v in enumerate(line, 1):
                c = ws.cell(row=r, column=i, value=v if v is not None else "/")
                c.border = BOX
                c.font = Font(size=10, bold=i == 1)
                c.alignment = Alignment(horizontal="left" if i == 1 else "center", vertical="center")
                if isinstance(v, float):
                    c.number_format = fmts.get(i, "0.00%")
            ws.row_dimensions[r].height = 22
        return r + 2

    def stat(title, sub):
        nums = numbers(sub)
        low = len([v for v in nums if v < 0.7995]) / len(nums) if nums else None
        return [title, len(set(x["name"] for x in sub)), len(sub), len(nums), mean(nums), low]

    text(1, f"永明分紅表現概覽（報告年度 {year}）", True, 14, 28, NAVY)
    text(2, f"自動生成｜擷取時間：{now}（香港時間）｜資料來源：香港永明金融網站")
    text(3, PAGE, color="0563C1")
    text(5, "資料概況與整體表現", True, 11, 22, NAVY)
    lines = []
    for label, rows in sheets:
        lines.append(stat(f"{year} {label}", rows))
        kinds = list(dict.fromkeys(x["kind"] for x in rows))
        if len(kinds) > 1:
            lines += [stat(f"　{k}", [x for x in rows if x["kind"] == k]) for k in kinds]
    r = table(6, ["指標項目", "產品數", "行數", "有數字格數", "整體平均值", "低於 80% 的比例"], lines, {6: "0.0%"})
    cats = list(dict.fromkeys(x["cat"] for _, rows in sheets for x in rows if x["cat"]))
    if cats:
        text(r, "按產品種類的整體平均值", True, 11, 22, NAVY)
        lines = [[cat] + [mean(numbers([x for x in rows if x["cat"] == cat])) for _, rows in sheets] for cat in cats]
        r = table(r + 1, ["產品種類"] + [label for label, _ in sheets], lines, {})
    text(r, "定義與顏色說明", True, 11, 22, NAVY)
    notes = [
        "總現金價值比率：歷年公佈之總現金價值，比對保單繕發時之預期總現金價值所計算的比率。",
        "分紅實現率：實際派發的非保證紅利，比對銷售時建議書上所示之金額所計算的比率，按周年紅利、歸原紅利、終期紅利/特別紅利分別列出。",
        "100% 或以上（綠色）：實際與投保時的預計一致或更高。低於 80%（紅色）：實際明顯低於投保時的預計。",
        "「已停售」「尚未推出」「不適用」為永明網站原文，表示該保單年度沒有相關數字。",
    ]
    for s in notes:
        r += 1
        text(r, s, height=32)
    r += 2
    text(r, "使用提示", True, 11, 22, NAVY)
    tips = [
        "兩張資料表第 3 行有篩選箭頭，可按產品種類或紅利類型篩選；產品名稱及標題列已固定。",
        "資料表底部的「持滿 N 年平均值」包含該保單年度及之後的所有數字；「當年平均值」只計算該保單年度一欄。",
        "平均值為固定數字，不隨篩選變化。",
        "過往表現並不能作為分紅保險產品未來表現的指引。",
    ]
    for s in tips:
        r += 1
        text(r, s, height=22)


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
    now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d %H:%M")

    def show(vals):
        return " ".join(f"{v:.2%}" if v is not None else "/" for v in vals)

    print("=== 總結 ===")
    print(f"資料表格：{log['tables']} 個｜報告年度：{dict(log['years'])}")
    print(f"保單年度欄 {len(ycols)} 個：", "、".join(ycols))
    for label, rows in sheets:
        kinds = dict(collections.Counter(r["kind"] or "—" for r in rows))
        print(f"{label}：產品 {len(set(r['name'] for r in rows))} 個｜行數 {len(rows)}"
              f"｜有數字格數 {len(numbers(rows))}｜類型 {kinds}")
        print(f"    當年平均值：{show(averages(rows, ycols))}")
        print(f"    持滿平均值：{show(cumulative(rows, ycols))}")
    print(f"無法識別標題的表格：{len(log['untitled'])} 個")
    for s in log["untitled"][:20]:
        print("   ", s)

    if not any(rows for _, rows in sheets):
        print("完全讀不到資料，沒有產生 Excel")
        sys.exit(1)
    wb = Workbook()
    write_overview(wb.active, sheets, year, now)
    wb.active.title = "概覽"
    for label, rows in sheets:
        write_sheet(wb.create_sheet(f"{year} {label}"), f"永明 {year} {label}", rows, ycols, ysubs)
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
