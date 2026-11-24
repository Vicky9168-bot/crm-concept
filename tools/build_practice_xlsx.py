"""Build the student practice workbook (.xlsx): one sheet per exercise part (A-D)
plus the three simulated CSVs as data sheets. Contains NO answers.

Questions come from src/main/resources/config/workbook_questions.json, the same
source used by tools/build_workbook.js, so the Word and Excel versions stay in sync.

Usage:
    python tools/build_practice_xlsx.py             # student version
    python tools/build_practice_xlsx.py --teacher   # A: answers + auto-check; B/C/D: reference answers
Output:
    docs/CRM_學員練習檔.xlsx  |  docs/CRM_教師版練習檔.xlsx
"""
import json
import math
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "output" / "data"
OUT = ROOT / "docs" / "CRM_學員練習檔.xlsx"
OUT_TEACHER = ROOT / "docs" / "CRM_教師版練習檔.xlsx"
ANSWERS = DATA / "answers.json"
REF_FILE = ROOT / "src/main/resources/config/reference_answers.json"
Q = json.loads((ROOT / "src/main/resources/config/workbook_questions.json").read_text(encoding="utf-8"))
BRAND = "drugstore"

FONT = "Arial"
F = lambda **k: Font(name=FONT, size=k.pop("size", 10), **k)
INPUT = PatternFill("solid", fgColor="FFFF99")      # yellow = student fills in
HEAD = PatternFill("solid", fgColor="1F4E79")
GREY = PatternFill("solid", fgColor="F2F4F7")
THIN = Side(style="thin", color="BBBBBB")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")


def put(ws, ref, value, font=None, fill=None, fmt=None, wrap=True, border=True):
    c = ws[ref]
    c.value = value
    c.font = font or F()
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if wrap:
        c.alignment = WRAP
    if border:
        c.border = BOX
    return c


def header(ws, row, labels, col=1):
    for i, t in enumerate(labels):
        put(ws, f"{get_column_letter(col + i)}{row}", t, font=F(bold=True, color="FFFFFF"), fill=HEAD)


def height(text, width_chars):
    """Rough row height for wrapped CJK text (CJK glyph ~2 width units)."""
    lines = sum(max(1, math.ceil(len(seg) * 1.9 / width_chars)) for seg in str(text).split("\n"))
    return max(18, 14 * lines + 4)


def title(ws, text, sub):
    ws["A1"].value = text
    ws["A1"].font = F(bold=True, size=14, color="1F4E79")
    ws["A2"].value = sub
    ws["A2"].font = F(color="666666")
    ws.sheet_view.showGridLines = False


def legend(ws, row):
    put(ws, f"A{row}", "", fill=INPUT)
    put(ws, f"B{row}", "黃色格 = 作答區，請在這裡填答案或公式；灰色斜體列 = 格式範例（非正式答案）", wrap=False, border=False)


# ---------------------------------------------------------------- 說明
def sheet_readme(wb):
    ws = wb.active
    ws.title = "說明"
    title(ws, "CRM 數據分析 學員練習檔", "全部為模擬資料（美妝生活通路型），不是任何品牌的真實數據")
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 90
    r = 4
    header(ws, r, ["分頁", "用途"])
    rows = [
        ("A_Excel", "A 部分：直接在這個檔案內用 data_rfm 做樞紐分析、COUNTIFS、SUMIFS，把答案填進黃色格。"),
        ("B_SQL", "B 部分：SQL 題。把 data_orders 匯入資料庫（表名 orders），在黃色格貼上你的 SQL 與結果。"),
        ("C_Python", "C 部分：Python / pandas 題。讀取 CSV（或由本檔匯出），貼上程式與結果。"),
        ("D_PowerBI", "D 部分：Power BI 題。匯入 data_orders 與 data_rfm，貼上 DAX 與核對數字。"),
        ("data_customers", f"會員主檔，{{n_c}} 列：customer_id、join_date、channel"),
        ("data_orders", "訂單明細，{n_o} 列：order_id、customer_id、order_date、amount、category、channel"),
        ("data_rfm", "每位會員一列，{n_c} 列；A 部分的主要資料，欄位見下表"),
    ]
    nc = len(pd.read_csv(DATA / f"{BRAND}_customers.csv"))
    no = len(pd.read_csv(DATA / f"{BRAND}_orders.csv"))
    for i, (a, b) in enumerate(rows, start=r + 1):
        put(ws, f"A{i}", a, font=F(bold=True))
        put(ws, f"B{i}", b.replace("{n_c}", f"{nc:,}").replace("{n_o}", f"{no:,}"))
        ws.row_dimensions[i].height = height(b, 90)
    r = r + len(rows) + 2
    put(ws, f"A{r}", "data_rfm 欄位", font=F(bold=True), border=False)
    r += 1
    header(ws, r, ["欄位（欄位字母）", "說明"])
    cols = [
        ("A customer_id", "會員編號"), ("B join_date", "加入日期"), ("C channel", "會員註冊渠道：App / Store / Web"),
        ("D recency", "距快照日 2026-10-31 的天數（越小越近）"), ("E frequency", "累計購買次數"),
        ("F monetary", "累計消費金額（NT$）"), ("G first_order", "首購日期"), ("H aov", "該會員平均客單（四捨五入）"),
        ("I R", "Recency 分數 1–5（5 最佳）"), ("J F", "Frequency 分數 1–5"), ("K M", "Monetary 分數 1–5"),
        ("L segment", "RFM 分群"), ("M lifecycle", "生命週期階段"),
        ("N ttm_spend", "近 365 天消費金額"), ("O tier", "會員等級：普卡 / 金卡（近 365 天消費 ≥ 10,000 為金卡）"),
    ]
    for a, b in cols:
        r += 1
        put(ws, f"A{r}", a)
        put(ws, f"B{r}", b)
    r += 2
    put(ws, f"A{r}", "使用提醒", font=F(bold=True), border=False)
    tips = [
        "快照日為 2026-10-31，所有「距今天數」都以這一天計算。",
        "回購率（約 82%）遠高於真實零售常見水準，各通路差異也極小，是模擬資料的特性；請把重點放在方法，不要把數字當成產業基準。",
        "請先複製本檔再作答，保留原檔供下一位使用。data_* 分頁請勿修改或排序，以免與題目口徑不一致。",
    ]
    for t in tips:
        r += 1
        put(ws, f"A{r}", "•", border=False)
        put(ws, f"B{r}", t, border=False)
        ws.row_dimensions[r].height = height(t, 90)
    legend(ws, r + 2)


# ---------------------------------------------------------------- A
# Teacher mode puts reference answers in G:I and auto-check formulas in K:M,
# mirroring the student input cells in B:D. Tolerances live on the 教師說明 sheet.
TOL_ROWS = {"int": 10, "aov": 11, "pct": 12, "gap": 13}
TOL_VALUES = {"int": 0, "aov": 1, "pct": 0.0005, "gap": 0.002}
TOL_LABELS = {
    "int": "人數 / 金額 / 會員數（整數）",
    "aov": "AOV（取整數的差距）",
    "pct": "比例（活躍率、回購率、營收占比；0.0005 = 0.05 個百分點）",
    "gap": "金卡 AOV 比普卡高 %（因參考值由取整後 AOV 推得，容許較大）",
}
BLUE = F(color="0000FF")


def sheet_a(wb, teacher=False, A=None):
    ws = wb.create_sheet("A_Excel")
    title(ws, "A. Excel 練習（使用 data_rfm 分頁）" + ("【教師版】" if teacher else ""),
          "請在本分頁黃色格作答；可在本檔內新增樞紐分析表分頁")
    for col, w in zip("ABCDE", [26, 18, 18, 18, 40]):
        ws.column_dimensions[col].width = w
    if teacher:
        for col, w in zip("FGHIJKLM", [3, 16, 16, 16, 3, 12, 12, 12]):
            ws.column_dimensions[col].width = w
    legend(ws, 3)
    put(ws, "A4", "欄位：A customer_id、B join_date、C channel、D recency、E frequency、F monetary、G first_order、H aov、I R、J F、K M、L segment、M lifecycle、N ttm_spend、O tier", wrap=False, border=False, font=F(color="666666"))
    ex = lambda ws, r, vals, fmts: [put(ws, f"{get_column_letter(i + 1)}{r}", v, font=F(italic=True, color="777777"), fill=GREY, fmt=fm) for i, (v, fm) in enumerate(zip(vals, fmts))]
    n_checks = [0]

    def label(r, first=True):
        """Teacher column headers on a table header row."""
        if teacher:
            header(ws, r, ["參考答案"] * 3, col=7)
            header(ws, r, ["核對"] * 3, col=11)

    def chk(r, col, ref, kind, fmt=None):
        """Reference answer (blue) + auto-check formula for the student cell in `col`."""
        if not teacher:
            return
        ci = ord(col) - 64
        rc, kc = get_column_letter(ci + 5), get_column_letter(ci + 9)
        put(ws, f"{rc}{r}", ref, font=BLUE, fmt=fmt)
        if kind == "text":
            f = f'=IF({col}{r}="","",IF(TRIM({col}{r})={rc}{r},"✓ 正確","✗ 再檢查"))'
        else:
            tol = f"教師說明!$C${TOL_ROWS[kind]}"
            f = f'=IF({col}{r}="","",IFERROR(IF(ABS({col}{r}-{rc}{r})<={tol},"✓ 正確","✗ 再檢查"),"✗ 再檢查"))'
        put(ws, f"{kc}{r}", f)
        n_checks[0] += 1

    r = 6
    # A1
    put(ws, f"A{r}", f"A1　{Q['A1']}", font=F(bold=True), border=False)
    ws.merge_cells(f"A{r}:E{r}")
    ws.row_dimensions[r].height = height(Q["A1"], 110)
    r += 1
    header(ws, r, ["segment", "人數", "累計消費 (NT$)", "營收占比"])
    label(r)
    r += 1
    ex(ws, r, ["範例（非答案）", 120, 85000, 0.125], [None, "#,##0", "#,##0", "0.0%"])
    segs = ["At Risk (VIP)", "Champions", "Hibernating", "Lost", "Loyal", "Need Attention", "New / Promising"]
    e1 = {row[0]: row for row in A["E1"]} if teacher else {}
    for sname in segs:
        r += 1
        put(ws, f"A{r}", sname)
        put(ws, f"B{r}", None, fill=INPUT, fmt="#,##0")
        put(ws, f"C{r}", None, fill=INPUT, fmt="#,##0")
        put(ws, f"D{r}", None, fill=INPUT, fmt="0.0%")
        if teacher:
            _, cnt, rev, share = e1[sname]
            chk(r, "B", cnt, "int", "#,##0")
            chk(r, "C", rev, "int", "#,##0")
            chk(r, "D", share / 100, "pct", "0.0%")
    r += 1
    put(ws, f"A{r}", "人數最多的群")
    put(ws, f"B{r}", None, fill=INPUT)
    if teacher:
        chk(r, "B", max(A["E1"], key=lambda x: x[1])[0], "text")
    r += 1
    put(ws, f"A{r}", "營收占比最高的群")
    put(ws, f"B{r}", None, fill=INPUT)
    if teacher:
        chk(r, "B", max(A["E1"], key=lambda x: x[2])[0], "text")

    # A2
    r += 2
    put(ws, f"A{r}", f"A2　{Q['A2']}", font=F(bold=True), border=False)
    ws.merge_cells(f"A{r}:E{r}")
    ws.row_dimensions[r].height = height(Q["A2"], 110)
    r += 1
    put(ws, f"A{r}", "180 天活躍率")
    put(ws, f"B{r}", None, fill=INPUT, fmt="0.0%")
    put(ws, f"C{r}", "範例：0.0%", font=F(italic=True, color="777777"), border=False)
    if teacher:
        chk(r, "B", A["E2"]["pct"] / 100, "pct", "0.0%")

    # A3
    r += 2
    put(ws, f"A{r}", f"A3　{Q['A3']}", font=F(bold=True), border=False)
    ws.merge_cells(f"A{r}:E{r}")
    ws.row_dimensions[r].height = height(Q["A3"], 110)
    r += 1
    header(ws, r, ["等級", "AOV (NT$)"])
    label(r)
    e3 = {row[0]: row[2] for row in A["E3"]} if teacher else {}
    for t in ["普卡", "金卡"]:
        r += 1
        put(ws, f"A{r}", t)
        put(ws, f"B{r}", None, fill=INPUT, fmt="#,##0")
        if teacher:
            chk(r, "B", e3[t], "aov", "#,##0")
    r += 1
    put(ws, f"A{r}", "金卡 AOV 比普卡高 (%)")
    put(ws, f"B{r}", None, fill=INPUT, fmt="0.0%")
    if teacher:
        chk(r, "B", round(e3["金卡"] / e3["普卡"] - 1, 4), "gap", "0.0%")

    # A4
    r += 2
    put(ws, f"A{r}", f"A4　{Q['A4']}", font=F(bold=True), border=False)
    ws.merge_cells(f"A{r}:E{r}")
    ws.row_dimensions[r].height = height(Q["A4"], 110)
    r += 1
    header(ws, r, ["channel", "會員數", "回購率", "AOV (NT$)"])
    label(r)
    e4 = {row[0]: row for row in A["E4"]} if teacher else {}
    for t in ["App", "Store", "Web"]:
        r += 1
        put(ws, f"A{r}", t)
        put(ws, f"B{r}", None, fill=INPUT, fmt="#,##0")
        put(ws, f"C{r}", None, fill=INPUT, fmt="0.0%")
        put(ws, f"D{r}", None, fill=INPUT, fmt="#,##0")
        if teacher:
            _, cnt, rep, aov = e4[t]
            chk(r, "B", cnt, "int", "#,##0")
            chk(r, "C", rep / 100, "pct", "0.0%")
            chk(r, "D", aov, "aov", "#,##0")
    r += 1
    put(ws, f"A{r}", "想一想：你的結論與理由")
    put(ws, f"B{r}", None, fill=INPUT)
    ws.merge_cells(f"B{r}:E{r}")
    ws.row_dimensions[r].height = 60
    if teacher:
        put(ws, f"G{r}", "見「教師說明」解讀要點", font=F(italic=True, color="777777"), border=False)
        last = r
        # score summary at the top right
        put(ws, "G2", "答對格數", font=F(bold=True), border=False)
        put(ws, "H2", f'=COUNTIF(K6:M{last},"✓*")', fmt="0")
        put(ws, "I2", "總格數", font=F(bold=True), border=False)
        put(ws, "J2", n_checks[0], font=BLUE, fmt="0")
        ws.column_dimensions["J"].width = 6
        put(ws, "K2", "得分率", font=F(bold=True), border=False)
        put(ws, "L2", "=IF(J2=0,0,H2/J2)", fmt="0%")


# ---------------------------------------------------------------- 教師說明
def sheet_teacher(wb, A):
    ws = wb.create_sheet("教師說明")
    title(ws, "教師說明（僅教師版）", "含 A–D 部分參考答案、A 部分自動核對與評分建議；請勿發給學員")
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 16
    put(ws, "A3", "使用方式", font=F(bold=True), border=False)
    how = [
        "複製本檔，把學員的答案貼進 A_Excel 分頁的黃色格（B:D 欄）；K:M 欄會即時顯示「✓ 正確」或「✗ 再檢查」，空白格不顯示。",
        "A_Excel 的 H2、L2 顯示答對格數與得分率；總格數 J2 為設計值（藍字），若增減題目格請同步修改。",
        f"參考答案（藍字）取自 {ANSWERS.relative_to(ROOT)}，由 tools/compute_answers.py 計算，SQL 與 pandas 已交叉比對一致。",
        "容許誤差見下表，可依需要調整；學員用取整後的值作答時，AOV 可能差 1 元以內。",
        "B、C、D 分頁的 E:G 欄為參考答案（結果）、參考 SQL / 程式碼 / DAX 與說明，只供人工比對，不自動核對。B 的 SQL 已在 SQLite 實際執行並與 pandas 比對一致；D 的 DAX 未在 Power BI 內實測，發放前請先跑過一遍。",
    ]
    for i, t in enumerate(how, start=4):
        put(ws, f"A{i}", "•", border=False)
        put(ws, f"B{i}", t, border=False)
        ws.row_dimensions[i].height = height(t, 70)
    header(ws, 9, ["類型", "適用說明", "容許誤差"])
    for k, row in TOL_ROWS.items():
        put(ws, f"A{row}", k)
        put(ws, f"B{row}", TOL_LABELS[k])
        put(ws, f"C{row}", TOL_VALUES[k], font=BLUE, fill=INPUT, fmt="General")
        ws.row_dimensions[row].height = height(TOL_LABELS[k], 70)
    r = 15
    put(ws, f"A{r}", "A4 解讀要點", font=F(bold=True), border=False)
    pts = [
        "三個通路的回購率與 AOV 差距都很小，且 Web 只有 %d 人，樣本太小，「不能下結論」是正確答案。" % next(x[1] for x in A["E4"] if x[0] == "Web"),
        "模擬資料中通路不影響行為；真實專案應先做顯著性檢定與樣本數檢查，再決定是否調整通路策略。",
        "加分：能指出「會員註冊渠道」與「訂單渠道」口徑不同（見 D3）。",
    ]
    for t in pts:
        r += 1
        put(ws, f"A{r}", "•", border=False)
        put(ws, f"B{r}", t, border=False)
        ws.row_dimensions[r].height = height(t, 70)
    r += 2
    put(ws, f"A{r}", "評分建議（A 部分共 20 分）", font=F(bold=True), border=False)
    r += 1
    header(ws, r, ["題號", "配分與重點", "分數"])
    for qid, txt, sc in [
        ("A1", "7 個分群的人數、累計消費、占比，以及人數最多 / 營收最高的群；每個正確格依比例給分", 8),
        ("A2", "180 天活躍率；公式口徑需為 recency ≤ 180 ÷ 全部會員", 3),
        ("A3", "兩個等級的 AOV 須以「累計消費 ÷ 累計次數」計算（非 aov 欄位平均），並算出金卡高出 %", 4),
        ("A4", "3 個通路的數字（3 分）＋「想一想」解讀（2 分）", 5),
    ]:
        r += 1
        put(ws, f"A{r}", qid)
        put(ws, f"B{r}", txt)
        put(ws, f"C{r}", sc, font=BLUE, fmt="0")
        ws.row_dimensions[r].height = height(txt, 70)
    r += 1
    put(ws, f"B{r}", "合計", font=F(bold=True))
    put(ws, f"C{r}", f"=SUM(C{r - 4}:C{r - 1})", fmt="0")
    r += 2
    put(ws, f"B{r}", "注意：自動核對只檢查數字；公式口徑（例如 AOV 是否用累計值）與「想一想」需教師人工判斷。", border=False, font=F(color="666666"))
    ws.row_dimensions[r].height = 30


# ---------------------------------------------------------------- B / C / D
def fmt_rows(rows, cols):
    """Render result rows as one line each: 'v1 | v2 | ...' (numbers get thousands separators)."""
    f = lambda v: f"{v:,}" if isinstance(v, int) else str(v)
    return "\n".join(" | ".join(f(v) for v in r[:cols]) for r in rows)


def reference(qid, A, REF):
    """(result text, code text, note text) for one B/C/D question, from answers.json + reference_answers.json."""
    code, notes = REF["code"], REF["notes"]
    if qid.startswith("B"):
        key = "S" + qid[1]
        res = {
            "B1": "month | revenue | orders | buyers | AOV\n" + fmt_rows(A["S1"], 5),
            "B2": f"{A['S2']}%",
            "B3": "customer_id | orders | monetary\n" + fmt_rows(A["S3"], 3),
            "B4": f"符合資格 {A['S4'][0]:,} 人；30 天內二購 {A['S4'][1]:,} 人；比例 {A['S4'][2]}%",
            "B5": "lifecycle | customers\n" + fmt_rows(A["S5"], 2),
            "B6": "category | AOV | 回購率 %\n" + fmt_rows(A["S6"], 3),
        }[qid]
        low = A["S6"][-1]
        note = {"B4": notes["B4"], "B6": f"回購率最低：{low[0]}（{low[2]}%）。" + notes["B6"]}.get(qid, "")
        return res, A["sql"][key], note
    if qid.startswith("C"):
        res = {
            "C1": f"cohort {A['P1']['size']} 人；M0–M3 留存率：" + "%、".join(str(v) for v in A["P1"]["ret"]) + "%",
            "C2": f"中位數 {A['P2']['median']:g} 天、平均 {A['P2']['mean']} 天（共 {A['P2']['n']:,} 個間隔）",
            "C3": "segment | 人數\n" + fmt_rows(list(A["P3"].items()), 2),
            "C4": f"金卡 {A['P4']['gold']} 人，占全部會員 {A['P4']['pct']}%（與 rfm.csv 的 tier=='金卡' 一致）",
        }[qid]
        note = {"C2": notes["C2"], "C3": notes["C3"], "C4": notes["C4"]}.get(qid, "")
        return res, code[qid], note
    o = A["B_overall"]
    res = {
        "D1": f"Buyers {o['buyers']:,}\nRevenue {o['revenue']:,}\nOrders {o['orders']:,}\nAOV {o['aov']:,}\nRepeat Rate {o['repeat']}%\nActive 180d {o['active180']}%",
        "D2": "Cohort = 2025-06 列，M0–M3：" + "%、".join(str(v) for v in A["P1"]["ret"]) + "%（與 C1 相同）",
        "D3": f"選 App：Buyers {A['B3']['buyers']:,}、Repeat Rate {A['B3']['repeat']}%、AOV {A['B3']['aov']:,}",
        "D4": "（無固定數字，依右欄評分重點）",
    }[qid]
    note = {"D2": notes["D2"], "D3": notes["D3"], "D4": notes["D4a"] + "\n" + notes["D4b"]}.get(qid, "")
    return res, code.get(qid, ""), note


def sheet_answers(wb, name, heading, sub, ids, col3, col4, intro=None, teacher=False, A=None, REF=None, code_label=""):
    ws = wb.create_sheet(name)
    title(ws, heading + ("【教師版】" if teacher else ""), sub)
    for col, w in zip("ABCD", [8, 48, 52, 34]):
        ws.column_dimensions[col].width = w
    if teacher:
        for col, w in zip("EFG", [40, 64, 46]):
            ws.column_dimensions[col].width = w
    legend(ws, 3)
    r = 4
    if intro:
        put(ws, f"A{r}", intro, wrap=False, border=False, font=F(color="666666"))
        r += 1
    r += 1
    header(ws, r, ["題號", "題目", col3, col4])
    if teacher:
        header(ws, r, ["參考答案（結果）", code_label, "說明 / 評分重點"], col=5)
    ws.freeze_panes = f"A{r + 1}"
    for qid in ids:
        r += 1
        put(ws, f"A{r}", qid, font=F(bold=True))
        put(ws, f"B{r}", Q[qid])
        put(ws, f"C{r}", None, fill=INPUT)
        put(ws, f"D{r}", None, fill=INPUT)
        h = max(120, height(Q[qid], 48))
        if teacher:
            res, code, note = reference(qid, A, REF)
            put(ws, f"E{r}", res, font=BLUE)
            put(ws, f"F{r}", code, font=Font(name="Consolas", size=9))
            put(ws, f"G{r}", note)
            h = max(h, 14 * (res.count("\n") + 1) + 8, 12 * (code.count("\n") + 1) + 10, height(note, 46))
        ws.row_dimensions[r].height = min(h, 409)


# ---------------------------------------------------------------- data
def sheet_data(wb, name, df, date_cols, widths):
    ws = wb.create_sheet(name)
    for j, col in enumerate(df.columns, start=1):
        c = ws.cell(row=1, column=j, value=col)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = HEAD
        ws.column_dimensions[get_column_letter(j)].width = widths.get(col, 14)
    for i, row in enumerate(df.itertuples(index=False), start=2):
        for j, v in enumerate(row, start=1):
            col = df.columns[j - 1]
            if col in date_cols:
                v = pd.Timestamp(v).to_pydatetime().date()
            elif hasattr(v, "item"):
                v = v.item()
            c = ws.cell(row=i, column=j, value=v)
            c.font = F()
            if col in date_cols:
                c.number_format = "yyyy-mm-dd"
            elif col in ("amount", "monetary", "ttm_spend", "aov"):
                c.number_format = "#,##0"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def main(teacher=False):
    A = json.loads(ANSWERS.read_text(encoding="utf-8")) if teacher else None
    REF = json.loads(REF_FILE.read_text(encoding="utf-8")) if teacher else None
    kw = dict(teacher=teacher, A=A, REF=REF)
    wb = Workbook()
    sheet_readme(wb)
    if teacher:
        sheet_teacher(wb, A)
    sheet_a(wb, teacher, A)
    sheet_answers(wb, "B_SQL", "B. SQL 練習（使用 orders 表）",
                  "把 data_orders 匯入資料庫，表名 orders；快照日 2026-10-31", ["B1", "B2", "B3", "B4", "B5", "B6"],
                  "你的 SQL", "你的答案（結果）",
                  "語法提醒：日期函數依資料庫而異（SQL Server 用 DATEDIFF，PostgreSQL 用日期相減）。",
                  code_label="參考 SQL（SQLite 語法，已實際執行）", **kw)
    sheet_answers(wb, "C_Python", "C. Python 練習（pandas）", "快照日 SNAP = pd.Timestamp('2026-10-31')",
                  ["C1", "C2", "C3", "C4"], "你的程式碼", "你的答案（結果）",
                  "起手：o = pd.read_csv(訂單 CSV, parse_dates=['order_date'])；rfm = pd.read_csv(RFM CSV)",
                  code_label="參考程式碼（pandas）", **kw)
    sheet_answers(wb, "D_PowerBI", "D. Power BI 練習",
                  "匯入 data_orders（表名 orders）與 data_rfm（表名 rfm），以 customer_id 建立 rfm 對 orders 的一對多關聯",
                  ["D1", "D2", "D3", "D4"], "你的 DAX / 作法", "核對數字或解讀",
                  code_label="參考 DAX（未在 Power BI 實測）", **kw)
    cu = pd.read_csv(DATA / f"{BRAND}_customers.csv")
    od = pd.read_csv(DATA / f"{BRAND}_orders.csv")
    rf = pd.read_csv(DATA / f"{BRAND}_rfm.csv")
    sheet_data(wb, "data_customers", cu, {"join_date"}, {"customer_id": 13, "join_date": 13})
    sheet_data(wb, "data_orders", od, {"order_date"}, {"order_id": 14, "customer_id": 13, "order_date": 13, "category": 14})
    sheet_data(wb, "data_rfm", rf, {"join_date", "first_order"},
               {"customer_id": 13, "join_date": 13, "first_order": 13, "segment": 16, "lifecycle": 22, "tier": 8})
    out = OUT_TEACHER if teacher else OUT
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print("wrote", out)


if __name__ == "__main__":
    main(teacher="--teacher" in sys.argv)
