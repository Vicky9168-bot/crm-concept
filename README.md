# CRM 數據分析實戰（教學教材）

CRM 課程用的講義、簡報、互動 Dashboard 與練習資料。**所有數據皆為模擬資料**，僅供教學；品牌名稱只代表零售型態的風格參考（服飾快時尚、美妝生活通路、居家生活），不是這些品牌的真實數據。

## 內容

| 路徑 | 說明 |
|---|---|
| `output/reports/CRM_教學講義.docx` | 8 章講義：會員制度、生命週期、分群、自動化旅程、VIP 與年度規劃、管道與廠商、Data Mining、必備工具 |
| `output/reports/CRM_課程簡報.pptx` | 對應講義的 23 頁簡報（含講者備註） |
| `output/crm_dashboard.html` | 互動 Dashboard，雙擊即可開啟（圖表套件由 CDN 載入，需要網路） |
| `output/data/*.csv` | 三種型態的模擬資料：`*_customers.csv`（會員）、`*_orders.csv`（訂單）、`*_rfm.csv`（每位會員的 RFM 與分群） |
| `docs/CRM_學員練習檔.xlsx` | 學員練習用 Excel，A–D 各一個分頁，內含資料 |
| `tools/` | 產生資料、Dashboard、講義與練習檔的腳本 |
| `src/main/resources/config/` | 資料參數與教材內容設定（見下方「調整參數」） |
| `src/main/html/dashboard.template.html` | Dashboard 的 HTML 模板 |

## 事前準備

- Python 3，套件：`pandas`、`numpy`、`openpyxl`（以 Python 3.11 測試）
- Node.js，套件：`docx`（以 Node.js 25 測試）
- 講義字型為 Microsoft JhengHei（微軟正黑體）；沒有安裝時，Word 會自動換用其他中文字型，版面可能略有不同

```bash
pip install pandas numpy openpyxl
npm install docx
```

## 重新產生資料與 Dashboard

資料與 Dashboard 由同一支腳本產生，在專案根目錄執行：

```bash
python tools/generate_crm_data.py
```

會寫出：

- `output/data/<型態>_customers.csv`、`<型態>_orders.csv`、`<型態>_rfm.csv`（型態為 `apparel`、`drugstore`、`lifestyle`）
- `output/data/dashboard_data.json`：Dashboard 使用的彙總資料
- `output/crm_dashboard.html`：把資料內嵌進 `src/main/html/dashboard.template.html` 後的單一網頁

亂數種子固定為 42，因此在相同的參數下，每次產生的資料都一樣。

### 調整參數

`src/main/resources/config/brands.json` 控制每種型態的模擬參數：

| 欄位 | 意義 |
|---|---|
| `customers` | 會員人數 |
| `base_gap_days` | 平均回購間隔（天） |
| `aov_median`、`aov_floor`、`first_order_factor` | 客單價的中位數、下限與首購折減 |
| `channels`、`categories` | 通路與品類的占比 |
| `tiers` | 會員等級與近 12 個月消費門檻 |

改完後重新執行上面的指令即可。快照日、起訖日期與亂數種子在 `tools/generate_crm_data.py` 開頭的常數。

### Dashboard 使用方式

直接用瀏覽器開啟 `output/crm_dashboard.html`。上方可切換品牌型態與通路，下方六個頁籤：總覽 KPI、新客 / 舊客 / 沉睡客、RFM 分群與推播、留存 Cohort、會員等級模擬、品類回購。

## 重新產生講義

```bash
npm install docx          # 只需第一次
node tools/build_handout.js
```

會寫出 `output/reports/CRM_教學講義.docx`。

- 講義內容寫在 `tools/build_handout.js`；版面函式在 `tools/docx_helpers.js`
- 第 2.1 節「訊號 → 動作 → 再追蹤」的表格內容來自 `src/main/resources/config/signal_actions.json`，改這個檔案就能更新表格
- 講義提到的檔案路徑（例如 `output/data/*.csv`）以本專案的目錄結構為準

## 重新產生學員練習檔

```bash
python tools/build_practice_xlsx.py
```

會寫出 `docs/CRM_學員練習檔.xlsx`：依序為說明、A_Excel、B_SQL、C_Python、D_PowerBI 與三份資料分頁。題目文字來自 `src/main/resources/config/workbook_questions.json`。請先執行過資料產生腳本，練習檔會讀取 `output/data` 底下的 `drugstore_*.csv`。

## 注意

- 本專案只含對學員安全的內容；教師版練習檔、參考答案與其建置腳本未附。
- `tools/build_deck.js`（簡報建置腳本）會讀取 `output/data/answers.json`，該檔含參考答案、未附在此，因此簡報無法在本專案內重新建置，請直接使用現成的 `CRM_課程簡報.pptx`。
