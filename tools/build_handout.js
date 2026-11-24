// Build the CRM training handout (.docx). Run: node tools/build_handout.js
const fs = require("fs");
const path = require("path");
const { PageBreak, Paragraph, TextRun, AlignmentType } = require("docx");
const { FONT, p, h1, h2, bl, note, code, table, gap, makeDoc, save } = require("./docx_helpers");

const OUT = path.resolve(__dirname, "../output/reports/CRM_教學講義.docx");
const SIGNALS = JSON.parse(fs.readFileSync(path.resolve(__dirname, "../src/main/resources/config/signal_actions.json"), "utf-8"));

const c = [];
// ---------- Cover ----------
c.push(
  new Paragraph({ spacing: { before: 2400, after: 200 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "CRM 數據分析實戰講義", font: FONT, size: 52, bold: true, color: "1F4E79" })] }),
  p("從會員制度、分群、自動化到 Dashboard", { run: { size: 28, color: "555555" } }),
  p("配套檔案：output/crm_dashboard.html（互動 Dashboard）、output/data/*.csv（模擬資料）", { run: { size: 20, color: "777777" } }),
  p("資料聲明：講義中的 Dashboard 與數字皆為模擬資料，僅供教學。品牌會員制度為公開資訊整理，可能隨官方調整，使用前請以官方公告為準。", { run: { size: 20, color: "777777" } }),
  new Paragraph({ children: [new PageBreak()] }),
);

// ---------- 1 ----------
c.push(h1("1. 會員制度設計：用 incentive 拉高回購與忠誠度"));
c.push(h2("1.1 設計框架（五步）"));
[
  "目標：先決定要提升哪一個數字——回購率、客單價（AOV）、還是留存率（Retention Rate）。一套制度很難同時最大化三者。",
  "分層：用「近 12 個月消費金額」或「消費次數」分級，門檻應落在約 5–15% 會員可達的位置，讓會員看得到下一級。",
  "Incentive：點數、折價券、專屬日、生日禮、免運、優先權、體驗活動。依成本與感知價值排序，優先用「低成本、高感知」的權益。",
  "維持條件：升等容易、降等要有緩衝（例如寶雅金卡「消費滿額自動展期」的設計），避免會員剛升等就失去動機。",
  "KPI 與成本：每季檢視權益成本占營收比、各等級回購率與 AOV，不賺錢的權益要下架。",
].forEach((t) => c.push(bl(t)));
c.push(h2("1.2 核心公式"));
[
  "回購率 = 期間內購買 ≥2 次的會員數 ÷ 期間內購買會員數",
  "留存率（Retention Rate）= 期末仍有購買的原會員數 ÷ 期初會員數（用 Cohort 看，見第 7 章）",
  "AOV = 營收 ÷ 訂單數；客單提升 = 加購 × 搭配 × 門檻設計（例如滿額贈）",
  "回饋成本率 = 發放點數與券的折算成本 ÷ 營收（建議設上限，例如 3–5%，依毛利調整）",
].forEach((t) => c.push(bl(t)));
c.push(h2("1.3 市面會員制度案例"));
c.push(table([1700, 3600, 4338], ["品牌", "公開資訊重點", "值得學習 / 注意"], [
  ["寶雅 POYA", "普卡（App 免費申辦，效期永久）與金卡（年度消費滿 NT$10,000 升等，之後滿 NT$8,000 自動展期一年）。每月 5 號會員日：普卡點數 3 倍、金卡 5 倍；生日當月 5 號前消費享 5 倍並加贈 NT$250 折價券。點數 300 點折 1 元。2026 年金卡會員於資格期間內在線上購買任一筆訂單，即可領當月 POYA GIFT 兌換券。", "用固定「會員日」養成到店習慣；金卡權益綁定線上購買，是把門市客導向 App / 線上的 omni-channel 手法；展期門檻低於升等門檻，降低流失感。"],
  ["雅詩蘭黛 Estée Lauder", "全球版為階層式忠誠度計畫 Estée E-List：除購買外，評論、與線上美容顧問聊天、社群合作等互動行為也能累積點數與升級。（台灣官網的實際等級與點數細節本次未能查到，請自行到官網與專櫃確認。）", "把「互動行為」也算進點數，不只獎勵消費；高單價品牌用體驗與專屬服務取代折扣，保護品牌價格。"],
  ["MUJI 無印良品", "MUJI passport 整合線上與線下會員，提供週間會員優惠與不定期會員限定活動。（香港地區的公開資訊顯示每消費 $1 得 1 積分、生日月首次消費雙倍積分；台灣細節請另行確認。）", "整合線上線下身分，是 CRM 的第一步；週間優惠可用來平衡平日與假日客流。"],
  ["UNIQLO", "本次搜尋未找到台灣會員制度的可靠細節，故不列入比較。", "教學時可用 Dashboard 的「服飾快時尚型」資料練習，不代表該品牌實際做法。"],
]));
c.push(gap());
c.push(h2("1.4 如何規劃更完善的制度（建議範本）"));
c.push(table([1400, 2500, 2800, 2938], ["等級", "進入條件（示例）", "核心權益", "要驗證的指標"], [
  ["入門", "加入即享", "首購禮、生日禮、會員日點數加倍", "首購→二購轉換率"],
  ["中階", "12 個月消費達 P60 分位", "免運、專屬券（需滿額）、新品試用", "AOV 提升、回購週期縮短"],
  ["高階", "12 個月消費達 P90 分位", "專人服務、優先購買、邀請制活動、推薦禮", "留存率、推薦人數、客單"],
]));
c.push(note("設計原則：1) 以滿額門檻拉 AOV，而不是無條件打折；2) 提醒「再 NT$XXX 即升等 / 再一單即領券」；3) 先小範圍 A/B 測試，再全面上線。"));

// ---------- 2 ----------
c.push(h1("2. 新客、舊客、沉睡客：追蹤指標"));
c.push(p("定義要先寫清楚並固定，否則每月報表不可比。以下為教學範本，請依品類購買週期調整天數（美妝週期短、家居週期長）。"));
c.push(table([2100, 2600, 4938], ["族群", "定義（範例）", "追蹤指標（判斷 CRM 是否有效運作）"], [
  ["新客", "首購 ≤90 天且僅 1 單", "新客數、首購 AOV、首購渠道、30/60/90 天二購轉換率、首購禮核銷率"],
  ["活躍舊客", "90 天內有購買且 ≥2 單", "購買頻率、AOV、回購週期、各等級升等率、品類交叉購買率"],
  ["預警客", "91–180 天未購", "預警→活躍轉換率（喚醒成效）、點數即將到期人數"],
  ["沉睡客", "181–365 天未購", "喚醒活動回應率與回購率、召回單的毛利（避免只買折扣）"],
  ["流失客", ">365 天未購", "占會員比、歷史 LTV；轉為低成本觸及或抑制名單"],
]));
c.push(gap());
[
  "新舊客營收占比：舊客占比持續下滑是警訊，代表成長靠買廣告而非靠關係。",
  "月度狀態轉移矩陣：統計上月「預警」有多少人本月回到「活躍」，才是 CRM 專案的真正成績單。",
  "活躍率 = 近 180 天有購買會員 ÷ 全部有購買會員。Dashboard「總覽」與「新客 / 舊客 / 沉睡客」頁可直接練習。",
].forEach((t) => c.push(bl(t)));
c.push(h2("2.1 追蹤之後：訊號 → 動作 → 再追蹤"));
c.push(p("指標本身不會改變結果。要先設好「看到什麼訊號就觸發什麼動作」，再回頭看動作有沒有效，才會形成循環。"));
c.push(table([1200, 2500, 3838, 2100], ["族群", "看到的訊號", "後續動作", "之後再追蹤"], SIGNALS.rows));
c.push(gap());
c.push(p("重點是看「狀態轉移」而不是單一指標：上月預警的人本月有多少回到活躍（喚醒成效）？新客有多少人在 90 天內成為活躍舊客（首購旅程成效）？"));
SIGNALS.habits.forEach((t) => c.push(bl(t)));
c.push(note("進階延伸：口徑、基準線與旅程穩定之後，可用 Python 做流失預測（如邏輯迴歸或樹模型），在客人變成沉睡前就觸發。一開始不建議，因為定義未穩時模型結果很難使用。"));

// ---------- 3 ----------
c.push(h1("3. 顧客分群（Segmentation）與客製化推播"));
c.push(h2("3.1 常見分群方式"));
c.push(table([2000, 3300, 4338], ["方式", "切法", "適用情境"], [
  ["RFM", "最近購買（R）、頻率（F）、金額（M）各五等分", "最通用的起手式；Dashboard「RFM 分群」頁"],
  ["生命週期", "新客 / 活躍 / 預警 / 沉睡 / 流失", "自動化旅程觸發"],
  ["價值分群", "LTV、毛利貢獻、等級", "資源分配、VIP 規劃"],
  ["行為分群", "瀏覽、加購、棄單、點擊、使用 App", "即時觸發與個人化推薦"],
  ["品類偏好", "購買品類、價位帶、折扣敏感度", "商品推薦、避免對非折扣客亂發券"],
  ["人口與渠道", "性別、年齡、地區、門市、通路偏好", "活動選址、門市邀約、渠道選擇"],
]));
c.push(gap());
c.push(h2("3.2 RFM 分群後的客製化推播"));
c.push(table([2200, 7438], ["分群", "推播重點"], [
  ["Champions（高 R、F、M）", "新品搶先、邀請制活動、推薦朋友禮。不用折扣，給身分與體驗。"],
  ["Loyal", "升等進度提醒、加購組合、點數加倍日。"],
  ["New / Promising", "7–14 天內二購引導：搭配推薦、回購券、使用教學內容。"],
  ["Need Attention", "補貨提醒、限時提醒、App 推播。"],
  ["At Risk (VIP)", "過去高價值但久未回訪：專人聯繫、專屬關懷券；優先級最高。"],
  ["Hibernating", "品牌新聞 + 小額誘因；兩次無回應就降頻。"],
  ["Lost", "低成本 Email 最後一搏，之後轉為抑制名單，避免傷害寄送信譽。"],
]));
c.push(note("實務提醒：分群名稱只是標籤，重點是每個群要有「不同的訊息、誘因、頻率與渠道」，否則分群沒有意義。"));

// ---------- 4 ----------
c.push(h1("4. CRM 自動化旅程（Automation）"));
c.push(table([1800, 2300, 5538], ["族群", "觸發", "旅程內容與目標"], [
  ["新加入會員（未購）", "註冊後", "D0 歡迎信 + 品牌故事 + 首購禮；D3 熱銷 / 新手推薦；D7 首購禮提醒；D14 最後提醒。目標：首購轉換。"],
  ["首購會員", "首購後", "D1 感謝 + 使用方式；D7 搭配推薦；依品類回購週期前 7 天提醒補貨；D30 評價邀請。目標：第二次購買。"],
  ["回購 ≥2 次會員", "第 2 單後", "會員等級介紹與升等進度；生日 / 會員日；交叉品類推薦；推薦朋友獎勵。目標：提高頻率與 AOV，培養忠誠。"],
  ["棄單會員", "加入購物車未結帳", "1 小時內提醒；24 小時內附商品評價或庫存提示；72 小時才考慮小額誘因。目標：挽回而非先打折。"],
  ["沉睡會員", "超過 N 天未購", "喚醒信 → 專屬回饋 → 最後通知（點數即將到期）；無回應則降頻。"],
]));
c.push(gap());
c.push(bl("每個旅程都要設「退出條件」（例如已購買就停止），並設全域頻率上限，避免同一人一天收到多封。"));
c.push(bl("評估：每個節點的開信 / 點擊 / 轉換，以及有無對照組（不發送群）的增量，才能證明旅程有創造額外營收。"));

// ---------- 5 ----------
c.push(h1("5. 高單價會員體驗與年度規劃"));
c.push(h2("5.1 高單價會員的體驗活動"));
[
  "邀請制私享會：新品預覽、品牌顧問一對一諮詢、小班體驗課（要有名額感與專屬感）。",
  "服務型權益：優先預約、專人聯繫、免費保養 / 修改、生日專屬禮（不靠折扣，保護價格）。",
  "宣傳力道：推薦朋友雙邊獎勵、鼓勵 UGC 分享、VIP 試用心得；追蹤被推薦人數與其首購 AOV。",
  "衡量：活動後 90 天回訪率、客單變化、推薦人數，與沒受邀的相似 VIP 做對照。",
].forEach((t) => c.push(bl(t)));
c.push(h2("5.2 年度 CRM 規劃：怎麼嘗試才正確"));
c.push(table([1700, 3100, 4838], ["季度", "重點", "要驗證的假設"], [
  ["Q1 打底", "資料盤點、會員 ID 整合、分群與 KPI 基準線、啟動歡迎與首購旅程", "資料乾淨嗎？基準回購率是多少？"],
  ["Q2 試驗", "二購、棄單、沉睡喚醒；A/B 測試誘因與頻率", "哪種誘因增量最高、毛利最佳？"],
  ["Q3 擴大", "會員等級調整、VIP 活動、品類交叉推薦", "升等提醒是否拉高 AOV？"],
  ["Q4 衝刺與檢討", "檔期（雙 11、年末）、點數到期提醒、年度檢討", "全年 Cohort 留存與 LTV 是否提升？"],
]));
c.push(note("正確的嘗試方式：一次只改一個變數、設對照組、預先寫下成功指標與停損條件、至少跑滿一個完整回購週期再下結論。"));

// ---------- 6 ----------
c.push(h1("6. 溝通管道特色與市面廠商"));
c.push(h2("6.1 管道特色"));
c.push(table([1600, 3300, 4738], ["管道", "特色", "適合與注意"], [
  ["Email", "成本低、可放長內容與圖片、可自動化", "品牌故事、帳單 / 通知、喚醒；開信率受隱私設定影響，要看點擊與轉換"],
  ["簡訊 SMS", "到達率高、即時，但單價高、字數有限", "驗證碼、限時提醒、到貨 / 取貨；頻率過高易被反感"],
  ["LINE 官方帳號", "台灣滲透率高，可做分眾推播、圖文選單、會員整合", "活動、客服、會員日；要控制訊息量避免封鎖"],
  ["App 推播", "免費、即時、可深度個人化，需先有 App 與授權", "棄單、補貨、升等進度；通知授權率是關鍵"],
  ["社群（IG、Facebook）", "擴散與互動，可再行銷", "拉新與口碑；個人化弱，需搭配名單廣告"],
  ["電話 / 專人服務", "最溫暖、成本最高", "VIP 關懷、高單價諮詢；需排優先序"],
  ["門市 / 專櫃", "面對面，連結線上身分", "掃碼入會、取貨、體驗活動；要記錄門市行為回 CRM"],
  ["網站 / App 內訊息", "在使用情境中觸達，不受外部渠道限制", "個人化推薦、橫幅、升等進度條"],
]));
c.push(gap());
c.push(h2("6.2 市面常見 CRM / 行銷自動化廠商（清單，非推薦）"));
c.push(p("以下為市面上常被討論的工具類型與代表廠商。我未逐一核實各家最新功能、價格與台灣在地服務，採購前請向廠商索取 demo 與案例，並用自己的資料做 PoC。"));
c.push(table([2300, 3500, 3838], ["類型", "代表（舉例）", "評估重點"], [
  ["全球級 CRM / 行銷雲", "Salesforce、Adobe、HubSpot", "整合度、客製彈性、導入成本與人力"],
  ["跨渠道互動 / 旅程", "Braze、MoEngage、Klaviyo、Emarsys", "即時觸發、分群能力、App / Email 整合"],
  ["CDP 與資料整合", "Segment、Treasure Data", "身分合併、資料串接、即時性"],
  ["台灣 / 亞洲電商與會員", "91APP、SHOPLINE、CYBERBIZ、Appier", "與台灣電商、金流、LINE 的整合、在地支援"],
  ["LINE 相關工具", "LINE 官方帳號 + 第三方 CRM 擴充", "會員綁定、分眾標籤、訊息成本"],
]));

// ---------- 7 ----------
c.push(h1("7. Data Mining 與互動 Dashboard"));
c.push(p("本教材的 Dashboard 使用模擬資料，風格參考三種零售型態：服飾快時尚（UNIQLO 風格）、美妝生活通路（寶雅風格）、居家生活（MUJI 風格）。這些品牌的實際會員資料並未公開，因此不能直接「用線上資料分析」，只能用公開資訊設定情境與參數。"));
c.push(h2("7.1 Data Mining 流程"));
[
  "1. 問題定義：想知道「誰會回購」「誰快流失」「哪些商品帶動二購」。",
  "2. 資料整理：會員表、訂單表、商品表、渠道與活動表，以 customer_id 串接；處理重複、退貨與測試單。",
  "3. 特徵：R/F/M、AOV、回購間隔、品類、渠道、折扣依賴度。",
  "4. 分析：RFM 分群、Cohort 留存、籃分析（Basket / Association）、流失預測（如邏輯迴歸 / 樹模型）。",
  "5. 行動：把分群結果輸出成名單，進自動化旅程，並用對照組衡量增量。",
].forEach((t) => c.push(p(t)));
c.push(h2("7.2 Dashboard 使用方式"));
c.push(table([2300, 7338], ["頁籤", "學員練習"], [
  ["總覽 KPI", "切換品牌與通路，比較回購率、AOV、活躍率，說明三種型態的差異。"],
  ["新客 / 舊客 / 沉睡客", "找出預警與沉睡的人數占比，設計對應喚醒活動。"],
  ["RFM 分群與推播", "點分群看推播建議，找出營收占比最高與最該挽回的群。"],
  ["留存 Cohort", "比較不同首購月與不同品牌的留存曲線。"],
  ["會員等級模擬", "調整門檻與 AOV 假設，看「差一點升等」會員的增量潛力。"],
  ["品類回購", "用回購率 × 客單的泡泡圖，選出引流品與升級品。"],
]));
c.push(note("限制：模擬資料的品類為隨機抽樣，品類間差異不代表真實商品表現。Dashboard 的等級模擬使用累計消費，真實制度應使用滾動 12 個月。"));
c.push(h2("7.3 重新產生資料"));
c.push(...code("python tools/generate_crm_data.py\n# 輸出：output/data/*.csv、output/crm_dashboard.html\n# 參數：src/main/resources/config/brands.json（客數、客單、回購週期、等級門檻）"));
c.push(gap());

// ---------- 8 ----------
c.push(h1("8. CRM 人必備的工具：Excel、Power BI、SQL、Python"));
c.push(table([1500, 4200, 3938], ["工具", "基本必會", "典型 CRM 用途"], [
  ["Excel", "樞紐分析表、VLOOKUP / XLOOKUP、SUMIFS / COUNTIFS、條件式格式、日期函數（DATEDIF）", "快速 RFM、活動成效表、名單清洗"],
  ["Power BI", "資料模型（星狀）、DAX 基本量值、時間智慧、篩選器與鑽取", "月報 Dashboard、Cohort 熱力圖"],
  ["SQL", "SELECT / JOIN / GROUP BY、視窗函數、CTE、日期處理", "從資料庫撈會員與訂單、建立 RFM 與名單"],
  ["Python", "pandas、matplotlib、scikit-learn 基礎、Jupyter", "自動化報表、分群與流失預測、A/B 測試顯著性"],
]));
c.push(gap());
c.push(h2("8.1 SQL 範例：RFM 基礎表"));
c.push(...code(`WITH base AS (
  SELECT customer_id,
         MAX(order_date)        AS last_order,
         COUNT(DISTINCT order_id) AS frequency,
         SUM(amount)            AS monetary
  FROM orders
  GROUP BY customer_id
)
SELECT customer_id,
       DATEDIFF(day, last_order, GETDATE()) AS recency,
       frequency, monetary,
       NTILE(5) OVER (ORDER BY last_order)  AS r_score,
       NTILE(5) OVER (ORDER BY frequency)   AS f_score,
       NTILE(5) OVER (ORDER BY monetary)    AS m_score
FROM base;`));
c.push(note("DATEDIFF / GETDATE 為 SQL Server 語法；PostgreSQL 與 MySQL 的日期函數略有不同。"));
c.push(h2("8.2 Python 範例：Cohort 留存"));
c.push(...code(`import pandas as pd
o = pd.read_csv("output/data/drugstore_orders.csv", parse_dates=["order_date"])
o["first"]  = o.groupby("customer_id")["order_date"].transform("min")
o["cohort"] = o["first"].dt.to_period("M")
o["idx"]    = (o["order_date"].dt.year - o["first"].dt.year) * 12 \\
            + (o["order_date"].dt.month - o["first"].dt.month)
pivot = o.pivot_table(index="cohort", columns="idx",
                      values="customer_id", aggfunc="nunique")
retention = pivot.div(pivot[0], axis=0).round(3)   # 每列除以 M0 人數`));
c.push(h2("8.3 Excel 練習"));
c.push(bl("用 output/data/*_rfm.csv 開樞紐分析：列＝segment，值＝人數與 monetary 加總，計算各群營收占比。"));
c.push(bl("用 COUNTIFS 算 180 天活躍率；用 SUMIFS 算各等級的 AOV。"));
c.push(h2("8.4 Power BI 練習"));
c.push(bl("匯入三個 CSV，建立 customers 與 orders 的關聯（customer_id），用 DAX 寫回購率：DIVIDE(客戶訂單數≥2 的人數, 購買人數)。"));
c.push(bl("用 Matrix 視覺效果做 Cohort 熱力圖，加入品牌 / 通路切片器。"));

save(makeDoc("CRM 數據分析實戰講義", c), OUT);
