// Build the CRM course deck (.pptx) from the handout content.
//   node tools/build_deck.js            -> output/reports/CRM_課程簡報.pptx
//   DECK_FONT="Heiti TC" DECK_OUT=/tmp/x.pptx node tools/build_deck.js   (render check only)
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const ROOT = path.resolve(__dirname, "..");
const OUT = process.env.DECK_OUT || path.join(ROOT, "output/reports/CRM_課程簡報.pptx");
const FONT = process.env.DECK_FONT || "Microsoft JhengHei";
const MONO = "Consolas";
const ANS = JSON.parse(fs.readFileSync(path.join(ROOT, "output/data/answers.json"), "utf-8"));
const SIGNALS = JSON.parse(fs.readFileSync(path.join(ROOT, "src/main/resources/config/signal_actions.json"), "utf-8"));

// palette: ink-navy dominant, warm sand content background, coral as the single sharp accent
const C = { dark: "10243E", mid: "1F4E79", teal: "2A9D8F", coral: "E4572E", sand: "F4F1EA", card: "FFFFFF", mute: "5B6776", line: "D9DEE5", gold: "E9B44C", code: "0B1A2E" };

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
pres.title = "CRM 數據分析實戰課程";
pres.author = "CRM Training";

const shadow = () => ({ type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12 });
let n = 0;

function base(tag, title, notes) {
  const s = pres.addSlide();
  n += 1;
  s.background = { color: C.sand };
  s.addText(tag, { x: 0.5, y: 0.28, w: 6, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: C.coral, margin: 0 });
  s.addText(title, { x: 0.5, y: 0.58, w: 9, h: 0.6, fontFace: FONT, fontSize: 26, bold: true, color: C.dark, margin: 0, valign: "middle" });
  s.addText("CRM 數據分析實戰", { x: 0.5, y: 5.25, w: 4, h: 0.25, fontFace: FONT, fontSize: 9, color: C.mute, margin: 0 });
  s.addText(String(n), { x: 8.5, y: 5.25, w: 1, h: 0.25, fontFace: FONT, fontSize: 9, color: C.mute, align: "right", margin: 0 });
  if (notes) s.addNotes(notes);
  return s;
}

// white card with a coral left edge (the deck's recurring motif)
function card(s, x, y, w, h, color = C.coral) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.card }, line: { color: C.line, width: 0.5 }, shadow: shadow() });
  s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.07, h, fill: { color }, line: { color, width: 0 } });
}
function cardText(s, x, y, w, h, head, body, opt = {}) {
  card(s, x, y, w, h, opt.color);
  s.addText(head, { x: x + 0.2, y: y + 0.1, w: w - 0.3, h: 0.35, fontFace: FONT, fontSize: opt.hs || 14, bold: true, color: C.dark, margin: 0, valign: "top" });
  s.addText(body, { x: x + 0.2, y: y + 0.5, w: w - 0.32, h: h - 0.58, fontFace: FONT, fontSize: opt.bs || 11, color: C.mute, margin: 0, valign: "top", paraSpaceAfter: 3 });
}
function circle(s, x, y, d, label, fill = C.mid, size = 14) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 } });
  s.addText(String(label), { x, y, w: d, h: d, fontFace: FONT, fontSize: size, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0 });
}
const th = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: C.mid }, fontFace: FONT, fontSize: 11, valign: "middle" } });
const td = (t, o = {}) => ({ text: t, options: { fontFace: FONT, fontSize: 10.5, color: C.dark, valign: "middle", ...o } });
const tableOpts = (x, y, colW) => ({ x, y, colW, border: { type: "solid", pt: 0.5, color: C.line }, fill: { color: C.card }, margin: [0.04, 0.08, 0.04, 0.08] });

// ============ 1 cover ============
{
  const s = pres.addSlide();
  n += 1;
  s.background = { color: C.dark };
  s.addText("CRM 數據分析實戰", { x: 0.6, y: 1.35, w: 5.6, h: 0.9, fontFace: FONT, fontSize: 38, bold: true, color: "FFFFFF", margin: 0 });
  s.addText("從會員制度、分群、自動化到 Dashboard", { x: 0.6, y: 2.3, w: 5.6, h: 0.5, fontFace: FONT, fontSize: 17, color: "CADCFC", margin: 0 });
  s.addText("課程內所有數字皆為模擬資料，品牌制度為公開資訊整理，請以官方公告為準。", { x: 0.6, y: 4.5, w: 5.4, h: 0.5, fontFace: FONT, fontSize: 10, color: "8FA3BF", margin: 0 });
  const tiles = [["8", "個主題章節"], ["4", "項必備工具"], ["3", "種零售型態模擬資料"]];
  tiles.forEach(([big, small], i) => {
    const y = 1.0 + i * 1.3;
    s.addShape(pres.shapes.RECTANGLE, { x: 6.7, y, w: 2.7, h: 1.1, fill: { color: "1B3556" }, line: { color: "2C4A73", width: 0.5 } });
    s.addShape(pres.shapes.RECTANGLE, { x: 6.7, y, w: 0.07, h: 1.1, fill: { color: C.coral }, line: { color: C.coral, width: 0 } });
    s.addText(big, { x: 6.95, y, w: 0.9, h: 1.1, fontFace: FONT, fontSize: 40, bold: true, color: "FFFFFF", valign: "middle", margin: 0 });
    s.addText(small, { x: 7.85, y, w: 1.5, h: 1.1, fontFace: FONT, fontSize: 12, color: "CADCFC", valign: "middle", margin: 0 });
  });
  s.addNotes("開場：說明本課程涵蓋 8 個主題，從制度設計到工具操作。強調所有 Dashboard 與練習資料皆為模擬資料。");
}

// ============ 2 course map ============
{
  const s = base("COURSE MAP", "課程地圖：8 個主題", "先帶過整體架構。1–5 是策略與方法，6 是管道與廠商，7–8 是資料與工具。");
  const items = [
    ["會員制度設計", "incentive、AOV、回購與留存"], ["新客・舊客・沉睡客", "生命週期與追蹤指標"],
    ["分群與客製推播", "RFM 與其他分群方式"], ["自動化旅程", "新客、首購、回購、棄單"],
    ["VIP 體驗與年度規劃", "高單價會員與全年節奏"], ["溝通管道與廠商", "管道特色、工具類型"],
    ["Data Mining 與 Dashboard", "流程與互動練習"], ["必備工具", "Excel、Power BI、SQL、Python"],
  ];
  items.forEach(([h, b], i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const x = 0.5 + col * 2.2, y = 1.5 + row * 1.85;
    card(s, x, y, 2.05, 1.65, i < 5 ? C.coral : C.teal);
    circle(s, x + 0.2, y + 0.15, 0.4, i + 1, C.mid, 13);
    s.addText(h, { x: x + 0.2, y: y + 0.62, w: 1.75, h: 0.52, fontFace: FONT, fontSize: 12, bold: true, color: C.dark, margin: 0, valign: "top" });
    s.addText(b, { x: x + 0.2, y: y + 1.17, w: 1.75, h: 0.45, fontFace: FONT, fontSize: 10.5, color: C.mute, margin: 0, valign: "top" });
  });
}

// ============ 3 five steps ============
{
  const s = base("01 會員制度設計", "制度設計五步：先定目標，再談權益", "一套制度很難同時最大化回購率、AOV、留存率，要先選一個主目標。分層門檻要讓會員看得到下一級。");
  const steps = [
    ["目標", "回購率、AOV、留存率三者難兼得，先選一個"],
    ["分層", "用近 12 個月消費分級，門檻約 5–15% 會員可達"],
    ["Incentive", "點數、券、專屬日、生日禮、免運、體驗；優先低成本高感知"],
    ["維持條件", "升等容易、降等有緩衝，避免剛升等就失去動機"],
    ["KPI 與成本", "每季看權益成本占比、各級回購率與 AOV，不賺錢的下架"],
  ];
  steps.forEach(([h, b], i) => {
    const x = 0.5 + i * 1.75;
    card(s, x, 1.55, 1.62, 3.1, i === 0 ? C.coral : C.mid);
    circle(s, x + 0.2, 1.75, 0.46, i + 1, i === 0 ? C.coral : C.mid, 15);
    s.addText(h, { x: x + 0.2, y: 2.4, w: 1.35, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: C.dark, margin: 0 });
    s.addText(b, { x: x + 0.2, y: 2.85, w: 1.35, h: 1.75, fontFace: FONT, fontSize: 12, color: C.mute, margin: 0, valign: "top" });
  });
}

// ============ 4 formulas ============
{
  const s = base("01 會員制度設計", "四個必會公式", "回饋成本率要設上限，依毛利調整。留存率用 Cohort 看，後面第 7 章會示範。");
  const f = [
    ["回購率", "期間內購買 ≥2 次的會員數 ÷ 購買會員數"],
    ["留存率", "期末仍有購買的原會員數 ÷ 期初會員數（用 Cohort 看）"],
    ["AOV 客單價", "營收 ÷ 訂單數；提升靠加購、搭配、滿額門檻設計"],
    ["回饋成本率", "點數與券的折算成本 ÷ 營收；建議設上限 3–5%，依毛利調整"],
  ];
  f.forEach(([h, b], i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.5 + Math.floor(i / 2) * 1.85;
    card(s, x, y, 4.4, 1.65, i === 3 ? C.gold : C.coral);
    s.addText(h, { x: x + 0.25, y: y + 0.15, w: 3.9, h: 0.5, fontFace: FONT, fontSize: 20, bold: true, color: C.mid, margin: 0 });
    s.addText(b, { x: x + 0.25, y: y + 0.75, w: 3.9, h: 0.8, fontFace: FONT, fontSize: 14, color: C.dark, margin: 0, valign: "top" });
  });
}

// ============ 5 POYA ============
{
  const s = base("01 會員制度設計 ・ 案例", "寶雅 POYA：用會員日與金卡養成習慣", "資料為 2026 年公開資訊整理，使用前請以官方公告為準。重點是展期門檻低於升等門檻，以及金卡權益綁定線上購買。");
  const stats = [
    ["NT$10,000", "金卡升等：年度消費門檻"], ["NT$8,000", "金卡展期：之後滿額自動展一年"],
    ["3 倍 / 5 倍", "每月 5 號會員日點數（普卡 / 金卡）"], ["300 點 = NT$1", "點數折抵換算"],
  ];
  stats.forEach(([big, small], i) => {
    const x = 0.5 + (i % 2) * 2.4, y = 1.5 + Math.floor(i / 2) * 1.6;
    card(s, x, y, 2.25, 1.45, i < 2 ? C.coral : C.teal);
    s.addText(big, { x: x + 0.2, y: y + 0.12, w: 1.95, h: 0.6, fontFace: FONT, fontSize: 19, bold: true, color: C.mid, margin: 0, valign: "middle" });
    s.addText(small, { x: x + 0.2, y: y + 0.78, w: 1.95, h: 0.6, fontFace: FONT, fontSize: 10.5, color: C.mute, margin: 0, valign: "top" });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.5, w: 4.1, h: 3.05, fill: { color: C.dark }, line: { color: C.dark, width: 0 } });
  s.addText("值得學習", { x: 5.65, y: 1.62, w: 3.6, h: 0.4, fontFace: FONT, fontSize: 15, bold: true, color: "FFFFFF", margin: 0 });
  s.addText([
    { text: "固定「會員日」養成到店習慣", options: { bullet: true, breakLine: true } },
    { text: "金卡權益綁定線上購買：把門市客導向 App / 線上", options: { bullet: true, breakLine: true } },
    { text: "展期門檻低於升等門檻，降低流失感", options: { bullet: true, breakLine: true } },
    { text: "生日當月 5 號前消費：5 倍點數 + NT$250 折價券", options: { bullet: true } },
  ], { x: 5.65, y: 2.1, w: 3.65, h: 2.35, fontFace: FONT, fontSize: 13, color: "DCE6F2", margin: 0, valign: "top", paraSpaceAfter: 10 });
  s.addText("普卡：App 免費申辦、效期永久。", { x: 0.5, y: 4.75, w: 8.5, h: 0.3, fontFace: FONT, fontSize: 10, color: C.mute, margin: 0 });
}

// ============ 6 other cases ============
{
  const s = base("01 會員制度設計 ・ 案例", "其他品牌：公開資訊有限，要標明來源範圍", "雅詩蘭黛台灣與 UNIQLO 台灣的細節本次沒查到，教學時要誠實說明，不要補猜。MUJI 的點數規則是香港資料。");
  const cols = [
    ["雅詩蘭黛 Estée Lauder", "全球版 Estée E-List 為階層制：評論、線上美容顧問聊天、社群合作也能累積點數與升級。台灣細節未查到，請自行確認。", "把互動行為算進點數；高單價用體驗取代折扣", C.coral],
    ["MUJI 無印良品", "MUJI passport 整合線上與線下會員，有週間會員優惠與限定活動。香港資料：每 $1 得 1 積分（台灣請另確認）。", "整合線上線下身分是 CRM 第一步", C.teal],
    ["UNIQLO", "本次搜尋未找到台灣會員制度的可靠細節，故不列入比較。", "改用 Dashboard 的服飾型模擬資料練習，不代表該品牌做法", C.mute],
  ];
  cols.forEach(([h, pub, learn, color], i) => {
    const x = 0.5 + i * 3.05;
    card(s, x, 1.5, 2.9, 3.4, color);
    s.addText(h, { x: x + 0.2, y: 1.6, w: 2.6, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: C.dark, margin: 0 });
    s.addText("公開資訊", { x: x + 0.2, y: 2.05, w: 2.6, h: 0.25, fontFace: FONT, fontSize: 10, bold: true, color: color, margin: 0 });
    s.addText(pub, { x: x + 0.2, y: 2.3, w: 2.6, h: 1.55, fontFace: FONT, fontSize: 12, color: C.dark, margin: 0, valign: "top" });
    s.addText("值得學習", { x: x + 0.2, y: 3.9, w: 2.6, h: 0.25, fontFace: FONT, fontSize: 10, bold: true, color: color, margin: 0 });
    s.addText(learn, { x: x + 0.2, y: 4.15, w: 2.6, h: 0.65, fontFace: FONT, fontSize: 11.5, color: C.mute, margin: 0, valign: "top" });
  });
}

// ============ 7 tier template ============
{
  const s = base("01 會員制度設計", "建議範本：三級制度，每級要有可驗證的指標", "門檻用分位數而不是固定金額，才能隨業績調整。設計原則三點：滿額門檻拉 AOV、提醒差多少升等、先小範圍 A/B 測試。");
  const tiers = [
    ["入門", "加入即享", "首購禮、生日禮、會員日點數加倍", "首購→二購轉換率", 1.7, C.teal],
    ["中階", "12 個月消費達 P60", "免運、專屬券（需滿額）、新品試用", "AOV 提升、回購週期縮短", 2.1, C.mid],
    ["高階", "12 個月消費達 P90", "專人服務、優先購買、邀請制活動、推薦禮", "留存率、推薦人數、客單", 2.5, C.dark],
  ];
  tiers.forEach(([h, cond, perk, kpi, hh, color], i) => {
    const x = 0.5 + i * 3.05, bottom = 4.0, y = bottom - hh;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.9, h: hh, fill: { color }, line: { color, width: 0 }, shadow: shadow() });
    s.addText(h, { x: x + 0.2, y: y + 0.1, w: 2.5, h: 0.4, fontFace: FONT, fontSize: 17, bold: true, color: "FFFFFF", margin: 0 });
    s.addText([
      { text: cond, options: { bold: true, breakLine: true } },
      { text: perk, options: { breakLine: true } },
      { text: "驗證：" + kpi, options: { color: "FFE3A3" } },
    ], { x: x + 0.2, y: y + 0.55, w: 2.5, h: hh - 0.65, fontFace: FONT, fontSize: 11.5, color: "FFFFFF", margin: 0, valign: "top", paraSpaceAfter: 5 });
  });
  card(s, 0.5, 4.2, 8.95, 0.85, C.coral);
  s.addText("設計原則　1) 用滿額門檻拉 AOV，不要無條件打折　2) 提醒「再 NT$XXX 即升等」　3) 先小範圍 A/B 測試",
    { x: 0.75, y: 4.2, w: 8.6, h: 0.85, fontFace: FONT, fontSize: 12, color: C.dark, margin: 0, valign: "middle" });
}

// ============ 8 lifecycle ============
{
  const s = base("02 新客・舊客・沉睡客", "生命週期五階段與追蹤指標", "天數是教學範本，要依品類購買週期調整：美妝短、家居長。定義要先寫清楚並固定，否則每月報表不可比。");
  const st = [
    ["新客", "首購 ≤90 天且僅 1 單", "新客數、首購 AOV、30/60/90 天二購轉換率", C.teal],
    ["活躍舊客", "90 天內有購買且 ≥2 單", "購買頻率、AOV、回購週期、升等率", C.mid],
    ["預警客", "91–180 天未購", "預警→活躍轉換率、點數即將到期人數", C.gold],
    ["沉睡客", "181–365 天未購", "喚醒回應率、召回單的毛利", C.coral],
    ["流失客", "超過 365 天未購", "占會員比、歷史 LTV；轉低成本觸及", C.mute],
  ];
  st.forEach(([h, d, m, color], i) => {
    const x = 0.5 + i * 1.8;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.5, w: 1.68, h: 0.55, fill: { color }, line: { color, width: 0 } });
    s.addText(h, { x, y: 1.5, w: 1.68, h: 0.55, fontFace: FONT, fontSize: 14, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0 });
    card(s, x, 2.05, 1.68, 2.2, color);
    s.addText(d, { x: x + 0.18, y: 2.15, w: 1.42, h: 0.75, fontFace: FONT, fontSize: 12, bold: true, color: C.dark, margin: 0, valign: "top" });
    s.addText(m, { x: x + 0.18, y: 2.95, w: 1.42, h: 1.25, fontFace: FONT, fontSize: 11, color: C.mute, margin: 0, valign: "top" });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.4, w: 8.95, h: 0.7, fill: { color: C.dark }, line: { color: C.dark, width: 0 } });
  s.addText("CRM 的真正成績單：上月「預警」有多少人，本月回到「活躍」？（月度狀態轉移矩陣）", { x: 0.7, y: 4.4, w: 8.6, h: 0.7, fontFace: FONT, fontSize: 12.5, bold: true, color: "FFFFFF", margin: 0, valign: "middle" });
}

// ============ 8b signal -> action -> track ============
{
  const s = base("02 新客・舊客・沉睡客", "追蹤之後：看到訊號，就觸發動作",
    "指標本身不會改變結果，要設好看到什麼訊號就觸發什麼動作，再回頭看動作有沒有效。完整的 13 條對照在講義第 2.1 節。口徑與天數要用自己的資料校正，並且每個動作都要留對照組。");
  const data = [[th("族群"), th("看到的訊號"), th("後續動作"), th("之後再追蹤")]].concat(SIGNALS.deck.map((r) => [
    td(r[0], { bold: true, color: C.mid }), td(r[1]), td(r[2]), td(r[3], { bold: true, color: C.coral }),
  ]));
  s.addTable(data, { ...tableOpts(0.5, 1.4, [1.3, 2.2, 3.65, 1.8]), rowH: 0.46 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.5, w: 8.95, h: 0.6, fill: { color: C.dark }, line: { color: C.dark, width: 0 } });
  s.addText("看「狀態轉移」，不只看單一指標　・　設警戒線再設觸發　・　每個動作留對照組",
    { x: 0.7, y: 4.5, w: 8.6, h: 0.6, fontFace: FONT, fontSize: 12.5, bold: true, color: "FFFFFF", margin: 0, valign: "middle" });
}

// ============ 9 segmentation methods ============
{
  const s = base("03 顧客分群", "六種分群方式：RFM 是最通用的起手式", "分群不是目的，重點是每個群要有不同的訊息、誘因、頻率與渠道。");
  const m = [
    ["RFM", "最近購買、頻率、金額，各五等分"], ["生命週期", "新客 / 活躍 / 預警 / 沉睡 / 流失"],
    ["價值分群", "LTV、毛利貢獻、等級，用於 VIP 規劃"], ["行為分群", "瀏覽、加購、棄單、點擊、使用 App"],
    ["品類偏好", "購買品類、價位帶、折扣敏感度"], ["人口與渠道", "性別、年齡、地區、門市、通路偏好"],
  ];
  m.forEach(([h, b], i) => {
    const x = 0.5 + (i % 3) * 3.05, y = 1.5 + Math.floor(i / 3) * 1.8;
    cardText(s, x, y, 2.9, 1.6, h, b, { color: i === 0 ? C.coral : C.mid, hs: 16, bs: 12.5 });
  });
}

// ============ 10 RFM playbook ============
{
  const s = base("03 顧客分群", "RFM 分群後：每一群用不同的推播", "At Risk (VIP) 是優先級最高的挽回對象。Champions 不用折扣，給身分與體驗。");
  const rows = [
    ["Champions", "新品搶先、邀請制活動、推薦朋友禮；不用折扣，給身分與體驗"],
    ["Loyal", "升等進度提醒、加購組合、點數加倍日"],
    ["New / Promising", "7–14 天內二購引導：搭配推薦、回購券、使用教學"],
    ["Need Attention", "補貨提醒、限時提醒、App 推播"],
    ["At Risk (VIP)", "過去高價值但久未回訪：專人聯繫、專屬關懷券，優先級最高"],
    ["Hibernating", "品牌新聞 + 小額誘因；兩次無回應就降頻"],
    ["Lost", "低成本 Email 最後一搏，之後轉抑制名單"],
  ];
  const data = [[th("分群"), th("推播重點")]].concat(rows.map(([a, b], i) => [
    td(a, { bold: true, color: i === 4 ? C.coral : C.mid }), td(b),
  ]));
  s.addTable(data, { ...tableOpts(0.5, 1.4, [2.1, 6.85]), rowH: 0.43 });
}

// ============ 11 journeys ============
{
  const s = base("04 自動化旅程", "五條旅程：不同族群，不同內容與目標", "旅程設計的核心是目標不同：新客求首購、首購客求二購、回購客求頻率與 AOV、棄單求挽回、沉睡求召回。");
  const rows = [
    // every journey cell is broken into two fixed lines so row height never depends on auto-wrapping
    ["新加入（未購）", "註冊後", "D0 歡迎 + 首購禮 → D3 熱銷推薦\n→ D7 提醒 → D14 最後提醒", "首購"],
    ["首購會員", "首購後", "D1 感謝 → D7 搭配推薦\n→ 補貨前 7 天提醒 → D30 評價邀請", "二購"],
    ["回購 ≥2 次", "第 2 單後", "等級與升等進度、生日 / 會員日\n交叉推薦、推薦朋友獎勵", "頻率與 AOV"],
    ["棄單會員", "購物車未結帳", "1 小時提醒 → 24 小時評價或庫存提示\n→ 72 小時才考慮小額誘因", "挽回"],
    ["沉睡會員", "超過 N 天未購", "喚醒信 → 專屬回饋 → 點數到期通知\n無回應則降頻", "召回"],
  ];
  // 10pt text, wide content column and a row height that still holds a wrapped 2-line cell,
  // so a longer font fallback in PowerPoint cannot grow a row into the footer area
  const data = [[th("族群"), th("觸發"), th("旅程內容"), th("目標")]].concat(rows.map((r) => [
    td(r[0], { bold: true, color: C.mid, fontSize: 10 }), td(r[1], { fontSize: 10 }), td(r[2], { fontSize: 10 }),
    td(r[3], { bold: true, color: C.coral, fontSize: 10 }),
  ]));
  s.addTable(data, { ...tableOpts(0.5, 1.4, [1.45, 1.35, 4.95, 1.2]), rowH: [0.42, 0.58, 0.58, 0.58, 0.58, 0.58] });
}

// ============ 12 journey rules ============
{
  const s = base("04 自動化旅程", "棄單挽回：先提醒，最後才給誘因", "挽回而非先打折。每個旅程都要有退出條件與全域頻率上限，並用對照組衡量增量。");
  const nodes = [["1 小時內", "提醒購物車還有商品"], ["24 小時內", "附商品評價或庫存提示"], ["72 小時", "才考慮小額誘因"]];
  s.addShape(pres.shapes.LINE, { x: 1.2, y: 2.1, w: 7.6, h: 0, line: { color: C.mid, width: 2, dashType: "dash" } });
  nodes.forEach(([h, b], i) => {
    const cx = 1.2 + i * 3.8;
    circle(s, cx - 0.25, 1.85, 0.5, i + 1, i === 2 ? C.coral : C.mid, 14);
    s.addText(h, { x: cx - 1.1, y: 2.45, w: 2.2, h: 0.35, fontFace: FONT, fontSize: 14, bold: true, color: C.dark, align: "center", margin: 0 });
    s.addText(b, { x: cx - 1.2, y: 2.8, w: 2.4, h: 0.5, fontFace: FONT, fontSize: 11, color: C.mute, align: "center", margin: 0, valign: "top" });
  });
  cardText(s, 0.5, 3.55, 4.4, 1.35, "退出條件與頻率上限", "已購買就停止旅程；設全域頻率上限，避免同一人一天收到多封。", { bs: 11.5 });
  cardText(s, 5.05, 3.55, 4.4, 1.35, "用對照組證明增量", "每個節點看開信、點擊、轉換；沒有不發送群的對照，就不能證明旅程創造額外營收。", { bs: 11.5, color: C.teal });
}

// ============ 13 VIP ============
{
  const s = base("05 高單價會員", "VIP 體驗：用專屬感取代折扣", "高單價會員要保護價格，用服務與體驗拉客單、回訪與口碑。活動成效要和沒受邀的相似 VIP 做對照。");
  const v = [
    ["邀請制私享會", "新品預覽、品牌顧問一對一、小班體驗課；要有名額感與專屬感"],
    ["服務型權益", "優先預約、專人聯繫、免費保養 / 修改、生日專屬禮；不靠折扣"],
    ["宣傳力道", "推薦朋友雙邊獎勵、鼓勵 UGC、VIP 試用心得；追蹤被推薦人數與首購 AOV"],
    ["如何衡量", "活動後 90 天回訪率、客單變化、推薦人數；與沒受邀的相似 VIP 對照"],
  ];
  v.forEach(([h, b], i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.5 + Math.floor(i / 2) * 1.85;
    cardText(s, x, y, 4.4, 1.65, h, b, { hs: 16, bs: 13, color: i === 3 ? C.teal : C.coral });
  });
}

// ============ 14 annual plan ============
{
  const s = base("05 年度規劃", "年度 CRM：一季一個問題，逐步驗證", "正確的嘗試方式：一次只改一個變數、設對照組、預先寫下成功指標與停損條件、至少跑滿一個完整回購週期再下結論。");
  const q = [
    ["Q1 打底", "資料盤點、會員 ID 整合、分群與 KPI 基準線、啟動歡迎與首購旅程", "資料乾淨嗎？基準回購率多少？"],
    ["Q2 試驗", "二購、棄單、沉睡喚醒；A/B 測試誘因與頻率", "哪種誘因增量最高、毛利最佳？"],
    ["Q3 擴大", "會員等級調整、VIP 活動、品類交叉推薦", "升等提醒是否拉高 AOV？"],
    ["Q4 衝刺與檢討", "檔期（雙 11、年末）、點數到期提醒、年度檢討", "全年留存與 LTV 是否提升？"],
  ];
  q.forEach(([h, f, hyp], i) => {
    const x = 0.5 + i * 2.28;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.5, w: 2.15, h: 0.5, fill: { color: [C.teal, C.mid, C.coral, C.dark][i] }, line: { color: C.mid, width: 0 } });
    s.addText(h, { x, y: 1.5, w: 2.15, h: 0.5, fontFace: FONT, fontSize: 14, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0 });
    card(s, x, 2.0, 2.15, 2.0, C.mid);
    s.addText(f, { x: x + 0.18, y: 2.1, w: 1.9, h: 1.0, fontFace: FONT, fontSize: 10.5, color: C.dark, margin: 0, valign: "top" });
    s.addText("要驗證", { x: x + 0.18, y: 3.15, w: 1.9, h: 0.22, fontFace: FONT, fontSize: 9.5, bold: true, color: C.coral, margin: 0 });
    s.addText(hyp, { x: x + 0.18, y: 3.38, w: 1.75, h: 0.55, fontFace: FONT, fontSize: 10.5, color: C.mute, margin: 0, valign: "top" });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.2, w: 8.95, h: 0.85, fill: { color: C.dark }, line: { color: C.dark, width: 0 } });
  s.addText("正確的嘗試：一次只改一個變數 ・ 設對照組 ・ 預寫成功指標與停損條件 ・ 跑滿一個完整回購週期再下結論",
    { x: 0.7, y: 4.2, w: 8.6, h: 0.85, fontFace: FONT, fontSize: 12, bold: true, color: "FFFFFF", margin: 0, valign: "middle" });
}

// ============ 15 channels ============
{
  const s = base("06 溝通管道", "八種管道：各有分工，不是越多越好", "管道選擇看目的與成本。LINE 在台灣滲透率高，但要控制訊息量避免封鎖。");
  const ch = [
    ["Email", "成本低、可放長內容、可自動化", "看點擊與轉換，不只開信率"],
    ["簡訊 SMS", "到達率高、即時，但單價高", "頻率過高易被反感"],
    ["LINE 官方帳號", "台灣滲透率高，可分眾推播", "控制訊息量避免封鎖"],
    ["App 推播", "免費、即時、可深度個人化", "通知授權率是關鍵"],
    ["社群 IG / FB", "擴散與互動，可再行銷", "個人化弱，需搭配名單廣告"],
    ["電話 / 專人", "最溫暖、成本最高", "留給 VIP 與高單價諮詢"],
    ["門市 / 專櫃", "面對面，連結線上身分", "門市行為要回寫 CRM"],
    ["站內訊息", "在使用情境中觸達", "個人化推薦、升等進度條"],
  ];
  ch.forEach(([h, a, b], i) => {
    const x = 0.5 + (i % 4) * 2.28, y = 1.5 + Math.floor(i / 4) * 1.8;
    card(s, x, y, 2.15, 1.65, i < 4 ? C.coral : C.teal);
    s.addText(h, { x: x + 0.2, y: y + 0.1, w: 1.85, h: 0.35, fontFace: FONT, fontSize: 13, bold: true, color: C.dark, margin: 0 });
    s.addText(a, { x: x + 0.2, y: y + 0.5, w: 1.85, h: 0.6, fontFace: FONT, fontSize: 10.5, color: C.dark, margin: 0, valign: "top" });
    s.addText(b, { x: x + 0.2, y: y + 1.1, w: 1.85, h: 0.5, fontFace: FONT, fontSize: 10, color: C.mute, margin: 0, valign: "top" });
  });
}

// ============ 16 vendors ============
{
  const s = base("06 市面廠商", "常見 CRM / 行銷自動化工具類型", "這是類型與代表清單，不是推薦。我沒有逐一核實各家最新功能、價格與台灣在地服務，採購前一定要做 PoC。");
  const rows = [
    ["全球級 CRM / 行銷雲", "Salesforce、Adobe、HubSpot", "整合度、客製彈性、導入成本與人力"],
    ["跨渠道互動 / 旅程", "Braze、MoEngage、Klaviyo、Emarsys", "即時觸發、分群能力、App / Email 整合"],
    ["CDP 與資料整合", "Segment、Treasure Data", "身分合併、資料串接、即時性"],
    ["台灣 / 亞洲電商與會員", "91APP、SHOPLINE、CYBERBIZ、Appier", "與電商、金流、LINE 整合、在地支援"],
    ["LINE 相關工具", "LINE 官方帳號 + 第三方 CRM 擴充", "會員綁定、分眾標籤、訊息成本"],
  ];
  const data = [[th("類型"), th("代表（舉例）"), th("評估重點")]].concat(rows.map((r) => [td(r[0], { bold: true, color: C.mid }), td(r[1]), td(r[2])]));
  s.addTable(data, { ...tableOpts(0.5, 1.4, [2.3, 3.2, 3.45]), rowH: 0.46 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.5, w: 8.95, h: 0.55, fill: { color: "FDECE7" }, line: { color: C.coral, width: 0.75 } });
  s.addText("提醒：以上為清單而非推薦；各家最新功能與價格未逐一核實。請向廠商索取 demo 與案例，並用自己的資料做 PoC。", { x: 0.7, y: 4.5, w: 8.6, h: 0.55, fontFace: FONT, fontSize: 11, color: C.dark, margin: 0, valign: "middle" });
}

// ============ 17 data mining flow ============
{
  const s = base("07 Data Mining", "五步流程：從問題到行動", "這些品牌的真實會員資料並未公開，所以不能直接用線上資料分析，只能用公開資訊設定情境與參數做模擬。");
  const st = [
    ["問題定義", "誰會回購、誰快流失、哪些商品帶動二購"],
    ["資料整理", "會員、訂單、商品、渠道表，以 customer_id 串接；處理重複與退貨"],
    ["特徵", "R/F/M、AOV、回購間隔、品類、渠道、折扣依賴度"],
    ["分析", "RFM、Cohort 留存、籃分析、流失預測"],
    ["行動", "輸出名單進旅程，用對照組衡量增量"],
  ];
  st.forEach(([h, b], i) => {
    const x = 0.5 + i * 1.8;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.5, w: 1.68, h: 0.6, fill: { color: i === 4 ? C.coral : C.mid }, line: { color: C.mid, width: 0 } });
    s.addText(`${i + 1}  ${h}`, { x, y: 1.5, w: 1.68, h: 0.6, fontFace: FONT, fontSize: 13, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0 });
    card(s, x, 2.1, 1.68, 1.65, i === 4 ? C.coral : C.mid);
    s.addText(b, { x: x + 0.18, y: 2.2, w: 1.42, h: 1.75, fontFace: FONT, fontSize: 11.5, color: C.dark, margin: 0, valign: "top" });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.95, w: 8.95, h: 0.95, fill: { color: C.dark }, line: { color: C.dark, width: 0 } });
  s.addText("Dashboard 用三種零售型態的模擬資料：服飾快時尚（UNIQLO 風格）、美妝生活通路（寶雅風格）、居家生活（MUJI 風格）。僅為風格參考，不是這些品牌的真實數據。",
    { x: 0.7, y: 3.95, w: 8.6, h: 0.95, fontFace: FONT, fontSize: 12, color: "FFFFFF", margin: 0, valign: "middle" });
}

// ============ 18 dashboard + chart ============
{
  const s = base("07 互動 Dashboard", "營收集中在少數群：看模擬資料的 RFM 分群", "資料為模擬的美妝生活通路型會員。重點：Champions 人數不到兩成，卻貢獻近四成營收；At Risk (VIP) 是優先挽回對象。");
  const e1 = [...ANS.E1].sort((a, b) => b[3] - a[3]);
  s.addChart(pres.charts.BAR, [{ name: "營收占比 %", labels: e1.map((r) => r[0]), values: e1.map((r) => r[3]) }], {
    x: 0.5, y: 1.4, w: 5.6, h: 3.3, barDir: "bar", chartColors: [C.mid, C.coral, C.mid, C.mid, C.mid, C.mid, C.mid],
    chartArea: { fill: { color: C.card } }, catAxisLabelColor: C.mute, valAxisLabelColor: C.mute, catAxisLabelFontFace: FONT, catAxisLabelFontSize: 10,
    valAxisLabelFontSize: 9, valGridLine: { color: C.line, size: 0.5 }, catGridLine: { style: "none" }, catAxisOrientation: "maxMin",
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.dark, dataLabelFontSize: 10, dataLabelFormatCode: "0.0", showLegend: false,
    showTitle: true, title: "各 RFM 分群營收占比 %（模擬資料）", titleFontFace: FONT, titleFontSize: 11, titleColor: C.mute,
  });
  const ch = e1.find((r) => r[0] === "Champions"), ar = e1.find((r) => r[0] === "At Risk (VIP)");
  const total = ANS.E1.reduce((a, r) => a + r[1], 0);
  [[ch, "最重要的群", C.mid], [ar, "最該優先挽回", C.coral]].forEach(([r, lab, color], i) => {
    const y = 1.4 + i * 1.7;
    card(s, 6.3, y, 3.15, 1.6, color);
    s.addText(`${r[3]}%`, { x: 6.55, y: y + 0.1, w: 2.8, h: 0.7, fontFace: FONT, fontSize: 34, bold: true, color, margin: 0 });
    s.addText(`${r[0]}：${r[1]} 人（占 ${(r[1] / total * 100).toFixed(0)}% 會員）貢獻的營收。${lab}。`, { x: 6.55, y: y + 0.85, w: 2.8, h: 0.75, fontFace: FONT, fontSize: 10.5, color: C.dark, margin: 0, valign: "top" });
  });
  s.addText("Dashboard 六頁籤：總覽 ・ 新客 / 舊客 / 沉睡客 ・ RFM 分群 ・ Cohort 留存 ・ 會員等級模擬 ・ 品類回購", { x: 0.5, y: 4.85, w: 9, h: 0.25, fontFace: FONT, fontSize: 10, color: C.mute, margin: 0 });
}

// ============ 19 tools ============
{
  const s = base("08 必備工具", "四項工具：各管一段流程", "Excel 做快速驗證，SQL 撈資料，Python 做自動化與分析，Power BI 做報表。");
  const t = [
    ["Excel", "樞紐分析表、XLOOKUP、SUMIFS / COUNTIFS、條件式格式、DATEDIF", "快速 RFM、活動成效表、名單清洗"],
    ["Power BI", "星狀資料模型、DAX 量值、時間智慧、篩選器與鑽取", "月報 Dashboard、Cohort 熱力圖"],
    ["SQL", "SELECT / JOIN / GROUP BY、視窗函數、CTE、日期處理", "從資料庫撈會員與訂單、建立 RFM 與名單"],
    ["Python", "pandas、matplotlib、scikit-learn 基礎、Jupyter", "自動化報表、分群與流失預測、A/B 測試顯著性"],
  ];
  t.forEach(([h, a, b], i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.45 + Math.floor(i / 2) * 1.85;
    card(s, x, y, 4.4, 1.7, [C.teal, C.gold, C.mid, C.coral][i]);
    s.addText(h, { x: x + 0.25, y: y + 0.1, w: 3.9, h: 0.4, fontFace: FONT, fontSize: 16, bold: true, color: C.dark, margin: 0 });
    s.addText("基本必會　" + a, { x: x + 0.25, y: y + 0.55, w: 3.95, h: 0.6, fontFace: FONT, fontSize: 10.5, color: C.dark, margin: 0, valign: "top" });
    s.addText("CRM 用途　" + b, { x: x + 0.25, y: y + 1.17, w: 3.95, h: 0.5, fontFace: FONT, fontSize: 10.5, color: C.mute, margin: 0, valign: "top" });
  });
}

// ============ 20 code ============
{
  const s = base("08 必備工具", "範例：SQL 做 RFM、Python 做 Cohort", "SQL 的 DATEDIFF / GETDATE 為 SQL Server 語法，其他資料庫日期函數略有不同。Cohort 用 M0 人數當分母。");
  const panel = (x, w, label, code) => {
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.4, w, h: 3.3, fill: { color: C.code }, line: { color: C.code, width: 0 } });
    s.addText(label, { x: x + 0.2, y: 1.48, w: w - 0.4, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: C.gold, margin: 0 });
    s.addText(code, { x: x + 0.2, y: 1.85, w: w - 0.4, h: 3.1, fontFace: MONO, fontSize: 10.5, color: "DCE6F2", margin: 0, valign: "top" });
  };
  panel(0.5, 4.4, "SQL ・ RFM", [
    "WITH base AS (", "  SELECT customer_id,", "    MAX(order_date) AS last_order,", "    COUNT(DISTINCT order_id) AS frequency,", "    SUM(amount) AS monetary",
    "  FROM orders GROUP BY customer_id", ")", "SELECT customer_id,", "  DATEDIFF(day, last_order, GETDATE()) AS recency,", "  frequency, monetary,", "  NTILE(5) OVER (ORDER BY frequency) AS f_score", "FROM base;",
  ].join("\n"));
  panel(5.05, 4.4, "Python ・ Cohort 留存", [
    "first = o.groupby('customer_id')", "          ['order_date'].transform('min')", "o['cohort'] = first.dt.to_period('M')", "o['idx'] = (", "  (o.order_date.dt.year - first.dt.year) * 12", "  + (o.order_date.dt.month - first.dt.month))",
    "pivot = o.pivot_table(index='cohort',", "  columns='idx', values='customer_id',", "  aggfunc='nunique')", "retention = pivot.div(pivot[0], axis=0)",
  ].join("\n"));
}

// ============ 21 practice ============
{
  const s = base("課後練習", "配套練習：四個工具各一部分", "題本與練習檔都用 output/data 的模擬資料。教師版另有 A 部分自動核對。");
  const p = [
    ["A  Excel", "20 分", "樞紐分析、COUNTIFS、SUMIFS", C.teal],
    ["B  SQL", "30 分", "月度 KPI、30 天二購、生命週期", C.mid],
    ["C  Python", "25 分", "Cohort、訂單間隔、RFM 重建", C.coral],
    ["D  Power BI", "25 分", "DAX 量值、Cohort 矩陣、切片器", C.gold],
  ];
  p.forEach(([h, sc, b, color], i) => {
    const x = 0.5 + i * 2.28;
    card(s, x, 1.5, 2.15, 2.25, color);
    s.addText(h, { x: x + 0.2, y: 1.6, w: 1.85, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: C.dark, margin: 0 });
    s.addText(sc, { x: x + 0.2, y: 2.05, w: 1.85, h: 0.6, fontFace: FONT, fontSize: 26, bold: true, color, margin: 0 });
    s.addText(b, { x: x + 0.2, y: 2.75, w: 1.85, h: 0.95, fontFace: FONT, fontSize: 11.5, color: C.mute, margin: 0, valign: "top" });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.0, w: 8.95, h: 1.05, fill: { color: C.dark }, line: { color: C.dark, width: 0 } });
  s.addText([
    { text: "docs/CRM_學員作業題本.docx　題目與參考答案", options: { breakLine: true } },
    { text: "docs/CRM_學員練習檔.xlsx　作答用 Excel（含資料）", options: { breakLine: true } },
    { text: "output/crm_dashboard.html　互動 Dashboard", options: {} },
  ], { x: 0.7, y: 4.0, w: 8.6, h: 1.05, fontFace: FONT, fontSize: 11, color: "FFFFFF", margin: 0, valign: "middle", paraSpaceAfter: 2 });
}

// ============ 22 closing ============
{
  const s = pres.addSlide();
  n += 1;
  s.background = { color: C.dark };
  s.addText("帶走三件事", { x: 0.6, y: 0.7, w: 8, h: 0.7, fontFace: FONT, fontSize: 32, bold: true, color: "FFFFFF", margin: 0 });
  const k = [
    ["定義口徑要固定", "新客、活躍、沉睡的天數先寫清楚，月報才可比"],
    ["分群要有不同動作", "每個群要有不同的訊息、誘因、頻率與渠道"],
    ["用對照組證明增量", "沒有對照組，就不能證明 CRM 創造了額外營收"],
  ];
  k.forEach(([h, b], i) => {
    const y = 1.7 + i * 1.1;
    circle(s, 0.6, y + 0.1, 0.6, i + 1, C.coral, 20);
    s.addText(h, { x: 1.5, y, w: 7.5, h: 0.45, fontFace: FONT, fontSize: 19, bold: true, color: "FFFFFF", margin: 0, valign: "middle" });
    s.addText(b, { x: 1.5, y: y + 0.45, w: 7.5, h: 0.4, fontFace: FONT, fontSize: 13, color: "CADCFC", margin: 0, valign: "top" });
  });
  s.addNotes("收尾三點：口徑、分群動作、對照組。");
}

fs.mkdirSync(path.dirname(OUT), { recursive: true });
pres.writeFile({ fileName: OUT }).then(() => console.log("wrote", OUT, "slides:", n));
