/**
 * 每日時間表產生器（Google Apps Script）
 * 最小單位：30 分鐘，時間範圍 06:00 ~ 22:00
 *
 * 使用方式：
 * 1. 開啟 Google 試算表 → 擴充功能 → Apps Script
 * 2. 貼上本程式碼並儲存
 * 3. 執行 createSchedule()（第一次需授權），或重新整理試算表後使用選單「時間表」
 */

// ===== 可自行調整的設定 =====
const DAY_START = '06:00';
const DAY_END = '22:00';
const SLOT_MIN = 30;

// 每個區塊：[開始, 結束, 活動, 備註]
// 活動名稱為「電動」「讀書」的區塊會被計入統計
const WEEKDAY_BLOCKS = [
  ['06:00', '07:00', '起床、早餐、準備出門', '7:00 準時出門'],
  ['07:00', '16:00', '上學', '在校 8 小時（含通勤）'],
  ['16:00', '16:30', '回家休息', '約 16:10 到家'],
  ['16:30', '18:00', '電動', ''],
  ['18:00', '19:00', '晚餐', ''],
  ['19:00', '20:30', '電動', ''],
  ['20:30', '21:00', '洗澡', ''],
  ['21:00', '22:00', '自由時間', '可再加打電動 / 準備睡覺'],
];

const HOLIDAY_BLOCKS = [
  ['06:00', '07:00', '起床、早餐', ''],
  ['07:00', '09:00', '電動', ''],
  ['09:00', '09:30', '休息', '伸展、喝水、護眼'],
  ['09:30', '10:30', '讀書', '專心 1 小時'],
  ['10:30', '12:00', '電動', ''],
  ['12:00', '13:00', '午餐', ''],
  ['13:00', '13:30', '休息', '午休、護眼'],
  ['13:30', '17:30', '電動', ''],
  ['17:30', '18:00', '休息', '伸展、走動'],
  ['18:00', '19:00', '晚餐', ''],
  ['19:00', '21:30', '電動', ''],
  ['21:30', '22:00', '洗澡', '準備睡覺'],
];

const COLORS = {
  '電動': '#d9ead3',
  '讀書': '#cfe2f3',
  '上學': '#f4cccc',
  '晚餐': '#fff2cc',
  '午餐': '#fff2cc',
  '洗澡': '#ead1dc',
  '休息': '#eeeeee',
  '回家休息': '#eeeeee',
  '自由時間': '#fce5cd',
};
const DEFAULT_COLOR = '#ffffff';

// ===== 主程式 =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('時間表')
    .addItem('建立 / 重建時間表', 'createSchedule')
    .addItem('檢查時間表是否符合規則', 'validateSchedule')
    .addToUi();
}

function createSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  buildSheet_(ss, '平日', WEEKDAY_BLOCKS, { 電動: 3 });
  buildSheet_(ss, '假日', HOLIDAY_BLOCKS, { 電動: 10, 讀書: 1 });

  // 今天是平日還是假日，就先切到對應的分頁
  const day = new Date().getDay();
  const name = (day === 0 || day === 6) ? '假日' : '平日';
  ss.setActiveSheet(ss.getSheetByName(name));

  validateSchedule();
}

function buildSheet_(ss, name, blocks, minHours) {
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(name);
  }

  const slots = expandBlocks_(blocks);
  const header = [['時間', '活動', '備註']];
  const rows = slots.map(s => [s.label, s.activity, s.note]);

  sheet.getRange(1, 1, 1, 3).setValues(header)
    .setFontWeight('bold').setBackground('#434343').setFontColor('#ffffff');
  const body = sheet.getRange(2, 1, rows.length, 3);
  body.setValues(rows);

  // 依活動上色
  slots.forEach((s, i) => {
    sheet.getRange(i + 2, 1, 1, 3).setBackground(COLORS[s.activity] || DEFAULT_COLOR);
  });

  // 統計區（用公式計算，之後手動改活動也會自動更新）
  const lastRow = rows.length + 1;
  const range = `B2:B${lastRow}`;
  sheet.getRange('E1:G1').setValues([['統計', '時數', '最低要求']])
    .setFontWeight('bold').setBackground('#434343').setFontColor('#ffffff');
  const stats = [
    ['電動', `=COUNTIF(${range},"電動")*${SLOT_MIN}/60`, minHours['電動'] || ''],
    ['讀書', `=COUNTIF(${range},"讀書")*${SLOT_MIN}/60`, minHours['讀書'] || ''],
    ['上學', `=COUNTIF(${range},"上學")*${SLOT_MIN}/60`, ''],
  ];
  sheet.getRange(2, 5, stats.length, 3).setValues(stats);

  // 若電動 / 讀書時數低於要求，數字變紅
  const rules = [];
  [2, 3].forEach(r => {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($G${r}<>"",$F${r}<$G${r})`)
      .setFontColor('#cc0000').setBold(true)
      .setRanges([sheet.getRange(`F${r}`)]).build());
  });
  sheet.setConditionalFormatRules(rules);

  // 活動欄加下拉選單，方便之後自行調整
  const options = Object.keys(COLORS).concat(['起床、早餐、準備出門', '起床、早餐', '其他']);
  sheet.getRange(range).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(options, true)
      .setAllowInvalid(true).build());

  sheet.getRange(1, 1, lastRow, 3).setBorder(true, true, true, true, true, true);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(5, 80);
  sheet.setFrozenRows(1);
}

// 把區塊展開成每 30 分鐘一格，並檢查是否完整覆蓋 06:00~22:00
function expandBlocks_(blocks) {
  const slots = [];
  let cursor = toMin_(DAY_START);
  const end = toMin_(DAY_END);

  blocks.forEach(b => {
    const s = toMin_(b[0]);
    const e = toMin_(b[1]);
    if (s !== cursor) {
      throw new Error(`時間不連續：${b[0]} 應該接在 ${toStr_(cursor)} 後面`);
    }
    if ((e - s) % SLOT_MIN !== 0) {
      throw new Error(`區塊 ${b[0]}-${b[1]} 不是 ${SLOT_MIN} 分鐘的倍數`);
    }
    for (let t = s; t < e; t += SLOT_MIN) {
      slots.push({
        label: `${toStr_(t)}–${toStr_(t + SLOT_MIN)}`,
        activity: b[2],
        note: t === s ? b[3] : '',
      });
    }
    cursor = e;
  });

  if (cursor !== end) {
    throw new Error(`時間表只排到 ${toStr_(cursor)}，應該排到 ${DAY_END}`);
  }
  return slots;
}

function validateSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const msgs = [];
  const checks = [
    ['平日', { 電動: 3 }],
    ['假日', { 電動: 10, 讀書: 1 }],
  ];
  checks.forEach(([name, req]) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { msgs.push(`找不到「${name}」分頁`); return; }
    const values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat();
    Object.keys(req).forEach(act => {
      const hours = values.filter(v => v === act).length * SLOT_MIN / 60;
      msgs.push(`${name}｜${act}：${hours} 小時（要求至少 ${req[act]} 小時）${hours >= req[act] ? ' ✅' : ' ❌'}`);
    });
  });
  SpreadsheetApp.getUi().alert(msgs.join('\n'));
}

// ===== 網頁版（部署為網頁應用程式時使用）=====
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('每日時間表')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 給 index.html 呼叫，回傳平日 / 假日的每格資料
function getScheduleData() {
  const pack = blocks => expandBlocks_(blocks).map(s => ({
    label: s.label,
    start: toMin_(s.label.split('–')[0]),
    activity: s.activity,
    note: s.note,
    color: COLORS[s.activity] || DEFAULT_COLOR,
  }));
  return { weekday: pack(WEEKDAY_BLOCKS), holiday: pack(HOLIDAY_BLOCKS) };
}

// ===== 小工具 =====
function toMin_(hhmm) {
  const [h, m] = hhmm.split(':').map(Number);
  return h * 60 + m;
}

function toStr_(min) {
  const h = String(Math.floor(min / 60)).padStart(2, '0');
  const m = String(min % 60).padStart(2, '0');
  return `${h}:${m}`;
}
