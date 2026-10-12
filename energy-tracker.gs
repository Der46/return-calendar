/**
 * 回歸帳號能量紀錄 - Google Apps Script
 *
 * 使用方式：
 * 1. 開一個新的 Google 試算表 → 擴充功能 → Apps Script
 * 2. 把這整個檔案貼進 Code.gs，儲存
 * 3. 回到試算表重新整理，上方會出現「能量紀錄」選單 → 點「初始化工作表」
 * 4. 之後只要在「紀錄」工作表填：回歸日期、帳號、(回歸能量)、離開日期、離開能量
 *    其餘欄位與「統計」工作表會自動計算
 *
 * 本次賺取的算法：
 * - 有填「回歸能量」→ 離開能量 - 回歸能量
 * - 沒填 → 離開能量 - 同帳號上一次的離開能量
 * - 第一次紀錄又沒填回歸能量 → 無法計算，留白
 */

const ACCOUNT_SHEET = '帳號';
const RECORD_SHEET = '紀錄';
const STATS_SHEET = '統計';

const DEFAULT_ACCOUNTS = [
  ['酷瑪', '歐秀婠'],
  ['威諺', '林傳勝'],
  ['椽家', '林傳勝'],
  ['Xin', 'XinXin'],
  ['沐晨', '林傳勝'],
  ['傳勝', '林傳勝'],
  ['家俊', 'Polly Chan'],
  ['家俊B', 'Polly Chan'],
  ['寶莉', 'Polly Chan'],
  ['雲秦', '林傳勝'],
  ['墨淵', '林傳勝'],
  ['智恩', '林傳勝'],
  ['若彤', '林傳勝'],
  ['Tony', 'Polly Chan'],
  ['芷涵', '林傳勝'],
  ['芷涵B', '林傳勝'],
  ['沐晨B', '林傳勝'],
  ['雲秦B', '林傳勝'],
  ['智恩B', '林傳勝'],
  ['傳勝B', '林傳勝'],
  ['威諺B', '林傳勝'],
];

// 紀錄工作表欄位（1-based）
const COL = {
  enterDate: 1,   // A 回歸日期（輸入）
  account: 2,     // B 帳號（輸入）
  enterEnergy: 3, // C 回歸能量（選填）
  leaveDate: 4,   // D 離開日期（輸入）
  leaveEnergy: 5, // E 離開能量（輸入）
  owner: 6,       // F 擁有者（自動）
  round: 7,       // G 第幾次（自動）
  baseEnergy: 8,  // H 起始能量（自動）
  earned: 9,      // I 本次賺取（自動）
  perDay: 10,     // J 每日平均（自動）
  gapDays: 11,    // K 距上次回歸天數（自動）
  note: 12,       // L 備註（輸入）
};
const RECORD_HEADERS = [
  '回歸日期', '帳號', '回歸能量(選填)', '離開日期', '離開能量',
  '擁有者', '第幾次', '起始能量', '本次賺取', '每日平均', '距上次回歸(天)', '備註',
];
const AUTO_FIRST = COL.owner;
const AUTO_COUNT = COL.gapDays - COL.owner + 1;
const MIN_GAP_DAYS = 30;
const DAY_MS = 24 * 60 * 60 * 1000;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('能量紀錄')
    .addItem('初始化工作表', 'setupSheets')
    .addItem('重新計算', 'recalculate')
    .addToUi();
}

function onEdit(e) {
  const sheet = e.range.getSheet();
  const name = sheet.getName();
  if (name === RECORD_SHEET || name === ACCOUNT_SHEET) {
    recalculate();
  }
}

function setupSheets() {
  const ss = SpreadsheetApp.getActive();

  let accSheet = ss.getSheetByName(ACCOUNT_SHEET);
  if (!accSheet) {
    accSheet = ss.insertSheet(ACCOUNT_SHEET);
    accSheet.getRange(1, 1, 1, 2).setValues([['帳號', '擁有者']]).setFontWeight('bold');
    accSheet.getRange(2, 1, DEFAULT_ACCOUNTS.length, 2).setValues(DEFAULT_ACCOUNTS);
    accSheet.setFrozenRows(1);
  }

  let recSheet = ss.getSheetByName(RECORD_SHEET);
  if (!recSheet) {
    recSheet = ss.insertSheet(RECORD_SHEET);
    recSheet.getRange(1, 1, 1, RECORD_HEADERS.length)
      .setValues([RECORD_HEADERS])
      .setFontWeight('bold');
    recSheet.setFrozenRows(1);
    // 自動欄位用灰底，提醒不要手動輸入
    recSheet.getRange(1, AUTO_FIRST, 1, AUTO_COUNT).setBackground('#d9d9d9');
    recSheet.getRange(2, AUTO_FIRST, 999, AUTO_COUNT).setBackground('#f3f3f3');
    recSheet.getRange(2, COL.enterDate, 999, 1).setNumberFormat('yyyy/mm/dd');
    recSheet.getRange(2, COL.leaveDate, 999, 1).setNumberFormat('yyyy/mm/dd');
    recSheet.getRange(2, COL.enterEnergy, 999, 1).setNumberFormat('#,##0');
    recSheet.getRange(2, COL.leaveEnergy, 999, 1).setNumberFormat('#,##0');
    recSheet.getRange(2, COL.baseEnergy, 999, 2).setNumberFormat('#,##0');
    recSheet.getRange(2, COL.perDay, 999, 1).setNumberFormat('#,##0.0');
  }

  // 帳號欄位下拉選單，來源是「帳號」工作表
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(accSheet.getRange('A2:A'), true)
    .setAllowInvalid(false)
    .build();
  recSheet.getRange(2, COL.account, 999, 1).setDataValidation(rule);

  if (!ss.getSheetByName(STATS_SHEET)) {
    ss.insertSheet(STATS_SHEET);
  }

  const defaultSheet = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  recalculate();
  ss.setActiveSheet(recSheet);
}

function recalculate() {
  const ss = SpreadsheetApp.getActive();
  const recSheet = ss.getSheetByName(RECORD_SHEET);
  const accSheet = ss.getSheetByName(ACCOUNT_SHEET);
  if (!recSheet || !accSheet) return;

  const ownerMap = readOwners(accSheet);
  const lastRow = recSheet.getLastRow();
  const rows = lastRow > 1
    ? recSheet.getRange(2, 1, lastRow - 1, RECORD_HEADERS.length).getValues()
    : [];

  // 依帳號分組，再依回歸日期排序（輸入順序不影響計算）
  const byAccount = {};
  rows.forEach((row, i) => {
    const account = String(row[COL.account - 1]).trim();
    if (!account) return;
    (byAccount[account] = byAccount[account] || []).push({ row, i });
  });

  const output = rows.map(() => new Array(AUTO_COUNT).fill(''));

  Object.keys(byAccount).forEach((account) => {
    const list = byAccount[account].sort((a, b) =>
      toTime(a.row[COL.enterDate - 1]) - toTime(b.row[COL.enterDate - 1]));

    let prevLeaveEnergy = null;
    let prevEnterDate = null;

    list.forEach((item, idx) => {
      const r = item.row;
      const enterDate = asDate(r[COL.enterDate - 1]);
      const leaveDate = asDate(r[COL.leaveDate - 1]);
      const enterEnergy = asNumber(r[COL.enterEnergy - 1]);
      const leaveEnergy = asNumber(r[COL.leaveEnergy - 1]);

      const base = enterEnergy !== null ? enterEnergy : prevLeaveEnergy;
      const earned = base !== null && leaveEnergy !== null ? leaveEnergy - base : '';

      let perDay = '';
      if (earned !== '' && enterDate && leaveDate) {
        const days = Math.round((leaveDate - enterDate) / DAY_MS) + 1;
        if (days > 0) perDay = earned / days;
      }

      const gap = enterDate && prevEnterDate
        ? Math.round((enterDate - prevEnterDate) / DAY_MS)
        : '';

      output[item.i] = [
        ownerMap[account] || '(未登記)',
        idx + 1,
        base === null ? '' : base,
        earned,
        perDay,
        gap,
      ];

      if (leaveEnergy !== null) prevLeaveEnergy = leaveEnergy;
      if (enterDate) prevEnterDate = enterDate;
    });
  });

  if (output.length) {
    recSheet.getRange(2, AUTO_FIRST, output.length, AUTO_COUNT).setValues(output);
    highlightRows(recSheet, output);
  }

  writeStats(ss, ownerMap, rows, output);
}

// 本次賺取為負數標紅；距上次回歸不足 30 天標橘
function highlightRows(sheet, output) {
  const earnedColors = output.map((o) => [typeof o[3] === 'number' && o[3] < 0 ? '#cc0000' : '#000000']);
  const gapBgs = output.map((o) => [typeof o[5] === 'number' && o[5] < MIN_GAP_DAYS ? '#f9cb9c' : '#f3f3f3']);
  sheet.getRange(2, COL.earned, output.length, 1).setFontColors(earnedColors);
  sheet.getRange(2, COL.gapDays, output.length, 1).setBackgrounds(gapBgs);
}

function writeStats(ss, ownerMap, rows, output) {
  const sheet = ss.getSheetByName(STATS_SHEET);
  if (!sheet) return;
  sheet.clear();

  const accStats = {};
  Object.keys(ownerMap).forEach((acc) => {
    accStats[acc] = { owner: ownerMap[acc], count: 0, total: 0, earnedCount: 0, lastEnter: null, lastLeave: null, lastEnergy: '' };
  });

  rows.forEach((row, i) => {
    const account = String(row[COL.account - 1]).trim();
    if (!account) return;
    const s = accStats[account] || (accStats[account] = {
      owner: '(未登記)', count: 0, total: 0, earnedCount: 0, lastEnter: null, lastLeave: null, lastEnergy: '',
    });
    s.count++;
    const earned = output[i][3];
    if (typeof earned === 'number') {
      s.total += earned;
      s.earnedCount++;
    }
    const enterDate = asDate(row[COL.enterDate - 1]);
    if (enterDate && (!s.lastEnter || enterDate > s.lastEnter)) {
      s.lastEnter = enterDate;
      s.lastLeave = asDate(row[COL.leaveDate - 1]);
      const le = asNumber(row[COL.leaveEnergy - 1]);
      s.lastEnergy = le === null ? '' : le;
    }
  });

  // 帳號統計
  const accHeader = ['帳號', '擁有者', '回歸次數', '總賺取', '平均每次賺取', '最近回歸日', '最近離開日', '目前能量', '最早下次回歸日'];
  const accRows = Object.keys(accStats).map((acc) => {
    const s = accStats[acc];
    return [
      acc,
      s.owner,
      s.count,
      s.total,
      s.earnedCount ? s.total / s.earnedCount : '',
      s.lastEnter || '',
      s.lastLeave || '',
      s.lastEnergy,
      s.lastEnter ? new Date(s.lastEnter.getTime() + MIN_GAP_DAYS * DAY_MS) : '',
    ];
  });

  sheet.getRange(1, 1).setValue('各帳號統計').setFontWeight('bold').setFontSize(12);
  sheet.getRange(2, 1, 1, accHeader.length).setValues([accHeader]).setFontWeight('bold').setBackground('#d9d9d9');
  if (accRows.length) {
    sheet.getRange(3, 1, accRows.length, accHeader.length).setValues(accRows);
    sheet.getRange(3, 4, accRows.length, 2).setNumberFormat('#,##0');
    sheet.getRange(3, 6, accRows.length, 2).setNumberFormat('yyyy/mm/dd');
    sheet.getRange(3, 8, accRows.length, 1).setNumberFormat('#,##0');
    sheet.getRange(3, 9, accRows.length, 1).setNumberFormat('yyyy/mm/dd');
  }

  // 擁有者統計
  const ownerStats = {};
  Object.keys(accStats).forEach((acc) => {
    const s = accStats[acc];
    const o = ownerStats[s.owner] || (ownerStats[s.owner] = { accounts: 0, count: 0, total: 0, energy: 0 });
    o.accounts++;
    o.count += s.count;
    o.total += s.total;
    if (typeof s.lastEnergy === 'number') o.energy += s.lastEnergy;
  });

  const ownerTop = accRows.length + 5;
  const ownerHeader = ['擁有者', '帳號數', '總回歸次數', '總賺取', '目前能量合計'];
  const ownerRows = Object.keys(ownerStats).map((name) => {
    const o = ownerStats[name];
    return [name, o.accounts, o.count, o.total, o.energy];
  });
  const grand = ownerRows.reduce((t, r) => [t[0], t[1] + r[1], t[2] + r[2], t[3] + r[3], t[4] + r[4]], ['合計', 0, 0, 0, 0]);
  ownerRows.push(grand);

  sheet.getRange(ownerTop, 1).setValue('各擁有者統計').setFontWeight('bold').setFontSize(12);
  sheet.getRange(ownerTop + 1, 1, 1, ownerHeader.length).setValues([ownerHeader]).setFontWeight('bold').setBackground('#d9d9d9');
  sheet.getRange(ownerTop + 2, 1, ownerRows.length, ownerHeader.length).setValues(ownerRows);
  sheet.getRange(ownerTop + 2, 4, ownerRows.length, 2).setNumberFormat('#,##0');
  sheet.getRange(ownerTop + 1 + ownerRows.length, 1, 1, ownerHeader.length).setFontWeight('bold');

  sheet.autoResizeColumns(1, accHeader.length);
}

function readOwners(accSheet) {
  const map = {};
  const last = accSheet.getLastRow();
  if (last < 2) return map;
  accSheet.getRange(2, 1, last - 1, 2).getValues().forEach(([acc, owner]) => {
    const key = String(acc).trim();
    if (key) map[key] = String(owner).trim();
  });
  return map;
}

function asDate(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'string' && v.trim()) {
    const d = new Date(v.trim().replace(/-/g, '/'));
    if (!isNaN(d)) return d;
  }
  return null;
}

function toTime(v) {
  const d = asDate(v);
  return d ? d.getTime() : Number.MAX_SAFE_INTEGER;
}

function asNumber(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = typeof v === 'number' ? v : Number(String(v).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}
