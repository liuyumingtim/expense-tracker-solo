// ================================================
// 雙人記帳本 - Google Apps Script 後端 v4
// v4 變更：
//  - add 改為 upsert（同 id 不再重複新增 → 支援重試與編輯，根治重複列）
//  - 新增 update / addMany（批次 upsert，給前端未同步佇列一次推送）
//  - 所有寫入包 LockService（兩支手機同時寫入不互踩、不刪錯列）
//  - 新增 getSettings / setSettings（使用者姓名 + 匯率雲端同步）
//  - getAll 去重（同 id 只回一筆，防舊資料殘留的重複）
//  - 清掉 setBudgets 的 B1 殘留；budgets / settings 採「合併」避免互相覆蓋
// ================================================

const HEADERS = ['id', 'date', 'kind', 'trip', 'cat', 'amount', 'curr', 'note', 'ts', 'user'];
const APP_SPREADSHEET_NAME = '錢錢有進有出';
const HEADER_BG = '#221e18';
const HEADER_FG = '#fbf9f2';
const TAB_COLOR = '#9a7b3c';
const BACKUP_TAB_COLOR = '#9a8f78';
const META_SHEET_BUDGETS   = '_budgets';
const META_SHEET_RECURRING = '_recurring';
const META_SHEET_SETTINGS  = '_settings';

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    let action, data;
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
      action = data.action;
    } else {
      const params = (e && e.parameter) || {};
      data = params.payload ? JSON.parse(params.payload) : normalizeRequestData(params);
      action = data.action || params.action;
    }

    let result;
    if      (action === 'ping')         result = { success: true, message: 'pong' };
    else if (action === 'getAll')       result = getAllRecords();
    else if (action === 'add')          result = upsertRecord(data);            // 改為 upsert
    else if (action === 'update')       result = upsertRecord(data);            // 新增（編輯走這裡，原子）
    else if (action === 'addMany')      result = addMany(data.records);         // 新增（批次 upsert）
    else if (action === 'delete')       result = deleteRecord(data.id);
    else if (action === 'getBudgets')   result = getBudgets();
    else if (action === 'setBudgets')   result = setBudgets(data.budgets);
    else if (action === 'getRecurring') result = getRecurring();
    else if (action === 'setRecurring') result = setRecurring(data.recurring);
    else if (action === 'getSettings')  result = getSettings();                 // 新增（姓名+匯率）
    else if (action === 'setSettings')  result = setSettings(data.settings);    // 新增
    else if (action === 'repairSchema') result = repairSchema();                // 整理月份分頁欄位
    else                                result = { success: false, message: 'Unknown action: ' + action };

    output.setContent(JSON.stringify(result));
  } catch (err) {
    output.setContent(JSON.stringify({ success: false, message: err.toString() }));
  }

  return output;
}

function normalizeRequestData(params) {
  const out = Object.assign({}, params || {});
  ['records', 'budgets', 'recurring', 'settings'].forEach(function(key) {
    if (typeof out[key] !== 'string') return;
    const value = out[key].trim();
    if (!value || !/^[\[{]/.test(value)) return;
    try { out[key] = JSON.parse(value); } catch(e) {}
  });
  return out;
}

// ================================================
// 工具：LockService 包住所有寫入（避免併發互踩）
// ================================================
function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // 最多等 20 秒
  try { return fn(); }
  finally { lock.releaseLock(); }
}

// ================================================
// 工具：日期格式統一
// ================================================
function normalizeDate(raw) {
  if (!raw) return '';
  const s = String(raw).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }
  return s;
}

// ================================================
// 工具：取得或建立「月份分頁」（例如 2025-05）
// ================================================
function getOrCreateMonthSheet(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monthKey = String(dateStr || '').slice(0, 7) || getThisMonth();
  let sheet = ss.getSheetByName(monthKey);
  if (!sheet) {
    sheet = ss.insertSheet(monthKey);
    ss.moveActiveSheet(ss.getNumSheets());
    sheet.appendRow(HEADERS);
    ensureMonthSheetShape(sheet);
  } else {
    ensureMonthSheetShape(sheet);
  }
  return sheet;
}

function ensureMonthSheetShape(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), HEADERS.length);
  let currentHeaders = [];
  if (lastRow >= 1) {
    currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(function(h){ return String(h || '').trim().toLowerCase(); });
  }
  const alreadyCurrent = HEADERS.every(function(h, idx) {
    return currentHeaders[idx] === h;
  });

  if (!alreadyCurrent && lastRow > 1) {
    const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const normalizedRows = rows.map(function(row) {
      return normalizeSheetRow(row, currentHeaders);
    });
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(2, 1, normalizedRows.length, HEADERS.length).setValues(normalizedRows);
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  const hr = sheet.getRange(1, 1, 1, HEADERS.length);
  hr.setBackground(HEADER_BG);
  hr.setFontColor(HEADER_FG);
  hr.setFontWeight('bold');
  hr.setHorizontalAlignment('center');
  sheet.setTabColor(TAB_COLOR);
  [170, 110, 90, 170, 120, 100, 90, 230, 150, 90].forEach(function(width, idx) {
    sheet.setColumnWidth(idx + 1, width);
  });
  const extraCols = sheet.getMaxColumns() - HEADERS.length;
  if (extraCols > 0) sheet.deleteColumns(HEADERS.length + 1, extraCols);
}

function normalizeSheetRow(row, headers) {
  const val = function(name) {
    const idx = headers.indexOf(name);
    return idx >= 0 ? row[idx] : '';
  };
  const kind = normalizeKind(val('kind') || val('type'));
  return [
    String(val('id') || ''),
    normalizeDate(val('date') || ''),
    kind,
    String(val('trip') || val('tour') || val('group') || ''),
    String(val('cat') || val('category') || ''),
    Number(val('amount') || 0),
    String(val('curr') || 'THB'),
    String(val('note') || ''),
    Number(val('ts') || 0),
    String(val('user') || '')
  ];
}

function normalizeKind(raw) {
  const s = String(raw || '').trim().toLowerCase();
  return s === 'income' ? 'income' : 'expense';
}

function backupSheetIfNeeded(ss, sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), HEADERS.length)).getValues()[0]
    .map(function(h){ return String(h || '').trim(); });
  const looksLikeTravel = headers.indexOf('trip') >= 0 || headers.indexOf('country') >= 0 || headers.indexOf('walletFrom') >= 0;
  const hasRows = sheet.getLastRow() > 1;
  if (!looksLikeTravel || !hasRows) return '';

  let backupName = sheet.getName() + '_travel_backup';
  let i = 2;
  while (ss.getSheetByName(backupName)) {
    backupName = sheet.getName() + '_travel_backup_' + i;
    i++;
  }
  sheet.copyTo(ss).setName(backupName).setTabColor(BACKUP_TAB_COLOR);
  return backupName;
}

function repairSchema() {
  return withLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.rename(APP_SPREADSHEET_NAME);
    cleanupDefaultSheets(ss);
    const backups = [];
    ss.getSheets().forEach(function(sheet) {
      if (!isMonthSheet(sheet.getName())) return;
      const backupName = backupSheetIfNeeded(ss, sheet);
      if (backupName) backups.push(backupName);
      ensureMonthSheetShape(sheet);
    });

    getOrCreateMonthSheet(getThisMonth());
    return { success: true, spreadsheet: ss.getName(), backups: backups, headers: HEADERS };
  });
}

function cleanupDefaultSheets(ss) {
  const defaults = ['工作表1', 'Sheet1'];
  defaults.forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet || ss.getSheets().length <= 1) return;
    if (isBlankSheet(sheet)) ss.deleteSheet(sheet);
  });
}

function isBlankSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return true;
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  return values.every(function(row) {
    return row.every(function(cell) { return String(cell || '').trim() === ''; });
  });
}

function getThisMonth() {
  const now = new Date();
  return now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
}

function isMonthSheet(name) {
  return /^\d{4}-\d{2}$/.test(name);
}

// 找某個 id 在哪個月份分頁的哪一列（1-based row）
function findRecordRow(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const target = String(id).trim();
  for (let s = 0; s < sheets.length; s++) {
    const sheet = sheets[s];
    if (!isMonthSheet(sheet.getName())) continue;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) continue;
    const headerRow = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
    const idCol = headerRow.indexOf('id');
    if (idCol < 0) continue;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]).trim() === target) {
        return { sheet: sheet, row: i + 1 };
      }
    }
  }
  return null;
}

// ================================================
// 記帳記錄：讀所有月份分頁（同 id 去重）
// ================================================
function getAllRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const byId = {};

  sheets.forEach(function(sheet) {
    if (!isMonthSheet(sheet.getName())) return;
    ensureMonthSheetShape(sheet);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const headerRow = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
    const col = function(name){ return headerRow.indexOf(name); };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = String(row[col('id')] || '').trim();
      if (!id) continue;
      byId[id] = {
        id:     id,
        date:   normalizeDate(row[col('date')]),
        kind:   normalizeKind(row[col('kind')]),
        trip:   String(row[col('trip')]   || ''),
        cat:    String(row[col('cat')]    || ''),
        amount: Number(row[col('amount')] || 0),
        curr:   String(row[col('curr')]   || 'THB'),
        note:   String(row[col('note')]   || ''),
        ts:     Number(row[col('ts')]     || 0),
        user:   String(row[col('user')]   || '')
      };
    }
  });

  const allRecords = Object.keys(byId).map(function(k){ return byId[k]; });
  allRecords.sort(function(a, b){ return a.date.localeCompare(b.date); });
  return { success: true, records: allRecords };
}

// ================================================
// 記帳記錄：upsert 核心（無 lock，給 add / update / addMany 共用）
//  - id 已存在：同月就地更新；跨月則刪舊列、寫到新月份分頁
//  - id 不存在：寫入對應月份分頁
// ================================================
function upsertOne(data) {
  const id = String(data.id || (Date.now() + '_' + Math.random().toString(36).slice(2)));
  const date = normalizeDate(data.date || '');
  const targetMonth = String(date).slice(0, 7) || getThisMonth();
  const rowVals = [
    id,
    date,
    normalizeKind(data.kind),
    String(data.trip   || ''),
    String(data.cat    || ''),
    Number(data.amount || 0),
    String(data.curr   || 'THB'),
    String(data.note   || ''),
    Number(data.ts     || Date.now()),
    String(data.user   || '')
  ];

  const existing = findRecordRow(id);
  if (existing) {
    if (existing.sheet.getName() === targetMonth) {
      existing.sheet.getRange(existing.row, 1, 1, HEADERS.length).setValues([rowVals]);
    } else {
      existing.sheet.deleteRow(existing.row);
      getOrCreateMonthSheet(date).appendRow(rowVals);
    }
  } else {
    getOrCreateMonthSheet(date).appendRow(rowVals);
  }
  return id;
}

function upsertRecord(data) {
  return withLock(function() {
    return { success: true, id: upsertOne(data) };
  });
}

function addMany(records) {
  if (!records || !records.length) return { success: true, count: 0 };
  return withLock(function() {
    records.forEach(function(r){ upsertOne(r); });
    return { success: true, count: records.length };
  });
}

// ================================================
// 記帳記錄：刪除（搜尋所有月份分頁）
// ================================================
function deleteRecord(id) {
  return withLock(function() {
    const found = findRecordRow(id);
    if (found) {
      found.sheet.deleteRow(found.row);
      return { success: true };
    }
    return { success: false, message: 'Record not found: ' + id };
  });
}

// ================================================
// 預算：讀 / 寫（存在 _budgets 分頁 A1 的 JSON；採合併避免互相覆蓋）
// ================================================
function getBudgets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(META_SHEET_BUDGETS);
  if (!sheet) return { success: true, budgets: {} };
  const val = sheet.getRange('A1').getValue();
  let budgets = {};
  try { if (val) budgets = JSON.parse(val); } catch(e) {}
  return { success: true, budgets: budgets };
}

function setBudgets(budgets) {
  return withLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(META_SHEET_BUDGETS);
    if (!sheet) sheet = ss.insertSheet(META_SHEET_BUDGETS);
    let cur = {};
    try { const v = sheet.getRange('A1').getValue(); if (v) cur = JSON.parse(v); } catch(e) {}
    const merged = Object.assign({}, cur, budgets || {});
    sheet.getRange('A1').setValue(JSON.stringify(merged));
    return { success: true, budgets: merged };
  });
}

// ================================================
// 定期付款：讀 / 寫（存在 _recurring 分頁 A1）
//  - 採「整包取代」：前端在讀取時已合併雲端+本機，推回的就是完整清單
//    （這樣刪除才能正確傳播，不會被雲端救回來）
// ================================================
function getRecurring() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(META_SHEET_RECURRING);
  if (!sheet) return { success: true, recurring: [] };
  const val = sheet.getRange('A1').getValue();
  let recurring = [];
  try { if (val) recurring = JSON.parse(val); } catch(e) {}
  return { success: true, recurring: recurring };
}

function setRecurring(recurring) {
  return withLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(META_SHEET_RECURRING);
    if (!sheet) sheet = ss.insertSheet(META_SHEET_RECURRING);
    sheet.getRange('A1').setValue(JSON.stringify(recurring || []));
    return { success: true };
  });
}

// ================================================
// 設定（使用者姓名 + 匯率）：讀 / 寫（存在 _settings 分頁 A1）
//  - setSettings 採欄位合併：一端只改 rates、另一端只改 usernames 不會互相清掉
// ================================================
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(META_SHEET_SETTINGS);
  if (!sheet) return { success: true, settings: {} };
  const val = sheet.getRange('A1').getValue();
  let settings = {};
  try { if (val) settings = JSON.parse(val); } catch(e) {}
  return { success: true, settings: settings };
}

function setSettings(settings) {
  return withLock(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(META_SHEET_SETTINGS);
    if (!sheet) sheet = ss.insertSheet(META_SHEET_SETTINGS);
    let cur = {};
    try { const v = sheet.getRange('A1').getValue(); if (v) cur = JSON.parse(v); } catch(e) {}
    const merged = Object.assign({}, cur, settings || {});
    sheet.getRange('A1').setValue(JSON.stringify(merged));
    return { success: true, settings: merged };
  });
}

// ================================================
// 資料遷移工具（一次性執行）
// 把舊的「記帳資料」分頁資料遷移到新的月份分頁格式
// 完成後可刪除此函數
// ================================================
function migrateOldSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const OLD_SHEET = '記帳資料';
  const oldSheet = ss.getSheetByName(OLD_SHEET);
  if (!oldSheet) {
    Logger.log('找不到舊分頁「記帳資料」，無需遷移');
    return;
  }

  const data = oldSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log('舊分頁沒有資料，無需遷移');
    return;
  }

  const headerRow = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
  const col = function(name){ return headerRow.indexOf(name); };
  let migrated = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = String(row[col('id')] || '').trim();
    if (!id) continue;

    const date = normalizeDate(row[col('date')]);
    const monthSheet = getOrCreateMonthSheet(date);

    const existing = monthSheet.getDataRange().getValues();
    const existingIds = existing.slice(1).map(function(r){ return String(r[0]).trim(); });
    if (existingIds.includes(id)) continue;

    monthSheet.appendRow([
      id,
      date,
      normalizeKind(row[col('kind')]),
      String(row[col('trip')]   || ''),
      String(row[col('cat')]    || ''),
      Number(row[col('amount')] || 0),
      String(row[col('curr')]   || 'THB'),
      String(row[col('note')]   || ''),
      Number(row[col('ts')]     || 0),
      String(row[col('user')]   || '')
    ]);
    migrated++;
  }

  Logger.log('遷移完成：共遷移 ' + migrated + ' 筆記錄');
  oldSheet.setName('記帳資料_backup');
  Logger.log('舊分頁已重新命名為「記帳資料_backup」，確認無誤後可手動刪除');
}
