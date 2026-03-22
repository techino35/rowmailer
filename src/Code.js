/**
 * RowMailer - Google Sheets Add-on
 */

function onInstall(e) { onOpen(e); }

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('サイドバーを開く / Open Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('今すぐ実行 / Run Now', 'runMailerManual')
    .addItem('テスト送信 / Test Send', 'runTestSend')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('src/Sidebar')
    .setTitle('RowMailer').setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

function runMailer(isTest) {
  isTest = isTest || false;
  const props = PropertiesService.getDocumentProperties();
  const configJson = props.getProperty('ROWMAILER_CONFIG');
  if (!configJson) return { success: false, message: '設定が見つかりません。' };

  let config;
  try { config = JSON.parse(configJson); }
  catch (e) { return { success: false, message: '設定データが破損: ' + e.message }; }

  const licenseInfo = validateLicense();
  const dailyLimit = licenseInfo.isPro ? 500 : 50;
  const todayCount = getTodaySentCount_();
  if (todayCount >= dailyLimit) return { success: false, message: `本日の送信上限（${dailyLimit}通）に達しました。` };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(config.sheetName || ss.getActiveSheet().getName());
  if (!sheet) return { success: false, message: `シート "${config.sheetName}" が見つかりません。` };

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, sent: 0, message: 'データ行がありません。' };

  const headers = data[0];
  const rows = data.slice(1);
  let sentCount = 0, errorCount = 0;
  const maxSend = isTest ? 1 : (dailyLimit - todayCount);

  const lock = LockService.getDocumentLock();
  try { lock.waitLock(30000); }
  catch (e) { return { success: false, message: '別の処理が実行中です。' }; }

  try {
    for (let i = 0; i < rows.length && sentCount < maxSend; i++) {
      const row = rows[i];
      const rowIndex = i + 2;
      if (!checkCondition_(row, headers, config.condition)) continue;
      if (config.logColumnIndex !== undefined && config.logColumnIndex >= 0) {
        if (row[config.logColumnIndex] && !isTest) continue;
      }
      const toEmail = isTest ? Session.getActiveUser().getEmail() : getCellValue_(row, headers, config.toColumn);
      if (!toEmail || !isValidEmail_(toEmail)) { errorCount++; continue; }
      const rowObj = buildRowObject_(row, headers);
      const subject = replacePlaceholders(config.subject || '', rowObj);
      const body = replacePlaceholders(config.body || '', rowObj);
      const ccEmail = config.ccFixed || getCellValue_(row, headers, config.ccColumn) || '';
      const bccEmail = config.bccFixed || getCellValue_(row, headers, config.bccColumn) || '';
      try {
        const opts = { name: config.senderName || 'RowMailer', htmlBody: config.useHtml ? body : undefined };
        if (ccEmail) opts.cc = ccEmail;
        if (bccEmail) opts.bcc = bccEmail;
        GmailApp.sendEmail(toEmail, subject, body, opts);
        if (config.logColumnIndex !== undefined && config.logColumnIndex >= 0 && !isTest) {
          sheet.getRange(rowIndex, config.logColumnIndex + 1)
            .setValue(new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' }));
        }
        sentCount++;
        incrementTodaySentCount_();
      } catch (mailErr) { errorCount++; }
    }
  } finally { lock.releaseLock(); }

  return { success: true, sent: sentCount, errors: errorCount, message: `送信完了: ${sentCount}通送信, ${errorCount}件エラー` };
}

function runMailerManual() {
  const result = runMailer(false);
  SpreadsheetApp.getUi().alert('RowMailer\n\n' + result.message);
}

function runTestSend() {
  const myEmail = Session.getActiveUser().getEmail();
  const result = runMailer(true);
  SpreadsheetApp.getUi().alert(`RowMailer テスト送信\n\n送信先: ${myEmail}\n${result.message}`);
}

function checkCondition_(row, headers, condition) {
  if (!condition) return true;
  const cellValue = row[condition.columnIndex];
  const condValue = condition.value;
  switch (condition.operator) {
    case '=': return String(cellValue) === String(condValue);
    case '!=': return String(cellValue) !== String(condValue);
    case '>': return Number(cellValue) > Number(condValue);
    case '<': return Number(cellValue) < Number(condValue);
    case '>=': return Number(cellValue) >= Number(condValue);
    case '<=': return Number(cellValue) <= Number(condValue);
    case 'contains': return String(cellValue).indexOf(String(condValue)) !== -1;
    case 'not_contains': return String(cellValue).indexOf(String(condValue)) === -1;
    case 'is_not_empty': return cellValue !== '' && cellValue !== null && cellValue !== undefined;
    case 'is_empty': return cellValue === '' || cellValue === null || cellValue === undefined;
    default: return true;
  }
}

function getCellValue_(row, headers, columnName) {
  if (!columnName) return '';
  const idx = headers.indexOf(columnName);
  return idx === -1 ? '' : String(row[idx] || '');
}

function buildRowObject_(row, headers) {
  const obj = {};
  headers.forEach(function(h, i) { obj[h] = row[i] !== undefined ? row[i] : ''; });
  return obj;
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email));
}

function getTodaySentCount_() {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  return parseInt(PropertiesService.getDocumentProperties().getProperty('ROWMAILER_SENT_' + today) || '0', 10);
}

function incrementTodaySentCount_() {
  const props = PropertiesService.getDocumentProperties();
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const key = 'ROWMAILER_SENT_' + today;
  props.setProperty(key, String(parseInt(props.getProperty(key) || '0', 10) + 1));
}

function saveConfig(config) {
  try {
    PropertiesService.getDocumentProperties().setProperty('ROWMAILER_CONFIG', JSON.stringify(config));
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function loadConfig() {
  const json = PropertiesService.getDocumentProperties().getProperty('ROWMAILER_CONFIG');
  return json ? JSON.parse(json) : {};
}

function getSheetInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const lastCol = activeSheet.getLastColumn();
  const headers = lastCol > 0 ? activeSheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(function(h) { return h !== ''; }) : [];
  return {
    sheetNames: ss.getSheets().map(function(s) { return s.getName(); }),
    activeSheet: activeSheet.getName(),
    headers: headers,
  };
}

function getLicenseInfo() { return validateLicense(); }

function getDailyStatus() {
  const licenseInfo = validateLicense();
  const limit = licenseInfo.isPro ? 500 : 50;
  const sent = getTodaySentCount_();
  return { sent: sent, limit: limit, remaining: limit - sent };
}
