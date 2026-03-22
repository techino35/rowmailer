/**
 * Template.js - {{列名}}プレースホルダー置換エンジン
 */

function replacePlaceholders(template, rowData) {
  if (!template) return '';
  if (!rowData) return template;
  return template.replace(/\{\{([^}]+)\}\}/g, function(match, key) {
    const trimmedKey = key.trim();
    if (Object.prototype.hasOwnProperty.call(rowData, trimmedKey)) {
      return formatCellValue_(rowData[trimmedKey]);
    }
    return match;
  });
}

function formatCellValue_(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  if (typeof value === 'boolean') return value ? 'はい' : 'いいえ';
  return String(value);
}

function extractPlaceholderKeys(template) {
  if (!template) return [];
  const keys = [];
  const regex = /\{\{([^}]+)\}\}/g;
  let match;
  while ((match = regex.exec(template)) !== null) {
    const key = match[1].trim();
    if (keys.indexOf(key) === -1) keys.push(key);
  }
  return keys;
}

function previewTemplate(subjectTemplate, bodyTemplate) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues();
  if (data.length < 2) return { subject: subjectTemplate, body: bodyTemplate, missingKeys: [], error: 'データ行がありません' };
  const headers = data[0];
  const rowObj = {};
  headers.forEach(function(h, i) { rowObj[h] = data[1][i] !== undefined ? data[1][i] : ''; });
  const allKeys = extractPlaceholderKeys(subjectTemplate).concat(extractPlaceholderKeys(bodyTemplate));
  return {
    subject: replacePlaceholders(subjectTemplate, rowObj),
    body: replacePlaceholders(bodyTemplate, rowObj),
    missingKeys: allKeys.filter(function(k) { return headers.indexOf(k) === -1; }),
  };
}

function saveTemplate(name, template) {
  if (!validateLicense().isPro) return { success: false, message: 'Proプランの機能です。' };
  try {
    const props = PropertiesService.getDocumentProperties();
    const lib = JSON.parse(props.getProperty('ROWMAILER_TEMPLATES') || '{}');
    lib[name] = template;
    props.setProperty('ROWMAILER_TEMPLATES', JSON.stringify(lib));
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function getTemplateLibrary() {
  if (!validateLicense().isPro) return {};
  try { return JSON.parse(PropertiesService.getDocumentProperties().getProperty('ROWMAILER_TEMPLATES') || '{}'); }
  catch (e) { return {}; }
}

function deleteTemplate(name) {
  if (!validateLicense().isPro) return { success: false, message: 'Proプランの機能です。' };
  try {
    const props = PropertiesService.getDocumentProperties();
    const lib = JSON.parse(props.getProperty('ROWMAILER_TEMPLATES') || '{}');
    delete lib[name];
    props.setProperty('ROWMAILER_TEMPLATES', JSON.stringify(lib));
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}
