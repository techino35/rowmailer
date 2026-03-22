/**
 * Trigger.js - タイマートリガー管理
 */

var TRIGGER_HANDLER = 'runMailer';
var TRIGGER_KEY = 'ROWMAILER_TRIGGER_ID';

function getTriggerInfo() {
  const props = PropertiesService.getDocumentProperties();
  const triggerId = props.getProperty(TRIGGER_KEY);
  if (!triggerId) return { exists: false };
  const trigger = ScriptApp.getProjectTriggers().find(function(t) { return t.getUniqueId() === triggerId; });
  if (!trigger) { props.deleteProperty(TRIGGER_KEY); return { exists: false }; }
  return {
    exists: true, triggerId: triggerId,
    intervalType: props.getProperty('ROWMAILER_TRIGGER_INTERVAL') || 'hourly',
    handlerFunction: trigger.getHandlerFunction(),
  };
}

function setTrigger(intervalType) {
  try {
    deleteTrigger();
    let trigger;
    switch (intervalType) {
      case 'hourly': trigger = ScriptApp.newTrigger(TRIGGER_HANDLER).timeBased().everyHours(1).create(); break;
      case 'every6hours': trigger = ScriptApp.newTrigger(TRIGGER_HANDLER).timeBased().everyHours(6).create(); break;
      case 'daily': trigger = ScriptApp.newTrigger(TRIGGER_HANDLER).timeBased().everyDays(1).atHour(9).nearMinute(0).inTimezone('Asia/Tokyo').create(); break;
      default: return { success: false, message: '不明なインターバル: ' + intervalType };
    }
    const props = PropertiesService.getDocumentProperties();
    props.setProperty(TRIGGER_KEY, trigger.getUniqueId());
    props.setProperty('ROWMAILER_TRIGGER_INTERVAL', intervalType);
    return { success: true, triggerId: trigger.getUniqueId(), intervalType: intervalType };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteTrigger() {
  try {
    const props = PropertiesService.getDocumentProperties();
    const triggerId = props.getProperty(TRIGGER_KEY);
    if (triggerId) {
      ScriptApp.getProjectTriggers().forEach(function(t) {
        if (t.getUniqueId() === triggerId) ScriptApp.deleteTrigger(t);
      });
      props.deleteProperty(TRIGGER_KEY);
      props.deleteProperty('ROWMAILER_TRIGGER_INTERVAL');
    }
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === TRIGGER_HANDLER) ScriptApp.deleteTrigger(t);
    });
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}
