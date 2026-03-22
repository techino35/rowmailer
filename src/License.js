/**
 * License.js - ライセンス管理（Free/Pro判定）
 */

var LICENSE_KEY_PROP = 'ROWMAILER_LICENSE_KEY';
var LICENSE_STATUS_PROP = 'ROWMAILER_LICENSE_STATUS';
var LICENSE_CACHE_PROP = 'ROWMAILER_LICENSE_CACHE_DATE';

function validateLicense() {
  const props = PropertiesService.getDocumentProperties();
  const licenseKey = props.getProperty(LICENSE_KEY_PROP);
  if (!licenseKey) return { isPro: false, plan: 'Free', message: 'ライセンスキー未設定 - Freeプランで動作中' };

  const cacheDate = props.getProperty(LICENSE_CACHE_PROP);
  const cachedStatus = props.getProperty(LICENSE_STATUS_PROP);
  if (cacheDate && cachedStatus) {
    const hoursElapsed = (new Date().getTime() - new Date(cacheDate).getTime()) / (1000 * 60 * 60);
    if (hoursElapsed < 24) return JSON.parse(cachedStatus);
  }

  const result = /^ROWMAILER-PRO-[A-Z0-9]{16}$/.test(licenseKey)
    ? { isPro: true, plan: 'Pro', message: 'Proプラン有効' }
    : { isPro: false, plan: 'Free', message: '無効なライセンスキー - Freeプランで動作中' };

  props.setProperty(LICENSE_STATUS_PROP, JSON.stringify(result));
  props.setProperty(LICENSE_CACHE_PROP, new Date().toISOString());
  return result;
}

function saveLicenseKey(licenseKey) {
  try {
    const props = PropertiesService.getDocumentProperties();
    if (!licenseKey || licenseKey.trim() === '') {
      [LICENSE_KEY_PROP, LICENSE_STATUS_PROP, LICENSE_CACHE_PROP].forEach(function(k) { props.deleteProperty(k); });
      return { success: true, result: { isPro: false, plan: 'Free', message: 'ライセンスキーを削除しました' } };
    }
    props.setProperty(LICENSE_KEY_PROP, licenseKey.trim().toUpperCase());
    props.deleteProperty(LICENSE_STATUS_PROP);
    props.deleteProperty(LICENSE_CACHE_PROP);
    return { success: true, result: validateLicense() };
  } catch (e) { return { success: false, message: e.message }; }
}
