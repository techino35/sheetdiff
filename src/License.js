/**
 * @fileoverview Free/Pro ライセンス判定モジュール
 */

const FREE_ROW_LIMIT = 100;

function isPro() {
  const props = PropertiesService.getUserProperties();
  const key = props.getProperty('SHEETDIFF_LICENSE_KEY');
  return key !== null && key.trim().length > 0;
}

function registerLicense(licenseKey) {
  if (!licenseKey || licenseKey.trim().length === 0) {
    return { success: false, message: 'License key is empty.' };
  }
  if (!licenseKey.trim().startsWith('SDPRO-')) {
    return { success: false, message: 'Invalid license key format.' };
  }
  PropertiesService.getUserProperties().setProperty('SHEETDIFF_LICENSE_KEY', licenseKey.trim());
  return { success: true, message: 'License registered successfully. Pro features unlocked.' };
}

function revokeLicense() {
  PropertiesService.getUserProperties().deleteProperty('SHEETDIFF_LICENSE_KEY');
}

function checkRowLimit(rowCount) {
  if (isPro()) return null;
  if (rowCount > FREE_ROW_LIMIT) {
    return (
      `Free plan is limited to ${FREE_ROW_LIMIT} rows. ` +
      `Your sheet has ${rowCount} rows. ` +
      `Upgrade to Pro for unlimited rows.`
    );
  }
  return null;
}