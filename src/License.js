/**
 * @fileoverview Free/Pro ライセンス判定モジュール
 * PropertiesService に SHEETDIFF_LICENSE_KEY を保存してPro判定を行う。
 */

/** Free プランの最大行数制限 */
const FREE_ROW_LIMIT = 100;

/**
 * 現在のユーザーが Pro ライセンスを持っているか判定する。
 * @return {boolean} Pro ライセンスを持つ場合 true
 */
function isPro() {
  const props = PropertiesService.getUserProperties();
  const key = props.getProperty('SHEETDIFF_LICENSE_KEY');
  return key !== null && key.trim().length > 0;
}

/**
 * Pro ライセンスキーを登録する。
 * @param {string} licenseKey - ライセンスキー文字列
 * @return {{success: boolean, message: string}}
 */
function registerLicense(licenseKey) {
  if (!licenseKey || licenseKey.trim().length === 0) {
    return { success: false, message: 'License key is empty.' };
  }
  // 簡易検証: "SDPRO-" プレフィックスを期待する（本番では外部API検証推奨）
  if (!licenseKey.trim().startsWith('SDPRO-')) {
    return { success: false, message: 'Invalid license key format.' };
  }
  PropertiesService.getUserProperties().setProperty('SHEETDIFF_LICENSE_KEY', licenseKey.trim());
  return { success: true, message: 'License registered successfully. Pro features unlocked.' };
}

/**
 * ライセンスキーを削除し Free プランに戻す。
 */
function revokeLicense() {
  PropertiesService.getUserProperties().deleteProperty('SHEETDIFF_LICENSE_KEY');
}

/**
 * 行数制限チェック。Free プランで上限を超える場合はエラーメッセージを返す。
 * @param {number} rowCount - 処理対象の行数
 * @return {string|null} エラーメッセージ、問題なければ null
 */
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
