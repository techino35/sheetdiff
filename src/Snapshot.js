/**
 * @fileoverview スナップショット保存・復元モジュール
 * 指定シートの現在状態を hidden シートにコピーし、後で比較元として利用できる。
 *
 * シート名規則: "_SheetDiff_Snap_<元シート名>_<タイムスタンプ>"
 */

/** @const {string} スナップショットシート名プレフィックス */
const SNAP_PREFIX = '_SheetDiff_Snap_';

/**
 * 指定シートのスナップショットを作成する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet - コピー元シート
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 作成したスナップショットシート
 */
function saveSnapshot(sourceSheet) {
  const ss = sourceSheet.getParent();
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const snapName = `${SNAP_PREFIX}${sourceSheet.getName()}_${ts}`;

  const snap = sourceSheet.copyTo(ss);
  snap.setName(snapName);
  snap.hideSheet();
  return snap;
}

/**
 * 現在のスプレッドシートに存在するスナップショット一覧を返す。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {{name: string, sheetId: number, createdAt: string}[]}
 */
function listSnapshots(ss) {
  return ss
    .getSheets()
    .filter((s) => s.getName().startsWith(SNAP_PREFIX))
    .map((s) => {
      const parts = s.getName().replace(SNAP_PREFIX, '').split('_');
      const createdAt = parts.slice(-2).join('_');
      const sourceName = parts.slice(0, -2).join('_');
      return { name: s.getName(), sheetId: s.getSheetId(), sourceName, createdAt };
    });
}

/**
 * スナップショットシートを名前で取得する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} snapSheetName - スナップショットシート名
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getSnapshot(ss, snapSheetName) {
  return ss.getSheetByName(snapSheetName) || null;
}

/**
 * スナップショットを削除する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} snapSheetName
 * @return {boolean} 削除成功なら true
 */
function deleteSnapshot(ss, snapSheetName) {
  const sheet = ss.getSheetByName(snapSheetName);
  if (!sheet) return false;
  ss.deleteSheet(sheet);
  return true;
}

/**
 * スナップショットシートのデータ（2次元配列）を取得する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} snapSheet
 * @return {Array[]}
 */
function getSnapshotData(snapSheet) {
  const range = snapSheet.getDataRange();
  return range.getValues();
}
