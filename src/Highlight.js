/**
 * @fileoverview 差分ハイライト処理モジュール
 * RowDiff[] を受け取り、比較先シートに背景色とコメントを付与する。
 *
 * 色仕様:
 *   追加行   : 緑  (#b7e1cd)
 *   削除行   : 赤  (#f4cccc)  ※比較先シートには存在しないため別行を挿入してマーク
 *   変更セル : 黄  (#fff2cc) + コメントに旧値記載
 */

/** @const {string} 追加行の背景色 */
const COLOR_ADDED = '#b7e1cd';

/** @const {string} 削除行の背景色 */
const COLOR_DELETED = '#f4cccc';

/** @const {string} 変更セルの背景色 */
const COLOR_MODIFIED = '#fff2cc';

/** @const {string} ハイライトなし（リセット用） */
const COLOR_NONE = '#ffffff';

/**
 * 指定シートの既存ハイライトとコメントをすべてクリアする。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function clearHighlights(sheet) {
  const range = sheet.getDataRange();
  range.setBackground(COLOR_NONE);
  // コメントクリア
  const notes = range.getNotes();
  const emptyNotes = notes.map((row) => row.map(() => ''));
  range.setNotes(emptyNotes);
}

/**
 * RowDiff[] に基づき比較先シートにハイライトとコメントを付与する。
 * 削除行は比較先には存在しないためスキップ（Report シートに記録）。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dstSheet - 比較先シート
 * @param {RowDiff[]}                          diffs    - computeDiff の結果
 */
function applyHighlights(dstSheet, diffs) {
  diffs.forEach((diff) => {
    if (diff.type === 'unchanged' || diff.dstRow === null) return;

    // dstRow は 0-based → Sheets API は 1-based
    const sheetRow = diff.dstRow + 1;

    if (diff.type === 'added') {
      // 行全体を緑にする
      const lastCol = Math.max(diff.dstData.length, 1);
      dstSheet
        .getRange(sheetRow, 1, 1, lastCol)
        .setBackground(COLOR_ADDED);
      return;
    }

    if (diff.type === 'modified') {
      diff.cells.forEach((cell) => {
        const cellRange = dstSheet.getRange(sheetRow, cell.col + 1);
        cellRange.setBackground(COLOR_MODIFIED);
        const oldDisplay = cell.oldVal === '' ? '(empty)' : String(cell.oldVal);
        cellRange.setNote(`Old value: ${oldDisplay}`);
      });
    }
  });
}

/**
 * ハイライトをすべてリセットし未差分の状態に戻す。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function resetHighlights(sheet) {
  clearHighlights(sheet);
}
