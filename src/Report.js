/**
 * @fileoverview 差分レポートシート生成モジュール
 * RowDiff[] を受け取り "_SheetDiff_Report" シートを生成する。
 */

/** @const {string} レポートシート名 */
const REPORT_SHEET_NAME = '_SheetDiff_Report';

/** @const {string[]} レポートヘッダー */
const REPORT_HEADERS = [
  'Cell',
  'Type',
  'Row (dst)',
  'Column',
  'Old Value',
  'New Value',
];

/**
 * 既存のレポートシートを削除して新規作成し、ヘッダーを設定する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - 対象スプレッドシート
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 新規レポートシート
 */
function createReportSheet_(ss) {
  const existing = ss.getSheetByName(REPORT_SHEET_NAME);
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet(REPORT_SHEET_NAME);
  const headerRange = sheet.getRange(1, 1, 1, REPORT_HEADERS.length);
  headerRange.setValues([REPORT_HEADERS]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#cfe2f3');
  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * 列インデックス (0-based) を A1 記法の列文字に変換する。
 * @param {number} colIndex - 0-based 列インデックス
 * @return {string} 列文字（例: 0 → "A", 26 → "AA"）
 */
function colToLetter_(colIndex) {
  let letter = '';
  let n = colIndex + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

/**
 * RowDiff[] からレポートシートを生成する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss   - 対象スプレッドシート
 * @param {RowDiff[]}                                diffs - computeDiff の結果
 * @param {string}                                   srcSheetName - 比較元シート名（表示用）
 * @param {string}                                   dstSheetName - 比較先シート名（表示用）
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 生成したレポートシート
 */
function generateReport(ss, diffs, srcSheetName, dstSheetName) {
  const sheet = createReportSheet_(ss);
  const rows = [];

  diffs.forEach((diff) => {
    if (diff.type === 'unchanged') return;

    if (diff.type === 'added') {
      const rowNum = diff.dstRow + 1;
      rows.push([
        `Row ${rowNum}`,
        'Added',
        rowNum,
        '',
        '',
        diff.dstData.join(', '),
      ]);
      return;
    }

    if (diff.type === 'deleted') {
      const rowNum = diff.srcRow + 1;
      rows.push([
        `Row ${rowNum} (src)`,
        'Deleted',
        '',
        '',
        diff.srcData.join(', '),
        '',
      ]);
      return;
    }

    if (diff.type === 'modified') {
      diff.cells.forEach((cell) => {
        const rowNum = cell.row + 1;
        const colLetter = colToLetter_(cell.col);
        rows.push([
          `${colLetter}${rowNum}`,
          'Modified',
          rowNum,
          colLetter,
          cell.oldVal,
          cell.newVal,
        ]);
      });
    }
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, REPORT_HEADERS.length).setValues(rows);
    // 種別列に色を付ける
    for (let i = 0; i < rows.length; i++) {
      const type = rows[i][1];
      const color =
        type === 'Added' ? COLOR_ADDED : type === 'Deleted' ? COLOR_DELETED : COLOR_MODIFIED;
      sheet.getRange(i + 2, 1, 1, REPORT_HEADERS.length).setBackground(color);
    }
  }

  // メタ情報をシート先頭コメントに残す
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.getRange(1, 1).setNote(
    `Generated: ${now}\nSource: ${srcSheetName}\nDest: ${dstSheetName}`
  );

  sheet.autoResizeColumns(1, REPORT_HEADERS.length);
  return sheet;
}
