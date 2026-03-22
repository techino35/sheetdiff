/**
 * @fileoverview 差分レポートシート生成モジュール
 */

const REPORT_SHEET_NAME = '_SheetDiff_Report';

const REPORT_HEADERS = [
  'Cell',
  'Type',
  'Row (dst)',
  'Column',
  'Old Value',
  'New Value',
];

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

function generateReport(ss, diffs, srcSheetName, dstSheetName) {
  const sheet = createReportSheet_(ss);
  const rows = [];

  diffs.forEach((diff) => {
    if (diff.type === 'unchanged') return;

    if (diff.type === 'added') {
      const rowNum = diff.dstRow + 1;
      rows.push([`Row ${rowNum}`, 'Added', rowNum, '', '', diff.dstData.join(', ')]);
      return;
    }

    if (diff.type === 'deleted') {
      const rowNum = diff.srcRow + 1;
      rows.push([`Row ${rowNum} (src)`, 'Deleted', '', '', diff.srcData.join(', '), '']);
      return;
    }

    if (diff.type === 'modified') {
      diff.cells.forEach((cell) => {
        const rowNum = cell.row + 1;
        const colLetter = colToLetter_(cell.col);
        rows.push([`${colLetter}${rowNum}`, 'Modified', rowNum, colLetter, cell.oldVal, cell.newVal]);
      });
    }
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, REPORT_HEADERS.length).setValues(rows);
    for (let i = 0; i < rows.length; i++) {
      const type = rows[i][1];
      const color =
        type === 'Added' ? COLOR_ADDED : type === 'Deleted' ? COLOR_DELETED : COLOR_MODIFIED;
      sheet.getRange(i + 2, 1, 1, REPORT_HEADERS.length).setBackground(color);
    }
  }

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.getRange(1, 1).setNote(
    `Generated: ${now}\nSource: ${srcSheetName}\nDest: ${dstSheetName}`
  );

  sheet.autoResizeColumns(1, REPORT_HEADERS.length);
  return sheet;
}