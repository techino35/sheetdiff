/**
 * @fileoverview 差分ハイライト処理モジュール
 */

const COLOR_ADDED = '#b7e1cd';
const COLOR_DELETED = '#f4cccc';
const COLOR_MODIFIED = '#fff2cc';
const COLOR_NONE = '#ffffff';

function clearHighlights(sheet) {
  const range = sheet.getDataRange();
  range.setBackground(COLOR_NONE);
  const notes = range.getNotes();
  const emptyNotes = notes.map((row) => row.map(() => ''));
  range.setNotes(emptyNotes);
}

function applyHighlights(dstSheet, diffs) {
  diffs.forEach((diff) => {
    if (diff.type === 'unchanged' || diff.dstRow === null) return;

    const sheetRow = diff.dstRow + 1;

    if (diff.type === 'added') {
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

function resetHighlights(sheet) {
  clearHighlights(sheet);
}