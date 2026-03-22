/**
 * @fileoverview スナップショット保存・復元モジュール
 */

const SNAP_PREFIX = '_SheetDiff_Snap_';

function saveSnapshot(sourceSheet) {
  const ss = sourceSheet.getParent();
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const snapName = `${SNAP_PREFIX}${sourceSheet.getName()}_${ts}`;

  const snap = sourceSheet.copyTo(ss);
  snap.setName(snapName);
  snap.hideSheet();
  return snap;
}

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

function getSnapshot(ss, snapSheetName) {
  return ss.getSheetByName(snapSheetName) || null;
}

function deleteSnapshot(ss, snapSheetName) {
  const sheet = ss.getSheetByName(snapSheetName);
  if (!sheet) return false;
  ss.deleteSheet(sheet);
  return true;
}

function getSnapshotData(snapSheet) {
  const range = snapSheet.getDataRange();
  return range.getValues();
}