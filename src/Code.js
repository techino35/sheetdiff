/**
 * @fileoverview SheetDiff - メインエントリーポイント
 * カスタムメニューの登録、ダイアログの表示、比較処理のオーケストレーションを行う。
 */

/**
 * スプレッドシート開封時にカスタムメニューを追加する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SheetDiff')
    .addItem('Compare Sheets...', 'showCompareDialog')
    .addSeparator()
    .addItem('Save Snapshot of Active Sheet', 'saveActiveSnapshot')
    .addItem('Compare with Snapshot...', 'showSnapshotDialog')
    .addSeparator()
    .addItem('Clear Highlights (Active Sheet)', 'clearActiveHighlights')
    .addItem('Delete Report Sheet', 'deleteReportSheet')
    .addSeparator()
    .addItem('Register Pro License...', 'showLicenseDialog')
    .addToUi();
}

// ---------------------------------------------------------------------------
// ダイアログ表示
// ---------------------------------------------------------------------------

/**
 * シート比較ダイアログを表示する。
 */
function showCompareDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(520)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'SheetDiff - Compare Sheets');
}

/**
 * スナップショット選択ダイアログを表示する。
 */
function showSnapshotDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snaps = listSnapshots(ss);
  if (snaps.length === 0) {
    SpreadsheetApp.getUi().alert('No snapshots found. Save a snapshot first.');
    return;
  }
  // スナップショット一覧を Dialog に渡してシンプルに処理
  const html = HtmlService.createHtmlOutput(buildSnapshotDialogHtml_(snaps))
    .setWidth(480)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'SheetDiff - Compare with Snapshot');
}

/**
 * ライセンス登録ダイアログを表示する。
 */
function showLicenseDialog() {
  const html = HtmlService.createHtmlOutput(buildLicenseDialogHtml_())
    .setWidth(400)
    .setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, 'SheetDiff - License');
}

// ---------------------------------------------------------------------------
// ダイアログから呼ばれる処理（google.script.run 経由）
// ---------------------------------------------------------------------------

/**
 * ダイアログから受け取った設定でシート比較を実行する。
 * @param {Object} config - ダイアログからの設定オブジェクト
 * @param {string} config.srcSheetName   - 比較元シート名
 * @param {string} config.dstSheetName   - 比較先シート名
 * @param {string} config.keyColumns     - キー列名をカンマ区切りにした文字列
 * @param {boolean} config.hasHeader     - 1行目をヘッダーとして扱うか
 * @return {{success: boolean, message: string, summary: Object}}
 */
function runCompare(config) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const srcSheet = ss.getSheetByName(config.srcSheetName);
    const dstSheet = ss.getSheetByName(config.dstSheetName);

    if (!srcSheet) return { success: false, message: `Source sheet "${config.srcSheetName}" not found.` };
    if (!dstSheet) return { success: false, message: `Dest sheet "${config.dstSheetName}" not found.` };

    const srcAllData = srcSheet.getDataRange().getValues();
    const dstAllData = dstSheet.getDataRange().getValues();

    let srcData = srcAllData;
    let dstData = dstAllData;
    let headerRow = [];

    if (config.hasHeader && srcAllData.length > 0) {
      headerRow = srcAllData[0];
      srcData = srcAllData.slice(1);
      dstData = dstAllData.slice(1);
    }

    // 行数制限チェック
    const maxRows = Math.max(srcData.length, dstData.length);
    const limitError = checkRowLimit(maxRows);
    if (limitError) return { success: false, message: limitError };

    // キー列解決
    const keyColNames = config.keyColumns
      ? config.keyColumns.split(',').map((s) => s.trim()).filter(Boolean)
      : [];
    const keyCols = resolveKeyColumns(keyColNames, headerRow);

    // 差分算出
    const diffs = computeDiff(srcData, dstData, keyCols);

    // ハイライト適用
    clearHighlights(dstSheet);
    applyHighlights(dstSheet, diffs);

    // レポート生成
    const reportSheet = generateReport(ss, diffs, config.srcSheetName, config.dstSheetName);

    // サマリー
    const added    = diffs.filter((d) => d.type === 'added').length;
    const deleted  = diffs.filter((d) => d.type === 'deleted').length;
    const modified = diffs.filter((d) => d.type === 'modified').length;

    return {
      success: true,
      message: 'Comparison complete.',
      summary: { added, deleted, modified, reportSheet: reportSheet.getName() },
    };
  } catch (e) {
    return { success: false, message: `Error: ${e.message}` };
  }
}

/**
 * スナップショットとアクティブシートを比較する。
 * @param {string} snapSheetName - スナップショットシート名
 * @param {string} dstSheetName  - 比較先シート名
 * @param {string} keyColumns    - キー列名カンマ区切り
 * @param {boolean} hasHeader
 * @return {{success: boolean, message: string, summary: Object}}
 */
function runCompareWithSnapshot(snapSheetName, dstSheetName, keyColumns, hasHeader) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snapSheet = getSnapshot(ss, snapSheetName);
  if (!snapSheet) return { success: false, message: 'Snapshot not found.' };

  return runCompare({
    srcSheetName: snapSheetName,
    dstSheetName,
    keyColumns,
    hasHeader,
  });
}

/**
 * ダイアログ用にシート名一覧を返す（非表示含む）。
 * @return {string[]}
 */
function getSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .filter((s) => !s.getName().startsWith(SNAP_PREFIX) && s.getName() !== REPORT_SHEET_NAME)
    .map((s) => s.getName());
}

/**
 * ライセンス登録を実行する（ダイアログから呼ばれる）。
 * @param {string} key
 * @return {{success: boolean, message: string}}
 */
function submitLicense(key) {
  return registerLicense(key);
}

/**
 * 現在の Pro ステータスを返す。
 * @return {{isPro: boolean}}
 */
function getLicenseStatus() {
  return { isPro: isPro() };
}

// ---------------------------------------------------------------------------
// メニューから直接呼ばれる処理
// ---------------------------------------------------------------------------

/**
 * アクティブシートのスナップショットを保存する。
 */
function saveActiveSnapshot() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const snap = saveSnapshot(sheet);
  SpreadsheetApp.getUi().alert(`Snapshot saved: ${snap.getName()}`);
}

/**
 * アクティブシートのハイライトをクリアする。
 */
function clearActiveHighlights() {
  clearHighlights(SpreadsheetApp.getActiveSheet());
  SpreadsheetApp.getUi().alert('Highlights cleared.');
}

/**
 * レポートシートを削除する。
 */
function deleteReportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPORT_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No report sheet found.');
    return;
  }
  ss.deleteSheet(sheet);
  SpreadsheetApp.getUi().alert('Report sheet deleted.');
}

// ---------------------------------------------------------------------------
// 内部ヘルパー: ダイアログ HTML 生成
// ---------------------------------------------------------------------------

/**
 * スナップショット選択ダイアログの HTML を生成する。
 * @param {{name: string, sourceName: string, createdAt: string}[]} snaps
 * @return {string}
 */
function buildSnapshotDialogHtml_(snaps) {
  const sheetNames = getSheetNames();
  const snapOptions = snaps
    .map((s) => `<option value="${s.name}">${s.sourceName} @ ${s.createdAt}</option>`)
    .join('');
  const sheetOptions = sheetNames
    .map((n) => `<option value="${n}">${n}</option>`)
    .join('');

  return `<!DOCTYPE html><html><head>
<style>body{font-family:Arial,sans-serif;padding:16px;}label{display:block;margin-top:12px;font-size:13px;}select,input{width:100%;padding:6px;box-sizing:border-box;}button{margin-top:16px;padding:8px 16px;background:#4a86e8;color:#fff;border:none;border-radius:4px;cursor:pointer;}</style>
</head><body>
<h3 style="margin:0 0 12px">Compare with Snapshot</h3>
<label>Snapshot (source):<select id="snap">${snapOptions}</select></label>
<label>Current Sheet (dest):<select id="dst">${sheetOptions}</select></label>
<label>Key Columns (comma-separated, leave empty for row-number matching):<input id="keys" placeholder="e.g. ID, Name"/></label>
<label><input type="checkbox" id="hdr" checked> First row is header</label>
<button onclick="run()">Compare</button>
<div id="msg" style="margin-top:12px;color:#666;font-size:12px;"></div>
<script>
function run(){
  var snap=document.getElementById('snap').value;
  var dst=document.getElementById('dst').value;
  var keys=document.getElementById('keys').value;
  var hdr=document.getElementById('hdr').checked;
  document.getElementById('msg').textContent='Running...';
  google.script.run.withSuccessHandler(function(r){
    if(r.success){
      var s=r.summary;
      document.getElementById('msg').textContent='Done! Added:'+s.added+' Deleted:'+s.deleted+' Modified:'+s.modified;
      setTimeout(function(){google.script.host.close();},2000);
    } else {
      document.getElementById('msg').style.color='#e00';
      document.getElementById('msg').textContent=r.message;
    }
  }).runCompareWithSnapshot(snap,dst,keys,hdr);
}
</script></body></html>`;
}

/**
 * ライセンス登録ダイアログの HTML を生成する。
 * @return {string}
 */
function buildLicenseDialogHtml_() {
  return `<!DOCTYPE html><html><head>
<style>body{font-family:Arial,sans-serif;padding:16px;}label{display:block;margin-top:12px;font-size:13px;}input{width:100%;padding:6px;box-sizing:border-box;}button{margin-top:16px;padding:8px 16px;background:#4a86e8;color:#fff;border:none;border-radius:4px;cursor:pointer;}</style>
</head><body>
<h3 style="margin:0 0 12px">SheetDiff License</h3>
<div id="status" style="font-size:13px;color:#666;margin-bottom:8px;">Checking...</div>
<label>License Key:<input id="key" placeholder="SDPRO-XXXX-XXXX-XXXX"/></label>
<button onclick="submit()">Register</button>
<div id="msg" style="margin-top:12px;font-size:12px;"></div>
<script>
google.script.run.withSuccessHandler(function(r){
  document.getElementById('status').textContent=r.isPro?'Status: PRO':'Status: Free (100 row limit)';
}).getLicenseStatus();
function submit(){
  var k=document.getElementById('key').value;
  google.script.run.withSuccessHandler(function(r){
    var el=document.getElementById('msg');
    el.style.color=r.success?'green':'#e00';
    el.textContent=r.message;
    if(r.success)setTimeout(function(){google.script.host.close();},1500);
  }).submitLicense(k);
}
</script></body></html>`;
}
