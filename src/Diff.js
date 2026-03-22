/**
 * @fileoverview 差分算出エンジン
 * LCS (Longest Common Subsequence) ベースの行マッチング + セル単位比較を実装する。
 */

/**
 * @typedef {Object} CellDiff
 * @property {number} row    - 比較先シート上の行インデックス (0-based)
 * @property {number} col    - 列インデックス (0-based)
 * @property {*}      oldVal - 旧値（比較元）
 * @property {*}      newVal - 新値（比較先）
 */

/**
 * @typedef {Object} RowDiff
 * @property {'added'|'deleted'|'modified'|'unchanged'} type
 * @property {number|null}   srcRow   - 比較元の行インデックス (0-based)、追加行は null
 * @property {number|null}   dstRow   - 比較先の行インデックス (0-based)、削除行は null
 * @property {CellDiff[]}    cells    - 変更されたセルの一覧（type が modified の場合のみ）
 * @property {Array}         srcData  - 比較元の行データ
 * @property {Array}         dstData  - 比較先の行データ
 */

/**
 * 2次元配列から指定列の値でキーを生成する。
 * @param {Array} row       - 行データ
 * @param {number[]} keyCols - キー列インデックス配列 (0-based)
 * @return {string} キー文字列
 */
function buildRowKey_(row, keyCols) {
  return keyCols.map((c) => String(row[c] !== undefined ? row[c] : '')).join('\x00');
}

/**
 * LCS テーブルを構築して一致行インデックスペアを返す。
 * キーが指定されている場合はキーベースマッチング、なければ行番号ベース。
 * @param {Array[]} srcRows  - 比較元行配列
 * @param {Array[]} dstRows  - 比較先行配列
 * @param {number[]} keyCols - キー列インデックス配列 (0-based)。空配列なら行番号ベース
 * @return {{srcIdx: number, dstIdx: number}[]} マッチしたペア
 */
function matchRows_(srcRows, dstRows, keyCols) {
  if (keyCols.length === 0) {
    // 行番号ベース: 両方に存在するインデックスをそのままペアにする
    const len = Math.min(srcRows.length, dstRows.length);
    const pairs = [];
    for (let i = 0; i < len; i++) {
      pairs.push({ srcIdx: i, dstIdx: i });
    }
    return pairs;
  }

  // キーベース LCS
  const srcKeys = srcRows.map((r) => buildRowKey_(r, keyCols));
  const dstKeys = dstRows.map((r) => buildRowKey_(r, keyCols));

  const m = srcKeys.length;
  const n = dstKeys.length;

  // LCS DP テーブル（メモリ節約のため 1D rolling を使用）
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (srcKeys[i - 1] === dstKeys[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1] + 1;
      } else {
        dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }
  }

  // バックトレース
  const pairs = [];
  let i = m, j = n;
  while (i > 0 && j > 0) {
    if (srcKeys[i - 1] === dstKeys[j - 1]) {
      pairs.unshift({ srcIdx: i - 1, dstIdx: j - 1 });
      i--;
      j--;
    } else if (dp[i - 1][j] >= dp[i][j - 1]) {
      i--;
    } else {
      j--;
    }
  }
  return pairs;
}

/**
 * 2行のセルを比較し差異のあるセルを返す。
 * @param {Array}   srcRow - 比較元の行データ
 * @param {Array}   dstRow - 比較先の行データ
 * @param {number}  dstRowIdx - 比較先行インデックス
 * @return {CellDiff[]}
 */
function compareCells_(srcRow, dstRow, dstRowIdx) {
  const cols = Math.max(srcRow.length, dstRow.length);
  const diffs = [];
  for (let c = 0; c < cols; c++) {
    const oldVal = srcRow[c] !== undefined ? srcRow[c] : '';
    const newVal = dstRow[c] !== undefined ? dstRow[c] : '';
    if (String(oldVal) !== String(newVal)) {
      diffs.push({ row: dstRowIdx, col: c, oldVal, newVal });
    }
  }
  return diffs;
}

/**
 * 2つの2次元データ配列を比較し RowDiff[] を返す。
 * @param {Array[]} srcData  - 比較元データ（2次元配列）
 * @param {Array[]} dstData  - 比較先データ（2次元配列）
 * @param {number[]} keyCols - キー列インデックス配列 (0-based)
 * @return {RowDiff[]}
 */
function computeDiff(srcData, dstData, keyCols) {
  const matchedPairs = matchRows_(srcData, dstData, keyCols);

  const matchedSrcSet = new Set(matchedPairs.map((p) => p.srcIdx));
  const matchedDstSet = new Set(matchedPairs.map((p) => p.dstIdx));

  // マッチペアを dstIdx 順に並べて差分リストを構築
  const pairBydst = new Map(matchedPairs.map((p) => [p.dstIdx, p]));

  /** @type {RowDiff[]} */
  const result = [];

  // dstData を走査しながら、マッチしていない src 行（削除）を挿入
  let srcPointer = 0;
  for (let di = 0; di < dstData.length; di++) {
    const pair = pairBydst.get(di);

    if (pair) {
      // マッチ前の削除行を挿入
      while (srcPointer < pair.srcIdx) {
        if (!matchedSrcSet.has(srcPointer)) {
          result.push({
            type: 'deleted',
            srcRow: srcPointer,
            dstRow: null,
            cells: [],
            srcData: srcData[srcPointer],
            dstData: [],
          });
        }
        srcPointer++;
      }
      srcPointer = pair.srcIdx + 1;

      // セル比較
      const cellDiffs = compareCells_(srcData[pair.srcIdx], dstData[di], di);
      if (cellDiffs.length > 0) {
        result.push({
          type: 'modified',
          srcRow: pair.srcIdx,
          dstRow: di,
          cells: cellDiffs,
          srcData: srcData[pair.srcIdx],
          dstData: dstData[di],
        });
      } else {
        result.push({
          type: 'unchanged',
          srcRow: pair.srcIdx,
          dstRow: di,
          cells: [],
          srcData: srcData[pair.srcIdx],
          dstData: dstData[di],
        });
      }
    } else {
      // 追加行
      result.push({
        type: 'added',
        srcRow: null,
        dstRow: di,
        cells: [],
        srcData: [],
        dstData: dstData[di],
      });
    }
  }

  // 残りの削除行を末尾に追加
  while (srcPointer < srcData.length) {
    if (!matchedSrcSet.has(srcPointer)) {
      result.push({
        type: 'deleted',
        srcRow: srcPointer,
        dstRow: null,
        cells: [],
        srcData: srcData[srcPointer],
        dstData: [],
      });
    }
    srcPointer++;
  }

  return result;
}

/**
 * キー列インデックス配列をヘッダー行から解決する。
 * @param {string[]} keyColNames - キー列名（ヘッダー文字列）
 * @param {Array}    headerRow   - ヘッダー行データ
 * @return {number[]} キー列インデックス (0-based)
 */
function resolveKeyColumns(keyColNames, headerRow) {
  if (!keyColNames || keyColNames.length === 0) return [];
  return keyColNames
    .map((name) => headerRow.indexOf(name))
    .filter((idx) => idx >= 0);
}
