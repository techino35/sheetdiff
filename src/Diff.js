/**
 * @fileoverview 差分算出エンジン
 * LCS (Longest Common Subsequence) ベースの行マッチング + セル単位比較を実装する。
 */

function buildRowKey_(row, keyCols) {
  return keyCols.map((c) => String(row[c] !== undefined ? row[c] : '')).join('\x00');
}

function matchRows_(srcRows, dstRows, keyCols) {
  if (keyCols.length === 0) {
    const len = Math.min(srcRows.length, dstRows.length);
    const pairs = [];
    for (let i = 0; i < len; i++) {
      pairs.push({ srcIdx: i, dstIdx: i });
    }
    return pairs;
  }

  const srcKeys = srcRows.map((r) => buildRowKey_(r, keyCols));
  const dstKeys = dstRows.map((r) => buildRowKey_(r, keyCols));

  const m = srcKeys.length;
  const n = dstKeys.length;

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

function computeDiff(srcData, dstData, keyCols) {
  const matchedPairs = matchRows_(srcData, dstData, keyCols);

  const matchedSrcSet = new Set(matchedPairs.map((p) => p.srcIdx));
  const matchedDstSet = new Set(matchedPairs.map((p) => p.dstIdx));

  const pairBydst = new Map(matchedPairs.map((p) => [p.dstIdx, p]));

  const result = [];

  let srcPointer = 0;
  for (let di = 0; di < dstData.length; di++) {
    const pair = pairBydst.get(di);

    if (pair) {
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

function resolveKeyColumns(keyColNames, headerRow) {
  if (!keyColNames || keyColNames.length === 0) return [];
  return keyColNames
    .map((name) => headerRow.indexOf(name))
    .filter((idx) => idx >= 0);
}