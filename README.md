# SheetDiff

Google Sheets アドオン — 2つのシートをセル単位で比較し、差分をハイライト＋レポート出力します。

A Google Sheets add-on that compares two sheets cell-by-cell, highlights differences, and generates a structured diff report.

---

## 目次 / Table of Contents

- [機能 / Features](#機能--features)
- [インストール / Installation](#インストール--installation)
- [使い方 / Usage](#使い方--usage)
- [ファイル構成 / File Structure](#ファイル構成--file-structure)
- [開発 / Development](#開発--development)
- [ライセンス / License](#ライセンス--license)

---

## 機能 / Features

| 機能 | 説明 | Feature | Description |
|------|------|---------|-------------|
| シート比較 | 同一ファイル内の2シートをセル単位で比較 | Sheet Compare | Cell-by-cell comparison of two sheets |
| ハイライト | 追加=緑 / 削除=赤 / 変更=黄 | Highlights | Added=green / Deleted=red / Modified=yellow |
| 差分レポート | `_SheetDiff_Report` シートを自動生成 | Diff Report | Auto-generates `_SheetDiff_Report` sheet |
| スナップショット | 現在のシートを hidden シートに保存し後で比較 | Snapshot | Save current state for later comparison |
| キー列指定 | 行のマッチングにキー列を使用（LCS） | Key Column | Row matching by key column via LCS |
| Free / Pro | Free: 100行まで / Pro: 無制限 | Plans | Free: 100 rows / Pro: Unlimited |

---

## インストール / Installation

### Google Workspace Marketplace から

1. Google Sheets を開く
2. 拡張機能 > アドオン > アドオンを取得
3. "SheetDiff" を検索してインストール

### clasp でデプロイ（開発者向け）

```bash
npm install -g @google/clasp
clasp login

cp .clasp.json.example .clasp.json
# .clasp.json の scriptId を編集

clasp push
```

---

## 使い方 / Usage

### シート比較

1. スプレッドシートのメニューから **SheetDiff > Compare Sheets...** を選択
2. 比較元（Source）と比較先（Dest）シートを選択
3. キー列（任意）を入力（例: `ID` または `ID, OrderNo`）
4. **Compare** をクリック

比較完了後:
- 比較先シートに色が付く
- `_SheetDiff_Report` シートが生成される

### スナップショット

1. 保存したいシートをアクティブにする
2. **SheetDiff > Save Snapshot of Active Sheet**
3. 後で **SheetDiff > Compare with Snapshot...** から比較

### ハイライトのクリア

**SheetDiff > Clear Highlights (Active Sheet)**

---

## ファイル構成 / File Structure

```
sheetdiff/
├── src/
│   ├── Code.js        # メインエントリー / Main entry (menu, dialog, orchestration)
│   ├── Diff.js        # 差分算出エンジン / LCS-based diff engine
│   ├── Highlight.js   # ハイライト処理 / Background color & comments
│   ├── Report.js      # レポートシート生成 / Report sheet generator
│   ├── Snapshot.js    # スナップショット / Snapshot save & restore
│   ├── License.js     # Free/Pro 判定 / License validation
│   └── Dialog.html    # 比較設定ダイアログ / Settings dialog
├── mock/
│   └── dialog_mock.html  # ブラウザで確認できる UI モック / Standalone UI preview
├── appsscript.json
├── .clasp.json.example
├── README.md
├── MARKETPLACE.md
└── PRIVACY_POLICY.md
```

---

## 開発 / Development

### 依存関係

- [clasp](https://github.com/google/clasp) — GAS プロジェクト管理
- Node.js 18+

### セットアップ

```bash
npm install -g @google/clasp
clasp login
cp .clasp.json.example .clasp.json
# scriptId を Google Apps Script のプロジェクト ID に書き換える
```

### デプロイ

```bash
clasp push        # ソースをアップロード
clasp open        # ブラウザでプロジェクトを開く
```

### UI モックの確認

`mock/dialog_mock.html` をブラウザで直接開くと、GAS を使わずに UI とサンプルデータを確認できます。

---

## Pro ライセンス

**SheetDiff > Register Pro License...** からライセンスキー（`SDPRO-XXXX-XXXX-XXXX` 形式）を入力してください。

ライセンスキーは `PropertiesService` にキー名 `SHEETDIFF_LICENSE_KEY` として安全に保存されます。

---

## ライセンス / License

MIT License — Copyright (c) 2025 Inoue Shudo
