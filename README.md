# SheetDiff

Google Sheets アドオン — 2つのシートをセル単位で比較し、差分をハイライト＋レポート出力します。

A Google Sheets add-on that compares two sheets cell-by-cell, highlights differences, and generates a structured diff report.

---

## Features

| Feature | Free | Pro |
|---------|------|-----|
| Row limit | 100 rows | Unlimited |
| Visual highlights (green/yellow) | Yes | Yes |
| Diff report sheet | Yes | Yes |
| Snapshot mode | Yes | Yes |
| Key column matching (LCS) | Yes | Yes |

## Setup

```bash
npm install -g @google/clasp
clasp login
cp .clasp.json.example .clasp.json
# Edit scriptId in .clasp.json
clasp push
```

## License

MIT License — Copyright (c) 2025 Inoue Shudo
