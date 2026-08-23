# CSVjoiner v2

**JOIN / SPLIT / TRANSFORM**

CSVの分割・結合・整形を、コードを書かずにローカルで処理する軽量GUIツールです。

> CSV加工でExcelを開く回数を減らす。

## What changed

旧版は `A.csv / B.csv / C.csv` の3ファイルと、B.csvだけを明細集約する特定用途の横結合ツールでした。

v2ではその前提を廃止し、CSV処理そのものを3つの操作へ整理しています。

- **JOIN**: 縦結合 / LEFT JOIN / INNER JOIN
- **SPLIT**: 指定列の値ごとにCSVを自動分割
- **TRANSFORM**: 列名・列順・値・日付形式などをGUIで整形

処理はすべて **CHECK / PREVIEW → EXPORT** の順で行い、実行前に結果を確認できます。

## Features

### JOIN

- 複数CSVの縦結合
- 左右で異なるキー列を指定可能
- LEFT JOIN / INNER JOIN
- 一致キー / 左のみ / 右のみを事前集計
- 左右の重複キーを検知
- 多対多JOINを警告
- 結果プレビュー

### SPLIT

- 1列をキーに自動分割
- 分割件数を事前プレビュー
- 空欄キーを除外 / `blank.csv` として出力
- Windowsで使えないファイル名文字を自動置換
- 同名ファイルは `_2`, `_3` のように衝突回避

### TRANSFORM

- 列名変更
- 列削除
- 列順変更
- 固定値列追加
- 文字列置換
- 前後空白削除
- 日付形式変換
- 0埋め
- 複数処理をパイプラインとして順番に適用
- BEFORE / AFTER プレビュー

### Common

- UTF-8 BOM / CP932 / UTF-8 の順でCSVを自動読込
- 文字コードを選んで出力
- CSV以外のファイルは対象外
- ローカル処理のみ

## Setup

```bash
python -m pip install -r requirements.txt
python CSVjoiner.py
```

Python 3.11+ を推奨します。

## Build Windows EXE

```bat
build_windows.bat
```

または手動で:

```bash
python -m pip install pyinstaller
python -m PyInstaller --clean --onefile --noconsole --name CSVjoiner CSVjoiner.py
```

生成先:

```text
dist/CSVjoiner.exe
```

## Repository structure

```text
CSVjoiner/
├─ CSVjoiner.py
├─ csvjoiner/
│  ├─ __init__.py
│  ├─ core.py
│  ├─ models.py
│  ├─ operations.py
│  └─ ui.py
├─ sample/
├─ tests/
├─ requirements.txt
├─ requirements-dev.txt
├─ build_windows.bat
└─ README.md
```

## Design concept

```text
CSV INPUT
   ↓
JOIN / SPLIT / TRANSFORM
   ↓
CONDITION
   ↓
CHECK / PREVIEW
   ↓
EXPORT
```

旧版の「CSVを結合するツール」から、v2では「CSV加工そのものを安全に扱うツール」へ刷新しています。

## License

MIT
