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


## UI refresh

v2ではGUIも作り直し、標準的なタブ型ユーティリティから **左ナビ + ワークベンチ型** の画面構成へ変更しています。

- JOIN / SPLIT / TRANSFORM を左ナビから切り替え
- 入力・設定・CHECK・PREVIEW・EXPORT をカード単位で整理
- JOINは一致 / 左のみ / 右のみ / 重複 / 出力行数をメトリクス表示
- SPLITは出力ファイル数 / 元行数 / 空欄キー行数を事前表示
- TRANSFORMは適用ステップ数と出力列数を表示
- 小さい画面ではワークエリアを縦スクロール可能
- 追加GUIライブラリなし（Tkinter / ttk）

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


## Sample data

`sample/` には、公開用の架空の生徒データを収録しています。実在する人物・学校とは関係ありません。

```text
sample/
├─ students.csv   # 生徒ID / 氏名 / クラス / 入学日
├─ scores.csv     # 受験者ID / 科目 / 点数 / 試験日
└─ clubs.csv      # 会員ID / 部活動 / 役割
```

試し方の例:

- **SPLIT**: `students.csv` を `クラス` で分割して `1-A.csv` / `1-B.csv` / `1-C.csv` を生成
- **JOIN**: `students.csv` の `生徒ID` と `scores.csv` の `受験者ID` を指定して LEFT / INNER JOIN
- **JOIN**: `students.csv` の `生徒ID` と `clubs.csv` の `会員ID` を指定し、左のみ / 右のみを確認
- **TRANSFORM**: `students.csv` の `入学日` を `2026-04-07` 形式へ変換、列名変更や列順変更を試す

`scores.csv` は同じ生徒IDを複数行持つため、JOIN時の **1対多 / 重複キー警告** も確認できます。

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

### 改行を含むCSVについて

- 引用符 (`"`) で囲まれたセル内改行は、そのまま保持して読み込み・出力します。
- プレビューではセル内改行を `↵` と表示し、表の行が崩れないようにしています。
- 出力CSVの行区切りは Windows / Excel で扱いやすい CRLF (`\r\n`) に統一します。
- 引用符が壊れているCSVや、未引用のセル内改行など構文自体が不正なCSVは自動修復しません。元CSVの修正が必要です。

