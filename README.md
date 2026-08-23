# CSVjoiner v2

**JOIN / SPLIT / TRANSFORM**

CSVの結合・分割・整形を、コードを書かずにローカルで処理するWindows向け軽量GUIツールです。

> CSV加工でExcelを開く回数を減らす。

## v2

旧版は `A.csv / B.csv / C.csv` と特定の列構成を前提にした結合ツールでした。
v2では業務固有の前提を外し、CSV加工を3つの操作へ整理しています。

- **JOIN** — CSVをまとめる
- **SPLIT** — CSVを分ける
- **TRANSFORM** — CSVを整える

すべての処理を **CSVを選ぶ → 条件を決める → 確認して保存** の3ステップで進めます。

## UI / UX

初見でも操作を選びやすいよう、v2の公開版では画面構成を簡素化しています。

### Home

起動直後に表示するのは次の3択だけです。

```text
CSVで何をしますか？

[ CSVをまとめる ]
[ CSVを分ける   ]
[ CSVを整える   ]
```

### Progressive disclosure

必要になるまで詳細を見せません。

- JOINの `LEFT JOIN / INNER JOIN` は、画面上では
  - `左のCSVをすべて残す`
  - `両方にある行だけ残す`
  と表示
- JOIN診断は、まず結果を1文で表示し、必要な場合だけ `詳細を見る`
- TRANSFORMの設定フォームは `＋ 変換を追加` を押したときだけ表示
- JOIN / SPLITは確認処理後、自動的に保存ステップへ移動

専門用語は補足として残しつつ、操作判断には不要な設計にしています。

## Features

### JOIN — CSVをまとめる

2種類のまとめ方があります。

**共通する列で情報をつなぐ**

- 左右で異なるキー列を指定可能
- LEFT JOIN / INNER JOIN
- JOIN判定時はキー前後の空白を無視（元データは保持）
- 空欄キー同士は一致させない
- 一致 / 左のみ / 右のみ / 重複キーを事前診断
- 多対多JOINを警告
- 条件変更後は古いPreviewを無効化

**行を下につなげる**

- 2ファイル以上を一括選択
- 列名を基準に列をそろえて縦結合
- 不足列は空欄で補完

### SPLIT — CSVを分ける

- 指定した1列の値ごとに自動分割
- 出力ファイル名と件数を保存前に確認
- 空欄キーを除外、または `blank.csv` へ出力
- Windowsで使えないファイル名文字を自動置換
- `CON` などのWindows予約名を回避
- 同名化したファイルは `_2`, `_3` のように衝突回避

### TRANSFORM — CSVを整える

変換を上から順番に追加できます。

- 列名変更
- 列削除
- 列順変更
- 固定値列追加
- 文字列置換
- 前後空白削除
- 日付形式変換
- 0埋め
- 列名変更後の新しい列を次の変換で利用可能
- 変更前 / 変更後を切り替えてPreview

### Common

- UTF-8 BOM → UTF-8 → CP932 の順で自動読込
- UTF-8 BOM / UTF-8 / CP932 を選んで出力
- 引用符で囲まれたセル内改行を保持
- Previewではセル内改行を `↵` 表示
- 出力CSVの改行コードはCRLF
- CSVはPC内だけで処理し、外部へアップロードしない

## Sample data

`sample/` には架空の生徒データを収録しています。実在する人物・学校とは関係ありません。

```text
sample/
├─ students.csv
├─ scores.csv
├─ clubs.csv
└─ validation/
```

試し方:

- **JOIN**: `students.csv` の `生徒ID` と `scores.csv` の `受験者ID`
- **SPLIT**: `students.csv` を `クラス` で分割
- **TRANSFORM**: `students.csv` の列名・入学日・列順などを変更

`validation/` には前後空白、空欄、重複、多対多、セル内改行、CP932、Windows禁止文字などの境界ケースを含めています。

## Validation

```bash
python -m pytest -q
python tools/self_check.py
```

現在の回帰テストは **18 tests**。

自己検証対象:

- 空欄キーをJOINしない
- キー前後空白の吸収
- 1対多 / 多対多検知
- SPLITファイル名sanitize / 衝突回避
- 列名変更後のTRANSFORM継続
- セル内改行の保持
- CP932読込
- CRLF出力と再読込
- 列構成が異なるCSVの縦結合
- 100,000行の簡易負荷確認

GUIの主要操作は `tools/ui_smoke.py` でも確認できます。

## Setup

```bash
python -m pip install -r requirements.txt
python CSVjoiner.py
```

Python 3.11+ 推奨。

## Build Windows EXE

Windows上でリポジトリ直下の `build_windows.bat` を実行します。

```bat
build_windows.bat
```

ビルド前に自動テストを実行し、PASSした場合だけPyInstallerで1ファイルEXEを生成します。

```text
dist/CSVjoiner.exe
```

`dist/` はGit管理対象外です。公開時はEXEをGitHub Releasesのassetとして添付する想定です。

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
│  └─ validation/
├─ tests/
├─ tools/
│  ├─ self_check.py
│  └─ ui_smoke.py
├─ CSVjoiner.spec
├─ version_info.txt
├─ requirements.txt
├─ requirements-dev.txt
├─ build_windows.bat
└─ README.md
```

## Design

```text
やりたいことを選ぶ
        ↓
1. CSVを選ぶ
        ↓
2. 条件を決める
        ↓
3. 確認して保存
```

旧版の「CSVを結合するコード」から、v2では **CSV加工を迷わず安全に行う小さなデスクトップツール** へ刷新しました。

## License

MIT
