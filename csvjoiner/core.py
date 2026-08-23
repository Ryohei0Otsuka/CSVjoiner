from __future__ import annotations

import csv
import re
from pathlib import Path

import pandas as pd

from .models import CsvDocument


ENCODING_CANDIDATES = ("utf-8-sig", "cp932", "utf-8")
EXPORT_ENCODINGS = {
    "UTF-8 BOM": "utf-8-sig",
    "UTF-8": "utf-8",
    "CP932": "cp932",
}
WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


class CsvToolError(Exception):
    """Expected user-facing error."""


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result.columns = [str(col).strip() for col in result.columns]
    if len(set(result.columns)) != len(result.columns):
        raise CsvToolError("列名を整形した結果、同名列が重複しました。列名を確認してください。")
    return result


def read_csv_flexible(path: Path) -> CsvDocument:
    if not path.exists():
        raise CsvToolError(f"ファイルが見つかりません: {path}")
    if path.suffix.lower() != ".csv":
        raise CsvToolError(f"CSVファイルを選択してください: {path.name}")

    last_error: Exception | None = None
    for encoding in ENCODING_CANDIDATES:
        try:
            df = pd.read_csv(
                path,
                encoding=encoding,
                dtype=str,
                keep_default_na=False,
                na_filter=False,
            )
            return CsvDocument(path=path, dataframe=normalize_columns(df), encoding=encoding)
        except UnicodeDecodeError as exc:
            last_error = exc
        except pd.errors.EmptyDataError as exc:
            raise CsvToolError(f"CSVが空です: {path.name}") from exc
        except Exception as exc:
            last_error = exc

    raise CsvToolError(f"CSVを読み込めませんでした: {path.name}\n{last_error}")


def export_csv(df: pd.DataFrame, path: Path, encoding_label: str = "UTF-8 BOM") -> None:
    encoding = EXPORT_ENCODINGS.get(encoding_label)
    if not encoding:
        raise CsvToolError(f"未対応の文字コードです: {encoding_label}")
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_csv(path, index=False, encoding=encoding, quoting=csv.QUOTE_MINIMAL)
    except UnicodeEncodeError as exc:
        raise CsvToolError(
            f"{encoding_label} では表現できない文字が含まれています。UTF-8 BOM での出力を試してください。"
        ) from exc


def sanitize_filename(value: object, fallback: str = "blank") -> str:
    text = str(value).strip()
    if not text:
        text = fallback
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    if not text:
        text = fallback
    if text.upper() in WINDOWS_RESERVED_NAMES:
        text = f"_{text}"
    return text[:120]


def unique_output_path(directory: Path, stem: str, used_names: set[str]) -> Path:
    base = sanitize_filename(stem)
    candidate = base
    index = 2
    while candidate.lower() in used_names or (directory / f"{candidate}.csv").exists():
        candidate = f"{base}_{index}"
        index += 1
    used_names.add(candidate.lower())
    return directory / f"{candidate}.csv"


def dataframe_preview(df: pd.DataFrame, limit: int = 100) -> pd.DataFrame:
    return df.head(limit).copy()
