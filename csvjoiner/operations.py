from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd

from .core import CsvToolError, export_csv, unique_output_path
from .models import JoinDiagnostics, TransformStep


def _require_column(df: pd.DataFrame, column: str, label: str = "CSV") -> None:
    if not column or column not in df.columns:
        raise CsvToolError(f"{label} に列「{column}」がありません。")


def vertical_join(dataframes: Iterable[pd.DataFrame]) -> pd.DataFrame:
    frames = [df.copy() for df in dataframes]
    if len(frames) < 2:
        raise CsvToolError("縦結合には2つ以上のCSVが必要です。")

    column_order: list[str] = []
    for frame in frames:
        for column in frame.columns:
            name = str(column)
            if name not in column_order:
                column_order.append(name)

    aligned = [frame.reindex(columns=column_order, fill_value="") for frame in frames]
    return pd.concat(aligned, ignore_index=True)


def analyze_key_join(
    left: pd.DataFrame,
    right: pd.DataFrame,
    left_key: str,
    right_key: str,
) -> JoinDiagnostics:
    _require_column(left, left_key, "左CSV")
    _require_column(right, right_key, "右CSV")

    left_values = left[left_key].astype(str).str.strip()
    right_values = right[right_key].astype(str).str.strip()

    left_nonblank = left_values[left_values != ""]
    right_nonblank = right_values[right_values != ""]

    left_counts = left_nonblank.value_counts()
    right_counts = right_nonblank.value_counts()
    left_keys = set(left_counts.index)
    right_keys = set(right_counts.index)
    matched = left_keys & right_keys

    left_dups = {k for k, v in left_counts.items() if v > 1}
    right_dups = {k for k, v in right_counts.items() if v > 1}
    many_to_many = left_dups & right_dups

    diagnostics = JoinDiagnostics(
        matched_keys=len(matched),
        left_only_keys=len(left_keys - right_keys),
        right_only_keys=len(right_keys - left_keys),
        left_duplicate_keys=len(left_dups),
        right_duplicate_keys=len(right_dups),
        many_to_many_keys=len(many_to_many),
    )

    if (left_values == "").any():
        diagnostics.warnings.append(f"左CSVに空欄キーが {(left_values == '').sum()} 行あります。")
    if (right_values == "").any():
        diagnostics.warnings.append(f"右CSVに空欄キーが {(right_values == '').sum()} 行あります。")
    if many_to_many:
        diagnostics.warnings.append(
            f"多対多になるキーが {len(many_to_many)} 件あります。結果行数が増える可能性があります。"
        )
    elif left_dups or right_dups:
        diagnostics.warnings.append("重複キーがあります。1対多の結合結果になる可能性があります。")
    return diagnostics


def key_join(
    left: pd.DataFrame,
    right: pd.DataFrame,
    left_key: str,
    right_key: str,
    how: str = "left",
) -> tuple[pd.DataFrame, JoinDiagnostics]:
    if how not in {"left", "inner"}:
        raise CsvToolError("JOIN TYPE は LEFT または INNER を指定してください。")

    diagnostics = analyze_key_join(left, right, left_key, right_key)
    result = pd.merge(
        left,
        right,
        how=how,
        left_on=left_key,
        right_on=right_key,
        suffixes=("_left", "_right"),
        sort=False,
    )
    diagnostics.result_rows = len(result)
    return result.fillna(""), diagnostics


def split_counts(df: pd.DataFrame, key_column: str, blank_label: str = "(空欄)") -> pd.DataFrame:
    _require_column(df, key_column)
    keys = df[key_column].astype(str).str.strip().replace("", blank_label)
    counts = keys.value_counts(dropna=False).rename_axis(key_column).reset_index(name="件数")
    return counts


def split_dataframe(
    df: pd.DataFrame,
    key_column: str,
    include_blank: bool = False,
) -> dict[str, pd.DataFrame]:
    _require_column(df, key_column)
    key_series = df[key_column].astype(str).str.strip()
    groups: dict[str, pd.DataFrame] = {}

    for raw_key in key_series.drop_duplicates().tolist():
        if raw_key == "" and not include_blank:
            continue
        label = raw_key if raw_key else "blank"
        groups[label] = df[key_series == raw_key].copy()
    return groups


def export_split_groups(
    groups: dict[str, pd.DataFrame],
    output_dir: Path,
    encoding_label: str = "UTF-8 BOM",
) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    used: set[str] = set()
    paths: list[Path] = []
    for label, frame in groups.items():
        path = unique_output_path(output_dir, label, used)
        export_csv(frame, path, encoding_label)
        paths.append(path)
    return paths


def apply_transform_step(df: pd.DataFrame, step: TransformStep) -> pd.DataFrame:
    result = df.copy()
    op = step.operation

    if op == "rename":
        _require_column(result, step.column)
        new_name = step.value1.strip()
        if not new_name:
            raise CsvToolError("新しい列名を入力してください。")
        if new_name != step.column and new_name in result.columns:
            raise CsvToolError(f"列「{new_name}」は既に存在します。")
        return result.rename(columns={step.column: new_name})

    if op == "delete":
        _require_column(result, step.column)
        return result.drop(columns=[step.column])

    if op == "add_fixed":
        new_name = step.value1.strip()
        if not new_name:
            raise CsvToolError("追加する列名を入力してください。")
        if new_name in result.columns:
            raise CsvToolError(f"列「{new_name}」は既に存在します。")
        result[new_name] = step.value2
        return result

    if op == "trim":
        _require_column(result, step.column)
        result[step.column] = result[step.column].astype(str).str.strip()
        return result

    if op == "replace":
        _require_column(result, step.column)
        if step.value1 == "":
            raise CsvToolError("置換前の文字列を入力してください。")
        result[step.column] = result[step.column].astype(str).str.replace(
            step.value1, step.value2, regex=False
        )
        return result

    if op == "date":
        _require_column(result, step.column)
        fmt = step.value1.strip() or "%Y-%m-%d"
        original = result[step.column].astype(str)
        parsed = pd.to_datetime(original, errors="coerce")
        converted = parsed.dt.strftime(fmt)
        result[step.column] = converted.where(parsed.notna(), original)
        return result

    if op == "zero_pad":
        _require_column(result, step.column)
        try:
            width = int(step.value1)
        except ValueError as exc:
            raise CsvToolError("0埋めの桁数は整数で入力してください。") from exc
        if width <= 0 or width > 100:
            raise CsvToolError("0埋めの桁数は1〜100で指定してください。")
        result[step.column] = result[step.column].astype(str).str.zfill(width)
        return result

    if op == "move":
        _require_column(result, step.column)
        target = step.value1.strip()
        columns = list(result.columns)
        columns.remove(step.column)
        if target == "先頭":
            columns.insert(0, step.column)
        elif target == "末尾":
            columns.append(step.column)
        else:
            try:
                pos = int(target)
            except ValueError as exc:
                raise CsvToolError("移動先は「先頭」「末尾」または1始まりの列番号で指定してください。") from exc
            pos = max(1, min(pos, len(columns) + 1))
            columns.insert(pos - 1, step.column)
        return result[columns]

    raise CsvToolError(f"未対応の変換処理です: {op}")


def apply_transform_pipeline(df: pd.DataFrame, steps: list[TransformStep]) -> pd.DataFrame:
    result = df.copy()
    for step in steps:
        result = apply_transform_step(result, step)
    return result
