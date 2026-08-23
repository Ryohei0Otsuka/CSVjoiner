from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


@dataclass
class CsvDocument:
    path: Path
    dataframe: pd.DataFrame
    encoding: str

    @property
    def rows(self) -> int:
        return len(self.dataframe)

    @property
    def columns(self) -> list[str]:
        return [str(c) for c in self.dataframe.columns]


@dataclass
class JoinDiagnostics:
    matched_keys: int = 0
    left_only_keys: int = 0
    right_only_keys: int = 0
    left_duplicate_keys: int = 0
    right_duplicate_keys: int = 0
    many_to_many_keys: int = 0
    result_rows: int = 0
    warnings: list[str] = field(default_factory=list)


@dataclass
class TransformStep:
    operation: str
    column: str = ""
    value1: str = ""
    value2: str = ""

    def describe(self) -> str:
        labels = {
            "rename": "列名変更",
            "delete": "列削除",
            "add_fixed": "固定値列追加",
            "trim": "前後空白削除",
            "replace": "文字列置換",
            "date": "日付形式変換",
            "zero_pad": "0埋め",
            "move": "列順変更",
        }
        label = labels.get(self.operation, self.operation)
        if self.operation == "rename":
            return f"{label}: {self.column} → {self.value1}"
        if self.operation == "delete":
            return f"{label}: {self.column}"
        if self.operation == "add_fixed":
            return f"{label}: {self.value1} = {self.value2}"
        if self.operation == "trim":
            return f"{label}: {self.column}"
        if self.operation == "replace":
            return f"{label}: {self.column} / {self.value1} → {self.value2}"
        if self.operation == "date":
            return f"{label}: {self.column} → {self.value1 or '%Y-%m-%d'}"
        if self.operation == "zero_pad":
            return f"{label}: {self.column} / {self.value1}桁"
        if self.operation == "move":
            return f"{label}: {self.column} → {self.value1}"
        return label
