from pathlib import Path

import pandas as pd

from csvjoiner.core import sanitize_filename
from csvjoiner.models import TransformStep
from csvjoiner.operations import (
    analyze_key_join,
    apply_transform_pipeline,
    key_join,
    split_dataframe,
    vertical_join,
)


def test_vertical_join_aligns_union_of_columns():
    a = pd.DataFrame({"id": ["1"], "name": ["A"]})
    b = pd.DataFrame({"id": ["2"], "amount": ["100"]})
    result = vertical_join([a, b])
    assert list(result.columns) == ["id", "name", "amount"]
    assert len(result) == 2
    assert result.iloc[1]["name"] == ""


def test_key_join_and_diagnostics():
    left = pd.DataFrame({"id": ["1", "2", "2", "3"], "name": ["A", "B", "B2", "C"]})
    right = pd.DataFrame({"staff_id": ["2", "2", "4"], "dept": ["X", "Y", "Z"]})
    diag = analyze_key_join(left, right, "id", "staff_id")
    assert diag.matched_keys == 1
    assert diag.left_only_keys == 2
    assert diag.right_only_keys == 1
    assert diag.left_duplicate_keys == 1
    assert diag.right_duplicate_keys == 1
    assert diag.many_to_many_keys == 1

    result, diag2 = key_join(left, right, "id", "staff_id", "left")
    assert len(result) == 6
    assert diag2.result_rows == 6


def test_split_dataframe_blank_policy():
    df = pd.DataFrame({"dept": ["Sales", "Dev", "Sales", ""], "name": ["A", "B", "C", "D"]})
    groups = split_dataframe(df, "dept", include_blank=False)
    assert set(groups) == {"Sales", "Dev"}
    assert len(groups["Sales"]) == 2
    groups2 = split_dataframe(df, "dept", include_blank=True)
    assert "blank" in groups2


def test_transform_pipeline():
    df = pd.DataFrame({"code": ["123", " 7 "], "date": ["2026/8/23", "bad"], "name": [" A ", "B"]})
    steps = [
        TransformStep("trim", "name"),
        TransformStep("zero_pad", "code", "6"),
        TransformStep("date", "date", "%Y-%m-%d"),
        TransformStep("rename", "name", "person"),
        TransformStep("add_fixed", "", "status", "ok"),
        TransformStep("move", "status", "先頭"),
    ]
    result = apply_transform_pipeline(df, steps)
    assert result.columns[0] == "status"
    assert result.loc[0, "code"] == "000123"
    assert result.loc[0, "date"] == "2026-08-23"
    assert result.loc[1, "date"] == "bad"
    assert result.loc[0, "person"] == "A"


def test_sanitize_filename_windows_rules():
    assert sanitize_filename('営業/東京:*?') == '営業_東京___'
    assert sanitize_filename('CON') == '_CON'
