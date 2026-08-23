from pathlib import Path

import pandas as pd

from csvjoiner.core import export_csv, read_csv_flexible
from csvjoiner.models import TransformStep
from csvjoiner.operations import (
    apply_transform_pipeline,
    key_join,
    split_dataframe,
    split_plan,
    vertical_join,
)

VALIDATION_DIR = Path(__file__).resolve().parents[1] / "sample" / "validation"


def test_join_validation_trims_keys_and_never_matches_blank_keys():
    left = read_csv_flexible(VALIDATION_DIR / "join_students.csv").dataframe
    right = read_csv_flexible(VALIDATION_DIR / "join_scores.csv").dataframe

    result, diag = key_join(left, right, "生徒ID", "受験者ID", "left")

    assert diag.matched_keys == 3
    assert diag.left_only_keys == 1
    assert diag.right_only_keys == 1
    assert diag.left_duplicate_keys == 1
    assert diag.right_duplicate_keys == 1
    assert diag.many_to_many_keys == 1
    assert len(result) == 8

    # 前後空白付きのS002は一致するが、元データ値は改変しない。
    s002 = result[result["生徒ID"] == " S002 "]
    assert len(s002) == 1
    assert s002.iloc[0]["科目"] == "数学"

    # 空欄キー同士は一致させない。
    blank = result[result["生徒ID"] == ""]
    assert len(blank) == 1
    assert blank.iloc[0]["科目"] == ""
    assert blank.iloc[0]["備考"] == ""


def test_inner_join_excludes_blank_keys():
    left = read_csv_flexible(VALIDATION_DIR / "join_students.csv").dataframe
    right = read_csv_flexible(VALIDATION_DIR / "join_scores.csv").dataframe
    result, _ = key_join(left, right, "生徒ID", "受験者ID", "inner")
    assert len(result) == 6
    assert (result["生徒ID"].astype(str).str.strip() != "").all()


def test_split_plan_shows_sanitized_and_collision_safe_names():
    df = read_csv_flexible(VALIDATION_DIR / "split_students.csv").dataframe
    plan = split_plan(df, "クラス", include_blank=True)

    assert plan["件数"].sum() == len(df)
    assert set(plan["出力ファイル"]) == {
        "1-A.csv",
        "1_B.csv",
        "1_B_2.csv",
        "_CON.csv",
        "blank.csv",
    }


def test_split_validation_groups_trim_key_whitespace():
    df = read_csv_flexible(VALIDATION_DIR / "split_students.csv").dataframe
    groups = split_dataframe(df, "クラス", include_blank=False)
    assert len(groups["1-A"]) == 2


def test_transform_validation_keeps_multiline_and_reuses_renamed_column():
    df = read_csv_flexible(VALIDATION_DIR / "transform_students.csv").dataframe
    steps = [
        TransformStep("rename", "氏名", "生徒名"),
        TransformStep("trim", "生徒名"),
        TransformStep("zero_pad", "番号", "6"),
        TransformStep("date", "入学日", "%Y-%m-%d"),
        TransformStep("add_fixed", "", "状態", "在籍"),
        TransformStep("move", "状態", "先頭"),
    ]
    result = apply_transform_pipeline(df, steps)

    assert result.columns[0] == "状態"
    assert result.loc[0, "生徒名"] == "青木 葵"
    assert result.loc[0, "番号"] == "000007"
    assert result.loc[0, "入学日"] == "2026-04-01"
    assert result.loc[2, "入学日"] == "not-a-date"
    assert "1行目" in result.loc[0, "メモ"] and "2行目" in result.loc[0, "メモ"]


def test_stack_validation_aligns_different_schemas():
    a = read_csv_flexible(VALIDATION_DIR / "stack_a.csv").dataframe
    b = read_csv_flexible(VALIDATION_DIR / "stack_b.csv").dataframe
    result = vertical_join([a, b])
    assert list(result.columns) == ["生徒ID", "氏名", "クラス", "部活動"]
    assert len(result) == 4
    assert result.loc[0, "部活動"] == ""
    assert result.loc[2, "クラス"] == ""


def test_cp932_validation_sample_is_readable():
    doc = read_csv_flexible(VALIDATION_DIR / "cp932_students.csv")
    assert doc.encoding == "cp932"
    assert doc.dataframe.loc[0, "氏名"] == "鈴木 海"


def test_export_validation_uses_crlf_and_roundtrips(tmp_path):
    df = read_csv_flexible(VALIDATION_DIR / "transform_students.csv").dataframe
    out = tmp_path / "roundtrip.csv"
    export_csv(df, out, "UTF-8 BOM")
    assert b"\r\n" in out.read_bytes()
    reread = read_csv_flexible(out).dataframe
    pd.testing.assert_frame_equal(reread, df)
