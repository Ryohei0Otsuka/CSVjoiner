from pathlib import Path

import pandas as pd

from csvjoiner.core import export_csv, format_preview_value, read_csv_flexible, sanitize_filename
from csvjoiner.models import TransformStep
from csvjoiner.operations import (
    analyze_key_join,
    apply_transform_pipeline,
    key_join,
    split_dataframe,
    vertical_join,
)


SAMPLE_DIR = Path(__file__).resolve().parents[1] / "sample"


def test_vertical_join_aligns_union_of_columns():
    a = pd.DataFrame({"生徒ID": ["S001"], "氏名": ["青木葵"]})
    b = pd.DataFrame({"生徒ID": ["S002"], "点数": ["91"]})
    result = vertical_join([a, b])
    assert list(result.columns) == ["生徒ID", "氏名", "点数"]
    assert len(result) == 2
    assert result.iloc[1]["氏名"] == ""


def test_key_join_and_diagnostics():
    left = pd.DataFrame(
        {"生徒ID": ["S001", "S002", "S002", "S003"], "氏名": ["A", "B", "B2", "C"]}
    )
    right = pd.DataFrame(
        {"受験者ID": ["S002", "S002", "S004"], "科目": ["国語", "数学", "国語"]}
    )
    diag = analyze_key_join(left, right, "生徒ID", "受験者ID")
    assert diag.matched_keys == 1
    assert diag.left_only_keys == 2
    assert diag.right_only_keys == 1
    assert diag.left_duplicate_keys == 1
    assert diag.right_duplicate_keys == 1
    assert diag.many_to_many_keys == 1

    result, diag2 = key_join(left, right, "生徒ID", "受験者ID", "left")
    assert len(result) == 6
    assert diag2.result_rows == 6


def test_split_dataframe_blank_policy():
    df = pd.DataFrame(
        {"クラス": ["1-A", "1-B", "1-A", ""], "氏名": ["A", "B", "C", "D"]}
    )
    groups = split_dataframe(df, "クラス", include_blank=False)
    assert set(groups) == {"1-A", "1-B"}
    assert len(groups["1-A"]) == 2
    groups2 = split_dataframe(df, "クラス", include_blank=True)
    assert "blank" in groups2


def test_transform_pipeline():
    df = pd.DataFrame(
        {"番号": ["123", " 7 "], "入学日": ["2026/8/23", "bad"], "氏名": [" 青木葵 ", "高橋蓮"]}
    )
    steps = [
        TransformStep("trim", "氏名"),
        TransformStep("zero_pad", "番号", "6"),
        TransformStep("date", "入学日", "%Y-%m-%d"),
        TransformStep("rename", "氏名", "生徒名"),
        TransformStep("add_fixed", "", "在籍状況", "在籍"),
        TransformStep("move", "在籍状況", "先頭"),
    ]
    result = apply_transform_pipeline(df, steps)
    assert result.columns[0] == "在籍状況"
    assert result.loc[0, "番号"] == "000123"
    assert result.loc[0, "入学日"] == "2026-08-23"
    assert result.loc[1, "入学日"] == "bad"
    assert result.loc[0, "生徒名"] == "青木葵"


def test_sanitize_filename_windows_rules():
    assert sanitize_filename('1-A/進学:*?') == '1-A_進学___'
    assert sanitize_filename('CON') == '_CON'


def test_sample_students_split_by_class():
    students = read_csv_flexible(SAMPLE_DIR / "students.csv").dataframe
    groups = split_dataframe(students, "クラス", include_blank=False)
    assert set(groups) == {"1-A", "1-B", "1-C"}
    assert len(groups["1-A"]) == 2
    assert len(groups["1-B"]) == 2


def test_sample_students_scores_join_uses_different_key_names():
    students = read_csv_flexible(SAMPLE_DIR / "students.csv").dataframe
    scores = read_csv_flexible(SAMPLE_DIR / "scores.csv").dataframe
    result, diag = key_join(students, scores, "生徒ID", "受験者ID", "left")
    assert diag.matched_keys == 4
    assert diag.left_only_keys == 2
    assert diag.right_only_keys == 1
    assert diag.right_duplicate_keys == 1
    assert diag.many_to_many_keys == 0
    assert len(result) == 7


def test_sample_students_clubs_join_has_left_and_right_only_keys():
    students = read_csv_flexible(SAMPLE_DIR / "students.csv").dataframe
    clubs = read_csv_flexible(SAMPLE_DIR / "clubs.csv").dataframe
    _, diag = key_join(students, clubs, "生徒ID", "会員ID", "inner")
    assert diag.matched_keys == 4
    assert diag.left_only_keys == 2
    assert diag.right_only_keys == 1


def test_embedded_newline_roundtrip_and_crlf_output(tmp_path):
    source = tmp_path / "multiline.csv"
    source.write_bytes(
        "生徒ID,メモ\r\nS001,\"1行目\r\n2行目\"\r\nS002,通常\r\n".encode("utf-8-sig")
    )
    doc = read_csv_flexible(source)
    assert len(doc.dataframe) == 2
    assert "1行目" in doc.dataframe.loc[0, "メモ"]
    assert "2行目" in doc.dataframe.loc[0, "メモ"]

    output = tmp_path / "out.csv"
    export_csv(doc.dataframe, output, "UTF-8 BOM")
    raw = output.read_bytes()
    assert b"\r\n" in raw

    reread = read_csv_flexible(output)
    assert reread.dataframe.equals(doc.dataframe)


def test_preview_marks_embedded_newlines_without_changing_data():
    value = "1行目\r\n2行目\n3行目"
    assert format_preview_value(value) == "1行目 ↵ 2行目 ↵ 3行目"
