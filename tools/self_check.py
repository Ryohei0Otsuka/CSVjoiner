from __future__ import annotations

from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from time import perf_counter

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from csvjoiner.core import export_csv, read_csv_flexible
from csvjoiner.models import TransformStep
from csvjoiner.operations import (
    apply_transform_pipeline,
    export_split_groups,
    key_join,
    split_dataframe,
    split_plan,
    vertical_join,
)

V = ROOT / "sample" / "validation"


def check(name: str, condition: bool) -> None:
    if not condition:
        raise AssertionError(name)
    print(f"[PASS] {name}")


def main() -> None:
    left = read_csv_flexible(V / "join_students.csv")
    right = read_csv_flexible(V / "join_scores.csv")
    joined, diag = key_join(left.dataframe, right.dataframe, "生徒ID", "受験者ID", "left")
    check("JOIN diagnostics", (diag.matched_keys, diag.left_only_keys, diag.right_only_keys) == (3, 1, 1))
    check("JOIN many-to-many detection", diag.many_to_many_keys == 1)
    check("JOIN trims key whitespace", joined.loc[joined["生徒ID"] == " S002 ", "科目"].iloc[0] == "数学")
    check("JOIN does not match blank keys", joined.loc[joined["生徒ID"] == "", "科目"].iloc[0] == "")

    split_doc = read_csv_flexible(V / "split_students.csv")
    plan = split_plan(split_doc.dataframe, "クラス", include_blank=True)
    check("SPLIT filename sanitization", "_CON.csv" in set(plan["出力ファイル"]))
    check("SPLIT collision avoidance", "1_B_2.csv" in set(plan["出力ファイル"]))

    transform_doc = read_csv_flexible(V / "transform_students.csv")
    transformed = apply_transform_pipeline(
        transform_doc.dataframe,
        [
            TransformStep("rename", "氏名", "生徒名"),
            TransformStep("trim", "生徒名"),
            TransformStep("zero_pad", "番号", "6"),
            TransformStep("date", "入学日", "%Y-%m-%d"),
            TransformStep("add_fixed", "", "状態", "在籍"),
            TransformStep("move", "状態", "先頭"),
        ],
    )
    check("TRANSFORM renamed column can be reused", transformed.loc[0, "生徒名"] == "青木 葵")
    check("TRANSFORM multiline preservation", "\n" in transformed.loc[0, "メモ"] or "\r" in transformed.loc[0, "メモ"])
    check("TRANSFORM invalid date preserved", transformed.loc[2, "入学日"] == "not-a-date")

    stacked = vertical_join([
        read_csv_flexible(V / "stack_a.csv").dataframe,
        read_csv_flexible(V / "stack_b.csv").dataframe,
    ])
    check("STACK aligns different schemas", list(stacked.columns) == ["生徒ID", "氏名", "クラス", "部活動"])

    cp932 = read_csv_flexible(V / "cp932_students.csv")
    check("CP932 detection", cp932.encoding == "cp932")

    with TemporaryDirectory() as td:
        out_dir = Path(td)
        groups = split_dataframe(split_doc.dataframe, "クラス", include_blank=True)
        paths = export_split_groups(groups, out_dir / "split", "UTF-8 BOM")
        check("SPLIT export count", len(paths) == 5)
        roundtrip = out_dir / "roundtrip.csv"
        export_csv(transformed, roundtrip, "UTF-8 BOM")
        reread = read_csv_flexible(roundtrip).dataframe
        check("CRLF/export roundtrip", reread.equals(transformed))

    # Moderate synthetic load to catch accidental O(n^2) behavior in core operations.
    n = 100_000
    master = pd.DataFrame({
        "id": [f"S{i:06d}" for i in range(n)],
        "class": [f"C{i % 20:02d}" for i in range(n)],
        "value": [str(i) for i in range(n)],
    })
    detail = pd.DataFrame({
        "student_id": [f"S{i:06d}" for i in range(0, n, 2)],
        "score": [str(i % 100) for i in range(0, n, 2)],
    })
    t0 = perf_counter()
    large_join, _ = key_join(master, detail, "id", "student_id", "left")
    t1 = perf_counter()
    large_groups = split_dataframe(master, "class")
    t2 = perf_counter()
    large_transform = apply_transform_pipeline(master, [TransformStep("zero_pad", "value", "8")])
    t3 = perf_counter()
    check("100k JOIN row count", len(large_join) == n)
    check("100k SPLIT group count", len(large_groups) == 20)
    check("100k TRANSFORM row count", len(large_transform) == n)
    print(f"[PERF] 100k JOIN {t1 - t0:.3f}s / SPLIT {t2 - t1:.3f}s / TRANSFORM {t3 - t2:.3f}s")
    print("SELF CHECK COMPLETE")


if __name__ == "__main__":
    main()
