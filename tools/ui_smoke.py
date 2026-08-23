from __future__ import annotations

from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from csvjoiner.core import read_csv_flexible
from csvjoiner.ui import CSVJoinerApp

V = ROOT / "sample" / "validation"


def main() -> None:
    app = CSVJoinerApp()
    app.withdraw()
    app.update_idletasks()

    # JOIN populated state
    app.join_left = read_csv_flexible(V / "join_students.csv")
    app.join_right = read_csv_flexible(V / "join_scores.csv")
    app.join_left_combo["values"] = app.join_left.columns
    app.join_right_combo["values"] = app.join_right.columns
    app.join_left_key.set("生徒ID")
    app.join_right_key.set("受験者ID")
    app._run_join_preview()
    assert app.join_result is not None and len(app.join_result) == 8
    assert app.join_metric_match.get() == "3"

    # Changing a condition must invalidate the previous result.
    app._set_join_type("inner")
    assert app.join_result is None
    app._set_join_type("left")
    app._run_join_preview()

    # SPLIT populated state
    app.split_doc = read_csv_flexible(V / "split_students.csv")
    app.split_key_combo["values"] = app.split_doc.columns
    app.split_key.set("クラス")
    app.split_include_blank.set(True)
    app.split_row_count.set(str(app.split_doc.rows))
    app._preview_split_counts()
    assert app.split_analysis_valid
    assert app.split_file_count.get() == "5"
    app.split_include_blank.set(False)
    app._invalidate_split_analysis()
    assert not app.split_analysis_valid

    # TRANSFORM: rename first, then reuse the new column in the next step.
    app.transform_doc = read_csv_flexible(V / "transform_students.csv")
    app.transform_col_combo["values"] = app.transform_doc.columns
    app.transform_column.set("氏名")
    app.transform_before.show_dataframe(app.transform_doc.dataframe)

    app.transform_operation.set("rename")
    app.transform_value1.set("生徒名")
    app._add_transform_step()
    assert "生徒名" in app.transform_col_combo["values"]

    app.transform_operation.set("trim")
    app.transform_column.set("生徒名")
    app._add_transform_step()
    assert app.transform_result is not None
    assert app.transform_result.loc[0, "生徒名"] == "青木 葵"

    app.destroy()
    print("UI SMOKE COMPLETE")


if __name__ == "__main__":
    main()
