from __future__ import annotations

import csv
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "CSV Joiner Neon"
APP_VERSION = "2.1.0"

DEFAULT_ENCODING_CANDIDATES = ["utf-8-sig", "cp932", "utf-8"]
DEFAULT_B_DATE_CANDIDATES = ["日付", "利用日", "年月日"]
DEFAULT_B_AMOUNT_CANDIDATES = ["交通費", "金額", "費用", "支給額"]

COLORS = {
    "bg": "#0b1020",
    "panel": "#12182b",
    "panel2": "#161f38",
    "text": "#e8ecff",
    "muted": "#99a3c7",
    "pink": "#ff4ecd",
    "cyan": "#38e1ff",
    "violet": "#8b5cf6",
    "success": "#37f0a8",
    "warn": "#ff7a90",
    "entry_bg": "#0f1530",
}


class CsvMergeError(Exception):
    pass


@dataclass
class MergeConfig:
    file_a: Path
    file_b: Path
    file_c: Path
    output_file: Path
    key_col: str
    name_col: str
    b_date_col: str
    b_amount_col: str


# =========================
# Core merge logic
# =========================

def read_csv_flexible(path: Path) -> pd.DataFrame:
    last_error: Exception | None = None
    for encoding in DEFAULT_ENCODING_CANDIDATES:
        try:
            return pd.read_csv(path, encoding=encoding, dtype=str).fillna("")
        except Exception as exc:
            last_error = exc
    raise CsvMergeError(f"CSVを読み込めませんでした: {path.name}\n{last_error}")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    return df


def require_columns(df: pd.DataFrame, required_cols: Iterable[str], file_label: str) -> None:
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise CsvMergeError(
            f"{file_label} に必要な列がありません: {', '.join(missing)}\n"
            f"存在する列: {', '.join(map(str, df.columns))}"
        )


def clean_common_columns(df: pd.DataFrame, key_col: str, name_col: str) -> pd.DataFrame:
    df = df.copy()
    df[key_col] = df[key_col].astype(str).str.strip()
    if name_col in df.columns:
        df[name_col] = df[name_col].astype(str).str.strip()
    else:
        df[name_col] = ""
    df = df[df[key_col] != ""].copy()
    return df


def validate_unique_key(df: pd.DataFrame, key_col: str, file_label: str) -> None:
    dup = df[df.duplicated(subset=[key_col], keep=False)]
    if not dup.empty:
        preview = dup[[key_col]].drop_duplicates().head(10)[key_col].tolist()
        raise CsvMergeError(
            f"{file_label} に {key_col} の重複があります。\n"
            f"重複キー例: {', '.join(map(str, preview))}"
        )


def choose_best_name(row: pd.Series, preferred_cols: list[str]) -> str:
    for col in preferred_cols:
        value = str(row.get(col, "")).strip()
        if value and value.lower() != "nan":
            return value
    return ""


def find_first_existing(columns: list[str], candidates: list[str]) -> str:
    for candidate in candidates:
        if candidate in columns:
            return candidate
    raise CsvMergeError(
        f"候補列が見つかりません。候補: {', '.join(candidates)} / 実列: {', '.join(columns)}"
    )


def aggregate_b(
    df_b: pd.DataFrame,
    key_col: str,
    name_col: str,
    date_col: str,
    amount_col: str,
) -> pd.DataFrame:
    df = df_b.copy()
    require_columns(df, [key_col, date_col, amount_col], "B.csv")

    if name_col not in df.columns:
        df[name_col] = ""

    df = clean_common_columns(df, key_col, name_col)
    df[amount_col] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0)
    df["__parsed_date"] = pd.to_datetime(df[date_col], errors="coerce")

    grouped = (
        df.groupby([key_col], as_index=False)
        .agg(
            B_名前=(name_col, "first"),
            明細金額合計=(amount_col, "sum"),
            明細件数=(amount_col, "count"),
            最初日付=("__parsed_date", "min"),
            最終日付=("__parsed_date", "max"),
        )
    )

    for col in ["最初日付", "最終日付"]:
        grouped[col] = grouped[col].dt.strftime("%Y-%m-%d").fillna("")

    return grouped


def prepare_single_row_csv(
    df: pd.DataFrame,
    key_col: str,
    name_col: str,
    file_label: str,
    prefix: str,
) -> pd.DataFrame:
    df = normalize_columns(df)
    require_columns(df, [key_col], file_label)

    if name_col not in df.columns:
        df[name_col] = ""

    df = clean_common_columns(df, key_col, name_col)
    validate_unique_key(df, key_col, file_label)

    rename_map: dict[str, str] = {}
    for col in df.columns:
        if col in [key_col, name_col]:
            continue
        rename_map[col] = f"{prefix}{col}" if col in ["明細金額合計", "明細件数", "最初日付", "最終日付", "B_名前"] else col

    return df.rename(columns=rename_map)


def merge_csvs(config: MergeConfig) -> pd.DataFrame:
    raw_a = normalize_columns(read_csv_flexible(config.file_a))
    raw_b = normalize_columns(read_csv_flexible(config.file_b))
    raw_c = normalize_columns(read_csv_flexible(config.file_c))

    df_a = prepare_single_row_csv(raw_a, config.key_col, config.name_col, "A.csv", "A_")
    df_c = prepare_single_row_csv(raw_c, config.key_col, config.name_col, "C.csv", "C_")
    df_b = aggregate_b(raw_b, config.key_col, config.name_col, config.b_date_col, config.b_amount_col)

    merged = pd.merge(df_a, df_b, on=config.key_col, how="outer")
    merged = pd.merge(merged, df_c, on=config.key_col, how="outer", suffixes=("", "_C"))

    a_name_col = config.name_col if config.name_col in merged.columns else ""
    c_name_candidates = [col for col in [config.name_col + "_C", config.name_col] if col in merged.columns]
    preferred_name_cols = [col for col in [a_name_col, "B_名前", *c_name_candidates] if col]
    merged[config.name_col] = merged.apply(lambda row: choose_best_name(row, preferred_name_cols), axis=1)

    drop_candidates = [col for col in ["B_名前", config.name_col + "_C"] if col in merged.columns]
    merged = merged.drop(columns=drop_candidates, errors="ignore")

    ordered_first = [config.key_col, config.name_col]
    other_cols = [col for col in merged.columns if col not in ordered_first]
    merged = merged[ordered_first + other_cols]
    merged = merged.sort_values(by=[config.key_col], kind="stable").reset_index(drop=True)
    return merged


def save_csv(df: pd.DataFrame, output_file: Path) -> None:
    df.to_csv(output_file, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_MINIMAL)


# =========================
# UI helpers
# =========================

def neon_box(parent: tk.Widget, title: str, color: str) -> tuple[tk.Frame, tk.Frame]:
    outer = tk.Frame(parent, bg=color, bd=0, highlightthickness=0)
    inner = tk.Frame(outer, bg=COLORS["panel"], bd=0, highlightthickness=0)
    inner.pack(fill="both", expand=True, padx=1, pady=1)
    title_label = tk.Label(
        inner,
        text=title,
        bg=COLORS["panel"],
        fg=color,
        font=("Yu Gothic UI", 12, "bold"),
        anchor="w",
    )
    title_label.pack(fill="x", padx=16, pady=(14, 8))
    body = tk.Frame(inner, bg=COLORS["panel"])
    body.pack(fill="both", expand=True, padx=16, pady=(0, 16))
    return outer, body


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_TITLE} v{APP_VERSION}")
        self.geometry("1040x700")
        self.minsize(980, 660)
        self.configure(bg=COLORS["bg"])

        self.file_a_var = tk.StringVar()
        self.file_b_var = tk.StringVar()
        self.file_c_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.key_col_var = tk.StringVar(value="社員番号")
        self.name_col_var = tk.StringVar(value="名前")
        self.b_date_col_var = tk.StringVar(value="日付")
        self.b_amount_col_var = tk.StringVar(value="交通費")
        self.status_var = tk.StringVar(value="A/C は基本CSV、B は明細CSVです。キー列を設定して実行してください。")

        self._setup_style()
        self._build_ui()

    def _setup_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TButton", font=("Yu Gothic UI", 10, "bold"), padding=(12, 8))
        style.configure("TEntry", padding=7, fieldbackground=COLORS["entry_bg"], foreground=COLORS["text"])

    def _build_ui(self) -> None:
        root = tk.Frame(self, bg=COLORS["bg"])
        root.pack(fill="both", expand=True, padx=18, pady=18)

        header = tk.Frame(root, bg=COLORS["bg"])
        header.pack(fill="x", pady=(0, 14))

        tk.Label(
            header,
            text="✦ CSV Joiner Neon",
            bg=COLORS["bg"],
            fg=COLORS["text"],
            font=("Yu Gothic UI", 24, "bold"),
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            header,
            text="モダンでネオン寄りのCSV横結合ツール。キー列は自由、B.csv は明細集約用。",
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=("Yu Gothic UI", 10),
            anchor="w",
        ).pack(anchor="w", pady=(4, 0))

        content = tk.Frame(root, bg=COLORS["bg"])
        content.pack(fill="both", expand=True)

        left = tk.Frame(content, bg=COLORS["bg"])
        left.pack(side="left", fill="both", expand=True)

        right = tk.Frame(content, bg=COLORS["bg"], width=300)
        right.pack(side="left", fill="y", padx=(14, 0))
        right.pack_propagate(False)

        box1, body1 = neon_box(left, "01 FILES", COLORS["cyan"])
        box1.pack(fill="x", pady=(0, 12))
        self._add_file_row(body1, 0, "A.csv", self.file_a_var)
        self._add_file_row(body1, 1, "B.csv", self.file_b_var)
        self._add_file_row(body1, 2, "C.csv", self.file_c_var)
        self._add_save_row(body1, 3, "出力先", self.output_var)

        box2, body2 = neon_box(left, "02 COMMON SETTINGS", COLORS["pink"])
        box2.pack(fill="x", pady=(0, 12))
        self._add_setting_row(body2, 0, "キー列", self.key_col_var, "社員番号 / 出席番号 / 会員ID")
        self._add_setting_row(body2, 1, "表示名列", self.name_col_var, "名前 / 氏名 / 生徒名")

        box3, body3 = neon_box(left, "03 B DETAIL SETTINGS", COLORS["violet"])
        box3.pack(fill="x")
        self._add_setting_row(body3, 0, "日付列", self.b_date_col_var, "日付 / 利用日 / 年月日")
        self._add_setting_row(body3, 1, "金額列", self.b_amount_col_var, "交通費 / 金額 / 費用")

        action_row = tk.Frame(left, bg=COLORS["bg"])
        action_row.pack(fill="x", pady=(14, 0))
        self._create_neon_button(action_row, "✨ 結合実行", COLORS["pink"], self.run_merge).pack(side="left")
        self._create_neon_button(action_row, "📍 出力先を自動入力", COLORS["cyan"], self.fill_default_output).pack(side="left", padx=(10, 0))
        self._create_neon_button(action_row, "🔎 B列候補を自動入力", COLORS["violet"], self.autofill_column_candidates).pack(side="left", padx=(10, 0))

        info1, infobody1 = neon_box(right, "SYSTEM", COLORS["cyan"])
        info1.pack(fill="x", pady=(0, 12))
        tk.Label(
            infobody1,
            text="A/C はキーごとに1行\nB はキーごとに複数行OK\nB は合計・件数・最小日付・最大日付へ集約",
            bg=COLORS["panel"], fg=COLORS["text"], justify="left", font=("Yu Gothic UI", 10),
        ).pack(anchor="w")

        info2, infobody2 = neon_box(right, "STATUS", COLORS["pink"])
        info2.pack(fill="x", pady=(0, 12))
        self.status_label = tk.Label(
            infobody2,
            textvariable=self.status_var,
            bg=COLORS["panel2"],
            fg=COLORS["text"],
            justify="left",
            wraplength=240,
            padx=12,
            pady=12,
            font=("Yu Gothic UI", 10),
        )
        self.status_label.pack(fill="x")

        info3, infobody3 = neon_box(right, "TIPS", COLORS["violet"])
        info3.pack(fill="x")
        tk.Label(
            infobody3,
            text="・名前列が無くてもOK\n・Bだけ設定が多いのは明細CSVだから\n・汎用キー列で学校/会員/勤怠にも転用可能",
            bg=COLORS["panel"], fg=COLORS["muted"], justify="left", font=("Yu Gothic UI", 10),
        ).pack(anchor="w")

    def _create_neon_button(self, parent: tk.Widget, text: str, color: str, command) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=color,
            fg="#ffffff",
            activebackground=color,
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            padx=14,
            pady=10,
            font=("Yu Gothic UI", 10, "bold"),
            cursor="hand2",
        )

    def _styled_entry(self, parent: tk.Widget, textvariable: tk.StringVar, width: int = 52) -> tk.Entry:
        return tk.Entry(
            parent,
            textvariable=textvariable,
            width=width,
            bg=COLORS["entry_bg"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            highlightthickness=1,
            highlightbackground=COLORS["panel2"],
            highlightcolor=COLORS["cyan"],
            font=("Yu Gothic UI", 10),
        )

    def _small_button(self, parent: tk.Widget, text: str, command) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=COLORS["panel2"],
            fg=COLORS["text"],
            activebackground=COLORS["cyan"],
            activeforeground="#08111f",
            relief="flat",
            bd=0,
            padx=10,
            pady=7,
            font=("Yu Gothic UI", 9, "bold"),
            cursor="hand2",
        )

    def _add_file_row(self, parent: tk.Frame, row: int, label: str, variable: tk.StringVar) -> None:
        frame = tk.Frame(parent, bg=COLORS["panel"])
        frame.grid(row=row, column=0, sticky="ew", pady=6)
        tk.Label(frame, text=label, bg=COLORS["panel"], fg=COLORS["text"], width=8, anchor="w", font=("Yu Gothic UI", 10)).pack(side="left")
        self._styled_entry(frame, variable, width=58).pack(side="left", fill="x", expand=True, padx=(0, 10))
        self._small_button(frame, "参照", lambda v=variable: self.choose_file(v)).pack(side="left")
        parent.grid_columnconfigure(0, weight=1)

    def _add_save_row(self, parent: tk.Frame, row: int, label: str, variable: tk.StringVar) -> None:
        frame = tk.Frame(parent, bg=COLORS["panel"])
        frame.grid(row=row, column=0, sticky="ew", pady=6)
        tk.Label(frame, text=label, bg=COLORS["panel"], fg=COLORS["text"], width=8, anchor="w", font=("Yu Gothic UI", 10)).pack(side="left")
        self._styled_entry(frame, variable, width=58).pack(side="left", fill="x", expand=True, padx=(0, 10))
        self._small_button(frame, "保存", lambda: self.choose_output(variable)).pack(side="left")
        parent.grid_columnconfigure(0, weight=1)

    def _add_setting_row(self, parent: tk.Frame, row: int, label: str, variable: tk.StringVar, hint: str) -> None:
        wrapper = tk.Frame(parent, bg=COLORS["panel"])
        wrapper.grid(row=row, column=0, sticky="ew", pady=6)
        top = tk.Frame(wrapper, bg=COLORS["panel"])
        top.pack(fill="x")
        tk.Label(top, text=label, bg=COLORS["panel"], fg=COLORS["text"], width=10, anchor="w", font=("Yu Gothic UI", 10, "bold")).pack(side="left")
        tk.Label(top, text=hint, bg=COLORS["panel"], fg=COLORS["muted"], anchor="w", font=("Yu Gothic UI", 9)).pack(side="left", padx=(8, 0))
        self._styled_entry(wrapper, variable, width=56).pack(fill="x", pady=(6, 0))
        parent.grid_columnconfigure(0, weight=1)

    def choose_file(self, variable: tk.StringVar) -> None:
        path = filedialog.askopenfilename(
            title="CSVを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            variable.set(path)

    def choose_output(self, variable: tk.StringVar) -> None:
        path = filedialog.asksaveasfilename(
            title="出力先を選択",
            defaultextension=".csv",
            initialfile="merged.csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            variable.set(path)

    def fill_default_output(self) -> None:
        base = None
        for candidate in [self.file_a_var.get(), self.file_b_var.get(), self.file_c_var.get()]:
            if candidate:
                base = Path(candidate).resolve().parent
                break
        if base is None:
            base = Path.cwd()
        self.output_var.set(str(base / "merged.csv"))
        self.status_var.set("出力先を自動入力しました。")

    def autofill_column_candidates(self) -> None:
        try:
            b_path = self.file_b_var.get().strip()
            if not b_path:
                raise CsvMergeError("先に B.csv を選択してください。")

            b_df = normalize_columns(read_csv_flexible(Path(b_path)))
            columns = list(map(str, b_df.columns))

            if not self.b_date_col_var.get().strip() or self.b_date_col_var.get().strip() not in columns:
                self.b_date_col_var.set(find_first_existing(columns, DEFAULT_B_DATE_CANDIDATES))

            if not self.b_amount_col_var.get().strip() or self.b_amount_col_var.get().strip() not in columns:
                self.b_amount_col_var.set(find_first_existing(columns, DEFAULT_B_AMOUNT_CANDIDATES))

            self.status_var.set("B.csv の候補列を自動入力しました。")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))

    def run_merge(self) -> None:
        try:
            config = self._get_config()
            merged = merge_csvs(config)
            save_csv(merged, config.output_file)
            self.status_var.set(f"出力完了: {config.output_file}")
            messagebox.showinfo(APP_TITLE, f"CSVを出力しました。\n{config.output_file}")
        except CsvMergeError as exc:
            self.status_var.set("エラーが発生しました。")
            messagebox.showerror(APP_TITLE, str(exc))
        except Exception as exc:
            self.status_var.set("想定外エラーが発生しました。")
            messagebox.showerror(APP_TITLE, f"想定外エラー:\n{exc}")

    def _get_config(self) -> MergeConfig:
        file_a_text = self.file_a_var.get().strip()
        file_b_text = self.file_b_var.get().strip()
        file_c_text = self.file_c_var.get().strip()
        output_text = self.output_var.get().strip()

        if not file_a_text:
            raise CsvMergeError("A.csv を選択してください。")
        if not file_b_text:
            raise CsvMergeError("B.csv を選択してください。")
        if not file_c_text:
            raise CsvMergeError("C.csv を選択してください。")
        if not output_text:
            raise CsvMergeError("出力先を指定してください。")

        file_a = Path(file_a_text)
        file_b = Path(file_b_text)
        file_c = Path(file_c_text)
        output_file = Path(output_text)

        for label, path in [("A.csv", file_a), ("B.csv", file_b), ("C.csv", file_c)]:
            if not path.exists():
                raise CsvMergeError(f"{label} が存在しません: {path}")

        key_col = self.key_col_var.get().strip()
        name_col = self.name_col_var.get().strip() or "名前"
        b_date_col = self.b_date_col_var.get().strip()
        b_amount_col = self.b_amount_col_var.get().strip()

        if not key_col:
            raise CsvMergeError("キー列を入力してください。")
        if not b_date_col:
            raise CsvMergeError("B の日付列を入力してください。")
        if not b_amount_col:
            raise CsvMergeError("B の金額列を入力してください。")

        return MergeConfig(
            file_a=file_a,
            file_b=file_b,
            file_c=file_c,
            output_file=output_file,
            key_col=key_col,
            name_col=name_col,
            b_date_col=b_date_col,
            b_amount_col=b_amount_col,
        )


# =========================
# CLI entry point
# =========================

def run_cli(args: list[str]) -> int:
    if len(args) < 4:
        print(
            "使い方:\n"
            "python CSVjoiner.py A.csv B.csv C.csv merged.csv [key_col] [name_col] [b_date_col] [b_amount_col]",
            file=sys.stderr,
        )
        return 1

    key_col = args[4] if len(args) >= 5 else "社員番号"
    name_col = args[5] if len(args) >= 6 else "名前"
    b_date_col = args[6] if len(args) >= 7 else "日付"
    b_amount_col = args[7] if len(args) >= 8 else "交通費"

    config = MergeConfig(
        file_a=Path(args[0]),
        file_b=Path(args[1]),
        file_c=Path(args[2]),
        output_file=Path(args[3]),
        key_col=key_col,
        name_col=name_col,
        b_date_col=b_date_col,
        b_amount_col=b_amount_col,
    )

    try:
        merged = merge_csvs(config)
        save_csv(merged, config.output_file)
        print(f"出力完了: {config.output_file}")
        return 0
    except Exception as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 2


def main() -> int:
    if len(sys.argv) > 1:
        return run_cli(sys.argv[1:])

    app = App()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())