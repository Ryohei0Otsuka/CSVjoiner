from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from . import __version__
from .core import (
    CsvToolError,
    EXPORT_ENCODINGS,
    dataframe_preview,
    export_csv,
    format_preview_value,
    read_csv_flexible,
)
from .models import CsvDocument, TransformStep
from .operations import (
    apply_transform_pipeline,
    export_split_groups,
    key_join,
    split_counts,
    split_dataframe,
    vertical_join,
)


APP_TITLE = "CSVjoiner v2"
COLORS = {
    "bg": "#0d1220",
    "panel": "#151c2f",
    "panel2": "#1c2540",
    "text": "#edf2ff",
    "muted": "#9aa8c7",
    "accent": "#48d7ff",
    "accent2": "#ff5cc8",
    "success": "#4be6a0",
    "warning": "#ffc857",
    "danger": "#ff7285",
}


class PreviewTable(ttk.Frame):
    def __init__(self, parent: tk.Widget, height: int = 12) -> None:
        super().__init__(parent)
        self.tree = ttk.Treeview(self, show="headings", height=height)
        ybar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        xbar = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

    def show_dataframe(self, df: pd.DataFrame, limit: int = 100) -> None:
        frame = dataframe_preview(df, limit)
        self.tree.delete(*self.tree.get_children())
        columns = [str(c) for c in frame.columns]
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, minwidth=70, stretch=True)
        for row in frame.itertuples(index=False, name=None):
            self.tree.insert("", "end", values=[format_preview_value(v) for v in row])

    def clear(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []


class CSVJoinerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_TITLE} {__version__}")
        self.geometry("1280x840")
        self.minsize(1080, 720)
        self.configure(bg=COLORS["bg"])
        self._setup_style()
        self._build_header()
        self._build_tabs()
        self.status_var = tk.StringVar(value="CSVを選択して処理を開始してください。")
        self._build_status_bar()

    def _setup_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background=COLORS["bg"])
        style.configure("Panel.TFrame", background=COLORS["panel"])
        style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["text"], font=("Yu Gothic UI", 10))
        style.configure("Muted.TLabel", background=COLORS["bg"], foreground=COLORS["muted"], font=("Yu Gothic UI", 9))
        style.configure("Panel.TLabel", background=COLORS["panel"], foreground=COLORS["text"], font=("Yu Gothic UI", 10))
        style.configure("Title.TLabel", background=COLORS["bg"], foreground=COLORS["text"], font=("Yu Gothic UI", 24, "bold"))
        style.configure("TButton", font=("Yu Gothic UI", 10, "bold"), padding=(11, 7))
        style.configure("Accent.TButton", font=("Yu Gothic UI", 10, "bold"), padding=(12, 8))
        style.configure("TNotebook", background=COLORS["bg"], borderwidth=0)
        style.configure("TNotebook.Tab", padding=(18, 9), font=("Yu Gothic UI", 10, "bold"))
        style.configure("Treeview", background="#11182a", fieldbackground="#11182a", foreground=COLORS["text"], rowheight=26)
        style.configure("Treeview.Heading", font=("Yu Gothic UI", 9, "bold"))
        style.map("Treeview", background=[("selected", "#2b5570")])

    def _build_header(self) -> None:
        header = ttk.Frame(self)
        header.pack(fill="x", padx=20, pady=(18, 10))
        ttk.Label(header, text="CSVjoiner v2", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="JOIN / SPLIT / TRANSFORM — CSV加工を、実行前に確認してから出力する。",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(3, 0))

    def _build_tabs(self) -> None:
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        self.join_tab = ttk.Frame(self.notebook)
        self.split_tab = ttk.Frame(self.notebook)
        self.transform_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.join_tab, text="JOIN")
        self.notebook.add(self.split_tab, text="SPLIT")
        self.notebook.add(self.transform_tab, text="TRANSFORM")
        self._build_join_tab()
        self._build_split_tab()
        self._build_transform_tab()

    def _build_status_bar(self) -> None:
        bar = tk.Frame(self, bg=COLORS["panel2"], height=34)
        bar.pack(fill="x", side="bottom")
        tk.Label(
            bar,
            textvariable=self.status_var,
            bg=COLORS["panel2"], fg=COLORS["text"],
            font=("Yu Gothic UI", 9), anchor="w", padx=14,
        ).pack(fill="x", pady=7)

    def _panel(self, parent: tk.Widget, title: str) -> tuple[tk.Frame, tk.Frame]:
        outer = tk.Frame(parent, bg=COLORS["panel"], highlightbackground="#293555", highlightthickness=1)
        tk.Label(
            outer, text=title, bg=COLORS["panel"], fg=COLORS["accent"],
            font=("Yu Gothic UI", 10, "bold"), anchor="w",
        ).pack(fill="x", padx=14, pady=(11, 6))
        body = tk.Frame(outer, bg=COLORS["panel"])
        body.pack(fill="both", expand=True, padx=14, pady=(0, 12))
        return outer, body

    def _entry(self, parent: tk.Widget, variable: tk.Variable, width: int = 34) -> tk.Entry:
        return tk.Entry(
            parent, textvariable=variable, width=width,
            bg="#0f1628", fg=COLORS["text"], insertbackground=COLORS["text"],
            relief="flat", highlightthickness=1, highlightbackground="#32405f",
            highlightcolor=COLORS["accent"], font=("Yu Gothic UI", 10),
        )

    def _label(self, parent: tk.Widget, text: str, width: int | None = None) -> tk.Label:
        return tk.Label(
            parent, text=text, width=width, bg=COLORS["panel"], fg=COLORS["text"],
            font=("Yu Gothic UI", 10), anchor="w",
        )

    def _browse_csv(self) -> str:
        return filedialog.askopenfilename(title="CSVを選択", filetypes=[("CSV files", "*.csv")])

    def _save_csv_path(self, initialfile: str) -> str:
        return filedialog.asksaveasfilename(
            title="CSV出力先",
            defaultextension=".csv",
            initialfile=initialfile,
            filetypes=[("CSV files", "*.csv")],
        )

    def _choose_directory(self) -> str:
        return filedialog.askdirectory(title="出力フォルダを選択")

    # JOIN ----------------------------------------------------------
    def _build_join_tab(self) -> None:
        self.join_mode = tk.StringVar(value="key")
        self.join_files: list[CsvDocument] = []
        self.join_left: CsvDocument | None = None
        self.join_right: CsvDocument | None = None
        self.join_result: pd.DataFrame | None = None
        self.join_left_key = tk.StringVar()
        self.join_right_key = tk.StringVar()
        self.join_type = tk.StringVar(value="left")
        self.join_encoding = tk.StringVar(value="UTF-8 BOM")
        self.join_output = tk.StringVar()
        self.join_diag = tk.StringVar(value="未解析")

        container = ttk.Frame(self.join_tab)
        container.pack(fill="both", expand=True, padx=4, pady=8)
        top = ttk.Frame(container)
        top.pack(fill="x")

        p1, b1 = self._panel(top, "01 MODE / FILES")
        p1.pack(side="left", fill="both", expand=True, padx=(0, 6))
        p2, b2 = self._panel(top, "02 SETTINGS / CHECK")
        p2.pack(side="left", fill="both", expand=True, padx=(6, 0))

        mode_row = tk.Frame(b1, bg=COLORS["panel"])
        mode_row.pack(fill="x", pady=(0, 8))
        for text, value in [("キー結合", "key"), ("縦結合", "vertical")]:
            tk.Radiobutton(
                mode_row, text=text, variable=self.join_mode, value=value,
                command=self._refresh_join_mode,
                bg=COLORS["panel"], fg=COLORS["text"], selectcolor="#263452",
                activebackground=COLORS["panel"], activeforeground=COLORS["text"],
            ).pack(side="left", padx=(0, 14))

        self.join_file_summary = tk.Label(
            b1, text="キー結合: 左CSVと右CSVを選択", justify="left", anchor="w",
            bg="#101729", fg=COLORS["text"], padx=10, pady=10,
            font=("Yu Gothic UI", 9),
        )
        self.join_file_summary.pack(fill="x", pady=(0, 8))
        action = tk.Frame(b1, bg=COLORS["panel"])
        action.pack(fill="x")
        ttk.Button(action, text="左CSV", command=lambda: self._load_join_side("left")).pack(side="left")
        ttk.Button(action, text="右CSV", command=lambda: self._load_join_side("right")).pack(side="left", padx=6)
        ttk.Button(action, text="縦結合CSVを複数選択", command=self._load_vertical_files).pack(side="left")

        row = tk.Frame(b2, bg=COLORS["panel"])
        row.pack(fill="x", pady=3)
        self._label(row, "左キー", 10).pack(side="left")
        self.join_left_combo = ttk.Combobox(row, textvariable=self.join_left_key, state="readonly", width=24)
        self.join_left_combo.pack(side="left", padx=(0, 10))
        self._label(row, "右キー", 10).pack(side="left")
        self.join_right_combo = ttk.Combobox(row, textvariable=self.join_right_key, state="readonly", width=24)
        self.join_right_combo.pack(side="left")

        row2 = tk.Frame(b2, bg=COLORS["panel"])
        row2.pack(fill="x", pady=(6, 3))
        self._label(row2, "JOIN TYPE", 10).pack(side="left")
        ttk.Combobox(row2, textvariable=self.join_type, state="readonly", values=["left", "inner"], width=10).pack(side="left")
        ttk.Button(row2, text="CHECK / PREVIEW", command=self._run_join_preview).pack(side="left", padx=10)

        tk.Label(
            b2, textvariable=self.join_diag, justify="left", anchor="w", wraplength=520,
            bg="#101729", fg=COLORS["text"], padx=10, pady=9, font=("Yu Gothic UI", 9),
        ).pack(fill="x", pady=(8, 0))

        preview_panel, preview_body = self._panel(container, "03 PREVIEW")
        preview_panel.pack(fill="both", expand=True, pady=(10, 0))
        self.join_preview = PreviewTable(preview_body, height=14)
        self.join_preview.pack(fill="both", expand=True)

        export_row = tk.Frame(preview_body, bg=COLORS["panel"])
        export_row.pack(fill="x", pady=(10, 0))
        self._entry(export_row, self.join_output, 54).pack(side="left", fill="x", expand=True)
        ttk.Button(export_row, text="保存先", command=self._choose_join_output).pack(side="left", padx=6)
        ttk.Combobox(export_row, textvariable=self.join_encoding, state="readonly", values=list(EXPORT_ENCODINGS), width=12).pack(side="left")
        ttk.Button(export_row, text="EXPORT", command=self._export_join).pack(side="left", padx=(6, 0))
        self._refresh_join_mode()

    def _refresh_join_mode(self) -> None:
        key_mode = self.join_mode.get() == "key"
        state = "readonly" if key_mode else "disabled"
        self.join_left_combo.configure(state=state)
        self.join_right_combo.configure(state=state)
        if key_mode:
            self.join_file_summary.configure(text="キー結合: 左CSVと右CSVを選択")
        else:
            self.join_file_summary.configure(text="縦結合: 2つ以上のCSVをまとめて選択")

    def _load_join_side(self, side: str) -> None:
        path = self._browse_csv()
        if not path:
            return
        try:
            doc = read_csv_flexible(Path(path))
            if side == "left":
                self.join_left = doc
                self.join_left_combo["values"] = doc.columns
                self.join_left_key.set(doc.columns[0] if doc.columns else "")
            else:
                self.join_right = doc
                self.join_right_combo["values"] = doc.columns
                self.join_right_key.set(doc.columns[0] if doc.columns else "")
            self.join_file_summary.configure(text=self._join_file_text())
            self.join_result = None
            self.join_preview.clear()
            self.status_var.set(f"{doc.path.name} を読み込みました。")
        except Exception as exc:
            self._show_error(exc)

    def _load_vertical_files(self) -> None:
        paths = filedialog.askopenfilenames(title="縦結合するCSVを選択", filetypes=[("CSV files", "*.csv")])
        if not paths:
            return
        try:
            self.join_files = [read_csv_flexible(Path(p)) for p in paths]
            self.join_file_summary.configure(text=self._join_file_text())
            self.status_var.set(f"{len(self.join_files)} ファイルを読み込みました。")
        except Exception as exc:
            self._show_error(exc)

    def _join_file_text(self) -> str:
        if self.join_mode.get() == "vertical":
            if not self.join_files:
                return "縦結合: 2つ以上のCSVをまとめて選択"
            return "\n".join(f"{d.path.name}  /  {d.rows:,} rows  /  {len(d.columns)} cols  /  {d.encoding}" for d in self.join_files)
        parts = []
        if self.join_left:
            parts.append(f"LEFT  {self.join_left.path.name}  /  {self.join_left.rows:,} rows  /  {self.join_left.encoding}")
        if self.join_right:
            parts.append(f"RIGHT {self.join_right.path.name}  /  {self.join_right.rows:,} rows  /  {self.join_right.encoding}")
        return "\n".join(parts) if parts else "キー結合: 左CSVと右CSVを選択"

    def _run_join_preview(self) -> None:
        try:
            if self.join_mode.get() == "vertical":
                self.join_result = vertical_join([d.dataframe for d in self.join_files])
                self.join_diag.set(
                    f"入力: {len(self.join_files)} files / 出力予定: {len(self.join_result):,} rows / {len(self.join_result.columns)} cols\n"
                    "不足列は空欄で補完して縦方向に結合します。"
                )
                base_dir = self.join_files[0].path.parent if self.join_files else Path.cwd()
                self.join_output.set(str(base_dir / "merged.csv"))
            else:
                if not self.join_left or not self.join_right:
                    raise CsvToolError("左CSVと右CSVを選択してください。")
                self.join_result, diag = key_join(
                    self.join_left.dataframe,
                    self.join_right.dataframe,
                    self.join_left_key.get(),
                    self.join_right_key.get(),
                    self.join_type.get(),
                )
                warning = " / ".join(diag.warnings) if diag.warnings else "警告なし"
                self.join_diag.set(
                    f"一致キー {diag.matched_keys:,} / 左のみ {diag.left_only_keys:,} / 右のみ {diag.right_only_keys:,}\n"
                    f"左重複 {diag.left_duplicate_keys:,} / 右重複 {diag.right_duplicate_keys:,} / 多対多 {diag.many_to_many_keys:,}\n"
                    f"出力予定 {diag.result_rows:,} rows / {warning}"
                )
                self.join_output.set(str(self.join_left.path.parent / "merged.csv"))
            self.join_preview.show_dataframe(self.join_result)
            self.status_var.set("JOIN結果をプレビューしました。内容を確認してから出力してください。")
        except Exception as exc:
            self._show_error(exc)

    def _choose_join_output(self) -> None:
        path = self._save_csv_path("merged.csv")
        if path:
            self.join_output.set(path)

    def _export_join(self) -> None:
        try:
            if self.join_result is None:
                raise CsvToolError("先に CHECK / PREVIEW を実行してください。")
            if not self.join_output.get().strip():
                raise CsvToolError("出力先を指定してください。")
            export_csv(self.join_result, Path(self.join_output.get()), self.join_encoding.get())
            self.status_var.set(f"出力完了: {self.join_output.get()}")
            messagebox.showinfo(APP_TITLE, "CSVを出力しました。")
        except Exception as exc:
            self._show_error(exc)

    # SPLIT ---------------------------------------------------------
    def _build_split_tab(self) -> None:
        self.split_doc: CsvDocument | None = None
        self.split_key = tk.StringVar()
        self.split_include_blank = tk.BooleanVar(value=False)
        self.split_output_dir = tk.StringVar()
        self.split_encoding = tk.StringVar(value="UTF-8 BOM")
        self.split_summary = tk.StringVar(value="CSV未選択")

        container = ttk.Frame(self.split_tab)
        container.pack(fill="both", expand=True, padx=4, pady=8)
        top = ttk.Frame(container)
        top.pack(fill="x")
        p1, b1 = self._panel(top, "01 INPUT")
        p1.pack(side="left", fill="both", expand=True, padx=(0, 6))
        p2, b2 = self._panel(top, "02 SPLIT KEY")
        p2.pack(side="left", fill="both", expand=True, padx=(6, 0))

        ttk.Button(b1, text="CSVを選択", command=self._load_split_file).pack(anchor="w")
        tk.Label(
            b1, textvariable=self.split_summary, justify="left", anchor="w",
            bg="#101729", fg=COLORS["text"], padx=10, pady=10, font=("Yu Gothic UI", 9),
        ).pack(fill="x", pady=(8, 0))

        row = tk.Frame(b2, bg=COLORS["panel"])
        row.pack(fill="x", pady=3)
        self._label(row, "キー列", 9).pack(side="left")
        self.split_key_combo = ttk.Combobox(row, textvariable=self.split_key, state="readonly", width=28)
        self.split_key_combo.pack(side="left")
        ttk.Button(row, text="件数CHECK", command=self._preview_split_counts).pack(side="left", padx=8)
        tk.Checkbutton(
            b2, text="空欄キーも blank.csv として出力", variable=self.split_include_blank,
            bg=COLORS["panel"], fg=COLORS["text"], selectcolor="#263452",
            activebackground=COLORS["panel"], activeforeground=COLORS["text"],
        ).pack(anchor="w", pady=(7, 0))

        p3, b3 = self._panel(container, "03 RESULT COUNT PREVIEW")
        p3.pack(fill="both", expand=True, pady=(10, 0))
        self.split_preview = PreviewTable(b3, height=15)
        self.split_preview.pack(fill="both", expand=True)

        export_row = tk.Frame(b3, bg=COLORS["panel"])
        export_row.pack(fill="x", pady=(10, 0))
        self._entry(export_row, self.split_output_dir, 54).pack(side="left", fill="x", expand=True)
        ttk.Button(export_row, text="出力フォルダ", command=self._choose_split_output).pack(side="left", padx=6)
        ttk.Combobox(export_row, textvariable=self.split_encoding, state="readonly", values=list(EXPORT_ENCODINGS), width=12).pack(side="left")
        ttk.Button(export_row, text="SPLIT EXPORT", command=self._export_split).pack(side="left", padx=(6, 0))

    def _load_split_file(self) -> None:
        path = self._browse_csv()
        if not path:
            return
        try:
            self.split_doc = read_csv_flexible(Path(path))
            self.split_key_combo["values"] = self.split_doc.columns
            self.split_key.set(self.split_doc.columns[0] if self.split_doc.columns else "")
            self.split_output_dir.set(str(self.split_doc.path.parent / f"{self.split_doc.path.stem}_split"))
            self.split_summary.set(
                f"{self.split_doc.path.name}\n{self.split_doc.rows:,} rows / {len(self.split_doc.columns)} cols / {self.split_doc.encoding}"
            )
            self.split_preview.show_dataframe(self.split_doc.dataframe)
            self.status_var.set("SPLIT対象CSVを読み込みました。")
        except Exception as exc:
            self._show_error(exc)

    def _preview_split_counts(self) -> None:
        try:
            if not self.split_doc:
                raise CsvToolError("CSVを選択してください。")
            counts = split_counts(self.split_doc.dataframe, self.split_key.get())
            self.split_preview.show_dataframe(counts, 500)
            self.split_summary.set(
                f"{self.split_doc.path.name}\n{self.split_doc.rows:,} rows / 分割候補 {len(counts):,} files"
            )
            self.status_var.set("分割件数をプレビューしました。")
        except Exception as exc:
            self._show_error(exc)

    def _choose_split_output(self) -> None:
        path = self._choose_directory()
        if path:
            self.split_output_dir.set(path)

    def _export_split(self) -> None:
        try:
            if not self.split_doc:
                raise CsvToolError("CSVを選択してください。")
            if not self.split_output_dir.get().strip():
                raise CsvToolError("出力フォルダを指定してください。")
            groups = split_dataframe(
                self.split_doc.dataframe,
                self.split_key.get(),
                include_blank=self.split_include_blank.get(),
            )
            if not groups:
                raise CsvToolError("出力対象がありません。")
            paths = export_split_groups(groups, Path(self.split_output_dir.get()), self.split_encoding.get())
            self.status_var.set(f"SPLIT完了: {len(paths)} files")
            messagebox.showinfo(APP_TITLE, f"{len(paths)} 個のCSVを出力しました。")
        except Exception as exc:
            self._show_error(exc)

    # TRANSFORM -----------------------------------------------------
    def _build_transform_tab(self) -> None:
        self.transform_doc: CsvDocument | None = None
        self.transform_steps: list[TransformStep] = []
        self.transform_result: pd.DataFrame | None = None
        self.transform_summary = tk.StringVar(value="CSV未選択")
        self.transform_operation = tk.StringVar(value="rename")
        self.transform_column = tk.StringVar()
        self.transform_value1 = tk.StringVar()
        self.transform_value2 = tk.StringVar()
        self.transform_output = tk.StringVar()
        self.transform_encoding = tk.StringVar(value="UTF-8 BOM")

        container = ttk.Frame(self.transform_tab)
        container.pack(fill="both", expand=True, padx=4, pady=8)
        top = ttk.Frame(container)
        top.pack(fill="x")
        p1, b1 = self._panel(top, "01 INPUT / PIPELINE")
        p1.pack(side="left", fill="both", expand=True, padx=(0, 6))
        p2, b2 = self._panel(top, "02 ADD TRANSFORM")
        p2.pack(side="left", fill="both", expand=True, padx=(6, 0))

        ttk.Button(b1, text="CSVを選択", command=self._load_transform_file).pack(anchor="w")
        tk.Label(
            b1, textvariable=self.transform_summary, justify="left", anchor="w",
            bg="#101729", fg=COLORS["text"], padx=10, pady=8, font=("Yu Gothic UI", 9),
        ).pack(fill="x", pady=(7, 7))
        self.transform_step_list = tk.Listbox(
            b1, height=6, bg="#101729", fg=COLORS["text"], selectbackground="#2b5570",
            relief="flat", font=("Yu Gothic UI", 9),
        )
        self.transform_step_list.pack(fill="both", expand=True)
        list_actions = tk.Frame(b1, bg=COLORS["panel"])
        list_actions.pack(fill="x", pady=(6, 0))
        ttk.Button(list_actions, text="選択削除", command=self._remove_transform_step).pack(side="left")
        ttk.Button(list_actions, text="全消去", command=self._clear_transform_steps).pack(side="left", padx=6)

        row1 = tk.Frame(b2, bg=COLORS["panel"])
        row1.pack(fill="x", pady=3)
        self._label(row1, "処理", 10).pack(side="left")
        op_values = [
            "rename | 列名変更", "delete | 列削除", "add_fixed | 固定値列追加", "trim | 前後空白削除",
            "replace | 文字列置換", "date | 日付形式変換", "zero_pad | 0埋め", "move | 列順変更",
        ]
        self.transform_op_combo = ttk.Combobox(row1, state="readonly", values=op_values, width=30)
        self.transform_op_combo.set(op_values[0])
        self.transform_op_combo.pack(side="left")
        self.transform_op_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_transform_operation())

        row2 = tk.Frame(b2, bg=COLORS["panel"])
        row2.pack(fill="x", pady=3)
        self._label(row2, "対象列", 10).pack(side="left")
        self.transform_col_combo = ttk.Combobox(row2, textvariable=self.transform_column, state="readonly", width=30)
        self.transform_col_combo.pack(side="left")

        row3 = tk.Frame(b2, bg=COLORS["panel"])
        row3.pack(fill="x", pady=3)
        self.transform_v1_label = self._label(row3, "値1", 10)
        self.transform_v1_label.pack(side="left")
        self._entry(row3, self.transform_value1, 32).pack(side="left", fill="x", expand=True)

        row4 = tk.Frame(b2, bg=COLORS["panel"])
        row4.pack(fill="x", pady=3)
        self.transform_v2_label = self._label(row4, "値2", 10)
        self.transform_v2_label.pack(side="left")
        self._entry(row4, self.transform_value2, 32).pack(side="left", fill="x", expand=True)

        ttk.Button(b2, text="ADD STEP", command=self._add_transform_step).pack(anchor="w", pady=(8, 2))
        self.transform_hint = tk.Label(
            b2, text="", justify="left", anchor="w", wraplength=520,
            bg="#101729", fg=COLORS["muted"], padx=10, pady=8, font=("Yu Gothic UI", 9),
        )
        self.transform_hint.pack(fill="x", pady=(5, 0))
        self._sync_transform_operation()

        p3, b3 = self._panel(container, "03 BEFORE / AFTER PREVIEW")
        p3.pack(fill="both", expand=True, pady=(10, 0))
        preview_split = tk.PanedWindow(b3, orient="horizontal", bg=COLORS["panel"], sashwidth=6)
        preview_split.pack(fill="both", expand=True)
        left = tk.Frame(preview_split, bg=COLORS["panel"])
        right = tk.Frame(preview_split, bg=COLORS["panel"])
        preview_split.add(left, stretch="always")
        preview_split.add(right, stretch="always")
        tk.Label(left, text="BEFORE", bg=COLORS["panel"], fg=COLORS["muted"], font=("Yu Gothic UI", 9, "bold")).pack(anchor="w")
        tk.Label(right, text="AFTER", bg=COLORS["panel"], fg=COLORS["success"], font=("Yu Gothic UI", 9, "bold")).pack(anchor="w")
        self.transform_before = PreviewTable(left, height=12)
        self.transform_after = PreviewTable(right, height=12)
        self.transform_before.pack(fill="both", expand=True, pady=(3, 0))
        self.transform_after.pack(fill="both", expand=True, pady=(3, 0))

        export_row = tk.Frame(b3, bg=COLORS["panel"])
        export_row.pack(fill="x", pady=(10, 0))
        ttk.Button(export_row, text="APPLY / PREVIEW", command=self._apply_transform_preview).pack(side="left")
        self._entry(export_row, self.transform_output, 48).pack(side="left", fill="x", expand=True, padx=(8, 0))
        ttk.Button(export_row, text="保存先", command=self._choose_transform_output).pack(side="left", padx=6)
        ttk.Combobox(export_row, textvariable=self.transform_encoding, state="readonly", values=list(EXPORT_ENCODINGS), width=12).pack(side="left")
        ttk.Button(export_row, text="EXPORT", command=self._export_transform).pack(side="left", padx=(6, 0))

    def _load_transform_file(self) -> None:
        path = self._browse_csv()
        if not path:
            return
        try:
            self.transform_doc = read_csv_flexible(Path(path))
            self.transform_steps.clear()
            self.transform_result = None
            self._refresh_transform_step_list()
            self.transform_col_combo["values"] = self.transform_doc.columns
            self.transform_column.set(self.transform_doc.columns[0] if self.transform_doc.columns else "")
            self.transform_output.set(str(self.transform_doc.path.parent / f"{self.transform_doc.path.stem}_transformed.csv"))
            self.transform_summary.set(
                f"{self.transform_doc.path.name}\n{self.transform_doc.rows:,} rows / {len(self.transform_doc.columns)} cols / {self.transform_doc.encoding}"
            )
            self.transform_before.show_dataframe(self.transform_doc.dataframe)
            self.transform_after.clear()
            self.status_var.set("TRANSFORM対象CSVを読み込みました。")
        except Exception as exc:
            self._show_error(exc)

    def _sync_transform_operation(self) -> None:
        raw = self.transform_op_combo.get()
        op = raw.split("|", 1)[0].strip() if raw else "rename"
        self.transform_operation.set(op)
        hints = {
            "rename": ("新しい列名", "未使用", "例: 氏名"),
            "delete": ("未使用", "未使用", "対象列を削除します。"),
            "add_fixed": ("追加列名", "固定値", "対象列の指定は不要です。例: 処理区分 / 済"),
            "trim": ("未使用", "未使用", "対象列の前後空白を削除します。"),
            "replace": ("置換前", "置換後", "部分一致の文字列を置換します。"),
            "date": ("出力形式", "未使用", "例: %Y-%m-%d。解析できない値は元のまま残します。"),
            "zero_pad": ("桁数", "未使用", "例: 6 → 123 を 000123 にします。"),
            "move": ("移動先", "未使用", "先頭 / 末尾 / 1始まりの列番号を指定します。"),
        }
        l1, l2, hint = hints[op]
        self.transform_v1_label.configure(text=l1)
        self.transform_v2_label.configure(text=l2)
        self.transform_hint.configure(text=hint)

    def _add_transform_step(self) -> None:
        try:
            if not self.transform_doc:
                raise CsvToolError("先にCSVを選択してください。")
            op = self.transform_operation.get()
            column = self.transform_column.get()
            if op != "add_fixed" and not column:
                raise CsvToolError("対象列を選択してください。")
            step = TransformStep(op, column, self.transform_value1.get(), self.transform_value2.get())
            # Validate by trying the pipeline with the new step.
            apply_transform_pipeline(self.transform_doc.dataframe, [*self.transform_steps, step])
            self.transform_steps.append(step)
            self._refresh_transform_step_list()
            self.transform_value1.set("")
            self.transform_value2.set("")
            self.status_var.set("変換ステップを追加しました。")
        except Exception as exc:
            self._show_error(exc)

    def _refresh_transform_step_list(self) -> None:
        self.transform_step_list.delete(0, "end")
        for index, step in enumerate(self.transform_steps, start=1):
            self.transform_step_list.insert("end", f"{index:02d}  {step.describe()}")

    def _remove_transform_step(self) -> None:
        selection = self.transform_step_list.curselection()
        if not selection:
            return
        del self.transform_steps[selection[0]]
        self._refresh_transform_step_list()

    def _clear_transform_steps(self) -> None:
        self.transform_steps.clear()
        self._refresh_transform_step_list()
        self.transform_after.clear()
        self.transform_result = None

    def _apply_transform_preview(self) -> None:
        try:
            if not self.transform_doc:
                raise CsvToolError("CSVを選択してください。")
            if not self.transform_steps:
                raise CsvToolError("変換ステップを1つ以上追加してください。")
            self.transform_result = apply_transform_pipeline(self.transform_doc.dataframe, self.transform_steps)
            self.transform_after.show_dataframe(self.transform_result)
            self.status_var.set(
                f"TRANSFORM結果をプレビューしました: {len(self.transform_result):,} rows / {len(self.transform_result.columns)} cols"
            )
        except Exception as exc:
            self._show_error(exc)

    def _choose_transform_output(self) -> None:
        path = self._save_csv_path("transformed.csv")
        if path:
            self.transform_output.set(path)

    def _export_transform(self) -> None:
        try:
            if self.transform_result is None:
                raise CsvToolError("先に APPLY / PREVIEW を実行してください。")
            if not self.transform_output.get().strip():
                raise CsvToolError("出力先を指定してください。")
            export_csv(self.transform_result, Path(self.transform_output.get()), self.transform_encoding.get())
            self.status_var.set(f"出力完了: {self.transform_output.get()}")
            messagebox.showinfo(APP_TITLE, "CSVを出力しました。")
        except Exception as exc:
            self._show_error(exc)

    def _show_error(self, exc: Exception) -> None:
        self.status_var.set("処理を中断しました。内容を確認してください。")
        if isinstance(exc, CsvToolError):
            messagebox.showerror(APP_TITLE, str(exc))
        else:
            messagebox.showerror(APP_TITLE, f"想定外エラー:\n{exc}")


def run_app() -> None:
    app = CSVJoinerApp()
    app.mainloop()
