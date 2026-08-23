from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from . import __version__
from .core import CsvToolError, EXPORT_ENCODINGS, dataframe_preview, export_csv, format_preview_value, read_csv_flexible
from .models import CsvDocument, TransformStep
from .operations import apply_transform_pipeline, export_split_groups, key_join, split_dataframe, split_plan, vertical_join

APP_TITLE = "CSVjoiner v2"
FONT = "Yu Gothic UI"
COLORS = {
    "app": "#F5F7FB",
    "surface": "#FFFFFF",
    "surface_alt": "#F0F3F8",
    "border": "#DCE3EE",
    "text": "#182233",
    "muted": "#66758A",
    "faint": "#93A0B2",
    "accent": "#356AE6",
    "accent_hover": "#2858C7",
    "accent_soft": "#EAF0FF",
    "success": "#168A68",
    "success_soft": "#E9F7F2",
    "warning": "#A36B00",
    "warning_soft": "#FFF6DF",
    "danger": "#C44656",
    "danger_soft": "#FCECEF",
}


class PreviewTable(tk.Frame):
    def __init__(self, parent: tk.Widget, height: int = 10) -> None:
        super().__init__(parent, bg=COLORS["surface"])
        self.tree = ttk.Treeview(self, show="headings", height=height, style="Data.Treeview")
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
            self.tree.column(col, width=145, minwidth=95, stretch=True, anchor="w")
        for i, row in enumerate(frame.itertuples(index=False, name=None)):
            self.tree.insert("", "end", values=[format_preview_value(v) for v in row], tags=("even" if i % 2 == 0 else "odd",))
        self.tree.tag_configure("even", background="#FFFFFF")
        self.tree.tag_configure("odd", background="#F8FAFD")

    def clear(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []


class CSVJoinerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_TITLE} {__version__}")
        self.geometry("1240x820")
        self.minsize(1040, 700)
        self.configure(bg=COLORS["app"])

        self._setup_style()
        self.pages: dict[str, tk.Frame] = {}
        self.page_canvases: dict[str, tk.Canvas] = {}
        self.current_page = "home"
        self.status_var = tk.StringVar(value="準備できました。やりたいことを選んでください。")

        self._build_shell()
        self._build_home_page()
        self._build_join_page()
        self._build_split_page()
        self._build_transform_page()
        self._show_page("home")
        self.bind_all("<MouseWheel>", self._on_mousewheel, add="+")

    # ---------- shared ----------
    def _setup_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Modern.TCombobox",
            fieldbackground="#FFFFFF",
            background="#FFFFFF",
            foreground=COLORS["text"],
            arrowcolor=COLORS["muted"],
            bordercolor=COLORS["border"],
            lightcolor=COLORS["border"],
            darkcolor=COLORS["border"],
            padding=8,
            font=(FONT, 10),
        )
        style.map(
            "Modern.TCombobox",
            fieldbackground=[("readonly", "#FFFFFF")],
            foreground=[("readonly", COLORS["text"])],
            selectbackground=[("readonly", "#FFFFFF")],
            selectforeground=[("readonly", COLORS["text"])],
        )
        style.configure(
            "Data.Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=COLORS["text"],
            rowheight=30,
            borderwidth=1,
            relief="flat",
            font=(FONT, 9),
        )
        style.configure(
            "Data.Treeview.Heading",
            background="#EEF2F7",
            foreground=COLORS["muted"],
            borderwidth=0,
            relief="flat",
            font=(FONT, 9, "bold"),
            padding=(8, 8),
        )
        style.map("Data.Treeview", background=[("selected", COLORS["accent_soft"])], foreground=[("selected", COLORS["text"])])

    def _build_shell(self) -> None:
        top = tk.Frame(self, bg=COLORS["surface"], height=64, highlightbackground=COLORS["border"], highlightthickness=0)
        top.pack(fill="x")
        top.pack_propagate(False)

        brand = tk.Frame(top, bg=COLORS["surface"])
        brand.pack(side="left", padx=26)
        tk.Label(brand, text="CSVjoiner", bg=COLORS["surface"], fg=COLORS["text"], font=(FONT, 17, "bold")).pack(side="left")
        tk.Label(brand, text=f"v{__version__}", bg=COLORS["surface"], fg=COLORS["faint"], font=(FONT, 9)).pack(side="left", padx=(8, 0), pady=(4, 0))

        self.home_button = self._button(top, "ホーム", lambda: self._show_page("home"), kind="ghost")
        self.home_button.pack(side="right", padx=26, pady=14)

        separator = tk.Frame(self, bg=COLORS["border"], height=1)
        separator.pack(fill="x")

        self.page_host = tk.Frame(self, bg=COLORS["app"])
        self.page_host.pack(fill="both", expand=True)

        status = tk.Frame(self, bg=COLORS["surface"], height=42)
        status.pack(fill="x")
        status.pack_propagate(False)
        tk.Label(status, text="●", bg=COLORS["surface"], fg=COLORS["success"], font=(FONT, 8)).pack(side="left", padx=(26, 7))
        tk.Label(status, textvariable=self.status_var, bg=COLORS["surface"], fg=COLORS["muted"], font=(FONT, 9), anchor="w").pack(side="left", fill="x", expand=True)
        tk.Label(status, text="ローカル処理", bg=COLORS["surface"], fg=COLORS["faint"], font=(FONT, 8)).pack(side="right", padx=26)

    def _page(self, key: str, scroll: bool = True) -> tk.Frame:
        page = tk.Frame(self.page_host, bg=COLORS["app"])
        self.pages[key] = page
        if not scroll:
            return page

        host = tk.Frame(page, bg=COLORS["app"])
        host.pack(fill="both", expand=True, padx=26, pady=22)
        canvas = tk.Canvas(host, bg=COLORS["app"], highlightthickness=0, bd=0)
        vbar = ttk.Scrollbar(host, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y", padx=(8, 0))
        body = tk.Frame(canvas, bg=COLORS["app"])
        window_id = canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window_id, width=e.width))
        self.page_canvases[key] = canvas
        page.body = body  # type: ignore[attr-defined]
        return page

    def _show_page(self, key: str) -> None:
        if key not in self.pages:
            return
        for page in self.pages.values():
            page.pack_forget()
        self.pages[key].pack(fill="both", expand=True)
        self.current_page = key
        if key == "home":
            self.home_button.pack_forget()
        elif not self.home_button.winfo_manager():
            self.home_button.pack(side="right", padx=26, pady=14)
        canvas = self.page_canvases.get(key)
        if canvas:
            canvas.yview_moveto(0)

    def _on_mousewheel(self, event) -> None:
        if getattr(event.widget, "winfo_class", lambda: "")() in {"Treeview", "Listbox", "TCombobox"}:
            return
        canvas = self.page_canvases.get(self.current_page)
        if canvas is not None and event.delta:
            canvas.yview_scroll(int(-event.delta / 120) * 3, "units")

    def _button(self, parent, text, command, kind="secondary", width=None) -> tk.Button:
        palette = {
            "primary": (COLORS["accent"], "#FFFFFF", COLORS["accent_hover"]),
            "secondary": (COLORS["surface_alt"], COLORS["text"], "#E4E9F2"),
            "ghost": (COLORS["surface"], COLORS["muted"], COLORS["surface_alt"]),
            "danger": (COLORS["danger_soft"], COLORS["danger"], "#F8DCE2"),
        }
        bg, fg, active = palette[kind]
        return tk.Button(parent, text=text, command=command, bg=bg, fg=fg, activebackground=active, activeforeground=fg,
                         relief="flat", bd=0, highlightthickness=0, font=(FONT, 9, "bold"), padx=15, pady=9,
                         width=width, cursor="hand2")

    def _card(self, parent, padding=20) -> tk.Frame:
        frame = tk.Frame(parent, bg=COLORS["surface"], highlightbackground=COLORS["border"], highlightthickness=1)
        frame.inner = tk.Frame(frame, bg=COLORS["surface"])  # type: ignore[attr-defined]
        frame.inner.pack(fill="both", expand=True, padx=padding, pady=padding)  # type: ignore[attr-defined]
        return frame

    def _workflow_header(self, parent, eyebrow: str, title: str, subtitle: str) -> None:
        tk.Label(parent, text=eyebrow, bg=COLORS["app"], fg=COLORS["accent"], font=(FONT, 8, "bold")).pack(anchor="w")
        tk.Label(parent, text=title, bg=COLORS["app"], fg=COLORS["text"], font=(FONT, 24, "bold")).pack(anchor="w", pady=(4, 2))
        tk.Label(parent, text=subtitle, bg=COLORS["app"], fg=COLORS["muted"], font=(FONT, 10)).pack(anchor="w")
        self._steps(parent)

    def _steps(self, parent) -> None:
        row = tk.Frame(parent, bg=COLORS["app"])
        row.pack(fill="x", pady=(18, 20))
        labels = [("1", "CSVを選ぶ"), ("2", "条件を決める"), ("3", "確認して保存")]
        for i, (num, text) in enumerate(labels):
            box = tk.Frame(row, bg=COLORS["app"])
            box.pack(side="left")
            tk.Label(box, text=num, bg=COLORS["accent_soft"], fg=COLORS["accent"], font=(FONT, 8, "bold"), width=3, pady=4).pack(side="left")
            tk.Label(box, text=text, bg=COLORS["app"], fg=COLORS["muted"], font=(FONT, 9, "bold")).pack(side="left", padx=(7, 0))
            if i < 2:
                tk.Label(row, text="→", bg=COLORS["app"], fg=COLORS["faint"], font=(FONT, 10)).pack(side="left", padx=16)

    def _section_title(self, parent, step: str, title: str, note: str = "") -> None:
        line = tk.Frame(parent, bg=COLORS["surface"])
        line.pack(fill="x", pady=(0, 12))
        tk.Label(line, text=step, bg=COLORS["accent_soft"], fg=COLORS["accent"], font=(FONT, 8, "bold"), width=4, pady=5).pack(side="left")
        text = tk.Frame(line, bg=COLORS["surface"])
        text.pack(side="left", fill="x", expand=True, padx=(10, 0))
        tk.Label(text, text=title, bg=COLORS["surface"], fg=COLORS["text"], font=(FONT, 11, "bold"), anchor="w").pack(fill="x")
        if note:
            tk.Label(text, text=note, bg=COLORS["surface"], fg=COLORS["muted"], font=(FONT, 8), anchor="w").pack(fill="x", pady=(2, 0))

    def _file_summary(self, parent, variable: tk.StringVar) -> tk.Label:
        return tk.Label(parent, textvariable=variable, bg=COLORS["surface_alt"], fg=COLORS["text"], font=(FONT, 9),
                        justify="left", anchor="w", padx=13, pady=11, wraplength=470)

    def _combo(self, parent, label: str, variable: tk.Variable, values=(), width=26) -> ttk.Combobox:
        wrap = tk.Frame(parent, bg=COLORS["surface"])
        tk.Label(wrap, text=label, bg=COLORS["surface"], fg=COLORS["muted"], font=(FONT, 8, "bold"), anchor="w").pack(fill="x", pady=(0, 5))
        combo = ttk.Combobox(wrap, textvariable=variable, values=values, state="readonly", width=width, style="Modern.TCombobox")
        combo.pack(fill="x")
        return combo

    def _entry(self, parent, label: str, variable: tk.Variable) -> tk.Entry:
        wrap = tk.Frame(parent, bg=COLORS["surface"])
        tk.Label(wrap, text=label, bg=COLORS["surface"], fg=COLORS["muted"], font=(FONT, 8, "bold"), anchor="w").pack(fill="x", pady=(0, 5))
        entry = tk.Entry(wrap, textvariable=variable, bg="#FFFFFF", fg=COLORS["text"], insertbackground=COLORS["text"],
                         relief="flat", bd=0, highlightthickness=1, highlightbackground=COLORS["border"],
                         highlightcolor=COLORS["accent"], font=(FONT, 10))
        entry.pack(fill="x", ipady=8)
        return entry

    def _browse_csv(self) -> str:
        return filedialog.askopenfilename(title="CSVを選択", filetypes=[("CSV files", "*.csv")])

    def _save_csv_path(self, initialfile: str) -> str:
        return filedialog.asksaveasfilename(title="CSV出力先", defaultextension=".csv", initialfile=initialfile, filetypes=[("CSV files", "*.csv")])

    def _scroll_to_widget(self, page_key: str, widget: tk.Widget) -> None:
        canvas = self.page_canvases.get(page_key)
        if canvas is None:
            return
        self.update_idletasks()
        parent = widget.master
        body_height = max(parent.winfo_reqheight(), 1)
        fraction = max(0.0, min(widget.winfo_y() / body_height, 1.0))
        canvas.yview_moveto(fraction)

    # ---------- home ----------
    def _build_home_page(self) -> None:
        page = self._page("home", scroll=False)
        wrap = tk.Frame(page, bg=COLORS["app"])
        wrap.pack(fill="both", expand=True, padx=42, pady=42)
        tk.Label(wrap, text="CSVで何をしますか？", bg=COLORS["app"], fg=COLORS["text"], font=(FONT, 27, "bold")).pack(anchor="w")
        tk.Label(wrap, text="やりたいことを1つ選ぶだけ。処理前に結果を確認できます。", bg=COLORS["app"], fg=COLORS["muted"], font=(FONT, 11)).pack(anchor="w", pady=(6, 28))

        grid = tk.Frame(wrap, bg=COLORS["app"])
        grid.pack(fill="x")
        for i in range(3):
            grid.grid_columnconfigure(i, weight=1, uniform="home")

        choices = [
            ("JOIN", "CSVをまとめる", "行を下につなぐ\nまたはIDで情報をつなぐ", "join"),
            ("SPLIT", "CSVを分ける", "クラス・カテゴリなど\n列の値ごとにファイル分割", "split"),
            ("TRANSFORM", "CSVを整える", "列名・文字列・日付などを\nコードなしで整形", "transform"),
        ]
        for i, (tag, title, desc, key) in enumerate(choices):
            card = self._card(grid, 22)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 7, 0 if i == 2 else 7))
            inner = card.inner  # type: ignore[attr-defined]
            tk.Label(inner, text=tag, bg=COLORS["surface"], fg=COLORS["accent"], font=(FONT, 8, "bold")).pack(anchor="w")
            tk.Label(inner, text=title, bg=COLORS["surface"], fg=COLORS["text"], font=(FONT, 17, "bold")).pack(anchor="w", pady=(10, 6))
            tk.Label(inner, text=desc, bg=COLORS["surface"], fg=COLORS["muted"], font=(FONT, 10), justify="left").pack(anchor="w")
            self._button(inner, "この操作を使う  →", lambda k=key: self._show_page(k), kind="primary").pack(fill="x", pady=(28, 0))

        note = tk.Frame(wrap, bg=COLORS["success_soft"])
        note.pack(fill="x", pady=(26, 0))
        tk.Label(note, text="✓  CSVはPC内だけで処理されます。アップロードはありません。", bg=COLORS["success_soft"], fg=COLORS["success"], font=(FONT, 9, "bold"), anchor="w").pack(fill="x", padx=16, pady=13)

    # ---------- JOIN ----------
    def _build_join_page(self) -> None:
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
        self.join_left_summary = tk.StringVar(value="左のCSV：未選択")
        self.join_right_summary = tk.StringVar(value="右のCSV：未選択")
        self.join_warning = tk.StringVar(value="")
        self.join_result_summary = tk.StringVar(value="CSVを選んで条件を決めると、ここに確認結果が出ます。")
        self.join_metric_match = tk.StringVar(value="—")
        self.join_metric_left = tk.StringVar(value="—")
        self.join_metric_right = tk.StringVar(value="—")
        self.join_metric_dups = tk.StringVar(value="—")
        self.join_metric_rows = tk.StringVar(value="—")
        self.join_details_open = False

        page = self._page("join")
        body = page.body  # type: ignore[attr-defined]
        self._workflow_header(body, "JOIN", "CSVをまとめる", "専門用語を知らなくても、2つのまとめ方から選べます。")

        # Step 1
        card1 = self._card(body)
        card1.pack(fill="x")
        c1 = card1.inner  # type: ignore[attr-defined]
        self._section_title(c1, "STEP 1", "まとめ方を選んでCSVを指定")
        mode_row = tk.Frame(c1, bg=COLORS["surface"])
        mode_row.pack(fill="x")
        mode_row.grid_columnconfigure(0, weight=1)
        mode_row.grid_columnconfigure(1, weight=1)
        self.join_key_mode_button = self._button(mode_row, "IDなど共通の列で情報をつなぐ", lambda: self._set_join_mode("key"), kind="primary")
        self.join_key_mode_button.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.join_vertical_mode_button = self._button(mode_row, "行を下につなげる", lambda: self._set_join_mode("vertical"))
        self.join_vertical_mode_button.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        self.join_key_files = tk.Frame(c1, bg=COLORS["surface"])
        self.join_key_files.pack(fill="x", pady=(16, 0))
        self.join_key_files.grid_columnconfigure(0, weight=1)
        self.join_key_files.grid_columnconfigure(1, weight=1)
        left = tk.Frame(self.join_key_files, bg=COLORS["surface"])
        left.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._file_summary(left, self.join_left_summary).pack(fill="x")
        self.join_left_file_button = self._button(left, "左のCSVを選ぶ", lambda: self._load_join_side("left"))
        self.join_left_file_button.pack(fill="x", pady=(7, 0))
        right = tk.Frame(self.join_key_files, bg=COLORS["surface"])
        right.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        self._file_summary(right, self.join_right_summary).pack(fill="x")
        self.join_right_file_button = self._button(right, "右のCSVを選ぶ", lambda: self._load_join_side("right"))
        self.join_right_file_button.pack(fill="x", pady=(7, 0))

        self.join_stack_files = tk.Frame(c1, bg=COLORS["surface"])
        self.join_stack_summary = tk.StringVar(value="CSV：未選択")
        self._file_summary(self.join_stack_files, self.join_stack_summary).pack(fill="x")
        self.join_stack_button = self._button(self.join_stack_files, "2つ以上のCSVをまとめて選ぶ", self._load_vertical_files)
        self.join_stack_button.pack(fill="x", pady=(7, 0))

        # Step 2
        card2 = self._card(body)
        card2.pack(fill="x", pady=(14, 0))
        c2 = card2.inner  # type: ignore[attr-defined]
        self._section_title(c2, "STEP 2", "条件を決める")

        self.join_key_settings = tk.Frame(c2, bg=COLORS["surface"])
        self.join_key_settings.pack(fill="x")
        keys = tk.Frame(self.join_key_settings, bg=COLORS["surface"])
        keys.pack(fill="x")
        keys.grid_columnconfigure(0, weight=1)
        keys.grid_columnconfigure(1, weight=1)
        lk = tk.Frame(keys, bg=COLORS["surface"]); lk.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        rk = tk.Frame(keys, bg=COLORS["surface"]); rk.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        self.join_left_combo = self._combo(lk, "左のCSVで共通にする列", self.join_left_key); self.join_left_combo.master.pack(fill="x")
        self.join_right_combo = self._combo(rk, "右のCSVで共通にする列", self.join_right_key); self.join_right_combo.master.pack(fill="x")
        self.join_left_combo.bind("<<ComboboxSelected>>", lambda _e: self._invalidate_join_result())
        self.join_right_combo.bind("<<ComboboxSelected>>", lambda _e: self._invalidate_join_result())

        tk.Label(self.join_key_settings, text="どの行を残しますか？", bg=COLORS["surface"], fg=COLORS["muted"], font=(FONT, 8, "bold")).pack(anchor="w", pady=(14, 6))
        keep = tk.Frame(self.join_key_settings, bg=COLORS["surface"]); keep.pack(fill="x")
        keep.grid_columnconfigure(0, weight=1); keep.grid_columnconfigure(1, weight=1)
        self.join_left_type_button = self._button(keep, "左のCSVをすべて残す", lambda: self._set_join_type("left"), kind="primary")
        self.join_left_type_button.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.join_inner_type_button = self._button(keep, "両方にある行だけ残す", lambda: self._set_join_type("inner"))
        self.join_inner_type_button.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        self.join_type_hint = tk.Label(self.join_key_settings, text="左のCSVを基準に情報を追加します（LEFT JOIN）", bg=COLORS["surface"], fg=COLORS["faint"], font=(FONT, 8))
        self.join_type_hint.pack(anchor="w", pady=(6, 0))

        self.join_vertical_settings = tk.Frame(c2, bg=COLORS["success_soft"])
        tk.Label(self.join_vertical_settings, text="列名を基準にそろえて、選んだCSVを上から順に連結します。足りない列は空欄で補います。",
                 bg=COLORS["success_soft"], fg=COLORS["success"], font=(FONT, 9), anchor="w", wraplength=900).pack(fill="x", padx=14, pady=12)

        self.join_preview_button = self._button(c2, "結果を確認する", self._run_join_preview, kind="primary")
        self.join_preview_button.pack(fill="x", pady=(16, 0))

        # Step 3
        card3 = self._card(body)
        self.join_step3_card = card3
        card3.pack(fill="both", expand=True, pady=(14, 18))
        c3 = card3.inner  # type: ignore[attr-defined]
        self._section_title(c3, "STEP 3", "確認して保存", "問題がなければそのままCSVを書き出します。")
        summary_box = tk.Frame(c3, bg=COLORS["accent_soft"])
        summary_box.pack(fill="x")
        tk.Label(summary_box, textvariable=self.join_result_summary, bg=COLORS["accent_soft"], fg=COLORS["text"], font=(FONT, 12, "bold"), anchor="w", justify="left", wraplength=950).pack(fill="x", padx=14, pady=(12, 5))
        tk.Label(summary_box, textvariable=self.join_warning, bg=COLORS["accent_soft"], fg=COLORS["warning"], font=(FONT, 8), anchor="w", justify="left", wraplength=950).pack(fill="x", padx=14, pady=(0, 12))
        self.join_detail_button = self._button(c3, "詳細を見る", self._toggle_join_details, kind="ghost")
        self.join_detail_button.pack(anchor="w", pady=(8, 0))
        self.join_details_slot = tk.Frame(c3, bg=COLORS["surface"])
        self.join_details_slot.pack(fill="x")
        self.join_details_frame = tk.Frame(self.join_details_slot, bg=COLORS["surface_alt"])
        metrics = [("一致するキー", self.join_metric_match), ("左だけ", self.join_metric_left), ("右だけ", self.join_metric_right), ("重複 左 / 右", self.join_metric_dups), ("出力行数", self.join_metric_rows)]
        for i, (label, var) in enumerate(metrics):
            self.join_details_frame.grid_columnconfigure(i, weight=1)
            box = tk.Frame(self.join_details_frame, bg=COLORS["surface_alt"])
            box.grid(row=0, column=i, sticky="ew", padx=8, pady=10)
            tk.Label(box, text=label, bg=COLORS["surface_alt"], fg=COLORS["muted"], font=(FONT, 8, "bold")).pack(anchor="w")
            tk.Label(box, textvariable=var, bg=COLORS["surface_alt"], fg=COLORS["text"], font=(FONT, 14, "bold")).pack(anchor="w")

        tk.Label(c3, text="プレビュー（先頭100行）", bg=COLORS["surface"], fg=COLORS["muted"], font=(FONT, 8, "bold")).pack(anchor="w", pady=(14, 6))
        self.join_preview = PreviewTable(c3, height=7)
        self.join_preview.pack(fill="both", expand=True)

        export = tk.Frame(c3, bg=COLORS["surface"]); export.pack(fill="x", pady=(12, 0))
        export.grid_columnconfigure(0, weight=1)
        out = tk.Frame(export, bg=COLORS["surface"]); out.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._entry(out, "保存先", self.join_output).master.pack(fill="x")
        self._button(export, "参照", self._choose_join_output).grid(row=0, column=1, sticky="s")
        enc = tk.Frame(export, bg=COLORS["surface"]); enc.grid(row=1, column=0, sticky="w", pady=(8, 0))
        self._combo(enc, "文字コード", self.join_encoding, list(EXPORT_ENCODINGS), width=14).master.pack(fill="x")
        self.join_export_button = self._button(export, "CSVを保存", self._export_join, kind="primary")
        self.join_export_button.grid(row=1, column=1, sticky="se", pady=(8, 0))
        self._set_join_mode("key")
        self._set_join_type("left")

    def _refresh_join_file_summaries(self) -> None:
        if self.join_left:
            self.join_left_summary.set(f"左のCSV\n{self.join_left.path.name}\n{self.join_left.rows:,}行 · {len(self.join_left.columns)}列 · {self.join_left.encoding}")
        else:
            self.join_left_summary.set("左のCSV：未選択")
        if self.join_right:
            self.join_right_summary.set(f"右のCSV\n{self.join_right.path.name}\n{self.join_right.rows:,}行 · {len(self.join_right.columns)}列 · {self.join_right.encoding}")
        else:
            self.join_right_summary.set("右のCSV：未選択")
        if self.join_files:
            names = "\n".join(f"• {d.path.name}  ({d.rows:,}行)" for d in self.join_files[:5])
            more = f"\nほか {len(self.join_files)-5} ファイル" if len(self.join_files) > 5 else ""
            self.join_stack_summary.set(f"{len(self.join_files)}ファイル選択済み\n{names}{more}")
        else:
            self.join_stack_summary.set("CSV：未選択")

    def _set_join_mode(self, mode: str) -> None:
        changed = self.join_mode.get() != mode
        self.join_mode.set(mode)
        is_key = mode == "key"
        self.join_key_mode_button.configure(bg=COLORS["accent"] if is_key else COLORS["surface_alt"], fg="#FFFFFF" if is_key else COLORS["text"])
        self.join_vertical_mode_button.configure(bg=COLORS["surface_alt"] if is_key else COLORS["accent"], fg=COLORS["text"] if is_key else "#FFFFFF")
        if is_key:
            self.join_stack_files.pack_forget(); self.join_vertical_settings.pack_forget()
            self.join_key_files.pack(fill="x", pady=(16, 0)); self.join_key_settings.pack(fill="x")
        else:
            self.join_key_files.pack_forget(); self.join_key_settings.pack_forget()
            self.join_stack_files.pack(fill="x", pady=(16, 0)); self.join_vertical_settings.pack(fill="x")
        if changed:
            self._invalidate_join_result("まとめ方が変わりました。もう一度結果を確認してください。")

    def _set_join_type(self, join_type: str) -> None:
        changed = self.join_type.get() != join_type
        self.join_type.set(join_type)
        is_left = join_type == "left"
        self.join_left_type_button.configure(bg=COLORS["accent"] if is_left else COLORS["surface_alt"], fg="#FFFFFF" if is_left else COLORS["text"])
        self.join_inner_type_button.configure(bg=COLORS["surface_alt"] if is_left else COLORS["accent"], fg=COLORS["text"] if is_left else "#FFFFFF")
        self.join_type_hint.configure(text="左のCSVを基準に情報を追加します（LEFT JOIN）" if is_left else "両方に存在するキーだけ残します（INNER JOIN）")
        if changed:
            self._invalidate_join_result("残す行の条件が変わりました。もう一度結果を確認してください。")

    def _invalidate_join_result(self, message="条件が変わりました。もう一度結果を確認してください。") -> None:
        self.join_result = None
        self.join_preview.clear()
        self._reset_join_metrics()
        self.join_result_summary.set("条件を確認して「結果を確認する」を押してください。")
        self.join_warning.set(message)

    def _reset_join_metrics(self) -> None:
        for var in (self.join_metric_match, self.join_metric_left, self.join_metric_right, self.join_metric_dups, self.join_metric_rows):
            var.set("—")

    def _load_join_side(self, side: str) -> None:
        path = self._browse_csv()
        if not path:
            return
        try:
            doc = read_csv_flexible(Path(path))
            if side == "left":
                self.join_left = doc; self.join_left_combo["values"] = doc.columns; self.join_left_key.set(doc.columns[0] if doc.columns else "")
            else:
                self.join_right = doc; self.join_right_combo["values"] = doc.columns; self.join_right_key.set(doc.columns[0] if doc.columns else "")
            self._refresh_join_file_summaries(); self._invalidate_join_result("")
            self.status_var.set(f"{doc.path.name} を読み込みました。")
        except Exception as exc:
            self._show_error(exc)

    def _load_vertical_files(self) -> None:
        paths = filedialog.askopenfilenames(title="つなげるCSVを選択", filetypes=[("CSV files", "*.csv")])
        if not paths:
            return
        try:
            self.join_files = [read_csv_flexible(Path(p)) for p in paths]
            if len(self.join_files) < 2:
                raise CsvToolError("2つ以上のCSVを選択してください。")
            self._refresh_join_file_summaries(); self._invalidate_join_result("")
            self.status_var.set(f"{len(self.join_files)}ファイルを読み込みました。")
        except Exception as exc:
            self._show_error(exc)

    def _run_join_preview(self) -> None:
        try:
            if self.join_mode.get() == "vertical":
                self.join_result = vertical_join([d.dataframe for d in self.join_files])
                self.join_metric_match.set("—"); self.join_metric_left.set(str(len(self.join_files))); self.join_metric_right.set("—"); self.join_metric_dups.set("—")
                self.join_metric_rows.set(f"{len(self.join_result):,}")
                self.join_result_summary.set(f"{len(self.join_files)}ファイルをまとめて、{len(self.join_result):,}行のCSVになります。")
                self.join_warning.set(f"出力列は {len(self.join_result.columns)}列。不足する列は空欄で補います。")
                base = self.join_files[0].path.parent if self.join_files else Path.cwd()
                self.join_output.set(str(base / "merged.csv"))
            else:
                if not self.join_left or not self.join_right:
                    raise CsvToolError("左と右のCSVを選択してください。")
                self.join_result, diag = key_join(self.join_left.dataframe, self.join_right.dataframe, self.join_left_key.get(), self.join_right_key.get(), self.join_type.get())
                self.join_metric_match.set(f"{diag.matched_keys:,}"); self.join_metric_left.set(f"{diag.left_only_keys:,}"); self.join_metric_right.set(f"{diag.right_only_keys:,}")
                self.join_metric_dups.set(f"{diag.left_duplicate_keys:,} / {diag.right_duplicate_keys:,}"); self.join_metric_rows.set(f"{diag.result_rows:,}")
                self.join_result_summary.set(f"{diag.result_rows:,}行のCSVになります。一致するキーは {diag.matched_keys:,}件です。")
                if diag.warnings:
                    self.join_warning.set("⚠ " + " / ".join(diag.warnings))
                elif diag.left_only_keys or diag.right_only_keys:
                    self.join_warning.set(f"左だけ {diag.left_only_keys:,}件、右だけ {diag.right_only_keys:,}件あります。必要なら詳細を確認してください。")
                else:
                    self.join_warning.set("✓ キー構成に問題は見つかりませんでした。")
                self.join_output.set(str(self.join_left.path.parent / "merged.csv"))
            self.join_preview.show_dataframe(self.join_result)
            self.status_var.set("確認結果を更新しました。内容を見て保存してください。")
            self._scroll_to_widget("join", self.join_step3_card)
        except Exception as exc:
            self._show_error(exc)

    def _toggle_join_details(self) -> None:
        self.join_details_open = not self.join_details_open
        if self.join_details_open:
            self.join_details_frame.pack(fill="x", pady=(6, 0))
            self.join_detail_button.configure(text="詳細を閉じる")
        else:
            self.join_details_frame.pack_forget()
            self.join_detail_button.configure(text="詳細を見る")

    def _choose_join_output(self) -> None:
        path = self._save_csv_path("merged.csv")
        if path:
            self.join_output.set(path)

    def _export_join(self) -> None:
        try:
            if self.join_result is None:
                raise CsvToolError("先に「結果を確認する」を押してください。")
            if not self.join_output.get().strip():
                raise CsvToolError("保存先を指定してください。")
            export_csv(self.join_result, Path(self.join_output.get()), self.join_encoding.get())
            self.status_var.set(f"保存しました: {self.join_output.get()}")
            messagebox.showinfo(APP_TITLE, "CSVを保存しました。")
        except Exception as exc:
            self._show_error(exc)

    # ---------- SPLIT ----------
    def _build_split_page(self) -> None:
        self.split_doc: CsvDocument | None = None
        self.split_key = tk.StringVar()
        self.split_include_blank = tk.BooleanVar(value=False)
        self.split_output_dir = tk.StringVar()
        self.split_encoding = tk.StringVar(value="UTF-8 BOM")
        self.split_summary = tk.StringVar(value="CSV：未選択")
        self.split_result_summary = tk.StringVar(value="CSVを選んで分ける列を決めると、ここに出力予定が表示されます。")
        self.split_file_count = tk.StringVar(value="—")
        self.split_row_count = tk.StringVar(value="—")
        self.split_blank_count = tk.StringVar(value="—")
        self.split_analysis_valid = False

        page = self._page("split"); body = page.body  # type: ignore[attr-defined]
        self._workflow_header(body, "SPLIT", "CSVを分ける", "選んだ列の値ごとに、自動で複数のCSVへ分けます。")

        card1 = self._card(body); card1.pack(fill="x")
        c1 = card1.inner  # type: ignore[attr-defined]
        self._section_title(c1, "STEP 1", "分けたいCSVを選ぶ")
        self._file_summary(c1, self.split_summary).pack(fill="x")
        self._button(c1, "CSVを選ぶ", self._load_split_file).pack(fill="x", pady=(8, 0))

        card2 = self._card(body); card2.pack(fill="x", pady=(14, 0))
        c2 = card2.inner  # type: ignore[attr-defined]
        self._section_title(c2, "STEP 2", "どの列で分けますか？")
        self.split_key_combo = self._combo(c2, "分ける基準の列", self.split_key); self.split_key_combo.master.pack(fill="x")
        self.split_key_combo.bind("<<ComboboxSelected>>", lambda _e: self._invalidate_split_analysis())
        blank = tk.Frame(c2, bg=COLORS["surface"]); blank.pack(fill="x", pady=(14, 0))
        tk.Label(blank, text="空欄のデータ", bg=COLORS["surface"], fg=COLORS["muted"], font=(FONT, 8, "bold")).pack(anchor="w", pady=(0, 6))
        tk.Radiobutton(blank, text="出力しない", variable=self.split_include_blank, value=False, command=self._invalidate_split_analysis,
                       bg=COLORS["surface"], fg=COLORS["text"], activebackground=COLORS["surface"], font=(FONT, 9), selectcolor=COLORS["surface"]).pack(side="left")
        tk.Radiobutton(blank, text="blank.csv にまとめる", variable=self.split_include_blank, value=True, command=self._invalidate_split_analysis,
                       bg=COLORS["surface"], fg=COLORS["text"], activebackground=COLORS["surface"], font=(FONT, 9), selectcolor=COLORS["surface"]).pack(side="left", padx=(18, 0))
        self._button(c2, "出力内容を確認する", self._preview_split_counts, kind="primary").pack(fill="x", pady=(16, 0))

        card3 = self._card(body); self.split_step3_card = card3
        card3.pack(fill="both", expand=True, pady=(14, 18))
        c3 = card3.inner  # type: ignore[attr-defined]
        self._section_title(c3, "STEP 3", "確認して分割")
        summary = tk.Frame(c3, bg=COLORS["accent_soft"]); summary.pack(fill="x")
        tk.Label(summary, textvariable=self.split_result_summary, bg=COLORS["accent_soft"], fg=COLORS["text"], font=(FONT, 12, "bold"), anchor="w").pack(fill="x", padx=14, pady=12)
        self.split_preview = PreviewTable(c3, height=7); self.split_preview.pack(fill="both", expand=True, pady=(12, 0))
        out = tk.Frame(c3, bg=COLORS["surface"]); out.pack(fill="x", pady=(12, 0)); out.grid_columnconfigure(0, weight=1)
        entry = tk.Frame(out, bg=COLORS["surface"]); entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._entry(entry, "保存先フォルダ", self.split_output_dir).master.pack(fill="x")
        self._button(out, "参照", self._choose_split_output).grid(row=0, column=1, sticky="s")
        enc = tk.Frame(out, bg=COLORS["surface"]); enc.grid(row=1, column=0, sticky="w", pady=(8, 0))
        self._combo(enc, "文字コード", self.split_encoding, list(EXPORT_ENCODINGS), width=14).master.pack(fill="x")
        self.split_export_button = self._button(out, "CSVを分割して保存", self._export_split, kind="primary")
        self.split_export_button.grid(row=1, column=1, sticky="se", pady=(8, 0))

    def _load_split_file(self) -> None:
        path = self._browse_csv()
        if not path:
            return
        try:
            self.split_doc = read_csv_flexible(Path(path))
            self.split_key_combo["values"] = self.split_doc.columns
            self.split_key.set(self.split_doc.columns[0] if self.split_doc.columns else "")
            self.split_output_dir.set(str(self.split_doc.path.parent / f"{self.split_doc.path.stem}_split"))
            self.split_summary.set(f"{self.split_doc.path.name}\n{self.split_doc.rows:,}行 · {len(self.split_doc.columns)}列 · {self.split_doc.encoding}")
            self.split_row_count.set(str(self.split_doc.rows)); self._invalidate_split_analysis()
            self.status_var.set(f"{self.split_doc.path.name} を読み込みました。")
        except Exception as exc:
            self._show_error(exc)

    def _invalidate_split_analysis(self) -> None:
        self.split_analysis_valid = False
        self.split_preview.clear()
        self.split_file_count.set("—"); self.split_blank_count.set("—")
        self.split_result_summary.set("条件を確認して「出力内容を確認する」を押してください。")

    def _preview_split_counts(self) -> None:
        try:
            if not self.split_doc:
                raise CsvToolError("先にCSVを選択してください。")
            if not self.split_key.get():
                raise CsvToolError("分ける列を選択してください。")
            plan = split_plan(self.split_doc.dataframe, self.split_key.get(), include_blank=self.split_include_blank.get())
            raw = self.split_doc.dataframe[self.split_key.get()].astype(str).str.strip()
            blank_count = int((raw == "").sum())
            self.split_file_count.set(str(len(plan))); self.split_blank_count.set(str(blank_count)); self.split_row_count.set(str(self.split_doc.rows))
            self.split_result_summary.set(f"{len(plan)}ファイルに分かれます。元データは {self.split_doc.rows:,}行です。")
            self.split_preview.show_dataframe(plan)
            self.split_analysis_valid = True
            self.status_var.set("出力予定を確認しました。")
            self._scroll_to_widget("split", self.split_step3_card)
        except Exception as exc:
            self._show_error(exc)

    def _choose_split_output(self) -> None:
        path = filedialog.askdirectory(title="保存先フォルダを選択")
        if path:
            self.split_output_dir.set(path)

    def _export_split(self) -> None:
        try:
            if not self.split_doc:
                raise CsvToolError("CSVを選択してください。")
            if not self.split_analysis_valid:
                raise CsvToolError("先に「出力内容を確認する」を押してください。")
            if not self.split_output_dir.get().strip():
                raise CsvToolError("保存先フォルダを指定してください。")
            groups = split_dataframe(self.split_doc.dataframe, self.split_key.get(), include_blank=self.split_include_blank.get())
            paths = export_split_groups(groups, Path(self.split_output_dir.get()), self.split_encoding.get())
            self.status_var.set(f"{len(paths)}ファイルを保存しました。")
            messagebox.showinfo(APP_TITLE, f"{len(paths)}ファイルを保存しました。")
        except Exception as exc:
            self._show_error(exc)

    # ---------- TRANSFORM ----------
    def _build_transform_page(self) -> None:
        self.transform_doc: CsvDocument | None = None
        self.transform_steps: list[TransformStep] = []
        self.transform_result: pd.DataFrame | None = None
        self.transform_summary = tk.StringVar(value="CSV：未選択")
        self.transform_operation = tk.StringVar(value="rename")
        self.transform_column = tk.StringVar()
        self.transform_value1 = tk.StringVar()
        self.transform_value2 = tk.StringVar()
        self.transform_output = tk.StringVar()
        self.transform_encoding = tk.StringVar(value="UTF-8 BOM")
        self.transform_step_count = tk.StringVar(value="0")
        self.transform_output_cols = tk.StringVar(value="—")
        self.transform_v1_caption = tk.StringVar(value="新しい列名")
        self.transform_v2_caption = tk.StringVar(value="値")
        self.transform_hint_text = tk.StringVar(value="")
        self.transform_editor_open = False
        self.transform_preview_after = True

        page = self._page("transform"); body = page.body  # type: ignore[attr-defined]
        self._workflow_header(body, "TRANSFORM", "CSVを整える", "変換を1つずつ追加。結果はその場で確認できます。")

        card1 = self._card(body); card1.pack(fill="x")
        c1 = card1.inner  # type: ignore[attr-defined]
        self._section_title(c1, "STEP 1", "整えたいCSVを選ぶ")
        self._file_summary(c1, self.transform_summary).pack(fill="x")
        self._button(c1, "CSVを選ぶ", self._load_transform_file).pack(fill="x", pady=(8, 0))

        card2 = self._card(body); card2.pack(fill="x", pady=(14, 0))
        c2 = card2.inner  # type: ignore[attr-defined]
        self._section_title(c2, "STEP 2", "変換内容を追加する", "上から順番に適用されます。")
        self.transform_step_list = tk.Listbox(c2, height=5, bg="#FFFFFF", fg=COLORS["text"], selectbackground=COLORS["accent_soft"],
                                              selectforeground=COLORS["text"], relief="flat", highlightthickness=1,
                                              highlightbackground=COLORS["border"], font=(FONT, 9))
        self.transform_step_list.pack(fill="x")
        step_actions = tk.Frame(c2, bg=COLORS["surface"]); step_actions.pack(fill="x", pady=(8, 0))
        self._button(step_actions, "＋ 変換を追加", self._show_transform_editor, kind="primary").pack(side="left")
        self._button(step_actions, "選択した変換を削除", self._remove_transform_step, kind="ghost").pack(side="left", padx=(8, 0))
        self._button(step_actions, "すべて消す", self._clear_transform_steps, kind="ghost").pack(side="right")
        self.transform_continue_button = self._button(c2, "変更後を確認して保存へ  ↓", lambda: self._scroll_to_widget("transform", self.transform_step3_card))
        self.transform_continue_button.pack(fill="x", pady=(10, 0))

        self.transform_editor = tk.Frame(c2, bg=COLORS["surface_alt"])
        editor_inner = tk.Frame(self.transform_editor, bg=COLORS["surface_alt"]); editor_inner.pack(fill="x", padx=14, pady=14)
        tk.Label(editor_inner, text="追加する変換", bg=COLORS["surface_alt"], fg=COLORS["text"], font=(FONT, 10, "bold")).pack(anchor="w")
        self.transform_op_labels = {
            "列名を変更": "rename", "列を削除": "delete", "固定値の列を追加": "add_fixed", "前後の空白を削除": "trim",
            "文字列を置換": "replace", "日付形式を変更": "date", "0埋め": "zero_pad", "列の位置を変更": "move",
        }
        self.transform_op_display = tk.StringVar(value="列名を変更")
        opwrap = tk.Frame(editor_inner, bg=COLORS["surface_alt"]); opwrap.pack(fill="x", pady=(10, 0))
        tk.Label(opwrap, text="何をしますか？", bg=COLORS["surface_alt"], fg=COLORS["muted"], font=(FONT, 8, "bold")).pack(anchor="w", pady=(0, 5))
        self.transform_op_combo = ttk.Combobox(opwrap, textvariable=self.transform_op_display, values=list(self.transform_op_labels), state="readonly", style="Modern.TCombobox")
        self.transform_op_combo.pack(fill="x"); self.transform_op_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_transform_operation())

        colwrap = tk.Frame(editor_inner, bg=COLORS["surface_alt"]); colwrap.pack(fill="x", pady=(10, 0))
        tk.Label(colwrap, text="対象の列", bg=COLORS["surface_alt"], fg=COLORS["muted"], font=(FONT, 8, "bold")).pack(anchor="w", pady=(0, 5))
        self.transform_col_combo = ttk.Combobox(colwrap, textvariable=self.transform_column, state="readonly", style="Modern.TCombobox")
        self.transform_col_combo.pack(fill="x")

        values = tk.Frame(editor_inner, bg=COLORS["surface_alt"]); values.pack(fill="x", pady=(10, 0)); values.grid_columnconfigure(0, weight=1); values.grid_columnconfigure(1, weight=1)
        self.transform_v1_wrap = tk.Frame(values, bg=COLORS["surface_alt"]); self.transform_v1_wrap.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        tk.Label(self.transform_v1_wrap, textvariable=self.transform_v1_caption, bg=COLORS["surface_alt"], fg=COLORS["muted"], font=(FONT, 8, "bold")).pack(anchor="w", pady=(0, 5))
        self.transform_v1_entry = tk.Entry(self.transform_v1_wrap, textvariable=self.transform_value1, bg="#FFFFFF", fg=COLORS["text"], relief="flat", highlightthickness=1, highlightbackground=COLORS["border"], font=(FONT, 10)); self.transform_v1_entry.pack(fill="x", ipady=8)
        self.transform_v2_wrap = tk.Frame(values, bg=COLORS["surface_alt"]); self.transform_v2_wrap.grid(row=0, column=1, sticky="ew", padx=(5, 0))
        tk.Label(self.transform_v2_wrap, textvariable=self.transform_v2_caption, bg=COLORS["surface_alt"], fg=COLORS["muted"], font=(FONT, 8, "bold")).pack(anchor="w", pady=(0, 5))
        self.transform_v2_entry = tk.Entry(self.transform_v2_wrap, textvariable=self.transform_value2, bg="#FFFFFF", fg=COLORS["text"], relief="flat", highlightthickness=1, highlightbackground=COLORS["border"], font=(FONT, 10)); self.transform_v2_entry.pack(fill="x", ipady=8)
        tk.Label(editor_inner, textvariable=self.transform_hint_text, bg=COLORS["surface_alt"], fg=COLORS["muted"], font=(FONT, 8), anchor="w").pack(fill="x", pady=(8, 0))
        buttons = tk.Frame(editor_inner, bg=COLORS["surface_alt"]); buttons.pack(fill="x", pady=(12, 0))
        self._button(buttons, "追加", self._add_transform_step, kind="primary").pack(side="right")
        self._button(buttons, "閉じる", self._hide_transform_editor, kind="ghost").pack(side="right", padx=(0, 8))
        self._sync_transform_operation()

        card3 = self._card(body); self.transform_step3_card = card3
        card3.pack(fill="both", expand=True, pady=(14, 18))
        c3 = card3.inner  # type: ignore[attr-defined]
        self._section_title(c3, "STEP 3", "確認して保存")
        top = tk.Frame(c3, bg=COLORS["surface"]); top.pack(fill="x", pady=(0, 6))
        tk.Label(top, textvariable=self.transform_step_count, bg=COLORS["accent_soft"], fg=COLORS["accent"], font=(FONT, 10, "bold"), padx=10, pady=5).pack(side="left")
        tk.Label(top, text="個の変換を適用中", bg=COLORS["surface"], fg=COLORS["muted"], font=(FONT, 9)).pack(side="left", padx=(6, 0))
        self.transform_before_button = self._button(top, "変更前", lambda: self._show_transform_preview(False), kind="ghost"); self.transform_before_button.pack(side="right")
        self.transform_after_button = self._button(top, "変更後", lambda: self._show_transform_preview(True), kind="primary"); self.transform_after_button.pack(side="right", padx=(0, 6))

        self.transform_preview_host = tk.Frame(c3, bg=COLORS["surface"]); self.transform_preview_host.pack(fill="both", expand=True)
        self.transform_before = PreviewTable(self.transform_preview_host, height=7)
        self.transform_after = PreviewTable(self.transform_preview_host, height=7)
        self.transform_after.pack(fill="both", expand=True)

        out = tk.Frame(c3, bg=COLORS["surface"]); out.pack(fill="x", pady=(12, 0)); out.grid_columnconfigure(0, weight=1)
        entry = tk.Frame(out, bg=COLORS["surface"]); entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._entry(entry, "保存先", self.transform_output).master.pack(fill="x")
        self._button(out, "参照", self._choose_transform_output).grid(row=0, column=1, sticky="s")
        enc = tk.Frame(out, bg=COLORS["surface"]); enc.grid(row=1, column=0, sticky="w", pady=(8, 0))
        self._combo(enc, "文字コード", self.transform_encoding, list(EXPORT_ENCODINGS), width=14).master.pack(fill="x")
        self.transform_export_button = self._button(out, "CSVを保存", self._export_transform, kind="primary")
        self.transform_export_button.grid(row=1, column=1, sticky="se", pady=(8, 0))

    def _show_transform_editor(self) -> None:
        if not self.transform_editor_open:
            self.transform_editor.pack(fill="x", pady=(12, 0))
            self.transform_editor_open = True

    def _hide_transform_editor(self) -> None:
        self.transform_editor.pack_forget(); self.transform_editor_open = False

    def _load_transform_file(self) -> None:
        path = self._browse_csv()
        if not path:
            return
        try:
            self.transform_doc = read_csv_flexible(Path(path)); self.transform_steps.clear(); self.transform_result = None
            self._refresh_transform_step_list(); self.transform_col_combo["values"] = self.transform_doc.columns
            self.transform_column.set(self.transform_doc.columns[0] if self.transform_doc.columns else "")
            self.transform_output.set(str(self.transform_doc.path.parent / f"{self.transform_doc.path.stem}_transformed.csv"))
            self.transform_summary.set(f"{self.transform_doc.path.name}\n{self.transform_doc.rows:,}行 · {len(self.transform_doc.columns)}列 · {self.transform_doc.encoding}")
            self.transform_before.show_dataframe(self.transform_doc.dataframe); self.transform_after.show_dataframe(self.transform_doc.dataframe)
            self.transform_output_cols.set(str(len(self.transform_doc.columns))); self._show_transform_preview(True)
            self.status_var.set(f"{self.transform_doc.path.name} を読み込みました。")
        except Exception as exc:
            self._show_error(exc)

    def _sync_transform_operation(self) -> None:
        if hasattr(self, "transform_op_display"):
            raw = self.transform_op_display.get()
            self.transform_operation.set(self.transform_op_labels.get(raw, self.transform_operation.get()))
        op = self.transform_operation.get()
        hints = {
            "rename": ("新しい列名", "", "例：氏名 → 生徒名"), "delete": ("", "", "選んだ列を削除します。"),
            "add_fixed": ("追加する列名", "固定する値", "例：確認状態 = 未確認"), "trim": ("", "", "前後の余分な空白だけを削除します。"),
            "replace": ("置換前", "置換後", "指定した文字列を置き換えます。"), "date": ("出力形式", "", "例：%Y-%m-%d。変換できない値は元のまま残します。"),
            "zero_pad": ("桁数", "", "例：6 → 123 を 000123 にします。"), "move": ("移動先", "", "「先頭」「末尾」または列番号を入力します。"),
        }
        c1, c2, hint = hints.get(op, ("値", "", "")); self.transform_v1_caption.set(c1); self.transform_v2_caption.set(c2); self.transform_hint_text.set(hint)
        self.transform_col_combo.configure(state="disabled" if op == "add_fixed" else "readonly")
        need1 = op in {"rename", "add_fixed", "replace", "date", "zero_pad", "move"}; need2 = op in {"add_fixed", "replace"}
        if need1:
            self.transform_v1_wrap.grid()
        else:
            self.transform_v1_wrap.grid_remove(); self.transform_value1.set("")
        if need2:
            self.transform_v2_wrap.grid()
        else:
            self.transform_v2_wrap.grid_remove(); self.transform_value2.set("")

    def _add_transform_step(self) -> None:
        try:
            if not self.transform_doc:
                raise CsvToolError("先にCSVを選択してください。")
            op = self.transform_operation.get(); column = self.transform_column.get()
            if op != "add_fixed" and not column:
                raise CsvToolError("対象の列を選択してください。")
            step = TransformStep(op, column, self.transform_value1.get(), self.transform_value2.get())
            preview = apply_transform_pipeline(self.transform_doc.dataframe, [*self.transform_steps, step])
            self.transform_steps.append(step); self.transform_result = preview; self._refresh_transform_step_list()
            self.transform_after.show_dataframe(preview); self.transform_output_cols.set(str(len(preview.columns)))
            self.transform_col_combo["values"] = list(preview.columns)
            if self.transform_column.get() not in preview.columns:
                self.transform_column.set(str(preview.columns[0]) if len(preview.columns) else "")
            self.transform_value1.set(""); self.transform_value2.set(""); self._show_transform_preview(True)
            self.status_var.set("変換を追加しました。変更後を確認してください。")
        except Exception as exc:
            self._show_error(exc)

    def _refresh_transform_step_list(self) -> None:
        self.transform_step_list.delete(0, "end")
        for i, step in enumerate(self.transform_steps, 1):
            self.transform_step_list.insert("end", f"{i}.  {step.describe()}")
        self.transform_step_count.set(str(len(self.transform_steps)))

    def _remove_transform_step(self) -> None:
        selection = self.transform_step_list.curselection()
        if not selection:
            return
        del self.transform_steps[selection[0]]; self._refresh_transform_step_list(); self._rebuild_transform_result()

    def _clear_transform_steps(self) -> None:
        self.transform_steps.clear(); self._refresh_transform_step_list(); self._rebuild_transform_result(); self.status_var.set("変換をすべて削除しました。")

    def _rebuild_transform_result(self) -> None:
        if not self.transform_doc:
            self.transform_result = None; self.transform_after.clear(); return
        preview = apply_transform_pipeline(self.transform_doc.dataframe, self.transform_steps)
        self.transform_result = preview if self.transform_steps else None
        self.transform_after.show_dataframe(preview); self.transform_col_combo["values"] = list(preview.columns)
        if self.transform_column.get() not in preview.columns:
            self.transform_column.set(str(preview.columns[0]) if len(preview.columns) else "")
        self.transform_output_cols.set(str(len(preview.columns))); self._show_transform_preview(True)

    def _apply_transform_preview(self) -> None:
        if not self.transform_doc:
            self._show_error(CsvToolError("CSVを選択してください。")); return
        if not self.transform_steps:
            self._show_error(CsvToolError("変換を1つ以上追加してください。")); return
        self.transform_result = apply_transform_pipeline(self.transform_doc.dataframe, self.transform_steps)
        self.transform_after.show_dataframe(self.transform_result); self._show_transform_preview(True)

    def _show_transform_preview(self, after: bool) -> None:
        self.transform_before.pack_forget(); self.transform_after.pack_forget()
        if after:
            self.transform_after.pack(fill="both", expand=True)
            self.transform_after_button.configure(bg=COLORS["accent"], fg="#FFFFFF"); self.transform_before_button.configure(bg=COLORS["surface"], fg=COLORS["muted"])
        else:
            self.transform_before.pack(fill="both", expand=True)
            self.transform_before_button.configure(bg=COLORS["accent"], fg="#FFFFFF"); self.transform_after_button.configure(bg=COLORS["surface"], fg=COLORS["muted"])
        self.transform_preview_after = after

    def _choose_transform_output(self) -> None:
        path = self._save_csv_path("transformed.csv")
        if path:
            self.transform_output.set(path)

    def _export_transform(self) -> None:
        try:
            if self.transform_result is None:
                raise CsvToolError("変換を1つ以上追加して、変更後を確認してください。")
            if not self.transform_output.get().strip():
                raise CsvToolError("保存先を指定してください。")
            export_csv(self.transform_result, Path(self.transform_output.get()), self.transform_encoding.get())
            self.status_var.set(f"保存しました: {self.transform_output.get()}")
            messagebox.showinfo(APP_TITLE, "CSVを保存しました。")
        except Exception as exc:
            self._show_error(exc)

    def _show_error(self, exc: Exception) -> None:
        self.status_var.set("処理を止めました。内容を確認してください。")
        messagebox.showerror(APP_TITLE, str(exc) if isinstance(exc, CsvToolError) else f"想定外エラー:\n{exc}")


def run_app() -> None:
    CSVJoinerApp().mainloop()
