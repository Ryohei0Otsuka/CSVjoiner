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
    split_plan,
    vertical_join,
)


APP_TITLE = "CSVjoiner v2"
FONT = "Yu Gothic UI"
COLORS = {
    "app": "#0B1020",
    "sidebar": "#0F1629",
    "surface": "#141D31",
    "surface_alt": "#19243A",
    "surface_soft": "#10182A",
    "border": "#263550",
    "border_focus": "#4C8DFF",
    "text": "#F4F7FC",
    "muted": "#91A0B8",
    "faint": "#63718A",
    "accent": "#4C8DFF",
    "accent_hover": "#397BE8",
    "accent_soft": "#1B315A",
    "success": "#4FD1A1",
    "warning": "#F2C66D",
    "danger": "#F27A8A",
    "row_alt": "#111A2C",
}


class PreviewTable(tk.Frame):
    def __init__(self, parent: tk.Widget, height: int = 12) -> None:
        super().__init__(parent, bg=COLORS["surface"])
        self.tree = ttk.Treeview(self, show="headings", height=height, style="Data.Treeview")
        ybar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview, style="Dark.Vertical.TScrollbar")
        xbar = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview, style="Dark.Horizontal.TScrollbar")
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
            self.tree.column(col, width=140, minwidth=90, stretch=True, anchor="w")
        for index, row in enumerate(frame.itertuples(index=False, name=None)):
            tag = "even" if index % 2 == 0 else "odd"
            self.tree.insert("", "end", values=[format_preview_value(v) for v in row], tags=(tag,))
        self.tree.tag_configure("even", background=COLORS["surface_soft"])
        self.tree.tag_configure("odd", background=COLORS["row_alt"])

    def clear(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []


class CSVJoinerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_TITLE} {__version__}")
        self.geometry("1360x860")
        self.minsize(1120, 720)
        self.configure(bg=COLORS["app"])

        self._setup_style()
        self.status_var = tk.StringVar(value="Ready — CSVを選択して処理を開始してください。")
        self.current_page = "join"
        self.pages: dict[str, tk.Frame] = {}
        self.page_canvases: dict[str, tk.Canvas] = {}
        self.nav_buttons: dict[str, tk.Button] = {}

        self._build_shell()
        self._build_join_page()
        self._build_split_page()
        self._build_transform_page()
        self._show_page("join")
        self.bind_all("<MouseWheel>", self._on_mousewheel, add="+")

    # ---------- shell / shared UI ----------
    def _setup_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Modern.TCombobox",
            fieldbackground=COLORS["surface_soft"],
            background=COLORS["surface_soft"],
            foreground=COLORS["text"],
            arrowcolor=COLORS["muted"],
            bordercolor=COLORS["border"],
            lightcolor=COLORS["border"],
            darkcolor=COLORS["border"],
            padding=7,
            font=(FONT, 10),
        )
        style.map(
            "Modern.TCombobox",
            fieldbackground=[("readonly", COLORS["surface_soft"])],
            foreground=[("readonly", COLORS["text"])],
            selectbackground=[("readonly", COLORS["surface_soft"])],
            selectforeground=[("readonly", COLORS["text"])],
        )
        style.configure(
            "Data.Treeview",
            background=COLORS["surface_soft"],
            fieldbackground=COLORS["surface_soft"],
            foreground=COLORS["text"],
            rowheight=30,
            borderwidth=0,
            relief="flat",
            font=(FONT, 9),
        )
        style.configure(
            "Data.Treeview.Heading",
            background=COLORS["surface_alt"],
            foreground=COLORS["muted"],
            borderwidth=0,
            relief="flat",
            font=(FONT, 9, "bold"),
            padding=(8, 8),
        )
        style.map(
            "Data.Treeview",
            background=[("selected", COLORS["accent_soft"])],
            foreground=[("selected", COLORS["text"])],
        )
        style.map(
            "Data.Treeview.Heading",
            background=[("active", COLORS["surface_alt"])],
        )
        for scrollbar_style in ("Dark.Vertical.TScrollbar", "Dark.Horizontal.TScrollbar"):
            style.configure(
                scrollbar_style,
                background=COLORS["surface_alt"],
                troughcolor=COLORS["surface_soft"],
                bordercolor=COLORS["surface_soft"],
                arrowcolor=COLORS["muted"],
                lightcolor=COLORS["surface_alt"],
                darkcolor=COLORS["surface_alt"],
                relief="flat",
            )
            style.map(
                scrollbar_style,
                background=[("active", COLORS["accent_soft"]), ("pressed", COLORS["accent"])],
            )

    def _build_shell(self) -> None:
        self.sidebar = tk.Frame(self, bg=COLORS["sidebar"], width=228)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        brand = tk.Frame(self.sidebar, bg=COLORS["sidebar"])
        brand.pack(fill="x", padx=20, pady=(24, 26))
        tk.Label(
            brand,
            text="CSVjoiner",
            bg=COLORS["sidebar"],
            fg=COLORS["text"],
            font=(FONT, 19, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            brand,
            text=f"v{__version__}  ·  CSV WORKBENCH",
            bg=COLORS["sidebar"],
            fg=COLORS["faint"],
            font=(FONT, 8, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(3, 0))

        tk.Label(
            self.sidebar,
            text="WORKFLOWS",
            bg=COLORS["sidebar"],
            fg=COLORS["faint"],
            font=(FONT, 8, "bold"),
            anchor="w",
        ).pack(fill="x", padx=20, pady=(0, 8))

        self._nav_button("join", "JOIN", "CSVを結合する")
        self._nav_button("split", "SPLIT", "キーで分割する")
        self._nav_button("transform", "TRANSFORM", "列を整形する")

        info = tk.Frame(self.sidebar, bg=COLORS["surface_soft"])
        info.pack(side="bottom", fill="x", padx=14, pady=14)
        tk.Label(
            info,
            text="LOCAL FIRST",
            bg=COLORS["surface_soft"],
            fg=COLORS["success"],
            font=(FONT, 8, "bold"),
            anchor="w",
        ).pack(fill="x", padx=12, pady=(10, 2))
        tk.Label(
            info,
            text="CSVはPC内で処理されます。\n実行前にPreviewで確認できます。",
            bg=COLORS["surface_soft"],
            fg=COLORS["muted"],
            font=(FONT, 8),
            justify="left",
            anchor="w",
        ).pack(fill="x", padx=12, pady=(0, 10))

        self.main = tk.Frame(self, bg=COLORS["app"])
        self.main.pack(side="left", fill="both", expand=True)

        self.page_host = tk.Frame(self.main, bg=COLORS["app"])
        self.page_host.pack(fill="both", expand=True, padx=26, pady=(22, 10))

        status = tk.Frame(self.main, bg=COLORS["app"])
        status.pack(fill="x", padx=26, pady=(0, 14))
        tk.Label(status, text="●", bg=COLORS["app"], fg=COLORS["success"], font=(FONT, 8)).pack(side="left")
        tk.Label(
            status,
            textvariable=self.status_var,
            bg=COLORS["app"],
            fg=COLORS["muted"],
            font=(FONT, 9),
            anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=(7, 0))

    def _nav_button(self, key: str, title: str, subtitle: str) -> None:
        holder = tk.Frame(self.sidebar, bg=COLORS["sidebar"])
        holder.pack(fill="x", padx=12, pady=3)
        button = tk.Button(
            holder,
            text=f"{title}\n{subtitle}",
            command=lambda: self._show_page(key),
            bg=COLORS["sidebar"],
            fg=COLORS["muted"],
            activebackground=COLORS["surface_alt"],
            activeforeground=COLORS["text"],
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=(FONT, 10, "bold"),
            justify="left",
            anchor="w",
            padx=14,
            pady=10,
            cursor="hand2",
        )
        button.pack(fill="x")
        self.nav_buttons[key] = button

    def _show_page(self, key: str) -> None:
        if key not in self.pages:
            return
        for page in self.pages.values():
            page.pack_forget()
        self.pages[key].pack(fill="both", expand=True)
        self.current_page = key
        for nav_key, button in self.nav_buttons.items():
            active = nav_key == key
            button.configure(
                bg=COLORS["accent_soft"] if active else COLORS["sidebar"],
                fg=COLORS["text"] if active else COLORS["muted"],
            )

    def _page(self, key: str, title: str, subtitle: str) -> tuple[tk.Frame, tk.Frame]:
        page = tk.Frame(self.page_host, bg=COLORS["app"])
        self.pages[key] = page

        header = tk.Frame(page, bg=COLORS["app"])
        header.pack(fill="x", pady=(0, 18))
        tk.Label(
            header,
            text=title,
            bg=COLORS["app"],
            fg=COLORS["text"],
            font=(FONT, 23, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            header,
            text=subtitle,
            bg=COLORS["app"],
            fg=COLORS["muted"],
            font=(FONT, 10),
            anchor="w",
        ).pack(fill="x", pady=(3, 0))

        scroll_host = tk.Frame(page, bg=COLORS["app"])
        scroll_host.pack(fill="both", expand=True)
        canvas = tk.Canvas(
            scroll_host,
            bg=COLORS["app"],
            highlightthickness=0,
            bd=0,
            relief="flat",
        )
        vbar = ttk.Scrollbar(
            scroll_host,
            orient="vertical",
            command=canvas.yview,
            style="Dark.Vertical.TScrollbar",
        )
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y", padx=(8, 0))

        body = tk.Frame(canvas, bg=COLORS["app"])
        window_id = canvas.create_window((0, 0), window=body, anchor="nw")

        def _sync_scroll_region(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _sync_body_width(event) -> None:
            canvas.itemconfigure(window_id, width=event.width)

        body.bind("<Configure>", _sync_scroll_region)
        canvas.bind("<Configure>", _sync_body_width)
        self.page_canvases[key] = canvas
        return page, body

    def _on_mousewheel(self, event) -> None:
        widget_class = event.widget.winfo_class() if hasattr(event.widget, "winfo_class") else ""
        if widget_class in {"Treeview", "Listbox", "TCombobox"}:
            return
        canvas = self.page_canvases.get(self.current_page)
        if canvas is None:
            return
        delta = int(-1 * (event.delta / 120)) if event.delta else 0
        if delta:
            canvas.yview_scroll(delta * 3, "units")

    def _card(self, parent: tk.Widget, title: str, caption: str = "") -> tuple[tk.Frame, tk.Frame]:
        card = tk.Frame(
            parent,
            bg=COLORS["surface"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        heading = tk.Frame(card, bg=COLORS["surface"])
        heading.pack(fill="x", padx=16, pady=(14, 10))
        tk.Label(
            heading,
            text=title,
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=(FONT, 10, "bold"),
            anchor="w",
        ).pack(fill="x")
        if caption:
            tk.Label(
                heading,
                text=caption,
                bg=COLORS["surface"],
                fg=COLORS["muted"],
                font=(FONT, 8),
                anchor="w",
            ).pack(fill="x", pady=(2, 0))
        body = tk.Frame(card, bg=COLORS["surface"])
        body.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        return card, body

    def _button(
        self,
        parent: tk.Widget,
        text: str,
        command,
        kind: str = "secondary",
        width: int | None = None,
    ) -> tk.Button:
        if kind == "primary":
            bg, fg, active = COLORS["accent"], "#FFFFFF", COLORS["accent_hover"]
        elif kind == "danger":
            bg, fg, active = COLORS["surface_alt"], COLORS["danger"], COLORS["surface_soft"]
        elif kind == "ghost":
            bg, fg, active = COLORS["surface"], COLORS["muted"], COLORS["surface_alt"]
        else:
            bg, fg, active = COLORS["surface_alt"], COLORS["text"], COLORS["accent_soft"]
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active,
            activeforeground=fg,
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=(FONT, 9, "bold"),
            padx=13,
            pady=8,
            width=width,
            cursor="hand2",
        )

    def _field(self, parent: tk.Widget, label: str, variable: tk.Variable, width: int = 26) -> tk.Entry:
        wrap = tk.Frame(parent, bg=COLORS["surface"])
        tk.Label(
            wrap,
            text=label,
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=(FONT, 8, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 5))
        entry = tk.Entry(
            wrap,
            textvariable=variable,
            width=width,
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["border_focus"],
            font=(FONT, 10),
        )
        entry.pack(fill="x", ipady=8)
        return entry

    def _combo(self, parent: tk.Widget, label: str, variable: tk.Variable, values=(), width: int = 26) -> ttk.Combobox:
        wrap = tk.Frame(parent, bg=COLORS["surface"])
        tk.Label(
            wrap,
            text=label,
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=(FONT, 8, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 5))
        combo = ttk.Combobox(
            wrap,
            textvariable=variable,
            values=values,
            state="readonly",
            width=width,
            style="Modern.TCombobox",
        )
        combo.pack(fill="x")
        return combo

    def _metric(self, parent: tk.Widget, label: str, variable: tk.StringVar, tone: str = "text") -> tk.Frame:
        frame = tk.Frame(parent, bg=COLORS["surface_soft"])
        tk.Label(
            frame,
            text=label,
            bg=COLORS["surface_soft"],
            fg=COLORS["muted"],
            font=(FONT, 8, "bold"),
        ).pack(anchor="w", padx=11, pady=(9, 0))
        tk.Label(
            frame,
            textvariable=variable,
            bg=COLORS["surface_soft"],
            fg=COLORS.get(tone, COLORS["text"]),
            font=(FONT, 16, "bold"),
        ).pack(anchor="w", padx=11, pady=(0, 8))
        return frame

    def _file_summary_label(self, parent: tk.Widget, variable: tk.StringVar) -> tk.Label:
        return tk.Label(
            parent,
            textvariable=variable,
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            font=(FONT, 9),
            justify="left",
            anchor="w",
            padx=12,
            pady=11,
            wraplength=480,
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
        self.join_left_summary = tk.StringVar(value="LEFT CSV\n未選択")
        self.join_right_summary = tk.StringVar(value="RIGHT CSV\n未選択")
        self.join_warning = tk.StringVar(value="ファイルを選択してPreviewを実行してください。")
        self.join_metric_match = tk.StringVar(value="—")
        self.join_metric_left = tk.StringVar(value="—")
        self.join_metric_right = tk.StringVar(value="—")
        self.join_metric_dups = tk.StringVar(value="—")
        self.join_metric_rows = tk.StringVar(value="—")

        _, body = self._page(
            "join",
            "Join CSV files",
            "キー結合または縦結合。結果件数と重複を確認してから出力します。",
        )

        top = tk.Frame(body, bg=COLORS["app"])
        top.pack(fill="x")
        top.grid_columnconfigure(0, weight=5)
        top.grid_columnconfigure(1, weight=4)

        files_card, files_body = self._card(top, "INPUT FILES", "左・右CSV、または縦結合する複数CSVを選択")
        files_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        settings_card, settings_body = self._card(top, "JOIN SETTINGS", "処理方式とキー列を指定")
        settings_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        mode = tk.Frame(files_body, bg=COLORS["surface"])
        mode.pack(fill="x", pady=(0, 10))
        self.join_key_mode_button = self._button(mode, "KEY JOIN", lambda: self._set_join_mode("key"), "primary")
        self.join_key_mode_button.pack(side="left")
        self.join_vertical_mode_button = self._button(mode, "STACK", lambda: self._set_join_mode("vertical"))
        self.join_vertical_mode_button.pack(side="left", padx=(8, 0))

        file_grid = tk.Frame(files_body, bg=COLORS["surface"])
        file_grid.pack(fill="x")
        file_grid.grid_columnconfigure(0, weight=1)
        file_grid.grid_columnconfigure(1, weight=1)

        left_box = tk.Frame(file_grid, bg=COLORS["surface"])
        left_box.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self._file_summary_label(left_box, self.join_left_summary).pack(fill="x")
        self.join_left_file_button = self._button(left_box, "左CSVを選択", lambda: self._load_join_side("left"), "secondary")
        self.join_left_file_button.pack(fill="x", pady=(7, 0))

        right_box = tk.Frame(file_grid, bg=COLORS["surface"])
        right_box.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        self._file_summary_label(right_box, self.join_right_summary).pack(fill="x")
        self.join_right_file_button = self._button(right_box, "右CSVを選択", lambda: self._load_join_side("right"), "secondary")
        self.join_right_file_button.pack(fill="x", pady=(7, 0))

        self.join_stack_button = self._button(files_body, "結合するCSVをまとめて選択", self._load_vertical_files, "secondary")
        self.join_stack_button.pack(fill="x", pady=(10, 0))

        settings_grid = tk.Frame(settings_body, bg=COLORS["surface"])
        settings_grid.pack(fill="x")
        settings_grid.grid_columnconfigure(0, weight=1)
        settings_grid.grid_columnconfigure(1, weight=1)

        left_key_wrap = tk.Frame(settings_grid, bg=COLORS["surface"])
        left_key_wrap.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.join_left_combo = self._combo(left_key_wrap, "LEFT KEY", self.join_left_key)
        self.join_left_combo.master.pack(fill="x")

        right_key_wrap = tk.Frame(settings_grid, bg=COLORS["surface"])
        right_key_wrap.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        self.join_right_combo = self._combo(right_key_wrap, "RIGHT KEY", self.join_right_key)
        self.join_right_combo.master.pack(fill="x")
        self.join_left_combo.bind("<<ComboboxSelected>>", lambda _e: self._invalidate_join_result("JOINキーが変更されました。再解析してください。"))
        self.join_right_combo.bind("<<ComboboxSelected>>", lambda _e: self._invalidate_join_result("JOINキーが変更されました。再解析してください。"))

        tk.Label(
            settings_body,
            text="JOIN TYPE",
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=(FONT, 8, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(13, 5))
        type_row = tk.Frame(settings_body, bg=COLORS["surface"])
        type_row.pack(fill="x")
        self.join_left_type_button = self._button(type_row, "LEFT", lambda: self._set_join_type("left"), "primary")
        self.join_left_type_button.pack(side="left", fill="x", expand=True)
        self.join_inner_type_button = self._button(type_row, "INNER", lambda: self._set_join_type("inner"))
        self.join_inner_type_button.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.join_preview_button = self._button(settings_body, "解析してプレビュー", self._run_join_preview, "primary")
        self.join_preview_button.pack(fill="x", pady=(14, 0))

        metrics_card, metrics_body = self._card(body, "CHECK", "一致・片側のみ・重複を実行前に確認")
        metrics_card.pack(fill="x", pady=(16, 0))
        metrics = tk.Frame(metrics_body, bg=COLORS["surface"])
        metrics.pack(fill="x")
        for i in range(5):
            metrics.grid_columnconfigure(i, weight=1)
        self._metric(metrics, "MATCHED KEYS", self.join_metric_match, "success").grid(row=0, column=0, sticky="ew", padx=(0, 5))
        self._metric(metrics, "LEFT ONLY", self.join_metric_left).grid(row=0, column=1, sticky="ew", padx=5)
        self._metric(metrics, "RIGHT ONLY", self.join_metric_right).grid(row=0, column=2, sticky="ew", padx=5)
        self._metric(metrics, "DUP KEYS L / R", self.join_metric_dups, "warning").grid(row=0, column=3, sticky="ew", padx=5)
        self._metric(metrics, "OUTPUT ROWS", self.join_metric_rows, "accent").grid(row=0, column=4, sticky="ew", padx=(5, 0))
        tk.Label(
            metrics_body,
            textvariable=self.join_warning,
            bg=COLORS["surface"],
            fg=COLORS["warning"],
            font=(FONT, 8),
            justify="left",
            anchor="w",
            wraplength=1000,
        ).pack(fill="x", pady=(10, 0))

        preview_card, preview_body = self._card(body, "PREVIEW", "先頭100行を表示。セル内改行は ↵ で表示")
        preview_card.pack(fill="both", expand=True, pady=(16, 0))
        self.join_preview = PreviewTable(preview_body, height=3)
        self.join_preview.pack(fill="both", expand=True)

        export = tk.Frame(preview_body, bg=COLORS["surface"])
        export.pack(fill="x", pady=(12, 0))
        export.grid_columnconfigure(0, weight=1)
        path_wrap = tk.Frame(export, bg=COLORS["surface"])
        path_wrap.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._field(path_wrap, "出力先", self.join_output).master.pack(fill="x")
        self._button(export, "参照", self._choose_join_output).grid(row=0, column=1, sticky="s")
        footer = tk.Frame(export, bg=COLORS["surface"])
        footer.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        footer.grid_columnconfigure(1, weight=1)
        enc_wrap = tk.Frame(footer, bg=COLORS["surface"])
        enc_wrap.grid(row=0, column=0, sticky="w")
        self._combo(enc_wrap, "文字コード", self.join_encoding, list(EXPORT_ENCODINGS), width=13).master.pack(fill="x")
        self.join_export_button = self._button(footer, "CSVを書き出す", self._export_join, "primary")
        self.join_export_button.grid(row=0, column=2, sticky="e")

        self._set_join_mode("key")
        self._set_join_type("left")

    def _refresh_join_file_summaries(self) -> None:
        if self.join_mode.get() == "vertical":
            if self.join_files:
                lines = [f"{d.path.name} · {d.rows:,} rows · {len(d.columns)} cols" for d in self.join_files[:3]]
                more = f"\n+ {len(self.join_files) - 3} more" if len(self.join_files) > 3 else ""
                self.join_left_summary.set(f"STACK INPUT\n{len(self.join_files)} files\n" + "\n".join(lines) + more)
            else:
                self.join_left_summary.set("STACK INPUT\n未選択")
            self.join_right_summary.set("STACK MODE\n右側CSVの指定は不要です。")
            return

        if self.join_left:
            self.join_left_summary.set(
                f"LEFT CSV\n{self.join_left.path.name}\n{self.join_left.rows:,} rows · {len(self.join_left.columns)} cols · {self.join_left.encoding}"
            )
        else:
            self.join_left_summary.set("LEFT CSV\n未選択")
        if self.join_right:
            self.join_right_summary.set(
                f"RIGHT CSV\n{self.join_right.path.name}\n{self.join_right.rows:,} rows · {len(self.join_right.columns)} cols · {self.join_right.encoding}"
            )
        else:
            self.join_right_summary.set("RIGHT CSV\n未選択")

    def _invalidate_join_result(self, message: str = "条件が変更されました。再解析してください。") -> None:
        self.join_result = None
        self.join_preview.clear()
        self._reset_join_metrics()
        self.join_warning.set(message)

    def _set_join_mode(self, mode: str) -> None:
        changed = self.join_mode.get() != mode
        self.join_mode.set(mode)
        is_key = mode == "key"
        self.join_key_mode_button.configure(
            bg=COLORS["accent"] if is_key else COLORS["surface_alt"],
            fg="#FFFFFF" if is_key else COLORS["text"],
        )
        self.join_vertical_mode_button.configure(
            bg=COLORS["surface_alt"] if is_key else COLORS["accent"],
            fg=COLORS["text"] if is_key else "#FFFFFF",
        )
        state = "readonly" if is_key else "disabled"
        self.join_left_combo.configure(state=state)
        self.join_right_combo.configure(state=state)
        self.join_left_type_button.configure(state="normal" if is_key else "disabled")
        self.join_inner_type_button.configure(state="normal" if is_key else "disabled")
        self.join_stack_button.configure(
            state="disabled" if is_key else "normal",
            bg=COLORS["surface_alt"] if not is_key else COLORS["surface_soft"],
            fg=COLORS["text"] if not is_key else COLORS["faint"],
        )
        self.join_left_file_button.configure(state="normal" if is_key else "disabled")
        self.join_right_file_button.configure(state="normal" if is_key else "disabled")
        self._refresh_join_file_summaries()
        if changed:
            self._invalidate_join_result(
                "左右のキーを指定し、重複や多対多を確認してから出力します。" if is_key
                else "STACKでは列名を基準に揃え、不足列を空欄で補完して縦結合します。"
            )
        elif is_key:
            self.join_warning.set("左右のキーを指定し、重複や多対多を確認してから出力します。")
        else:
            self.join_warning.set("STACKでは列名を基準に揃え、不足列を空欄で補完して縦結合します。")

    def _set_join_type(self, join_type: str) -> None:
        changed = self.join_type.get() != join_type
        self.join_type.set(join_type)
        is_left = join_type == "left"
        self.join_left_type_button.configure(bg=COLORS["accent"] if is_left else COLORS["surface_alt"])
        self.join_inner_type_button.configure(bg=COLORS["surface_alt"] if is_left else COLORS["accent"])
        if changed:
            self._invalidate_join_result("JOIN TYPEが変更されました。再解析してください。")

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
            self._refresh_join_file_summaries()
            self.join_result = None
            self.join_preview.clear()
            self._reset_join_metrics()
            self.status_var.set(f"Loaded {doc.path.name}")
        except Exception as exc:
            self._show_error(exc)

    def _load_vertical_files(self) -> None:
        paths = filedialog.askopenfilenames(title="縦結合するCSVを選択", filetypes=[("CSV files", "*.csv")])
        if not paths:
            return
        try:
            self.join_files = [read_csv_flexible(Path(p)) for p in paths]
            if len(self.join_files) < 2:
                raise CsvToolError("STACKには2つ以上のCSVを選択してください。")
            self._refresh_join_file_summaries()
            self.join_result = None
            self.join_preview.clear()
            self._reset_join_metrics()
            self.status_var.set(f"Loaded {len(self.join_files)} CSV files for STACK")
        except Exception as exc:
            self._show_error(exc)

    def _reset_join_metrics(self) -> None:
        for var in (
            self.join_metric_match,
            self.join_metric_left,
            self.join_metric_right,
            self.join_metric_dups,
            self.join_metric_rows,
        ):
            var.set("—")

    def _run_join_preview(self) -> None:
        try:
            if self.join_mode.get() == "vertical":
                self.join_result = vertical_join([d.dataframe for d in self.join_files])
                self.join_metric_match.set("N/A")
                self.join_metric_left.set(str(len(self.join_files)))
                self.join_metric_right.set("N/A")
                self.join_metric_dups.set("N/A")
                self.join_metric_rows.set(f"{len(self.join_result):,}")
                self.join_warning.set(
                    f"{len(self.join_files)} filesを結合。出力列は {len(self.join_result.columns)} columns。不足列は空欄で補完します。"
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
                self.join_metric_match.set(f"{diag.matched_keys:,}")
                self.join_metric_left.set(f"{diag.left_only_keys:,}")
                self.join_metric_right.set(f"{diag.right_only_keys:,}")
                self.join_metric_dups.set(f"{diag.left_duplicate_keys:,} / {diag.right_duplicate_keys:,}")
                self.join_metric_rows.set(f"{diag.result_rows:,}")
                self.join_warning.set(" / ".join(diag.warnings) if diag.warnings else "No warnings — キー構成に問題は見つかりませんでした。")
                self.join_output.set(str(self.join_left.path.parent / "merged.csv"))
            self.join_preview.show_dataframe(self.join_result)
            self.status_var.set("JOIN preview ready — 内容を確認してからExportしてください。")
        except Exception as exc:
            self._show_error(exc)

    def _choose_join_output(self) -> None:
        path = self._save_csv_path("merged.csv")
        if path:
            self.join_output.set(path)

    def _export_join(self) -> None:
        try:
            if self.join_result is None:
                raise CsvToolError("先に Analyze & Preview を実行してください。")
            if not self.join_output.get().strip():
                raise CsvToolError("出力先を指定してください。")
            export_csv(self.join_result, Path(self.join_output.get()), self.join_encoding.get())
            self.status_var.set(f"Exported: {self.join_output.get()}")
            messagebox.showinfo(APP_TITLE, "CSVを出力しました。")
        except Exception as exc:
            self._show_error(exc)

    # ---------- SPLIT ----------
    def _build_split_page(self) -> None:
        self.split_doc: CsvDocument | None = None
        self.split_key = tk.StringVar()
        self.split_include_blank = tk.BooleanVar(value=False)
        self.split_output_dir = tk.StringVar()
        self.split_encoding = tk.StringVar(value="UTF-8 BOM")
        self.split_summary = tk.StringVar(value="CSV未選択")
        self.split_file_count = tk.StringVar(value="—")
        self.split_row_count = tk.StringVar(value="—")
        self.split_blank_count = tk.StringVar(value="—")
        self.split_analysis_valid = False

        _, body = self._page(
            "split",
            "Split a CSV",
            "指定列の値ごとにCSVを分割。出力ファイル数を確認してから生成します。",
        )

        top = tk.Frame(body, bg=COLORS["app"])
        top.pack(fill="x")
        top.grid_columnconfigure(0, weight=1)
        top.grid_columnconfigure(1, weight=1)

        input_card, input_body = self._card(top, "INPUT", "分割するCSVを選択")
        input_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        settings_card, settings_body = self._card(top, "SPLIT SETTINGS", "キー列と空欄の扱い")
        settings_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        self._file_summary_label(input_body, self.split_summary).pack(fill="x")
        self._button(input_body, "CSVを選択", self._load_split_file, "primary").pack(fill="x", pady=(10, 0))

        combo_wrap = tk.Frame(settings_body, bg=COLORS["surface"])
        combo_wrap.pack(fill="x")
        self.split_key_combo = self._combo(combo_wrap, "SPLIT KEY", self.split_key)
        self.split_key_combo.master.pack(fill="x")
        self.split_key_combo.bind("<<ComboboxSelected>>", lambda _e: self._invalidate_split_analysis())

        blank_row = tk.Frame(settings_body, bg=COLORS["surface"])
        blank_row.pack(fill="x", pady=(13, 0))
        self.split_blank_check = tk.Checkbutton(
            blank_row,
            text="空欄キーを blank.csv として出力",
            variable=self.split_include_blank,
            bg=COLORS["surface"],
            fg=COLORS["text"],
            selectcolor=COLORS["surface_soft"],
            activebackground=COLORS["surface"],
            activeforeground=COLORS["text"],
            font=(FONT, 9),
            relief="flat",
            command=self._invalidate_split_analysis,
        )
        self.split_blank_check.pack(anchor="w")
        self._button(settings_body, "分割内容を確認", self._preview_split_counts, "primary").pack(fill="x", pady=(12, 0))

        check_card, check_body = self._card(body, "CHECK", "分割前の見込み")
        check_card.pack(fill="x", pady=(16, 0))
        metrics = tk.Frame(check_body, bg=COLORS["surface"])
        metrics.pack(fill="x")
        for i in range(3):
            metrics.grid_columnconfigure(i, weight=1)
        self._metric(metrics, "OUTPUT FILES", self.split_file_count, "accent").grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._metric(metrics, "SOURCE ROWS", self.split_row_count).grid(row=0, column=1, sticky="ew", padx=6)
        self._metric(metrics, "BLANK KEY ROWS", self.split_blank_count, "warning").grid(row=0, column=2, sticky="ew", padx=(6, 0))

        preview_card, preview_body = self._card(body, "GROUP PREVIEW", "キー別の出力件数")
        preview_card.pack(fill="both", expand=True, pady=(16, 0))
        self.split_preview = PreviewTable(preview_body, height=4)
        self.split_preview.pack(fill="both", expand=True)

        export = tk.Frame(preview_body, bg=COLORS["surface"])
        export.pack(fill="x", pady=(12, 0))
        export.grid_columnconfigure(0, weight=1)
        path_wrap = tk.Frame(export, bg=COLORS["surface"])
        path_wrap.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._field(path_wrap, "出力フォルダ", self.split_output_dir).master.pack(fill="x")
        self._button(export, "参照", self._choose_split_output).grid(row=0, column=1, sticky="s")
        footer = tk.Frame(export, bg=COLORS["surface"])
        footer.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        footer.grid_columnconfigure(1, weight=1)
        enc_wrap = tk.Frame(footer, bg=COLORS["surface"])
        enc_wrap.grid(row=0, column=0, sticky="w")
        self._combo(enc_wrap, "文字コード", self.split_encoding, list(EXPORT_ENCODINGS), width=13).master.pack(fill="x")
        self.split_export_button = self._button(footer, "分割CSVを書き出す", self._export_split, "primary")
        self.split_export_button.grid(row=0, column=2, sticky="e")

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
                f"{self.split_doc.path.name}\n{self.split_doc.rows:,} rows · {len(self.split_doc.columns)} cols · {self.split_doc.encoding}"
            )
            self.split_row_count.set(f"{self.split_doc.rows:,}")
            self.split_file_count.set("—")
            self.split_blank_count.set("—")
            self.split_analysis_valid = False
            self.split_preview.clear()
            self.status_var.set(f"Loaded {self.split_doc.path.name} — キーを選んで分割内容を確認してください。")
        except Exception as exc:
            self._show_error(exc)

    def _invalidate_split_analysis(self) -> None:
        self.split_analysis_valid = False
        self.split_file_count.set("—")
        self.split_blank_count.set("—")
        if hasattr(self, "split_preview"):
            self.split_preview.clear()
        if self.split_doc:
            self.status_var.set("SPLIT条件が変更されました。分割内容を再確認してください。")

    def _preview_split_counts(self) -> None:
        try:
            if not self.split_doc:
                raise CsvToolError("CSVを選択してください。")
            plan = split_plan(
                self.split_doc.dataframe,
                self.split_key.get(),
                include_blank=self.split_include_blank.get(),
            )
            blank_rows = int((self.split_doc.dataframe[self.split_key.get()].astype(str).str.strip() == "").sum())
            self.split_file_count.set(f"{len(plan):,}")
            self.split_row_count.set(f"{self.split_doc.rows:,}")
            self.split_blank_count.set(f"{blank_rows:,}")
            self.split_preview.show_dataframe(plan, 500)
            self.split_analysis_valid = True
            self.status_var.set("SPLIT analysis ready — 出力ファイル名と件数を確認してください。")
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
            if not self.split_analysis_valid:
                raise CsvToolError("先に「分割内容を確認」を実行してください。")
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
            self.status_var.set(f"Exported {len(paths)} split CSV files")
            messagebox.showinfo(APP_TITLE, f"{len(paths)} 個のCSVを出力しました。")
        except Exception as exc:
            self._show_error(exc)

    # ---------- TRANSFORM ----------
    def _build_transform_page(self) -> None:
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
        self.transform_hint_text = tk.StringVar(value="")
        self.transform_step_count = tk.StringVar(value="0")
        self.transform_output_cols = tk.StringVar(value="—")

        _, body = self._page(
            "transform",
            "Transform a CSV",
            "列操作をステップとして積み、Before / Afterを比較してから出力します。",
        )

        top = tk.Frame(body, bg=COLORS["app"])
        top.pack(fill="x")
        top.grid_columnconfigure(0, weight=5)
        top.grid_columnconfigure(1, weight=5)

        pipeline_card, pipeline_body = self._card(top, "PIPELINE", "入力CSVと適用予定ステップ")
        pipeline_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        add_card, add_body = self._card(top, "ADD TRANSFORM", "処理を1つずつ追加")
        add_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        input_row = tk.Frame(pipeline_body, bg=COLORS["surface"])
        input_row.pack(fill="x")
        self._file_summary_label(input_row, self.transform_summary).pack(side="left", fill="both", expand=True)
        self._button(input_row, "CSVを選択", self._load_transform_file, "primary").pack(side="left", padx=(8, 0), fill="y")

        tk.Label(
            pipeline_body,
            text="STEPS",
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=(FONT, 8, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(12, 5))
        self.transform_step_list = tk.Listbox(
            pipeline_body,
            height=3,
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            selectbackground=COLORS["accent_soft"],
            selectforeground=COLORS["text"],
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            font=(FONT, 9),
            activestyle="none",
        )
        self.transform_step_list.pack(fill="both", expand=True)
        actions = tk.Frame(pipeline_body, bg=COLORS["surface"])
        actions.pack(fill="x", pady=(8, 0))
        self._button(actions, "選択したステップを削除", self._remove_transform_step, "ghost").pack(side="left")
        self._button(actions, "すべてクリア", self._clear_transform_steps, "danger").pack(side="left", padx=(7, 0))

        op_values = [
            "rename | 列名変更",
            "delete | 列削除",
            "add_fixed | 固定値列追加",
            "trim | 前後空白削除",
            "replace | 文字列置換",
            "date | 日付形式変換",
            "zero_pad | 0埋め",
            "move | 列順変更",
        ]

        op_wrap = tk.Frame(add_body, bg=COLORS["surface"])
        op_wrap.pack(fill="x")
        self.transform_op_combo = self._combo(op_wrap, "OPERATION", tk.StringVar(), op_values)
        self.transform_op_combo.master.pack(fill="x")
        self.transform_op_combo.set(op_values[0])
        self.transform_op_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_transform_operation())

        col_wrap = tk.Frame(add_body, bg=COLORS["surface"])
        col_wrap.pack(fill="x", pady=(10, 0))
        self.transform_col_combo = self._combo(col_wrap, "TARGET COLUMN", self.transform_column)
        self.transform_col_combo.master.pack(fill="x")

        values = tk.Frame(add_body, bg=COLORS["surface"])
        values.pack(fill="x", pady=(10, 0))
        values.grid_columnconfigure(0, weight=1)
        values.grid_columnconfigure(1, weight=1)
        self.transform_v1_wrap = tk.Frame(values, bg=COLORS["surface"])
        self.transform_v1_wrap.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        self.transform_v1_caption = tk.StringVar(value="VALUE 1")
        self.transform_v1_entry = self._labeled_dynamic_field(self.transform_v1_wrap, self.transform_v1_caption, self.transform_value1)
        self.transform_v2_wrap = tk.Frame(values, bg=COLORS["surface"])
        self.transform_v2_wrap.grid(row=0, column=1, sticky="ew", padx=(5, 0))
        self.transform_v2_caption = tk.StringVar(value="VALUE 2")
        self.transform_v2_entry = self._labeled_dynamic_field(self.transform_v2_wrap, self.transform_v2_caption, self.transform_value2)

        tk.Label(
            add_body,
            textvariable=self.transform_hint_text,
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=(FONT, 8),
            justify="left",
            anchor="w",
            wraplength=470,
        ).pack(fill="x", pady=(9, 0))
        self._button(add_body, "ステップを追加", self._add_transform_step, "primary").pack(fill="x", pady=(11, 0))
        self._sync_transform_operation()

        stats_card, stats_body = self._card(body, "PIPELINE CHECK", "適用ステップと出力列")
        stats_card.pack(fill="x", pady=(16, 0))
        stats = tk.Frame(stats_body, bg=COLORS["surface"])
        stats.pack(fill="x")
        stats.grid_columnconfigure(0, weight=1)
        stats.grid_columnconfigure(1, weight=1)
        self._metric(stats, "STEPS", self.transform_step_count, "accent").grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._metric(stats, "OUTPUT COLUMNS", self.transform_output_cols).grid(row=0, column=1, sticky="ew", padx=(6, 0))

        preview_card, preview_body = self._card(body, "BEFORE / AFTER", "変換結果を比較")
        preview_card.pack(fill="both", expand=True, pady=(16, 0))
        split = tk.PanedWindow(preview_body, orient="horizontal", bg=COLORS["border"], sashwidth=5, bd=0, relief="flat")
        split.pack(fill="both", expand=True)
        before = tk.Frame(split, bg=COLORS["surface"])
        after = tk.Frame(split, bg=COLORS["surface"])
        split.add(before, stretch="always")
        split.add(after, stretch="always")
        tk.Label(before, text="BEFORE", bg=COLORS["surface"], fg=COLORS["muted"], font=(FONT, 8, "bold")).pack(anchor="w", pady=(0, 5))
        tk.Label(after, text="AFTER", bg=COLORS["surface"], fg=COLORS["success"], font=(FONT, 8, "bold")).pack(anchor="w", pady=(0, 5))
        self.transform_before = PreviewTable(before, height=3)
        self.transform_after = PreviewTable(after, height=3)
        self.transform_before.pack(fill="both", expand=True)
        self.transform_after.pack(fill="both", expand=True)

        export = tk.Frame(preview_body, bg=COLORS["surface"])
        export.pack(fill="x", pady=(12, 0))
        export.grid_columnconfigure(0, weight=1)
        path_wrap = tk.Frame(export, bg=COLORS["surface"])
        path_wrap.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._field(path_wrap, "出力先", self.transform_output).master.pack(fill="x")
        self._button(export, "参照", self._choose_transform_output).grid(row=0, column=1, sticky="s")
        footer = tk.Frame(export, bg=COLORS["surface"])
        footer.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        footer.grid_columnconfigure(1, weight=1)
        tk.Label(
            footer,
            text="ステップ追加・削除時にAfterを自動更新します。",
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=(FONT, 8),
        ).grid(row=0, column=0, sticky="w")
        enc_wrap = tk.Frame(footer, bg=COLORS["surface"])
        enc_wrap.grid(row=0, column=2, sticky="e", padx=(12, 8))
        self._combo(enc_wrap, "文字コード", self.transform_encoding, list(EXPORT_ENCODINGS), width=13).master.pack(fill="x")
        self.transform_export_button = self._button(footer, "CSVを書き出す", self._export_transform, "primary")
        self.transform_export_button.grid(row=0, column=3, sticky="e")

    def _labeled_dynamic_field(self, parent: tk.Widget, caption: tk.StringVar, variable: tk.StringVar) -> tk.Entry:
        tk.Label(
            parent,
            textvariable=caption,
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=(FONT, 8, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 5))
        entry = tk.Entry(
            parent,
            textvariable=variable,
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["border_focus"],
            font=(FONT, 10),
        )
        entry.pack(fill="x", ipady=8)
        return entry

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
                f"{self.transform_doc.path.name}\n{self.transform_doc.rows:,} rows · {len(self.transform_doc.columns)} cols · {self.transform_doc.encoding}"
            )
            self.transform_before.show_dataframe(self.transform_doc.dataframe)
            self.transform_after.clear()
            self.transform_result = None
            self.transform_step_count.set("0")
            self.transform_output_cols.set(f"{len(self.transform_doc.columns):,}")
            self.status_var.set(f"Loaded {self.transform_doc.path.name} for TRANSFORM")
        except Exception as exc:
            self._show_error(exc)

    def _sync_transform_operation(self) -> None:
        raw = self.transform_op_combo.get()
        op = raw.split("|", 1)[0].strip() if raw else "rename"
        self.transform_operation.set(op)
        hints = {
            "rename": ("NEW COLUMN NAME", "UNUSED", "列名だけを変更します。例: name → student_name"),
            "delete": ("UNUSED", "UNUSED", "対象列を削除します。"),
            "add_fixed": ("NEW COLUMN", "FIXED VALUE", "対象列は不要です。例: status / checked"),
            "trim": ("UNUSED", "UNUSED", "対象列の前後空白を削除します。"),
            "replace": ("FIND", "REPLACE WITH", "部分一致の文字列を置換します。"),
            "date": ("OUTPUT FORMAT", "UNUSED", "例: %Y-%m-%d。解析できない値は元のまま残します。"),
            "zero_pad": ("DIGITS", "UNUSED", "例: 6 → 123 を 000123 にします。"),
            "move": ("MOVE TO", "UNUSED", "先頭 / 末尾 / 1始まりの列番号を指定します。"),
        }
        c1, c2, hint = hints[op]
        self.transform_v1_caption.set(c1)
        self.transform_v2_caption.set(c2)
        self.transform_hint_text.set(hint)
        if op == "add_fixed":
            self.transform_col_combo.configure(state="disabled")
        else:
            self.transform_col_combo.configure(state="readonly")

        needs_v1 = op in {"rename", "add_fixed", "replace", "date", "zero_pad", "move"}
        needs_v2 = op in {"add_fixed", "replace"}
        self.transform_v1_entry.configure(state="normal" if needs_v1 else "disabled")
        self.transform_v2_entry.configure(state="normal" if needs_v2 else "disabled")
        if not needs_v1:
            self.transform_value1.set("")
        if not needs_v2:
            self.transform_value2.set("")
        self.transform_v1_wrap.grid_remove()
        self.transform_v2_wrap.grid_remove()
        if needs_v1 and needs_v2:
            self.transform_v1_wrap.grid(row=0, column=0, sticky="ew", padx=(0, 5))
            self.transform_v2_wrap.grid(row=0, column=1, sticky="ew", padx=(5, 0))
        elif needs_v1:
            self.transform_v1_wrap.grid(row=0, column=0, columnspan=2, sticky="ew", padx=0)

    def _add_transform_step(self) -> None:
        try:
            if not self.transform_doc:
                raise CsvToolError("先にCSVを選択してください。")
            op = self.transform_operation.get()
            column = self.transform_column.get()
            if op != "add_fixed" and not column:
                raise CsvToolError("対象列を選択してください。")
            step = TransformStep(op, column, self.transform_value1.get(), self.transform_value2.get())
            preview = apply_transform_pipeline(self.transform_doc.dataframe, [*self.transform_steps, step])
            self.transform_steps.append(step)
            self._refresh_transform_step_list()
            self.transform_result = preview
            self.transform_after.show_dataframe(preview)
            self.transform_step_count.set(str(len(self.transform_steps)))
            self.transform_output_cols.set(str(len(preview.columns)))
            self.transform_col_combo["values"] = list(preview.columns)
            if self.transform_column.get() not in preview.columns:
                self.transform_column.set(str(preview.columns[0]) if len(preview.columns) else "")
            self.transform_value1.set("")
            self.transform_value2.set("")
            self.status_var.set(f"Transform step added — Afterを更新しました ({len(preview):,} rows)")
        except Exception as exc:
            self._show_error(exc)

    def _refresh_transform_step_list(self) -> None:
        self.transform_step_list.delete(0, "end")
        for index, step in enumerate(self.transform_steps, start=1):
            self.transform_step_list.insert("end", f"{index:02d}   {step.describe()}")
        self.transform_step_count.set(str(len(self.transform_steps)))

    def _remove_transform_step(self) -> None:
        selection = self.transform_step_list.curselection()
        if not selection:
            return
        del self.transform_steps[selection[0]]
        self._refresh_transform_step_list()
        self.transform_result = None
        self.transform_after.clear()
        if self.transform_doc:
            try:
                preview = apply_transform_pipeline(self.transform_doc.dataframe, self.transform_steps)
                self.transform_output_cols.set(str(len(preview.columns)))
                self.transform_col_combo["values"] = list(preview.columns)
                if self.transform_steps:
                    self.transform_result = preview
                    self.transform_after.show_dataframe(preview)
                if self.transform_column.get() not in preview.columns:
                    self.transform_column.set(str(preview.columns[0]) if len(preview.columns) else "")
            except CsvToolError:
                self.transform_output_cols.set("—")

    def _clear_transform_steps(self) -> None:
        self.transform_steps.clear()
        self._refresh_transform_step_list()
        self.transform_after.clear()
        self.transform_result = None
        self.transform_output_cols.set(str(len(self.transform_doc.columns)) if self.transform_doc else "—")
        if self.transform_doc:
            self.transform_col_combo["values"] = self.transform_doc.columns
            self.transform_column.set(self.transform_doc.columns[0] if self.transform_doc.columns else "")
        self.status_var.set("Transform steps cleared")

    def _apply_transform_preview(self) -> None:
        try:
            if not self.transform_doc:
                raise CsvToolError("CSVを選択してください。")
            if not self.transform_steps:
                raise CsvToolError("変換ステップを1つ以上追加してください。")
            self.transform_result = apply_transform_pipeline(self.transform_doc.dataframe, self.transform_steps)
            self.transform_after.show_dataframe(self.transform_result)
            self.transform_output_cols.set(str(len(self.transform_result.columns)))
            self.status_var.set(
                f"TRANSFORM preview ready — {len(self.transform_result):,} rows · {len(self.transform_result.columns)} cols"
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
                raise CsvToolError("変換ステップを追加してAfterを確認してください。")
            if not self.transform_output.get().strip():
                raise CsvToolError("出力先を指定してください。")
            export_csv(self.transform_result, Path(self.transform_output.get()), self.transform_encoding.get())
            self.status_var.set(f"Exported: {self.transform_output.get()}")
            messagebox.showinfo(APP_TITLE, "CSVを出力しました。")
        except Exception as exc:
            self._show_error(exc)

    # ---------- errors ----------
    def _show_error(self, exc: Exception) -> None:
        self.status_var.set("Stopped — 内容を確認してください。")
        if isinstance(exc, CsvToolError):
            messagebox.showerror(APP_TITLE, str(exc))
        else:
            messagebox.showerror(APP_TITLE, f"想定外エラー:\n{exc}")


def run_app() -> None:
    app = CSVJoinerApp()
    app.mainloop()
