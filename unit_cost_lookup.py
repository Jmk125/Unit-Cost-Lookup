from __future__ import annotations

import ctypes
from ctypes import wintypes
import json
import queue
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox, ttk

import requests
import win32clipboard
import win32com.client


APP_NAME = "UNIT COST TERMINAL"
CONFIG_PATH = Path(__file__).with_name("config.json")

WM_HOTKEY = 0x0312
WM_QUIT = 0x0012
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002

HOTKEY_TOGGLE_ID = 1
HOTKEY_INSERT_ID = 2
VK_U = 0x55
VK_RETURN = 0x0D

BG = "#101418"
PANEL = "#171d22"
FIELD = "#0b0f12"
FG = "#d7e0e5"
MUTED = "#84939c"
ACCENT = "#72d572"
SELECT = "#294a33"
BORDER = "#3c474e"
ERROR = "#ff7b72"


@dataclass
class UnitCost:
    id: int
    division: str
    name: str
    cost_per_unit: float
    output_unit: str
    published_at: str
    comments: str = ""
    has_stale_material: bool = False


class HotkeyThread(threading.Thread):
    def __init__(self, events: queue.Queue[str]) -> None:
        super().__init__(daemon=True)
        self.events = events
        self.thread_id: int | None = None
        self.user32 = ctypes.windll.user32
        self.kernel32 = ctypes.windll.kernel32

    def run(self) -> None:
        self.thread_id = self.kernel32.GetCurrentThreadId()

        if not self.user32.RegisterHotKey(
            None, HOTKEY_TOGGLE_ID, MOD_CONTROL | MOD_ALT, VK_U
        ):
            self.events.put("toggle_failed")

        if not self.user32.RegisterHotKey(
            None, HOTKEY_INSERT_ID, MOD_CONTROL | MOD_ALT, VK_RETURN
        ):
            self.events.put("insert_failed")

        msg = wintypes.MSG()

        while self.user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            if msg.message == WM_HOTKEY:
                if msg.wParam == HOTKEY_TOGGLE_ID:
                    self.events.put("toggle")
                elif msg.wParam == HOTKEY_INSERT_ID:
                    self.events.put("insert")

        self.user32.UnregisterHotKey(None, HOTKEY_TOGGLE_ID)
        self.user32.UnregisterHotKey(None, HOTKEY_INSERT_ID)

    def stop(self) -> None:
        if self.thread_id:
            self.user32.PostThreadMessageW(self.thread_id, WM_QUIT, 0, 0)


class UnitCostLookupApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("920x690")
        self.root.minsize(760, 560)
        self.root.configure(bg=BG)

        self.config = self.load_config()
        self.all_costs: list[UnitCost] = []
        self.visible_costs: list[UnitCost] = []
        self.selected_cost: UnitCost | None = None
        self.pending_insert = False
        self.loading = False

        self.events: queue.Queue[str] = queue.Queue()
        self.hotkeys = HotkeyThread(self.events)
        self.hotkeys.start()

        self.configure_style()
        self.build_ui()

        self.root.protocol("WM_DELETE_WINDOW", self.hide)
        self.root.after(100, self.process_events)
        self.root.after(200, self.refresh_data)

    def load_config(self) -> dict:
        defaults = {
            "server_url": "http://YOUR-PI-IP:3077",
            "currency_format": "$#,##0.00",
        }

        try:
            loaded = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            return {**defaults, **loaded}
        except Exception:
            return defaults

    def configure_style(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(".", background=BG, foreground=FG, font=("Consolas", 10))
        style.configure("TFrame", background=BG)
        style.configure("Panel.TFrame", background=PANEL)
        style.configure(
            "TLabel",
            background=BG,
            foreground=FG,
            font=("Consolas", 10),
        )
        style.configure(
            "Title.TLabel",
            background=BG,
            foreground=ACCENT,
            font=("Consolas", 16, "bold"),
        )
        style.configure(
            "Muted.TLabel",
            background=BG,
            foreground=MUTED,
            font=("Consolas", 9),
        )
        style.configure(
            "TButton",
            background=PANEL,
            foreground=FG,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            relief="flat",
            padding=(10, 7),
            font=("Consolas", 9, "bold"),
        )
        style.map(
            "TButton",
            background=[("active", SELECT), ("pressed", FIELD)],
            foreground=[("active", ACCENT)],
        )
        style.configure(
            "Accent.TButton",
            background=SELECT,
            foreground=ACCENT,
            bordercolor=ACCENT,
        )
        style.map(
            "Accent.TButton",
            background=[("active", ACCENT)],
            foreground=[("active", BG)],
        )
        style.configure(
            "TEntry",
            fieldbackground=FIELD,
            foreground=FG,
            insertcolor=ACCENT,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            padding=7,
        )
        style.configure(
            "TCombobox",
            fieldbackground=FIELD,
            background=PANEL,
            foreground=FG,
            arrowcolor=ACCENT,
            bordercolor=BORDER,
            padding=5,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", FIELD)],
            foreground=[("readonly", FG)],
            selectbackground=[("readonly", FIELD)],
            selectforeground=[("readonly", FG)],
        )
        style.configure(
            "Treeview",
            background=FIELD,
            fieldbackground=FIELD,
            foreground=FG,
            bordercolor=BORDER,
            rowheight=26,
            font=("Consolas", 10),
        )
        style.map(
            "Treeview",
            background=[("selected", SELECT)],
            foreground=[("selected", ACCENT)],
        )
        style.configure(
            "Treeview.Heading",
            background=PANEL,
            foreground=ACCENT,
            bordercolor=BORDER,
            relief="flat",
            font=("Consolas", 10, "bold"),
        )
        style.map("Treeview.Heading", background=[("active", PANEL)])
        style.configure(
            "TCheckbutton",
            background=BG,
            foreground=FG,
            indicatorcolor=FIELD,
            indicatorrelief="flat",
            font=("Consolas", 9),
        )
        style.map(
            "TCheckbutton",
            background=[("active", BG)],
            foreground=[("active", ACCENT)],
            indicatorcolor=[("selected", ACCENT)],
        )
        style.configure(
            "TLabelframe",
            background=BG,
            foreground=ACCENT,
            bordercolor=BORDER,
            relief="solid",
        )
        style.configure(
            "TLabelframe.Label",
            background=BG,
            foreground=ACCENT,
            font=("Consolas", 9, "bold"),
        )

    def build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 10))

        ttk.Label(header, text="UNIT COST TERMINAL", style="Title.TLabel").pack(
            side="left"
        )
        ttk.Label(
            header,
            text="CTRL+ALT+U  SHOW/HIDE",
            style="Muted.TLabel",
        ).pack(side="right")

        filter_frame = ttk.Frame(outer)
        filter_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(filter_frame, text="SEARCH").grid(row=0, column=0, sticky="w")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(
            filter_frame,
            textvariable=self.search_var,
        )
        self.search_entry.grid(row=1, column=0, sticky="ew", padx=(0, 10))
        self.search_var.trace_add("write", lambda *_: self.apply_filters())

        ttk.Label(filter_frame, text="DIVISION").grid(row=0, column=1, sticky="w")
        self.division_var = tk.StringVar(value="ALL DIVISIONS")
        self.division_combo = ttk.Combobox(
            filter_frame,
            textvariable=self.division_var,
            state="readonly",
            width=22,
            values=["ALL DIVISIONS"],
        )
        self.division_combo.grid(row=1, column=1, sticky="ew", padx=(0, 10))
        self.division_combo.bind(
            "<<ComboboxSelected>>",
            lambda _event: self.apply_filters(),
        )

        ttk.Button(
            filter_frame,
            text="REFRESH",
            command=self.refresh_data,
        ).grid(row=1, column=2, sticky="ew")

        filter_frame.columnconfigure(0, weight=1)

        table_frame = ttk.Frame(outer)
        table_frame.pack(fill="both", expand=True)

        columns = ("division", "name", "cost", "unit", "published")
        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
        )
        self.tree.heading("division", text="DIV")
        self.tree.heading("name", text="UNIT COST")
        self.tree.heading("cost", text="COST")
        self.tree.heading("unit", text="UNIT")
        self.tree.heading("published", text="LAST PUBLISHED")

        self.tree.column("division", width=65, anchor="center", stretch=False)
        self.tree.column("name", width=430, anchor="w")
        self.tree.column("cost", width=105, anchor="e", stretch=False)
        self.tree.column("unit", width=70, anchor="center", stretch=False)
        self.tree.column("published", width=145, anchor="center", stretch=False)

        scrollbar = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.tree.yview,
        )
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", lambda _event: self.begin_choose_cell())

        comment_frame = ttk.LabelFrame(
            outer,
            text=" COMMENT / EXCEL NOTE ",
            padding=8,
        )
        comment_frame.pack(fill="x", pady=(10, 8))

        self.comment_box = tk.Text(
            comment_frame,
            height=6,
            wrap="word",
            bg=FIELD,
            fg=FG,
            insertbackground=ACCENT,
            selectbackground=SELECT,
            selectforeground=ACCENT,
            relief="flat",
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
            font=("Consolas", 10),
            padx=8,
            pady=8,
        )
        self.comment_box.pack(fill="both", expand=True)

        options = ttk.Frame(outer)
        options.pack(fill="x", pady=(0, 8))

        self.include_comment_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options,
            text="INCLUDE COMMENT IN EXCEL NOTE",
            variable=self.include_comment_var,
        ).pack(side="left")

        self.status_var = tk.StringVar(value="INITIALIZING...")
        ttk.Label(
            options,
            textvariable=self.status_var,
            style="Muted.TLabel",
        ).pack(side="right")

        actions = ttk.Frame(outer)
        actions.pack(fill="x")

        ttk.Button(
            actions,
            text="POPULATE EXCEL",
            style="Accent.TButton",
            command=self.begin_choose_cell,
        ).pack(side="left")

        ttk.Button(
            actions,
            text="CURRENT CELL",
            command=self.populate_current_cell,
        ).pack(side="left", padx=(6, 0))

        ttk.Button(
            actions,
            text="VIEW DETAILS",
            command=self.view_details,
        ).pack(side="left", padx=(6, 0))

        ttk.Button(
            actions,
            text="COPY COST",
            command=self.copy_cost,
        ).pack(side="left", padx=(16, 0))

        ttk.Button(
            actions,
            text="COPY FULL",
            command=self.copy_full,
        ).pack(side="left", padx=(6, 0))

        ttk.Button(
            actions,
            text="EXIT",
            command=self.exit_app,
        ).pack(side="right")

        ttk.Button(
            actions,
            text="HIDE",
            command=self.hide,
        ).pack(side="right", padx=(0, 6))

    def refresh_data(self) -> None:
        if self.loading:
            return

        self.loading = True
        self.status_var.set("REFRESHING...")
        self.root.update_idletasks()

        try:
            server_url = str(self.config["server_url"]).rstrip("/")

            if "YOUR-PI-IP" in server_url:
                raise RuntimeError("Set server_url in config.json.")

            response = requests.get(
                server_url + "/api/unit-cost-lookup",
                timeout=8,
            )
            response.raise_for_status()

            self.all_costs = [
                UnitCost(
                    id=int(row.get("id", 0)),
                    division=str(row.get("division", "")),
                    name=str(row.get("name", "")),
                    cost_per_unit=float(row.get("cost_per_unit", 0)),
                    output_unit=str(row.get("output_unit", "")),
                    published_at=str(row.get("published_at", "")),
                    comments=str(row.get("comments", "") or ""),
                    has_stale_material=bool(row.get("has_stale_material", False)),
                )
                for row in response.json()
            ]

            divisions = sorted(
                {str(item.division).zfill(2) for item in self.all_costs},
                key=lambda value: int(value) if value.isdigit() else value,
            )

            self.division_combo["values"] = ["ALL DIVISIONS"] + divisions

            if self.division_var.get() not in self.division_combo["values"]:
                self.division_var.set("ALL DIVISIONS")

            self.apply_filters()
            self.status_var.set(f"LOADED {len(self.all_costs)} UNIT COST(S)")

        except Exception as exc:
            self.status_var.set("REFRESH FAILED")
            messagebox.showerror(
                APP_NAME,
                f"Could not load unit costs.\n\n{exc}",
            )
        finally:
            self.loading = False

    def apply_filters(self) -> None:
        search = self.search_var.get().strip().lower()
        selected_division = self.division_var.get()

        self.visible_costs = []

        for item in self.all_costs:
            division = str(item.division).zfill(2)

            if (
                selected_division != "ALL DIVISIONS"
                and division != selected_division
            ):
                continue

            haystack = (
                f"{division} {item.name} {item.output_unit} {item.comments}"
            ).lower()

            if search and search not in haystack:
                continue

            self.visible_costs.append(item)

        self.refresh_tree()

    def refresh_tree(self) -> None:
        for child in self.tree.get_children():
            self.tree.delete(child)

        for index, item in enumerate(self.visible_costs):
            self.tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    str(item.division).zfill(2),
                    item.name,
                    f"${item.cost_per_unit:,.2f}",
                    item.output_unit,
                    self.format_publish_status(item),
                ),
            )

        self.selected_cost = None
        self.comment_box.delete("1.0", "end")
        self.status_var.set(f"{len(self.visible_costs)} MATCHES")

        if self.visible_costs:
            first = self.tree.get_children()[0]
            self.tree.selection_set(first)
            self.tree.focus(first)
            self.tree.see(first)
            self.on_select()

    def format_publish_status(self, item: UnitCost) -> str:
        published = str(item.published_at or "").strip()

        if published:
            published = published[:10]
        else:
            published = "NOT PUBLISHED"

        if item.has_stale_material:
            published += "  !"

        return published

    def on_select(self, _event=None) -> None:
        selection = self.tree.selection()

        if not selection:
            self.selected_cost = None
            self.comment_box.delete("1.0", "end")
            return

        index = int(selection[0])

        if index >= len(self.visible_costs):
            return

        self.selected_cost = self.visible_costs[index]
        self.comment_box.delete("1.0", "end")
        self.comment_box.insert("1.0", self.selected_cost.comments)

    def ensure_selection(self) -> bool:
        if not self.selected_cost:
            messagebox.showinfo(APP_NAME, "Select a unit cost first.")
            return False
        return True

    def view_details(self) -> None:
        if not self.ensure_selection():
            return

        try:
            server_url = str(self.config["server_url"]).rstrip("/")
            response = requests.get(
                f"{server_url}/api/unit-cost-lookup/{self.selected_cost.id}",
                timeout=8,
            )
            response.raise_for_status()
            details = response.json()
        except Exception as exc:
            messagebox.showerror(
                APP_NAME,
                f"Could not load unit-cost details.\n\n{exc}",
            )
            return

        window = tk.Toplevel(self.root)
        window.title(f"UNIT COST DETAILS :: {self.selected_cost.name}")
        window.geometry("980x760")
        window.minsize(820, 620)
        window.configure(bg=BG)

        outer = ttk.Frame(window, padding=14)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 10))

        ttk.Label(
            header,
            text=str(details.get("name", self.selected_cost.name)).upper(),
            style="Title.TLabel",
        ).pack(side="left")

        ttk.Label(
            header,
            text=(
                f"DIV {str(details.get('division', '')).zfill(2)}   "
                f"${float(details.get('cost_per_unit', 0)):,.2f} / "
                f"{details.get('output_unit', '')}"
            ),
            style="Title.TLabel",
        ).pack(side="right")

        meta = ttk.Frame(outer)
        meta.pack(fill="x", pady=(0, 10))

        metadata = [
            ("STATUS", details.get("status", "")),
            ("OUTPUT QTY", details.get("output_quantity", "")),
            ("MISC/BOND", f"{float(details.get('misc_bond_pct', 0)):g}%"),
            ("ESCALATION", f"{float(details.get('escalation_pct', 0)):g}%"),
            ("MARKUP", f"{float(details.get('markup_pct', 0)):g}%"),
            ("UPDATED", details.get("updated_at", "")),
        ]

        for index, (label, value) in enumerate(metadata):
            block = ttk.Frame(meta, style="Panel.TFrame", padding=7)
            block.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 4, 0))
            ttk.Label(block, text=label, style="Muted.TLabel").pack(anchor="w")
            ttk.Label(block, text=str(value)).pack(anchor="w")
            meta.columnconfigure(index, weight=1)

        notebook = ttk.Notebook(outer)
        notebook.pack(fill="both", expand=True)

        overview_tab = ttk.Frame(notebook, padding=10)
        materials_tab = ttk.Frame(notebook, padding=10)
        labor_tab = ttk.Frame(notebook, padding=10)
        history_tab = ttk.Frame(notebook, padding=10)

        notebook.add(overview_tab, text=" OVERVIEW ")
        notebook.add(materials_tab, text=" MATERIALS ")
        notebook.add(labor_tab, text=" LABOR ")
        notebook.add(history_tab, text=" PUBLICATIONS ")

        overview_tab.columnconfigure(0, weight=1)
        overview_tab.rowconfigure(1, weight=1)

        comments_frame = ttk.LabelFrame(
            overview_tab,
            text=" COMMENTS ",
            padding=8,
        )
        comments_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 8))

        comments_text = tk.Text(
            comments_frame,
            height=6,
            wrap="word",
            bg=FIELD,
            fg=FG,
            insertbackground=ACCENT,
            relief="flat",
            highlightthickness=1,
            highlightbackground=BORDER,
            font=("Consolas", 10),
            padx=8,
            pady=8,
        )
        comments_text.pack(fill="both", expand=True)
        comments_text.insert("1.0", str(details.get("comments", "") or ""))
        comments_text.configure(state="disabled")

        scratch_frame = ttk.LabelFrame(
            overview_tab,
            text=" CALCULATION SCRATCH ",
            padding=8,
        )
        scratch_frame.grid(row=1, column=0, sticky="nsew")

        scratch_text = tk.Text(
            scratch_frame,
            wrap="none",
            bg=FIELD,
            fg=FG,
            insertbackground=ACCENT,
            relief="flat",
            highlightthickness=1,
            highlightbackground=BORDER,
            font=("Consolas", 10),
            padx=8,
            pady=8,
        )
        scratch_y = ttk.Scrollbar(
            scratch_frame,
            orient="vertical",
            command=scratch_text.yview,
        )
        scratch_x = ttk.Scrollbar(
            scratch_frame,
            orient="horizontal",
            command=scratch_text.xview,
        )
        scratch_text.configure(
            yscrollcommand=scratch_y.set,
            xscrollcommand=scratch_x.set,
        )
        scratch_text.grid(row=0, column=0, sticky="nsew")
        scratch_y.grid(row=0, column=1, sticky="ns")
        scratch_x.grid(row=1, column=0, sticky="ew")
        scratch_frame.columnconfigure(0, weight=1)
        scratch_frame.rowconfigure(0, weight=1)
        scratch_text.insert("1.0", str(details.get("calc_scratch", "") or ""))
        scratch_text.configure(state="disabled")

        self.populate_detail_tree(
            materials_tab,
            columns=[
                ("order", "ORDER", 65, "center"),
                ("description", "DESCRIPTION", 350, "w"),
                ("unit", "UNIT", 75, "center"),
                ("price", "PRICE", 95, "e"),
                ("qty", "QTY", 90, "e"),
                ("multiplier", "MULT", 75, "e"),
                ("extended", "EXTENDED", 105, "e"),
                ("updated", "PRICE UPDATED", 135, "center"),
            ],
            rows=[
                (
                    row.get("line_order", ""),
                    row.get("description", ""),
                    row.get("purchase_unit", ""),
                    f"${float(row.get('price_per_unit') or 0):,.2f}",
                    f"{float(row.get('quantity') or 0):,.4g}",
                    f"{float(row.get('multiplier') or 0):,.4g}",
                    f"${float(row.get('extended') or 0):,.2f}",
                    self.format_material_update(row),
                )
                for row in details.get("materials", [])
            ],
        )

        self.populate_detail_tree(
            labor_tab,
            columns=[
                ("order", "ORDER", 65, "center"),
                ("type", "LABOR TYPE", 300, "w"),
                ("crew", "CREW", 80, "e"),
                ("hours", "HOURS", 90, "e"),
                ("total", "TOTAL HRS", 100, "e"),
                ("rate", "RATE", 95, "e"),
                ("extended", "EXTENDED", 105, "e"),
            ],
            rows=[
                (
                    row.get("line_order", ""),
                    row.get("labor_type", ""),
                    f"{float(row.get('crew_size') or 0):,.4g}",
                    f"{float(row.get('hours') or 0):,.4g}",
                    f"{float(row.get('hours_total') or 0):,.4g}",
                    f"${float(row.get('labor_rate') or 0):,.2f}",
                    f"${float(row.get('extended') or 0):,.2f}",
                )
                for row in details.get("labor", [])
            ],
        )

        self.populate_detail_tree(
            history_tab,
            columns=[
                ("date", "PUBLISHED", 155, "w"),
                ("project", "PROJECT", 280, "w"),
                ("estimator", "ESTIMATOR", 150, "w"),
                ("cost", "COST", 100, "e"),
                ("unit", "UNIT", 70, "center"),
                ("notes", "NOTES", 320, "w"),
            ],
            rows=[
                (
                    row.get("published_at", ""),
                    row.get("project_name", ""),
                    row.get("estimator", ""),
                    f"${float(row.get('cost_per_unit') or 0):,.2f}",
                    row.get("output_unit", ""),
                    row.get("notes", ""),
                )
                for row in details.get("publications", [])
            ],
        )

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(10, 0))

        ttk.Button(
            footer,
            text="CLOSE",
            command=window.destroy,
        ).pack(side="right")

    def format_material_update(self, row: dict) -> str:
        updated = str(row.get("date_updated", "") or "").strip()

        if updated:
            updated = updated[:10]
        else:
            updated = "UNKNOWN"

        if bool(row.get("is_stale", False)):
            updated += "  !"

        return updated

    def populate_detail_tree(self, parent, columns, rows) -> None:
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)

        keys = [item[0] for item in columns]
        tree = ttk.Treeview(
            frame,
            columns=keys,
            show="headings",
            selectmode="browse",
        )

        for key, heading, width, anchor in columns:
            tree.heading(key, text=heading)
            tree.column(
                key,
                width=width,
                anchor=anchor,
                stretch=key in {"description", "type", "project", "notes"},
            )

        scrollbar = ttk.Scrollbar(
            frame,
            orient="vertical",
            command=tree.yview,
        )
        tree.configure(yscrollcommand=scrollbar.set)

        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for row in rows:
            tree.insert("", "end", values=row)

    def begin_choose_cell(self) -> None:
        if not self.ensure_selection():
            return

        self.pending_insert = True
        self.hide()
        self.show_pending_popup()

    def show_pending_popup(self) -> None:
        self.pending_popup = tk.Toplevel(self.root)
        self.pending_popup.title(APP_NAME)
        self.pending_popup.geometry("470x155")
        self.pending_popup.configure(bg=BG)
        self.pending_popup.attributes("-topmost", True)
        self.pending_popup.resizable(False, False)

        frame = ttk.Frame(self.pending_popup, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text=(
                "SELECT A DESTINATION CELL IN EXCEL\n\n"
                "THEN PRESS  CTRL + ALT + ENTER"
            ),
            justify="center",
            style="Title.TLabel",
        ).pack(fill="x")

        ttk.Button(
            frame,
            text="CANCEL",
            command=self.cancel_pending,
        ).pack(pady=(14, 0))

        self.pending_popup.protocol("WM_DELETE_WINDOW", self.cancel_pending)

    def cancel_pending(self) -> None:
        self.pending_insert = False

        if (
            hasattr(self, "pending_popup")
            and self.pending_popup.winfo_exists()
        ):
            self.pending_popup.destroy()

        self.show()

    def populate_current_cell(self) -> None:
        if self.ensure_selection():
            self.write_to_excel()

    def write_to_excel(self) -> None:
        if not self.selected_cost:
            return

        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            cell = excel.ActiveCell

            if cell is None:
                raise RuntimeError("Excel does not have an active cell.")

            edited_comment = self.comment_box.get("1.0", "end").strip()

            note = (
                f"Unit cost: {self.selected_cost.name}\n"
                f"Unit: {self.selected_cost.output_unit}\n"
                f"Published: {self.selected_cost.published_at}"
            )

            if self.include_comment_var.get() and edited_comment:
                note += f"\n\nComments:\n{edited_comment}"

            cell.Value = float(self.selected_cost.cost_per_unit)
            cell.NumberFormat = str(
                self.config.get("currency_format", "$#,##0.00")
            )

            try:
                if cell.Comment is not None:
                    cell.Comment.Delete()
            except Exception:
                pass

            cell.AddComment(note)
            excel.Visible = True

            self.pending_insert = False

            if (
                hasattr(self, "pending_popup")
                and self.pending_popup.winfo_exists()
            ):
                self.pending_popup.destroy()

            self.status_var.set(
                f"INSERTED {self.selected_cost.name} INTO {cell.Address}"
            )

        except Exception as exc:
            messagebox.showerror(
                APP_NAME,
                "Could not populate Excel.\n\n"
                "Make sure desktop Excel is open and a cell is selected.\n\n"
                f"{exc}",
            )

    def set_clipboard(self, text: str) -> None:
        win32clipboard.OpenClipboard()

        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text)
        finally:
            win32clipboard.CloseClipboard()

        self.status_var.set("COPIED TO CLIPBOARD")

    def copy_cost(self) -> None:
        if self.ensure_selection():
            self.set_clipboard(f"{self.selected_cost.cost_per_unit:.2f}")

    def copy_full(self) -> None:
        if not self.ensure_selection():
            return

        edited_comment = self.comment_box.get("1.0", "end").strip()

        text = (
            f"{self.selected_cost.name}\n"
            f"Division: {str(self.selected_cost.division).zfill(2)}\n"
            f"Cost: ${self.selected_cost.cost_per_unit:,.2f} / "
            f"{self.selected_cost.output_unit}\n"
            f"Published: {self.selected_cost.published_at}"
        )

        if edited_comment:
            text += f"\nComments: {edited_comment}"

        self.set_clipboard(text)

    def toggle_visibility(self) -> None:
        try:
            visible = bool(self.root.winfo_viewable())
        except tk.TclError:
            visible = False

        if visible:
            self.hide()
        else:
            self.show()

    def show(self) -> None:
        self.root.deiconify()
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(
            150,
            lambda: self.root.attributes("-topmost", False),
        )
        self.search_entry.focus_set()

    def hide(self) -> None:
        self.root.withdraw()

    def process_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()

                if event == "toggle":
                    self.toggle_visibility()

                elif event == "insert" and self.pending_insert:
                    self.write_to_excel()

                elif event == "toggle_failed":
                    messagebox.showwarning(
                        APP_NAME,
                        "Ctrl+Alt+U could not be registered.\n\n"
                        "Another application may already be using it.",
                    )

                elif event == "insert_failed":
                    messagebox.showwarning(
                        APP_NAME,
                        "Ctrl+Alt+Enter could not be registered.\n\n"
                        "Another application may already be using it.",
                    )

        except queue.Empty:
            pass

        self.root.after(100, self.process_events)

    def exit_app(self) -> None:
        self.hotkeys.stop()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    app = UnitCostLookupApp(root)
    app.show()
    root.mainloop()


if __name__ == "__main__":
    main()
