from __future__ import annotations

import json
import queue
import re
import threading
import tkinter as tk
import xml.etree.ElementTree as ET
import sys
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from zipfile import ZipFile

import ttkbootstrap as tb

from scheduler_engine import SchedulerError, solve_all_scenarios
from timetable_tool import read_json, run_solve, write_json

DAY_ORDER = ["Mon", "Tue", "Wed", "Thu", "Fri"]
DAY_JP = {"Mon": "月", "Tue": "火", "Wed": "水", "Thu": "木", "Fri": "金"}
MATRIX_PERIODS = [1, 2, 3, 4, 5, 6]
DEFAULT_SUBJECTS = ["国語", "数学", "英語", "理科", "社会", "美術", "音楽","技術", "家庭", "保体", "学活", "道徳", "総合"]
TEACHER_SUBJECT_OPTIONS = ["国語", "数学", "英語", "理科", "社会", "保体", "美術", "音楽", "技術", "家庭", "個別"]
TECH_SUBJECTS = ["音楽", "美術", "音美", "技術", "家庭", "技家", "保体", "学活", "道徳", "総合"]
AUTO_EXEMPT_SUBJECTS = list(TECH_SUBJECTS) + ["個別"]
SKILL_SUBJECT_KEYWORDS = set(AUTO_EXEMPT_SUBJECTS) | {"体育", "保健体育"}
SCENARIO_ID_UPPER = "1grade-music-art"
SCENARIO_ID_LOWER = "1grade-general"
LEGACY_SCENARIO_ID_MAP = {
    "r6_upper": SCENARIO_ID_UPPER,
    "r6_lower": SCENARIO_ID_LOWER,
    "１年音美": SCENARIO_ID_UPPER,
    "１年総合": SCENARIO_ID_LOWER,
}
# 技能科目シート由来の固定色
SUBJECT_COLOR_MAP = {
    "国語": "#ffcccc",
    "数学": "#ccffff",
    "英語": "#4472c4",
    "理科": "#ffff00",
    "社会": "#ff7c80",
    "美術": "#9999ff",
    "音楽": "#ff99ff",
    "音美": "#cc99ff",
    "技術": "#92d050",
    "家庭": "#92d050",
    "技家": "#92d050",
    "保体": "#ffc000",
    "学活": "#FFF2CC",
    "道徳": "#FFF2CC",
    "総合": "#FFF2CC",
}


def _normalize_subject_name(subject: str) -> str:
    return str(subject or "").replace("　", "").replace(" ", "").strip()


def _is_skill_subject_name(subject: str) -> bool:
    normalized = _normalize_subject_name(subject)
    if not normalized:
        return False
    if normalized in SKILL_SUBJECT_KEYWORDS:
        return True
    return any(keyword in normalized for keyword in SKILL_SUBJECT_KEYWORDS)

DAY_SEPARATOR_COLOR = "#000000"
DAY_SEPARATOR_WIDTH = 3
EMPTY_TUE_THU_COLOR = "#f2f2f2"
MATRIX_CELL_WIDTH = 4
TT_HIGHLIGHT_BORDER_COLOR = "#d32f2f"
TT_DEFAULT_BORDER_COLOR = "#000000"
TT_HIGHLIGHT_BORDER_WIDTH = 2


def _normalize_scenario_id(sid: str) -> str:
    key = sid.strip()
    return LEGACY_SCENARIO_ID_MAP.get(key, key)


def _scenario_id_to_block(sid: str) -> str:
    normalized = _normalize_scenario_id(sid)
    if normalized == SCENARIO_ID_UPPER:
        return "upper"
    if normalized == SCENARIO_ID_LOWER:
        return "lower"
    return "upper" if normalized.endswith("upper") else "lower"


def split_csv(text: str) -> list[str]:
    if not text.strip():
        return []
    return [x.strip() for x in text.split(",") if x.strip()]


def join_csv(values: list[str]) -> str:
    return ", ".join(values)


class EntryTableGrid(ttk.Frame):
    def __init__(
        self,
        parent: tk.Misc,
        title: str,
        columns: list[tuple[str, str]],
        rows: int = 120,
        cell_width: int = 8,
        dropdown_values: dict[str, list[str]] | None = None,
    ):
        super().__init__(parent)
        self.columns = columns
        self.rows = rows
        self.cell_width = cell_width
        self.dropdown_values = dropdown_values or {}
        self.vars: list[list[tk.StringVar]] = []
        self.widgets: list[list[tk.Widget]] = []
        self.cell_widgets: dict[tuple[int, int], tk.Widget] = {}
        self.combo_style_name = "TeacherGrid.TCombobox"
        st = ttk.Style()
        st.configure(self.combo_style_name, fieldbackground="white", background="white", foreground="black")
        st.map(
            self.combo_style_name,
            fieldbackground=[("readonly", "white"), ("focus", "white"), ("active", "white"), ("!disabled", "white")],
            background=[("readonly", "white"), ("focus", "white"), ("active", "white"), ("!disabled", "white")],
            foreground=[("readonly", "black"), ("focus", "black"), ("active", "black"), ("!disabled", "black"), ("disabled", "black")],
            selectbackground=[("readonly", "white"), ("focus", "white"), ("active", "white"), ("!disabled", "white")],
            selectforeground=[("readonly", "black"), ("focus", "black"), ("active", "black"), ("!disabled", "black"), ("disabled", "black")],
        )

        ttk.Label(self, text=title).pack(anchor="w")

        outer = ttk.Frame(self)
        outer.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(outer, bg="white")
        ybar = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        xbar = ttk.Scrollbar(outer, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)

        ybar.pack(side="right", fill="y")
        xbar.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.grid_frame = tk.Frame(self.canvas, bg="white")
        self.window_id = self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
        self.grid_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Enter>", lambda _e: self._bind_mousewheel())
        self.canvas.bind("<Leave>", lambda _e: self._unbind_mousewheel())

        self._build_grid()

    def _on_frame_configure(self, _evt=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event) -> None:
        self.canvas.itemconfig(self.window_id, width=max(event.width, 800))

    def _bind_mousewheel(self) -> None:
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)

    def _unbind_mousewheel(self) -> None:
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Shift-MouseWheel>")

    def _on_mousewheel(self, event) -> str:
        delta = -1 * int(event.delta / 120) if event.delta else 0
        if delta:
            self.canvas.yview_scroll(delta, "units")
        return "break"

    def _on_shift_mousewheel(self, event) -> str:
        delta = -1 * int(event.delta / 120) if event.delta else 0
        if delta:
            self.canvas.xview_scroll(delta, "units")
        return "break"

    def _header(self, text: str, r: int, c: int, w: int = 8) -> None:
        tk.Label(
            self.grid_frame,
            text=text,
            width=w,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground="black",
            bg="white",
            fg="black",
            font=("Meiryo UI", 8),
        ).grid(row=r, column=c, sticky="nsew")

    def _build_grid(self) -> None:
        for w in self.grid_frame.winfo_children():
            w.destroy()
        self.vars.clear()
        self.widgets.clear()
        self.cell_widgets.clear()

        for c, (_key, label) in enumerate(self.columns):
            self._header(label, 0, c, w=max(6, min(14, self.cell_width + 2)))

        for r in range(self.rows):
            row_vars: list[tk.StringVar] = []
            row_widgets: list[tk.Widget] = []
            for c in range(len(self.columns)):
                v = tk.StringVar(value="")
                row_vars.append(v)
                key, _ = self.columns[c]
                if key in self.dropdown_values:
                    w = ttk.Combobox(
                        self.grid_frame,
                        textvariable=v,
                        width=self.cell_width,
                        values=self.dropdown_values.get(key, []),
                        state="readonly",
                        font=("Meiryo UI", 8),
                        style=self.combo_style_name,
                        foreground="black",
                    )
                else:
                    w = tk.Entry(
                        self.grid_frame,
                        textvariable=v,
                        width=self.cell_width,
                        relief="flat",
                        bd=0,
                        highlightthickness=1,
                        highlightbackground="black",
                        highlightcolor="black",
                        bg="white",
                        fg="black",
                        insertbackground="black",
                        justify="left",
                        font=("Meiryo UI", 8),
                    )
                w.grid(row=r + 1, column=c, sticky="nsew")
                row_widgets.append(w)
                self.cell_widgets[(r, c)] = w
                self._bind_arrow_nav(w, r, c)
            self.vars.append(row_vars)
            self.widgets.append(row_widgets)

    def set_rows(self, rows: list[dict]) -> None:
        if len(rows) > self.rows:
            self.rows = len(rows)
            self._build_grid()
        for r in range(self.rows):
            for c in range(len(self.columns)):
                self.vars[r][c].set("")
        for r, row in enumerate(rows[: self.rows]):
            for c, (key, _label) in enumerate(self.columns):
                self.vars[r][c].set(str(row.get(key, "")))

    def get_rows(self, drop_empty: bool = True) -> list[dict]:
        out = []
        for r in range(self.rows):
            row: dict[str, str] = {}
            for c, (key, _label) in enumerate(self.columns):
                row[key] = self.vars[r][c].get().strip()
            if drop_empty and all(not v for v in row.values()):
                continue
            out.append(row)
        return out

    def set_dropdown_values(self, key: str, values: list[str]) -> None:
        self.dropdown_values[key] = values
        col_idx = None
        for i, (k, _label) in enumerate(self.columns):
            if k == key:
                col_idx = i
                break
        if col_idx is None:
            return
        for r in range(len(self.widgets)):
            w = self.widgets[r][col_idx]
            if isinstance(w, ttk.Combobox):
                w.configure(values=values)

    def _bind_arrow_nav(self, widget: tk.Widget, r: int, c: int) -> None:
        widget.bind("<Up>", lambda e, rr=r, cc=c: self._move_focus(rr - 1, cc))
        widget.bind("<Down>", lambda e, rr=r, cc=c: self._move_focus(rr + 1, cc))
        widget.bind("<Left>", lambda e, rr=r, cc=c: self._move_focus(rr, cc - 1))
        widget.bind("<Right>", lambda e, rr=r, cc=c: self._move_focus(rr, cc + 1))
        widget.bind("<FocusIn>", lambda _e, w=widget: self._ensure_visible(w))
        widget.bind("<Button-1>", lambda _e, w=widget: self._ensure_visible(w))

    def _move_focus(self, r: int, c: int) -> str:
        w = self.cell_widgets.get((r, c))
        if w is not None:
            w.focus_set()
            self._ensure_visible(w)
        return "break"

    def _ensure_visible(self, widget: tk.Widget) -> None:
        self.update_idletasks()
        wx, wy = widget.winfo_x(), widget.winfo_y()
        ww, wh = widget.winfo_width(), widget.winfo_height()
        vx0, vy0 = self.canvas.canvasx(0), self.canvas.canvasy(0)
        vx1, vy1 = vx0 + self.canvas.winfo_width(), vy0 + self.canvas.winfo_height()
        cw = max(self.grid_frame.winfo_reqwidth(), 1)
        ch = max(self.grid_frame.winfo_reqheight(), 1)

        if wx < vx0:
            self.canvas.xview_moveto(max(0.0, wx / cw))
        elif wx + ww > vx1:
            self.canvas.xview_moveto(max(0.0, (wx + ww - self.canvas.winfo_width()) / cw))

        if wy < vy0:
            self.canvas.yview_moveto(max(0.0, wy / ch))
        elif wy + wh > vy1:
            self.canvas.yview_moveto(max(0.0, (wy + wh - self.canvas.winfo_height()) / ch))

    def dump_cells(self) -> dict[tuple[str, str], tuple[str, bool]]:
        out: dict[tuple[str, str], tuple[str, bool]] = {}
        for key, var in self.vars.items():
            hours = var.get().strip()
            tt = key in self.tt_marks
            if hours or tt:
                out[key] = (hours, tt)
        return out

    def set_cells(self, data: dict[tuple[str, str], tuple[str, bool]]) -> None:
        for key, var in self.vars.items():
            var.set("")
        self.tt_marks.clear()
        for key, (hours, tt) in data.items():
            if key in self.vars:
                self.vars[key].set(hours)
                if tt:
                    self.tt_marks.add(key)
        for key in self.entries:
            self._apply_cell_style(key)


class TimetableMatrixGrid(ttk.Frame):
    _instance_seq = 0

    def __init__(
        self,
        parent: tk.Misc,
        title: str,
        row_labels: list[str],
        click_x_only: bool = False,
        day_periods: dict[str, int] | None = None,
        subject_values: list[str] | None = None,
        subject_colors: dict[str, str] | None = None,
        readonly: bool = False,
        merge_day_header: bool = False,
        show_day_separators: bool = False,
        readonly_cell_scale: float = 1.0,
    ):
        super().__init__(parent)
        self.row_labels = list(row_labels)
        self.click_x_only = click_x_only
        self.readonly = readonly
        self.merge_day_header = merge_day_header
        self.show_day_separators = show_day_separators
        self.day_periods = day_periods or {"Mon": 5, "Tue": 6, "Wed": 5, "Thu": 6, "Fri": 6}
        self.slot_keys = self._build_slot_keys()
        self.vars: dict[tuple[str, str], tk.StringVar] = {}
        self.cell_widgets: dict[tuple[int, int], tk.Widget] = {}
        self.key_widgets: dict[tuple[str, str], tk.Widget] = {}
        self.combo_styles: dict[tk.Widget, str] = {}
        self.subject_values = list(subject_values or [])
        self.subject_colors = dict(subject_colors or {})
        self._combo_style_counter = 0
        self._instance_id = TimetableMatrixGrid._instance_seq
        TimetableMatrixGrid._instance_seq += 1
        self.readonly_cell_scale = max(1.0, float(readonly_cell_scale))
        self.tt_highlight_keys: set[tuple[str, str]] = set()

        ttk.Label(self, text=title).pack(anchor="w")

        outer = ttk.Frame(self)
        outer.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(outer, bg="white")
        ybar = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        xbar = ttk.Scrollbar(outer, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)

        ybar.pack(side="right", fill="y")
        xbar.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.grid_frame = tk.Frame(self.canvas, bg="white")
        self.window_id = self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
        self.grid_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Enter>", lambda _e: self._bind_mousewheel())
        self.canvas.bind("<Leave>", lambda _e: self._unbind_mousewheel())

        self._build_grid()

    def _build_slot_keys(self) -> list[str]:
        keys = []
        for d in DAY_ORDER:
            for p in range(1, int(self.day_periods.get(d, 6)) + 1):
                keys.append(f"{d}-{p}")
        return keys

    def _on_frame_configure(self, _evt=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event) -> None:
        self.canvas.itemconfig(self.window_id, width=max(event.width, 800))

    def _bind_mousewheel(self) -> None:
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)

    def _unbind_mousewheel(self) -> None:
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Shift-MouseWheel>")

    def _on_mousewheel(self, event) -> str:
        delta = -1 * int(event.delta / 120) if event.delta else 0
        if delta:
            self.canvas.yview_scroll(delta, "units")
        return "break"

    def _on_shift_mousewheel(self, event) -> str:
        delta = -1 * int(event.delta / 120) if event.delta else 0
        if delta:
            self.canvas.xview_scroll(delta, "units")
        return "break"

    def _header_label(self, text: str, row: int, col: int, width: int = 3, columnspan: int = 1) -> None:
        tk.Label(
            self.grid_frame,
            text=text,
            width=width,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground="black",
            bg="white",
            fg="black",
            font=("Meiryo UI", 8),
        ).grid(row=row, column=col, columnspan=max(1, int(columnspan)), sticky="nsew")

    def _add_day_separator(self, row: int, col: int) -> None:
        sep = tk.Frame(
            self.grid_frame,
            bg=DAY_SEPARATOR_COLOR,
            width=DAY_SEPARATOR_WIDTH,
            bd=0,
            highlightthickness=0,
        )
        sep.grid(row=row, column=col, sticky="nsew")
        self.grid_frame.grid_columnconfigure(col, minsize=DAY_SEPARATOR_WIDTH)

    def _build_grid(self) -> None:
        for w in self.grid_frame.winfo_children():
            w.destroy()
        self.vars.clear()
        self.cell_widgets.clear()
        self.key_widgets.clear()
        self.combo_styles.clear()
        self.slot_keys = self._build_slot_keys()
        old_tt_marks = set(self.tt_highlight_keys)

        day_items = [(day, int(self.day_periods.get(day, 6))) for day in DAY_ORDER]
        last_day_idx = len(day_items) - 1
        separator_cols: list[int] = []

        if self.merge_day_header:
            # 曜日は1セルに結合、左端の「組」は1行下へ配置
            self._header_label("", 0, 0, width=6)
            self._header_label("組", 1, 0, width=6)
            col = 1
            for day_idx, (day, day_cols) in enumerate(day_items):
                self._header_label(DAY_JP[day], 0, col, width=max(3, day_cols * 2), columnspan=day_cols)
                for p in range(1, day_cols + 1):
                    self._header_label(str(p), 1, col, width=3)
                    col += 1
                if self.show_day_separators and day_idx < last_day_idx:
                    separator_cols.append(col)
                    col += 1
        else:
            self._header_label("組", 0, 0, width=6)
            col = 1
            for day_idx, (day, day_cols) in enumerate(day_items):
                for p in range(1, day_cols + 1):
                    day_txt = DAY_JP[day] if p == 1 else ""
                    self._header_label(day_txt, 0, col, width=3)
                    self._header_label(str(p), 1, col, width=3)
                    col += 1
                if self.show_day_separators and day_idx < last_day_idx:
                    separator_cols.append(col)
                    col += 1

        for ridx, cls in enumerate(self.row_labels, start=2):
            self._header_label(cls, ridx, 0, width=6)
            col = 1
            logical_row = ridx - 2
            logical_col = 0
            for day_idx, (day, day_cols) in enumerate(day_items):
                for p in range(1, day_cols + 1):
                    slot = f"{day}-{p}"
                    var = tk.StringVar(value="")
                    self.vars[(cls, slot)] = var

                    if self.click_x_only:
                        ent = tk.Entry(
                            self.grid_frame,
                            textvariable=var,
                            width=3,
                            relief="flat",
                            bd=0,
                            highlightthickness=1,
                            highlightbackground="black",
                            highlightcolor="black",
                            readonlybackground="white",
                            bg="white",
                            fg="black",
                            disabledforeground="black",
                            justify="center",
                            state="readonly",
                            font=("Meiryo UI", 8),
                        )
                        ent.grid(row=ridx, column=col, sticky="nsew")
                        self.cell_widgets[(logical_row, logical_col)] = ent
                        self.key_widgets[(cls, slot)] = ent
                        self._bind_arrow_nav(ent, logical_row, logical_col)

                        def _toggle(_evt, key=(cls, slot)):
                            cur = self.vars[key].get().strip()
                            self.vars[key].set("" if cur else "✕")
                            self._apply_x_cell_color(key)

                        ent.bind("<Button-1>", _toggle)
                    elif self.subject_values:
                        style_name = f"GridCell_{self._instance_id}_{self._combo_style_counter}.TCombobox"
                        self._combo_style_counter += 1
                        style = ttk.Style()
                        style.configure(style_name, fieldbackground="white", background="white", foreground="black")
                        style.map(
                            style_name,
                            fieldbackground=[("readonly", "white"), ("focus", "white"), ("active", "white"), ("!disabled", "white")],
                            background=[("readonly", "white"), ("focus", "white"), ("active", "white"), ("!disabled", "white")],
                            foreground=[("readonly", "black"), ("focus", "black"), ("active", "black"), ("!disabled", "black"), ("disabled", "black")],
                            selectbackground=[("readonly", "white"), ("focus", "white"), ("active", "white"), ("!disabled", "white")],
                            selectforeground=[("readonly", "black"), ("focus", "black"), ("active", "black"), ("!disabled", "black"), ("disabled", "black")],
                        )
                        cmb = ttk.Combobox(
                            self.grid_frame,
                            textvariable=var,
                            width=MATRIX_CELL_WIDTH,
                            values=[""] + self.subject_values,
                            state="readonly",
                            justify="center",
                            font=("Meiryo UI", 8),
                            style=style_name,
                            foreground="black",
                        )
                        cmb.grid(row=ridx, column=col, sticky="nsew")
                        self.cell_widgets[(logical_row, logical_col)] = cmb
                        self.combo_styles[cmb] = style_name
                        self._bind_arrow_nav(cmb, logical_row, logical_col)
                        var.trace_add("write", lambda *_a, key=(cls, slot): self._apply_subject_cell_color(key))
                        cmb.bind("<<ComboboxSelected>>", lambda _e, key=(cls, slot): self._apply_subject_cell_color(key))
                        cmb.bind("<FocusIn>", lambda _e, w=cmb: self._ensure_visible(w))
                        cmb.bind("<Button-1>", lambda _e, w=cmb: self._ensure_visible(w))
                    elif self.readonly:
                        readonly_width = max(1, int(round(MATRIX_CELL_WIDTH * self.readonly_cell_scale)))
                        readonly_ipady = max(0, int(round((self.readonly_cell_scale - 1.0) * 8)))
                        lbl = tk.Label(
                            self.grid_frame,
                            textvariable=var,
                            width=readonly_width,
                            relief="flat",
                            bd=0,
                            highlightthickness=1,
                            highlightbackground="black",
                            bg="white",
                            fg="black",
                            font=("Meiryo UI", 8),
                        )
                        lbl.grid(row=ridx, column=col, sticky="nsew", ipady=readonly_ipady)
                        self.cell_widgets[(logical_row, logical_col)] = lbl
                    else:
                        ent = tk.Entry(
                            self.grid_frame,
                            textvariable=var,
                            width=3,
                            justify="center",
                            relief="flat",
                            bd=0,
                            highlightthickness=1,
                            highlightbackground="black",
                            highlightcolor="black",
                            bg="white",
                            fg="black",
                            insertbackground="black",
                            font=("Meiryo UI", 8),
                        )
                        ent.grid(row=ridx, column=col, sticky="nsew")
                        self.cell_widgets[(logical_row, logical_col)] = ent
                        self._bind_arrow_nav(ent, logical_row, logical_col)
                    col += 1
                    logical_col += 1
                if self.show_day_separators and day_idx < last_day_idx:
                    col += 1

        if self.show_day_separators and separator_cols:
            total_rows = 2 + len(self.row_labels)
            for sep_col in separator_cols:
                for rr in range(total_rows):
                    self._add_day_separator(rr, sep_col)

        for key in self.vars:
            self._apply_subject_cell_color(key)

    def set_row_labels(self, row_labels: list[str]) -> None:
        self.set_layout(row_labels=row_labels)

    def set_day_periods(self, day_periods: dict[str, int]) -> None:
        self.set_layout(day_periods=day_periods)

    def set_layout(
        self,
        row_labels: list[str] | None = None,
        day_periods: dict[str, int] | None = None,
    ) -> None:
        new_row_labels = self.row_labels if row_labels is None else list(row_labels)
        new_day_periods = self.day_periods if day_periods is None else dict(day_periods)
        if new_row_labels == self.row_labels and new_day_periods == self.day_periods:
            return
        old = self.get_value_map()
        self.row_labels = new_row_labels
        self.day_periods = new_day_periods
        self._build_grid()
        self.set_value_map(old)

    def set_subject_values(self, values: list[str], colors: dict[str, str] | None = None) -> None:
        new_values = list(values)
        new_colors = dict(colors) if colors is not None else None
        if new_values == self.subject_values and (new_colors is None or new_colors == self.subject_colors):
            if new_colors is not None:
                self.subject_colors = new_colors
                self.refresh_subject_colors()
            return
        old = self.get_value_map()
        self.subject_values = new_values
        if new_colors is not None:
            self.subject_colors = new_colors
        self._build_grid()
        self.set_value_map(old)

    def get_value_map(self) -> dict[tuple[str, str], str]:
        out: dict[tuple[str, str], str] = {}
        for key, var in self.vars.items():
            v = var.get().strip()
            if v:
                out[key] = v
        return out

    def set_value_map(self, values: dict[tuple[str, str], str]) -> None:
        for k, v in values.items():
            if k in self.vars:
                self.vars[k].set(v)
        self.refresh_subject_colors()

    def clear(self) -> None:
        for v in self.vars.values():
            v.set("")
        self.tt_highlight_keys.clear()
        if self.click_x_only:
            for key in self.key_widgets:
                self._apply_x_cell_color(key)
        for key in self.vars:
            self._apply_subject_cell_color(key)

    def set_from_assignments(self, assignments: list[dict]) -> None:
        self.clear()
        for a in assignments:
            cls = a.get("class", "").strip()
            slot = a.get("slot", "").strip()
            subj = a.get("subject", "").strip()
            k = (cls, slot)
            if k in self.vars:
                self.vars[k].set(subj)
        self.refresh_subject_colors()

    def refresh_subject_colors(self) -> None:
        for key in self.vars:
            self._apply_subject_cell_color(key)

    @staticmethod
    def _strip_tt_suffix(subject_text: str) -> str:
        text = str(subject_text or "").strip()
        return text[:-3].strip() if text.endswith("(T)") else text

    @staticmethod
    def _with_tt_suffix(subject_text: str) -> str:
        base = TimetableMatrixGrid._strip_tt_suffix(subject_text)
        return f"{base}(T)" if base else base

    def dump_assignments(self) -> list[dict]:
        out = []
        for (cls, slot), var in self.vars.items():
            text = self._strip_tt_suffix(var.get().strip())
            if text:
                out.append({"class": cls, "slot": slot, "subject": text})
        return out

    def set_x_marks(self, marks: set[tuple[str, str]]) -> None:
        self.clear()
        for key in marks:
            if key in self.vars:
                self.vars[key].set("✕")
                self._apply_x_cell_color(key)

    def dump_x_marks(self) -> set[tuple[str, str]]:
        out: set[tuple[str, str]] = set()
        for key, var in self.vars.items():
            if var.get().strip() == "✕":
                out.add(key)
        return out

    def set_tt_highlights(self, marks: set[tuple[str, str]]) -> None:
        self.tt_highlight_keys = {key for key in marks if key in self.vars}
        self.refresh_subject_colors()

    def clear_tt_highlights(self) -> None:
        self.tt_highlight_keys.clear()
        self.refresh_subject_colors()

    def _resolve_empty_cell_color(self, slot: str) -> str:
        day = slot.split("-", 1)[0] if slot else ""
        return EMPTY_TUE_THU_COLOR if day in {"Tue", "Thu"} else "#ffffff"

    def _apply_x_cell_color(self, key: tuple[str, str]) -> None:
        w = self.key_widgets.get(key)
        if w is None:
            return
        cls, slot = key
        is_x = self.vars.get((cls, slot)).get().strip() == "✕" if (cls, slot) in self.vars else False
        color = "#fde2e2" if is_x else self._resolve_empty_cell_color(slot)  # ✕は薄赤、空セルは曜日で既定色
        try:
            w.configure(bg=color, readonlybackground=color)
        except Exception:
            try:
                w.configure(bg=color)
            except Exception:
                pass

    def _apply_subject_cell_color(self, key: tuple[str, str]) -> None:
        if self.click_x_only:
            self._apply_x_cell_color(key)
            return
        row_labels = self.row_labels
        cls, slot = key
        if cls not in row_labels:
            return
        row = row_labels.index(cls)
        slot_list = self.slot_keys
        if slot not in slot_list:
            return
        col = slot_list.index(slot)
        widget = self.cell_widgets.get((row, col))
        if widget is None:
            return
        subj = self._strip_tt_suffix(self.vars.get(key).get().strip() if key in self.vars else "")
        color = self.subject_colors.get(subj) if subj else self._resolve_empty_cell_color(slot)
        if not color:
            color = "#ffffff"
        try:
            if isinstance(widget, ttk.Combobox):
                style_name = self.combo_styles.get(widget)
                if style_name:
                    style = ttk.Style()
                    style.configure(style_name, fieldbackground=color, background=color, foreground="black")
                    style.map(
                        style_name,
                        fieldbackground=[("readonly", color), ("focus", color), ("active", color)],
                        background=[("readonly", color), ("focus", color), ("active", color)],
                        foreground=[("readonly", "black"), ("focus", "black"), ("active", "black")],
                        selectbackground=[("readonly", color), ("focus", color), ("active", color)],
                        selectforeground=[("readonly", "black"), ("focus", "black"), ("active", "black")],
                    )
                    widget.configure(style=style_name)
            else:
                widget.configure(bg=color, activebackground=color, highlightbackground="#000000", highlightcolor="#000000", highlightthickness=1)
        except Exception:
            try:
                widget.configure(bg=color)
            except Exception:
                pass

    def _bind_arrow_nav(self, widget: tk.Widget, r: int, c: int) -> None:
        widget.bind("<Up>", lambda e, rr=r, cc=c: self._move_focus(rr - 1, cc))
        widget.bind("<Down>", lambda e, rr=r, cc=c: self._move_focus(rr + 1, cc))
        widget.bind("<Left>", lambda e, rr=r, cc=c: self._move_focus(rr, cc - 1))
        widget.bind("<Right>", lambda e, rr=r, cc=c: self._move_focus(rr, cc + 1))
        widget.bind("<FocusIn>", lambda _e, w=widget: self._ensure_visible(w))
        widget.bind("<Button-1>", lambda _e, w=widget: self._ensure_visible(w))

    def _move_focus(self, r: int, c: int) -> str:
        w = self.cell_widgets.get((r, c))
        if w is not None:
            w.focus_set()
            self._ensure_visible(w)
        return "break"

    def _ensure_visible(self, widget: tk.Widget) -> None:
        self.update_idletasks()
        wx, wy = widget.winfo_x(), widget.winfo_y()
        ww, wh = widget.winfo_width(), widget.winfo_height()
        vx0, vy0 = self.canvas.canvasx(0), self.canvas.canvasy(0)
        vx1, vy1 = vx0 + self.canvas.winfo_width(), vy0 + self.canvas.winfo_height()
        cw = max(self.grid_frame.winfo_reqwidth(), 1)
        ch = max(self.grid_frame.winfo_reqheight(), 1)

        if wx < vx0:
            self.canvas.xview_moveto(max(0.0, wx / cw))
        elif wx + ww > vx1:
            self.canvas.xview_moveto(max(0.0, (wx + ww - self.canvas.winfo_width()) / cw))

        if wy < vy0:
            self.canvas.yview_moveto(max(0.0, wy / ch))
        elif wy + wh > vy1:
            self.canvas.yview_moveto(max(0.0, (wy + wh - self.canvas.winfo_height()) / ch))


class ClassAssignGrid(ttk.Frame):
    def __init__(self, parent: tk.Misc, title: str, row_labels: list[str], col_labels: list[str]):
        super().__init__(parent)
        self.row_labels = list(row_labels)
        self.col_labels = list(col_labels)
        self.vars: dict[tuple[str, str], tk.StringVar] = {}
        self.entries: dict[tuple[str, str], tk.Entry] = {}
        self.tt_marks: set[tuple[str, str]] = set()
        self.cell_widgets: dict[tuple[int, int], tk.Entry] = {}

        ttk.Label(self, text=title).pack(anchor="w")

        outer = ttk.Frame(self)
        outer.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(outer, bg="white")
        ybar = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        xbar = ttk.Scrollbar(outer, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        ybar.pack(side="right", fill="y")
        xbar.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.grid_frame = tk.Frame(self.canvas, bg="white")
        self.window_id = self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
        self.grid_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Enter>", lambda _e: self._bind_mousewheel())
        self.canvas.bind("<Leave>", lambda _e: self._unbind_mousewheel())
        self._build_grid()

    def _on_frame_configure(self, _evt=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event) -> None:
        self.canvas.itemconfig(self.window_id, width=max(event.width, 800))

    def _bind_mousewheel(self) -> None:
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)

    def _unbind_mousewheel(self) -> None:
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Shift-MouseWheel>")

    def _on_mousewheel(self, event) -> str:
        delta = -1 * int(event.delta / 120) if event.delta else 0
        if delta:
            self.canvas.yview_scroll(delta, "units")
        return "break"

    def _on_shift_mousewheel(self, event) -> str:
        delta = -1 * int(event.delta / 120) if event.delta else 0
        if delta:
            self.canvas.xview_scroll(delta, "units")
        return "break"

    def _header(self, text: str, row: int, col: int, width: int) -> None:
        tk.Label(
            self.grid_frame,
            text=text,
            width=width,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground="black",
            bg="white",
            fg="black",
            font=("Meiryo UI", 8),
        ).grid(row=row, column=col, sticky="nsew")

    def _build_grid(self) -> None:
        for w in self.grid_frame.winfo_children():
            w.destroy()
        self.vars.clear()
        self.entries.clear()
        self.tt_marks.clear()
        self.cell_widgets.clear()

        self._header("教員", 0, 0, width=10)
        for c, cls in enumerate(self.col_labels, start=1):
            self._header(cls, 0, c, width=4)

        for r, teacher in enumerate(self.row_labels, start=1):
            self._header(teacher, r, 0, width=10)
            for c, cls in enumerate(self.col_labels, start=1):
                var = tk.StringVar(value="")
                self.vars[(teacher, cls)] = var
                ent = tk.Entry(
                    self.grid_frame,
                    textvariable=var,
                    width=4,
                    relief="flat",
                    bd=0,
                    highlightthickness=1,
                    highlightbackground="black",
                    highlightcolor="black",
                    bg="white",
                    fg="black",
                    justify="center",
                    font=("Meiryo UI", 8),
                )
                ent.grid(row=r, column=c, sticky="nsew")
                self.entries[(teacher, cls)] = ent
                self.cell_widgets[(r - 1, c - 1)] = ent
                self._bind_arrow_nav(ent, r - 1, c - 1)

                def _toggle_tt(_evt, key=(teacher, cls)):
                    if key in self.tt_marks:
                        self.tt_marks.remove(key)
                    else:
                        self.tt_marks.add(key)
                    self._apply_cell_style(key)

                ent.bind("<Double-Button-1>", _toggle_tt)

    def _apply_cell_style(self, key: tuple[str, str]) -> None:
        ent = self.entries.get(key)
        if ent is None:
            return
        if key in self.tt_marks:
            ent.configure(bg="#dbeafe")  # 薄い青
        else:
            ent.configure(bg="white")

    def set_axes(self, row_labels: list[str], col_labels: list[str]) -> None:
        if list(row_labels) == self.row_labels and list(col_labels) == self.col_labels:
            return
        old = self.dump_cells()
        self.row_labels = list(row_labels)
        self.col_labels = list(col_labels)
        self._build_grid()
        self.set_cells(old)

    def _bind_arrow_nav(self, widget: tk.Widget, r: int, c: int) -> None:
        widget.bind("<Up>", lambda e, rr=r, cc=c: self._move_focus(rr - 1, cc))
        widget.bind("<Down>", lambda e, rr=r, cc=c: self._move_focus(rr + 1, cc))
        widget.bind("<Left>", lambda e, rr=r, cc=c: self._move_focus(rr, cc - 1))
        widget.bind("<Right>", lambda e, rr=r, cc=c: self._move_focus(rr, cc + 1))
        widget.bind("<FocusIn>", lambda _e, w=widget: self._ensure_visible(w))
        widget.bind("<Button-1>", lambda _e, w=widget: self._ensure_visible(w))

    def _move_focus(self, r: int, c: int) -> str:
        w = self.cell_widgets.get((r, c))
        if w is not None:
            w.focus_set()
            self._ensure_visible(w)
        return "break"

    def _ensure_visible(self, widget: tk.Widget) -> None:
        self.update_idletasks()
        wx, wy = widget.winfo_x(), widget.winfo_y()
        ww, wh = widget.winfo_width(), widget.winfo_height()
        vx0, vy0 = self.canvas.canvasx(0), self.canvas.canvasy(0)
        vx1, vy1 = vx0 + self.canvas.winfo_width(), vy0 + self.canvas.winfo_height()
        cw = max(self.grid_frame.winfo_reqwidth(), 1)
        ch = max(self.grid_frame.winfo_reqheight(), 1)

        if wx < vx0:
            self.canvas.xview_moveto(max(0.0, wx / cw))
        elif wx + ww > vx1:
            self.canvas.xview_moveto(max(0.0, (wx + ww - self.canvas.winfo_width()) / cw))

        if wy < vy0:
            self.canvas.yview_moveto(max(0.0, wy / ch))
        elif wy + wh > vy1:
            self.canvas.yview_moveto(max(0.0, (wy + wh - self.canvas.winfo_height()) / ch))

    def dump_cells(self) -> dict[tuple[str, str], tuple[str, bool]]:
        out: dict[tuple[str, str], tuple[str, bool]] = {}
        for key, var in self.vars.items():
            hours = var.get().strip()
            tt = key in self.tt_marks
            if hours or tt:
                out[key] = (hours, tt)
        return out

    def set_cells(self, data: dict[tuple[str, str], tuple[str, bool]]) -> None:
        for key, var in self.vars.items():
            var.set("")
        self.tt_marks.clear()
        for key, (hours, tt) in data.items():
            if key in self.vars:
                self.vars[key].set(hours)
                if tt:
                    self.tt_marks.add(key)
        for key in self.entries:
            self._apply_cell_style(key)


class ScenarioFrame(ttk.Frame):
    def __init__(
        self,
        parent: tk.Misc,
        scenario_id: str,
        title: str,
        classes: list[str],
        day_periods: dict[str, int],
    ):
        super().__init__(parent, padding=6)
        self.scenario_id = scenario_id
        self.classes = list(classes)
        self.day_periods = dict(day_periods)

        ttk.Label(self, text=title).pack(anchor="w")
        pan = ttk.Panedwindow(self, orient="vertical")
        pan.pack(fill="both", expand=True)

        f1 = ttk.Frame(pan, padding=4)
        f2 = ttk.Frame(pan, padding=4)
        f3 = ttk.Frame(pan, padding=4)
        pan.add(f1, weight=1)
        pan.add(f2, weight=3)
        pan.add(f3, weight=3)

        self.tbl_weekly = EntryTableGrid(
            f1,
            "週時数（PDF入力）",
            [("class", "クラス"), ("subject", "教科"), ("hours", "週時数")],
            rows=260,
            cell_width=10,
        )
        self.grid_fixed = TimetableMatrixGrid(
            f2, "固定教科入力（予備シート形式）", self.classes, click_x_only=False, day_periods=self.day_periods
        )
        self.grid_manual = TimetableMatrixGrid(
            f3, "技能系手入力（予備シート形式）", self.classes, click_x_only=False, day_periods=self.day_periods
        )

        self.tbl_weekly.pack(fill="both", expand=True)
        self.grid_fixed.pack(fill="both", expand=True)
        self.grid_manual.pack(fill="both", expand=True)
        self.grid_manual.pack(fill="both", expand=True)

    def set_classes(self, classes: list[str]) -> None:
        self.classes = list(classes)
        self.grid_fixed.set_row_labels(classes)
        self.grid_manual.set_row_labels(classes)

    def set_day_periods(self, day_periods: dict[str, int]) -> None:
        self.day_periods = dict(day_periods)
        self.grid_fixed.set_day_periods(day_periods)
        self.grid_manual.set_day_periods(day_periods)

    def load(self, sc: dict) -> None:
        wr_rows = []
        for c, subject_map in sc.get("weekly_requirements", {}).items():
            for subject, hours in subject_map.items():
                wr_rows.append({"class": c, "subject": subject, "hours": str(hours)})
        self.tbl_weekly.set_rows(wr_rows)
        self.grid_fixed.set_from_assignments(sc.get("fixed_assignments", []))
        self.grid_manual.set_from_assignments(sc.get("manual_skill_assignments", []))

    def dump(self) -> dict:
        wr: dict[str, dict[str, float]] = {}
        for r in self.tbl_weekly.get_rows(drop_empty=True):
            c = r["class"].strip()
            s = r["subject"].strip()
            h = r["hours"].strip()
            if not c or not s or not h:
                continue
            try:
                hv = float(h)
            except ValueError:
                continue
            wr.setdefault(c, {})[s] = hv

        return {
            "id": self.scenario_id,
            "target_block": _scenario_id_to_block(self.scenario_id),
            "weekly_requirements": wr,
            "fixed_assignments": self.grid_fixed.dump_assignments(),
            "manual_skill_assignments": self.grid_manual.dump_assignments(),
        }


class TimetableGUI(tb.Window):
    def __init__(self) -> None:
        super().__init__(themename="flatly")
        self.title("時間割作成 GUI")
        self.state("zoomed")
        self._apply_window_icon()
        self._setup_styles()

        self.config_path = Path("sample_config.json")
        self.config_data = self.default_config()
        self.subject_options = list(DEFAULT_SUBJECTS)
        self.teacher_subject_options = list(TEACHER_SUBJECT_OPTIONS)
        self.tech_subject_options = list(TECH_SUBJECTS)
        self.subject_colors = dict(SUBJECT_COLOR_MAP)
        self._built_tabs: set[str] = set()

        self.create_widgets()
        self._apply_initial_common_values()

        self._assign_queue: queue.Queue | None = None
        self._assign_running = False

    def _resource_path(self, relative_path: str) -> Path:
        base = getattr(sys, "_MEIPASS", None)
        if base:
            return Path(base) / relative_path
        return Path(__file__).resolve().parent / relative_path

    def _apply_window_icon(self) -> None:
        # Use icon.ico if available (copied from iron.ico), fallback to iron.ico
        for name in ("icon.ico", "iron.ico"):
            path = self._resource_path(name)
            if path.exists():
                try:
                    self.iconbitmap(str(path))
                except Exception:
                    pass
                break

    def _setup_styles(self) -> None:
        # テーマのデフォルト配色を使用する
        tb.Style()

    def _set_startup_message(self, message: str) -> None:
        if self._startup_splash is None:
            return
        if self._startup_message_var is not None:
            self._startup_message_var.set(str(message).strip() or "初期化中...")
        self._startup_splash.update_idletasks()

    def _close_startup_splash(self) -> None:
        if self._startup_progress is not None:
            try:
                self._startup_progress.stop()
            except Exception:
                pass
            self._startup_progress = None
        if self._startup_splash is not None:
            try:
                self._startup_splash.destroy()
            except Exception:
                pass
            self._startup_splash = None

    def default_config(self) -> dict:
        return {
            "year": 6,
            "output_path": "./output/timetable.xlsx",
            "seed": 42,
            "subjects": list(DEFAULT_SUBJECTS),
            "classes": ["1-1", "1-2", "1-3", "1-4", "2-1", "2-2", "2-3", "2-4", "3-1", "3-2", "3-3"],
            "day_periods": {"Mon": 5, "Tue": 6, "Wed": 5, "Thu": 6, "Fri": 6},
            "teachers": [],
            "class_subject_teacher": {},
            "class_subject_unavailable": [],
            "solver": {"engine": "cp_sat", "time_limit_sec": 90, "random_restarts": 8},
            "scenarios": [
                {"id": SCENARIO_ID_UPPER, "target_block": "upper", "weekly_requirements": {}, "fixed_assignments": [], "manual_skill_assignments": []},
                {"id": SCENARIO_ID_LOWER, "target_block": "lower", "weekly_requirements": {}, "fixed_assignments": [], "manual_skill_assignments": []},
            ],
        }

    def _apply_initial_common_values(self) -> None:
        d = self.config_data
        self.var_output.set(d.get("output_path", ""))
        self.var_subjects.set(join_csv(d.get("subjects", list(DEFAULT_SUBJECTS))))
        self.var_classes.set(join_csv(d.get("classes", [])))
        dp = d.get("day_periods", {})
        self.var_mon.set(str(dp.get("Mon", 5)))
        self.var_tue.set(str(dp.get("Tue", 6)))
        self.var_wed.set(str(dp.get("Wed", 5)))
        self.var_thu.set(str(dp.get("Thu", 6)))
        self.var_fri.set(str(dp.get("Fri", 6)))
        self.var_load_progress_text.set("起動: 完了")

    def _build_tab_if_needed(self, key: str) -> None:
        if key in self._built_tabs:
            return
        builder_map = {
            "common": self.build_common_tab,
            "teachers": self.build_teachers_tab,
            "skill_input": self.build_skill_input_tab,
            "other_assign": self.build_other_assign_tab,
            "confirm": self.build_confirm_tab,
        }
        builder = builder_map.get(key)
        if builder is None:
            return
        builder()
        self._built_tabs.add(key)

    def _ensure_tabs_built(self, *keys: str) -> None:
        for key in keys:
            self._build_tab_if_needed(key)

    def _ensure_data_tabs_built(self) -> None:
        self._ensure_tabs_built("teachers", "skill_input", "other_assign")

    def _ensure_all_tabs_built(self) -> None:
        self._ensure_data_tabs_built()
        self._ensure_tabs_built("confirm")

    def _on_notebook_tab_changed(self, _event=None) -> None:
        current = self.nb.select()
        if not current:
            return
        mapping = {
            str(self.tab_common): "common",
            str(self.tab_teachers): "teachers",
            str(self.tab_skill_input): "skill_input",
            str(self.tab_other_assign): "other_assign",
            str(self.tab_confirm): "confirm",
        }
        key = mapping.get(current)
        if key:
            self._build_tab_if_needed(key)
            if key == "confirm":
                self.refresh_confirm(rebuild=True)

    def create_widgets(self) -> None:
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="新規", command=self.new_config).pack(side="left", padx=4)
        ttk.Button(top, text="読込", command=self.open_config).pack(side="left", padx=4)
        ttk.Button(top, text="保存", command=self.save_config).pack(side="left", padx=4)
        ttk.Button(top, text="基本設定反映", command=self.apply_basic_info_to_tabs).pack(side="left", padx=4)
        ttk.Button(top, text="Excelに出力", command=self.run_generate).pack(side="right", padx=4)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=8)

        self.tab_common = ttk.Frame(self.nb, padding=8)
        self.tab_teachers = ttk.Frame(self.nb, padding=8)
        self.tab_skill_input = ttk.Frame(self.nb, padding=8)
        self.tab_other_assign = ttk.Frame(self.nb, padding=8)
        self.tab_confirm = ttk.Frame(self.nb, padding=8)

        self.nb.add(self.tab_common, text="基本設定")
        self.nb.add(self.tab_teachers, text="教員")
        self.nb.add(self.tab_skill_input, text="技能科目入力")
        self.nb.add(self.tab_other_assign, text="その他科目割り振り")
        self.nb.add(self.tab_confirm, text="確認")

        self._build_tab_if_needed("common")
        self.nb.bind("<<NotebookTabChanged>>", self._on_notebook_tab_changed)

        footer = ttk.Frame(self, padding=(8, 0, 8, 8))
        footer.pack(fill="x", side="bottom")
        self.var_load_progress_text = tk.StringVar(value="JSON読込: 待機中")
        ttk.Label(footer, textvariable=self.var_load_progress_text).pack(side="left", padx=(0, 8))
        self.prg_load = ttk.Progressbar(footer, orient="horizontal", mode="determinate", maximum=100, length=260)
        self.prg_load.pack(side="left", padx=(12, 6))

    def build_common_tab(self) -> None:
        f = self.tab_common
        self.var_output = tk.StringVar()
        self.var_subjects = tk.StringVar()
        self.var_classes = tk.StringVar()
        self.var_mon = tk.StringVar()
        self.var_tue = tk.StringVar()
        self.var_wed = tk.StringVar()
        self.var_thu = tk.StringVar()
        self.var_fri = tk.StringVar()

        row = 0
        ttk.Label(f, text="出力Excel").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=self.var_output, width=90).grid(row=row, column=1, sticky="ew")
        ttk.Button(f, text="参照", command=self.pick_output).grid(row=row, column=2, padx=4)
        row += 1

        ttk.Label(f, text="教科(カンマ区切り)").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=self.var_subjects, width=90).grid(row=row, column=1, sticky="ew")
        row += 1

        ttk.Label(f, text="クラス一覧(カンマ区切り)").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=self.var_classes, width=90).grid(row=row, column=1, sticky="ew")
        row += 1

        ttk.Label(f, text="曜日ごとの時限数").grid(row=row, column=0, sticky="w", pady=4)
        dp = ttk.Frame(f)
        dp.grid(row=row, column=1, sticky="w")
        for i, (lbl, var) in enumerate([("月", self.var_mon), ("火", self.var_tue), ("水", self.var_wed), ("木", self.var_thu), ("金", self.var_fri)]):
            ttk.Label(dp, text=lbl).grid(row=0, column=i * 2, padx=(0, 2))
            ttk.Entry(dp, textvariable=var, width=5).grid(row=0, column=i * 2 + 1, padx=(0, 8))
        row += 1

        # 案内文の前に約3行分の余白を追加
        ttk.Frame(f, height=1).grid(row=row, column=0, columnspan=2, sticky="ew", pady=(0, 36))
        row += 1

        ttk.Label(f, text="上記がすべて入力出来たら、基本設定反映ボタンを押してください。").grid(row=row, column=0, columnspan=2, sticky="w")
        row += 1

        f.columnconfigure(1, weight=1)

    def build_teachers_tab(self) -> None:
        day_periods = self._current_day_periods()
        pan = ttk.Panedwindow(self.tab_teachers, orient="vertical")
        pan.pack(fill="both", expand=True)

        f1 = ttk.Frame(pan, padding=4)
        f2 = ttk.Frame(pan, padding=4)
        f3 = ttk.Frame(pan, padding=4)
        pan.add(f1, weight=1)
        pan.add(f2, weight=2)
        pan.add(f3, weight=3)

        action = ttk.Frame(f1)
        action.pack(fill="x", pady=(0, 6))
        ttk.Button(
            action,
            text="教員名を確定して不可コマ欄を生成",
            command=self.confirm_teacher_names,
        ).pack(side="right")

        self.tbl_teachers = EntryTableGrid(
            f1,
            "教員名・担当教科（苗字重複葉登録不可）",
            [("name", "教員名"), ("subject", "担当教科")],
            rows=240,
            cell_width=10,
            dropdown_values={"subject": self.teacher_subject_options},
        )
        self.tbl_teachers.pack(fill="both", expand=True)

        self.grid_teacher_unavail = TimetableMatrixGrid(
            f2,
            "教員の不可コマ（クリックで✕を入れる）",
            row_labels=[],
            click_x_only=True,
            day_periods=day_periods,
            merge_day_header=True,
            show_day_separators=True,
        )
        self.grid_teacher_unavail.pack(fill="both", expand=True)

        self.grid_teacher_class = ClassAssignGrid(
            f3,
            "担当クラス（時間数を数字入力。セルクリックでTT=背景色薄青",
            row_labels=[],
            col_labels=[],
        )
        self.grid_teacher_class.pack(fill="both", expand=True)

    def build_skill_input_tab(self) -> None:
        current_classes = split_csv(self.var_classes.get()) or self.config_data.get("classes", [])
        day_periods = self._current_day_periods()
        top = ttk.Frame(self.tab_skill_input)
        top.pack(fill="x", pady=(0, 6))
        ttk.Button(top, text="１年音美 → １年総合にコピー", command=self.copy_skill_upper_to_lower).pack(side="left", padx=4)

        pan = ttk.Panedwindow(self.tab_skill_input, orient="vertical")
        pan.pack(fill="both", expand=True)
        f1 = ttk.Frame(pan, padding=4)
        f2 = ttk.Frame(pan, padding=4)
        pan.add(f1, weight=1)
        pan.add(f2, weight=1)

        ttk.Label(f1, text="１年音美").pack(anchor="w")
        self.skill_upper_grid = TimetableMatrixGrid(
            f1,
            "時間割",
            row_labels=current_classes,
            day_periods=day_periods,
            subject_values=self.tech_subject_options,
            subject_colors=self.subject_colors,
            merge_day_header=True,
            show_day_separators=True,
        )
        self.skill_upper_grid.pack(fill="both", expand=True)

        ttk.Label(f2, text="１年総合").pack(anchor="w")
        self.skill_lower_grid = TimetableMatrixGrid(
            f2,
            "時間割",
            row_labels=current_classes,
            day_periods=day_periods,
            subject_values=self.tech_subject_options,
            subject_colors=self.subject_colors,
            merge_day_header=True,
            show_day_separators=True,
        )
        self.skill_lower_grid.pack(fill="both", expand=True)

    def build_other_assign_tab(self) -> None:
        current_classes = split_csv(self.var_classes.get()) or self.config_data.get("classes", [])
        day_periods = self._current_day_periods()

        top = ttk.Frame(self.tab_other_assign)
        top.pack(fill="x", pady=(0, 6))
        ttk.Button(top, text="技能科目入力タブをコピー", command=self.copy_from_skill_input).pack(side="left", padx=4)
        ttk.Button(top, text="教員設定を基に他教科割り当て", command=self.assign_other_subjects).pack(side="left", padx=4)
        self.var_assign_progress_text = tk.StringVar(value="待機中")
        self.prg_assign = ttk.Progressbar(top, orient="horizontal", mode="determinate", maximum=100, length=260)
        self.prg_assign.pack(side="left", padx=(12, 6))
        ttk.Label(top, textvariable=self.var_assign_progress_text).pack(side="left", padx=(0, 4))

        pan = ttk.Panedwindow(self.tab_other_assign, orient="vertical")
        pan.pack(fill="both", expand=True)
        f1 = ttk.Frame(pan, padding=4)
        f2 = ttk.Frame(pan, padding=4)
        pan.add(f1, weight=1)
        pan.add(f2, weight=1)

        ttk.Label(f1, text="１年音美").pack(anchor="w")
        self.other_upper_grid = TimetableMatrixGrid(
            f1,
            "各クラス × 曜日 × 時間数（読み取り専用）",
            row_labels=current_classes,
            day_periods=day_periods,
            readonly=True,
            readonly_cell_scale=1.5,
            subject_colors=self.subject_colors,
            merge_day_header=True,
            show_day_separators=True,
        )
        self.other_upper_grid.pack(fill="both", expand=True)

        ttk.Label(f2, text="１年総合").pack(anchor="w")
        self.other_lower_grid = TimetableMatrixGrid(
            f2,
            "各クラス × 曜日 × 時間数（読み取り専用）",
            row_labels=current_classes,
            day_periods=day_periods,
            readonly=True,
            readonly_cell_scale=1.5,
            subject_colors=self.subject_colors,
            merge_day_header=True,
            show_day_separators=True,
        )
        self.other_lower_grid.pack(fill="both", expand=True)

    def build_confirm_tab(self) -> None:
        ttk.Label(self.tab_confirm, text="確認画面: サマリとJSONを確認後に実行してください").pack(anchor="w")
        self.txt_confirm = tk.Text(
            self.tab_confirm,
            wrap="none",
            fg="black",
            bg="white",
            insertbackground="black",
        )
        self.txt_confirm.pack(fill="both", expand=True)

    def pick_output(self) -> None:
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            self.var_output.set(p)

    def new_config(self) -> None:
        if not messagebox.askyesno("確認", "現在の入力を破棄して新規作成しますか？"):
            return
        self.config_data = self.default_config()
        self.load_to_ui()

    def open_config(self) -> None:
        p = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not p:
            return
        try:
            self._load_config_from_path(Path(p))
        except Exception as e:
            self._set_load_progress(0, "JSON読込: エラー")
            messagebox.showerror("読込エラー", str(e))

    def save_config(self) -> None:
        p = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not p:
            return
        try:
            save_path = Path(p)
            snapshot = self._build_config_snapshot()
            self._save_config_to_path(save_path, snapshot)
            messagebox.showinfo("保存", f"保存しました:\n{save_path}")
        except Exception as e:
            messagebox.showerror("保存エラー", str(e))

    def _load_config_from_path(self, path: Path) -> None:
        self._set_load_progress(5, "JSON読込: ファイル読込中...")
        self.config_data = read_json(path)
        self._set_load_progress(35, "JSON読込: UIへ反映中...")
        self.config_path = path
        self.load_to_ui(progress_cb=self._set_load_progress)
        self._set_load_progress(100, "JSON読込: 完了")

    def _save_config_to_path(self, path: Path, snapshot: dict | None = None) -> None:
        self.config_data = snapshot or self._build_config_snapshot()
        write_json(path, self.config_data)
        self.config_path = path

    def _get_scenario(self, sid: str) -> dict:
        normalized_sid = _normalize_scenario_id(sid)
        for sc in self.config_data.get("scenarios", []):
            sc_id = _normalize_scenario_id(str(sc.get("id", "")))
            if sc_id == normalized_sid:
                # 旧IDで保存された設定を読み込んだときも、新IDへ正規化して扱う。
                sc["id"] = normalized_sid
                if "target_block" not in sc or not sc.get("target_block"):
                    sc["target_block"] = _scenario_id_to_block(normalized_sid)
                return sc
        sc = {
            "id": normalized_sid,
            "target_block": _scenario_id_to_block(normalized_sid),
            "weekly_requirements": {},
            "fixed_assignments": [],
            "manual_skill_assignments": [],
        }
        self.config_data.setdefault("scenarios", []).append(sc)
        return sc

    def _sync_matrix_labels(self) -> None:
        classes = split_csv(self.var_classes.get())
        day_periods = self._current_day_periods()
        self.skill_upper_grid.set_layout(row_labels=classes, day_periods=day_periods)
        self.skill_lower_grid.set_layout(row_labels=classes, day_periods=day_periods)
        self.other_upper_grid.set_layout(row_labels=classes, day_periods=day_periods)
        self.other_lower_grid.set_layout(row_labels=classes, day_periods=day_periods)

        teacher_names = [r["name"].strip() for r in self.tbl_teachers.get_rows(drop_empty=True) if r["name"].strip()]
        self.grid_teacher_unavail.set_layout(row_labels=teacher_names, day_periods=day_periods)
        self.grid_teacher_class.set_axes(teacher_names, classes)

    def _sync_subject_options(self) -> None:
        subjects = split_csv(self.var_subjects.get())
        if not subjects:
            subjects = list(DEFAULT_SUBJECTS)
        self.subject_options = subjects
        self.teacher_subject_options = list(TEACHER_SUBJECT_OPTIONS)
        self.tech_subject_options = list(TECH_SUBJECTS)
        self.subject_colors = dict(SUBJECT_COLOR_MAP)
        self.tbl_teachers.set_dropdown_values("subject", self.teacher_subject_options)
        self.skill_upper_grid.set_subject_values(self.tech_subject_options, self.subject_colors)
        self.skill_lower_grid.set_subject_values(self.tech_subject_options, self.subject_colors)

    def _current_day_periods(self) -> dict[str, int]:
        def _to_int(v: str, default: int) -> int:
            try:
                x = int(v)
                return x if x > 0 else default
            except Exception:
                return default

        return {
            "Mon": _to_int(self.var_mon.get(), 5),
            "Tue": _to_int(self.var_tue.get(), 6),
            "Wed": _to_int(self.var_wed.get(), 5),
            "Thu": _to_int(self.var_thu.get(), 6),
            "Fri": _to_int(self.var_fri.get(), 6),
        }

    def apply_basic_info_to_tabs(self, silent: bool = False, sync_labels: bool = True) -> None:
        try:
            self._ensure_data_tabs_built()
            if sync_labels:
                self._sync_matrix_labels()
            if not silent:
                messagebox.showinfo("完了", "基本情報を他タブへ反映しました。")
        except Exception as e:
            if not silent:
                messagebox.showerror("反映エラー", str(e))

    def copy_from_skill_input(self) -> None:
        # 先にコピー先の過去入力をクリアしてからコピーする
        self.other_upper_grid.clear()
        self.other_lower_grid.clear()
        self.other_upper_grid.set_value_map(self.skill_upper_grid.get_value_map())
        self.other_lower_grid.set_value_map(self.skill_lower_grid.get_value_map())
        messagebox.showinfo("完了", "技能科目入力タブの内容をコピーしました。")

    def copy_skill_upper_to_lower(self) -> None:
        self.skill_lower_grid.clear()
        self.skill_lower_grid.set_value_map(self.skill_upper_grid.get_value_map())
        messagebox.showinfo("完了", "１年音美の内容を１年総合へコピーしました。")

    def assign_other_subjects(self) -> None:
        if self._assign_running:
            return
        try:
            self.collect_from_ui()
            self.prg_assign.configure(value=0)
            self.prg_assign.configure(mode="indeterminate")
            self.prg_assign.start(12)
            self.var_assign_progress_text.set("開始準備中...（探索中）")
            self.update_idletasks()

            # 提案探索条件はGUI入力を廃止し、従来の既定値で固定運用
            suggestion_options = {
                "max_one_step": 5,
                "max_two_step": 3,
                "spec_limit": 240,
                "two_pool_limit": 40,
                "exhaustive_two_step": False,
            }

            self._assign_queue = queue.Queue()
            self._assign_running = True

            def _worker() -> None:
                try:
                    def _on_progress(p: int, msg: str) -> None:
                        if self._assign_queue is not None:
                            self._assign_queue.put(("progress", int(p), str(msg)))

                    solved = solve_all_scenarios(
                        self.config_data,
                        exempt_subjects=set(AUTO_EXEMPT_SUBJECTS),
                        progress_callback=_on_progress,
                        suggestion_options=suggestion_options,
                    )
                    if self._assign_queue is not None:
                        self._assign_queue.put(("done", solved, ""))
                except SchedulerError as e:
                    if self._assign_queue is not None:
                        self._assign_queue.put(("error", str(e), ""))
                except Exception as e:
                    if self._assign_queue is not None:
                        self._assign_queue.put(("error", f"予期しないエラー: {e}", ""))

            threading.Thread(target=_worker, daemon=True).start()
            self.after(80, self._poll_assign_queue)
        except Exception as e:
            self.prg_assign.stop()
            self.prg_assign.configure(mode="determinate", value=0)
            self.var_assign_progress_text.set("エラー")
            self._show_assign_diagnostics_popup("割り当てエラー", f"入力準備でエラー:\n{e}")

    @staticmethod
    def _slot_sort_key(slot: str) -> tuple[int, int, str]:
        try:
            day, per = slot.split("-", 1)
            day_idx = DAY_ORDER.index(day) if day in DAY_ORDER else len(DAY_ORDER)
            return (day_idx, int(per), slot)
        except Exception:
            return (len(DAY_ORDER), 99, slot)

    @staticmethod
    def _class_sort_key(cls: str) -> tuple[int, int, str]:
        parts = cls.split("-", 1)
        if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
            return (int(parts[0]), int(parts[1]), cls)
        return (99, 99, cls)

    def _get_solved_block_data(self, solved: dict, candidates: list[str]) -> dict:
        for sid_candidate in candidates:
            if sid_candidate in solved:
                return solved[sid_candidate]
        return {}

    def _collect_teacher_tt_info(self) -> tuple[dict[str, str], dict[tuple[str, str], int], dict[str, dict[str, str]]]:
        teacher_subjects: dict[str, str] = {}
        tt_demands: dict[tuple[str, str], int] = {}
        non_tt_map: dict[str, dict[str, str]] = {}

        for teacher in self.config_data.get("teachers", []):
            name = str(teacher.get("name", "")).strip()
            if not name:
                continue
            subjects = teacher.get("subjects", [])
            subject = str(subjects[0]).strip() if subjects else ""
            teacher_subjects[name] = subject

            for ca in teacher.get("class_assignments", []):
                cls = str(ca.get("class", "")).strip()
                if not cls:
                    continue
                raw_hours = ca.get("hours", 0)
                try:
                    hours = int(raw_hours)
                except Exception:
                    try:
                        hours = int(float(raw_hours))
                    except Exception:
                        hours = 0

                if hours <= 0:
                    continue

                if bool(ca.get("tt", False)):
                    tt_demands[(name, cls)] = tt_demands.get((name, cls), 0) + hours
                elif subject:
                    non_tt_map.setdefault(cls, {})[subject] = name

        return teacher_subjects, tt_demands, non_tt_map

    def _allocate_tt_marks(
        self,
        assignments: dict[tuple[str, str], str],
        teacher_assignments: dict[tuple[str, str], str],
        teacher_subjects: dict[str, str],
        tt_demands: dict[tuple[str, str], int],
        non_tt_class_subject_teacher: dict[str, dict[str, str]],
    ) -> tuple[set[tuple[str, str]], dict[tuple[str, str], int]]:
        tt_marks: set[tuple[str, str]] = set()
        unmet: dict[tuple[str, str], int] = {}
        if not assignments or not tt_demands:
            return tt_marks, unmet

        classes_in_block = {cls for cls, _slot in assignments}
        scoped_demands = {
            key: hours
            for key, hours in tt_demands.items()
            if key[1] in classes_in_block and hours > 0
        }
        if not scoped_demands:
            return tt_marks, unmet

        busy_slots: dict[str, set[str]] = {}
        for (_cls, slot), teacher in teacher_assignments.items():
            tname = str(teacher).strip()
            if tname:
                busy_slots.setdefault(tname, set()).add(slot)

        for (cls, slot), subj in assignments.items():
            tname = str(non_tt_class_subject_teacher.get(cls, {}).get(subj, "")).strip()
            if tname:
                busy_slots.setdefault(tname, set()).add(slot)

        class_subject_slots: dict[tuple[str, str], list[str]] = {}
        for (cls, slot), subj in assignments.items():
            subject = str(subj).strip()
            if subject:
                class_subject_slots.setdefault((cls, subject), []).append(slot)
        for key in class_subject_slots:
            class_subject_slots[key].sort(key=self._slot_sort_key)

        tt_busy_slots: dict[str, set[str]] = {}
        for (teacher, cls), hours in sorted(scoped_demands.items()):
            remain = int(hours)
            subject = teacher_subjects.get(teacher, "").strip()
            if remain <= 0:
                continue
            if not subject:
                unmet[(teacher, cls)] = remain
                continue

            candidates = class_subject_slots.get((cls, subject), [])
            for slot in candidates:
                if remain <= 0:
                    break
                if slot in busy_slots.get(teacher, set()):
                    continue
                if slot in tt_busy_slots.setdefault(teacher, set()):
                    continue
                tt_marks.add((cls, slot))
                tt_busy_slots[teacher].add(slot)
                remain -= 1

            if remain > 0:
                unmet[(teacher, cls)] = remain

        return tt_marks, unmet

    def _summarize_solver_tt_assignments(
        self,
        assignments: dict[tuple[str, str], str],
        tt_assignments: dict[tuple[str, str], list[str]] | None,
        tt_demands: dict[tuple[str, str], int],
    ) -> tuple[set[tuple[str, str]], dict[tuple[str, str], int]]:
        tt_marks: set[tuple[str, str]] = set()
        unmet: dict[tuple[str, str], int] = {}
        if not assignments or not tt_demands:
            return tt_marks, unmet

        classes_in_block = {cls for cls, _slot in assignments}
        scoped_demands = {
            key: int(hours)
            for key, hours in tt_demands.items()
            if key[1] in classes_in_block and int(hours) > 0
        }
        if not scoped_demands:
            return tt_marks, unmet

        actual_counts: dict[tuple[str, str], int] = {}
        for (cls, slot), teachers in (tt_assignments or {}).items():
            if cls not in classes_in_block:
                continue
            teacher_list = [str(t).strip() for t in teachers if str(t).strip()]
            if not teacher_list:
                continue
            tt_marks.add((cls, slot))
            for teacher in teacher_list:
                key = (teacher, cls)
                actual_counts[key] = actual_counts.get(key, 0) + 1

        for key, hours in scoped_demands.items():
            remain = int(hours) - actual_counts.get(key, 0)
            if remain > 0:
                unmet[key] = remain

        return tt_marks, unmet

    def _poll_assign_queue(self) -> None:
        if not self._assign_running or self._assign_queue is None:
            return
        keep_poll = True
        try:
            while True:
                kind, a, _b = self._assign_queue.get_nowait()
                if kind == "progress":
                    msg = str(_b).strip() if _b else ""
                    self.var_assign_progress_text.set(f"{int(a)}% {msg}".strip())
                elif kind == "done":
                    solved = a

                    with open("debug_log.txt", "a", encoding='utf-8') as log:
                        log.write(f"\n=== GUI QUEUE RECEIVED ===\n")
                        log.write(f"solved keys: {list(solved.keys())}\n")
                        for sid, data in solved.items():
                            log.write(f"  {sid}: {list(data.keys())}\n")
                            if "assignments" in data:
                                log.write(f"    assignments count: {len(data['assignments'])}\n")
                                for k, v in list(data['assignments'].items())[:5]:
                                    log.write(f"      {k} -> {v}\n")

                    up_map: dict[tuple[str, str], str] = {}
                    lo_map: dict[tuple[str, str], str] = {}

                    up_data = self._get_solved_block_data(solved, [SCENARIO_ID_UPPER, "１年音美"])
                    lo_data = self._get_solved_block_data(solved, [SCENARIO_ID_LOWER, "１年総合"])

                    up_assignments = up_data.get("assignments", {})
                    for (cls, slot), subj in up_assignments.items():
                        up_map[(cls, slot)] = subj

                    lo_assignments = lo_data.get("assignments", {})
                    for (cls, slot), subj in lo_assignments.items():
                        lo_map[(cls, slot)] = subj

                    teacher_subjects, tt_demands, non_tt_map = self._collect_teacher_tt_info()
                    if "tt_assignments" in up_data:
                        up_tt_marks, up_unmet = self._summarize_solver_tt_assignments(
                            assignments=up_assignments,
                            tt_assignments=up_data.get("tt_assignments", {}),
                            tt_demands=tt_demands,
                        )
                    else:
                        up_tt_marks, up_unmet = self._allocate_tt_marks(
                            assignments=up_assignments,
                            teacher_assignments=up_data.get("teacher_assignments", {}),
                            teacher_subjects=teacher_subjects,
                            tt_demands=tt_demands,
                            non_tt_class_subject_teacher=non_tt_map,
                        )
                    if "tt_assignments" in lo_data:
                        lo_tt_marks, lo_unmet = self._summarize_solver_tt_assignments(
                            assignments=lo_assignments,
                            tt_assignments=lo_data.get("tt_assignments", {}),
                            tt_demands=tt_demands,
                        )
                    else:
                        lo_tt_marks, lo_unmet = self._allocate_tt_marks(
                            assignments=lo_assignments,
                            teacher_assignments=lo_data.get("teacher_assignments", {}),
                            teacher_subjects=teacher_subjects,
                            tt_demands=tt_demands,
                            non_tt_class_subject_teacher=non_tt_map,
                        )

                    with open("debug_log.txt", "a", encoding='utf-8') as log:
                        log.write(f"up_map size: {len(up_map)}, lo_map size: {len(lo_map)}\n")

                    up_display_map = {
                        key: (TimetableMatrixGrid._with_tt_suffix(subj) if key in up_tt_marks else TimetableMatrixGrid._strip_tt_suffix(subj))
                        for key, subj in up_map.items()
                    }
                    lo_display_map = {
                        key: (TimetableMatrixGrid._with_tt_suffix(subj) if key in lo_tt_marks else TimetableMatrixGrid._strip_tt_suffix(subj))
                        for key, subj in lo_map.items()
                    }

                    self.other_upper_grid.clear_tt_highlights()
                    self.other_lower_grid.clear_tt_highlights()
                    self.other_upper_grid.set_tt_highlights(up_tt_marks)
                    self.other_lower_grid.set_tt_highlights(lo_tt_marks)
                    self.other_upper_grid.set_value_map(up_display_map)
                    self.other_lower_grid.set_value_map(lo_display_map)
                    self.prg_assign.stop()
                    self.prg_assign.configure(mode="determinate", value=100)
                    unmet_total = sum(up_unmet.values()) + sum(lo_unmet.values())
                    self.var_assign_progress_text.set("完了" if unmet_total == 0 else f"完了（TT不足 {unmet_total}）")
                    self._assign_running = False
                    keep_poll = False
                    if unmet_total == 0:
                        messagebox.showinfo("完了", "他教科の自動割り当てが完了しました。")
                    else:
                        lines = ["他教科の自動割り当てが完了しました。", "", f"TT不足コマ数: {unmet_total}", ""]
                        for (teacher, cls), remain in sorted({**up_unmet, **lo_unmet}.items()):
                            lines.append(f"- {teacher} / {cls}: 残り {remain}")
                        messagebox.showwarning("完了（TT一部未割当）", "\n".join(lines))
                elif kind == "error":
                    self.prg_assign.stop()
                    self.prg_assign.configure(mode="determinate", value=0)
                    self.var_assign_progress_text.set("エラー")
                    self._assign_running = False
                    keep_poll = False
                    self._show_assign_diagnostics_popup("割り当てエラー", str(a))
        except queue.Empty:
            pass

        if keep_poll and self._assign_running:
            self.after(80, self._poll_assign_queue)

    def confirm_teacher_names(self) -> None:
        teacher_names = [r["name"].strip() for r in self.tbl_teachers.get_rows(drop_empty=True) if r["name"].strip()]
        if not teacher_names:
            messagebox.showwarning("確認", "教員名が入力されていません。")
            return
        classes = split_csv(self.var_classes.get())
        self.grid_teacher_unavail.set_row_labels(teacher_names)
        self.grid_teacher_class.set_axes(teacher_names, classes)
        messagebox.showinfo("完了", f"不可コマ入力欄を {len(teacher_names)} 名分に更新しました。")

    def _set_load_progress(self, percent: int, message: str) -> None:
        if hasattr(self, "prg_load"):
            self.prg_load.configure(value=max(0, min(100, int(percent))))
        if hasattr(self, "var_load_progress_text"):
            self.var_load_progress_text.set(message)
        self.update_idletasks()

    def load_to_ui(self, progress_cb=None) -> None:
        def _pg(p: int, msg: str) -> None:
            if progress_cb:
                progress_cb(p, msg)

        d = self.config_data
        _pg(40, "JSON読込: 基本設定反映")
        self.var_output.set(d.get("output_path", ""))
        d.setdefault("solver", {})
        d["solver"]["engine"] = "cp_sat"
        self.var_subjects.set(join_csv(d.get("subjects", list(DEFAULT_SUBJECTS))))
        self.var_classes.set(join_csv(d.get("classes", [])))

        _pg(52, "JSON読込: 曜日時限反映")
        dp = d.get("day_periods", {})
        self.var_mon.set(str(dp.get("Mon", 5)))
        self.var_tue.set(str(dp.get("Tue", 6)))
        self.var_wed.set(str(dp.get("Wed", 5)))
        self.var_thu.set(str(dp.get("Thu", 6)))
        self.var_fri.set(str(dp.get("Fri", 6)))

        _pg(58, "JSON読込: タブ準備")
        self._ensure_data_tabs_built()
        self._sync_subject_options()

        _pg(62, "JSON読込: 教員データ反映")
        teacher_rows = []
        teacher_marks: set[tuple[str, str]] = set()
        teacher_class_cells: dict[tuple[str, str], tuple[str, bool]] = {}
        for t in d.get("teachers", []):
            subj_list = t.get("subjects", [])
            teacher_rows.append({"name": t.get("name", ""), "subject": (subj_list[0] if subj_list else "")})
            name = t.get("name", "").strip()
            for slot in t.get("unavailable_slots", []):
                teacher_marks.add((name, slot))
            if t.get("class_assignments"):
                for ca in t.get("class_assignments", []):
                    cls = ca.get("class", "").strip()
                    hrs = str(ca.get("hours", "")).strip()
                    tt = bool(ca.get("tt", False))
                    if cls:
                        teacher_class_cells[(name, cls)] = (hrs, tt)
            else:
                # 旧データ互換: assigned_classes があれば hours=1, tt=False として復元
                for cls in t.get("assigned_classes", []):
                    teacher_class_cells[(name, cls)] = ("1", False)
        self.tbl_teachers.set_rows(teacher_rows)

        _pg(75, "JSON読込: タブ同期")
        self._sync_matrix_labels()
        self.grid_teacher_unavail.set_x_marks(teacher_marks)
        self.grid_teacher_class.set_cells(teacher_class_cells)
        sc_up = self._get_scenario(SCENARIO_ID_UPPER)
        sc_lo = self._get_scenario(SCENARIO_ID_LOWER)
        _pg(86, "JSON読込: 技能/割当反映")
        self.skill_upper_grid.set_from_assignments(sc_up.get("manual_skill_assignments", []))
        self.skill_lower_grid.set_from_assignments(sc_lo.get("manual_skill_assignments", []))

        def _split_tt(assignments: list[dict]) -> tuple[list[dict], set[tuple[str, str]]]:
            cleaned: list[dict] = []
            marks: set[tuple[str, str]] = set()
            for a in assignments:
                cls = a.get("class")
                slot = a.get("slot")
                subj = str(a.get("subject", "")).strip()
                if subj.endswith("(T)"):
                    marks.add((cls, slot))
                    subj = TimetableMatrixGrid._strip_tt_suffix(subj)
                cleaned.append({"class": cls, "slot": slot, "subject": subj})
            return cleaned, marks

        up_fixed, up_marks = _split_tt(sc_up.get("fixed_assignments", []))
        lo_fixed, lo_marks = _split_tt(sc_lo.get("fixed_assignments", []))
        self.other_upper_grid.set_from_assignments(up_fixed)
        self.other_lower_grid.set_from_assignments(lo_fixed)
        if sc_up.get("tt_marks"):
            self.other_upper_grid.set_tt_highlights(
                {(m.get("class"), m.get("slot")) for m in sc_up.get("tt_marks", []) if m.get("class") and m.get("slot")}
            )
        if sc_lo.get("tt_marks"):
            self.other_lower_grid.set_tt_highlights(
                {(m.get("class"), m.get("slot")) for m in sc_lo.get("tt_marks", []) if m.get("class") and m.get("slot")}
            )
        if up_marks:
            self.other_upper_grid.set_tt_highlights(self.other_upper_grid.tt_highlight_keys | up_marks)
        if lo_marks:
            self.other_lower_grid.set_tt_highlights(self.other_lower_grid.tt_highlight_keys | lo_marks)

        self._apply_tt_suffix_to_grid(self.other_upper_grid)
        self._apply_tt_suffix_to_grid(self.other_lower_grid)
        if "confirm" in self._built_tabs and self.nb.select() == str(self.tab_confirm):
            _pg(96, "JSON読込: 確認タブ更新")
            self.refresh_confirm(rebuild=True)
        _pg(100, "JSON読込: 完了")

    def _build_config_snapshot(self) -> dict:
        self._ensure_data_tabs_built()
        d = dict(self.config_data)
        d["output_path"] = self.var_output.get().strip()
        d["solver"] = dict(d.get("solver", {}))
        d["solver"]["engine"] = "cp_sat"
        self._sync_subject_options()
        d["subjects"] = list(self.subject_options)
        d["classes"] = split_csv(self.var_classes.get())

        d["day_periods"] = {
            "Mon": int(self.var_mon.get() or 5),
            "Tue": int(self.var_tue.get() or 6),
            "Wed": int(self.var_wed.get() or 5),
            "Thu": int(self.var_thu.get() or 6),
            "Fri": int(self.var_fri.get() or 6),
        }
        self._sync_matrix_labels()

        teachers = []
        teacher_x = self.grid_teacher_unavail.dump_x_marks()
        class_cells = self.grid_teacher_class.dump_cells()
        class_subject_teacher: dict[str, dict[str, str]] = {}
        for r in self.tbl_teachers.get_rows(drop_empty=True):
            name = r["name"].strip()
            if not name:
                continue
            slots = sorted(slot for tname, slot in teacher_x if tname == name)
            class_assignments = []
            for (tname, cls), (hours_text, tt) in class_cells.items():
                if tname != name:
                    continue
                hours = 0
                if hours_text:
                    try:
                        hours = int(hours_text)
                    except ValueError:
                        try:
                            hours = int(float(hours_text))
                        except Exception:
                            hours = 0
                if hours > 0 or tt:
                    class_assignments.append({"class": cls, "hours": hours, "tt": tt})
            class_assignments.sort(key=lambda x: x["class"])
            assigned_classes = [x["class"] for x in class_assignments if x.get("hours", 0) > 0]
            subj = r.get("subject", "").strip()
            subjects = [subj] if subj else []
            if subj:
                for cls in assigned_classes:
                    class_subject_teacher.setdefault(cls, {})[subj] = name
            teachers.append(
                {
                    "name": name,
                    "subjects": subjects,
                    "assigned_classes": assigned_classes,
                    "class_assignments": class_assignments,
                    "unavailable_slots": slots,
                }
            )
        d["teachers"] = teachers
        d["class_subject_teacher"] = class_subject_teacher

        # 不可コマは教員単位で管理（クラス単位不可は統合により廃止）
        d["class_subject_unavailable"] = []

        old_map: dict[str, dict] = {}
        for s in d.get("scenarios", []):
            normalized_id = _normalize_scenario_id(str(s.get("id", "")))
            if normalized_id:
                old_map[normalized_id] = s
        d["scenarios"] = [
            {
                "id": SCENARIO_ID_UPPER,
                "target_block": "upper",
                "weekly_requirements": old_map.get(SCENARIO_ID_UPPER, {}).get("weekly_requirements", {}),
                "fixed_assignments": self.other_upper_grid.dump_assignments(),
                "manual_skill_assignments": self.skill_upper_grid.dump_assignments(),
                "tt_marks": [
                    {"class": cls, "slot": slot} for cls, slot in sorted(self.other_upper_grid.tt_highlight_keys)
                ],
            },
            {
                "id": SCENARIO_ID_LOWER,
                "target_block": "lower",
                "weekly_requirements": old_map.get(SCENARIO_ID_LOWER, {}).get("weekly_requirements", {}),
                "fixed_assignments": self.other_lower_grid.dump_assignments(),
                "manual_skill_assignments": self.skill_lower_grid.dump_assignments(),
                "tt_marks": [
                    {"class": cls, "slot": slot} for cls, slot in sorted(self.other_lower_grid.tt_highlight_keys)
                ],
            },
        ]

        return d

    def collect_from_ui(self) -> None:
        self.config_data = self._build_config_snapshot()

    def refresh_confirm(self, rebuild: bool = False) -> None:
        try:
            if rebuild:
                self.collect_from_ui()
        except Exception as e:
            messagebox.showerror("入力エラー", str(e))
            return

        d = self.config_data
        lines = [
            "=== 入力サマリ ===",
            f"year: {d.get('year')}",
            f"output_path: {d.get('output_path')}",
            f"classes: {len(d.get('classes', []))}",
            f"teachers: {len(d.get('teachers', []))}",
            f"class_subject_teacher rows: {sum(len(v) for v in d.get('class_subject_teacher', {}).values())}",
            f"class_subject_unavailable rows: {len(d.get('class_subject_unavailable', []))}",
            f"scenarios: {len(d.get('scenarios', []))}",
            "",
            "=== JSON Preview ===",
            json.dumps(d, ensure_ascii=False, indent=2),
        ]
        self.txt_confirm.delete("1.0", "end")
        self.txt_confirm.insert("1.0", "\n".join(lines))

    def _apply_tt_suffix_to_grid(self, grid: TimetableMatrixGrid) -> None:
        if not grid.tt_highlight_keys:
            return
        for key in grid.tt_highlight_keys:
            var = grid.vars.get(key)
            if var is None:
                continue
            text = var.get().strip()
            if not text:
                continue
            if text.endswith("(T)"):
                continue
            var.set(TimetableMatrixGrid._with_tt_suffix(text))

    def _find_empty_slots(self) -> list[str]:
        d = self.config_data
        classes = [str(cls).strip() for cls in d.get("classes", []) if str(cls).strip()]
        day_periods = d.get("day_periods", {})
        all_slots: list[str] = []
        for day in DAY_ORDER:
            for p in range(1, int(day_periods.get(day, 0)) + 1):
                all_slots.append(f"{day}-{p}")

        other_hours_by_class: dict[str, int] = {}
        for t in d.get("teachers", []):
            subjects = t.get("subjects", [])
            subject = str(subjects[0]).strip() if subjects else ""
            if not subject or _is_skill_subject_name(subject):
                continue
            for ca in t.get("class_assignments", []):
                cls = str(ca.get("class", "")).strip()
                if not cls:
                    continue
                if bool(ca.get("tt", False)):
                    continue
                hours = ca.get("hours", 0)
                try:
                    hours = int(hours)
                except Exception:
                    try:
                        hours = int(float(hours))
                    except Exception:
                        hours = 0
                if hours > 0:
                    other_hours_by_class[cls] = other_hours_by_class.get(cls, 0) + hours

        missing: list[str] = []
        for sc in d.get("scenarios", []):
            sid = sc.get("id", "")
            assigned: set[tuple[str, str]] = set()
            assigned_other_hours_by_class: dict[str, int] = {}
            for a in sc.get("fixed_assignments", []):
                cls = str(a.get("class", "")).strip()
                slot = str(a.get("slot", "")).strip()
                subj = TimetableMatrixGrid._strip_tt_suffix(str(a.get("subject", "")).strip())
                if cls and slot:
                    assigned.add((cls, slot))
                if cls and subj and not _is_skill_subject_name(subj):
                    assigned_other_hours_by_class[cls] = assigned_other_hours_by_class.get(cls, 0) + 1
            for a in sc.get("manual_skill_assignments", []):
                cls = str(a.get("class", "")).strip()
                slot = str(a.get("slot", "")).strip()
                if cls and slot:
                    assigned.add((cls, slot))
            for cls in classes:
                shortage = other_hours_by_class.get(cls, 0) - assigned_other_hours_by_class.get(cls, 0)
                if shortage <= 0:
                    continue
                empty_slots = [slot for slot in all_slots if (cls, slot) not in assigned]
                for slot in empty_slots[:shortage]:
                    missing.append(f"{sid}: {cls} {slot}")
                if shortage > len(empty_slots):
                    for idx in range(shortage - len(empty_slots)):
                        missing.append(f"{sid}: {cls} その他教科不足 {idx + 1}")
        return missing

    def _prompt_excel_variant_settings(self) -> dict | None:
        defaults = {
            "①": {"onbi_start_music": True, "tech_subject": "技術"},
            "②": {"onbi_start_music": False, "tech_subject": "家庭"},
            "③": {"onbi_start_music": None, "tech_subject": "技術"},
            "④": {"onbi_start_music": True, "tech_subject": "家庭"},
            "⑤": {"onbi_start_music": False, "tech_subject": "技術"},
            "⑥": {"onbi_start_music": None, "tech_subject": "家庭"},
        }
        saved = self.config_data.get("excel_variant_settings", {}) or {}
        for key, cfg in saved.items():
            if key in defaults:
                if "onbi_start_music" in cfg:
                    defaults[key]["onbi_start_music"] = cfg["onbi_start_music"]
                if "tech_subject" in cfg:
                    defaults[key]["tech_subject"] = cfg["tech_subject"]

        classes = self.config_data.get("classes", [])
        grade1_classes = [c for c in classes if str(c).startswith("1-")]
        grade1_classes.sort(key=lambda x: self._class_sort_key(x))

        def _slot_sort_key(slot: str) -> tuple[int, int]:
            day, p = slot.split("-")
            return (DAY_ORDER.index(day), int(p))

        def _collect_skill_items(sid: str) -> list[dict]:
            sc = self._get_scenario(sid)
            merged: dict[tuple[str, str, str], dict] = {}

            def _add_items(items: list[dict]) -> None:
                for a in items:
                    cls = str(a.get("class", "")).strip()
                    slot = str(a.get("slot", "")).strip()
                    subj = str(a.get("subject", "")).strip()
                    if subj.endswith("(T)"):
                        subj = subj[:-3]
                    if not (cls and slot and subj):
                        continue
                    key = (cls, slot, subj)
                    if key not in merged:
                        merged[key] = {"class": cls, "slot": slot, "subject": subj}

            # manual_skill_assignments を優先して重複を排除
            _add_items(sc.get("manual_skill_assignments", []))
            _add_items(sc.get("fixed_assignments", []))

            cleaned = list(merged.values())
            cleaned.sort(key=lambda x: (self._class_sort_key(x["class"]), _slot_sort_key(x["slot"])))
            return cleaned

        def _default_onbi_map(start_music: bool | None) -> dict[str, str]:
            if start_music is None:
                return {}
            out = {}
            for idx, cls in enumerate(grade1_classes):
                if start_music:
                    out[cls] = "音楽" if idx % 2 == 0 else "美術"
                else:
                    out[cls] = "美術" if idx % 2 == 0 else "音楽"
            return out

        variants = [
            {"id": "①", "use_lower": False},
            {"id": "②", "use_lower": False},
            {"id": "③", "use_lower": True},
            {"id": "④", "use_lower": False},
            {"id": "⑤", "use_lower": False},
            {"id": "⑥", "use_lower": True},
        ]

        onbi_options = ["音楽", "美術"]
        tech_options = ["技術", "家庭"]

        result: dict | None = None
        win = tk.Toplevel(self)
        win.title("技能科目の扱い")
        win.transient(self)
        win.grab_set()
        win.resizable(True, True)

        win.grid_rowconfigure(0, weight=1)
        win.grid_columnconfigure(0, weight=1)

        container = ttk.Frame(win, padding=12)
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(2, weight=1)
        ttk.Label(container, text="①〜⑥の配置コマに対応する 技家 / 音美 の扱いを選択してください。").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(container, text="③・⑥は音美が無いため技家のみ選択します。").grid(
            row=1, column=0, sticky="w", pady=(0, 8)
        )

        content_outer = ttk.Frame(container)
        content_outer.grid(row=2, column=0, sticky="nsew")
        content_outer.grid_rowconfigure(0, weight=1)
        content_outer.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(content_outer, highlightthickness=0, borderwidth=0)
        ybar = ttk.Scrollbar(content_outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=ybar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")

        grid_frame = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=grid_frame, anchor="nw")
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.columnconfigure(1, weight=1)
        grid_frame.columnconfigure(2, weight=1)

        def _sync_dialog_layout(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfigure(window_id, width=max(canvas.winfo_width(), grid_frame.winfo_reqwidth()))

        grid_frame.bind("<Configure>", _sync_dialog_layout)
        canvas.bind("<Configure>", _sync_dialog_layout)

        override_vars: dict[tuple[str, str, str, str], tk.StringVar] = {}

        for idx, v in enumerate(variants):
            vid = v["id"]
            use_lower = v["use_lower"]
            sc_id = SCENARIO_ID_LOWER if use_lower else SCENARIO_ID_UPPER
            items = _collect_skill_items(sc_id)
            onbi_items = [x for x in items if x["subject"] == "音美"]
            tech_items = [x for x in items if x["subject"] == "技家"]

            r = idx // 3
            c = idx % 3
            frame = ttk.LabelFrame(grid_frame, text=f"{vid}")
            frame.grid(row=r, column=c, sticky="nsew", padx=6, pady=6)

            ttk.Label(frame, text="対象").grid(row=0, column=0, sticky="w", padx=(6, 12))
            ttk.Label(frame, text="選択").grid(row=0, column=1, sticky="w")

            row_idx = 1
            saved_cfg = saved.get(vid, {})
            onbi_overrides = saved_cfg.get("onbi_overrides", {}) or {}
            tech_overrides = saved_cfg.get("tech_overrides", {}) or {}
            onbi_default_map = _default_onbi_map(defaults[vid]["onbi_start_music"])
            tech_default = defaults[vid]["tech_subject"]

            def _add_row(label: str, options: list[str], default_value: str, key: tuple[str, str, str, str]) -> None:
                nonlocal row_idx
                ttk.Label(frame, text=label).grid(row=row_idx, column=0, sticky="w", padx=(6, 12))
                var = tk.StringVar(value=default_value)
                combo = ttk.Combobox(frame, textvariable=var, values=options, width=10, state="readonly")
                combo.grid(row=row_idx, column=1, sticky="w", padx=(0, 6))
                override_vars[key] = var
                row_idx += 1

            if not onbi_items and not tech_items:
                ttk.Label(frame, text="該当なし").grid(row=row_idx, column=0, sticky="w", padx=(6, 12))
                row_idx += 1
                continue

            for item in onbi_items:
                cls = item["class"]
                slot = item["slot"]
                key = f"{cls}|{slot}"
                default_value = onbi_overrides.get(key) or onbi_default_map.get(cls) or "音楽"
                _add_row(f"{cls} {slot} 音美", onbi_options, default_value, (vid, "音美", cls, slot))

            for item in tech_items:
                cls = item["class"]
                slot = item["slot"]
                key = f"{cls}|{slot}"
                default_value = tech_overrides.get(key) or tech_default
                _add_row(f"{cls} {slot} 技家", tech_options, default_value, (vid, "技家", cls, slot))

        buttons = ttk.Frame(container)
        buttons.grid(row=3, column=0, sticky="ew", pady=(12, 0))

        def _ok() -> None:
            nonlocal result
            settings: dict[str, dict[str, object]] = {}
            for vid in ["①", "②", "③", "④", "⑤", "⑥"]:
                cfg = {
                    "onbi_start_music": defaults[vid]["onbi_start_music"],
                    "tech_subject": defaults[vid]["tech_subject"],
                    "onbi_overrides": {},
                    "tech_overrides": {},
                }
                settings[vid] = cfg
            for (vid, subj, cls, slot), var in override_vars.items():
                key = f"{cls}|{slot}"
                if subj == "音美":
                    settings[vid]["onbi_overrides"][key] = var.get().strip()
                else:
                    settings[vid]["tech_overrides"][key] = var.get().strip()
            result = settings
            win.destroy()

        def _cancel() -> None:
            nonlocal result
            result = None
            win.destroy()

        ttk.Button(buttons, text="キャンセル", command=_cancel).pack(side="right", padx=(6, 0))
        ttk.Button(buttons, text="OK", command=_ok).pack(side="right")

        win.update_idletasks()
        req_w = max(container.winfo_reqwidth(), 1200)
        req_h = max(container.winfo_reqheight(), 480)
        screen_w = max(win.winfo_screenwidth() - 80, 800)
        screen_h = max(win.winfo_screenheight() - 120, 500)
        width = min(req_w, screen_w)
        height = min(req_h, screen_h)
        x = max((win.winfo_screenwidth() - width) // 2, 0)
        y = max((win.winfo_screenheight() - height) // 2, 0)
        win.geometry(f"{width}x{height}+{x}+{y}")
        win.minsize(min(req_w, 1200), min(req_h, 480))
        _sync_dialog_layout()

        win.wait_window()
        return result

    def run_generate(self) -> None:
        try:
            self.collect_from_ui()
        except Exception as e:
            messagebox.showerror("入力エラー", str(e))
            return

        if not messagebox.askyesno("最終確認", "この設定でExcelを生成しますか？"):
            return

        missing = self._find_empty_slots()
        if missing:
            preview = "\n".join(missing[:30])
            more = f"\n... 他 {len(missing) - 30} 件" if len(missing) > 30 else ""
            messagebox.showwarning(
                "空きコマ検出",
                "未割当のコマがあるためExcel出力を中止しました。\n"
                f"件数: {len(missing)}\n\n{preview}{more}",
            )
            return

        if self.config_data.get("year") == 6:
            settings = self._prompt_excel_variant_settings()
            if settings is None:
                return
            self.config_data["excel_variant_settings"] = settings

        work_cfg = Path("./_gui_current_config.json")
        try:
            write_json(work_cfg, self.config_data)
            run_solve(work_cfg)
            messagebox.showinfo("完了", f"生成完了:\n{self.config_data.get('output_path')}\n各クラスの担任や学活・総合・道徳は手動で割り当ててください")
        except PermissionError:
            messagebox.showerror("生成エラー", "Excelファイルが開かれているため保存できません。\nExcelを閉じてから再実行してください。")
        except Exception as e:
            messagebox.showerror("生成エラー", str(e))

    def _show_assign_diagnostics_popup(self, title: str, message: str) -> None:
        win = tk.Toplevel(self)
        win.title(title)
        win.geometry("980x620")
        win.transient(self)

        container = ttk.Frame(win, padding=10)
        container.pack(fill="both", expand=True)

        ttk.Label(
            container,
            text="割り当て診断結果（コピーして共有できます）",
            anchor="w",
        ).pack(fill="x", pady=(0, 8))

        text_frame = ttk.Frame(container)
        text_frame.pack(fill="both", expand=True)

        ybar = ttk.Scrollbar(text_frame, orient="vertical")
        ybar.pack(side="right", fill="y")

        txt = tk.Text(
            text_frame,
            wrap="word",
            yscrollcommand=ybar.set,
            font=("Meiryo UI", 10),
        )
        txt.pack(side="left", fill="both", expand=True)
        ybar.configure(command=txt.yview)

        txt.insert("1.0", str(message).strip() or "(診断メッセージなし)")
        txt.configure(state="disabled")

        btns = ttk.Frame(container)
        btns.pack(fill="x", pady=(8, 0))

        def _copy_all() -> None:
            body = txt.get("1.0", "end-1c")
            self.clipboard_clear()
            self.clipboard_append(body)

        ttk.Button(btns, text="全文コピー", command=_copy_all).pack(side="left")
        ttk.Button(btns, text="閉じる", command=win.destroy).pack(side="right")

        win.grab_set()


def main() -> None:
    app = TimetableGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
