#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import csv
import os
import sys
from dataclasses import dataclass
from datetime import date
from typing import Callable, Dict, List, Tuple

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QHBoxLayout,
    QVBoxLayout,
    QGridLayout,
    QGroupBox,
    QLabel,
    QLineEdit,
    QComboBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QHeaderView,
    QFrame,
    QFileDialog,
    QScrollArea,
    QSplitter,
    QSizePolicy,
    QCheckBox,
)

# ============================================================
# App metadata / branding
# ============================================================

APP_NAME = "Automation Roadmap Tool"
APP_VERSION = "v1.1.1"
RELEASE_DATE = "2026-01-17"  # YYYY-MM-DD

LOGO_PATH = "vesuvius_logo.png"


def resource_path(relative_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base, relative_path)


# ============================================================
# Model
# ============================================================

PHASES = ["Never", "Phase 1", "Phase 2", "Phase 3"]


@dataclass
class Inputs:
    kk: str
    heat_per_day: float
    plate_life: float
    cnt_life: float
    in_life: float
    pp_life: float
    o2_success: float
    working_days_year: float
    working_days_month: float


@dataclass
class OperationDef:
    name: str
    ops_per_day_fn: Callable[[Inputs], float]
    default_time_min: float
    default_phase: str
    default_cost_k_eur: float


@dataclass
class PhaseResults:
    remaining_h_per_day: float
    saving_h_per_day: float
    saving_pct: float
    saving_h_per_month: float
    saving_h_per_year: float
    solutions_used: List[str]
    investment_k_eur_total: float
    investment_k_eur_incremental: float


def ops_definitions() -> List[OperationDef]:
    def cylinder(inputs: Inputs) -> float:
        return inputs.heat_per_day * 2

    def tip_clean(inputs: Inputs) -> float:
        return 0.0 if inputs.cnt_life == 1 else inputs.heat_per_day

    def o2_lancing(inputs: Inputs) -> float:
        return inputs.heat_per_day

    def plate_inspection(inputs: Inputs) -> float:
        return inputs.heat_per_day

    def cnt_exchange(inputs: Inputs) -> float:
        return 0.0 if inputs.kk.lower() == "yes" else (inputs.heat_per_day / inputs.cnt_life)

    def plate_exchange(inputs: Inputs) -> float:
        return inputs.heat_per_day / inputs.plate_life

    def plate_cementing(inputs: Inputs) -> float:
        return inputs.heat_per_day / inputs.plate_life

    def in_bottom_clean(inputs: Inputs) -> float:
        return inputs.heat_per_day / inputs.plate_life

    def in_exchange(inputs: Inputs) -> float:
        return inputs.heat_per_day / inputs.in_life

    def pp_exchange(inputs: Inputs) -> float:
        return inputs.heat_per_day / inputs.pp_life

    return [
        OperationDef("Cylinder manipulation", cylinder, 1, "Phase 1", 100),
        OperationDef("CNT tip cleaning", tip_clean, 3, "Phase 2", 100),
        OperationDef("O₂ lancing", o2_lancing, 4, "Phase 1", 100),
        OperationDef("Plate inspection", plate_inspection, 1, "Phase 2", 100),
        OperationDef("CNT exchange", cnt_exchange, 3, "Phase 1", 100),
        OperationDef("Plate exchange", plate_exchange, 7, "Phase 1", 100),
        OperationDef("Plate cementing", plate_cementing, 2, "Phase 1", 100),
        OperationDef("IN & bottom plate surface cleaning", in_bottom_clean, 3, "Phase 3", 100),
        OperationDef("IN exchange", in_exchange, 15, "Phase 3", 100),
        OperationDef("PP exchange", pp_exchange, 15, "Phase 3", 100),
    ]


def phase_index(phase: str) -> int:
    if phase == "Phase 1":
        return 1
    if phase == "Phase 2":
        return 2
    if phase == "Phase 3":
        return 3
    return 999  # Never


def workload_h_per_day(ops_per_day: float, time_min: float) -> float:
    return ops_per_day * time_min / 60.0


def remaining_ops_for_phase(
    op_name: str,
    baseline_ops: float,
    selected_phase: str,
    phase_n: int,
    inputs: Inputs
) -> float:
    auto_at = phase_index(selected_phase)
    if phase_n < auto_at:
        return baseline_ops

    # automated
    if op_name == "O₂ lancing":
        return (1.0 - inputs.o2_success) * baseline_ops
    return 0.0


def compute_results(
    inputs: Inputs,
    times_min: Dict[str, float],
    phases: Dict[str, str],
    costs_k_eur: Dict[str, float],
    enabled: Dict[str, bool],
) -> Tuple[float, Dict[int, PhaseResults]]:
    defs = ops_definitions()

    baseline_h = 0.0
    remaining_h = {1: 0.0, 2: 0.0, 3: 0.0}

    # Solutions used up to each phase (cumulative)
    solutions_up_to: Dict[int, List[str]] = {1: [], 2: [], 3: []}
    for phase_n in (1, 2, 3):
        used = []
        for d in defs:
            if not enabled.get(d.name, True):
                continue
            sel_phase = phases.get(d.name, d.default_phase)
            if phase_index(sel_phase) <= phase_n:
                used.append(d.name)
        solutions_up_to[phase_n] = used

    # Investment per phase (cumulative) — only enabled
    invest_total = {1: 0.0, 2: 0.0, 3: 0.0}
    for phase_n in (1, 2, 3):
        s = 0.0
        for d in defs:
            if not enabled.get(d.name, True):
                continue
            sel_phase = phases.get(d.name, d.default_phase)
            c = float(costs_k_eur.get(d.name, d.default_cost_k_eur))
            if phase_index(sel_phase) <= phase_n:
                s += c
        invest_total[phase_n] = s

    for d in defs:
        if not enabled.get(d.name, True):
            continue

        ops = float(d.ops_per_day_fn(inputs))
        tmin = float(times_min.get(d.name, d.default_time_min))
        baseline_h += workload_h_per_day(ops, tmin)

        sel_phase = phases.get(d.name, d.default_phase)
        for phase_n in (1, 2, 3):
            rem_ops = remaining_ops_for_phase(d.name, ops, sel_phase, phase_n, inputs)
            remaining_h[phase_n] += workload_h_per_day(rem_ops, tmin)

    phase_results: Dict[int, PhaseResults] = {}
    for phase_n in (1, 2, 3):
        rem = remaining_h[phase_n]
        saving_h = baseline_h - rem
        saving_pct = (saving_h / baseline_h) if baseline_h > 0 else 0.0

        saving_month = saving_h * inputs.working_days_month
        saving_year = saving_h * inputs.working_days_year

        inv_total_k = invest_total[phase_n]
        inv_incr_k = inv_total_k - (invest_total[phase_n - 1] if phase_n > 1 else 0.0)

        phase_results[phase_n] = PhaseResults(
            remaining_h_per_day=rem,
            saving_h_per_day=saving_h,
            saving_pct=saving_pct,
            saving_h_per_month=saving_month,
            saving_h_per_year=saving_year,
            solutions_used=solutions_up_to[phase_n],
            investment_k_eur_total=inv_total_k,
            investment_k_eur_incremental=inv_incr_k,
        )

    return baseline_h, phase_results


# ============================================================
# UI helpers
# ============================================================

def big_font(size: int, bold: bool = True) -> QFont:
    f = QFont()
    f.setPointSize(size)
    f.setBold(bold)
    return f


def make_card(title: str) -> Tuple[QGroupBox, QLabel, QLabel, QLabel, QLabel, QLabel, QLabel]:
    box = QGroupBox(title)
    layout = QVBoxLayout(box)
    layout.setSpacing(8)

    main = QLabel("-")
    main.setFont(big_font(22, True))

    remaining = QLabel("Remaining workload: -")
    saving = QLabel("Workload reduction: -")
    extrap = QLabel("Extrapolation: -")
    invest = QLabel("Investment: -")
    solutions = QLabel("Solutions used: -")
    solutions.setWordWrap(True)

    for w in (remaining, saving, extrap, invest, solutions):
        w.setFont(big_font(11, False))

    layout.addWidget(main)
    layout.addWidget(remaining)
    layout.addWidget(saving)

    line = QFrame()
    line.setFrameShape(QFrame.Shape.HLine)
    line.setFrameShadow(QFrame.Shadow.Sunken)
    layout.addWidget(line)

    layout.addWidget(extrap)
    layout.addWidget(invest)
    layout.addWidget(solutions)

    return box, main, remaining, saving, extrap, invest, solutions


def fmt_h_day(x: float) -> str:
    return f"{x:.2f} h/day"


def fmt_h(x: float) -> str:
    return f"{x:.1f} h"


def fmt_pct(x: float) -> str:
    return f"{x * 100:.0f} %"


def fmt_k(x: float) -> str:
    return f"{x:.0f} k€"


# ============================================================
# Main Window
# ============================================================

class MainWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()

        self.setWindowTitle(f"{APP_NAME} — {APP_VERSION} ({RELEASE_DATE})")
        self.setMinimumSize(1180, 720)

        self._last_inputs: Inputs | None = None
        self._last_times: Dict[str, float] | None = None
        self._last_phases: Dict[str, str] | None = None
        self._last_costs: Dict[str, float] | None = None
        self._last_enabled: Dict[str, bool] | None = None
        self._last_baseline_h: float | None = None
        self._last_phase_res: Dict[int, PhaseResults] | None = None

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # Header with logo + title + version
        header = QHBoxLayout()
        header.setSpacing(10)

        self.logo = QLabel()
        self.logo.setFixedHeight(52)
        self.logo.setFixedWidth(160)  # ensures space for the logo (fixes your "empty slot")
        self.logo.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.logo.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self._load_logo()

        title_box = QVBoxLayout()
        title_lbl = QLabel(APP_NAME)
        title_lbl.setFont(big_font(16, True))
        meta_lbl = QLabel(f"{APP_VERSION} • Release date: {RELEASE_DATE}")
        meta_lbl.setStyleSheet("color: #666;")
        title_box.addWidget(title_lbl)
        title_box.addWidget(meta_lbl)

        header.addWidget(self.logo, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        header.addLayout(title_box, 1)

        # Export button in header (right)
        self.btn_export = QPushButton("Export CSV")
        self.btn_export.clicked.connect(self.export_csv)
        header.addWidget(self.btn_export, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        root.addLayout(header)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter, 1)

        # LEFT (Scenario) — scrollable if screen is small
        left_container = QWidget()
        left_layout = QVBoxLayout(left_container)
        left_layout.setSpacing(12)

        # Process assumptions
        proc_box = QGroupBox("Scenario builder – Process assumptions")
        proc_layout = QGridLayout(proc_box)
        proc_layout.setHorizontalSpacing(10)
        proc_layout.setVerticalSpacing(8)

        self.kk = QComboBox()
        self.kk.addItems(["no", "yes"])

        self.heat = QLineEdit("20")
        self.plate = QLineEdit("2")
        self.cnt = QLineEdit("1")
        self.inlife = QLineEdit("9")
        self.pplife = QLineEdit("20")
        self.o2 = QLineEdit("0.95")

        # NEW: show/hide costs toggle
        self.chk_show_costs = QCheckBox("Show cost & investment")
        self.chk_show_costs.setChecked(True)
        self.chk_show_costs.stateChanged.connect(self.on_toggle_costs)

        rows = [
            ("KK", self.kk, "CNT exchange = 0 when KK=yes"),
            ("Heat/day", self.heat, "Daily production volume"),
            ("Plate life", self.plate, "Heats per plate"),
            ("CNT life", self.cnt, "Heats per CNT"),
            ("IN life", self.inlife, "Heats per IN"),
            ("PP life", self.pplife, "Heats per PP"),
            ("O₂ success rate", self.o2, "Residual manual rate for O₂ lancing"),
        ]
        for r, (lab, w, hint) in enumerate(rows):
            proc_layout.addWidget(QLabel(lab), r, 0)
            proc_layout.addWidget(w, r, 1)
            h = QLabel(hint)
            h.setStyleSheet("color: #666;")
            proc_layout.addWidget(h, r, 2)

        # place checkbox below
        proc_layout.addWidget(self.chk_show_costs, len(rows), 0, 1, 3)

        left_layout.addWidget(proc_box)

        # Extrapolation assumptions
        extra_box = QGroupBox("Scenario builder – Extrapolation")
        extra_layout = QGridLayout(extra_box)
        extra_layout.setHorizontalSpacing(10)
        extra_layout.setVerticalSpacing(8)

        self.days_month = QLineEdit("22")
        self.days_year = QLineEdit("250")

        extra_layout.addWidget(QLabel("Working days / month"), 0, 0)
        extra_layout.addWidget(self.days_month, 0, 1)
        extra_layout.addWidget(QLabel("Default: 22"), 0, 2)

        extra_layout.addWidget(QLabel("Working days / year"), 1, 0)
        extra_layout.addWidget(self.days_year, 1, 1)
        extra_layout.addWidget(QLabel("Default: 250"), 1, 2)

        left_layout.addWidget(extra_box)

        # Automation scope (Use / Function / Phase / Time / Cost)
        scope_box = QGroupBox("Scenario builder – Automation scope")
        scope_layout = QVBoxLayout(scope_box)
        scope_layout.setSpacing(8)

        self.scope_table = QTableWidget()
        # NEW: first column is "Use"
        self.scope_table.setColumnCount(5)
        self.scope_table.setHorizontalHeaderLabels(["Use", "Function", "Automation phase", "Time/op [min]", "Cost [k€]"])
        self.scope_table.setAlternatingRowColors(True)
        self.scope_table.verticalHeader().setVisible(False)

        self.scope_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.scope_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.scope_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.scope_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.scope_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)

        scope_layout.addWidget(self.scope_table)

        btn_row = QHBoxLayout()
        self.btn_calc = QPushButton("Recalculate")
        self.btn_calc.clicked.connect(self.on_calculate)

        self.btn_reset = QPushButton("Reset defaults")
        self.btn_reset.clicked.connect(self.reset_defaults)

        btn_row.addWidget(self.btn_calc)
        btn_row.addWidget(self.btn_reset)
        btn_row.addStretch(1)
        scope_layout.addLayout(btn_row)

        left_layout.addWidget(scope_box)
        left_layout.addStretch(1)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setWidget(left_container)
        splitter.addWidget(left_scroll)

        # RIGHT (Executive results)
        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setSpacing(12)

        title = QLabel("Executive results – Workload & investment by phase")
        title.setFont(big_font(14, True))
        right_layout.addWidget(title)

        base_box = QGroupBox("Baseline")
        base_layout = QVBoxLayout(base_box)
        base_layout.setSpacing(8)
        self.lbl_baseline_main = QLabel("-")
        self.lbl_baseline_main.setFont(big_font(26, True))
        self.lbl_baseline_sub = QLabel("Manual workload today")
        self.lbl_baseline_sub.setStyleSheet("color: #666;")
        base_layout.addWidget(self.lbl_baseline_main)
        base_layout.addWidget(self.lbl_baseline_sub)
        right_layout.addWidget(base_box)

        self.card_p1 = make_card("Phase 1 – Minimal viable automation")
        self.card_p2 = make_card("Phase 2 – Extended automation")
        self.card_p3 = make_card("Phase 3 – Full automation")

        right_layout.addWidget(self.card_p1[0])
        right_layout.addWidget(self.card_p2[0])
        right_layout.addWidget(self.card_p3[0])

        self.lbl_note = QLabel("Note: Investment is cumulative by phase. “Solutions used” lists what is automated up to that phase.")
        self.lbl_note.setStyleSheet("color: #666;")
        right_layout.addWidget(self.lbl_note)
        right_layout.addStretch(1)

        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        right_scroll.setWidget(right_container)
        splitter.addWidget(right_scroll)

        splitter.setSizes([600, 580])

        # Table widgets maps
        self._use_widgets: Dict[str, QCheckBox] = {}
        self._phase_widgets: Dict[str, QComboBox] = {}
        self._time_items: Dict[str, QTableWidgetItem] = {}
        self._cost_items: Dict[str, QTableWidgetItem] = {}

        self.populate_scope_table()

        # Ensure the scope table is fully visible (no vertical scrolling needed for the rows)
        self._resize_scope_table_to_content()

        self.on_toggle_costs()
        self.on_calculate()

    # ----------------------------
    # Branding
    # ----------------------------

    def _load_logo(self) -> None:
        p = resource_path(LOGO_PATH)
        if os.path.exists(p):
            pm = QPixmap(p)
            if not pm.isNull():
                # scale nicely inside the fixed slot
                pm = pm.scaled(150, 48, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                self.logo.setPixmap(pm)
                self.logo.setToolTip("Vesuvius")
                return
        self.logo.setText("")

    # ----------------------------
    # UI setup
    # ----------------------------

    def populate_scope_table(self) -> None:
        defs = ops_definitions()
        self.scope_table.setRowCount(len(defs))

        for r, d in enumerate(defs):
            # Use checkbox
            chk = QCheckBox()
            chk.setChecked(True)
            chk.stateChanged.connect(lambda _=None: self.on_calculate())
            self.scope_table.setCellWidget(r, 0, chk)
            self._use_widgets[d.name] = chk

            # Function name
            name_item = QTableWidgetItem(d.name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.scope_table.setItem(r, 1, name_item)

            # Phase
            phase_cb = QComboBox()
            phase_cb.addItems(PHASES)
            phase_cb.setCurrentText(d.default_phase)
            phase_cb.currentIndexChanged.connect(lambda _=None: self.on_calculate())
            self.scope_table.setCellWidget(r, 2, phase_cb)
            self._phase_widgets[d.name] = phase_cb

            # Time
            time_item = QTableWidgetItem(f"{d.default_time_min:g}")
            time_item.setTextAlignment(int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter))
            self.scope_table.setItem(r, 3, time_item)
            self._time_items[d.name] = time_item

            # Cost
            cost_item = QTableWidgetItem(f"{d.default_cost_k_eur:g}")
            cost_item.setTextAlignment(int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter))
            self.scope_table.setItem(r, 4, cost_item)
            self._cost_items[d.name] = cost_item

        self.scope_table.resizeRowsToContents()

    def _resize_scope_table_to_content(self) -> None:
        """Make the table tall enough to show all rows without scrolling."""
        header_h = self.scope_table.horizontalHeader().height()
        rows_h = 0
        for r in range(self.scope_table.rowCount()):
            rows_h += self.scope_table.rowHeight(r)
        # add a small margin
        total = header_h + rows_h + 6
        self.scope_table.setMinimumHeight(total)
        self.scope_table.setMaximumHeight(total)

    def reset_defaults(self) -> None:
        defs = ops_definitions()
        for d in defs:
            self._use_widgets[d.name].setChecked(True)
            self._phase_widgets[d.name].setCurrentText(d.default_phase)
            self._time_items[d.name].setText(f"{d.default_time_min:g}")
            self._cost_items[d.name].setText(f"{d.default_cost_k_eur:g}")

        self.days_month.setText("22")
        self.days_year.setText("250")
        self.chk_show_costs.setChecked(True)

        self._resize_scope_table_to_content()
        self.on_calculate()

    # ----------------------------
    # Costs toggle
    # ----------------------------

    def on_toggle_costs(self) -> None:
        show = self.chk_show_costs.isChecked()

        # hide/show Cost column
        self.scope_table.setColumnHidden(4, not show)

        # hide/show investment line on cards
        for card in (self.card_p1, self.card_p2, self.card_p3):
            invest_label = card[5]
            invest_label.setVisible(show)

        self.lbl_note.setVisible(show)
        self.on_calculate()

    # ----------------------------
    # Reading inputs
    # ----------------------------

    def _f(self, w: QLineEdit) -> float:
        return float(w.text().strip())

    def read_inputs(self) -> Inputs:
        return Inputs(
            kk=self.kk.currentText().strip().lower(),
            heat_per_day=self._f(self.heat),
            plate_life=self._f(self.plate),
            cnt_life=self._f(self.cnt),
            in_life=self._f(self.inlife),
            pp_life=self._f(self.pplife),
            o2_success=self._f(self.o2),
            working_days_year=self._f(self.days_year),
            working_days_month=self._f(self.days_month),
        )

    def read_scope(self) -> Tuple[Dict[str, bool], Dict[str, float], Dict[str, str], Dict[str, float]]:
        enabled: Dict[str, bool] = {}
        times: Dict[str, float] = {}
        phases: Dict[str, str] = {}
        costs: Dict[str, float] = {}

        for name, chk in self._use_widgets.items():
            enabled[name] = chk.isChecked()

        for name, item in self._time_items.items():
            times[name] = float(item.text().strip())

        for name, cb in self._phase_widgets.items():
            phases[name] = cb.currentText()

        for name, item in self._cost_items.items():
            costs[name] = float(item.text().strip())

        return enabled, times, phases, costs

    # ----------------------------
    # Compute & update UI
    # ----------------------------

    def on_calculate(self) -> None:
        try:
            inputs = self.read_inputs()
            enabled, times, phases, costs = self.read_scope()

            if inputs.heat_per_day <= 0:
                raise ValueError("Heat/day must be > 0.")
            if not (0.0 <= inputs.o2_success <= 1.0):
                raise ValueError("O₂ success rate must be between 0 and 1.")
            for k, v in times.items():
                if v <= 0 and enabled.get(k, True):
                    raise ValueError(f"Time/op must be > 0 (check '{k}').")
            for k, v in costs.items():
                if v < 0 and enabled.get(k, True):
                    raise ValueError(f"Cost must be >= 0 (check '{k}').")
            for life_name, life_val in [
                ("Plate life", inputs.plate_life),
                ("CNT life", inputs.cnt_life),
                ("IN life", inputs.in_life),
                ("PP life", inputs.pp_life),
            ]:
                if life_val <= 0:
                    raise ValueError(f"{life_name} must be > 0.")
            if inputs.working_days_month <= 0 or inputs.working_days_year <= 0:
                raise ValueError("Working days/month and days/year must be > 0.")

            baseline_h, phase_res = compute_results(inputs, times, phases, costs, enabled)

        except Exception as e:
            QMessageBox.critical(self, "Input error", f"Could not calculate:\n{e}")
            return

        self._last_inputs = inputs
        self._last_times = times
        self._last_phases = phases
        self._last_costs = costs
        self._last_enabled = enabled
        self._last_baseline_h = baseline_h
        self._last_phase_res = phase_res

        self.lbl_baseline_main.setText(fmt_h_day(baseline_h))

        self._update_phase_card(self.card_p1, phase_res[1], phase_n=1)
        self._update_phase_card(self.card_p2, phase_res[2], phase_n=2)
        self._update_phase_card(self.card_p3, phase_res[3], phase_n=3)

    def _update_phase_card(self, card_tuple, res: PhaseResults, phase_n: int) -> None:
        _, main, remaining, saving, extrap, invest, solutions = card_tuple
        show_costs = self.chk_show_costs.isChecked()

        main.setText(f"{res.saving_h_per_day:.2f} h/day saved")
        remaining.setText(f"Remaining workload:   {fmt_h_day(res.remaining_h_per_day)}")
        saving.setText(f"Workload reduction:   {fmt_pct(res.saving_pct)}")
        extrap.setText(
            f"Extrapolation:        {fmt_h(res.saving_h_per_month)} / month   |   {fmt_h(res.saving_h_per_year)} / year"
        )

        if show_costs:
            if phase_n == 1:
                invest.setText(f"Investment:           {fmt_k(res.investment_k_eur_total)} total")
            else:
                invest.setText(
                    f"Investment:           {fmt_k(res.investment_k_eur_incremental)} incremental  |  {fmt_k(res.investment_k_eur_total)} total"
                )

        used = res.solutions_used
        solutions.setText("Solutions used:       " + (", ".join(used) if used else "(none)"))

    # ----------------------------
    # Export CSV
    # ----------------------------

    def export_csv(self) -> None:
        if not (self._last_inputs and self._last_times and self._last_phases and self._last_costs and self._last_phase_res and self._last_baseline_h is not None and self._last_enabled):
            QMessageBox.warning(self, "Export CSV", "No results available yet. Click Recalculate first.")
            return

        show_costs = self.chk_show_costs.isChecked()

        default_name = f"automation_roadmap_{APP_VERSION}_{RELEASE_DATE}.csv".replace(" ", "_")
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export results to CSV",
            default_name,
            "CSV Files (*.csv);;All Files (*)"
        )
        if not path:
            return

        i = self._last_inputs
        times = self._last_times
        phases = self._last_phases
        costs = self._last_costs
        enabled = self._last_enabled
        phase_res = self._last_phase_res
        baseline_h = self._last_baseline_h

        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow([APP_NAME, APP_VERSION, "Release", RELEASE_DATE])
                w.writerow(["Export date", date.today().isoformat()])
                w.writerow([])

                w.writerow(["SCENARIO INPUTS"])
                w.writerow(["KK", i.kk])
                w.writerow(["Heat/day", i.heat_per_day])
                w.writerow(["Plate life", i.plate_life])
                w.writerow(["CNT life", i.cnt_life])
                w.writerow(["IN life", i.in_life])
                w.writerow(["PP life", i.pp_life])
                w.writerow(["O2 success rate", i.o2_success])
                w.writerow(["Working days/month", i.working_days_month])
                w.writerow(["Working days/year", i.working_days_year])
                w.writerow([])
                w.writerow(["Baseline workload (h/day)", f"{baseline_h:.4f}"])
                w.writerow([])

                w.writerow(["OPERATIONS (scope)"])
                if show_costs:
                    w.writerow(["Use", "Function", "Phase", "Time/op [min]", "Cost [k€]"])
                else:
                    w.writerow(["Use", "Function", "Phase", "Time/op [min]"])

                for d in ops_definitions():
                    name = d.name
                    if show_costs:
                        w.writerow([
                            "yes" if enabled.get(name, True) else "no",
                            name,
                            phases.get(name, d.default_phase),
                            f"{times.get(name, d.default_time_min):.4f}",
                            f"{costs.get(name, d.default_cost_k_eur):.4f}",
                        ])
                    else:
                        w.writerow([
                            "yes" if enabled.get(name, True) else "no",
                            name,
                            phases.get(name, d.default_phase),
                            f"{times.get(name, d.default_time_min):.4f}",
                        ])
                w.writerow([])

                w.writerow(["RESULTS BY PHASE"])
                if show_costs:
                    w.writerow([
                        "Phase",
                        "Saving [h/day]",
                        "Remaining [h/day]",
                        "Reduction [%]",
                        "Saving [h/month]",
                        "Saving [h/year]",
                        "Investment incremental [k€]",
                        "Investment total [k€]",
                        "Solutions used"
                    ])
                else:
                    w.writerow([
                        "Phase",
                        "Saving [h/day]",
                        "Remaining [h/day]",
                        "Reduction [%]",
                        "Saving [h/month]",
                        "Saving [h/year]",
                        "Solutions used"
                    ])

                for p in (1, 2, 3):
                    r = phase_res[p]
                    if show_costs:
                        w.writerow([
                            f"Phase {p}",
                            f"{r.saving_h_per_day:.6f}",
                            f"{r.remaining_h_per_day:.6f}",
                            f"{(r.saving_pct*100):.2f}",
                            f"{r.saving_h_per_month:.3f}",
                            f"{r.saving_h_per_year:.3f}",
                            f"{r.investment_k_eur_incremental:.3f}",
                            f"{r.investment_k_eur_total:.3f}",
                            "; ".join(r.solutions_used),
                        ])
                    else:
                        w.writerow([
                            f"Phase {p}",
                            f"{r.saving_h_per_day:.6f}",
                            f"{r.remaining_h_per_day:.6f}",
                            f"{(r.saving_pct*100):.2f}",
                            f"{r.saving_h_per_month:.3f}",
                            f"{r.saving_h_per_year:.3f}",
                            "; ".join(r.solutions_used),
                        ])

            QMessageBox.information(self, "Export CSV", f"Exported successfully:\n{path}")

        except Exception as e:
            QMessageBox.critical(self, "Export CSV", f"Failed to export:\n{e}")


# ============================================================
# Run
# ============================================================

def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    w = MainWindow()
    w.resize(1320, 780)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

