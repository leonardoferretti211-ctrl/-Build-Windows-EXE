#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Jan 17 02:41:48 2026

@author: ferrettileonardo
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from typing import Callable, Dict, List, Tuple

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
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
)


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
    solutions_used: List[str]  # automated up to this phase (cumulative)
    investment_k_eur_total: float
    investment_k_eur_incremental: float


def ops_definitions() -> List[OperationDef]:
    # Baseline ops/day formulas (from your Excel)
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

    # Default times (min) from Excel; default costs 100 k€ as requested.
    # Default phases chosen to resemble your original sheet.
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
    """
    Generic automation rule:
    - before automation phase: remaining = baseline
    - at/after automation phase: remaining = 0
    Special case kept from Excel: O₂ lancing becomes residual = (1-success)*baseline when automated.
    """
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
) -> Tuple[float, Dict[int, PhaseResults]]:
    defs = ops_definitions()

    baseline_h = 0.0
    remaining_h = {1: 0.0, 2: 0.0, 3: 0.0}

    # Solutions used up to each phase (cumulative)
    solutions_up_to: Dict[int, List[str]] = {1: [], 2: [], 3: []}
    for phase_n in (1, 2, 3):
        used = []
        for d in defs:
            sel_phase = phases.get(d.name, d.default_phase)
            if phase_index(sel_phase) <= phase_n:
                used.append(d.name)
        solutions_up_to[phase_n] = used

    # Investment per phase (cumulative)
    invest_total = {1: 0.0, 2: 0.0, 3: 0.0}
    for phase_n in (1, 2, 3):
        s = 0.0
        for d in defs:
            sel_phase = phases.get(d.name, d.default_phase)
            c = float(costs_k_eur.get(d.name, d.default_cost_k_eur))
            if phase_index(sel_phase) <= phase_n:
                s += c
        invest_total[phase_n] = s

    for d in defs:
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
    """
    Returns a card widget and references to labels you can update:
      - main KPI label
      - remaining label
      - saving label
      - extrapolation label
      - investment label
      - solutions label
    """
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
        self.setWindowTitle("Automation Roadmap – Workload & Investment Tool")

        root = QHBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(14)

        # LEFT (Scenario)
        left = QVBoxLayout()
        left.setSpacing(12)
        root.addLayout(left, 1)

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

        left.addWidget(proc_box)

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

        left.addWidget(extra_box)

        # Automation scope (Function / Phase / Time / Cost)
        scope_box = QGroupBox("Scenario builder – Automation scope")
        scope_layout = QVBoxLayout(scope_box)
        scope_layout.setSpacing(8)

        self.scope_table = QTableWidget()
        self.scope_table.setColumnCount(4)
        self.scope_table.setHorizontalHeaderLabels(["Function", "Automation phase", "Time/op [min]", "Cost [k€]"])
        self.scope_table.setAlternatingRowColors(True)
        self.scope_table.verticalHeader().setVisible(False)
        self.scope_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.scope_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.scope_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.scope_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)

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

        left.addWidget(scope_box)
        left.addStretch(1)

        # RIGHT (Executive results)
        right = QVBoxLayout()
        right.setSpacing(12)
        root.addLayout(right, 1)

        title = QLabel("Executive results – Workload & investment by phase")
        title.setFont(big_font(16, True))
        right.addWidget(title)

        # Baseline card
        base_box = QGroupBox("Baseline")
        base_layout = QVBoxLayout(base_box)
        base_layout.setSpacing(8)
        self.lbl_baseline_main = QLabel("-")
        self.lbl_baseline_main.setFont(big_font(26, True))
        self.lbl_baseline_sub = QLabel("Manual workload today")
        self.lbl_baseline_sub.setStyleSheet("color: #666;")
        base_layout.addWidget(self.lbl_baseline_main)
        base_layout.addWidget(self.lbl_baseline_sub)
        right.addWidget(base_box)

        # Phase cards
        self.card_p1 = make_card("Phase 1 – Minimal viable automation")
        self.card_p2 = make_card("Phase 2 – Extended automation")
        self.card_p3 = make_card("Phase 3 – Full automation")

        right.addWidget(self.card_p1[0])
        right.addWidget(self.card_p2[0])
        right.addWidget(self.card_p3[0])

        self.lbl_note = QLabel(
            "Note: Investment is cumulative by phase. “Solutions used” lists what is automated up to that phase."
        )
        self.lbl_note.setStyleSheet("color: #666;")
        right.addWidget(self.lbl_note)

        self._phase_widgets: Dict[str, QComboBox] = {}
        self._time_items: Dict[str, QTableWidgetItem] = {}
        self._cost_items: Dict[str, QTableWidgetItem] = {}

        self.populate_scope_table()
        self.on_calculate()

    # ----------------------------
    # UI setup
    # ----------------------------

    def populate_scope_table(self) -> None:
        defs = ops_definitions()
        self.scope_table.setRowCount(len(defs))

        for r, d in enumerate(defs):
            # Function name
            name_item = QTableWidgetItem(d.name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.scope_table.setItem(r, 0, name_item)

            # Phase selector
            phase_cb = QComboBox()
            phase_cb.addItems(PHASES)
            phase_cb.setCurrentText(d.default_phase)
            self.scope_table.setCellWidget(r, 1, phase_cb)
            self._phase_widgets[d.name] = phase_cb

            # Time/op [min]
            time_item = QTableWidgetItem(f"{d.default_time_min:g}")
            time_item.setTextAlignment(int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter))
            self.scope_table.setItem(r, 2, time_item)
            self._time_items[d.name] = time_item

            # Cost [k€]
            cost_item = QTableWidgetItem(f"{d.default_cost_k_eur:g}")
            cost_item.setTextAlignment(int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter))
            self.scope_table.setItem(r, 3, cost_item)
            self._cost_items[d.name] = cost_item

        self.scope_table.resizeRowsToContents()

    def reset_defaults(self) -> None:
        defs = ops_definitions()
        for d in defs:
            self._phase_widgets[d.name].setCurrentText(d.default_phase)
            self._time_items[d.name].setText(f"{d.default_time_min:g}")
            self._cost_items[d.name].setText(f"{d.default_cost_k_eur:g}")
        self.days_month.setText("22")
        self.days_year.setText("250")
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

    def read_scope(self) -> Tuple[Dict[str, float], Dict[str, str], Dict[str, float]]:
        times: Dict[str, float] = {}
        phases: Dict[str, str] = {}
        costs: Dict[str, float] = {}

        for name, item in self._time_items.items():
            times[name] = float(item.text().strip())

        for name, cb in self._phase_widgets.items():
            phases[name] = cb.currentText()

        for name, item in self._cost_items.items():
            costs[name] = float(item.text().strip())

        return times, phases, costs

    # ----------------------------
    # Compute & update UI
    # ----------------------------

    def on_calculate(self) -> None:
        try:
            inputs = self.read_inputs()
            times, phases, costs = self.read_scope()

            # Validation
            if inputs.heat_per_day <= 0:
                raise ValueError("Heat/day must be > 0.")
            if not (0.0 <= inputs.o2_success <= 1.0):
                raise ValueError("O₂ success rate must be between 0 and 1.")
            for k, v in times.items():
                if v <= 0:
                    raise ValueError(f"Time/op must be > 0 (check '{k}').")
            for k, v in costs.items():
                if v < 0:
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

            baseline_h, phase_res = compute_results(inputs, times, phases, costs)
        except Exception as e:
            QMessageBox.critical(self, "Input error", f"Could not calculate:\n{e}")
            return

        # Baseline
        self.lbl_baseline_main.setText(fmt_h_day(baseline_h))

        # Update cards
        self._update_phase_card(self.card_p1, phase_res[1], phase_n=1)
        self._update_phase_card(self.card_p2, phase_res[2], phase_n=2)
        self._update_phase_card(self.card_p3, phase_res[3], phase_n=3)

    def _update_phase_card(self, card_tuple, res: PhaseResults, phase_n: int) -> None:
        _, main, remaining, saving, extrap, invest, solutions = card_tuple

        main.setText(f"{res.saving_h_per_day:.2f} h/day saved")

        remaining.setText(f"Remaining workload:   {fmt_h_day(res.remaining_h_per_day)}")
        saving.setText(f"Workload reduction:   {fmt_pct(res.saving_pct)}")

        extrap.setText(
            f"Extrapolation:        {fmt_h(res.saving_h_per_month)} / month   |   {fmt_h(res.saving_h_per_year)} / year"
        )

        if phase_n == 1:
            invest.setText(f"Investment:           {fmt_k(res.investment_k_eur_total)} total")
        else:
            invest.setText(
                f"Investment:           {fmt_k(res.investment_k_eur_incremental)} incremental  |  {fmt_k(res.investment_k_eur_total)} total"
            )

        used = res.solutions_used
        if not used:
            solutions.setText("Solutions used:       (none)")
        else:
            solutions.setText("Solutions used:       " + ", ".join(used))


# ============================================================
# Run
# ============================================================

def main() -> None:
    app = QApplication(sys.argv)
    w = MainWindow()
    w.resize(1350, 780)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()