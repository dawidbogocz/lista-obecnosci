#!/usr/bin/env python3
"""
Attendance Sheet App (Lista Obecności)
PySide6 GUI desktop app — monthly attendance with DOCX export.
"""

import sys
import os
import re
from datetime import date, timedelta
from calendar import monthrange, day_name, month_name

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QLineEdit, QPushButton, QComboBox,
    QScrollArea, QMessageBox, QFileDialog, QSpinBox,
    QFrame, QSizePolicy, QCheckBox
)
from PySide6.QtCore import Qt, QDate, QSize, Signal, QRectF, QPointF, QRegularExpression
from PySide6.QtGui import (
    QPainter, QPen, QColor, QFont, QImage, QPixmap, QAction,
    QPainterPath, QPageSize, QTransform, QRegularExpressionValidator
)

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml


# ─────────────────────────────────────────────
# Polish holidays
# ─────────────────────────────────────────────

def _easter(year: int) -> date:
    """Computes Easter Sunday using the Anonymous Gregorian algorithm."""
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def polish_holidays(year: int) -> set:
    """Return set of dates that are Polish official non-working days."""
    easter = _easter(year)
    fixed = {
        date(year, 1, 1),   # Nowy Rok
        date(year, 1, 6),   # Święto Trzech Króli
        date(year, 5, 1),   # Święto Pracy
        date(year, 5, 3),   # Święto Konstytucji 3 Maja
        date(year, 8, 15),  # Wniebowzięcie NMP
        date(year, 11, 1),  # Wszystkich Świętych
        date(year, 11, 11), # Narodowe Święto Niepodległości
        date(year, 12, 25), # Boże Narodzenie
        date(year, 12, 26), # Boże Narodzenie (drugi dzień)
    }
    movable = {
        easter,                                  # Wielkanoc
        easter + timedelta(days=1),              # Poniedziałek Wielkanocny
        easter + timedelta(days=49),             # Zielone Świątki (Pentecost)
        easter + timedelta(days=60),             # Boże Ciało (Corpus Christi)
    }
    return fixed | movable


def holiday_name(d: date) -> str:
    """Return human-readable name of a Polish holiday, or empty string."""
    holidays = {
        (1, 1): "Nowy Rok",
        (6, 1): "Święto Trzech Króli",
        (1, 5): "Święto Pracy",
        (3, 5): "Święto Konstytucji 3 Maja",
        (15, 8): "Wniebowzięcie NMP",
        (1, 11): "Wszystkich Świętych",
        (11, 11): "Narodowe Święto Niepodległości",
        (25, 12): "Boże Narodzenie",
        (26, 12): "Boże Narodzenie (drugi dzień)",
    }
    key = (d.day, d.month)
    if key in holidays:
        return holidays[key]
    easter = _easter(d.year)
    if d == easter:
        return "Wielkanoc"
    if d == easter + timedelta(days=1):
        return "Poniedziałek Wielkanocny"
    if d == easter + timedelta(days=49):
        return "Zielone Świątki"
    if d == easter + timedelta(days=60):
        return "Boże Ciało"
    return ""


# ─────────────────────────────────────────────
# Status constants
# ─────────────────────────────────────────────

STATUS_OPTIONS = [
    ("", "—"),
    ("obecny", "Obecny"),
    ("home_office", "Home Office"),
    ("urlop", "Urlop"),
    ("wolne_swieto", "Wolne za święto"),
    ("nieobecny", "Nieobecny"),
    ("inne", "Inne"),
]

STATUS_WEEKEND = "dzień wolny"
STATUS_HOLIDAY = "wolne za święto"

WEEKEND_COLOR = QColor(240, 240, 240)
HOLIDAY_COLOR = QColor(252, 228, 214)

# Time regex validator for HH:MM format
TIME_REGEX = QRegularExpression("^([01]?[0-9]|2[0-3]):[0-5][0-9]$")


# ─────────────────────────────────────────────
# Day row widget
# ─────────────────────────────────────────────

class DayRow(QFrame):
    """A single row for one day in the attendance sheet."""

    def __init__(self, day_date: date, is_holiday: bool, holiday_name_str: str = "", parent=None):
        super().__init__(parent)
        self.day_date = day_date
        self._is_weekend = day_date.weekday() >= 5
        self._is_holiday = is_holiday
        self._holiday_name = holiday_name_str

        self.setFrameStyle(QFrame.Shape.NoFrame)
        self.setMinimumHeight(32)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 1, 4, 1)
        layout.setSpacing(6)

        # Day number + name
        day_name_pl = {
            0: "Pon", 1: "Wt", 2: "Śr", 3: "Czw", 4: "Pt", 5: "Sob", 6: "Niedz"
        }
        dn = day_name_pl[day_date.weekday()]
        self.date_label = QLabel(f"{day_date.day:02d} ({dn})")
        self.date_label.setMinimumWidth(75)
        self.date_label.setFont(QFont("Arial", 9, QFont.Weight.Bold))
        layout.addWidget(self.date_label)

        # Status combo
        self.status_combo = QComboBox()
        for val, label in STATUS_OPTIONS:
            self.status_combo.addItem(label, val)
        self.status_combo.setMinimumWidth(150)
        self.status_combo.currentIndexChanged.connect(self._on_status_changed)
        layout.addWidget(self.status_combo)

        # Entry time — QLineEdit with input mask for natural typing
        self.time_in = QLineEdit("08:00")
        self.time_in.setInputMask("99:99")
        self.time_in.setPlaceholderText("HH:MM")
        self.time_in.setMinimumWidth(70)
        self.time_in.setMaxLength(5)
        self.time_in.setEnabled(True)
        layout.addWidget(self.time_in)

        # Separator
        sep = QLabel("→")
        sep.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sep.setMaximumWidth(20)
        layout.addWidget(sep)

        # Exit time
        self.time_out = QLineEdit("16:00")
        self.time_out.setInputMask("99:99")
        self.time_out.setPlaceholderText("HH:MM")
        self.time_out.setMinimumWidth(70)
        self.time_out.setMaxLength(5)
        self.time_out.setEnabled(True)
        layout.addWidget(self.time_out)

        # Location
        self.location_edit = QLineEdit()
        self.location_edit.setPlaceholderText("Lokacja")
        self.location_edit.setMinimumWidth(120)
        self.location_edit.setText("Tychy")
        layout.addWidget(self.location_edit)

        # Apply holiday auto-fill if needed
        if is_holiday:
            self._apply_holiday()

        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

    def _on_status_changed(self, idx):
        val = self.status_combo.currentData()
        if val in ("urlop", "wolne_swieto", "nieobecny"):
            self.time_in.setEnabled(False)
            self.time_out.setEnabled(False)
            self.location_edit.setEnabled(False)
        elif val == "home_office":
            self.time_in.setEnabled(True)
            self.time_out.setEnabled(True)
            self.location_edit.setEnabled(False)
            self.location_edit.setText("")
        else:  # obecny, inne, empty
            self.time_in.setEnabled(True)
            self.time_out.setEnabled(True)
            self.location_edit.setEnabled(True)

    def _apply_holiday(self):
        # Set to "wolne za święto" and show holiday name
        for i in range(self.status_combo.count()):
            if self.status_combo.itemData(i) == "wolne_swieto":
                self.status_combo.setCurrentIndex(i)
                break
        self.status_combo.setEnabled(False)
        self.time_in.setEnabled(False)
        self.time_out.setEnabled(False)
        self.location_edit.setEnabled(False)
        if self._holiday_name:
            self.location_edit.setText(self._holiday_name)
        self.setStyleSheet(f"background-color: {HOLIDAY_COLOR.name()}; border-radius: 2px;")

    def set_present_defaults(self):
        """Fill this row with default 'present' values."""
        self.status_combo.setEnabled(True)
        for i in range(self.status_combo.count()):
            if self.status_combo.itemData(i) == "obecny":
                self.status_combo.setCurrentIndex(i)
                break
        self.time_in.setText("08:00")
        self.time_in.setEnabled(True)
        self.time_out.setText("16:00")
        self.time_out.setEnabled(True)
        self.location_edit.setEnabled(True)
        self.location_edit.setText("Tychy")
        self.date_label.setStyleSheet("")
        self.setStyleSheet("")

    def is_workday(self) -> bool:
        """True if this day is a weekday (not weekend) and not a holiday."""
        return not self._is_weekend and not self._is_holiday

    def get_data(self) -> dict:
        return {
            "date": self.day_date,
            "status": self.status_combo.currentData(),
            "status_label": self.status_combo.currentText(),
            "time_in": self.time_in.text().strip(),
            "time_out": self.time_out.text().strip(),
            "location": self.location_edit.text().strip(),
            "is_weekend": self._is_weekend,
            "is_holiday": self._is_holiday,
            "holiday_name": self._holiday_name,
        }


# ─────────────────────────────────────────────
# Main window
# ─────────────────────────────────────────────

class AttendanceApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Lista Obecności")
        self.setMinimumSize(900, 700)
        self._day_rows = []
        self._setup_ui()

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(8)

        # ── Top bar: month picker + employee info ──
        top_bar = QHBoxLayout()

        # Month/Year
        self.month_spin = QSpinBox()
        self.month_spin.setRange(1, 12)
        self.month_spin.setValue(date.today().month)
        self.month_spin.setPrefix("Miesiąc: ")
        self.month_spin.setMinimumWidth(140)

        self.year_spin = QSpinBox()
        self.year_spin.setRange(2020, 2100)
        self.year_spin.setValue(date.today().year)
        self.year_spin.setPrefix("Rok: ")
        self.year_spin.setMinimumWidth(120)

        self.refresh_btn = QPushButton("Odśwież")
        self.refresh_btn.clicked.connect(self._rebuild_table)

        top_bar.addWidget(self.month_spin)
        top_bar.addWidget(self.year_spin)
        top_bar.addWidget(self.refresh_btn)
        top_bar.addStretch()

        # Employee fields
        top_bar.addWidget(QLabel("Imię i nazwisko:"))
        self.name_edit = QLineEdit("Dawid Bogocz")
        self.name_edit.setMinimumWidth(140)
        top_bar.addWidget(self.name_edit)

        top_bar.addWidget(QLabel("Dział:"))
        self.dept_edit = QLineEdit("Dział IT")
        self.dept_edit.setMinimumWidth(120)
        top_bar.addWidget(self.dept_edit)

        main_layout.addLayout(top_bar)

        # ── Separator ──
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)
        main_layout.addWidget(sep)

        # ── Table header ──
        header = QHBoxLayout()
        header.setSpacing(6)
        for text, w in [
            ("Data", 75), ("Status", 150), ("Wejście", 70),
            ("", 20), ("Wyjście", 70), ("Lokacja / Uwagi", 120)
        ]:
            lbl = QLabel(text)
            lbl.setFont(QFont("Arial", 9, QFont.Weight.Bold))
            lbl.setMinimumWidth(w)
            header.addWidget(lbl)
        main_layout.addLayout(header)

        # ── Scroll area with day rows ──
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)

        self.table_widget = QWidget()
        self.table_layout = QVBoxLayout(self.table_widget)
        self.table_layout.setSpacing(1)
        self.table_layout.setContentsMargins(0, 0, 0, 0)
        self.table_layout.addStretch()
        self.scroll.setWidget(self.table_widget)

        main_layout.addWidget(self.scroll, stretch=1)

        # ── Auto-fill and Export buttons ──
        btn_layout = QHBoxLayout()

        self.auto_fill_btn = QPushButton("⚡ Auto-fill workdays")
        self.auto_fill_btn.setMinimumHeight(36)
        self.auto_fill_btn.setStyleSheet("font-size: 13px; font-weight: bold; padding: 6px 16px;")
        self.auto_fill_btn.clicked.connect(self._auto_fill_workdays)
        btn_layout.addWidget(self.auto_fill_btn)

        btn_layout.addStretch()

        export_docx_btn = QPushButton("📄 Zapisz DOCX")
        export_docx_btn.setMinimumHeight(36)
        export_docx_btn.setStyleSheet("font-size: 14px; font-weight: bold; padding: 6px 20px;")
        export_docx_btn.clicked.connect(self._export_docx)
        btn_layout.addWidget(export_docx_btn)

        main_layout.addLayout(btn_layout)

        # Build initial table
        self._rebuild_table()

    def _rebuild_table(self):
        """Rebuild day rows for the selected month/year."""
        month = self.month_spin.value()
        year = self.year_spin.value()

        # Remove old rows
        for row in self._day_rows:
            self.table_layout.removeWidget(row)
            row.deleteLater()
        self._day_rows.clear()

        # Get holidays for this year
        holidays = polish_holidays(year)

        days_in_month = monthrange(year, month)[1]
        for day_num in range(1, days_in_month + 1):
            d = date(year, month, day_num)
            is_hol = d in holidays
            h_name = holiday_name(d) if is_hol else ""

            row = DayRow(d, is_hol, h_name)
            self._day_rows.append(row)
            # Insert before the stretch
            self.table_layout.insertWidget(self.table_layout.count() - 1, row)

    def _auto_fill_workdays(self):
        """Fill all workdays with 'Obecny', 8:00-16:00, Tychy."""
        for row in self._day_rows:
            if row.is_workday():
                row.set_present_defaults()
        QMessageBox.information(self, "Auto-fill", "Dni robocze wypełnione: Obecny, 8:00-16:00, Tychy")

    def _collect_data(self) -> list:
        """Collect all day data into a list of dicts."""
        return [row.get_data() for row in self._day_rows]

    # ──────────────────────────────────────────
    # DOCX export
    # ──────────────────────────────────────────

    def _export_docx(self):
        month = self.month_spin.value()
        year = self.year_spin.value()
        name = self.name_edit.text().strip() or "Pracownik"
        dept = self.dept_edit.text().strip() or ""

        default_filename = f"lista_obecnosci_{month:02d}-{year}.docx"
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Zapisz DOCX", os.path.expanduser(f"~/{default_filename}"),
            "Word (*.docx)"
        )
        if not filepath:
            return

        data = self._collect_data()

        try:
            self._build_docx(filepath, data, name, dept, month, year)
            QMessageBox.information(self, "Sukces", f"DOKUMENT ZAPISANY:\n{filepath}")
        except Exception as e:
            QMessageBox.critical(self, "Błąd", f"Nie udało się zapisać DOCX:\n{e}")

    def _build_docx(self, filepath, data, name, dept, month, year):
        doc = Document()

        # Set default font to support Polish characters
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(10)

        # ── Title ──
        month_names = [
            "", "Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
            "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"
        ]
        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run(f"LISTA OBECNOŚCI - {month_names[month]} {year}")
        run.bold = True
        run.font.size = Pt(16)
        run.font.name = 'Calibri'

        # ── Employee info ──
        info_line = doc.add_paragraph()
        info_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        info_text = name
        if dept:
            info_text += f" — {dept}"
        run = info_line.add_run(info_text)
        run.font.size = Pt(10)
        run.font.name = 'Calibri'

        # Spacer
        doc.add_paragraph()

        # ── Table ──
        day_names_short = ["Pon", "Wt", "Śr", "Czw", "Pt", "Sob", "Niedz"]

        table_data = []
        # Header row
        table_data.append(["Data", "Status", "Wejście", "Wyjście", "Lokacja / Uwagi"])

        for row_data in data:
            d = row_data["date"]
            date_str = f"{d.day:02d} {day_names_short[d.weekday()]}"

            status = row_data["status_label"]
            if row_data["is_weekend"] and not row_data["status"]:
                # Weekend with no status set by user — label it
                status = "Dzień wolny od pracy"
            elif row_data["is_holiday"] and row_data["status"] == "wolne_swieto":
                hname = row_data.get("holiday_name", "")
                if hname:
                    status = f"Wolne: {hname}"
                else:
                    status = "Wolne za święto"

            time_in = row_data["time_in"] if row_data["status"] in ("obecny", "home_office", "inne", "") else "—"
            time_out = row_data["time_out"] if row_data["status"] in ("obecny", "home_office", "inne", "") else "—"
            location = row_data.get("location", "")
            if row_data["is_weekend"] and not row_data["status"]:
                location = "—"
            elif row_data["is_holiday"] and row_data["holiday_name"]:
                location = row_data["holiday_name"]

            table_data.append([date_str, status, time_in, time_out, location])

        # Create table
        table = doc.add_table(rows=len(table_data), cols=5)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = 'Table Grid'

        # Set column widths
        page_width = Cm(18.5)  # A4 minus margins
        col_widths = [page_width * 0.16, page_width * 0.30, page_width * 0.12, page_width * 0.12, page_width * 0.30]

        for row_idx, row_data in enumerate(table_data):
            row = table.rows[row_idx]
            for col_idx, cell_text in enumerate(row_data):
                cell = row.cells[col_idx]
                cell.text = ""
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx != 4 else WD_ALIGN_PARAGRAPH.LEFT
                run = p.add_run(str(cell_text))
                run.font.size = Pt(9)
                run.font.name = 'Calibri'

                if row_idx == 0:
                    # Header row
                    run.bold = True
                    self._set_cell_shading(cell, "D9D9D9")
                else:
                    # Data rows - color weekends and holidays
                    data_idx = row_idx - 1
                    if data_idx < len(data):
                        r = data[data_idx]
                        if r["is_weekend"] and not r["status"]:
                            self._set_cell_shading(cell, "F2F2F2")
                        elif r["is_holiday"]:
                            self._set_cell_shading(cell, "FCE4D6")

        # No signature — as requested

        doc.save(filepath)

    def _set_cell_shading(self, cell, hex_color):
        """Set background color on a table cell."""
        shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        cell._tc.get_or_add_tcPr().append(shading_elm)


# ─────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = AttendanceApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()