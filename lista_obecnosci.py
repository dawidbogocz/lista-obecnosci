#!/usr/bin/env python3
"""
Attendance Sheet App (Lista Obecności)
PySide6 GUI — per-row signature, delegacja status, DOCX export.
"""

import sys
import os
from datetime import date, timedelta
from calendar import monthrange

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox,
    QScrollArea, QMessageBox, QFileDialog, QSpinBox,
    QFrame, QSizePolicy
)
from PySide6.QtCore import Qt
from PySide6.QtGui import (
    QPainter, QPen, QColor, QFont, QImage, QPainterPath
)

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml


# ─────────────────────────────────────────────
# Polish holidays
# ─────────────────────────────────────────────

def _easter(year: int) -> date:
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
    easter = _easter(year)
    fixed = {
        date(year, 1, 1), date(year, 1, 6),
        date(year, 5, 1), date(year, 5, 3), date(year, 8, 15),
        date(year, 11, 1), date(year, 11, 11),
        date(year, 12, 25), date(year, 12, 26),
    }
    movable = {easter, easter + timedelta(days=1),
               easter + timedelta(days=49), easter + timedelta(days=60)}
    return fixed | movable


def holiday_name(d: date) -> str:
    h = {
        (1, 1): "Nowy Rok", (6, 1): "Święto Trzech Króli",
        (1, 5): "Święto Pracy", (3, 5): "Święto Konstytucji 3 Maja",
        (15, 8): "Wniebowzięcie NMP", (1, 11): "Wszystkich Świętych",
        (11, 11): "Narodowe Święto Niepodległości",
        (25, 12): "Boże Narodzenie", (26, 12): "Boże Narodzenie (drugi dzień)",
    }
    key = (d.day, d.month)
    if key in h:
        return h[key]
    easter = _easter(d.year)
    if d == easter: return "Wielkanoc"
    if d == easter + timedelta(days=1): return "Poniedziałek Wielkanocny"
    if d == easter + timedelta(days=49): return "Zielone Świątki"
    if d == easter + timedelta(days=60): return "Boże Ciało"
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
    ("delegacja", "Delegacja"),
    ("nieobecny", "Nieobecny"),
    ("inne", "Inne"),
]

WEEKEND_COLOR = QColor(200, 215, 240)
HOLIDAY_COLOR = QColor(255, 210, 180)


# ─────────────────────────────────────────────
# Signature canvas
# ─────────────────────────────────────────────

class SignatureCanvas(QWidget):
    CANVAS_W = 400
    CANVAS_H = 120

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(self.CANVAS_W, self.CANVAS_H)
        self.setMaximumHeight(160)
        self.setStyleSheet("background-color: white; border: 1px solid #aaa; border-radius: 4px;")
        self._path = QPainterPath()
        self._has_strokes = False

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.fillRect(self.rect(), Qt.GlobalColor.white)
        pen = QPen(QColor(180, 180, 180), 1, Qt.PenStyle.DashLine)
        painter.setPen(pen)
        y = self.height() - 25
        painter.drawLine(10, y, self.width() - 10, y)
        painter.setPen(QColor(120, 120, 120))
        painter.setFont(QFont("Arial", 9))
        painter.drawText(12, y - 4, "Podpis:")
        if not self._path.isEmpty():
            pen2 = QPen(QColor(0, 0, 140), 2, Qt.PenStyle.SolidLine,
                        Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
            painter.setPen(pen2)
            painter.drawPath(self._path)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._path.moveTo(event.position())
            self._has_strokes = False
            self.update()

    def mouseMoveEvent(self, event):
        if event.buttons() & Qt.MouseButton.LeftButton:
            self._path.lineTo(event.position())
            self._has_strokes = True
            self.update()

    def clear_signature(self):
        self._path = QPainterPath()
        self._has_strokes = False
        self.update()

    def has_signature(self) -> bool:
        return self._has_strokes and not self._path.isEmpty()

    def render_for_docx(self, filepath: str) -> bool:
        """Render signature onto a white PNG. Simple and reliable."""
        if not self.has_signature():
            return False
        scale = 3
        w = self.CANVAS_W * scale
        h = self.CANVAS_H * scale
        img = QImage(w, h, QImage.Format.Format_RGB32)
        img.fill(Qt.GlobalColor.white)
        p = QPainter(img)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.scale(scale, scale)
        pen = QPen(QColor(0, 0, 140), 2, Qt.PenStyle.SolidLine,
                   Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
        p.setPen(pen)
        p.drawPath(self._path)
        p.end()
        return img.save(filepath, "PNG")


# ─────────────────────────────────────────────
# Day row
# ─────────────────────────────────────────────

class DayRow(QFrame):
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

        # Date label — dd.mm.rrrr format
        self.date_label = QLabel(day_date.strftime("%d.%m.%Y"))
        self.date_label.setMinimumWidth(85)
        self.date_label.setFont(QFont("Arial", 9, QFont.Weight.Bold))
        layout.addWidget(self.date_label)

        # Status
        self.status_combo = QComboBox()
        for val, label in STATUS_OPTIONS:
            self.status_combo.addItem(label, val)
        self.status_combo.setMinimumWidth(150)
        self.status_combo.currentIndexChanged.connect(self._on_status_changed)
        layout.addWidget(self.status_combo)

        # Location — only usable for delegacja
        self.location_edit = QLineEdit()
        self.location_edit.setPlaceholderText("Miejsce delegacji")
        self.location_edit.setMinimumWidth(180)
        self.location_edit.setEnabled(False)
        layout.addWidget(self.location_edit)

        if is_holiday:
            self._style_holiday()
        elif self._is_weekend:
            self._style_weekend()

        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

    def _on_status_changed(self, idx):
        val = self.status_combo.currentData()
        if val == "delegacja":
            self.location_edit.setEnabled(True)
        else:
            self.location_edit.setEnabled(False)
            if val != "obecny":
                self.location_edit.setText("")

    def _style_holiday(self):
        for i in range(self.status_combo.count()):
            if self.status_combo.itemData(i) == "wolne_swieto":
                self.status_combo.setCurrentIndex(i)
                break
        if self._holiday_name:
            self.location_edit.setText(self._holiday_name)
        self.setStyleSheet(f"background-color: {HOLIDAY_COLOR.name()}; border-radius: 2px;")

    def _style_weekend(self):
        self.setStyleSheet(f"background-color: {WEEKEND_COLOR.name()}; border-radius: 2px;")

    def set_present_defaults(self):
        for i in range(self.status_combo.count()):
            if self.status_combo.itemData(i) == "obecny":
                self.status_combo.setCurrentIndex(i)
                break
        self.location_edit.setEnabled(False)
        self.location_edit.setText("")
        self.setStyleSheet("")

    def is_workday(self) -> bool:
        return not self._is_weekend and not self._is_holiday

    def get_data(self) -> dict:
        return {
            "date": self.day_date,
            "status": self.status_combo.currentData(),
            "status_label": self.status_combo.currentText(),
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
        self.setMinimumSize(750, 700)
        self._day_rows = []
        self._setup_ui()

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(8)

        # Top bar
        top = QHBoxLayout()
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
        top.addWidget(self.month_spin)
        top.addWidget(self.year_spin)
        top.addWidget(self.refresh_btn)
        top.addStretch()
        top.addWidget(QLabel("Imię i nazwisko:"))
        self.name_edit = QLineEdit("Dawid Bogocz")
        self.name_edit.setMinimumWidth(140)
        top.addWidget(self.name_edit)
        top.addWidget(QLabel("Dział:"))
        self.dept_edit = QLineEdit("Dział IT")
        self.dept_edit.setMinimumWidth(120)
        top.addWidget(self.dept_edit)
        main_layout.addLayout(top)

        # Separator
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)
        main_layout.addWidget(sep)

        # Header
        header = QHBoxLayout()
        header.setSpacing(6)
        for text, w in [("Data", 85), ("Status", 150), ("Lokacja / Uwagi", 180)]:
            lbl = QLabel(text)
            lbl.setFont(QFont("Arial", 9, QFont.Weight.Bold))
            lbl.setMinimumWidth(w)
            header.addWidget(lbl)
        main_layout.addLayout(header)

        # Scroll
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

        # Signature
        sig_layout = QHBoxLayout()
        sig_layout.addWidget(QLabel("Podpis (zostanie dodany do każdego wiersza w dokumencie):"))
        self.sig_canvas = SignatureCanvas()
        sig_layout.addWidget(self.sig_canvas, stretch=1)
        clear_btn = QPushButton("Wyczyść")
        clear_btn.clicked.connect(self.sig_canvas.clear_signature)
        sig_layout.addWidget(clear_btn)
        main_layout.addLayout(sig_layout)

        # Buttons
        btn_layout = QHBoxLayout()
        self.auto_fill_btn = QPushButton("Auto-fill workdays")
        self.auto_fill_btn.setMinimumHeight(36)
        self.auto_fill_btn.setStyleSheet("font-size: 13px; font-weight: bold; padding: 6px 16px;")
        self.auto_fill_btn.clicked.connect(self._auto_fill_workdays)
        btn_layout.addWidget(self.auto_fill_btn)
        btn_layout.addStretch()
        export_btn = QPushButton("Zapisz DOCX")
        export_btn.setMinimumHeight(36)
        export_btn.setStyleSheet("font-size: 14px; font-weight: bold; padding: 6px 20px;")
        export_btn.clicked.connect(self._export_docx)
        btn_layout.addWidget(export_btn)
        main_layout.addLayout(btn_layout)

        self._rebuild_table()

    def _rebuild_table(self):
        month = self.month_spin.value()
        year = self.year_spin.value()
        for row in self._day_rows:
            self.table_layout.removeWidget(row)
            row.deleteLater()
        self._day_rows.clear()
        holidays = polish_holidays(year)
        for day_num in range(1, monthrange(year, month)[1] + 1):
            d = date(year, month, day_num)
            is_hol = d in holidays
            row = DayRow(d, is_hol, holiday_name(d) if is_hol else "")
            self._day_rows.append(row)
            self.table_layout.insertWidget(self.table_layout.count() - 1, row)

    def _auto_fill_workdays(self):
        filled = 0
        for row in self._day_rows:
            if row.is_workday():
                row.set_present_defaults()
                filled += 1
        QMessageBox.information(self, "Auto-fill",
                                f"Wypełniono {filled} dni: Obecny, Tychy")

    def _collect_data(self) -> list:
        return [row.get_data() for row in self._day_rows]

    # ──────────────────────────────────────────
    # DOCX export
    # ──────────────────────────────────────────

    def _export_docx(self):
        month = self.month_spin.value()
        year = self.year_spin.value()
        name = self.name_edit.text().strip() or "Pracownik"
        dept = self.dept_edit.text().strip() or ""

        fp, _ = QFileDialog.getSaveFileName(
            self, "Zapisz DOCX",
            os.path.expanduser(f"~/lista_obecnosci_{month:02d}-{year}.docx"),
            "Word (*.docx)")
        if not fp:
            return

        data = self._collect_data()
        sig_path = "/tmp/_attendance_sig_temp.png"
        has_sig = self.sig_canvas.render_for_docx(sig_path)

        # Verify the file was actually created
        if has_sig and not os.path.exists(sig_path):
            has_sig = False  # sanity check

        try:
            self._build_docx(fp, data, name, dept, month, year,
                             sig_path if has_sig else None)
            QMessageBox.information(self, "Sukces", f"DOKUMENT ZAPISANY:\n{fp}")
        except Exception as e:
            QMessageBox.critical(self, "Błąd", f"Nie udało się zapisać DOCX:\n{e}")

    def _build_docx(self, filepath, data, name, dept, month, year, sig_img_path=None):
        doc = Document()

        style = doc.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(10)

        # Title
        month_names = ["", "Styczeń", "Luty", "Marzec", "Kwiecień", "Maj",
                       "Czerwiec", "Lipiec", "Sierpień", "Wrzesień",
                       "Październik", "Listopad", "Grudzień"]
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"LISTA OBECNOŚCI - {month_names[month]} {year}")
        r.bold = True
        r.font.size = Pt(16)
        r.font.name = 'Calibri'

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = name + (f" — {dept}" if dept else "")
        r = p.add_run(txt)
        r.font.size = Pt(10)

        doc.add_paragraph()

        # Table: 4 columns — Data, Status, Podpis, Lokacja
        cols = ["Data", "Status", "Podpis", "Lokacja"]
        table = doc.add_table(rows=len(data) + 1, cols=4)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = 'Table Grid'

        # Header
        for ci, h in enumerate(cols):
            cell = table.rows[0].cells[ci]
            cell.text = ""
            r = cell.paragraphs[0].add_run(h)
            r.bold = True
            r.font.size = Pt(9)
            r.font.name = 'Calibri'
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._shade(cell, "D9D9D9")

        # Column widths
        cw = [Cm(2.5), Cm(4.0), Cm(4.5), Cm(7.5)]

        # Data rows
        days_short = ["Pon", "Wt", "Śr", "Czw", "Pt", "Sob", "Niedz"]

        for ri, rd in enumerate(data):
            d = rd["date"]
            date_str = d.strftime("%d.%m.%Y")

            status = rd["status_label"]
            if not rd["status"] and rd["is_weekend"]:
                status = "Dzień wolny od pracy"
            elif rd["is_holiday"] and rd["status"] == "wolne_swieto":
                hn = rd.get("holiday_name", "")
                status = f"Wolne: {hn}" if hn else "Wolne za święto"

            location = rd.get("location", "")
            if rd["is_holiday"] and rd["holiday_name"] and not rd["status"] in ("delegacja", "obecny"):
                location = rd["holiday_name"]

            # Should this row show the signature?
            show_sig = sig_img_path and rd["status"] in ("obecny", "home_office", "delegacja")

            values = [date_str, status, "", location]
            doc_row = table.rows[ri + 1]

            for ci in range(4):
                cell = doc_row.cells[ci]
                cell.text = ""
                par = cell.paragraphs[0]
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci != 3 else WD_ALIGN_PARAGRAPH.LEFT

                if ci == 2 and show_sig:
                    # Signature image
                    run = par.add_run()
                    run.add_picture(sig_img_path, width=Cm(4.0), height=Cm(1.2))
                elif ci == 2:
                    # Empty podpis cell — keep blank
                    pass
                else:
                    run = par.add_run(str(values[ci]))
                    run.font.size = Pt(9)
                    run.font.name = 'Calibri'

                # Shade weekends/holidays
                if rd["is_holiday"]:
                    self._shade(cell, "FCE4D6")
                elif rd["is_weekend"] and not rd["status"]:
                    self._shade(cell, "DAE8FC")

        doc.save(filepath)

    def _shade(self, cell, color):
        cell._tc.get_or_add_tcPr().append(
            parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>'))


# ─────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    w = AttendanceApp()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()