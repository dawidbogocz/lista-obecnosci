#!/usr/bin/env python3
"""
Attendance Sheet App (Lista Obecności)
Single-page DOCX output, per-row signatures, L4, delegacja, uwaga for inne.
"""

import sys
import os
import tempfile
from datetime import date, timedelta
from calendar import monthrange

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox,
    QScrollArea, QMessageBox, QFileDialog, QSpinBox,
    QFrame, QSizePolicy
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPainter, QPen, QColor, QFont, QImage, QPainterPath

from docx import Document
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml


# ─────────────────────────────────────────────
# Polish holidays
# ─────────────────────────────────────────────

def _easter(year: int) -> date:
    a = year % 19; b = year // 100; c = year % 100
    d = b // 4; e = b % 4; f = (b + 8) // 25
    g = (b - f + 1) // 3; h = (19 * a + b - d - g + 15) % 30
    i = c // 4; k = c % 4; l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def polish_holidays(year: int) -> set:
    e = _easter(year)
    fixed = {date(year, 1, 1), date(year, 1, 6), date(year, 5, 1),
             date(year, 5, 3), date(year, 8, 15), date(year, 11, 1),
             date(year, 11, 11), date(year, 12, 25), date(year, 12, 26)}
    mov = {e, e + timedelta(days=1), e + timedelta(days=49), e + timedelta(days=60)}
    return fixed | mov


def holiday_name(d: date) -> str:
    h = {(1, 1): "Nowy Rok", (6, 1): "Święto Trzech Króli", (1, 5): "Święto Pracy",
         (3, 5): "Święto Konstytucji 3 Maja", (15, 8): "Wniebowzięcie NMP",
         (1, 11): "Wszystkich Świętych", (11, 11): "Narodowe Święto Niepodległości",
         (25, 12): "Boże Narodzenie", (26, 12): "Boże Narodzenie (drugi dzień)"}
    key = (d.day, d.month)
    if key in h: return h[key]
    e = _easter(d.year)
    if d == e: return "Wielkanoc"
    if d == e + timedelta(days=1): return "Poniedziałek Wielkanocny"
    if d == e + timedelta(days=49): return "Zielone Świątki"
    if d == e + timedelta(days=60): return "Boże Ciało"
    return ""


STATUS_OPTIONS = [
    ("", "—"),
    ("obecny", "Obecny"),
    ("home_office", "Home Office"),
    ("urlop", "Urlop"),
    ("l4", "L4"),
    ("wolne_swieto", "Wolne za święto"),
    ("delegacja", "Delegacja"),
    ("nieobecny", "Nieobecny"),
    ("inne", "Inne"),
]

# More saturated backgrounds for better contrast
WEEKEND_COLOR = QColor(180, 200, 235)
HOLIDAY_COLOR = QColor(255, 195, 160)


# ─────────────────────────────────────────────
# Signature canvas
# ─────────────────────────────────────────────

class SignatureCanvas(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(400, 120)
        self.setMaximumHeight(160)
        self.setStyleSheet("background-color: white; border: 1px solid #aaa; border-radius: 4px;")
        self._path = QPainterPath()
        self._stroked = False

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.fillRect(self.rect(), Qt.GlobalColor.white)
        pen = QPen(QColor(180, 180, 180), 1, Qt.PenStyle.DashLine)
        p.setPen(pen)
        y = self.height() - 25
        p.drawLine(10, y, self.width() - 10, y)
        p.setPen(QColor(120, 120, 120))
        p.setFont(QFont("Arial", 9))
        p.drawText(12, y - 4, "Podpis:")
        if not self._path.isEmpty():
            pen2 = QPen(QColor(0, 0, 140), 2, Qt.PenStyle.SolidLine,
                        Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
            p.setPen(pen2)
            p.drawPath(self._path)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._path.moveTo(event.position())
            self._stroked = False
            self.update()

    def mouseMoveEvent(self, event):
        if event.buttons() & Qt.MouseButton.LeftButton:
            self._path.lineTo(event.position())
            self._stroked = True
            self.update()

    def clear(self):
        self._path = QPainterPath()
        self._stroked = False
        self.update()

    def has_sig(self) -> bool:
        return self._stroked and not self._path.isEmpty()

    def save_png(self, path: str) -> bool:
        """Grab widget pixels and crop to signature area — guaranteed WYSIWYG."""
        if not self.has_sig():
            return False
        pixmap = self.grab()
        if pixmap.isNull():
            return False
        # Crop above guide line (signature drawing area)
        # Guide line is at height - 25; crop 5px above "Podpis:" text
        crop_h = self.height() - 35
        if crop_h < 10:
            crop_h = self.height() - 10
        cropped = pixmap.copy(0, 0, pixmap.width(), crop_h)
        if cropped.isNull():
            return False
        # Scale up for document quality
        scaled = cropped.scaled(
            cropped.width() * 3, cropped.height() * 3,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation)
        return scaled.save(path, "PNG")


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
        self.setMinimumHeight(30)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 1, 0, 1)
        layout.setSpacing(6)

        self.date_label = QLabel(day_date.strftime("%d-%m-%Y"))
        self.date_label.setMinimumWidth(85)
        self.date_label.setFont(QFont("Arial", 9, QFont.Weight.Bold))
        layout.addWidget(self.date_label)

        self.status_combo = QComboBox()
        for val, label in STATUS_OPTIONS:
            self.status_combo.addItem(label, val)
        self.status_combo.setMinimumWidth(150)
        self.status_combo.currentIndexChanged.connect(self._on_status_changed)
        layout.addWidget(self.status_combo)

        self.uwaga_edit = QLineEdit()
        self.uwaga_edit.setPlaceholderText("Uwaga")
        self.uwaga_edit.setMinimumWidth(200)
        self.uwaga_edit.setEnabled(False)
        layout.addWidget(self.uwaga_edit, stretch=1)

        if is_holiday:
            self._apply_holiday()
        elif self._is_weekend:
            self._style_weekend()

        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

    def _on_status_changed(self, idx):
        val = self.status_combo.currentData()
        if val == "delegacja":
            self.uwaga_edit.setEnabled(True)
            self.uwaga_edit.setPlaceholderText("Miejsce delegacji")
        elif val == "inne":
            self.uwaga_edit.setEnabled(True)
            self.uwaga_edit.setPlaceholderText("Opis")
        else:
            self.uwaga_edit.setEnabled(False)
            if not (val == "obecny" and self.uwaga_edit.text() == "Tychy"):
                self.uwaga_edit.setText("")

    def _apply_holiday(self):
        for i in range(self.status_combo.count()):
            if self.status_combo.itemData(i) == "wolne_swieto":
                self.status_combo.setCurrentIndex(i)
                break
        if self._holiday_name:
            self.uwaga_edit.setText(self._holiday_name)
        self.setStyleSheet(f"background-color: {HOLIDAY_COLOR.name()}; border-radius: 2px;")

    def _style_weekend(self):
        self.setStyleSheet(f"background-color: {WEEKEND_COLOR.name()}; border-radius: 2px;")

    def set_present(self):
        for i in range(self.status_combo.count()):
            if self.status_combo.itemData(i) == "obecny":
                self.status_combo.setCurrentIndex(i)
                break
        self.uwaga_edit.setEnabled(False)
        self.uwaga_edit.setText("")
        self.setStyleSheet("")

    def is_workday(self) -> bool:
        return not self._is_weekend and not self._is_holiday

    def get_data(self) -> dict:
        return {
            "date": self.day_date,
            "status": self.status_combo.currentData(),
            "status_label": self.status_combo.currentText(),
            "uwaga": self.uwaga_edit.text().strip(),
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
        self.setMinimumSize(700, 700)
        self._rows = []
        self._setup_ui()

    def _setup_ui(self):
        c = QWidget()
        self.setCentralWidget(c)
        ml = QVBoxLayout(c)
        ml.setSpacing(8)

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
        self.refresh_btn.clicked.connect(self._rebuild)
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
        ml.addLayout(top)

        # Separator
        s = QFrame()
        s.setFrameShape(QFrame.Shape.HLine)
        s.setFrameShadow(QFrame.Shadow.Sunken)
        ml.addWidget(s)

        # Header row — widths match DayRow widgets exactly
        hdr = QHBoxLayout()
        hdr.setSpacing(6)
        hdr.setContentsMargins(0, 0, 0, 0)
        for t, w in [("Data", 85), ("Status", 150), ("Uwaga", 200)]:
            lbl = QLabel(t)
            lbl.setFont(QFont("Arial", 9, QFont.Weight.Bold))
            lbl.setMinimumWidth(w)
            hdr.addWidget(lbl)
        hdr.addStretch()
        ml.addLayout(hdr)

        # Scroll area
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.tw = QWidget()
        self.tl = QVBoxLayout(self.tw)
        self.tl.setSpacing(1)
        self.tl.setContentsMargins(0, 0, 0, 0)
        self.tl.addStretch()
        self.scroll.setWidget(self.tw)
        ml.addWidget(self.scroll, stretch=1)

        # Signature
        sl = QHBoxLayout()
        sl.addWidget(QLabel("Podpis:"))
        self.sig = SignatureCanvas()
        sl.addWidget(self.sig, stretch=1)
        cb = QPushButton("Wyczyść")
        cb.clicked.connect(self.sig.clear)
        sl.addWidget(cb)
        ml.addLayout(sl)

        # Buttons
        bl = QHBoxLayout()
        af = QPushButton("Auto-fill workdays")
        af.setMinimumHeight(36)
        af.setStyleSheet("font-size: 13px; font-weight: bold; padding: 6px 16px;")
        af.clicked.connect(self._auto_fill)
        bl.addWidget(af)
        bl.addStretch()
        ex = QPushButton("Zapisz DOCX")
        ex.setMinimumHeight(36)
        ex.setStyleSheet("font-size: 14px; font-weight: bold; padding: 6px 20px;")
        ex.clicked.connect(self._export)
        bl.addWidget(ex)
        ml.addLayout(bl)

        self._rebuild()

    def _rebuild(self):
        month = self.month_spin.value()
        year = self.year_spin.value()
        for r in self._rows:
            self.tl.removeWidget(r)
            r.deleteLater()
        self._rows.clear()
        hols = polish_holidays(year)
        for dn in range(1, monthrange(year, month)[1] + 1):
            d = date(year, month, dn)
            is_hol = d in hols
            r = DayRow(d, is_hol, holiday_name(d) if is_hol else "")
            self._rows.append(r)
            self.tl.insertWidget(self.tl.count() - 1, r)

    def _auto_fill(self):
        n = 0
        for r in self._rows:
            if r.is_workday():
                r.set_present()
                n += 1
        QMessageBox.information(self, "Auto-fill", f"Wypełniono {n} dni: Obecny")

    def _collect(self) -> list:
        return [r.get_data() for r in self._rows]

    def _export(self):
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

        data = self._collect()
        sig_path = tempfile.NamedTemporaryFile(suffix='.png', delete=False).name
        has_sig = self.sig.save_png(sig_path)

        try:
            self._build_docx(fp, data, name, dept, month, year,
                             sig_path if has_sig else None)
            QMessageBox.information(self, "Sukces", f"DOKUMENT ZAPISANY:\n{fp}")
        except Exception as e:
            QMessageBox.critical(self, "Błąd", f"Nie udało się zapisać DOCX:\n{e}")

    def _build_docx(self, fp, data, name, dept, month, year, sig_path=None):
        doc = Document()

        # Tight margins for single-page fit
        for section in doc.sections:
            section.top_margin = Cm(0.7)
            section.bottom_margin = Cm(0.5)
            section.left_margin = Cm(1.2)
            section.right_margin = Cm(1.0)

        doc.styles['Normal'].font.name = 'Calibri'
        doc.styles['Normal'].font.size = Pt(10)

        # Title — compact
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = Pt(12)
        r = p.add_run(f"LISTA OBECNOŚCI - {month:02d}-{year}")
        r.bold = True
        r.font.size = Pt(12)
        r.font.name = 'Calibri'

        # Employee info
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = Pt(12)
        r = p.add_run(name + (f" - {dept}" if dept else ""))
        r.font.size = Pt(10)
        r.font.name = 'Calibri'

        # 3 columns: Data | Wejście | Wyjście
        headers = ["Data", "Wejście", "Wyjście"]
        nrows = len(data) + 1
        table = doc.add_table(rows=nrows, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = 'Table Grid'

        # Reduce cell padding
        tbl_pr = table._tbl.tblPr
        tbl_pr.append(parse_xml(
            f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))

        for ci, h in enumerate(headers):
            cell = table.rows[0].cells[ci]
            cell.text = ""
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].paragraph_format.space_after = Pt(0)
            cell.paragraphs[0].paragraph_format.space_before = Pt(0)
            # Set cell margins to 0
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_mar = parse_xml(
                f'<w:tcMar {nsdecls("w")}>'
                f'<w:top w:w="0" w:type="dxa"/>'
                f'<w:left w:w="0" w:type="dxa"/>'
                f'<w:bottom w:w="0" w:type="dxa"/>'
                f'<w:right w:w="0" w:type="dxa"/>'
                f'</w:tcMar>')
            tc_pr.append(tc_mar)
            r = cell.paragraphs[0].add_run(h)
            r.bold = True; r.font.size = Pt(10); r.font.name = 'Calibri'
            self._shade(cell, "D9D9D9")

        # Use atLeast row height so rows match content

        for ri, rd in enumerate(data):
            d = rd["date"]
            st = rd["status"]

            # Determine cell content
            if st in ("obecny", "home_office", "delegacja"):
                if st == "obecny":
                    wej = ""
                    wyj = ""
                elif st == "home_office":
                    wej = "Home Office"
                    wyj = "Home Office"
                else:
                    loc = rd.get("uwaga", "")
                    txt = f"Delegacja - {loc}" if loc else "Delegacja"
                    wej = txt; wyj = txt
            elif st in ("urlop", "l4", "wolne_swieto", "nieobecny", "inne", ""):
                if rd["is_holiday"] and st == "wolne_swieto":
                    hn = rd.get("holiday_name", "")
                    wej = f"Wolne: {hn}" if hn else "Wolne za święto"
                elif rd["is_weekend"] and not st:
                    wej = "dzień wolny od pracy"
                elif st == "inne":
                    uw = rd.get("uwaga", "")
                    wej = f"Inne - {uw}" if uw else "Inne"
                else:
                    wej = rd["status_label"] if rd["status_label"] else "—"
                wyj = wej
            else:
                wej = rd["status_label"]; wyj = wej

            show_sig = sig_path and os.path.exists(sig_path) and st in ("obecny", "home_office", "delegacja")

            doc_row = table.rows[ri + 1]
            cells_data = [d.strftime("%d-%m-%Y"), wej, wyj]

            for ci in range(3):
                cell = doc_row.cells[ci]
                cell.text = ""
                # Set cell margins to 0
                tc_pr = cell._tc.get_or_add_tcPr()
                tc_mar = parse_xml(
                    f'<w:tcMar {nsdecls("w")}>'
                    f'<w:top w:w="0" w:type="dxa"/>'
                    f'<w:left w:w="0" w:type="dxa"/>'
                    f'<w:bottom w:w="0" w:type="dxa"/>'
                    f'<w:right w:w="0" w:type="dxa"/>'
                    f'</w:tcMar>')
                tc_pr.append(tc_mar)
                par = cell.paragraphs[0]
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                par.paragraph_format.space_after = Pt(0)
                par.paragraph_format.space_before = Pt(0)
                par.paragraph_format.line_spacing = Pt(12)

                if ci >= 1 and show_sig:
                    # Signature in Wejście/Wyjście
                    if st == "home_office":
                        r = par.add_run("Home Office\n")
                        r.font.size = Pt(8); r.font.name = 'Calibri'
                    elif st == "delegacja":
                        loc = rd.get("uwaga", "")
                        label = f"Delegacja - {loc}\n" if loc else "Delegacja\n"
                        r = par.add_run(label)
                        r.font.size = Pt(8); r.font.name = 'Calibri'
                    r = par.add_run()
                    r.add_picture(sig_path, width=Cm(2.0), height=Cm(0.5))
                else:
                    txt = str(cells_data[ci])
                    if txt:
                        r = par.add_run(txt)
                        r.font.size = Pt(10); r.font.name = 'Calibri'

                # Background colors
                if rd["is_holiday"]:
                    self._shade(cell, "FCE4D6")
                elif rd["is_weekend"] and not rd["status"]:
                    self._shade(cell, "B0C4DE")

        doc.save(fp)

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