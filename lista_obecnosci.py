#!/usr/bin/env python3
"""
Attendance Sheet App (Lista Obecnosci)
PySide6 GUI — PDF export (QPainter, perfect one-page), HTML fallback, DOCX fallback.
"""

import sys
import os
import base64
from datetime import date, timedelta
from calendar import monthrange

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox,
    QScrollArea, QMessageBox, QFileDialog, QSpinBox,
    QFrame, QSizePolicy
)
from PySide6.QtCore import Qt, QBuffer, QIODevice, QRectF, QMarginsF
from PySide6.QtGui import (
    QPainter, QPen, QColor, QFont, QImage, QPainterPath, QPixmap,
    QPageSize, QPageLayout, QTextDocument
)
from PySide6.QtPrintSupport import QPrinter

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml


# ─────────────────────────────────────────────
# Polish holidays
# ─────────────────────────────────────────────

def _easter(year):
    a = year % 19; b = year // 100; c = year % 100
    d = b // 4; e = b % 4; f = (b + 8) // 25
    g = (b - f + 1) // 3; h = (19 * a + b - d - g + 15) % 30
    i = c // 4; k = c % 4; l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def polish_holidays(year):
    e = _easter(year)
    fixed = {date(year, 1, 1), date(year, 1, 6), date(year, 5, 1),
             date(year, 5, 3), date(year, 8, 15), date(year, 11, 1),
             date(year, 11, 11), date(year, 12, 25), date(year, 12, 26)}
    return fixed | {e, e + timedelta(days=1), e + timedelta(days=49), e + timedelta(days=60)}


def holiday_name(d):
    h = {(1, 1): "Nowy Rok", (6, 1): "Swieto Trzech Kroli", (1, 5): "Swieto Pracy",
         (3, 5): "Swieto Konstytucji 3 Maja", (15, 8): "Wniebowziecie NMP",
         (1, 11): "Wszystkich Swietych", (11, 11): "Narodowe Swieto Niepodleglosci",
         (25, 12): "Boze Narodzenie", (26, 12): "Boze Narodzenie (drugi dzien)"}
    key = (d.day, d.month)
    if key in h: return h[key]
    e = _easter(d.year)
    if d == e: return "Wielkanoc"
    if d == e + timedelta(days=1): return "Poniedzialek Wielkanocny"
    if d == e + timedelta(days=49): return "Zielone Swiatki"
    if d == e + timedelta(days=60): return "Boze Cialo"
    return ""


STATUS_OPTIONS = [
    ("", "-"),
    ("obecny", "Obecny"),
    ("home_office", "Home Office"),
    ("urlop", "Urlop"),
    ("l4", "L4"),
    ("wolne_swieto", "Wolne za swieto"),
    ("delegacja", "Delegacja"),
    ("nieobecny", "Nieobecny"),
    ("inne", "Inne"),
]

WEEKEND_COLOR = QColor(180, 200, 235)
HOLIDAY_COLOR = QColor(255, 195, 160)
HEADER_COLOR = QColor(217, 217, 217)


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
        p.setPen(QPen(QColor(180, 180, 180), 1, Qt.PenStyle.DashLine))
        y = self.height() - 25
        p.drawLine(10, y, self.width() - 10, y)
        p.setPen(QColor(120, 120, 120))
        p.setFont(QFont("Arial", 9))
        p.drawText(12, y - 4, "Podpis:")
        if not self._path.isEmpty():
            p.setPen(QPen(QColor(0, 0, 140), 2, Qt.PenStyle.SolidLine,
                         Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
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

    def has_sig(self):
        return self._stroked and not self._path.isEmpty()

    def to_data_url(self):
        """Render path on clean white canvas, return data URL."""
        if not self.has_sig():
            return None
        scale = 3
        w = self.width() * scale
        h = self.height() * scale
        img = QImage(w, h, QImage.Format.Format_RGB32)
        img.fill(Qt.GlobalColor.white)
        p = QPainter(img)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.scale(scale, scale)
        p.setPen(QPen(QColor(0, 0, 140), 2, Qt.PenStyle.SolidLine,
                     Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
        p.drawPath(self._path)
        p.end()
        buf = QBuffer()
        buf.open(QIODevice.OpenModeFlag.WriteOnly)
        img.save(buf, "PNG")
        buf.close()
        return f"data:image/png;base64,{base64.b64encode(buf.data()).decode()}"


# ─────────────────────────────────────────────
# Day row
# ─────────────────────────────────────────────

class DayRow(QFrame):
    def __init__(self, day_date, is_holiday, holiday_name_str="", parent=None):
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

    def is_workday(self):
        return not self._is_weekend and not self._is_holiday

    def get_data(self):
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
        self.setWindowTitle("Lista Obecnosci")
        self.setMinimumSize(700, 700)
        self._rows = []
        self._setup_ui()

    def _setup_ui(self):
        c = QWidget()
        self.setCentralWidget(c)
        ml = QVBoxLayout(c)
        ml.setSpacing(8)

        top = QHBoxLayout()
        self.month_spin = QSpinBox()
        self.month_spin.setRange(1, 12)
        self.month_spin.setValue(date.today().month)
        self.month_spin.setPrefix("Miesiac: ")
        self.month_spin.setMinimumWidth(140)
        self.year_spin = QSpinBox()
        self.year_spin.setRange(2020, 2100)
        self.year_spin.setValue(date.today().year)
        self.year_spin.setPrefix("Rok: ")
        self.year_spin.setMinimumWidth(120)
        self.refresh_btn = QPushButton("Odswiez")
        self.refresh_btn.clicked.connect(self._rebuild)
        top.addWidget(self.month_spin)
        top.addWidget(self.year_spin)
        top.addWidget(self.refresh_btn)
        top.addStretch()
        top.addWidget(QLabel("Imie i nazwisko:"))
        self.name_edit = QLineEdit("Dawid Bogocz")
        self.name_edit.setMinimumWidth(140)
        top.addWidget(self.name_edit)
        top.addWidget(QLabel("Dzial:"))
        self.dept_edit = QLineEdit("Dzial IT")
        self.dept_edit.setMinimumWidth(120)
        top.addWidget(self.dept_edit)
        ml.addLayout(top)

        s = QFrame()
        s.setFrameShape(QFrame.Shape.HLine)
        s.setFrameShadow(QFrame.Shadow.Sunken)
        ml.addWidget(s)

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

        sl = QHBoxLayout()
        sl.addWidget(QLabel("Podpis:"))
        self.sig = SignatureCanvas()
        sl.addWidget(self.sig, stretch=1)
        cb = QPushButton("Wyczysc")
        cb.clicked.connect(self.sig.clear)
        sl.addWidget(cb)
        ml.addLayout(sl)

        bl = QHBoxLayout()
        af = QPushButton("Auto-fill workdays")
        af.setMinimumHeight(36)
        af.setStyleSheet("font-size: 13px; font-weight: bold; padding: 6px 16px;")
        af.clicked.connect(self._auto_fill)
        bl.addWidget(af)
        bl.addStretch()
        for txt, handler in [("Zapisz PDF", self._export_pdf),
                              ("Zapisz HTML", self._export_html),
                              ("Zapisz DOCX", self._export_docx)]:
            btn = QPushButton(txt)
            btn.setMinimumHeight(36)
            btn.setStyleSheet("font-size: 13px; font-weight: bold; padding: 6px 14px;")
            btn.clicked.connect(handler)
            bl.addWidget(btn)
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
        QMessageBox.information(self, "Auto-fill",
                                f"Wypelniono {n} dni: Obecny")

    def _collect(self):
        return [r.get_data() for r in self._rows]

    def _cell_info(self, rd):
        """Return (wej, wyj, show_sig, label) for a row's cells."""
        st = rd["status"]
        sl = rd["status_label"]
        if st in ("obecny", "home_office", "delegacja"):
            label = ""
            if st == "home_office": label = "Home Office"
            elif st == "delegacja":
                loc = rd.get("uwaga", "")
                label = f"Delegacja - {loc}" if loc else "Delegacja"
            return ("", "", True, label)
        else:
            if rd["is_holiday"] and st == "wolne_swieto":
                hn = rd.get("holiday_name", "")
                t = f"Wolne: {hn}" if hn else "Wolne za swieto"
            elif rd["is_weekend"] and not st:
                t = "dzien wolny od pracy"
            elif st == "inne":
                uw = rd.get("uwaga", "")
                t = f"Inne - {uw}" if uw else "Inne"
            else:
                t = sl if sl else "-"
            return (t, t, False, "")

    # ─── PDF export (QPainter — exact, one page) ───

    def _export_pdf(self):
        month = self.month_spin.value()
        year = self.year_spin.value()
        name = self.name_edit.text().strip() or "Pracownik"
        dept = self.dept_edit.text().strip() or ""

        fp, _ = QFileDialog.getSaveFileName(
            self, "Zapisz PDF",
            os.path.expanduser(f"~/lista_obecnosci_{month:02d}-{year}.pdf"),
            "PDF (*.pdf)")
        if not fp:
            return

        data = self._collect()
        sig_data = self.sig.to_data_url()

        try:
            self._render_pdf(fp, data, name, dept, month, year, sig_data)
            QMessageBox.information(self, "Sukces", f"PDF ZAPISANY:\n{fp}")
        except Exception as e:
            QMessageBox.critical(self, "Blad", f"Nie udalo sie zapisac PDF:\n{e}")

    def _render_pdf(self, fp, data, name, dept, month, year, sig_data):
        """Render table to PDF using QPrinter + QPainter for exact one-page output."""
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setPageSize(QPageSize(QPageSize.PageSizeID.A4))
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        printer.setOutputFileName(fp)
        printer.setPageMargins(QMarginsF(12, 8, 12, 6), QPageLayout.Unit.Millimeter)

        # Build HTML table so we can use QTextDocument for layout
        mn = ["", "Styczen", "Luty", "Marzec", "Kwiecien", "Maj",
              "Czerwiec", "Lipiec", "Sierpien", "Wrzesien",
              "Pazdziernik", "Listopad", "Grudzien"]
        info = name + (f" - {dept}" if dept else "")

        rows = ""
        for rd in data:
            bg = ""
            if rd["is_holiday"]:
                bg = ' bgcolor="#FCE4D6"'
            elif rd["is_weekend"] and not rd["status"]:
                bg = ' bgcolor="#B0C4DE"'

            wej, wyj, show_sig, label = self._cell_info(rd)
            date_str = rd["date"].strftime("%d-%m-%Y")

            def cell_html(text, has_sig, lbl):
                parts = []
                if lbl:
                    parts.append(f'<div style="font-size:9pt;margin-bottom:1px">{lbl}</div>')
                if has_sig and sig_data:
                    parts.append(f'<img src="{sig_data}" style="height:1.0cm;display:block;margin:1px auto">')
                elif not lbl and text:
                    parts.append(f'<span style="font-size:9pt">{text}</span>')
                return "".join(parts) if parts else "&nbsp;"

            wj = cell_html(wej, show_sig, label)
            wy = cell_html(wyj, show_sig, label)
            rows += f"""<tr{bg}><td align=center style="font-size:9pt">{date_str}</td>
<td align=center style="font-size:9pt">{wj}</td>
<td align=center style="font-size:9pt">{wy}</td></tr>\n"""

        html = f"""<!DOCTYPE html><html><head><meta charset=utf-8>
<style>
body {{ font-family:Calibri,Arial,sans-serif; font-size:9pt; margin:0; padding:0; }}
h1 {{ text-align:center; font-size:13pt; margin:0 0 1px 0; }}
.info {{ text-align:center; font-size:9pt; margin:0 0 3px 0; }}
table {{ width:100%; border-collapse:collapse; }}
th, td {{ border:1px solid #888; padding:1px 3px; }}
th {{ background:#D9D9D9; font-size:9pt; text-align:center; }}
</style></head><body>
<h1>LISTA OBECNOSCI - {month:02d}-{year}</h1>
<div class="info">{info}</div>
<table><tr><th style=width:18%>Data</th><th style=width:41%>Wejscie</th><th style=width:41%>Wyjscie</th></tr>
{rows}</table></body></html>"""

        doc = QTextDocument()
        doc.setHtml(html)
        # Set page size to A4 in points (1 point = 1/72 inch)
        doc.setPageSize(QPageSize(QPageSize.PageSizeID.A4).size(QPageSize.SizeUnit.Point))
        doc.print_(printer)

    # ─── HTML export ───

    def _export_html(self):
        month = self.month_spin.value()
        year = self.year_spin.value()
        name = self.name_edit.text().strip() or "Pracownik"
        dept = self.dept_edit.text().strip() or ""

        fp, _ = QFileDialog.getSaveFileName(
            self, "Zapisz HTML",
            os.path.expanduser(f"~/lista_obecnosci_{month:02d}-{year}.html"),
            "HTML (*.html)")
        if not fp:
            return

        data = self._collect()
        sig_url = self.sig.to_data_url()

        try:
            html = self._build_html(data, name, dept, month, year, sig_url)
            with open(fp, "w", encoding="utf-8") as f:
                f.write(html)
            QMessageBox.information(self, "Sukces", f"HTML ZAPISANY:\n{fp}")
        except Exception as e:
            QMessageBox.critical(self, "Blad", f"Nie udalo sie zapisac HTML:\n{e}")

    def _build_html(self, data, name, dept, month, year, sig_url):
        mn = ["", "Styczen", "Luty", "Marzec", "Kwiecien", "Maj",
              "Czerwiec", "Lipiec", "Sierpien", "Wrzesien",
              "Pazdziernik", "Listopad", "Grudzien"]
        info = name + (f" - {dept}" if dept else "")

        rows = ""
        for rd in data:
            bg = ""
            if rd["is_holiday"]: bg = ' style="background:#FCE4D6"'
            elif rd["is_weekend"] and not rd["status"]: bg = ' style="background:#B0C4DE"'

            wej, wyj, show_sig, label = self._cell_info(rd)
            date_str = rd["date"].strftime("%d-%m-%Y")

            def cell(text, has_sig, lbl):
                parts = []
                if lbl: parts.append(f'<div style="font-size:10pt">{lbl}</div>')
                if has_sig and sig_url:
                    parts.append(f'<img src="{sig_url}" style="height:1.1cm;display:block;margin:2px auto">')
                elif not lbl and text:
                    parts.append(f'<span style="font-size:10pt">{text}</span>')
                return "".join(parts) if parts else "&nbsp;"

            wj = cell(wej, show_sig, label)
            wy = cell(wyj, show_sig, label)
            rows += f"<tr{bg}><td align=center style=font-size:10pt>{date_str}</td><td align=center style=font-size:10pt>{wj}</td><td align=center style=font-size:10pt>{wy}</td></tr>\n"

        return f"""<!DOCTYPE html><html lang=pl><head><meta charset=utf-8><style>
@page {{ size:A4; margin:0; }}
body {{ font-family:Calibri,Arial,sans-serif; font-size:10pt; }}
h1 {{ text-align:center; font-size:14pt; margin:0 0 2px 0; }}
.info {{ text-align:center; font-size:10pt; margin:0 0 4px 0; }}
table {{ width:100%; border-collapse:collapse; }}
th,td {{ border:1px solid #888; padding:2px 4px; }}
th {{ background:#D9D9D9; font-size:10pt; text-align:center; }}
</style></head><body>
<h1>LISTA OBECNOSCI - {month:02d}-{year}</h1>
<div class="info">{info}</div>
<table><tr><th style=width:18%>Data</th><th style=width:41%>Wejscie</th><th style=width:41%>Wyjscie</th></tr>
{rows}</table></body></html>"""

    # ─── DOCX export (no signature) ───

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

        data = self._collect()

        try:
            self._build_docx(fp, data, name, dept, month, year)
            QMessageBox.information(self, "Sukces", f"DOCX ZAPISANY:\n{fp}")
        except Exception as e:
            QMessageBox.critical(self, "Blad", f"Nie udalo sie zapisac DOCX:\n{e}")

    def _build_docx(self, fp, data, name, dept, month, year):
        doc = Document()
        for section in doc.sections:
            section.top_margin = Cm(0.7)
            section.bottom_margin = Cm(0.5)
            section.left_margin = Cm(1.2)
            section.right_margin = Cm(1.0)

        doc.styles['Normal'].font.name = 'Calibri'
        doc.styles['Normal'].font.size = Pt(10)

        mn = ["", "Styczen", "Luty", "Marzec", "Kwiecien", "Maj",
              "Czerwiec", "Lipiec", "Sierpien", "Wrzesien",
              "Pazdziernik", "Listopad", "Grudzien"]

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(f"LISTA OBECNOSCI - {month:02d}-{year}")
        r.bold = True; r.font.size = Pt(12); r.font.name = 'Calibri'

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(name + (f" - {dept}" if dept else ""))
        r.font.size = Pt(10); r.font.name = 'Calibri'

        # 3 cols: Data, Wejscie, Wyjscie
        table = doc.add_table(rows=len(data) + 1, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = 'Table Grid'

        for ci, h in enumerate(["Data", "Wejscie", "Wyjscie"]):
            cell = table.rows[0].cells[ci]
            cell.text = ""
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = cell.paragraphs[0].add_run(h)
            r.bold = True; r.font.size = Pt(9); r.font.name = 'Calibri'
            self._shade_cell(cell, "D9D9D9")

        for ri, rd in enumerate(data):
            wej, wyj, show_sig, label = self._cell_info(rd)
            # DOCX has no signature — just show the label for sig rows
            if show_sig and label:
                cell_text = label
            elif show_sig and not label:
                cell_text = "Obecny"
            else:
                cell_text = wej
            doc_row = table.rows[ri + 1]
            vals = [rd["date"].strftime("%d-%m-%Y"), cell_text, cell_text if not show_sig else cell_text]
            for ci in range(3):
                cell = doc_row.cells[ci]
                cell.text = ""
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = cell.paragraphs[0].add_run(str(vals[ci]))
                r.font.size = Pt(9); r.font.name = 'Calibri'
                if rd["is_holiday"]:
                    self._shade_cell(cell, "FCE4D6")
                elif rd["is_weekend"] and not rd["status"]:
                    self._shade_cell(cell, "DAE8FC")

        doc.save(fp)

    def _shade_cell(self, cell, color):
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