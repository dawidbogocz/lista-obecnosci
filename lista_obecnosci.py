#!/usr/bin/env python3
"""
Attendance Sheet App (Lista Obecnosci)
PySide6 GUI — HTML export with embedded signature, prints to one A4 page.
"""

import sys
import os
import base64
import io
from datetime import date, timedelta
from calendar import monthrange

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox,
    QScrollArea, QMessageBox, QFileDialog, QSpinBox,
    QFrame, QSizePolicy
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPainter, QPen, QColor, QFont, QImage, QPainterPath, QPixmap


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
    ("", "—"),
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
        """Return signature as a base64 data URL — no temp files, no cropping issues."""
        if not self.has_sig():
            return None
        # Grab what's on screen — guaranteed WYSIWYG
        px = self.grab()
        if px.isNull():
            return None
        # Crop above guide line
        ch = self.height() - 35
        if ch < 10: ch = self.height() - 10
        c = px.copy(0, 0, px.width(), ch)
        if c.isNull():
            return None
        # Scale up 3x
        s = c.scaled(c.width() * 3, c.height() * 3,
                     Qt.AspectRatioMode.KeepAspectRatio,
                     Qt.TransformationMode.SmoothTransformation)
        buf = io.BytesIO()
        s.save(buf, "PNG")
        b64 = base64.b64encode(buf.getvalue()).decode()
        return f"data:image/png;base64,{b64}"


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
        ex = QPushButton("Zapisz HTML")
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
        QMessageBox.information(self, "Auto-fill", f"Wypelniono {n} dni: Obecny")

    def _collect(self):
        return [r.get_data() for r in self._rows]

    def _cell_text(self, rd):
        """Return (wejscie_text, wyjscie_text, show_sig, sig_label) for a row."""
        st = rd["status"]
        sl = rd["status_label"]
        if st in ("obecny", "home_office", "delegacja"):
            wej = ""; wyj = ""
            label = ""
            if st == "home_office": label = "Home Office"
            elif st == "delegacja":
                loc = rd.get("uwaga", "")
                label = f"Delegacja - {loc}" if loc else "Delegacja"
            return (wej, wyj, True, label)
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
                t = sl if sl else "—"
            return (t, t, False, "")

    def _export(self):
        month = self.month_spin.value()
        year = self.year_spin.value()
        name = self.name_edit.text().strip() or "Pracownik"
        dept = self.dept_edit.text().strip() or ""

        fp, _ = QFileDialog.getSaveFileName(
            self, "Zapisz HTML",
            os.path.expanduser(f"~/lista_obecnosci_{month:02d}-{year}.html"),
            "HTML (*.html *.htm)")
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

        rows_html = ""
        row_n = 1
        for rd in data:
            bg = ""
            if rd["is_holiday"]:
                bg = ' style="background:#FCE4D6"'
            elif rd["is_weekend"] and not rd["status"]:
                bg = ' style="background:#B0C4DE"'

            wej, wyj, show_sig, sig_label = self._cell_text(rd)
            date_str = rd["date"].strftime("%d-%m-%Y")

            # Determine cell content for Wejscie and Wyjscie
            def make_cell(text, has_sig, label):
                parts = []
                if label:
                    parts.append(f'<div style="font-size:10pt">{label}</div>')
                if has_sig and sig_url:
                    parts.append(f'<img src="{sig_url}" style="height:0.7cm; display:block; margin:0 auto">')
                elif not label and text:
                    parts.append(f'<span style="font-size:10pt">{text}</span>')
                return "".join(parts) if parts else "&nbsp;"

            wj = make_cell(wej, show_sig, sig_label)
            wy = make_cell(wyj, show_sig, sig_label)

            rows_html += f"""        <tr{bg}>
            <td style="text-align:center;font-size:10pt;padding:2px 4px">{date_str}</td>
            <td style="text-align:center;font-size:10pt;padding:2px 4px">{wj}</td>
            <td style="text-align:center;font-size:10pt;padding:2px 4px">{wy}</td>
        </tr>
"""
            row_n += 1

        info = name + (f" - {dept}" if dept else "")

        return f"""<!DOCTYPE html>
<html lang="pl">
<head>
<meta charset="utf-8">
<style>
    @page {{ size: A4; margin: 0.7cm 1cm 0.5cm 1cm; }}
    @media print {{ body {{ margin: 0; padding: 0; }} }}
    body {{ font-family: Calibri, Arial, sans-serif; font-size: 10pt; margin: 0.5cm; }}
    h1 {{ text-align: center; font-size: 14pt; margin: 0 0 2px 0; }}
    .info {{ text-align: center; font-size: 10pt; margin: 0 0 4px 0; }}
    table {{ width: 100%; border-collapse: collapse; }}
    th, td {{ border: 1px solid #888; padding: 2px 4px; }}
    th {{ background: #D9D9D9; font-size: 10pt; text-align: center; }}
    td {{ vertical-align: middle; }}
    tr:nth-child(even) td {{ background: inherit; }}
</style>
</head>
<body>
<h1>LISTA OBECNOSCI - {month:02d}-{year}</h1>
<div class="info">{info}</div>
<table>
    <tr>
        <th style="width:18%">Data</th>
        <th style="width:41%">Wejscie</th>
        <th style="width:41%">Wyjscie</th>
    </tr>
{rows_html}</table>
</body>
</html>"""


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