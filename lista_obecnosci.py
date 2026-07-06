#!/home/dzeyd/.hermes/hermes-agent/venv/bin/python3
"""
Attendance Sheet App (Lista Obecności)
PySide6 GUI desktop app — calendar-based monthly attendance with signature canvas and PDF export.
"""

import sys
import os
import math
from datetime import date, timedelta
from calendar import monthrange, day_name, month_name

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QLineEdit, QPushButton, QComboBox,
    QTimeEdit, QScrollArea, QMessageBox, QFileDialog, QSpinBox,
    QFrame, QSizePolicy
)
from PySide6.QtCore import Qt, QDate, QSize, Signal, QRectF, QPointF
from PySide6.QtGui import (
    QPainter, QPen, QColor, QFont, QImage, QPixmap, QAction,
    QPainterPath, QPageSize, QTransform
)

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import reportlab.lib.fonts

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


# ─────────────────────────────────────────────
# Signature canvas widget
# ─────────────────────────────────────────────

class SignatureCanvas(QWidget):
    """A simple drawable canvas for a signature."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(400, 120)
        self.setMaximumHeight(160)
        self.setStyleSheet("background-color: white; border: 1px solid #aaa; border-radius: 4px;")
        self._path = QPainterPath()
        self._points = []

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        # background
        painter.fillRect(self.rect(), Qt.GlobalColor.white)
        # dashed guide line
        pen = QPen(QColor(180, 180, 180), 1, Qt.PenStyle.DashLine)
        painter.setPen(pen)
        y = self.height() - 25
        painter.drawLine(10, y, self.width() - 10, y)
        # label
        painter.setPen(QColor(120, 120, 120))
        font = QFont("Arial", 9)
        painter.setFont(font)
        painter.drawText(12, y - 4, "Podpis:")
        # signature path
        if not self._path.isEmpty():
            pen2 = QPen(QColor(0, 0, 140), 2, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
            painter.setPen(pen2)
            painter.drawPath(self._path)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._path = QPainterPath()
            self._points = [(event.position().x(), event.position().y())]
            self._path.moveTo(event.position())
            self.update()

    def mouseMoveEvent(self, event):
        if event.buttons() & Qt.MouseButton.LeftButton:
            self._points.append((event.position().x(), event.position().y()))
            self._path.lineTo(event.position())
            self.update()

    def clear_signature(self):
        self._path = QPainterPath()
        self._points = []
        self.update()

    def to_qimage(self) -> QImage:
        """Render the signature to a QImage with transparent background."""
        if self._path.isEmpty():
            return QImage()
        rect = self._path.boundingRect().adjusted(-5, -5, 5, 5).toRect()
        if rect.width() < 5 or rect.height() < 5:
            return QImage()
        img = QImage(rect.size(), QImage.Format.Format_ARGB32_Premultiplied)
        img.fill(Qt.GlobalColor.transparent)
        painter = QPainter(img)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.translate(-rect.topLeft())
        pen = QPen(QColor(0, 0, 140), 2, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        painter.drawPath(self._path)
        painter.end()
        return img

    def has_signature(self) -> bool:
        return not self._path.isEmpty()


# ─────────────────────────────────────────────
# Day row widget
# ─────────────────────────────────────────────

class DayRow(QFrame):
    """A single row for one day in the attendance sheet."""

    def __init__(self, day_date: date, is_weekend: bool, is_holiday: bool, holiday_name_str: str = "", parent=None):
        super().__init__(parent)
        self.day_date = day_date
        self._is_weekend = is_weekend
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

        # Entry time
        self.time_in = QTimeEdit()
        self.time_in.setDisplayFormat("HH:mm")
        self.time_in.setTime(self.time_in.time().fromString("07:30", "HH:mm"))
        self.time_in.setMinimumWidth(70)
        self.time_in.setEnabled(True)
        layout.addWidget(self.time_in)

        # Separator
        sep = QLabel("→")
        sep.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sep.setMaximumWidth(20)
        layout.addWidget(sep)

        # Exit time
        self.time_out = QTimeEdit()
        self.time_out.setDisplayFormat("HH:mm")
        self.time_out.setTime(self.time_out.time().fromString("15:30", "HH:mm"))
        self.time_out.setMinimumWidth(70)
        self.time_out.setEnabled(True)
        layout.addWidget(self.time_out)

        # Location
        self.location_edit = QLineEdit()
        self.location_edit.setPlaceholderText("Lokacja")
        self.location_edit.setMinimumWidth(120)
        layout.addWidget(self.location_edit)

        # Auto-fill for weekends/holidays
        if is_weekend:
            self._apply_weekend()
        elif is_holiday:
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

    def _apply_weekend(self):
        idx = self.status_combo.findData("")
        self.status_combo.setCurrentIndex(idx)
        self.status_combo.setEnabled(False)
        self.time_in.setEnabled(False)
        self.time_out.setEnabled(False)
        self.location_edit.setEnabled(False)
        self.date_label.setStyleSheet("color: #999;")
        self.setStyleSheet(f"background-color: {WEEKEND_COLOR.name()}; border-radius: 2px;")

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

    def get_data(self) -> dict:
        return {
            "date": self.day_date,
            "status": self.status_combo.currentData(),
            "status_label": self.status_combo.currentText(),
            "time_in": self.time_in.time().toString("HH:mm"),
            "time_out": self.time_out.time().toString("HH:mm"),
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

        # ── Signature ──
        sig_layout = QHBoxLayout()
        sig_layout.addWidget(QLabel("Podpis:"))
        self.sig_canvas = SignatureCanvas()
        sig_layout.addWidget(self.sig_canvas, stretch=1)
        clear_sig_btn = QPushButton("Wyczyść")
        clear_sig_btn.clicked.connect(self.sig_canvas.clear_signature)
        sig_layout.addWidget(clear_sig_btn)
        main_layout.addLayout(sig_layout)

        # ── Bottom buttons ──
        btn_layout = QHBoxLayout()
        export_pdf_btn = QPushButton("📄 Zapisz PDF")
        export_pdf_btn.setMinimumHeight(36)
        export_pdf_btn.setStyleSheet("font-size: 14px; font-weight: bold; padding: 6px 20px;")
        export_pdf_btn.clicked.connect(self._export_pdf)
        btn_layout.addStretch()
        btn_layout.addWidget(export_pdf_btn)
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
            is_weekend = d.weekday() >= 5
            is_hol = d in holidays
            h_name = holiday_name(d) if is_hol else ""

            row = DayRow(d, is_weekend, is_hol, h_name)
            self._day_rows.append(row)
            # Insert before the stretch
            self.table_layout.insertWidget(self.table_layout.count() - 1, row)

    def _collect_data(self) -> list:
        """Collect all day data into a list of dicts."""
        return [row.get_data() for row in self._day_rows]

    # ──────────────────────────────────────────
    # PDF export
    # ──────────────────────────────────────────

    def _export_pdf(self):
        month = self.month_spin.value()
        year = self.year_spin.value()
        name = self.name_edit.text().strip() or "Pracownik"
        dept = self.dept_edit.text().strip() or ""

        default_filename = f"lista_obecnosci_{month:02d}-{year}.pdf"
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Zapisz PDF", os.path.expanduser(f"~/{default_filename}"),
            "PDF (*.pdf)"
        )
        if not filepath:
            return

        data = self._collect_data()

        # Render signature to temp image
        sig_img_path = None
        if self.sig_canvas.has_signature():
            qimg = self.sig_canvas.to_qimage()
            if not qimg.isNull():
                sig_img_path = "/tmp/_attendance_sig_temp.png"
                qimg.save(sig_img_path)

        try:
            self._build_pdf(filepath, data, name, dept, month, year, sig_img_path)
            QMessageBox.information(self, "Sukces", f"PDF zapisany:\n{filepath}")
        except Exception as e:
            QMessageBox.critical(self, "Błąd", f"Nie udało się zapisać PDF:\n{e}")

    def _build_pdf(self, filepath, data, name, dept, month, year, sig_img_path=None):
        doc = SimpleDocTemplate(
            filepath, pagesize=A4,
            leftMargin=15 * mm, rightMargin=10 * mm,
            topMargin=15 * mm, bottomMargin=15 * mm,
        )

        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(
            name='CenterTitle', parent=styles['Normal'],
            fontSize=16, alignment=TA_CENTER, spaceAfter=4 * mm,
            fontName="Helvetica-Bold"
        ))
        styles.add(ParagraphStyle(
            name='InfoLine', parent=styles['Normal'],
            fontSize=10, alignment=TA_CENTER, spaceAfter=2 * mm,
        ))

        elements = []

        # Title
        month_names = [
            "", "Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
            "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"
        ]
        title = f"LISTA OBECNOŚCI - {month_names[month]} {year}"
        elements.append(Paragraph(title, styles['CenterTitle']))

        info_line = f"{name}"
        if dept:
            info_line += f" — {dept}"
        elements.append(Paragraph(info_line, styles['InfoLine']))
        elements.append(Spacer(1, 4 * mm))

        # Table
        table_data = [["Data", "Status", "Wejście", "Wyjście", "Lokacja / Uwagi"]]

        for row_data in data:
            d = row_data["date"]
            day_names_short = ["Pon", "Wt", "Śr", "Czw", "Pt", "Sob", "Niedz"]
            date_str = f"{d.day:02d} {day_names_short[d.weekday()]}"

            status = row_data["status_label"]
            if row_data["is_weekend"]:
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
            if row_data["is_weekend"]:
                location = "—"
            elif row_data["is_holiday"] and row_data["holiday_name"]:
                location = row_data["holiday_name"]

            table_data.append([date_str, status, time_in, time_out, location])

        # Calculate column widths
        page_width = A4[0] - 25 * mm
        col_widths = [page_width * 0.16, page_width * 0.30, page_width * 0.12, page_width * 0.12, page_width * 0.30]

        tbl = Table(table_data, colWidths=col_widths, repeatRows=1)

        # Style
        style_cmds = [
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (4, 1), (4, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.85, 0.85, 0.85)),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]

        # Color weekends and holidays in table
        for i, row_data in enumerate(data):
            row_idx = i + 1  # +1 for header
            if row_data["is_weekend"]:
                style_cmds.append(
                    ('BACKGROUND', (0, row_idx), (-1, row_idx),
                     colors.Color(0.95, 0.95, 0.95))
                )
            elif row_data["is_holiday"]:
                style_cmds.append(
                    ('BACKGROUND', (0, row_idx), (-1, row_idx),
                     colors.Color(0.99, 0.92, 0.88))
                )

        tbl.setStyle(TableStyle(style_cmds))
        elements.append(tbl)
        elements.append(Spacer(1, 6 * mm))

        # Signature
        if sig_img_path and os.path.exists(sig_img_path):
            sig_elements = []
            sig_elements.append(Paragraph("Podpis pracownika:", styles['InfoLine']))
            sig_img = Image(sig_img_path, width=80 * mm, height=20 * mm)
            sig_elements.append(sig_img)
            elements.extend(sig_elements)
        else:
            sig_line = Table([["Podpis pracownika: ...................................................."]],
                             colWidths=[page_width])
            sig_line.setStyle(TableStyle([
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('TOPPADDING', (0, 0), (-1, -1), 6),
            ]))
            elements.append(sig_line)

        doc.build(elements)


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