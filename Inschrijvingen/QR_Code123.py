#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Maak een A4-PDF met drie QR-codes naast elkaar.
Onder de QR-codes staat enkel "1", "2" en "3".
"""

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.graphics import renderPDF
from reportlab.lib.units import mm
from dataclasses import dataclass
# ---------- CONFIG ----------
OUTPUT_PDF = "QRcodes_123.pdf"

# Waarden die de scanner leest (en ook als label getoond worden)
ITEMS = ["1", "2", "3"]

QR_SIZE_MM = 60           # grootte QR-code (vierkant)
PAGE_MARGIN_MM = 10       # linker/rechter marge
LABEL_FONT = "Helvetica-Bold"
LABEL_FONT_SIZE = 18
LABEL_OFFSET_MM = 10      # ruimte tussen QR en label
ROW_Y_POSITION = 198      # Y-positie in mm vanaf onderkant pagina (ongeveer midden A4)
# ---------------------------


@dataclass
class Layout:
    page_width: float
    page_height: float
    margin: float
    qr_size: float
    spacing: float
    y: float


def mm_(x):
    """Helper: mm naar punten (reportlab werkt intern met punten)."""
    return x * mm


def draw_qr(c, x, y, size, text):
    """Teken een QR-code met inhoud 'text' op (x,y) met zijde 'size'."""
    qr_code = qr.QrCodeWidget(text)
    bounds = qr_code.getBounds()
    w = bounds[2] - bounds[0]
    h = bounds[3] - bounds[1]
    d = Drawing(size, size, transform=[size / w, 0, 0, size / h, 0, 0])
    d.add(qr_code)
    renderPDF.draw(d, c, x, y)


def main():
    page = A4
    c = canvas.Canvas(OUTPUT_PDF, pagesize=page)
    width, height = page

    n = len(ITEMS)
    qr_size = mm_(QR_SIZE_MM)
    margin = mm_(PAGE_MARGIN_MM)

    # Bereken tussenruimte zodat alles netjes past
    total_qr_width = n * qr_size
    spacing = (width - 2 * margin - total_qr_width) / (n - 1)
    y = mm_(ROW_Y_POSITION)

    # X-posities van de QR-codes
    x_positions = [margin + i * (qr_size + spacing) for i in range(n)]

    for i, val in enumerate(ITEMS):
        x = x_positions[i]

        # QR-code tekenen
        draw_qr(c, x, y, qr_size, val)

        # Label onder QR
        c.setFont(LABEL_FONT, LABEL_FONT_SIZE)
        label_y = y - mm_(LABEL_OFFSET_MM)
        c.drawCentredString(x + qr_size / 2, label_y, val)

    c.save()
    print(f"PDF aangemaakt: {OUTPUT_PDF}")


if __name__ == "__main__":
    main()
