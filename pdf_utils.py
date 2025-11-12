#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

from typing import List, Dict
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.pdfbase.pdfmetrics import stringWidth

from datetime import datetime

from config import (
    PDF_MARGIN_LEFT,
    PDF_MARGIN_RIGHT,
    PDF_MARGIN_TOP,
    PDF_MARGIN_BOTTOM,
)


def _para(text: str, style_name: str = "Normal") -> Paragraph:
    styles = getSampleStyleSheet()
    if style_name == "NormalWrap":
        st = ParagraphStyle("NormalWrap", parent=styles["Normal"], fontSize=10, leading=12)
        return Paragraph(text.replace("&", "&amp;"), st)
    return Paragraph(text.replace("&", "&amp;"), styles[style_name])


def _auto_col_widths(rows: List[List[str]], page_width: int, padd: int = 16) -> List[int]:
    # Calcula anchos según el contenido + padding, y ajusta al ancho disponible
    # Fuente asumida: Helvetica 10
    font_name = "Helvetica"
    font_size = 10
    n_cols = len(rows[0])
    maxw = [0] * n_cols
    for r in rows:
        for i, cell in enumerate(r):
            s = str(cell)
            w = stringWidth(s, font_name, font_size) + padd
            if w > maxw[i]:
                maxw[i] = w
    total = sum(maxw)
    avail = page_width - PDF_MARGIN_LEFT - PDF_MARGIN_RIGHT
    if total <= avail:
        return maxw
    # Escalar proporcionalmente
    scale = avail / total if total else 1.0
    return [max(60, int(w * scale)) for w in maxw]


def build_vale_pdf(filename: str, vale_data: List[Dict], emission_time: datetime) -> None:
    doc = SimpleDocTemplate(
        filename,
        pagesize=letter,
        rightMargin=PDF_MARGIN_RIGHT,
        leftMargin=PDF_MARGIN_LEFT,
        topMargin=PDF_MARGIN_TOP,
        bottomMargin=PDF_MARGIN_BOTTOM,
    )
    styles = getSampleStyleSheet()
    elements = []

    title_style = ParagraphStyle("TitleStyle", parent=styles["Title"], fontSize=16, alignment=1)
    elements.append(Paragraph("<b>VALE DE CONSUMO – BIOPLATES</b>", title_style))
    elements.append(Spacer(1, 12))

    metadata = [
        ["Fecha de Emisión:", emission_time.strftime("%d-%m-%Y")],
        ["Hora de Emisión:", emission_time.strftime("%H:%M:%S")],
    ]

    metadata_table = Table(metadata, colWidths=[150, 300])
    metadata_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    elements.append(metadata_table)
    elements.append(Spacer(1, 18))

    # Datos de tabla: Producto, Lote, Ubicacion, Vencimiento, Cantidad
    headers = ["Producto", "Lote", "Ubicación", "Vencimiento", "Cantidad"]
    rows = [headers]
    for item in vale_data:
        rows.append(
            [
                item.get("Producto", ""),
                item.get("Lote", ""),
                item.get("Ubicacion", ""),
                item.get("Vencimiento", ""),
                str(item.get("Cantidad", "")),
            ]
        )

    # Anchos automáticos
    page_w, _ = letter
    col_widths = _auto_col_widths(rows, page_w)

    # Convertir Producto y Ubicación a Paragraph para permitir wrap
    table_rows = [rows[0]]
    for r in rows[1:]:
        table_rows.append([
            _para(r[0], "NormalWrap"),  # Producto
            r[1],                        # Lote
            _para(r[2], "NormalWrap"),  # Ubicacion
            r[3],                        # Vencimiento
            r[4],                        # Cantidad
        ])

    table = Table(table_rows, colWidths=col_widths)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("ALIGN", (0, 0), (0, -1), "LEFT"),  # Producto
                ("ALIGN", (2, 0), (2, -1), "LEFT"),  # Ubicacion
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ]
        )
    )

    elements.append(Paragraph("<b>Detalle de Productos Retirados:</b>", styles["Normal"]))
    elements.append(Spacer(1, 6))
    elements.append(table)
    elements.append(Spacer(1, 48))

    signature_data = [["________________________", "________________________"], ["Entregado por:", "Recibido por:"]]
    signature_table = Table(signature_data, colWidths=[250, 250])
    signature_table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica-Bold"),
                ("TOPPADDING", (0, 0), (-1, 0), 10),
            ]
        )
    )
    elements.append(signature_table)

    doc.build(elements)

