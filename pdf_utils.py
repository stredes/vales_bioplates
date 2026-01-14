#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

import logging
from datetime import datetime
from typing import Dict, Iterable, List

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from config import PDF_MARGIN_BOTTOM, PDF_MARGIN_LEFT, PDF_MARGIN_RIGHT, PDF_MARGIN_TOP

logger = logging.getLogger(__name__)


def _para(text: str, style_name: str = "Normal") -> Paragraph:
    styles = getSampleStyleSheet()
    if style_name == "NormalWrap":
        st = ParagraphStyle("NormalWrap", parent=styles["Normal"], fontSize=10, leading=12)
        return Paragraph(text.replace("&", "&amp;"), st)
    return Paragraph(text.replace("&", "&amp;"), styles[style_name])

def _format_date_ddmmyyyy(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    text = str(value).strip()
    if not text:
        return ""
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y")
        except Exception:
            continue
    return text


def _auto_col_widths(rows: Iterable[Iterable[str]], page_width: int, padd: int = 16) -> List[int]:
    # Calcula anchos segun el contenido + padding, y ajusta al ancho disponible
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


def build_vale_pdf(filename: str, vale_data_with_users: Dict, emission_time: datetime) -> None:
    """Genera un PDF de solicitud de productos (uso bodega).
    
    Args:
        filename: Ruta donde guardar el PDF
        vale_data_with_users: Dict con 'solicitante', 'usuario_bodega', 'numero_correlativo' e 'items'
        emission_time: Fecha y hora de emisión
    """
    footer_height = 90
    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        rightMargin=PDF_MARGIN_RIGHT,
        leftMargin=PDF_MARGIN_LEFT,
        topMargin=PDF_MARGIN_TOP,
        bottomMargin=PDF_MARGIN_BOTTOM + footer_height,
    )
    styles = getSampleStyleSheet()
    elements = []

    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Title"],
        fontSize=16,
        alignment=1,
        textColor=colors.HexColor("#1f2a44"),
        leading=18,
    )
    elements.append(Paragraph("<b>SOLICITUD DE PRODUCTOS<br/>(USO BODEGA)</b>", title_style))
    elements.append(Spacer(1, 14))

    # Obtener datos
    solicitante = vale_data_with_users.get('solicitante', '')
    usuario_bodega = vale_data_with_users.get('usuario_bodega', '')
    numero_correlativo = vale_data_with_users.get('numero_correlativo', '')
    vale_data = vale_data_with_users.get('items', [])

    label_style = ParagraphStyle(
        "MetaLabel",
        parent=styles["Normal"],
        fontSize=9,
        textColor=colors.HexColor("#5a5a5a"),
        leading=12,
    )
    value_style = ParagraphStyle(
        "MetaValue",
        parent=styles["Normal"],
        fontSize=11,
        textColor=colors.HexColor("#111111"),
        leading=12,
    )
    meta_rows = [
        [
            Paragraph("N° Solicitud", label_style),
            Paragraph("Fecha de Emision", label_style),
            Paragraph("Hora de Emision", label_style),
        ],
        [
            Paragraph(str(numero_correlativo), value_style),
            Paragraph(emission_time.strftime("%d/%m/%Y"), value_style),
            Paragraph(emission_time.strftime("%H:%M:%S"), value_style),
        ],
        [
            Paragraph("Solicitado por", label_style),
            Paragraph("Preparado por", label_style),
            "",
        ],
        [
            Paragraph(str(solicitante), value_style),
            Paragraph(str(usuario_bodega), value_style),
            "",
        ],
    ]

    metadata_table = Table(meta_rows, colWidths=[150, 170, 130])
    metadata_table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("LINEBELOW", (0, 1), (2, 1), 0.6, colors.HexColor("#e0e0e0")),
                ("LINEBELOW", (0, 3), (1, 3), 0.6, colors.HexColor("#e0e0e0")),
            ]
        )
    )
    elements.append(metadata_table)
    elements.append(Spacer(1, 10))
    elements.append(HRFlowable(width="100%", thickness=0.6, color=colors.lightgrey, spaceBefore=2, spaceAfter=12))

    # Datos de tabla: Codigo, Producto, Lote, Ubicacion, Vencimiento, Stock, Cantidad
    headers = ["Codigo", "Producto", "Lote", "Ubicacion", "Vencimiento", "Stock", "Cantidad"]
    rows = [headers]
    for item in vale_data:
        rows.append(
            [
                item.get("Codigo", ""),
                item.get("Producto", ""),
                item.get("Lote", ""),
                item.get("Ubicacion", ""),
                _format_date_ddmmyyyy(item.get("Vencimiento", "")),
                str(item.get("Stock", "")),
                str(item.get("Cantidad", "")),
            ]
        )

    # Anchos automÃ¡ticos
    page_w, _ = A4
    col_widths = _auto_col_widths(rows, page_w)

    # Convertir Producto y Ubicacion a Paragraph para permitir wrap
    table_rows = [rows[0]]
    for r in rows[1:]:
        table_rows.append([
            r[0],                        # Codigo
            _para(r[1], "NormalWrap"),  # Producto
            r[2],                        # Lote
            _para(r[3], "NormalWrap"),  # Ubicacion
            r[4],                        # Vencimiento
            r[5],                        # Stock
            r[6],                        # Cantidad
        ])

    table = Table(table_rows, colWidths=col_widths)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2f3b52")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("ALIGN", (1, 0), (1, -1), "LEFT"),  # Producto
                ("ALIGN", (3, 0), (3, -1), "LEFT"),  # Ubicacion
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 10),
                ("TOPPADDING", (0, 0), (-1, 0), 6),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#f7f7f7")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#ffffff"), colors.HexColor("#f1f3f6")]),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#c9c9c9")),
            ]
        )
    )

    section_style = ParagraphStyle(
        "SectionTitle",
        parent=styles["Normal"],
        fontSize=11,
        textColor=colors.HexColor("#1f2a44"),
        spaceAfter=6,
    )
    elements.append(Paragraph("<b>Detalle de Productos Retirados</b>", section_style))
    elements.append(Spacer(1, 4))
    elements.append(table)
    elements.append(Spacer(1, 18))

    def _draw_footer(canvas, _doc) -> None:
        canvas.saveState()
        page_w, _ = A4
        left = PDF_MARGIN_LEFT
        right = page_w - PDF_MARGIN_RIGHT

        footer_top = PDF_MARGIN_BOTTOM + footer_height - 10
        canvas.setStrokeColor(colors.lightgrey)
        canvas.setLineWidth(0.6)
        canvas.line(left, footer_top, right, footer_top)

        line_y = PDF_MARGIN_BOTTOM + 55
        label_y = line_y - 12
        line_len = 180
        left_start = left + 10
        right_end = right - 10
        right_start = right_end - line_len

        canvas.setStrokeColor(colors.black)
        canvas.setLineWidth(0.8)
        canvas.line(left_start, line_y, left_start + line_len, line_y)
        canvas.line(right_start, line_y, right_end, line_y)

        canvas.setFont("Helvetica-Bold", 9)
        canvas.setFillColor(colors.black)
        canvas.drawCentredString(left_start + (line_len / 2), label_y, "Entregado por")
        canvas.drawCentredString(right_start + (line_len / 2), label_y, "Recibido por")

        doc_label_y = PDF_MARGIN_BOTTOM + 16
        canvas.setStrokeColor(colors.grey)
        canvas.setLineWidth(0.6)
        canvas.setDash(3, 3)
        canvas.line(left, doc_label_y + 2, right, doc_label_y + 2)
        canvas.setDash()
        canvas.setFont("Helvetica", 9)
        canvas.setFillColor(colors.grey)
        canvas.drawCentredString(
            (left + right) / 2,
            doc_label_y - 8,
            "Doc. Asociado: _______________________",
        )
        canvas.restoreState()

    doc.build(elements, onFirstPage=_draw_footer, onLaterPages=_draw_footer)
    logger.info("PDF generado correctamente en %s", filename)



def build_unified_vale_pdf(filename: str, rows: List[Dict], emission_time: datetime) -> None:
    """Genera un PDF de vale unificado en una sola tabla.

    Espera filas con claves: Producto, Lote, Ubicacion, Vencimiento, Cantidad, Origen
    """
    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        rightMargin=PDF_MARGIN_RIGHT,
        leftMargin=PDF_MARGIN_LEFT,
        topMargin=PDF_MARGIN_TOP,
        bottomMargin=PDF_MARGIN_BOTTOM,
    )
    styles = getSampleStyleSheet()
    elements = []

    title_style = ParagraphStyle("TitleStyle", parent=styles["Title"], fontSize=16, alignment=1)
    elements.append(Paragraph("<b>SOLICITUD DE PRODUCTOS UNIFICADA<br/>(USO BODEGA)</b>", title_style))
    elements.append(Spacer(1, 12))

    metadata = [
        ["Fecha de Emision:", emission_time.strftime("%d/%m/%Y")],
        ["Hora de Emision:", emission_time.strftime("%H:%M:%S")],
        ["Documentos origen:", str(len({r.get('Origen','') for r in rows if r.get('Origen')}))],
    ]
    metadata_table = Table(metadata, colWidths=[150, 300])
    metadata_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    elements.append(metadata_table)
    elements.append(Spacer(1, 18))

    headers = ["Origen", "Producto", "Lote", "Ubicacion", "Vencimiento", "Cantidad"]
    raw_rows = [headers]
    for r in rows:
        raw_rows.append([
            r.get("Origen", ""),
            r.get("Producto", ""),
            r.get("Lote", ""),
            r.get("Ubicacion", ""),
            _format_date_ddmmyyyy(r.get("Vencimiento", "")),
            str(r.get("Cantidad", "")),
        ])

    page_w, _ = A4
    col_widths = _auto_col_widths(raw_rows, page_w)

    table_rows = [raw_rows[0]]
    for r in raw_rows[1:]:
        table_rows.append([
            r[0],
            _para(r[1], "NormalWrap"),  # Producto
            r[2],                        # Lote
            _para(r[3], "NormalWrap"),  # Ubicacion
            r[4],                        # Vencimiento
            r[5],                        # Cantidad
        ])

    table = Table(table_rows, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("Fin del vale unificado.", styles["Normal"]))
    doc.build(elements)


def build_vales_list_pdf(filename: str, title: str, rows: List[Dict]) -> None:
    """Genera un listado de vales (Numero, Estado, Fecha, Archivo, Items)."""
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
    elements.append(Paragraph(title, title_style))
    elements.append(Spacer(1, 12))

    headers = ["Numero", "Estado", "Fecha", "Archivo", "Items"]
    raw_rows = [headers]
    for r in rows:
        raw_rows.append([
            str(r.get('number', '')),
            str(r.get('status', '')),
            str(r.get('created_at', '')),
            str(r.get('pdf', '')),
            str(r.get('items_count', '')),
        ])

    page_w, _ = letter
    col_widths = _auto_col_widths(raw_rows, page_w)

    table = Table(raw_rows, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
    ]))
    elements.append(table)
    doc.build(elements)
