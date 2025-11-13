#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os

# Archivo de inventario por defecto (puede ser reemplazado desde la UI)
INVENTORY_FILE = "Informe_stock_fisico.xlsx"

# Filtro de área por defecto
AREA_FILTER = "Bioplates"

# Carpeta para guardar vales generados
HISTORY_DIR = "Vales_Historial"

# Márgenes PDF (en puntos)
PDF_MARGIN_LEFT = 50
PDF_MARGIN_RIGHT = 50
PDF_MARGIN_TOP = 50
PDF_MARGIN_BOTTOM = 50

# Detección simple de Windows para impresión automática
try:
    import win32api  # type: ignore
    WINDOWS_OS = True
except Exception:
    WINDOWS_OS = False

# Ruta opcional a SumatraPDF.exe para impresión silenciosa en Windows.
# Puede configurarse mediante la variable de entorno 'SUMATRA_PDF_PATH'.
SUMATRA_PDF_PATH = os.environ.get("SUMATRA_PDF_PATH")

# Mostrar emojis en la UI (puede causar texto corrupto en algunos Windows)
USE_EMOJIS = False


