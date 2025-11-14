#!/usr/bin/env python
# -*- coding: utf-8 -*-

import logging
import os
from typing import Final

# Archivo de inventario por defecto (puede ser reemplazado desde la UI)
INVENTORY_FILE: Final[str] = "Informe_stock_fisico.xlsx"

# Filtro de área por defecto
AREA_FILTER: Final[str] = "Bioplates"

# Carpeta para guardar vales generados
HISTORY_DIR: Final[str] = "Vales_Historial"

# Márgenes PDF (en puntos)
PDF_MARGIN_LEFT: Final[int] = 50
PDF_MARGIN_RIGHT: Final[int] = 50
PDF_MARGIN_TOP: Final[int] = 50
PDF_MARGIN_BOTTOM: Final[int] = 50

# Detección simple de Windows para impresión automática
try:
    import win32api  # type: ignore
    WINDOWS_OS: Final[bool] = True
except Exception:
    WINDOWS_OS = False

# Ruta opcional a SumatraPDF.exe para impresión silenciosa en Windows.
# Puede configurarse mediante la variable de entorno 'SUMATRA_PDF_PATH'.
SUMATRA_PDF_PATH = os.environ.get("SUMATRA_PDF_PATH")

# Mostrar emojis en la UI (puede causar texto corrupto en algunos Windows)
USE_EMOJIS: Final[bool] = False

# Logging
LOG_LEVEL = os.environ.get("VALE_LOG_LEVEL", "INFO").upper()
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format="[%(asctime)s] %(levelname)s %(name)s: %(message)s",
)


