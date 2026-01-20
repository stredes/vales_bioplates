#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Optional

import pandas as pd


@dataclass
class FilterOptions:
    producto: str = ""
    lote: str = ""
    ubicacion: str = ""
    venc_desde: str = ""
    venc_hasta: str = ""
    subfamilia: str = "(Todas)"
    solo_con_stock: bool = False

    def apply(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return df

        out = df

        # Solo con stock (filtro barato primero)
        if self.solo_con_stock and 'Stock' in out.columns:
            out = out[out['Stock'] > 0]
            if out.empty:
                return out

        # Subfamilia exacta
        if self.subfamilia and self.subfamilia != '(Todas)' and 'Subfamilia' in out.columns:
            col = out['_subfam'] if '_subfam' in out.columns else out['Subfamilia'].astype(str)
            out = out[col == self.subfamilia]
            if out.empty:
                return out

        # Vencimiento rango (YYYY-MM-DD)
        d1 = pd.to_datetime(self.venc_desde, errors='coerce') if self.venc_desde else pd.NaT
        d2 = pd.to_datetime(self.venc_hasta, errors='coerce') if self.venc_hasta else pd.NaT
        if (pd.notna(d1) or pd.notna(d2)) and 'Vencimiento' in out.columns:
            vdt = out['_venc_dt'] if '_venc_dt' in out.columns else pd.to_datetime(out['Vencimiento'], errors='coerce')
            if pd.notna(d1):
                out = out[vdt >= d1]
            if pd.notna(d2):
                out = out[vdt <= d2]
            if out.empty:
                return out

        # Lote
        lote_t = (self.lote or "").strip().lower()
        if lote_t and 'Lote' in out.columns:
            col = out['_lc_lote'] if '_lc_lote' in out.columns else out['Lote'].fillna('').astype(str).str.lower()
            out = out[col.str.contains(lote_t, na=False, regex=False)]
            if out.empty:
                return out

        # Ubicacion
        ubi_t = (self.ubicacion or "").strip().lower()
        if ubi_t and 'Ubicacion' in out.columns:
            col = out['_lc_ubicacion'] if '_lc_ubicacion' in out.columns else out['Ubicacion'].fillna('').astype(str).str.lower()
            out = out[col.str.contains(ubi_t, na=False, regex=False)]
            if out.empty:
                return out

        # Texto de producto/codigo (al final para reducir filas)
        term = (self.producto or "").strip().lower()
        if term:
            prod_col = out['_lc_producto'] if '_lc_producto' in out.columns else out['Nombre_del_Producto'].fillna('').astype(str).str.lower()
            if 'Codigo' in out.columns:
                code_col = out['_lc_codigo'] if '_lc_codigo' in out.columns else out['Codigo'].fillna('').astype(str).str.lower()
                out = out[prod_col.str.contains(term, na=False, regex=False) | code_col.str.contains(term, na=False, regex=False)]
            else:
                out = out[prod_col.str.contains(term, na=False, regex=False)]

        return out

