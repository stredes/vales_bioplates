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

        mask = pd.Series(True, index=df.index)

        # Texto de producto
        term = (self.producto or "").strip().lower()
        if term:
            col = df['_lc_producto'] if '_lc_producto' in df.columns else df['Nombre_del_Producto'].fillna('').astype(str).str.lower()
            mask &= col.str.contains(term, na=False)

        # Subfamilia exacta
        if self.subfamilia and self.subfamilia != '(Todas)' and 'Subfamilia' in df.columns:
            col = df['_subfam'] if '_subfam' in df.columns else df['Subfamilia'].astype(str)
            mask &= col == self.subfamilia

        # Lote
        lote_t = (self.lote or "").strip().lower()
        if lote_t and 'Lote' in df.columns:
            col = df['_lc_lote'] if '_lc_lote' in df.columns else df['Lote'].fillna('').astype(str).str.lower()
            mask &= col.str.contains(lote_t, na=False)

        # Ubicacion
        ubi_t = (self.ubicacion or "").strip().lower()
        if ubi_t and 'Ubicacion' in df.columns:
            col = df['_lc_ubicacion'] if '_lc_ubicacion' in df.columns else df['Ubicacion'].fillna('').astype(str).str.lower()
            mask &= col.str.contains(ubi_t, na=False)

        # Vencimiento rango (YYYY-MM-DD)
        d1 = pd.to_datetime(self.venc_desde, errors='coerce') if self.venc_desde else pd.NaT
        d2 = pd.to_datetime(self.venc_hasta, errors='coerce') if self.venc_hasta else pd.NaT
        if (pd.notna(d1) or pd.notna(d2)) and 'Vencimiento' in df.columns:
            vdt = df['_venc_dt'] if '_venc_dt' in df.columns else pd.to_datetime(df['Vencimiento'], errors='coerce')
            if pd.notna(d1):
                mask &= vdt >= d1
            if pd.notna(d2):
                mask &= vdt <= d2

        # Solo con stock
        if self.solo_con_stock and 'Stock' in df.columns:
            mask &= df['Stock'] > 0

        return df[mask]

