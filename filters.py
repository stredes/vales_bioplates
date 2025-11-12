#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
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
        out = df.copy()

        # Texto de producto
        term = (self.producto or "").strip().lower()
        if term:
            out = out[out['Nombre_del_Producto'].str.lower().str.contains(term, na=False)]

        # Subfamilia exacta
        if self.subfamilia and self.subfamilia != '(Todas)' and 'Subfamilia' in out.columns:
            out = out[out['Subfamilia'].astype(str) == self.subfamilia]

        # Lote
        lote_t = (self.lote or "").strip().lower()
        if lote_t:
            out = out[out['Lote'].astype(str).str.lower().str.contains(lote_t, na=False)]

        # Ubicacion
        ubi_t = (self.ubicacion or "").strip().lower()
        if ubi_t and 'Ubicacion' in out.columns:
            out = out[out['Ubicacion'].astype(str).str.lower().str.contains(ubi_t, na=False)]

        # Vencimiento rango (YYYY-MM-DD)
        def _parse(d: str):
            try:
                return datetime.strptime(d, '%Y-%m-%d')
            except Exception:
                return None
        d1 = _parse(self.venc_desde) if self.venc_desde else None
        d2 = _parse(self.venc_hasta) if self.venc_hasta else None
        if d1 or d2:
            vdt = out['Vencimiento'].astype(str)
            vdt = vdt.where(vdt.ne('NaT'), None)

            def _ok(s: str) -> bool:
                try:
                    t = datetime.strptime(s, '%Y-%m-%d')
                except Exception:
                    return False
                if d1 and t < d1:
                    return False
                if d2 and t > d2:
                    return False
                return True

            out = out[vdt.apply(_ok)]

        # Solo con stock
        if self.solo_con_stock and 'Stock' in out.columns:
            out = out[out['Stock'] > 0]

        return out

