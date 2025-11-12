#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Dict, Optional

import pandas as pd

from data_loader import load_inventory
from pdf_utils import build_vale_pdf


@dataclass
class ValeManager:
    bioplates_inventory: pd.DataFrame = field(default_factory=pd.DataFrame)
    current_vale: List[Dict] = field(default_factory=list)

    def load(self, file_path: str, area_filter: Optional[str] = None) -> pd.DataFrame:
        self.bioplates_inventory = load_inventory(file_path, area_filter)
        return self.bioplates_inventory

    def add_to_vale(self, item_index: int, quantity: int) -> Dict:
        df = self.bioplates_inventory
        product_data = df.loc[item_index]
        current_stock = int(product_data['Stock'])
        if quantity > current_stock:
            raise ValueError(f"No hay suficiente stock. Stock disponible: {current_stock}")

        # Descontar stock
        self.bioplates_inventory.loc[item_index, 'Stock'] = current_stock - quantity

        new_item = {
            'Producto': product_data['Nombre_del_Producto'],
            'Lote': product_data['Lote'],
            'Vencimiento': product_data['Vencimiento'],
            'Ubicacion': product_data.get('Ubicacion', ''),
            'Cantidad': int(quantity),
            'Stock_Original_Index': int(item_index),
        }

        # Consolidar si mismo producto/lote/ubicacion
        for it in self.current_vale:
            if (
                it['Producto'] == new_item['Producto'] and
                it['Lote'] == new_item['Lote'] and
                it.get('Ubicacion', '') == new_item.get('Ubicacion', '')
            ):
                it['Cantidad'] += int(quantity)
                return it

        self.current_vale.append(new_item)
        return new_item

    def remove_from_vale(self, vale_index: int) -> Dict:
        item = self.current_vale.pop(vale_index)
        stock_index = item['Stock_Original_Index']
        qty = int(item['Cantidad'])
        self.bioplates_inventory.loc[stock_index, 'Stock'] = int(self.bioplates_inventory.loc[stock_index, 'Stock']) + qty
        return item

    def clear_vale(self) -> None:
        for it in self.current_vale:
            idx = it['Stock_Original_Index']
            self.bioplates_inventory.loc[idx, 'Stock'] = int(self.bioplates_inventory.loc[idx, 'Stock']) + int(it['Cantidad'])
        self.current_vale = []

    def generate_pdf(self, filename: str, emission_time: datetime) -> None:
        build_vale_pdf(filename, self.current_vale, emission_time)

