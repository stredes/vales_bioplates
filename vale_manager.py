#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
import logging
from typing import List, Optional, TypedDict

import pandas as pd

from data_loader import load_inventory
from pdf_utils import build_vale_pdf

logger = logging.getLogger(__name__)


class ValeItem(TypedDict):
    Producto: str
    Lote: str
    Vencimiento: str
    Ubicacion: str
    Cantidad: int
    Stock_Original_Index: int


@dataclass
class ValeManager:
    bioplates_inventory: pd.DataFrame = field(default_factory=pd.DataFrame)
    current_vale: List[ValeItem] = field(default_factory=list)

    def load(self, file_path: str, area_filter: Optional[str] = None) -> pd.DataFrame:
        logger.info("Solicitando carga de inventario (archivo=%s, area=%s)", file_path, area_filter)
        self.bioplates_inventory = load_inventory(file_path, area_filter)
        return self.bioplates_inventory

    def is_vale_empty(self) -> bool:
        return not self.current_vale

    def add_to_vale(self, item_index: int, quantity: int) -> ValeItem:
        df = self.bioplates_inventory
        product_data = df.loc[item_index]
        current_stock = int(product_data['Stock'])
        if quantity > current_stock:
            raise ValueError(f"No hay suficiente stock. Stock disponible: {current_stock}")

        self.bioplates_inventory.loc[item_index, 'Stock'] = current_stock - quantity

        new_item: ValeItem = {
            'Producto': str(product_data['Nombre_del_Producto']),
            'Lote': str(product_data['Lote']),
            'Vencimiento': str(product_data['Vencimiento']),
            'Ubicacion': str(product_data.get('Ubicacion', '')),
            'Cantidad': int(quantity),
            'Stock_Original_Index': int(item_index),
        }

        for it in self.current_vale:
            if (
                it['Producto'] == new_item['Producto']
                and it['Lote'] == new_item['Lote']
                and it.get('Ubicacion', '') == new_item.get('Ubicacion', '')
            ):
                it['Cantidad'] += int(quantity)
                logger.debug(
                    "Consolidado item %s (lote %s) nueva cantidad=%s",
                    it['Producto'],
                    it['Lote'],
                    it['Cantidad'],
                )
                return it

        logger.debug("Agregado item %s (lote %s) cantidad=%s", new_item['Producto'], new_item['Lote'], new_item['Cantidad'])
        self.current_vale.append(new_item)
        return new_item

    def remove_from_vale(self, vale_index: int) -> ValeItem:
        item = self.current_vale.pop(vale_index)
        self._restore_stock(item)
        logger.debug("Item removido del vale: %s (cantidad %s)", item['Producto'], item['Cantidad'])
        return item

    def clear_vale(self) -> None:
        for it in self.current_vale:
            self._restore_stock(it)
        self.current_vale = []

    def generate_pdf(self, filename: str, vale_data_with_users: Dict, emission_time: datetime) -> None:
        """Genera el PDF con la información del vale y usuarios."""
        build_vale_pdf(filename, vale_data_with_users, emission_time)

    def serialize_current_vale(self, emission_time: datetime) -> dict:
        return {
            'filename': None,
            'emission_time': emission_time.isoformat(timespec='seconds'),
            'item_count': self.total_items(),
            'total_quantity': self.total_quantity(),
            'items': self.current_vale,
        }

    def finalize_vale(self) -> None:
        self.current_vale = []

    def _restore_stock(self, item: ValeItem) -> None:
        stock_index = item['Stock_Original_Index']
        qty = int(item['Cantidad'])
        current = int(self.bioplates_inventory.loc[stock_index, 'Stock'])
        self.bioplates_inventory.loc[stock_index, 'Stock'] = current + qty

