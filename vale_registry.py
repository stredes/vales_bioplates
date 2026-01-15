#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

import json
import os
from datetime import datetime
from typing import Any, Dict, List, Optional


class ValeRegistry:
    """Registro simple basado en JSON para numerar y llevar estados de solicitudes.

    Estructura del archivo:
    {
      "sequence": 15,
      "vales": [
        {
          "number": 1,
          "status": "Pendiente",  # Pendiente | Descontado | Anulado
          "created_at": "2025-11-13T15:04:09",
          "pdf": "solicitud_001_20251113_150409.pdf",
          "json": "solicitud_001_20251113_150409.json",
          "items_count": 3
        },
        ...
      ]
    }
    """

    def __init__(self, history_dir: str) -> None:
        self.history_dir = history_dir
        self.index_path = os.path.join(history_dir, 'vales_index.json')
        self.data: Dict[str, Any] = {}
        self._load()

    # ---------------- Internals ----------------
    def _load(self) -> None:
        if not os.path.exists(self.history_dir):
            try:
                os.makedirs(self.history_dir, exist_ok=True)
            except Exception:
                pass
        if os.path.exists(self.index_path):
            try:
                with open(self.index_path, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
            except Exception:
                self.data = {}
        if not isinstance(self.data, dict):
            self.data = {}
        self.data.setdefault('sequence', 0)
        self.data.setdefault('vales', [])

    def _save(self) -> None:
        try:
            with open(self.index_path, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ---------------- API pública ----------------
    def next_number(self) -> int:
        n = int(self.data.get('sequence', 0)) + 1
        self.data['sequence'] = n
        self._save()
        return n

    def register_with_number(self, number: int, pdf_filename: str, json_filename: Optional[str], items_count: int) -> Dict[str, Any]:
        """Registra un vale usando un numero ya reservado.

        Asegura que la secuencia no quede por debajo del numero asignado.
        """
        try:
            cur_seq = int(self.data.get('sequence', 0))
        except Exception:
            cur_seq = 0
        if number > cur_seq:
            self.data['sequence'] = number
        entry = {
            'number': int(number),
            'status': 'Pendiente',
            'created_at': datetime.now().isoformat(timespec='seconds'),
            'pdf': os.path.basename(pdf_filename),
            'json': os.path.basename(json_filename) if json_filename else '',
            'items_count': int(items_count),
        }
        self.data['vales'].append(entry)
        self._save()
        return entry

    def register_voucher(self, pdf_filename: str, json_filename: Optional[str], items_count: int) -> Dict[str, Any]:
        number = self.next_number()
        entry = {
            'number': number,
            'status': 'Pendiente',
            'created_at': datetime.now().isoformat(timespec='seconds'),
            'pdf': os.path.basename(pdf_filename),
            'json': os.path.basename(json_filename) if json_filename else '',
            'items_count': int(items_count),
        }
        self.data['vales'].append(entry)
        self._save()
        return entry

    def list(self, status: Optional[str] = None) -> List[Dict[str, Any]]:
        items: List[Dict[str, Any]] = list(self.data.get('vales', []))
        if status:
            items = [x for x in items if x.get('status') == status]
        # ordenar por fecha desc
        def _key(e: Dict[str, Any]):
            return e.get('created_at', '')
        items.sort(key=_key, reverse=True)
        return items

    def update_status(self, numbers: List[int], new_status: str) -> int:
        count = 0
        for e in self.data.get('vales', []):
            if int(e.get('number', -1)) in numbers:
                e['status'] = new_status
                count += 1
        if count:
            self._save()
        return count

    def update_entry(self, number: int, **updates: Any) -> bool:
        changed = False
        for e in self.data.get('vales', []):
            if int(e.get('number', -1)) == int(number):
                for k, v in updates.items():
                    if e.get(k) != v:
                        e[k] = v
                        changed = True
                break
        if changed:
            self._save()
        return changed

    def find_by_number(self, number: int) -> Optional[Dict[str, Any]]:
        for e in self.data.get('vales', []):
            if int(e.get('number', -1)) == number:
                return e
        return None

    # ---------------- Importación de vales antiguos ----------------
    def _has_pdf(self, pdf_name: str) -> bool:
        pdf_base = os.path.basename(pdf_name)
        for e in self.data.get('vales', []):
            if os.path.basename(e.get('pdf', '')) == pdf_base:
                return True
        return False

    def reindex(self) -> Dict[str, int]:
        """Escanea la carpeta de historial y agrega al índice los PDFs que no estén
        registrados. Asigna números correlativos y estado Pendiente.

        Devuelve {'added': n, 'skipped': m}
        """
        added = 0
        skipped = 0
        try:
            files = [f for f in os.listdir(self.history_dir) if f.lower().endswith('.pdf')]
        except Exception:
            files = []
        def _mtime(name: str) -> float:
            try:
                return os.path.getmtime(os.path.join(self.history_dir, name))
            except Exception:
                return 0.0
        files.sort(key=_mtime)
        for f in files:
            if self._has_pdf(f):
                skipped += 1
                continue
            base, _ = os.path.splitext(f)
            jpath = os.path.join(self.history_dir, base + '.json')
            items_count = 0
            created_iso = None
            if os.path.exists(jpath):
                try:
                    with open(jpath, 'r', encoding='utf-8') as jf:
                        data = json.load(jf)
                    items = data.get('items', [])
                    items_count = int(len(items)) if isinstance(items, list) else 0
                    created_iso = data.get('emission_time')
                except Exception:
                    pass
            if not created_iso:
                try:
                    created_iso = datetime.fromtimestamp(_mtime(f)).isoformat(timespec='seconds')
                except Exception:
                    created_iso = datetime.now().isoformat(timespec='seconds')
            number = self.next_number()
            entry = {
                'number': number,
                'status': 'Pendiente',
                'created_at': created_iso,
                'pdf': f,
                'json': (os.path.basename(base + '.json') if os.path.exists(jpath) else ''),
                'items_count': items_count,
            }
            self.data['vales'].append(entry)
            added += 1
        if added:
            self._save()
        return {'added': added, 'skipped': skipped}
