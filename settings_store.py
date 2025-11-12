#!/usr/bin/env python
# -*- coding: utf-8 -*-

import json
import os
from typing import Optional

SETTINGS_FILE = 'app_settings.json'


def _load() -> dict:
    try:
        with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _save(data: dict) -> None:
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def get_last_inventory_dir() -> Optional[str]:
    data = _load()
    path = data.get('last_inventory_dir')
    if path and os.path.isdir(path):
        return path
    return None


def set_last_inventory_dir(path: str) -> None:
    data = _load()
    data['last_inventory_dir'] = path
    _save(data)

