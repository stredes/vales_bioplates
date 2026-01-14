#!/usr/bin/env python
# -*- coding: utf-8 -*-

import json
import os
from typing import Optional
try:
    import config as _config
except Exception:  # fallback si no hay config en runtime
    class _Dummy:
        HISTORY_DIR = 'Vales_Historial'
    _config = _Dummy()

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


# --- Printing preferences ---
def get_auto_print() -> bool:
    return bool(_load().get('auto_print', False))


def set_auto_print(enabled: bool) -> None:
    data = _load()
    data['auto_print'] = bool(enabled)
    _save(data)


def get_sumatra_path() -> str | None:
    val = _load().get('sumatra_path')
    return val if isinstance(val, str) and val else None


def set_sumatra_path(path: str) -> None:
    data = _load()
    data['sumatra_path'] = path
    _save(data)


# --- File label reminder preferences ---
def get_reminder_enabled() -> bool:
    return bool(_load().get('reminder_enabled', True))


def set_reminder_enabled(enabled: bool) -> None:
    data = _load()
    data['reminder_enabled'] = bool(enabled)
    _save(data)


def get_reminder_text() -> str:
    val = _load().get('reminder_text')
    return val if isinstance(val, str) and val else 'Recordar solicitar archivo actualizado'


def set_reminder_text(text: str) -> None:
    data = _load()
    data['reminder_text'] = text
    _save(data)


# --- History folder preferences ---
def get_history_dir() -> str:
    val = _load().get('history_dir')
    if isinstance(val, str) and val:
        return val
    return getattr(_config, 'HISTORY_DIR', 'Vales_Historial')


def set_history_dir(path: str) -> None:
    data = _load()
    data['history_dir'] = path
    _save(data)
