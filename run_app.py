#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Punto de entrada para la app de Vales de Consumo (Bioplates).

Ejecuta la interfaz Tkinter importando `run_app` desde `vale_consumo_bioplates`.

Uso:
  python run_app.py
"""

import os, sys
# Asegura que el directorio del script está en sys.path para resolver imports locales
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

try:
    # Pylance/pyright puede no resolver este import en modo script suelto.
    # type: ignore[import-not-found]
    from vale_consumo_bioplates import run_app  # type: ignore[import-not-found]
except Exception:
    # Fallback: cargar por ruta si el import directo falla en entornos del analizador
    mod_path = os.path.join(SCRIPT_DIR, 'vale_consumo_bioplates.py')
    if os.path.exists(mod_path):
        import importlib.util
        spec = importlib.util.spec_from_file_location('vale_consumo_bioplates', mod_path)
        if spec and spec.loader:
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)  # type: ignore[attr-defined]
            run_app = getattr(module, 'run_app')
        else:
            raise SystemExit('No se pudo cargar vale_consumo_bioplates.py')
    else:
        raise SystemExit('No se encontro vale_consumo_bioplates.py')


if __name__ == "__main__":
    run_app()
