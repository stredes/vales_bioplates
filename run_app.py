#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Punto de entrada para la app de Vales de Consumo (Bioplates).

Ejecuta la interfaz Tkinter importando `run_app` desde `vale_consumo_bioplates`.

Uso:
  python run_app.py
"""

try:
    from vale_consumo_bioplates import run_app
except Exception as e:
    raise SystemExit(f"Error importando la aplicación: {e}")


if __name__ == "__main__":
    run_app()

