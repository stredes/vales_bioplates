#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import subprocess
from tkinter import messagebox

import config as app_config



def _find_sumatra_exe():
    """Return path to SumatraPDF.exe if present via config/env or typical paths."""
    try:
        sp = getattr(app_config, 'SUMATRA_PDF_PATH', None)
    except Exception:
        sp = None
    if sp and os.path.exists(sp):
        return sp
    candidates = [
        os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "SumatraPDF", "SumatraPDF.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "SumatraPDF", "SumatraPDF.exe"),
    ]
    for p in candidates:
        if p and os.path.exists(p):
            return p
    return None

def print_pdf_windows(filename: str, copies: int = 1) -> None:
    """Imprime un PDF en Windows con varias rutas de respaldo.

    Intentos:
      1) ShellExecute 'print' (requiere asociación de impresión para .pdf)
      2) ShellExecute 'printto' con impresora predeterminada
      3) SumatraPDF en modo silencioso si está instalado
    """
    if not getattr(app_config, 'WINDOWS_OS', False):
        messagebox.showerror(
            "Error de Impresión",
            "La impresión automática requiere Windows y el módulo pywin32.",
        )
        return

    if not os.path.exists(filename):
        messagebox.showerror("Error de Impresión", "El archivo PDF no existe.")
        return

    try:
        import win32api  # type: ignore
        import win32print  # type: ignore

        # Intento prioritario: SumatraPDF si está disponible (evita errores de asociación)
        runner_pref = None
        try:
            # función añadida arriba
            runner_pref = _find_sumatra_exe()
        except Exception:
            runner_pref = None
        if runner_pref:
            try:
                for _ in range(max(1, int(copies))):
                    cmd = [runner_pref, "-silent", "-print-to-default", filename]
                    subprocess.run(cmd, check=True, creationflags=0x08000000)
                messagebox.showinfo(
                    "Impresión Enviada",
                    f"Se han enviado {copies} copias del vale a la impresora predeterminada.",
                )
                return
            except Exception:
                # seguirá con ShellExecute
                pass

        # Impresora predeterminada
        try:
            default_printer = win32print.GetDefaultPrinter()
        except Exception:
            default_printer = None

        if not default_printer:
            messagebox.showerror(
                "Error de Impresión",
                "No hay impresora predeterminada configurada en Windows. Configure una e intente nuevamente.",
            )
            return

        last_err = None
        for _ in range(max(1, int(copies))):
            try:
                # 1) Verbo 'print'
                win32api.ShellExecute(0, "print", filename, None, ".", 0)
                last_err = None
            except Exception as e1:
                last_err = e1
                # 2) Verbo 'printto' con nombre de impresora
                try:
                    win32api.ShellExecute(0, "printto", filename, f'"{default_printer}"', ".", 0)
                    last_err = None
                except Exception as e2:
                    last_err = e2

            if last_err:
                break

        if last_err:
            # 3) SumatraPDF (si existe o se definió vía config)
            runner = _find_sumatra_exe()
            if runner:
                try:
                    cmd = [runner, "-silent", "-print-to", default_printer, filename]
                    subprocess.run(cmd, check=True, creationflags=0x08000000)
                    last_err = None
                except Exception as e3:
                    last_err = e3

        if last_err:
            msg = str(last_err)
            if "31" in msg:  # SE_ERR_NOASSOC
                messagebox.showerror(
                    "Error de Impresión",
                    "No hay asociación para imprimir PDFs (error 31).\n"
                    "Soluciones: instale/establezca como predeterminado un lector PDF con soporte de impresión (Adobe Reader, SumatraPDF, etc.)\n"
                    "o configure la asociación de impresión para archivos PDF.",
                )
            else:
                messagebox.showerror("Error de Impresión", f"Fallo al enviar a la impresora: {last_err}")
            return

        messagebox.showinfo(
            "Impresión Enviada",
            f"Se han enviado {copies} copias del vale a la impresora predeterminada.",
        )
    except Exception as e:
        messagebox.showerror("Error de Impresión", f"Fallo al enviar a la impresora: {e}")


