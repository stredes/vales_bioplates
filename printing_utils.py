#!/usr/bin/env python
# -*- coding: utf-8 -*-

import logging
import os
import subprocess
import threading
import time
import sys
from typing import Optional
from tkinter import messagebox

import config as app_config
import settings_store as settings

logger = logging.getLogger(__name__)


def _resolve_printer_name(name: Optional[str]) -> Optional[str]:
    if not name:
        return None
    try:
        import win32print  # type: ignore

        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        printers = [p[2] for p in win32print.EnumPrinters(flags)]
        for p in printers:
            if p == name:
                return p
        for p in printers:
            if p.lower() == name.lower():
                return p
        return None
    except Exception:
        return name


def _find_sumatra_exe() -> Optional[str]:
    """Return path to SumatraPDF.exe if present via config/env or typical paths."""
    try:
        sp = getattr(app_config, 'SUMATRA_PDF_PATH', None)
    except Exception:
        sp = None
    if sp and os.path.exists(sp):
        return sp
    candidates = []
    try:
        base_dirs = []
        if getattr(sys, "frozen", False):
            base_dirs.append(os.path.dirname(sys.executable))
            meipass = getattr(sys, "_MEIPASS", None)
            if meipass:
                base_dirs.append(meipass)
        else:
            base_dirs.append(os.path.dirname(os.path.abspath(__file__)))
        for base in base_dirs:
            candidates += [
                os.path.join(base, "SumatraPDF.exe"),
                os.path.join(base, "tools", "SumatraPDF.exe"),
                os.path.join(base, "bin", "SumatraPDF.exe"),
            ]
    except Exception:
        pass
    candidates += [
        os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "SumatraPDF", "SumatraPDF.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "SumatraPDF", "SumatraPDF.exe"),
    ]
    for p in candidates:
        if p and os.path.exists(p):
            return p
    return None


def _close_preview_window(filename: str, delay_sec: float = 2.0) -> None:
    def _worker() -> None:
        time.sleep(max(0.0, delay_sec))
        try:
            import win32con  # type: ignore
            import win32gui  # type: ignore
        except Exception:
            return
        base = os.path.basename(filename)
        if not base:
            return

        def _enum(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return
                title = win32gui.GetWindowText(hwnd)
                if base in title:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            except Exception:
                pass

        try:
            win32gui.EnumWindows(_enum, None)
        except Exception:
            pass

    threading.Thread(target=_worker, daemon=True).start()


def print_pdf_windows(filename: str, copies: int = 1, preview: bool = False) -> None:
    """Imprime un PDF en Windows con varias rutas de respaldo.

    Intentos:
      1) ShellExecute 'print' (requiere asociaci?n de impresi?n para .pdf)
      2) ShellExecute 'printto' con impresora predeterminada
      3) SumatraPDF en modo silencioso si est? instalado
    """
    if not getattr(app_config, 'WINDOWS_OS', False):
        messagebox.showerror(
            "Error de Impresi?n",
            "La impresi?n autom?tica requiere Windows y el m?dulo pywin32.",
        )
        logger.warning("Intento de impresi?n en OS no soportado")
        return

    if not os.path.exists(filename):
        messagebox.showerror("Error de Impresi?n", "El archivo PDF no existe.")
        logger.error("Archivo para impresi?n no existe: %s", filename)
        return

    try:
        import win32api  # type: ignore
        import win32print  # type: ignore

        # Impresora configurada (opcional)
        try:
            raw_printer = settings.get_printer_name()
        except Exception:
            raw_printer = None
        configured_printer = _resolve_printer_name(raw_printer)
        if raw_printer and not configured_printer:
            logger.warning("Impresora configurada no encontrada: %s", raw_printer)

        printer_label = configured_printer or "impresora predeterminada"

        # Intento prioritario: SumatraPDF si est? disponible (evita errores de asociaci?n)
        runner_pref = None
        try:
            runner_pref = _find_sumatra_exe()
        except Exception:
            runner_pref = None
        if runner_pref:
            try:
                if preview:
                    if configured_printer:
                        cmd = [runner_pref, "-print-to", configured_printer, "-exit-on-print", filename]
                    else:
                        cmd = [runner_pref, "-print-to-default", "-exit-on-print", filename]
                    if int(copies) > 1:
                        cmd += ["-print-settings", f"copies={int(copies)}"]
                    subprocess.Popen(cmd)
                    messagebox.showinfo(
                        "Impresi?n Enviada",
                        f"Se han enviado {copies} copias del vale a la {printer_label}.",
                    )
                    logger.info("PDF enviado a SumatraPDF con preview copias=%s", copies)
                    return
                for _ in range(max(1, int(copies))):
                    if configured_printer:
                        cmd = [runner_pref, "-silent", "-print-to", configured_printer, filename]
                    else:
                        cmd = [runner_pref, "-silent", "-print-to-default", filename]
                    subprocess.run(cmd, check=True, creationflags=0x08000000)
                messagebox.showinfo(
                    "Impresi?n Enviada",
                    f"Se han enviado {copies} copias del vale a la {printer_label}.",
                )
                logger.info("PDF enviado a SumatraPDF (default printer) copias=%s", copies)
                return
            except Exception:
                # seguir? con ShellExecute
                logger.warning("SumatraPDF no pudo imprimir, se intentara ShellExecute", exc_info=True)
                pass

        # Impresora predeterminada / configurada
        if configured_printer:
            default_printer = configured_printer
        else:
            try:
                default_printer = win32print.GetDefaultPrinter()
            except Exception:
                default_printer = None

        if not default_printer:
            messagebox.showerror(
                "Error de Impresi?n",
                "No hay impresora predeterminada configurada en Windows. Configure una e intente nuevamente.",
            )
            logger.error("No hay impresora predeterminada configurada en Windows")
            return

        opened_preview = False
        if preview:
            try:
                os.startfile(filename)
                opened_preview = True
            except Exception:
                opened_preview = False

        last_err = None
        for _ in range(max(1, int(copies))):
            try:
                if configured_printer:
                    win32api.ShellExecute(0, "printto", filename, f'"{default_printer}"', ".", 0)
                else:
                    # 1) Verbo 'print'
                    win32api.ShellExecute(0, "print", filename, None, ".", 0)
                last_err = None
            except Exception as e1:
                last_err = e1
                if not configured_printer:
                    # 2) Verbo 'printto' con nombre de impresora
                    try:
                        win32api.ShellExecute(0, "printto", filename, f'"{default_printer}"', ".", 0)
                        last_err = None
                    except Exception as e2:
                        last_err = e2

            if last_err:
                break

        if last_err:
            # 3) SumatraPDF (si existe o se defini? v?a config)
            runner = _find_sumatra_exe()
            if runner:
                try:
                    cmd = [runner, "-silent", "-print-to", default_printer, filename]
                    subprocess.run(cmd, check=True, creationflags=0x08000000)
                    last_err = None
                    logger.info("PDF enviado via SumatraPDF a impresora %s", default_printer)
                except Exception as e3:
                    last_err = e3
                    logger.error("SumatraPDF fallo al imprimir", exc_info=True)
            else:
                logger.warning("SumatraPDF no disponible para fallback de impresi?n")

        if last_err:
            msg = str(last_err)
            if "31" in msg:  # SE_ERR_NOASSOC
                messagebox.showerror(
                    "Error de Impresion",
                    "No hay asociacion para imprimir PDFs (error 31).\n"
                    "Soluciones: instale/establezca como predeterminado un lector PDF con soporte de impresion (Adobe Reader, SumatraPDF, etc.)\n"
                    "o configure la asociacion de impresion para archivos PDF.\n"
                    "Alternativa: coloque SumatraPDF.exe junto al ejecutable o defina SUMATRA_PDF_PATH.",
                )
            else:
                messagebox.showerror("Error de Impresi?n", f"Fallo al enviar a la impresora: {last_err}")
            return

        if opened_preview:
            _close_preview_window(filename, delay_sec=2.0)
        messagebox.showinfo(
            "Impresi?n Enviada",
            f"Se han enviado {copies} copias del vale a la {printer_label}.",
        )
        logger.info("PDF enviado a impresora predeterminada via ShellExecute copias=%s", copies)
    except Exception as e:
        messagebox.showerror("Error de Impresi?n", f"Fallo al enviar a la impresora: {e}")
        logger.exception("Fallo inesperado al enviar a la impresora")
