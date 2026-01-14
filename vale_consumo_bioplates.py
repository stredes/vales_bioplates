#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Interfaz de Solicitud de Productos (Uso Bodega - Bioplates) en ASCII seguro (sin emojis) para
evitar texto corrupto en Windows. Incluye:
- Filtros a la derecha con scroll
- Solicitud en curso con acciones
- Historial con abrir, reimprimir y unificar varios PDFs
- Gestión de solicitantes y usuarios de bodega
- run_app() como punto de entrada
"""

from __future__ import annotations

import logging
import queue
import threading
import os
import subprocess
import sys
from datetime import datetime, timedelta
from typing import Optional, Callable

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import pandas as pd

from config import AREA_FILTER, HISTORY_DIR, INVENTORY_FILE, WINDOWS_OS
from vale_manager import ValeManager
from filters import FilterOptions
from printing_utils import print_pdf_windows
import settings_store as settings
from vale_registry import ValeRegistry
from user_manager import UserManager

MSG_SELECT_HISTORY = "Seleccione una solicitud del listado."
MSG_SELECT_PRODUCT = "Seleccione un producto de la tabla."


class ValeConsumoApp:
    def __init__(self, master: tk.Tk) -> None:
        self.master = master
        self.master.title(f"Solicitud Productos(Uso Bodega)")
        try:
            self.master.state('zoomed')
        except Exception:
            self.master.geometry('1280x800')

        self.log = logging.getLogger("vale_consumo_bioplates")
        self._apply_auto_scaling()
        self._load_queue = queue.Queue()
        self._load_progress = None
        self._load_progress_bar = None
        self._load_progress_label = None
        self._loading_inventory = False
        self.manager = ValeManager()
        # Carpeta de historial desde ajustes (fallback a config)
        try:
            self.history_dir = settings.get_history_dir() or HISTORY_DIR
        except Exception:
            self.history_dir = HISTORY_DIR
        self.registry = ValeRegistry(self.history_dir)
        self.user_manager = UserManager(self.history_dir)
        self.filtered_df: pd.DataFrame = pd.DataFrame()
        self.current_file: Optional[str] = None
        self.show_ubicacion = tk.BooleanVar(value=True)  # Control para mostrar/ocultar ubicaciones
        self.ubicacion_exclude_vars = {}
        self.ubicacion_exclude_search_var = None
        self.ubicacion_checklist_frame = None
        self.ubicacion_checklist_canvas = None
        self.ubicacion_checklist_scroll = None
        self.ubicacion_checklist_inner = None
        self.ubicacion_checklist_window = None
        self.ubicacion_checklist_visible = tk.BooleanVar(value=False)
        self.ubicacion_toggle_btn = None
        self.ubicacion_search_entry = None

        style = ttk.Style(self.master)
        try:
            # Tema y escalado suave en pantallas densas
            try:
                style.theme_use('clam')
            except Exception:
                pass

            # TipografÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â­as y colores base
            base_font = ('Segoe UI', 11)
            style.configure('TLabel', font=base_font)
            style.configure('TLabelframe.Label', font=('Segoe UI', 10, 'bold'))
            style.configure('TButton', font=base_font, padding=(6, 4))
            style.configure('TEntry', font=base_font)
            style.configure('TCombobox', font=base_font)
            style.configure('Treeview', font=base_font, rowheight=28, background='#ffffff', fieldbackground='#ffffff')
            style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'), padding=(8, 6), background='#f0f0f0')
            style.map('Treeview', background=[('selected', '#e1ecff')], foreground=[('selected', '#000000')])

            # BotÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n de acciÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n acentuado
            style.configure('Accent.TButton', background='#4a90e2', foreground='#ffffff')
            style.map('Accent.TButton', background=[('active', '#3d7fcc'), ('pressed', '#346dac')])
            style.configure('Small.TLabel', font=('Segoe UI', 10))
            style.configure('SmallBold.TLabel', font=('Segoe UI', 10, 'bold'))
            style.configure('Action.TButton', font=('Segoe UI', 10), padding=(6, 3))
            style.configure('ActionAccent.TButton', font=('Segoe UI', 10), padding=(6, 3), background='#4a90e2', foreground='#ffffff')
            style.map('ActionAccent.TButton', background=[('active', '#3d7fcc'), ('pressed', '#346dac')])
        except Exception:
            pass
        
        # Configurar umbrales de vencimiento (dias)
        self.days_vencimiento_rojo_min = 60
        self.days_vencimiento_rojo_max = 90

        self._build_menu()

        container = ttk.Frame(self.master)
        container.grid(row=0, column=0, sticky='nsew')
        self.master.rowconfigure(0, weight=1)
        self.master.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=0)  # topbar
        container.rowconfigure(1, weight=3)  # productos
        container.rowconfigure(2, weight=4)  # vale/historial
        container.columnconfigure(0, weight=1)

        # Barra superior: seleccion de archivo + recordatorio
        self.topbar = ttk.Frame(container, padding=(8, 6))
        self.topbar.grid(row=0, column=0, sticky='ew')
        self.topbar.columnconfigure(1, weight=1)

        self.select_btn = ttk.Button(self.topbar, text="Seleccionar archivo de inventario...", command=self.select_inventory_file)
        self.select_btn.grid(row=0, column=0, sticky='w', padx=(0, 10))

        self.file_label = ttk.Label(self.topbar, text="(ningun archivo cargado)")
        self.file_label.grid(row=0, column=1, sticky='w')

        # Acceso rapido a instrucciones
        try:
            self.topbar.columnconfigure(2, weight=0)
            ttk.Button(self.topbar, text="Instrucciones", command=self._open_instructions).grid(row=0, column=2, sticky='e')
        except Exception:
            pass

        if settings.get_reminder_enabled():
            lbl_txt = settings.get_reminder_text()
            self.reminder_label = ttk.Label(self.topbar, text=lbl_txt, foreground="#666666")
            self.reminder_label.grid(row=1, column=0, columnspan=2, sticky='w', pady=(4, 0))
        else:
            self.reminder_label = None

        # Area superior: productos + filtros
        self._build_products_area(container)
        # Area inferior: Notebook Vale / Historial
        self._build_vale_and_history(container)

        # Atajos de teclado globales
        try:
            self.master.bind('<Control-f>', lambda e: self.search_entry.focus_set())
            self.master.bind('<Control-h>', lambda e: self.vale_notebook.select(self.hist_tab))
            self.master.bind('<Control-m>', lambda e: self.vale_notebook.select(self.mgr_tab))
            self.master.bind('<Control-g>', lambda e: self.generate_and_print_vale())
            # Suprimir item seleccionado del vale con tecla Supr
            try:
                self.vale_tree.bind('<Delete>', lambda e: self.remove_from_vale())
                self.vale_tree.bind('<Double-1>', lambda e: self.edit_vale_item())
            except Exception:
                pass
        except Exception:
            pass

        self._restore_last_inventory()

    def _build_menu(self) -> None:
        menubar = tk.Menu(self.master)
        self.master.config(menu=menubar)

        m_archivo = tk.Menu(menubar, tearoff=0)
        m_archivo.add_command(label="Seleccionar inventario...", command=self.select_inventory_file)
        m_archivo.add_separator()
        m_archivo.add_command(label="Salir", command=self.master.destroy)
        menubar.add_cascade(label="Archivo", menu=m_archivo)

        m_conf = tk.Menu(menubar, tearoff=0)
        m_conf.add_command(label="Ajustes...", command=self._open_settings_dialog)
        m_conf.add_separator()
        m_conf.add_command(label="Gestionar Solicitantes...", command=self._open_solicitantes_dialog)
        m_conf.add_command(label="Gestionar Usuarios Bodega...", command=self._open_usuarios_dialog)
        m_conf.add_separator()
        m_conf.add_command(label="Configuracion de impresora...", command=self._menu_printer_settings)
        m_conf.add_command(label="Herramientas de lectura de label...", command=self._menu_label_tools)
        menubar.add_cascade(label="Configuracion", menu=m_conf)

        # Menu Herramientas
        m_tools = tk.Menu(menubar, tearoff=0)
        m_tools.add_command(label="Limpiar base de datos...", command=self._clean_database)
        menubar.add_cascade(label="Herramientas", menu=m_tools)

        # Menu Ayuda
        m_help = tk.Menu(menubar, tearoff=0)
        m_help.add_command(label="Instrucciones de uso...", command=self._open_instructions)
        m_help.add_command(label="Atajos de teclado...", command=self._open_shortcuts)
        menubar.add_cascade(label="Ayuda", menu=m_help)

    def _apply_auto_scaling(self) -> None:
        try:
            screen_w = float(self.master.winfo_screenwidth())
            screen_h = float(self.master.winfo_screenheight())
            base_w, base_h = 1280.0, 800.0
            scale_by_dim = min(screen_w / base_w, screen_h / base_h)
            pixels_per_inch = float(self.master.winfo_fpixels('1i'))
            scale_by_dpi = pixels_per_inch / 72.0
            scaling = max(scale_by_dim, scale_by_dpi)
            scaling = max(0.9, min(1.6, scaling))
            self.master.tk.call('tk', 'scaling', scaling)
            self.log.info(
                "Auto scaling aplicado: %.2f (dpi=%.2f, dim=%.2f)",
                scaling,
                scale_by_dpi,
                scale_by_dim,
            )
        except Exception:
            pass

    def _restore_last_inventory(self) -> None:
        try:
            path = settings.get_last_inventory_file()
        except Exception:
            path = None
        if not path or not os.path.exists(path):
            return
        self.current_file = path
        try:
            base_dir = os.path.dirname(path)
            if base_dir:
                settings.set_last_inventory_dir(base_dir)
        except Exception:
            pass
        self._load_inventory(path)

    def _menu_printer_settings(self) -> None:
        messagebox.showinfo(
            "Impresora",
            "Impresion automatica usa SumatraPDF si esta disponible.\n"
            "Configure SUMATRA_PDF_PATH en config.py o settings_store si desea forzar ruta."
        )

    def _menu_label_tools(self) -> None:
        messagebox.showinfo(
            "Lector de Label",
            "Proximamente: utilidades para probar y configurar el lector de etiquetas."
        )

    def _clean_database(self) -> None:
        """Limpia la base de datos de solicitudes y elimina todos los archivos del historial."""
        confirm = messagebox.askyesno(
            "Limpiar Base de Datos",
            "¿Está seguro que desea limpiar la base de datos?\n\n"
            "Esto eliminará:\n"
            "• Todos los registros de solicitudes del sistema\n"
            "• TODOS los archivos PDF y JSON del historial\n\n"
            "La numeración de solicitudes se reiniciará desde 000.",
            icon='warning'
        )
        if not confirm:
            return
        
        # Segunda confirmación para estar seguros
        confirm2 = messagebox.askyesno(
            "Confirmar limpieza",
            "¿Está COMPLETAMENTE SEGURO?\n\n"
            "Se eliminarán TODOS los archivos del historial.\n"
            "Esta acción NO se puede deshacer.",
            icon='warning'
        )
        if not confirm2:
            return
        
        try:
            deleted_files = 0
            errors = []
            
            # Eliminar archivos PDF y JSON del historial
            if os.path.exists(self.history_dir):
                for filename in os.listdir(self.history_dir):
                    filepath = os.path.join(self.history_dir, filename)
                    # Eliminar solo PDFs y JSONs, no el vales_index.json aún
                    if filename.lower().endswith(('.pdf', '.json')) and filename != 'vales_index.json':
                        try:
                            os.remove(filepath)
                            deleted_files += 1
                        except Exception as e:
                            errors.append(f"{filename}: {e}")
            
            # Limpiar el registro
            self.registry.data = {
                'sequence': 0,
                'vales': []
            }
            self.registry._save()
            
            # Refrescar vistas
            self.refresh_history()
            if hasattr(self, 'refresh_manager'):
                self.refresh_manager()
            
            msg = f"Base de datos limpiada exitosamente.\n\n"
            msg += f"• Archivos eliminados: {deleted_files}\n"
            msg += f"• La numeración comenzará desde 000"
            
            if errors:
                msg += f"\n\nAdvertencias ({len(errors)} archivos no pudieron eliminarse):\n"
                msg += "\n".join(errors[:5])  # Mostrar solo los primeros 5 errores
                if len(errors) > 5:
                    msg += f"\n... y {len(errors) - 5} más"
            
            messagebox.showinfo("Base de Datos Limpiada", msg)
            
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"No se pudo limpiar la base de datos:\n{e}"
            )

    def _open_settings_dialog(self) -> None:
        dlg = tk.Toplevel(self.master)
        dlg.title('Ajustes')
        dlg.transient(self.master)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=12)
        frm.grid(row=0, column=0, sticky='nsew')
        dlg.columnconfigure(0, weight=1)
        dlg.rowconfigure(0, weight=1)

        # Vars
        ap_var = tk.BooleanVar(value=settings.get_auto_print())
        try:
            from config import SUMATRA_PDF_PATH as CFG_SUM
        except Exception:
            CFG_SUM = ''
        sum_var = tk.StringVar(value=(settings.get_sumatra_path() or CFG_SUM or ''))
        rem_en_var = tk.BooleanVar(value=settings.get_reminder_enabled())
        rem_text_var = tk.StringVar(value=settings.get_reminder_text())
        hist_var = tk.StringVar(value=(settings.get_history_dir()))

        r = 0
        ttk.Checkbutton(frm, text='Impresion automatica al generar', variable=ap_var).grid(row=r, column=0, columnspan=3, sticky='w')
        r += 1

        ttk.Label(frm, text='Ruta SumatraPDF.exe (opcional):').grid(row=r, column=0, sticky='w', pady=(8,0))
        e_sum = ttk.Entry(frm, textvariable=sum_var, width=60)
        e_sum.grid(row=r, column=1, sticky='ew', padx=(8,6), pady=(8,0))
        ttk.Button(frm, text='Examinar...', command=lambda: sum_var.set(filedialog.askopenfilename(title='Seleccionar SumatraPDF.exe', filetypes=[('Ejecutable','*.exe')]) or sum_var.get())).grid(row=r, column=2, sticky='w', pady=(8,0))
        r += 1

        ttk.Checkbutton(frm, text='Mostrar recordatorio superior', variable=rem_en_var).grid(row=r, column=0, columnspan=3, sticky='w', pady=(8,0))
        r += 1
        ttk.Label(frm, text='Texto del recordatorio:').grid(row=r, column=0, sticky='w')
        e_rem = ttk.Entry(frm, textvariable=rem_text_var, width=60)
        e_rem.grid(row=r, column=1, columnspan=2, sticky='ew', padx=(8,0))
        r += 1

        ttk.Label(frm, text='Carpeta Historial de vales:').grid(row=r, column=0, sticky='w', pady=(8,0))
        e_hist = ttk.Entry(frm, textvariable=hist_var, width=60)
        e_hist.grid(row=r, column=1, sticky='ew', padx=(8,6), pady=(8,0))
        ttk.Button(frm, text='Seleccionar...', command=lambda: hist_var.set(filedialog.askdirectory(title='Seleccionar carpeta de historial') or hist_var.get())).grid(row=r, column=2, sticky='w', pady=(8,0))
        r += 1

        for c in range(0,3):
            frm.columnconfigure(c, weight=(1 if c==1 else 0))

        btns = ttk.Frame(frm)
        btns.grid(row=r, column=0, columnspan=3, sticky='e', pady=(12,0))
        def _on_save():
            try:
                # Persistir
                settings.set_auto_print(bool(ap_var.get()))
                settings.set_sumatra_path(sum_var.get().strip())
                settings.set_reminder_enabled(bool(rem_en_var.get()))
                settings.set_reminder_text(rem_text_var.get())
                new_hist = hist_var.get().strip()
                if new_hist:
                    settings.set_history_dir(new_hist)
                # Aplicar en runtime
                try:
                    import config as _cfg
                    _cfg.SUMATRA_PDF_PATH = sum_var.get().strip()
                except Exception:
                    pass
                # Recordatorio en UI
                try:
                    if rem_en_var.get():
                        if not self.reminder_label:
                            self.reminder_label = ttk.Label(self.topbar, text=rem_text_var.get(), foreground="#666666")
                            self.reminder_label.grid(row=1, column=0, columnspan=2, sticky='w', pady=(4, 0))
                        else:
                            self.reminder_label.configure(text=rem_text_var.get())
                    else:
                        if self.reminder_label:
                            self.reminder_label.destroy()
                            self.reminder_label = None
                except Exception:
                    pass
                # Historial (reubicar si cambio)
                try:
                    new_dir = settings.get_history_dir()
                    if new_dir and new_dir != getattr(self, 'history_dir', None):
                        self.history_dir = new_dir
                        os.makedirs(self.history_dir, exist_ok=True)
                        self.registry = ValeRegistry(self.history_dir)
                    self.refresh_history()
                    try:
                        self.refresh_manager()
                    except Exception:
                        pass
                except Exception:
                    pass
            finally:
                dlg.destroy()

        ttk.Button(btns, text='Guardar', command=_on_save).pack(side='right')
        ttk.Button(btns, text='Cancelar', command=dlg.destroy).pack(side='right', padx=(0,8))

    def _open_solicitantes_dialog(self) -> None:
        """Abre diálogo para gestionar solicitantes."""
        dlg = tk.Toplevel(self.master)
        dlg.title('Gestionar Solicitantes')
        dlg.geometry('400x500')
        dlg.transient(self.master)
        dlg.grab_set()
        
        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill='both', expand=True)
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(1, weight=1)
        
        ttk.Label(frm, text='Solicitantes registrados:', font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, sticky='w', pady=(0,6))
        
        listbox_frame = ttk.Frame(frm)
        listbox_frame.grid(row=1, column=0, sticky='nsew', pady=(0,10))
        listbox_frame.columnconfigure(0, weight=1)
        listbox_frame.rowconfigure(0, weight=1)
        
        listbox = tk.Listbox(listbox_frame, height=15)
        listbox.grid(row=0, column=0, sticky='nsew')
        scroll = ttk.Scrollbar(listbox_frame, orient='vertical', command=listbox.yview)
        scroll.grid(row=0, column=1, sticky='ns')
        listbox.configure(yscrollcommand=scroll.set)
        
        def refresh_list():
            listbox.delete(0, tk.END)
            for s in self.user_manager.get_solicitantes():
                listbox.insert(tk.END, s)
        
        refresh_list()
        
        input_frame = ttk.Frame(frm)
        input_frame.grid(row=2, column=0, sticky='ew', pady=(0,10))
        input_frame.columnconfigure(1, weight=1)
        
        ttk.Label(input_frame, text='Nombre:').grid(row=0, column=0, sticky='w', padx=(0,6))
        nombre_var = tk.StringVar()
        nombre_entry = ttk.Entry(input_frame, textvariable=nombre_var)
        nombre_entry.grid(row=0, column=1, sticky='ew')
        
        def add_solicitante():
            nombre = nombre_var.get().strip()
            if not nombre:
                messagebox.showwarning('Agregar Solicitante', 'Ingrese un nombre.')
                return
            try:
                if self.user_manager.add_solicitante(nombre):
                    refresh_list()
                    nombre_var.set('')
                    self._refresh_solicitantes_combo()  # Sincronizar combobox
                else:
                    messagebox.showinfo('Agregar Solicitante', 'Este solicitante ya existe.')
            except Exception as e:
                messagebox.showerror('Error', str(e))
        
        def remove_solicitante():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning('Eliminar', 'Seleccione un solicitante.')
                return
            nombre = listbox.get(sel[0])
            if messagebox.askyesno('Confirmar', f'¿Eliminar a "{nombre}"?'):
                self.user_manager.remove_solicitante(nombre)
                refresh_list()
                self._refresh_solicitantes_combo()  # Sincronizar combobox
        
        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=3, column=0, sticky='ew')
        ttk.Button(btn_frame, text='Agregar', command=add_solicitante).pack(side='left', padx=(0,6))
        ttk.Button(btn_frame, text='Eliminar Seleccionado', command=remove_solicitante).pack(side='left')
        ttk.Button(btn_frame, text='Cerrar', command=dlg.destroy).pack(side='right')
    
    def _open_usuarios_dialog(self) -> None:
        """Abre diálogo para gestionar usuarios de bodega."""
        dlg = tk.Toplevel(self.master)
        dlg.title('Gestionar Usuarios Bodega')
        dlg.geometry('400x500')
        dlg.transient(self.master)
        dlg.grab_set()
        
        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill='both', expand=True)
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(1, weight=1)
        
        ttk.Label(frm, text='Usuarios de bodega registrados:', font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, sticky='w', pady=(0,6))
        
        listbox_frame = ttk.Frame(frm)
        listbox_frame.grid(row=1, column=0, sticky='nsew', pady=(0,10))
        listbox_frame.columnconfigure(0, weight=1)
        listbox_frame.rowconfigure(0, weight=1)
        
        listbox = tk.Listbox(listbox_frame, height=15)
        listbox.grid(row=0, column=0, sticky='nsew')
        scroll = ttk.Scrollbar(listbox_frame, orient='vertical', command=listbox.yview)
        scroll.grid(row=0, column=1, sticky='ns')
        listbox.configure(yscrollcommand=scroll.set)
        
        def refresh_list():
            listbox.delete(0, tk.END)
            for u in self.user_manager.get_usuarios_bodega():
                listbox.insert(tk.END, u)
        
        refresh_list()
        
        input_frame = ttk.Frame(frm)
        input_frame.grid(row=2, column=0, sticky='ew', pady=(0,10))
        input_frame.columnconfigure(1, weight=1)
        
        ttk.Label(input_frame, text='Nombre:').grid(row=0, column=0, sticky='w', padx=(0,6))
        nombre_var = tk.StringVar()
        nombre_entry = ttk.Entry(input_frame, textvariable=nombre_var)
        nombre_entry.grid(row=0, column=1, sticky='ew')
        
        def add_usuario():
            nombre = nombre_var.get().strip()
            if not nombre:
                messagebox.showwarning('Agregar Usuario', 'Ingrese un nombre.')
                return
            try:
                if self.user_manager.add_usuario_bodega(nombre):
                    refresh_list()
                    nombre_var.set('')
                    self._refresh_usuarios_combo()  # Sincronizar combobox
                else:
                    messagebox.showinfo('Agregar Usuario', 'Este usuario ya existe.')
            except Exception as e:
                messagebox.showerror('Error', str(e))
        
        def remove_usuario():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning('Eliminar', 'Seleccione un usuario.')
                return
            nombre = listbox.get(sel[0])
            if messagebox.askyesno('Confirmar', f'¿Eliminar a "{nombre}"?'):
                self.user_manager.remove_usuario_bodega(nombre)
                refresh_list()
                self._refresh_usuarios_combo()  # Sincronizar combobox
        
        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=3, column=0, sticky='ew')
        ttk.Button(btn_frame, text='Agregar', command=add_usuario).pack(side='left', padx=(0,6))
        ttk.Button(btn_frame, text='Eliminar Seleccionado', command=remove_usuario).pack(side='left')
        ttk.Button(btn_frame, text='Cerrar', command=dlg.destroy).pack(side='right')

    # --- Ayuda / Instrucciones ---
    def _open_instructions(self) -> None:
        text = None
        # Buscar instrucciones.txt en carpeta del script o cwd; fallback a texto embebido
        try:
            here = os.path.dirname(os.path.abspath(__file__))
            candidates = [
                os.path.join(here, 'instrucciones.txt'),
                os.path.join(os.getcwd(), 'instrucciones.txt')
            ]
            for p in candidates:
                if os.path.exists(p):
                    with open(p, 'r', encoding='utf-8') as f:
                        text = f.read()
                        break
        except Exception:
            text = None
        if not text:
            text = (
                "Uso basico:\n\n"
                "1) Seleccione el archivo de inventario (Excel).\n"
                "2) Aplique filtros (producto, subfamilia, lote, ubicacion, fechas).\n"
                "3) Seleccione Solicitante y Usuario de Bodega.\n"
                "4) Ingrese cantidad y pulse 'Agregar a Solicitud'.\n"
                "5) Pulse 'Generar e Imprimir Solicitud' para crear el PDF e imprimir.\n\n"
                "Historial y Manager:\n"
                "- En Historial: abrir (doble clic), reimprimir y unificar varias solicitudes.\n"
                "- Ordene columnas haciendo clic en los encabezados.\n"
                "- En Manager: cambiar estados (Pendiente/Descontado/Anulado) y exportar listados.\n\n"
                "Ajustes:\n"
                "- Gestione solicitantes y usuarios de bodega desde el menu Configuracion.\n"
                "- Configure cantidad de copias a imprimir.\n"
                "- Excluya ubicaciones desde el checklist.\n"
            )
        dlg = tk.Toplevel(self.master)
        dlg.title('Instrucciones de uso')
        dlg.geometry('820x520')
        dlg.transient(self.master)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=10)
        frm.pack(fill='both', expand=True)
        txt = tk.Text(frm, wrap='word')
        ysb = ttk.Scrollbar(frm, orient='vertical', command=txt.yview)
        txt.configure(yscrollcommand=ysb.set)
        txt.pack(side='left', fill='both', expand=True)
        ysb.pack(side='right', fill='y')
        try:
            txt.insert('1.0', text)
        except Exception:
            txt.insert('1.0', 'No se pudieron cargar las instrucciones.')
        txt.configure(state='disabled')

    def _open_shortcuts(self) -> None:
        info = (
            "Atajos de teclado:\n\n"
            "Ctrl+F : Enfocar busqueda de productos\n"
            "Ctrl+H : Ir a pestaña Historial\n"
            "Ctrl+M : Ir a pestaña Manager Solicitudes\n"
            "Ctrl+G : Generar e Imprimir Solicitud\n"
            "Supr   : Eliminar item seleccionado de la solicitud\n"
        )
        messagebox.showinfo('Atajos de teclado', info)

    # -------- Productos y filtros --------
    def _build_products_area(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent, padding=(8, 0))
        frame.grid(row=1, column=0, sticky='nsew')
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        # Tabla a la izquierda
        table_frame = ttk.Frame(frame)
        table_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 8))
        frame.columnconfigure(0, weight=1)

        self._init_product_tree(table_frame)
        self._init_filter_panel(frame)

    def _init_product_tree(self, table_frame: ttk.Frame) -> None:
        self.product_tree = ttk.Treeview(
            table_frame,
            columns=("Producto", "Codigo", "Lote", "Bodega", "Ubicacion", "Vencimiento", "Stock"),
            show='headings',
            selectmode='browse',
        )
        headers = (
            ('Producto', 'Producto', 280, 'w'),
            ('Codigo', 'Codigo', 110, 'center'),
            ('Lote', 'Lote', 120, 'center'),
            ('Bodega', 'Bodega', 120, 'center'),
            ('Ubicacion', 'Ubicacion', 160, 'center'),
            ('Vencimiento', 'Vencimiento', 130, 'center'),
            ('Stock', 'Stock', 80, 'center'),
        )
        for col, text, width, anchor in headers:
            self.product_tree.heading(
                col,
                text=text,
                command=lambda c=col: self._sort_treeview(self.product_tree, c),
            )
            self.product_tree.column(col, width=width, minwidth=max(60, width - 40), anchor=anchor, stretch=True)

        self.product_tree.grid(row=0, column=0, sticky='nsew')
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.product_tree.yview)
        vsb.grid(row=0, column=1, sticky='ns')
        hsb = ttk.Scrollbar(table_frame, orient='horizontal', command=self.product_tree.xview)
        hsb.grid(row=1, column=0, sticky='ew')
        self.product_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.product_tree.bind('<Configure>', self._autosize_product_columns)
        
        # Configurar tags de color para vencimiento
        self.product_tree.tag_configure('vencido', background='#ffb3b3', foreground='#b00000')  # Rojo: vencido
        self.product_tree.tag_configure('vencimiento_proximo', background='#cfe5ff', foreground='#004a99')  # Azul: proximo a vencer

    def _init_filter_panel(self, frame: ttk.Frame) -> None:
        # Panel de control con scroll
        filters_lf = ttk.LabelFrame(frame, text="Filtros y Acciones", padding=0, width=260)
        filters_lf.grid(row=0, column=1, sticky='ns')
        try:
            filters_lf.grid_propagate(False)
        except Exception:
            pass

        filters_canvas = tk.Canvas(filters_lf, borderwidth=0, highlightthickness=0)
        filters_vsb = ttk.Scrollbar(filters_lf, orient='vertical', command=filters_canvas.yview)
        filters_canvas.configure(yscrollcommand=filters_vsb.set)
        filters_vsb.pack(side='right', fill='y')
        filters_canvas.pack(side='left', fill='both', expand=True)

        self.control_frame = ttk.Frame(filters_canvas, padding=8)
        self._filters_window = filters_canvas.create_window((0, 0), window=self.control_frame, anchor='nw')

        def _sync_scroll_region(_: 'tk.Event') -> None:
            try:
                filters_canvas.configure(scrollregion=filters_canvas.bbox('all'))
            except Exception:
                pass

        def _fit_canvas_width(event: 'tk.Event') -> None:
            try:
                filters_canvas.itemconfigure(self._filters_window, width=event.width)
            except Exception:
                pass

        self.control_frame.bind('<Configure>', _sync_scroll_region)
        filters_canvas.bind('<Configure>', _fit_canvas_width)
        filters_lf.bind('<Enter>', lambda _: self._bind_filter_mousewheel(filters_canvas))
        filters_lf.bind('<Leave>', lambda _: filters_canvas.unbind_all('<MouseWheel>'))

        self._build_filter_controls()

    def _bind_filter_mousewheel(self, canvas: tk.Canvas) -> None:
        def _on_mousewheel(event: 'tk.Event') -> None:
            try:
                delta = 0
                if getattr(event, 'delta', None):
                    delta = int(-1 * (event.delta / 120))
                elif getattr(event, 'num', None) == 4:
                    delta = -1
                elif getattr(event, 'num', None) == 5:
                    delta = 1
                if delta:
                    canvas.yview_scroll(delta, 'units')
            except Exception:
                pass

        canvas.bind_all('<MouseWheel>', _on_mousewheel)

    def _build_filter_controls(self) -> None:
        # Buscar
        ttk.Label(self.control_frame, text="Buscar producto:", font=('Segoe UI', 10)).grid(row=0, column=0, sticky='w', pady=(0, 4))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(self.control_frame, textvariable=self.search_var, width=26, font=('Segoe UI', 10))  # Agregada fuente
        self.search_entry.grid(row=1, column=0, sticky='ew', pady=(0, 8))
        self.search_var.trace_add('write', lambda *_: self.filter_products())

        # Cantidad y acciones
        ttk.Label(self.control_frame, text="Cantidad a retirar:", font=('Segoe UI', 10)).grid(row=2, column=0, sticky='w')
        self.quantity_entry = ttk.Entry(self.control_frame, width=12, font=('Segoe UI', 10))  # Aumentado width y agregada fuente
        self.quantity_entry.insert(0, '1')
        self.quantity_entry.grid(row=3, column=0, sticky='w', pady=(0, 8))
        ttk.Button(self.control_frame, text="Agregar a Solicitud", style='ActionAccent.TButton', command=self.add_to_vale, width=26).grid(row=4, column=0, sticky='ew', pady=(6, 5))
        ttk.Button(self.control_frame, text="Generar e Imprimir Solicitud", style='ActionAccent.TButton', command=self.generate_and_print_vale, width=26).grid(row=5, column=0, sticky='ew', pady=(2, 10))

        # Subfamilia
        ttk.Label(self.control_frame, text="Subfamilia:", style='Small.TLabel').grid(row=6, column=0, sticky='w', pady=(4,0))
        self.subfam_var = tk.StringVar(value='(Todas)')
        self.subfam_combo = ttk.Combobox(self.control_frame, textvariable=self.subfam_var, state='readonly', width=24, font=('Segoe UI', 10))  # Agregada fuente
        self.subfam_combo.grid(row=7, column=0, sticky='ew', pady=(0, 8))
        self.subfam_combo.bind('<<ComboboxSelected>>', lambda *_: self.filter_products())

        # Lote / Ubicacion
        ttk.Label(self.control_frame, text="Lote:", style='Small.TLabel').grid(row=8, column=0, sticky='w')
        self.lote_var = tk.StringVar()
        ttk.Entry(self.control_frame, textvariable=self.lote_var, width=26, font=('Segoe UI', 10)).grid(row=9, column=0, sticky='ew', pady=(0, 8))

        ttk.Label(self.control_frame, text="Ubicacion:", style='Small.TLabel').grid(row=10, column=0, sticky='w')
        self.ubi_var = tk.StringVar()
        ttk.Entry(self.control_frame, textvariable=self.ubi_var, width=26, font=('Segoe UI', 10)).grid(row=11, column=0, sticky='ew', pady=(0, 8))

        # Rango de vencimiento
        ttk.Label(self.control_frame, text="Vencimiento desde (YYYY-MM-DD):", style='Small.TLabel').grid(row=12, column=0, sticky='w')
        self.vdesde_var = tk.StringVar()
        ttk.Entry(self.control_frame, textvariable=self.vdesde_var, width=26, font=('Segoe UI', 10)).grid(row=13, column=0, sticky='ew', pady=(0, 8))
        ttk.Label(self.control_frame, text="Vencimiento hasta (YYYY-MM-DD):", style='Small.TLabel').grid(row=14, column=0, sticky='w')
        self.vhasta_var = tk.StringVar()
        ttk.Entry(self.control_frame, textvariable=self.vhasta_var, width=26, font=('Segoe UI', 10)).grid(row=15, column=0, sticky='ew', pady=(0, 8))

        self.stock_only_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.control_frame, text='Solo con stock', variable=self.stock_only_var, command=self.filter_products).grid(row=16, column=0, sticky='w', pady=(6, 8))

        # Excluir ubicaciones (checklist desplegable)
        ubic_head = ttk.Frame(self.control_frame)
        ubic_head.grid(row=17, column=0, sticky='ew')
        ubic_head.columnconfigure(0, weight=1)
        ttk.Label(ubic_head, text="Excluir ubicaciones:", style='Small.TLabel').grid(row=0, column=0, sticky='w')
        self.ubicacion_toggle_btn = ttk.Button(ubic_head, text="Mostrar", command=self._toggle_ubicaciones_checklist, width=10)
        self.ubicacion_toggle_btn.grid(row=0, column=1, sticky='e')

        self.ubicacion_exclude_search_var = tk.StringVar()
        self.ubicacion_exclude_search_var.trace_add('write', lambda *_: self._render_ubicaciones_checklist())
        self.ubicacion_search_entry = ttk.Entry(self.control_frame, textvariable=self.ubicacion_exclude_search_var, width=26, font=('Segoe UI', 10))
        self.ubicacion_search_entry.grid(row=18, column=0, sticky='ew', pady=(0, 6))
        self.ubicacion_checklist_frame = ttk.Frame(self.control_frame)
        self.ubicacion_checklist_frame.grid(row=19, column=0, sticky='ew', pady=(0, 12))
        if not self.ubicacion_checklist_visible.get():
            self.ubicacion_search_entry.grid_remove()
            self.ubicacion_checklist_frame.grid_remove()

        # Separador
        ttk.Separator(self.control_frame, orient='horizontal').grid(row=20, column=0, sticky='ew', pady=(8,12))

        # Solicitante
        ttk.Label(self.control_frame, text="Solicitante:", style='SmallBold.TLabel').grid(row=21, column=0, sticky='w')
        self.solicitante_var = tk.StringVar()
        self.solicitante_combo = ttk.Combobox(self.control_frame, textvariable=self.solicitante_var, width=24, font=('Segoe UI', 10))
        self.solicitante_combo.grid(row=22, column=0, sticky='ew', pady=(0, 8))
        self._refresh_solicitantes_combo()

        # Usuario Bodega
        ttk.Label(self.control_frame, text="Usuario Bodega:", style='SmallBold.TLabel').grid(row=23, column=0, sticky='w')
        self.usuario_bodega_var = tk.StringVar()
        self.usuario_bodega_combo = ttk.Combobox(self.control_frame, textvariable=self.usuario_bodega_var, width=24, font=('Segoe UI', 10))
        self.usuario_bodega_combo.grid(row=24, column=0, sticky='ew', pady=(0, 8))
        self._refresh_usuarios_combo()

        # Cantidad de copias
        ttk.Label(self.control_frame, text="Copias a imprimir:", style='SmallBold.TLabel').grid(row=25, column=0, sticky='w')
        self.copias_var = tk.IntVar(value=1)
        copias_spinbox = ttk.Spinbox(self.control_frame, from_=1, to=10, textvariable=self.copias_var, width=12, font=('Segoe UI', 10))
        copias_spinbox.grid(row=26, column=0, sticky='w', pady=(0, 12))

        # Separador
        ttk.Separator(self.control_frame, orient='horizontal').grid(row=27, column=0, sticky='ew', pady=(8,12))

        # Leyenda de vencimiento
        ttk.Label(self.control_frame, text="Leyenda:", style='SmallBold.TLabel').grid(row=28, column=0, sticky='w')
        
        legend_frame = ttk.Frame(self.control_frame)
        legend_frame.grid(row=29, column=0, sticky='ew', pady=(6,12))
        
        vencido_lbl = tk.Label(legend_frame, text=" Vencido ", bg='#ffb3b3', fg='#b00000', font=('Segoe UI', 9), relief='solid', borderwidth=1, pady=4)
        vencido_lbl.pack(fill='x', pady=3)
        proximo_lbl = tk.Label(legend_frame, text=" Proximo a vencer ", bg='#cfe5ff', fg='#004a99', font=('Segoe UI', 9), relief='solid', borderwidth=1, pady=4)
        proximo_lbl.pack(fill='x', pady=3)
        estable_lbl = tk.Label(legend_frame, text=" Fecha estable ", bg='#ffffff', fg='#333333', font=('Segoe UI', 9), relief='solid', borderwidth=1, pady=4)
        estable_lbl.pack(fill='x', pady=3)
        # Limpiar filtros
        ttk.Button(self.control_frame, text="Limpiar filtros", command=self._clear_filters, width=26).grid(row=30, column=0, sticky='ew', pady=(4, 4))

        for i in range(0, 31):
            self.control_frame.rowconfigure(i, weight=0)
        self.control_frame.columnconfigure(0, weight=1)

    def _autosize_product_columns(self, event=None):
        try:
            tw = self.product_tree
            tw.update_idletasks()
            width = tw.winfo_width()
            if not width or width < 300:
                return
            # Ajustar según si ubicación está visible
            if self.show_ubicacion.get():
                specs = [
                    ('Producto',    0.34, 220),
                    ('Codigo',      0.11,  80),
                    ('Lote',        0.13,  80),
                    ('Bodega',      0.11,  90),
                    ('Ubicacion',   0.14, 120),
                    ('Vencimiento', 0.10, 100),
                    ('Stock',       0.07,  60),
                ]
            else:
                specs = [
                    ('Producto',    0.38, 220),
                    ('Codigo',      0.12,  80),
                    ('Lote',        0.16,  80),
                    ('Bodega',      0.14,  90),
                    ('Vencimiento', 0.12, 100),
                    ('Stock',       0.08,  60),
                ]
            for col, frac, minw in specs:
                if col in tw['columns']:
                    tw.column(col, width=max(int(width * frac), minw), stretch=True)
        except Exception:
            pass

    def _toggle_ubicacion_column(self):
        """Muestra u oculta la columna de Ubicación."""
        try:
            if self.show_ubicacion.get():
                # Mostrar columna
                self.product_tree['displaycolumns'] = ("Producto", "Codigo", "Lote", "Bodega", "Ubicacion", "Vencimiento", "Stock")
                self.vale_tree['displaycolumns'] = ("Producto", "Lote", "Ubicacion", "Vencimiento", "Cantidad")
            else:
                # Ocultar columna
                self.product_tree['displaycolumns'] = ("Producto", "Codigo", "Lote", "Bodega", "Vencimiento", "Stock")
                self.vale_tree['displaycolumns'] = ("Producto", "Lote", "Vencimiento", "Cantidad")
            self._autosize_product_columns()
        except Exception:
            pass

    def _refresh_solicitantes_combo(self):
        """Actualiza la lista de solicitantes en el combobox."""
        try:
            solicitantes = self.user_manager.get_solicitantes()
            self.solicitante_combo['values'] = solicitantes
            if not self.solicitante_var.get() and solicitantes:
                self.solicitante_var.set(solicitantes[0])
        except Exception:
            pass

    def _refresh_usuarios_combo(self):
        """Actualiza la lista de usuarios de bodega en el combobox."""
        try:
            usuarios = self.user_manager.get_usuarios_bodega()
            self.usuario_bodega_combo['values'] = usuarios
            if not self.usuario_bodega_var.get() and usuarios:
                self.usuario_bodega_var.set(usuarios[0])
        except Exception:
            pass

    def _refresh_ubicaciones_checklist(self) -> None:
        try:
            df = self.manager.bioplates_inventory
            if df is None or df.empty or 'Ubicacion' not in df.columns:
                ubicaciones = []
            else:
                ubicaciones = sorted(
                    {str(x).strip() for x in df['Ubicacion'].dropna().astype(str) if str(x).strip()}
                )
        except Exception:
            ubicaciones = []
        existing = self.ubicacion_exclude_vars or {}
        new_vars = {}
        for ubi in ubicaciones:
            var = existing.get(ubi)
            if var is None:
                var = tk.BooleanVar(value=False)
            new_vars[ubi] = var
        self.ubicacion_exclude_vars = new_vars
        self._render_ubicaciones_checklist()

    def _render_ubicaciones_checklist(self) -> None:
        frame = getattr(self, 'ubicacion_checklist_frame', None)
        if not frame:
            return
        if not self.ubicacion_checklist_canvas:
            frame.columnconfigure(0, weight=1)
            self.ubicacion_checklist_canvas = tk.Canvas(frame, height=96, highlightthickness=0)
            self.ubicacion_checklist_scroll = ttk.Scrollbar(
                frame, orient='vertical', command=self.ubicacion_checklist_canvas.yview
            )
            self.ubicacion_checklist_canvas.configure(yscrollcommand=self.ubicacion_checklist_scroll.set)
            self.ubicacion_checklist_inner = ttk.Frame(self.ubicacion_checklist_canvas)
            self.ubicacion_checklist_window = self.ubicacion_checklist_canvas.create_window(
                (0, 0), window=self.ubicacion_checklist_inner, anchor='nw'
            )

            def _on_inner_configure(_event):
                self.ubicacion_checklist_canvas.configure(
                    scrollregion=self.ubicacion_checklist_canvas.bbox("all")
                )

            def _on_canvas_configure(event):
                self.ubicacion_checklist_canvas.itemconfigure(
                    self.ubicacion_checklist_window, width=event.width
                )

            self.ubicacion_checklist_inner.bind("<Configure>", _on_inner_configure)
            self.ubicacion_checklist_canvas.bind("<Configure>", _on_canvas_configure)

            self.ubicacion_checklist_canvas.grid(row=0, column=0, sticky='ew')
            self.ubicacion_checklist_scroll.grid(row=0, column=1, sticky='ns')

        for child in self.ubicacion_checklist_inner.winfo_children():
            child.destroy()
        ubicaciones = sorted(self.ubicacion_exclude_vars.keys())
        term = ''
        try:
            term = (self.ubicacion_exclude_search_var.get() if self.ubicacion_exclude_search_var else '').strip().lower()
        except Exception:
            term = ''
        shown = 0
        for ubi in ubicaciones:
            if term and term not in ubi.lower():
                continue
            var = self.ubicacion_exclude_vars.get(ubi)
            cb = ttk.Checkbutton(self.ubicacion_checklist_inner, text=ubi, variable=var, command=self.filter_products)
            cb.pack(anchor='w')
            shown += 1
        if shown == 0:
            ttk.Label(self.ubicacion_checklist_inner, text="(sin ubicaciones)").pack(anchor='w')

    def _toggle_ubicaciones_checklist(self) -> None:
        try:
            visible = bool(self.ubicacion_checklist_visible.get())
            self.ubicacion_checklist_visible.set(not visible)
            if self.ubicacion_checklist_visible.get():
                if self.ubicacion_search_entry:
                    self.ubicacion_search_entry.grid()
                if self.ubicacion_checklist_frame:
                    self.ubicacion_checklist_frame.grid()
                if self.ubicacion_toggle_btn:
                    self.ubicacion_toggle_btn.configure(text="Ocultar")
            else:
                if self.ubicacion_search_entry:
                    self.ubicacion_search_entry.grid_remove()
                if self.ubicacion_checklist_frame:
                    self.ubicacion_checklist_frame.grid_remove()
                if self.ubicacion_toggle_btn:
                    self.ubicacion_toggle_btn.configure(text="Mostrar")
        except Exception:
            pass

    def _get_excluded_ubicaciones(self):
        try:
            return [ubi for ubi, var in self.ubicacion_exclude_vars.items() if var.get()]
        except Exception:
            return []

    def _clear_ubicaciones_excluidas(self) -> None:
        try:
            for var in self.ubicacion_exclude_vars.values():
                var.set(False)
        except Exception:
            pass
        self._render_ubicaciones_checklist()

    def _clear_filters(self) -> None:
        self.search_var.set("")
        self.subfam_var.set('(Todas)')
        self.lote_var.set("")
        self.ubi_var.set("")
        self.vdesde_var.set("")
        self.vhasta_var.set("")
        self.stock_only_var.set(False)
        try:
            if self.ubicacion_exclude_search_var is not None:
                self.ubicacion_exclude_search_var.set("")
        except Exception:
            pass
        self._clear_ubicaciones_excluidas()
        self.filter_products()

    # -------- Vale / Historial --------
    def _build_vale_and_history(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent, padding=(8, 6))
        frame.grid(row=2, column=0, sticky='nsew')
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self.vale_notebook = ttk.Notebook(frame)
        self.vale_notebook.grid(row=0, column=0, sticky='nsew')

        # Tab Solicitud
        self.vale_tab = ttk.Frame(self.vale_notebook)
        self.vale_notebook.add(self.vale_tab, text='Solicitud')
        self.vale_tab.columnconfigure(0, weight=1)
        self.vale_tab.rowconfigure(0, weight=1)

        vale_table_frame = ttk.Frame(self.vale_tab)
        vale_table_frame.grid(row=0, column=0, sticky='nsew')
        vale_table_frame.columnconfigure(0, weight=1)
        vale_table_frame.rowconfigure(0, weight=1)

        self.vale_tree = ttk.Treeview(
            vale_table_frame,
            columns=("Producto", "Lote", "Ubicacion", "Vencimiento", "Cantidad"),
            show='headings',
            selectmode='browse',
        )
        for col, text, w in (
            ('Producto', 'Producto', 380),
            ('Lote', 'Lote', 120),
            ('Ubicacion', 'Ubicacion', 160),
            ('Vencimiento', 'Vencimiento', 130),
            ('Cantidad', 'Cantidad', 100),
        ):
            self.vale_tree.heading(
                col,
                text=text,
                command=lambda c=col: self._sort_treeview(self.vale_tree, c),
            )
            anchor = 'center' if col != 'Producto' else 'w'
            self.vale_tree.column(col, width=w, anchor=anchor, stretch=True)

        self.vale_tree.grid(row=0, column=0, sticky='nsew')
        v_vsb = ttk.Scrollbar(vale_table_frame, orient='vertical', command=self.vale_tree.yview)
        v_vsb.grid(row=0, column=1, sticky='ns')
        self.vale_tree.configure(yscrollcommand=v_vsb.set)

        self.vale_actions_frame = ttk.Frame(self.vale_tab, padding=12)
        self.vale_actions_frame.grid(row=0, column=1, sticky='ns', padx=(8, 0))
        ttk.Button(self.vale_actions_frame, text="Editar Cantidad", command=self.edit_vale_item, width=26, style='Action.TButton').pack(pady=8, fill='x')
        ttk.Button(self.vale_actions_frame, text="Eliminar Producto", command=self.remove_from_vale, width=26, style='Action.TButton').pack(pady=8, fill='x')
        ttk.Button(self.vale_actions_frame, text="Generar e Imprimir Solicitud", command=self.generate_and_print_vale, width=26, style='Action.TButton').pack(pady=8, fill='x')
        ttk.Button(self.vale_actions_frame, text="Limpiar Solicitud", command=self.clear_vale, width=26, style='Action.TButton').pack(pady=8, fill='x')

        self._init_history_tab()

    def _init_history_tab(self) -> None:
        self.hist_tab = ttk.Frame(self.vale_notebook)
        self.vale_notebook.add(self.hist_tab, text='Historial')
        self._build_history_ui()
        self.refresh_history()

        # Tab Manager de Solicitudes
        try:
            self.mgr_tab = ttk.Frame(self.vale_notebook)
            self.vale_notebook.add(self.mgr_tab, text='Manager Solicitudes')
            self._build_manager_ui()
            self.refresh_manager()
        except Exception:
            pass

    def _build_history_ui(self) -> None:
        if not os.path.exists(self.history_dir):
            os.makedirs(self.history_dir, exist_ok=True)

        self.history_frame = ttk.Frame(self.hist_tab, padding=12)
        self.history_frame.pack(fill='both', expand=True)
        self.history_frame.columnconfigure(0, weight=1)
        # Fila 1 (árbol) crecerá
        self.history_frame.rowconfigure(1, weight=1)

        # Barra de búsqueda
        top = ttk.Frame(self.history_frame)
        top.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 8))
        ttk.Label(top, text='Buscar:', font=('Segoe UI', 10)).pack(side='left')
        self.hist_search_var = tk.StringVar(value='')
        hist_entry = ttk.Entry(top, textvariable=self.hist_search_var, width=32, font=('Segoe UI', 10))
        hist_entry.pack(side='left', padx=(6, 0))
        self.hist_search_var.trace_add('write', lambda *_: self.refresh_history())

        # Historial basado en el registro: Numero, Estado, Fecha, Archivo, Items
        cols = ("Numero", "Estado", "Fecha", "Archivo", "Items")
        self.history_tree = ttk.Treeview(self.history_frame, columns=cols, show='headings', selectmode='extended')
        
        # Hacer headers clicables para ordenar
        self._tree_sort_state = {}
        
        for col in cols:
            self.history_tree.heading(
                col,
                text=col,
                command=lambda c=col: self._sort_treeview(self.history_tree, c),
            )
        
        self.history_tree.column('Numero', width=80, anchor='center', stretch=False)
        self.history_tree.column('Estado', width=110, anchor='center', stretch=False)
        self.history_tree.column('Fecha', width=170, anchor='center', stretch=False)
        self.history_tree.column('Archivo', width=520, anchor='w', stretch=True)
        self.history_tree.column('Items', width=80, anchor='center', stretch=False)

        self.history_tree.grid(row=1, column=0, sticky='nsew')
        h_vsb = ttk.Scrollbar(self.history_frame, orient='vertical', command=self.history_tree.yview)
        h_vsb.grid(row=1, column=1, sticky='ns')
        self.history_tree.configure(yscrollcommand=h_vsb.set)
        
        # Doble clic para abrir PDF
        self.history_tree.bind('<Double-Button-1>', lambda e: self.open_selected_history())
        
        try:
            self.history_tree.bind('<Configure>', self._autosize_history_columns)
        except Exception:
            pass

        act = ttk.Frame(self.history_frame)
        act.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(10, 0))
        ttk.Button(act, text="Unificar seleccionados", command=self.merge_selected_history).pack(side='left', padx=(0, 8))
        ttk.Button(act, text="Abrir PDF", command=self.open_selected_history).pack(side='left', padx=(0, 8))
        ttk.Button(act, text="Reimprimir", command=self.print_selected_history).pack(side='left')

    def _autosize_history_columns(self, event: Optional['tk.Event'] = None) -> None:
        """Distribuye el ancho de la tabla de historial manteniendo visible la fecha."""
        try:
            tw = self.history_tree
            tw.update_idletasks()
            width = tw.winfo_width()
            if not width:
                return
            fecha_min = 180
            fecha_w = max(fecha_min, int(width * 0.22))
            archivo_w = max(200, width - fecha_w - 20)
            tw.column('Archivo', width=archivo_w, anchor='w', stretch=True)
            tw.column('Fecha', width=fecha_w, anchor='center', stretch=False)
        except Exception:
            pass

    def _ensure_history_dir(self) -> None:
        if not os.path.exists(HISTORY_DIR):
            os.makedirs(HISTORY_DIR, exist_ok=True)
            self.log.debug("Creada carpeta de historial en %s", HISTORY_DIR)

    def refresh_history(self) -> None:
        # Cargar desde el registro; si esta vacio y existen PDFs, reindexar primero
        for i in self.history_tree.get_children():
            self.history_tree.delete(i)
        try:
            rows = self.registry.list()
            if not rows:
                try:
                    if os.path.isdir(self.history_dir) and any(fn.lower().endswith('.pdf') for fn in os.listdir(self.history_dir)):
                        self.registry.reindex()
                        rows = self.registry.list()
                except Exception:
                    pass
            # Filtrar por término de búsqueda si se ingresó
            try:
                term = (self.hist_search_var.get() if hasattr(self, 'hist_search_var') else '').strip().lower()
            except Exception:
                term = ''
            if term:
                def _hit(e):
                    return (
                        term in str(e.get('number', '')).lower() or
                        term in str(e.get('status', '')).lower() or
                        term in str(e.get('created_at', '')).lower() or
                        term in str(e.get('pdf', '')).lower()
                    )
                rows = [e for e in rows if _hit(e)]
            
            # Ordenar por fecha (más reciente primero) por defecto
            try:
                rows = sorted(rows, key=lambda x: x.get('created_at', ''), reverse=True)
            except Exception:
                pass
            
            for e in rows:
                iid = str(e.get('number'))
                vals = (e.get('number'), e.get('status'), e.get('created_at'), e.get('pdf'), e.get('items_count'))
                self.history_tree.insert('', 'end', iid=iid, values=vals)
            self._apply_stripes(self.history_tree)
        except Exception:
            pass

    def _sort_treeview(self, tree: ttk.Treeview, col: str) -> None:
        """Ordena dinamicamente por columna (A-Z / Z-A y mayor-menor)."""
        try:
            items = [(tree.set(k, col), k) for k in tree.get_children('')]

            state_key = (id(tree), col)
            reverse = bool(self._tree_sort_state.get(state_key, False))
            self._tree_sort_state[state_key] = not reverse

            def sort_key(item):
                val = item[0]
                try:
                    return int(val)
                except Exception:
                    try:
                        return float(val)
                    except Exception:
                        return str(val).lower()

            items.sort(key=sort_key, reverse=reverse)

            for index, (_val, k) in enumerate(items):
                tree.move(k, '', index)

            self._apply_stripes(tree)
        except Exception:
            pass

    def _with_history_selection(self, action: Callable[[str, str], None]) -> None:
        cur = self.history_tree.focus()
        if not cur:
            messagebox.showwarning("Historial", MSG_SELECT_HISTORY)
            return
        try:
            num = int(cur)
            e = self.registry.find_by_number(num)
            if not e:
                messagebox.showwarning("Historial", "No se encontro informacion del vale seleccionado.")
                return
            path = os.path.join(self.history_dir, e.get('pdf', ''))
            if os.path.exists(path):
                action(path, str(e.get('pdf', '')))
            else:
                messagebox.showwarning("Historial", "No se encontro el PDF en el historial.")
        except Exception as e:
            messagebox.showerror("Abrir PDF", f"No se pudo abrir el PDF: {e}")

    def open_selected_history(self) -> None:
        def _open(path: str, _name: str) -> None:
            self._open_path(path, title="Abrir PDF")

        self._with_history_selection(_open)

    def print_selected_history(self) -> None:
        cur = self.history_tree.focus()
        if not cur:
            messagebox.showwarning("Historial", "Seleccione una solicitud del listado.")
            return
        try:
            num = int(cur)
            e = self.registry.find_by_number(num)
            if not e:
                messagebox.showwarning("Historial", "No se encontro el registro de la solicitud.")
                return
            path = os.path.join(self.history_dir, e.get('pdf', ''))
            if os.path.exists(path) and WINDOWS_OS:
                # Obtener cantidad de copias
                copias = max(1, int(self.copias_var.get()))
                print_pdf_windows(path, copies=copias, preview=True)
            elif os.path.exists(path):
                self._open_path(path, title="Reimprimir")
        except Exception as e:
            messagebox.showerror("Reimprimir", f"Error al imprimir: {e}")

    def merge_selected_history(self) -> None:
        sels = list(self.history_tree.selection())
        if not sels or len(sels) < 2:
            messagebox.showwarning("Historial", "Seleccione al menos dos vales para unificar.")
            return
        # Resolver rutas desde el registro
        input_paths = []
        for iid in sels:
            try:
                num = int(iid)
            except Exception:
                continue
            e = self.registry.find_by_number(num)
            if not e:
                continue
            p = os.path.join(self.history_dir, e.get('pdf', ''))
            if os.path.exists(p) and p.lower().endswith('.pdf'):
                input_paths.append(p)
        if len(input_paths) < 2:
            messagebox.showwarning("Historial", "No hay suficientes PDFs validos para unificar.")
            return

        # Intentar tabla unificada a partir de sidecars JSON (consolidada)
        unified_rows = []
        missing_json = []
        try:
            import json
        except Exception:
            json = None
        if json is not None:
            acc = {}
            for p in input_paths:
                base, _ = os.path.splitext(p)
                jpath = base + '.json'
                if os.path.exists(jpath):
                    try:
                        with open(jpath, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                        pdf_name = os.path.basename(p)
                        origin_num = None
                        try:
                            parts = pdf_name.split('_')
                            if len(parts) >= 2 and parts[1].isdigit():
                                origin_num = parts[1].lstrip('0') or '0'
                        except Exception:
                            origin_num = None
                        for it in data.get('items', []):
                            key = (
                                it.get('Producto', ''),
                                it.get('Lote', ''),
                                it.get('Ubicacion', ''),
                                it.get('Vencimiento', ''),
                            )
                            try:
                                qty = int(it.get('Cantidad', 0))
                            except Exception:
                                qty = 0
                            cur = acc.get(key)
                            if not cur:
                                cur = {'Cantidad': 0, 'Origenes': set()}
                                acc[key] = cur
                            cur['Cantidad'] += qty
                            cur['Origenes'].add(str(origin_num) if origin_num is not None else pdf_name)
                    except Exception:
                        missing_json.append(os.path.basename(p))
                else:
                    missing_json.append(os.path.basename(p))
            for (prod, lote, ubi, venc), info in acc.items():
                origenes = sorted(list(info.get('Origenes', [])), key=lambda x: (len(x), x))
                origen_txt = '+'.join(origenes) if origenes else ''
                unified_rows.append({
                    'Origen': origen_txt,
                    'Producto': prod,
                    'Lote': lote,
                    'Ubicacion': ubi,
                    'Vencimiento': venc,
                    'Cantidad': info.get('Cantidad', 0),
                })

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        # Nombre con origenes destacados (limitado a 5 tokens)
        try:
            origins_in_name = ''
            if unified_rows:
                tokens = []
                for r in unified_rows:
                    o = (r.get('Origen') or '').split('+')
                    for t in o:
                        if t and t not in tokens:
                            tokens.append(t)
                        if len(tokens) >= 5:
                            break
                    if len(tokens) >= 5:
                        break
                if tokens:
                    origins_in_name = '_(' + '+'.join(tokens) + ')'
            out_name = f"solicitud_unificada{origins_in_name}_{ts}.pdf"
        except Exception:
            out_name = f"solicitud_unificada_{ts}.pdf"
        out_file = os.path.join(self.history_dir, out_name)

        if unified_rows:
            try:
                from pdf_utils import build_unified_vale_pdf

                build_unified_vale_pdf(out_file, unified_rows, datetime.now())
                self.refresh_history()
                try:
                    self._open_path(out_file, title="Unificar")
                except Exception:
                    pass
                if missing_json:
                    messagebox.showinfo('Unificar', 'Se unificaron datos pero faltaron: ' + ', '.join(missing_json))
                self.log.info("Vale unificado generado (tabla) -> %s", out_file)
                return
            except Exception as e:
                self.log.exception("Fallo al crear vale unificado estructurado")
                messagebox.showwarning('Unificar', f'Fallo al crear tabla unificada ({e}). Se intentara concatenar PDFs...')

        ok, err = self._merge_pdfs(input_paths, out_file)
        if not ok:
            messagebox.showerror("Unificar", f"No se pudo unificar: {err}\nInstale 'pypdf' o 'PyPDF2'.")
            self.log.error("No se pudo unificar PDFs: %s", err)
            return
        try:
            self.refresh_history()
            self._open_path(out_file, title="Unificar")
        except Exception:
            pass
        self.log.info("Vale unificado generado (concat) -> %s", out_file)

    def _merge_pdfs(self, inputs: list[str], output: str) -> tuple[bool, str | None]:
        """Unifica PDFs con el mejor backend disponible.
        Prioriza pypdf (Merger o PdfMerger) y si falla cae a PyPDF2 o a un Writer/Reader manual.
        """
        try:
            import importlib.util as _iu

            # 1) Intentar con pypdf (varias APIs segun version)
            if _iu.find_spec('pypdf') is not None:
                # a) pypdf exporta PdfMerger
                try:
                    from pypdf import PdfMerger  # type: ignore
                    merger = PdfMerger()
                    for p in inputs:
                        merger.append(p)
                    with open(output, 'wb') as f:
                        merger.write(f)
                    try:
                        merger.close()
                    except Exception:
                        pass
                    self.log.info("PDFs unificados con pypdf -> %s", output)
                    return True, None
                except Exception:
                    # b) Algunas versiones exponen Merger en pypdf.merger
                    try:
                        from pypdf.merger import Merger  # type: ignore
                        merger = Merger()
                        for p in inputs:
                            merger.append(p)
                        with open(output, 'wb') as f:
                            merger.write(f)
                        try:
                            merger.close()
                        except Exception:
                            pass
                        self.log.info("PDFs unificados con pypdf.Merger -> %s", output)
                        return True, None
                    except Exception:
                        # c) Fallback manual con Writer/Reader
                        try:
                            from pypdf import PdfWriter, PdfReader  # type: ignore
                            writer = PdfWriter()
                            for p in inputs:
                                reader = PdfReader(p)
                                for page in getattr(reader, 'pages', []):
                                    writer.add_page(page)
                            with open(output, 'wb') as f:
                                writer.write(f)
                            self.log.info("PDFs unificados manualmente (pypdf writer) -> %s", output)
                            return True, None
                        except Exception as e3:
                            last_err = e3  # noqa: F841

            # 2) Intentar con PyPDF2
            if _iu.find_spec('PyPDF2') is not None:
                try:
                    from PyPDF2 import PdfMerger  # type: ignore
                    merger = PdfMerger()
                    for p in inputs:
                        merger.append(p)
                    with open(output, 'wb') as f:
                        merger.write(f)
                    try:
                        merger.close()
                    except Exception:
                        pass
                    self.log.info("PDFs unificados con PyPDF2 -> %s", output)
                    return True, None
                except Exception:
                    try:
                        from PyPDF2 import PdfWriter, PdfReader  # type: ignore
                        writer = PdfWriter()
                        for p in inputs:
                            reader = PdfReader(p)
                            for page in getattr(reader, 'pages', []):
                                writer.add_page(page)
                        with open(output, 'wb') as f:
                            writer.write(f)
                        self.log.info("PDFs unificados manualmente (PyPDF2 writer) -> %s", output)
                        return True, None
                    except Exception as e2:
                        return False, f"PyPDF2 fallo: {e2}"

            self.log.error("No se encontraron modulos para unir PDF")
            return False, "No se encontraron modulos 'pypdf' ni 'PyPDF2' en este interprete."
        except Exception as e:
            self.log.exception("Error inesperado al unir PDFs")
            return False, str(e)

    # -------- Interacciones --------
    def select_inventory_file(self) -> None:
        # Prefer last dir; fall back to Downloads if it exists, else cwd.
        initialdir = settings.get_last_inventory_dir()
        if not initialdir or not os.path.isdir(initialdir):
            downloads_dir = os.path.expanduser("~/Downloads")
            initialdir = downloads_dir if os.path.isdir(downloads_dir) else os.getcwd()
        path = filedialog.askopenfilename(
            title='Seleccionar archivo de inventario',
            initialdir=initialdir,
            filetypes=[('Excel', '*.xlsx *.xls'), ('Todos', '*.*')]
        )
        if not path:
            self.log.info("Seleccion de inventario cancelada")
            return
        self.current_file = path
        try:
            settings.set_last_inventory_file(path)
        except Exception:
            pass
        try:
            base_dir = os.path.dirname(path)
            if base_dir:
                settings.set_last_inventory_dir(base_dir)
                self.log.debug("Ultimo directorio de inventario actualizado: %s", base_dir)
        except Exception:
            pass
        self._load_inventory(path)

    def _load_inventory(self, path: str) -> None:
        if self._loading_inventory:
            return
        self._loading_inventory = True
        self._open_load_progress()

        def _progress(processed: int, total: Optional[int], message: str) -> None:
            self._load_queue.put(("progress", processed, total, message))

        def _worker() -> None:
            try:
                df = self.manager.load(path, AREA_FILTER, progress_cb=_progress)
                self._load_queue.put(("done", df))
            except Exception as e:
                self._load_queue.put(("error", e))

        threading.Thread(target=_worker, daemon=True).start()
        self._poll_load_queue(path)

    def _open_load_progress(self) -> None:
        if self._load_progress and self._load_progress.winfo_exists():
            return
        dlg = tk.Toplevel(self.master)
        dlg.title("Cargando inventario")
        dlg.geometry("360x120")
        dlg.resizable(False, False)
        dlg.transient(self.master)
        try:
            dlg.grab_set()
        except Exception:
            pass
        self._load_progress = dlg
        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill='both', expand=True)
        self._load_progress_label = ttk.Label(frm, text="Cargando archivo...")
        self._load_progress_label.pack(anchor='w')
        self._load_progress_bar = ttk.Progressbar(frm, mode='indeterminate')
        self._load_progress_bar.pack(fill='x', pady=(10, 0))
        self._load_progress_bar.start(10)

    def _close_load_progress(self) -> None:
        try:
            if self._load_progress_bar:
                self._load_progress_bar.stop()
        except Exception:
            pass
        try:
            if self._load_progress and self._load_progress.winfo_exists():
                self._load_progress.destroy()
        except Exception:
            pass
        self._load_progress = None
        self._load_progress_bar = None
        self._load_progress_label = None

    def _update_load_progress(self, processed: int, total: Optional[int], message: str) -> None:
        if not self._load_progress_bar:
            return
        if total and total > 0:
            if str(self._load_progress_bar["mode"]) != "determinate":
                self._load_progress_bar.stop()
                self._load_progress_bar.configure(mode="determinate", maximum=total)
            self._load_progress_bar["value"] = min(processed, total)
            pct = int((processed / total) * 100) if total else 0
            label = f"{message} {processed}/{total} ({pct}%)"
        else:
            if str(self._load_progress_bar["mode"]) != "indeterminate":
                self._load_progress_bar.configure(mode="indeterminate")
                self._load_progress_bar.start(10)
            label = message
        if self._load_progress_label:
            self._load_progress_label.configure(text=label)

    def _poll_load_queue(self, path: str) -> None:
        try:
            while True:
                kind, *payload = self._load_queue.get_nowait()
                if kind == "progress":
                    processed, total, message = payload
                    self._update_load_progress(processed, total, message)
                elif kind == "done":
                    df = payload[0]
                    self._loading_inventory = False
                    self._close_load_progress()
                    self.file_label.configure(text=os.path.basename(path))
                    self._populate_subfamilies(df)
                elif kind == "error":
                    err = payload[0]
                    self._loading_inventory = False
                    self._close_load_progress()
                    self.log.exception("No se pudo cargar inventario %s", path)
                    messagebox.showerror('Carga de Inventario', f'No se pudo cargar el archivo:\n{err}')
        except queue.Empty:
            pass
        if self._loading_inventory:
            self.master.after(120, lambda: self._poll_load_queue(path))

    def _populate_subfamilies(self, df: pd.DataFrame) -> None:
        try:
            uniq = sorted([x for x in pd.Series(df.get('Subfamilia', [])).dropna().astype(str).unique() if x])
            self.subfam_combo['values'] = ['(Todas)'] + uniq
        except Exception:
            self.subfam_combo['values'] = ['(Todas)']
        self.subfam_combo.set('(Todas)')

        self._refresh_ubicaciones_checklist()
        self.filter_products()

    def filter_products(self) -> None:
        df = self.manager.bioplates_inventory
        if df is None or df.empty:
            self._populate_products(pd.DataFrame())
            self.log.info("Filtros aplicados sin inventario cargado")
            return
        search_term = self.search_var.get().strip()
        opts = FilterOptions(
            producto=search_term,
            lote=self.lote_var.get().strip(),
            ubicacion=self.ubi_var.get().strip(),
            venc_desde=self.vdesde_var.get().strip(),
            venc_hasta=self.vhasta_var.get().strip(),
            subfamilia=self.subfam_var.get().strip() or '(Todas)',
            solo_con_stock=bool(self.stock_only_var.get()),
        )
        try:
            out = opts.apply(df)
        except Exception as exc:
            self.log.error("Filtro fallo, se muestra inventario completo: %s", exc)
            out = df.copy()
        excluded = self._get_excluded_ubicaciones()
        if excluded and 'Ubicacion' in out.columns:
            excluded_set = {str(x).strip().lower() for x in excluded}
            out = out[~out['Ubicacion'].fillna('').astype(str).str.strip().str.lower().isin(excluded_set)]
        if search_term:
            out = self._sort_by_proximidad(out)
        self.filtered_df = out
        self._populate_products(out)
        self.log.info("Filtro aplicado -> %d filas visibles", len(out))

    def _parse_vencimiento_value(self, value):
        if value is None:
            return None
        if isinstance(value, datetime):
            return value.date()
        try:
            return value.to_pydatetime().date()
        except Exception:
            pass
        try:
            num = float(value)
            if not pd.isna(num):
                return (datetime(1899, 12, 30) + timedelta(days=num)).date()
        except Exception:
            pass
        venc_str = str(value).strip()
        if not venc_str:
            return None
        for fmt in ['%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%Y/%m/%d', '%Y-%m-%d %H:%M:%S']:
            try:
                return datetime.strptime(venc_str, fmt).date()
            except Exception:
                continue
        return None

    def _sort_by_proximidad(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty or 'Vencimiento' not in df.columns:
            return df
        today = datetime.now().date()
        parsed = df['Vencimiento'].apply(self._parse_vencimiento_value)
        productos = df['Nombre_del_Producto'].fillna('').astype(str).str.strip().str.lower()
        earliest_by_producto = {}
        for prod, venc_date in zip(productos, parsed):
            if not prod or not venc_date:
                continue
            prev = earliest_by_producto.get(prod)
            if not prev or venc_date < prev:
                earliest_by_producto[prod] = venc_date
        proximo_flags = []
        days_list = []
        for prod, venc_date in zip(productos, parsed):
            if prod and venc_date and venc_date == earliest_by_producto.get(prod):
                proximo_flags.append(1)
                days_list.append((venc_date - today).days)
            else:
                proximo_flags.append(0)
                days_list.append(999999)
        out = df.copy()
        out['_proximo'] = proximo_flags
        out['_days'] = days_list
        out['_orig'] = range(len(out))
        out = out.sort_values(by=['_proximo', '_days', '_orig'], ascending=[False, True, True], kind='mergesort')
        return out.drop(columns=['_proximo', '_days', '_orig'])

    def _populate_products(self, df: pd.DataFrame) -> None:
        for i in self.product_tree.get_children():
            self.product_tree.delete(i)
        if df is None or df.empty:
            return

        today = datetime.now().date()
        earliest_by_producto = {}
        parsed_dates = {}
        for idx, row in df.iterrows():
            producto = str(row.get('Nombre_del_Producto', '')).strip()
            producto_key = producto.lower()
            venc_date = self._parse_vencimiento_value(row.get('Vencimiento', ''))
            parsed_dates[idx] = venc_date
            if producto_key and venc_date:
                prev = earliest_by_producto.get(producto_key)
                if not prev or venc_date < prev:
                    earliest_by_producto[producto_key] = venc_date

        for idx, row in df.iterrows():
            values = [
                row.get('Nombre_del_Producto', ''),
                row.get('Codigo', ''),
                row.get('Lote', ''),
                row.get('Bodega', ''),
                row.get('Ubicacion', ''),
                row.get('Vencimiento', ''),
                row.get('Stock', ''),
            ]

            # Determinar vencimiento (rojo vencido, azul proximo) solo para el mas proximo por producto
            tags = []
            venc_date = parsed_dates.get(idx)
            producto = str(row.get('Nombre_del_Producto', '')).strip()
            producto_key = producto.lower()
            if producto_key and venc_date and venc_date == earliest_by_producto.get(producto_key):
                try:
                    dias_restantes = (venc_date - today).days
                    if dias_restantes < 0:
                        tags.append('vencido')
                    else:
                        tags.append('vencimiento_proximo')
                except Exception:
                    pass
            
            self.product_tree.insert('', 'end', iid=str(int(idx)), values=values, tags=tuple(tags) if tags else ())
        
        self._apply_stripes(self.product_tree)

    def add_to_vale(self) -> None:
        sel = self.product_tree.focus()
        if not sel:
            messagebox.showwarning('Seleccion', MSG_SELECT_PRODUCT)
            return
        try:
            qty = int(self.quantity_entry.get().strip())
            if qty <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror('Cantidad', 'Ingrese una cantidad valida (> 0).')
            return
        try:
            item_index = int(sel)
            self.manager.add_to_vale(item_index, qty)
            self.log.info("Agregado item al vale (idx=%s qty=%s)", item_index, qty)
        except Exception as e:
            messagebox.showerror('Agregar a Solicitud', str(e))
            return
        # Refrescar vistas manteniendo filtros
        self.update_vale_treeview()
        self.filter_products()

    def update_vale_treeview(self) -> None:
        for i in self.vale_tree.get_children():
            self.vale_tree.delete(i)
        for i, it in enumerate(self.manager.current_vale):
            self.vale_tree.insert(
                '', 'end', iid=f'val-{i}',
                values=[it.get('Producto',''), it.get('Lote',''), it.get('Ubicacion',''), it.get('Vencimiento',''), it.get('Cantidad','')]
            )
        self._apply_stripes(self.vale_tree)

    def _apply_stripes(self, tree: ttk.Treeview) -> None:
        """Aplica zebra stripes al Treeview para mejorar legibilidad, respetando tags de vencimiento."""
        try:
            tree.tag_configure('evenrow', background='#ffffff')
            tree.tag_configure('oddrow', background='#f7f7f7')
            for n, iid in enumerate(tree.get_children()):
                # Verificar si tiene tag de vencimiento
                current_tags = tree.item(iid, 'tags')
                if current_tags and ('vencimiento_proximo' in current_tags or 'vencido' in current_tags):
                    # Mantener el tag de vencimiento
                    continue
                else:
                    # Aplicar zebra stripe
                    tree.item(iid, tags=('evenrow' if n % 2 == 0 else 'oddrow',))
        except Exception:
            pass

    # --------------- Manager de Vales ---------------
    def _build_manager_ui(self) -> None:
        frame = ttk.Frame(self.mgr_tab, padding=12)
        frame.pack(fill='both', expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        # Filtros superiores (estado)
        top = ttk.Frame(frame)
        top.grid(row=0, column=0, sticky='ew', pady=(0,8))
        ttk.Label(top, text='Estado:', font=('Segoe UI', 10)).pack(side='left')
        self.mgr_estado = tk.StringVar(value='(Todos)')
        ttk.Combobox(top, textvariable=self.mgr_estado, values=['(Todos)','Pendiente','Descontado','Anulado'], state='readonly', width=14, font=('Segoe UI', 10)).pack(side='left', padx=(6,12))
        ttk.Button(top, text='Actualizar', command=self.refresh_manager).pack(side='left')
        ttk.Button(top, text='Reindexar', command=self._mgr_reindex).pack(side='left', padx=(12,0))

        # Búsqueda
        ttk.Label(top, text=' Buscar:', font=('Segoe UI', 10)).pack(side='left', padx=(12,0))
        self.mgr_search_var = tk.StringVar(value='')
        e_msrch = ttk.Entry(top, textvariable=self.mgr_search_var, width=28, font=('Segoe UI', 10))
        e_msrch.pack(side='left', padx=(6,0))
        self.mgr_search_var.trace_add('write', lambda *_: self.refresh_manager())

        cols = ("Numero","Estado","Fecha","Archivo","Items")
        self.mgr_tree = ttk.Treeview(frame, columns=cols, show='headings', selectmode='extended')
        for c, w, anc, st in (
            ('Numero', 80, 'center', False),
            ('Estado', 110, 'center', False),
            ('Fecha', 170, 'center', False),
            ('Archivo', 520, 'w', True),
            ('Items', 80, 'center', False),
        ):
            self.mgr_tree.heading(
                c,
                text=c,
                command=lambda col=c: self._sort_treeview(self.mgr_tree, col),
            )
            self.mgr_tree.column(c, width=w, anchor=anc, stretch=st)
        self.mgr_tree.grid(row=1, column=0, sticky='nsew')
        vbar = ttk.Scrollbar(frame, orient='vertical', command=self.mgr_tree.yview)
        vbar.grid(row=1, column=1, sticky='ns')
        self.mgr_tree.configure(yscrollcommand=vbar.set)

        # Acciones
        act = ttk.Frame(frame)
        act.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(10,0))
        ttk.Button(act, text='Marcar Pendiente', command=lambda: self._mgr_set_status('Pendiente')).pack(side='left', padx=(0,8))
        ttk.Button(act, text='Marcar Descontado', command=lambda: self._mgr_set_status('Descontado')).pack(side='left', padx=(0,8))
        ttk.Button(act, text='Marcar Anulado', command=lambda: self._mgr_set_status('Anulado')).pack(side='left', padx=(0,8))
        ttk.Button(act, text='Listado PDF (Pendientes)', command=lambda: self._mgr_export_pdf('Pendiente')).pack(side='right', padx=(8,0))
        ttk.Button(act, text='Listado PDF (Descontados)', command=lambda: self._mgr_export_pdf('Descontado')).pack(side='right', padx=(8,0))
        ttk.Button(act, text='Exportar Excel (Pendientes)', command=lambda: self._mgr_export_excel('Pendiente')).pack(side='right', padx=(8,0))
        ttk.Button(act, text='Exportar Excel (Descontados)', command=lambda: self._mgr_export_excel('Descontado')).pack(side='right', padx=(8,0))

    def refresh_manager(self) -> None:
        try:
            for i in self.mgr_tree.get_children():
                self.mgr_tree.delete(i)
            estado = self.mgr_estado.get() if hasattr(self, 'mgr_estado') else '(Todos)'
            entries = self.registry.list(None if estado in (None,'', '(Todos)') else estado)
            if not entries:
                try:
                    if os.path.isdir(self.history_dir) and any(fn.lower().endswith('.pdf') for fn in os.listdir(self.history_dir)):
                        self.registry.reindex()
                        entries = self.registry.list(None if estado in (None,'', '(Todos)') else estado)
                except Exception:
                    pass
            # Filtrar por búsqueda
            try:
                term = (self.mgr_search_var.get() if hasattr(self, 'mgr_search_var') else '').strip().lower()
            except Exception:
                term = ''
            if term:
                def _hit(e):
                    return (
                        term in str(e.get('number', '')).lower() or
                        term in str(e.get('status', '')).lower() or
                        term in str(e.get('created_at', '')).lower() or
                        term in str(e.get('pdf', '')).lower()
                    )
                entries = [e for e in entries if _hit(e)]
            for e in entries:
                self.mgr_tree.insert('', 'end', iid=str(e.get('number')), values=(e.get('number'), e.get('status'), e.get('created_at'), e.get('pdf'), e.get('items_count')))
            self._apply_stripes(self.mgr_tree)
        except Exception:
            pass


    def _autosize_history_columns(self, event=None) -> None:
        """Ajusta columnas del Historial para evitar espacios en blanco."""
        try:
            tw = self.history_tree
            tw.update_idletasks()
            width = tw.winfo_width()
            if not width:
                return
            fecha_min = 180
            fecha_w = max(fecha_min, int(width * 0.22))
            archivo_w = max(200, width - fecha_w - 20)
            tw.column('Archivo', width=archivo_w, anchor='w', stretch=True)
            tw.column('Fecha', width=fecha_w, anchor='center', stretch=False)
        except Exception:
            pass

    def _mgr_reindex(self) -> None:
        try:
            res = self.registry.reindex()
            self.refresh_manager()
            self.refresh_history()
            try:
                messagebox.showinfo('Reindexar', f"Agregados: {res.get('added',0)}\nOmitidos: {res.get('skipped',0)}")
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror('Reindexar', f'No se pudo reindexar: {e}')
    def _mgr_selected_numbers(self) -> list[int]:
        try:
            return [int(i) for i in self.mgr_tree.selection()]
        except Exception:
            return []

    def _mgr_set_status(self, new_status: str) -> None:
        nums = self._mgr_selected_numbers()
        if not nums:
            messagebox.showwarning('Manager Solicitudes', 'Seleccione una o mas solicitudes.')
            return
        try:
            changed = self.registry.update_status(nums, new_status)
            if changed:
                self.refresh_manager()
                self.refresh_history()
        except Exception as e:
            messagebox.showerror('Manager Solicitudes', f'No se pudo actualizar el estado: {e}')

    def _mgr_export_excel(self, status: str) -> None:
        try:
            rows = self.registry.list(status)
            if not rows:
                messagebox.showinfo('Exportar', f'No hay solicitudes con estado {status}.')
                return
            import pandas as _pd
            df = _pd.DataFrame(rows)
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            out = os.path.join(self.history_dir, f'solicitudes_{status.lower()}_{ts}.xlsx')
            df.to_excel(out, index=False)
            try:
                self._open_path(out, title="Exportar")
            except Exception:
                pass
            messagebox.showinfo('Exportar', f'Listado exportado: {os.path.basename(out)}')
        except Exception as e:
            messagebox.showerror('Exportar', f'Error al exportar: {e}')

    def _mgr_export_pdf(self, status: str) -> None:
        try:
            rows = self.registry.list(status)
            if not rows:
                messagebox.showinfo('Listado', f'No hay solicitudes con estado {status}.')
                return
            from pdf_utils import build_vales_list_pdf
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            title = f"Listado de solicitudes - {status}"
            out = os.path.join(self.history_dir, f'listado_solicitudes_{status.lower()}_{ts}.pdf')
            build_vales_list_pdf(out, title, rows)
            try:
                self._open_path(out, title="Listado")
            except Exception:
                pass
            messagebox.showinfo('Listado', f'PDF generado: {os.path.basename(out)}')
        except Exception as e:
            messagebox.showerror('Listado', f'Error al generar PDF: {e}')

    def remove_from_vale(self) -> None:
        sel = self.vale_tree.focus()
        if not sel or not sel.startswith('val-'):
            messagebox.showwarning('Solicitud', 'Seleccione un item de la solicitud.')
            return
        try:
            idx = int(sel.split('-')[1])
            self.manager.remove_from_vale(idx)
            self.log.info("Item removido del vale idx=%s", idx)
        except Exception as e:
            self.log.warning("No se pudo remover item: %s", e)
            messagebox.showerror('Eliminar', str(e))
            return
        self.update_vale_treeview()
        self.filter_products()

    def edit_vale_item(self) -> None:
        sel = self.vale_tree.focus()
        if not sel or not sel.startswith('val-'):
            messagebox.showwarning('Solicitud', 'Seleccione un item de la solicitud.')
            return
        try:
            idx = int(sel.split('-')[1])
            item = self.manager.current_vale[idx]
            current_qty = int(item.get('Cantidad', 0))
        except Exception:
            messagebox.showerror('Editar', 'No se pudo leer el item seleccionado.')
            return
        new_qty = simpledialog.askinteger(
            "Editar Cantidad",
            "Nueva cantidad:",
            initialvalue=current_qty,
            minvalue=1,
        )
        if new_qty is None:
            return
        try:
            self.manager.update_vale_quantity(idx, int(new_qty))
        except Exception as e:
            messagebox.showerror('Editar', str(e))
            return
        self.update_vale_treeview()
        self.filter_products()

    def clear_vale(self) -> None:
        self.manager.clear_vale()
        self.log.info("Vale en curso limpiado (se restauro stock)")
        self.update_vale_treeview()
        self.filter_products()

    def generate_and_print_vale(self) -> None:
        if not self.manager.current_vale:
            messagebox.showwarning('Solicitud', 'No hay productos en la solicitud.')
            return
        
        # Validar que haya solicitante y usuario seleccionados
        solicitante = self.solicitante_var.get().strip()
        usuario_bodega = self.usuario_bodega_var.get().strip()
        if not solicitante:
            messagebox.showwarning('Solicitud', 'Debe seleccionar un solicitante.')
            return
        if not usuario_bodega:
            messagebox.showwarning('Solicitud', 'Debe seleccionar un usuario de bodega.')
            return
        
        if not os.path.exists(self.history_dir):
            os.makedirs(self.history_dir, exist_ok=True)
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        # Reservar numero para nombrar el archivo
        try:
            number = self.registry.next_number()
        except Exception:
            number = None
        if number is not None:
            padded = f"{int(number):03d}"
            base_pdf = f"solicitud_{padded}_{ts}.pdf"
            base_json = f"solicitud_{padded}_{ts}.json"
        else:
            base_pdf = f"solicitud_{ts}.pdf"
            base_json = f"solicitud_{ts}.json"
        filename = os.path.join(self.history_dir, base_pdf)
        try:
            # Agregar información de solicitante y usuario al vale
            vale_data_with_users = {
                'solicitante': solicitante,
                'usuario_bodega': usuario_bodega,
                'numero_correlativo': padded if number is not None else '',
                'items': self.manager.current_vale
            }
            
            # Generar PDF (via ValeManager -> pdf_utils)
            from pdf_utils import build_vale_pdf
            build_vale_pdf(filename, vale_data_with_users, datetime.now())
            
            # Guardar datos estructurados del vale (sidecar JSON)
            try:
                import json
                sidecar = os.path.join(self.history_dir, base_json)
                payload = {
                    'filename': os.path.basename(filename),
                    'emission_time': datetime.now().isoformat(timespec='seconds'),
                    'solicitante': solicitante,
                    'usuario_bodega': usuario_bodega,
                    'numero_correlativo': padded if number is not None else '',
                    'items': self.manager.current_vale,
                }
                with open(sidecar, 'w', encoding='utf-8') as f:
                    json.dump(payload, f, ensure_ascii=False, indent=2)
            except Exception:
                sidecar = ''
            
            # Obtener cantidad de copias
            copias = max(1, int(self.copias_var.get()))
            
            # Impresion directa (sin abrir PDF viewer)
            if WINDOWS_OS:
                print_pdf_windows(filename, copies=copias, preview=True)
            
            messagebox.showinfo('Solicitud Generada', f'Solicitud N° {padded if number is not None else "N/A"} generada correctamente.\n{copias} copia(s) enviada(s) a la impresora.')
            
            # Registrar en el indice con el numero reservado y refrescar vistas
            try:
                if number is not None:
                    self.registry.register_with_number(
                        int(number),
                        os.path.basename(filename),
                        os.path.basename(sidecar) if sidecar else '',
                        len(self.manager.current_vale),
                    )
            except Exception:
                pass
            # Limpiar vale e historial
            self.manager.current_vale = []
            self.update_vale_treeview()
            self.refresh_history()
            try:
                self.refresh_manager()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror('Generar Solicitud', f'Ocurrio un error: {e}')

    def _write_vale_sidecar(self, timestamp: str, payload: dict) -> None:
        try:
            import json

            sidecar = os.path.join(HISTORY_DIR, f'vale_{timestamp}.json')
            with open(sidecar, 'w', encoding='utf-8') as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception:
            self.log.exception("No se pudo guardar JSON sidecar para vale %s", timestamp)

    def _open_pdf_safe(self, filename: str) -> None:
        self._open_path(filename, title="Abrir PDF")

    def _open_path(self, path: str, title: str = "Abrir archivo") -> None:
        try:
            if WINDOWS_OS and hasattr(os, "startfile"):
                os.startfile(path)
                return
            if sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            self.log.warning("No se pudo abrir archivo: %s", path, exc_info=True)
            messagebox.showerror(title, f"No se pudo abrir el archivo: {e}")


def run_app() -> None:
    root = tk.Tk()
    _ = ValeConsumoApp(root)
    root.mainloop()


if __name__ == '__main__':
    run_app()
