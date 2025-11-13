#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Interfaz de Vales de Consumo (Bioplates) en ASCII seguro (sin emojis) para
evitar texto corrupto en Windows. Incluye:
- Filtros a la derecha con scroll
- Vale en curso con acciones
- Historial con abrir, reimprimir y unificar varios PDFs
- run_app() como punto de entrada
"""

from __future__ import annotations

import os
from datetime import datetime
from typing import Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

from config import INVENTORY_FILE, AREA_FILTER, HISTORY_DIR, WINDOWS_OS
from vale_manager import ValeManager
from filters import FilterOptions
from printing_utils import print_pdf_windows
import settings_store as settings


class ValeConsumoApp:
    def __init__(self, master: tk.Tk) -> None:
        self.master = master
        self.master.title("Vale de Consumo - Bioplates")
        try:
            self.master.state('zoomed')
        except Exception:
            self.master.geometry('1280x800')

        self.manager = ValeManager()
        self.filtered_df: pd.DataFrame = pd.DataFrame()
        self.current_file: Optional[str] = None

        style = ttk.Style(self.master)
        try:
            # Tema y escalado suave en pantallas densas
            try:
                style.theme_use('clam')
            except Exception:
                pass
            try:
                # 1.15–1.25 suele ser cómodo para 1080p/2k
                self.master.tk.call('tk', 'scaling', 1.15)
            except Exception:
                pass

            # Tipografías y colores base
            base_font = ('Segoe UI', 10)
            style.configure('TLabel', font=base_font)
            style.configure('TLabelframe.Label', font=('Segoe UI', 10, 'bold'))
            style.configure('TButton', font=base_font)
            style.configure('Treeview', font=base_font, rowheight=24, background='#ffffff', fieldbackground='#ffffff')
            style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'), padding=(6, 4), background='#f0f0f0')
            style.map('Treeview', background=[('selected', '#e1ecff')], foreground=[('selected', '#000000')])

            # Botón de acción acentuado
            style.configure('Accent.TButton', background='#4a90e2', foreground='#ffffff')
            style.map('Accent.TButton', background=[('active', '#3d7fcc'), ('pressed', '#346dac')])
        except Exception:
            pass

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
        topbar = ttk.Frame(container, padding=(8, 6))
        topbar.grid(row=0, column=0, sticky='ew')
        topbar.columnconfigure(1, weight=1)

        self.select_btn = ttk.Button(topbar, text="Seleccionar archivo de inventario...", command=self.select_inventory_file)
        self.select_btn.grid(row=0, column=0, sticky='w', padx=(0, 10))

        self.file_label = ttk.Label(topbar, text="(ningun archivo cargado)")
        self.file_label.grid(row=0, column=1, sticky='w')

        if settings.get_reminder_enabled():
            lbl_txt = settings.get_reminder_text()
            self.reminder_label = ttk.Label(topbar, text=lbl_txt, foreground="#666666")
            self.reminder_label.grid(row=1, column=0, columnspan=2, sticky='w', pady=(4, 0))
        else:
            self.reminder_label = None

        # Area superior: productos + filtros
        self._build_products_area(container)
        # Area inferior: Notebook Vale / Historial
        self._build_vale_and_history(container)

    def _build_menu(self) -> None:
        menubar = tk.Menu(self.master)
        self.master.config(menu=menubar)

        m_archivo = tk.Menu(menubar, tearoff=0)
        m_archivo.add_command(label="Seleccionar inventario...", command=self.select_inventory_file)
        m_archivo.add_separator()
        m_archivo.add_command(label="Salir", command=self.master.destroy)
        menubar.add_cascade(label="Archivo", menu=m_archivo)

        m_conf = tk.Menu(menubar, tearoff=0)
        m_conf.add_command(label="Configuracion de impresora...", command=self._menu_printer_settings)
        m_conf.add_command(label="Herramientas de lectura de label...", command=self._menu_label_tools)
        menubar.add_cascade(label="Configuracion", menu=m_conf)

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

    # -------- Productos y filtros --------
    def _build_products_area(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent, padding=(8, 0))
        frame.grid(row=1, column=0, sticky='nsew')
        frame.columnconfigure(0, weight=1)

        # Tabla a la izquierda
        table_frame = ttk.Frame(frame)
        table_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 8))
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self.product_tree = ttk.Treeview(
            table_frame,
            columns=("Producto", "Lote", "Ubicacion", "Vencimiento", "Stock"),
            show='headings',
            selectmode='browse'
        )
        self.product_tree.heading('Producto', text='Producto')
        self.product_tree.heading('Lote', text='Lote')
        self.product_tree.heading('Ubicacion', text='Ubicacion')
        self.product_tree.heading('Vencimiento', text='Vencimiento')
        self.product_tree.heading('Stock', text='Stock')

        self.product_tree.column('Producto', width=280, minwidth=220, anchor='w', stretch=True)
        self.product_tree.column('Lote', width=120, minwidth=90, anchor='center', stretch=True)
        self.product_tree.column('Ubicacion', width=160, minwidth=120, anchor='center', stretch=True)
        self.product_tree.column('Vencimiento', width=130, minwidth=110, anchor='center', stretch=True)
        self.product_tree.column('Stock', width=80, minwidth=60, anchor='center', stretch=True)

        self.product_tree.grid(row=0, column=0, sticky='nsew')
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.product_tree.yview)
        vsb.grid(row=0, column=1, sticky='ns')
        hsb = ttk.Scrollbar(table_frame, orient='horizontal', command=self.product_tree.xview)
        hsb.grid(row=1, column=0, sticky='ew')
        self.product_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.product_tree.bind('<Configure>', self._autosize_product_columns)

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

        self.control_frame = ttk.Frame(filters_canvas, padding=10)
        self._filters_window = filters_canvas.create_window((0, 0), window=self.control_frame, anchor='nw')

        def _on_cf_configure(event):
            try:
                filters_canvas.configure(scrollregion=filters_canvas.bbox('all'))
            except Exception:
                pass

        def _on_canvas_configure(event):
            try:
                filters_canvas.itemconfigure(self._filters_window, width=event.width)
            except Exception:
                pass

        self.control_frame.bind('<Configure>', _on_cf_configure)
        filters_canvas.bind('<Configure>', _on_canvas_configure)

        # Rueda del mouse sobre el panel
        def _on_mousewheel(event):
            try:
                delta = 0
                if hasattr(event, 'delta') and event.delta:
                    delta = int(-1 * (event.delta / 120))
                elif getattr(event, 'num', None) == 4:
                    delta = -1
                elif getattr(event, 'num', None) == 5:
                    delta = 1
                if delta:
                    filters_canvas.yview_scroll(delta, 'units')
            except Exception:
                pass

        filters_lf.bind('<Enter>', lambda e: filters_canvas.bind_all('<MouseWheel>', _on_mousewheel))
        filters_lf.bind('<Leave>', lambda e: filters_canvas.unbind_all('<MouseWheel>'))

        # Buscar
        ttk.Label(self.control_frame, text="Buscar producto:").grid(row=0, column=0, sticky='w', pady=(0, 2))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(self.control_frame, textvariable=self.search_var, width=28)
        self.search_entry.grid(row=1, column=0, sticky='ew', pady=(0, 6))
        self.search_var.trace_add('write', lambda *_: self.filter_products())

        # Cantidad y acciones
        ttk.Label(self.control_frame, text="Cantidad a retirar:").grid(row=2, column=0, sticky='w')
        self.quantity_entry = ttk.Entry(self.control_frame, width=10)
        self.quantity_entry.insert(0, '1')
        self.quantity_entry.grid(row=3, column=0, sticky='w', pady=(0, 6))
        ttk.Button(self.control_frame, text="Agregar al Vale", style='Accent.TButton', command=self.add_to_vale, width=26).grid(row=4, column=0, sticky='ew', pady=(4, 4))
        ttk.Button(self.control_frame, text="Generar e Imprimir Vale", style='Accent.TButton', command=self.generate_and_print_vale, width=26).grid(row=5, column=0, sticky='ew', pady=(2, 8))

        # Subfamilia
        ttk.Label(self.control_frame, text="Subfamilia:").grid(row=6, column=0, sticky='w')
        self.subfam_var = tk.StringVar(value='(Todas)')
        self.subfam_combo = ttk.Combobox(self.control_frame, textvariable=self.subfam_var, state='readonly', width=26)
        self.subfam_combo.grid(row=7, column=0, sticky='ew', pady=(0, 6))
        self.subfam_combo.bind('<<ComboboxSelected>>', lambda *_: self.filter_products())

        # Lote / Ubicacion
        ttk.Label(self.control_frame, text="Lote:").grid(row=8, column=0, sticky='w')
        self.lote_var = tk.StringVar()
        ttk.Entry(self.control_frame, textvariable=self.lote_var, width=28).grid(row=9, column=0, sticky='ew', pady=(0, 6))

        ttk.Label(self.control_frame, text="Ubicacion:").grid(row=10, column=0, sticky='w')
        self.ubi_var = tk.StringVar()
        ttk.Entry(self.control_frame, textvariable=self.ubi_var, width=28).grid(row=11, column=0, sticky='ew', pady=(0, 6))

        # Rango de vencimiento
        ttk.Label(self.control_frame, text="Vencimiento desde (YYYY-MM-DD):").grid(row=12, column=0, sticky='w')
        self.vdesde_var = tk.StringVar()
        ttk.Entry(self.control_frame, textvariable=self.vdesde_var, width=28).grid(row=13, column=0, sticky='ew', pady=(0, 6))
        ttk.Label(self.control_frame, text="Vencimiento hasta (YYYY-MM-DD):").grid(row=14, column=0, sticky='w')
        self.vhasta_var = tk.StringVar()
        ttk.Entry(self.control_frame, textvariable=self.vhasta_var, width=28).grid(row=15, column=0, sticky='ew', pady=(0, 6))

        # Solo con stock
        self.stock_only_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.control_frame, text='Solo con stock', variable=self.stock_only_var, command=self.filter_products).grid(row=16, column=0, sticky='w', pady=(4, 10))

        # Limpiar filtros
        ttk.Button(self.control_frame, text="Limpiar filtros", command=self._clear_filters, width=26).grid(row=17, column=0, sticky='ew', pady=(2, 2))

        for i in range(0, 18):
            self.control_frame.rowconfigure(i, weight=0)
        self.control_frame.columnconfigure(0, weight=1)

    def _autosize_product_columns(self, event=None):
        try:
            tw = self.product_tree
            tw.update_idletasks()
            width = tw.winfo_width()
            if not width or width < 300:
                return
            specs = [
                ('Producto',    0.45, 220),
                ('Lote',        0.15,  80),
                ('Ubicacion',   0.20, 120),
                ('Vencimiento', 0.12, 100),
                ('Stock',       0.08,  60),
            ]
            for col, frac, minw in specs:
                tw.column(col, width=max(int(width * frac), minw), stretch=True)
        except Exception:
            pass

    def _clear_filters(self) -> None:
        self.search_var.set("")
        self.subfam_var.set('(Todas)')
        self.lote_var.set("")
        self.ubi_var.set("")
        self.vdesde_var.set("")
        self.vhasta_var.set("")
        self.stock_only_var.set(False)
        self.filter_products()

    # -------- Vale / Historial --------
    def _build_vale_and_history(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent, padding=(8, 6))
        frame.grid(row=2, column=0, sticky='nsew')
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self.vale_notebook = ttk.Notebook(frame)
        self.vale_notebook.grid(row=0, column=0, sticky='nsew')

        # Tab Vale
        self.vale_tab = ttk.Frame(self.vale_notebook)
        self.vale_notebook.add(self.vale_tab, text='Vale')
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
            selectmode='browse'
        )
        for col, text, w in (
            ('Producto', 'Producto', 380),
            ('Lote', 'Lote', 120),
            ('Ubicacion', 'Ubicacion', 160),
            ('Vencimiento', 'Vencimiento', 130),
            ('Cantidad', 'Cantidad', 100),
        ):
            self.vale_tree.heading(col, text=text)
            self.vale_tree.column(col, width=w, anchor='center' if col != 'Producto' else 'w', stretch=True)

        self.vale_tree.grid(row=0, column=0, sticky='nsew')
        v_vsb = ttk.Scrollbar(vale_table_frame, orient='vertical', command=self.vale_tree.yview)
        v_vsb.grid(row=0, column=1, sticky='ns')
        self.vale_tree.configure(yscrollcommand=v_vsb.set)

        self.vale_actions_frame = ttk.Frame(self.vale_tab, padding=10)
        self.vale_actions_frame.grid(row=0, column=1, sticky='ns', padx=(8, 0))
        ttk.Button(self.vale_actions_frame, text="Eliminar Producto", command=self.remove_from_vale, width=26).pack(pady=6, fill='x')
        ttk.Button(self.vale_actions_frame, text="Generar e Imprimir Vale", command=self.generate_and_print_vale, width=26).pack(pady=6, fill='x')
        ttk.Button(self.vale_actions_frame, text="Limpiar Vale", command=self.clear_vale, width=26).pack(pady=6, fill='x')

        # Tab Historial
        self.hist_tab = ttk.Frame(self.vale_notebook)
        self.vale_notebook.add(self.hist_tab, text='Historial')
        self._build_history_ui()
        self.refresh_history()

    def _build_history_ui(self) -> None:
        if not os.path.exists(HISTORY_DIR):
            os.makedirs(HISTORY_DIR, exist_ok=True)

        self.history_frame = ttk.Frame(self.hist_tab, padding=10)
        self.history_frame.pack(fill='both', expand=True)
        self.history_frame.columnconfigure(0, weight=1)
        self.history_frame.rowconfigure(0, weight=1)

        cols = ("Archivo", "Fecha")
        self.history_tree = ttk.Treeview(self.history_frame, columns=cols, show='headings', selectmode='extended')
        self.history_tree.heading('Archivo', text='Archivo')
        self.history_tree.heading('Fecha', text='Fecha')
        self.history_tree.column('Archivo', width=520, anchor='w', stretch=True)
        self.history_tree.column('Fecha', width=180, anchor='center', stretch=False)
        self.history_tree.grid(row=0, column=0, sticky='nsew')
        h_vsb = ttk.Scrollbar(self.history_frame, orient='vertical', command=self.history_tree.yview)
        h_vsb.grid(row=0, column=1, sticky='ns')
        self.history_tree.configure(yscrollcommand=h_vsb.set)
        try:
            self.history_tree.bind('<Configure>', self._autosize_history_columns)
        except Exception:
            pass

        act = ttk.Frame(self.history_frame)
        act.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(8, 0))
        ttk.Button(act, text="Unificar seleccionados", command=self.merge_selected_history).pack(side='left', padx=(0, 6))
        ttk.Button(act, text="Abrir PDF", command=self.open_selected_history).pack(side='left', padx=(0, 6))
        ttk.Button(act, text="Reimprimir", command=self.print_selected_history).pack(side='left')

    def refresh_history(self) -> None:
        for i in self.history_tree.get_children():
            self.history_tree.delete(i)
        try:
            files = [f for f in os.listdir(HISTORY_DIR) if f.lower().endswith('.pdf')]
        except Exception:
            files = []
        files.sort(reverse=True)
        for f in files:
            p = os.path.join(HISTORY_DIR, f)
            try:
                ts = datetime.fromtimestamp(os.path.getmtime(p)).strftime('%Y-%m-%d %H:%M:%S')
            except Exception:
                ts = ''
            self.history_tree.insert('', 'end', iid=f, values=(f, ts))

    def open_selected_history(self) -> None:
        cur = self.history_tree.focus()
        if not cur:
            messagebox.showwarning("Historial", "Seleccione un vale del listado.")
            return
        path = os.path.join(HISTORY_DIR, cur)
        try:
            if os.path.exists(path):
                os.startfile(path)
        except Exception as e:
            messagebox.showerror("Abrir PDF", f"No se pudo abrir el PDF: {e}")

    def print_selected_history(self) -> None:
        cur = self.history_tree.focus()
        if not cur:
            messagebox.showwarning("Historial", "Seleccione un vale del listado.")
            return
        path = os.path.join(HISTORY_DIR, cur)
        try:
            if os.path.exists(path) and WINDOWS_OS:
                print_pdf_windows(path, copies=1)
            elif os.path.exists(path):
                os.startfile(path)
        except Exception as e:
            messagebox.showerror("Reimprimir", f"Error al imprimir: {e}")

    def merge_selected_history(self) -> None:
        sels = list(self.history_tree.selection())
        if not sels or len(sels) < 2:
            messagebox.showwarning("Historial", "Seleccione al menos dos vales para unificar.")
            return
        input_paths = []
        for iid in sels:
            p = os.path.join(HISTORY_DIR, iid)
            if os.path.exists(p) and p.lower().endswith('.pdf'):
                input_paths.append(p)
        if len(input_paths) < 2:
            messagebox.showwarning("Historial", "No hay suficientes PDFs validos para unificar.")
            return

        # Intentar tabla unificada a partir de sidecars JSON
        unified_rows = []
        missing_json = []
        try:
            import json
        except Exception:
            json = None
        if json is not None:
            for p in input_paths:
                base, _ = os.path.splitext(p)
                jpath = base + '.json'
                if os.path.exists(jpath):
                    try:
                        with open(jpath, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                        for it in data.get('items', []):
                            unified_rows.append({
                                'Origen': os.path.basename(p),
                                'Producto': it.get('Producto', ''),
                                'Lote': it.get('Lote', ''),
                                'Ubicacion': it.get('Ubicacion', ''),
                                'Vencimiento': it.get('Vencimiento', ''),
                                'Cantidad': it.get('Cantidad', ''),
                            })
                    except Exception:
                        missing_json.append(os.path.basename(p))
                else:
                    missing_json.append(os.path.basename(p))

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        out_file = os.path.join(HISTORY_DIR, f"vale_unificado_{ts}.pdf")

        if unified_rows:
            try:
                from pdf_utils import build_unified_vale_pdf
                build_unified_vale_pdf(out_file, unified_rows, datetime.now())
                self.refresh_history()
                try:
                    os.startfile(out_file)
                except Exception:
                    pass
                if missing_json:
                    messagebox.showinfo('Unificar', 'Se unificaron datos pero faltaron: ' + ', '.join(missing_json))
                return
            except Exception as e:
                messagebox.showwarning('Unificar', f'Fallo al crear tabla unificada ({e}). Se intentara concatenar PDFs...')

        ok, err = self._merge_pdfs(input_paths, out_file)
        if not ok:
            messagebox.showerror("Unificar", f"No se pudo unificar: {err}\nInstale 'pypdf' o 'PyPDF2'.")
            return
        try:
            self.refresh_history()
            os.startfile(out_file)
        except Exception:
            pass

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
                        return True, None
                    except Exception as e2:
                        return False, f"PyPDF2 fallo: {e2}"

            return False, "No se encontraron modulos 'pypdf' ni 'PyPDF2' en este interprete."
        except Exception as e:
            return False, str(e)

    # -------- Interacciones --------
    def select_inventory_file(self) -> None:
        initialdir = settings.get_last_inventory_dir() or os.getcwd()
        path = filedialog.askopenfilename(
            title='Seleccionar archivo de inventario',
            initialdir=initialdir,
            filetypes=[('Excel', '*.xlsx;*.xls')]
        )
        if not path:
            return
        self.current_file = path
        try:
            base_dir = os.path.dirname(path)
            if base_dir:
                settings.set_last_inventory_dir(base_dir)
        except Exception:
            pass
        self._load_inventory(path)

    def _load_inventory(self, path: str) -> None:
        try:
            df = self.manager.load(path, AREA_FILTER)
        except Exception as e:
            messagebox.showerror('Carga de Inventario', f'No se pudo cargar el archivo:\n{e}')
            return
        self.file_label.configure(text=os.path.basename(path))
        # Subfamilias
        try:
            uniq = sorted([x for x in pd.Series(df.get('Subfamilia', [])).dropna().astype(str).unique() if x])
            self.subfam_combo['values'] = ['(Todas)'] + uniq
            self.subfam_combo.set('(Todas)')
        except Exception:
            self.subfam_combo['values'] = ['(Todas)']
            self.subfam_combo.set('(Todas)')

        self.filter_products()

    def filter_products(self) -> None:
        df = self.manager.bioplates_inventory
        if df is None or df.empty:
            self._populate_products(pd.DataFrame())
            return
        opts = FilterOptions(
            producto=self.search_var.get().strip(),
            lote=self.lote_var.get().strip(),
            ubicacion=self.ubi_var.get().strip(),
            venc_desde=self.vdesde_var.get().strip(),
            venc_hasta=self.vhasta_var.get().strip(),
            subfamilia=self.subfam_var.get().strip() or '(Todas)',
            solo_con_stock=bool(self.stock_only_var.get()),
        )
        try:
            out = opts.apply(df)
        except Exception:
            out = df.copy()
        self.filtered_df = out
        self._populate_products(out)

    def _populate_products(self, df: pd.DataFrame) -> None:
        for i in self.product_tree.get_children():
            self.product_tree.delete(i)
        if df is None or df.empty:
            return
        for idx, row in df.iterrows():
            values = [
                row.get('Nombre_del_Producto', ''),
                row.get('Lote', ''),
                row.get('Ubicacion', ''),
                row.get('Vencimiento', ''),
                row.get('Stock', ''),
            ]
            self.product_tree.insert('', 'end', iid=str(int(idx)), values=values)
        self._apply_stripes(self.product_tree)

    def add_to_vale(self) -> None:
        sel = self.product_tree.focus()
        if not sel:
            messagebox.showwarning('Seleccion', 'Seleccione un producto de la tabla.')
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
        except Exception as e:
            messagebox.showerror('Agregar al Vale', str(e))
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
        """Aplica zebra stripes al Treeview para mejorar legibilidad."""
        try:
            tree.tag_configure('evenrow', background='#ffffff')
            tree.tag_configure('oddrow', background='#fafafa')
            for n, iid in enumerate(tree.get_children()):
                tree.item(iid, tags=('evenrow' if n % 2 == 0 else 'oddrow',))
        except Exception:
            pass

    def remove_from_vale(self) -> None:
        sel = self.vale_tree.focus()
        if not sel or not sel.startswith('val-'):
            messagebox.showwarning('Vale', 'Seleccione un item del vale.')
            return
        try:
            idx = int(sel.split('-')[1])
            self.manager.remove_from_vale(idx)
        except Exception as e:
            messagebox.showerror('Eliminar', str(e))
            return
        self.update_vale_treeview()
        self.filter_products()

    def clear_vale(self) -> None:
        self.manager.clear_vale()
        self.update_vale_treeview()
        self.filter_products()

    def generate_and_print_vale(self) -> None:
        if not self.manager.current_vale:
            messagebox.showwarning('Vale', 'No hay productos en el vale.')
            return
        if not os.path.exists(HISTORY_DIR):
            os.makedirs(HISTORY_DIR, exist_ok=True)
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = os.path.join(HISTORY_DIR, f'vale_{ts}.pdf')
        try:
            # Generar PDF (via ValeManager -> pdf_utils)
            self.manager.generate_pdf(filename, datetime.now())
            # Guardar datos estructurados del vale (sidecar JSON)
            try:
                import json
                sidecar = os.path.join(HISTORY_DIR, f'vale_{ts}.json')
                payload = {
                    'filename': os.path.basename(filename),
                    'emission_time': datetime.now().isoformat(timespec='seconds'),
                    'items': self.manager.current_vale,
                }
                with open(sidecar, 'w', encoding='utf-8') as f:
                    json.dump(payload, f, ensure_ascii=False, indent=2)
            except Exception:
                pass
            # Impresion automatica si esta habilitada
            if WINDOWS_OS and settings.get_auto_print():
                print_pdf_windows(filename, copies=1)
            # Abrir para vista previa
            try:
                os.startfile(filename)
            except Exception:
                pass
            messagebox.showinfo('Vale Generado', 'Vale generado correctamente.')
            # Limpiar vale e historial
            self.manager.current_vale = []
            self.update_vale_treeview()
            self.refresh_history()
        except Exception as e:
            messagebox.showerror('Generar Vale', f'Ocurrio un error: {e}')


def run_app() -> None:
    root = tk.Tk()
    _ = ValeConsumoApp(root)
    root.mainloop()


if __name__ == '__main__':
    run_app()

    def _autosize_history_columns(self, event=None) -> None:
        """Distribuye el ancho de la tabla de historial para ocupar el espacio disponible.
        Deja 'Fecha' con un minimo y entrega el resto a 'Archivo'.
        """
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
