#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Aplicación de Escritorio para Gestión de Vales de Consumo - Bioplates
# Desarrollado con Python y Tkinter.

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
import os
import sys

from config import INVENTORY_FILE, AREA_FILTER, HISTORY_DIR, WINDOWS_OS
from data_loader import load_inventory
from pdf_utils import build_vale_pdf
from settings_store import get_last_inventory_dir, set_last_inventory_dir

class ValeConsumoApp:
    def __init__(self, master):
        self.master = master
        master.title("Gestión de Vale de Consumo - Bioplates")
        master.geometry("1200x700")

        # Variables de estado
        self.inventory_df = pd.DataFrame()
        self.bioplates_inventory = pd.DataFrame()
        self.current_vale = [] # Lista de diccionarios para el vale: [{'Producto': '...', 'Lote': '...', 'Vencimiento': '...', 'Cantidad': x, 'Stock_Original': y}, ...]
        self.inventory_file = INVENTORY_FILE  # Archivo de inventario seleccionado

        # Configuración de la interfaz
        self.setup_ui()

        # Carga inicial de datos (no fatal si falla; permite seleccionar archivo luego)
        self.load_inventory_data(fatal=False)

        # Crear directorio de historial si no existe
        if not os.path.exists(HISTORY_DIR):
            os.makedirs(HISTORY_DIR)

    def load_inventory_data(self, filepath=None, fatal=True):
        """Carga el archivo Excel usando data_loader y actualiza la UI."""
        try:
            if filepath is not None:
                self.inventory_file = filepath
            file_to_read = self.inventory_file

            # Cargar inventario normalizado
            self.bioplates_inventory = load_inventory(file_to_read, AREA_FILTER)

            self.update_product_treeview()
            messagebox.showinfo("Carga Exitosa", f"Inventario cargado y filtrado. {len(self.bioplates_inventory)} productos de Bioplates disponibles.")
            if hasattr(self, 'file_label'):
                try:
                    shown = os.path.basename(self.inventory_file)
                except Exception:
                    shown = str(self.inventory_file)
                self.file_label.configure(text=f"Archivo: {shown}")

        except FileNotFoundError:
            messagebox.showerror("Error de Archivo", f"El archivo de inventario '{self.inventory_file}' no se encontro.")
            if fatal:
                self.master.quit()
        except Exception as e:
            messagebox.showerror("Error de Carga", f"Ocurrio un error al leer el archivo Excel: {e}")
            if fatal:
                self.master.quit()

    # --- Configuración de la Interfaz (Tkinter) ---
    def setup_ui(self):
        # Paneles principales
        self.main_frame = ttk.Frame(self.master, padding="10")
        self.main_frame.pack(fill="both", expand=True)

        # Barra superior: Selecci3n de archivo de inventario
        self.file_frame = ttk.Frame(self.main_frame, padding="5")
        self.file_frame.pack(side="top", fill="x", padx=5, pady=5)

        ttk.Button(self.file_frame, text="Seleccionar archivo de inventario", command=self.browse_inventory_file).pack(side="left", padx=(0,10))
        try:
            initial_shown = os.path.basename(self.inventory_file)
        except Exception:
            initial_shown = str(self.inventory_file)
        self.file_label = ttk.Label(self.file_frame, text=f"Archivo: {initial_shown}")
        self.file_label.pack(side="left")

        # Marco superior: Productos disponibles
        self.products_frame = ttk.LabelFrame(self.main_frame, text=f"Productos Disponibles ({AREA_FILTER})", padding="10")
        self.products_frame.pack(side="top", fill="both", expand=True, padx=5, pady=5)

        # Marco inferior: Vale en curso
        self.vale_frame = ttk.LabelFrame(self.main_frame, text="Vale de Consumo en Curso", padding="10")
        self.vale_frame.pack(side="bottom", fill="x", padx=5, pady=5)

        # --- Productos Disponibles (Treeview) ---
        self.product_tree = ttk.Treeview(self.products_frame, columns=('Producto', 'Lote', 'Ubicacion', 'Vencimiento', 'Stock'), show='headings')
        self.product_tree.heading('Producto', text='Nombre del Producto')
        self.product_tree.heading('Lote', text='Lote')
        self.product_tree.heading('Ubicacion', text='Ubicacion')
        self.product_tree.heading('Vencimiento', text='Fecha de Vencimiento')
        self.product_tree.heading('Stock', text='Stock')
        
        self.product_tree.column('Producto', width=300, anchor='w')
        self.product_tree.column('Lote', width=150, anchor='center')
        self.product_tree.column('Ubicacion', width=140, anchor='center')
        self.product_tree.column('Vencimiento', width=150, anchor='center')
        self.product_tree.column('Stock', width=100, anchor='center')

        self.product_tree.pack(side="left", fill="both", expand=True)
        
        # Scrollbar para la tabla de productos
        vsb = ttk.Scrollbar(self.products_frame, orient="vertical", command=self.product_tree.yview)
        vsb.pack(side='right', fill='y')
        self.product_tree.configure(yscrollcommand=vsb.set)

        # Botones y entrada de cantidad
        self.control_frame = ttk.Frame(self.products_frame, padding="10")
        self.control_frame.pack(side="right", fill="y")

        ttk.Label(self.control_frame, text="Cantidad a Retirar:").pack(pady=5)
        self.quantity_entry = ttk.Entry(self.control_frame, width=10)
        self.quantity_entry.pack(pady=5)
        self.quantity_entry.insert(0, "1") # Valor por defecto

        ttk.Button(self.control_frame, text="Agregar al Vale", command=self.add_to_vale).pack(pady=10)
        
        # Campo de búsqueda
        ttk.Label(self.control_frame, text="Buscar Producto:").pack(pady=5, padx=10)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(self.control_frame, textvariable=self.search_var)
        self.search_entry.pack(pady=5, padx=10)
        self.search_var.trace_add("write", lambda name, index, mode: self.filter_products())
        
        ttk.Button(self.control_frame, text="Limpiar Búsqueda", command=lambda: self.search_var.set("")).pack(pady=5)
        # Filtros integrados
        ttk.Label(self.control_frame, text='Lote:').pack(pady=3, padx=10)
        self.lote_var = tk.StringVar()
        ttk.Entry(self.control_frame, textvariable=self.lote_var).pack(pady=3, padx=10)

        ttk.Label(self.control_frame, text='Ubicacion:').pack(pady=3, padx=10)
        self.ubi_var = tk.StringVar()
        ttk.Entry(self.control_frame, textvariable=self.ubi_var).pack(pady=3, padx=10)

        ttk.Label(self.control_frame, text='Vencimiento desde (YYYY-MM-DD):').pack(pady=3, padx=10)
        self.vdesde_var = tk.StringVar()
        ttk.Entry(self.control_frame, textvariable=self.vdesde_var).pack(pady=3, padx=10)
        ttk.Label(self.control_frame, text='Vencimiento hasta (YYYY-MM-DD):').pack(pady=3, padx=10)
        self.vhasta_var = tk.StringVar()
        ttk.Entry(self.control_frame, textvariable=self.vhasta_var).pack(pady=3, padx=10)

        self.stock_only_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.control_frame, text='Solo con stock', variable=self.stock_only_var, command=self.filter_products).pack(pady=3, padx=10)

        # Aplicar filtros al escribir
        self.lote_var.trace_add('write', lambda *args: self.filter_products())
        self.ubi_var.trace_add('write', lambda *args: self.filter_products())
        self.vdesde_var.trace_add('write', lambda *args: self.filter_products())
        self.vhasta_var.trace_add('write', lambda *args: self.filter_products())

        ttk.Button(self.control_frame, text='Limpiar filtros', command=self._clear_filters).pack(pady=6)



        # --- Vale en Curso (Treeview) ---
        self.vale_tree = ttk.Treeview(self.vale_frame, columns=('Producto', 'Lote', 'Vencimiento', 'Cantidad'), show='headings', height=5)
        self.vale_tree.heading('Producto', text='Nombre del Producto')
        self.vale_tree.heading('Lote', text='Lote')
        self.vale_tree.heading('Vencimiento', text='Fecha de Vencimiento')
        self.vale_tree.heading('Cantidad', text='Cantidad')

        self.vale_tree.column('Producto', width=400, anchor='w')
        self.vale_tree.column('Lote', width=150, anchor='center')
        self.vale_tree.column('Vencimiento', width=150, anchor='center')
        self.vale_tree.column('Cantidad', width=100, anchor='center')
        
        self.vale_tree.pack(side="left", fill="x", expand=True)

        # Botones de acción del vale
        self.vale_actions_frame = ttk.Frame(self.vale_frame, padding="10")
        self.vale_actions_frame.pack(side="right", fill="y")

        ttk.Button(self.vale_actions_frame, text="Eliminar Producto", command=self.remove_from_vale).pack(pady=5, fill='x')
        ttk.Button(self.vale_actions_frame, text="Generar e Imprimir Vale", command=self.generate_and_print_vale).pack(pady=10, fill='x')
        ttk.Button(self.vale_actions_frame, text="Limpiar Vale", command=self.clear_vale).pack(pady=5, fill='x')

    def browse_inventory_file(self):
        """Abre un di��logo para seleccionar el archivo de inventario y recarga los datos."""
        filetypes = [
            ("Archivos de Excel", "*.xlsx;*.xls"),
            ("Todos los archivos", "*.*"),
        ]
        initial_dir = get_last_inventory_dir() or os.path.dirname(self.inventory_file) or os.getcwd()
        selected = filedialog.askopenfilename(title="Seleccionar archivo de inventario", filetypes=filetypes, initialdir=initial_dir)
        if not selected:
            return
        # Actualizar etiqueta inmediatamente
        try:
            shown = os.path.basename(selected)
        except Exception:
            shown = str(selected)
        if hasattr(self, 'file_label'):
            self.file_label.configure(text=f"Archivo: {shown}")
        # Guardar la ubicación del inventario seleccionado
        try:
            dirpath = os.path.dirname(selected)
            if dirpath:
                set_last_inventory_dir(dirpath)
        except Exception:
            pass

        # Cargar datos desde el archivo seleccionado (no fatal si falla)
        self.load_inventory_data(filepath=selected, fatal=False)

    def _clear_filters(self):
        if hasattr(self, 'search_var'): self.search_var.set('')
        if hasattr(self, 'lote_var'): self.lote_var.set('')
        if hasattr(self, 'ubi_var'): self.ubi_var.set('')
        if hasattr(self, 'vdesde_var'): self.vdesde_var.set('')
        if hasattr(self, 'vhasta_var'): self.vhasta_var.set('')
        if hasattr(self, 'stock_only_var'): self.stock_only_var.set(False)
        self.update_product_treeview(self.bioplates_inventory)



    def update_product_treeview(self, filtered_df=None):
        """Actualiza la tabla de productos disponibles."""
        
        # Limpiar tabla actual
        for item in self.product_tree.get_children():
            self.product_tree.delete(item)
        
        df_to_show = filtered_df if filtered_df is not None else self.bioplates_inventory
        
        # Insertar nuevos datos
        for index, row in df_to_show.iterrows():
            self.product_tree.insert('', 'end', values=(row['Nombre_del_Producto'], row['Lote'], row.get('Ubicacion', ''), row['Vencimiento'], row['Stock']), iid=index)

    def filter_products(self):
        """Filtra los productos disponibles segun los filtros integrados."""
        df = self.bioplates_inventory.copy()

        # Texto de producto
        term = self.search_var.get().strip().lower() if hasattr(self, 'search_var') else ''
        if term:
            mask_p = df['Nombre_del_Producto'].str.lower().str.contains(term, na=False)
            df = df[mask_p]

        # Lote
        lote_t = self.lote_var.get().strip().lower() if hasattr(self, 'lote_var') else ''
        if lote_t:
            mask_l = df['Lote'].str.lower().str.contains(lote_t, na=False)
            df = df[mask_l]

        # Ubicacion
        ubi_t = self.ubi_var.get().strip().lower() if hasattr(self, 'ubi_var') else ''
        if ubi_t and 'Ubicacion' in df.columns:
            mask_u = df['Ubicacion'].astype(str).str.lower().str.contains(ubi_t, na=False)
            df = df[mask_u]

        # Vencimiento rango
        from datetime import datetime
        vdesde = self.vdesde_var.get().strip() if hasattr(self, 'vdesde_var') else ''
        vhasta = self.vhasta_var.get().strip() if hasattr(self, 'vhasta_var') else ''
        def _parse(d):
            try:
                return datetime.strptime(d, '%Y-%m-%d')
            except Exception:
                return None
        d1 = _parse(vdesde) if vdesde else None
        d2 = _parse(vhasta) if vhasta else None
        if d1 or d2:
            # Vencimiento esta en formato YYYY-MM-DD
            vdt = df['Vencimiento'].astype(str)
            vdt = vdt.where(vdt.ne('NaT'), None)
            def _ok(s):
                try:
                    t = datetime.strptime(s, '%Y-%m-%d')
                except Exception:
                    return False
                if d1 and t < d1: return False
                if d2 and t > d2: return False
                return True
            df = df[vdt.apply(_ok)]

        # Stock
        if hasattr(self, 'stock_only_var') and self.stock_only_var.get():
            if 'Stock' in df.columns:
                df = df[df['Stock'] > 0]

        self.update_product_treeview(df)


    # --- Lógica de Gestión de Vale y Validaciones ---
    def add_to_vale(self):
        """Añade el producto seleccionado al vale, aplicando validaciones."""
        selected_item = self.product_tree.focus()
        if not selected_item:
            messagebox.showwarning("Advertencia", "Debe seleccionar un producto de la lista de disponibles.")
            return

        try:
            quantity_to_remove = int(self.quantity_entry.get())
        except ValueError:
            messagebox.showerror("Error de Entrada", "La cantidad debe ser un número entero.")
            return
        
        if quantity_to_remove <= 0:
            messagebox.showerror("Error de Entrada", "La cantidad a retirar debe ser positiva.")
            return

        # Obtener los datos del producto seleccionado
        item_index = int(selected_item) # El iid es el índice del DataFrame
        product_data = self.bioplates_inventory.loc[item_index]
        
        current_stock = product_data['Stock']

        # Validación de stock
        if quantity_to_remove > current_stock:
            messagebox.showerror("Error de Stock", f"No hay suficiente stock. Stock disponible: {current_stock}")
            return

        # Actualizar el stock en el DataFrame en memoria
        self.bioplates_inventory.loc[item_index, 'Stock'] -= quantity_to_remove
        
        # Actualizar la tabla de productos disponibles
        self.update_product_treeview()

        # Preparar el ítem para el vale
        new_item = {
            'Producto': product_data['Nombre_del_Producto'],
            'Lote': product_data['Lote'],
            'Vencimiento': product_data['Vencimiento'],
            'Ubicacion': product_data.get('Ubicacion', ''),
            'Cantidad': quantity_to_remove,
            'Stock_Original_Index': item_index # Guardar el índice para la reversión
        }

        # Comprobar si el producto (por Nombre y Lote) ya está en el vale
        found = False
        for item in self.current_vale:
            if (
                item['Producto'] == new_item['Producto'] and
                item['Lote'] == new_item['Lote'] and
                item.get('Ubicacion', '') == new_item.get('Ubicacion', '')
            ):
                item['Cantidad'] += quantity_to_remove
                found = True
                break
        if not found:
            self.current_vale.append(new_item)

        self.update_vale_treeview()
        self.quantity_entry.delete(0, tk.END)
        self.quantity_entry.insert(0, '1')
        messagebox.showinfo("Producto Agregado", f"Se han agregado {quantity_to_remove} unidades de '{product_data['Nombre_del_Producto']}' al vale.")

    def remove_from_vale(self):
        """Elimina un producto del vale y revierte el stock."""
        selected_item = self.vale_tree.focus()
        if not selected_item:
            messagebox.showwarning("Advertencia", "Debe seleccionar un producto del vale en curso para eliminar.")
            return
            
        # El iid en vale_tree es el índice de la lista self.current_vale
        item_index = self.vale_tree.index(selected_item)
        item_to_remove = self.current_vale.pop(item_index)

        # Revertir el stock
        stock_index = item_to_remove['Stock_Original_Index']
        quantity_reverted = item_to_remove['Cantidad']
        
        self.bioplates_inventory.loc[stock_index, 'Stock'] += quantity_reverted
        
        self.update_product_treeview()
        self.update_vale_treeview()
        
        messagebox.showinfo("Producto Eliminado", f"Se ha revertido el stock de {quantity_reverted} unidades de '{item_to_remove['Producto']}'.")

    def clear_vale(self):
        """Limpia el vale en curso y revierte todo el stock."""
        if not self.current_vale:
            messagebox.showinfo("Información", "El vale ya está vacío.")
            return

        if not messagebox.askyesno("Confirmar Limpieza", "¿Está seguro de que desea limpiar el vale y revertir todo el stock retirado?"):
            return

        for item in self.current_vale:
            stock_index = item['Stock_Original_Index']
            quantity_reverted = item['Cantidad']
            self.bioplates_inventory.loc[stock_index, 'Stock'] += quantity_reverted

        self.current_vale = []
        self.update_product_treeview()
        self.update_vale_treeview()
        messagebox.showinfo("Vale Limpiado", "El vale ha sido limpiado y el stock revertido.")

    def update_vale_treeview(self):
        """Actualiza la tabla del vale en curso."""
        for item in self.vale_tree.get_children():
            self.vale_tree.delete(item)
            
        for i, item in enumerate(self.current_vale):
            self.vale_tree.insert('', 'end', values=(item['Producto'], item['Lote'], item['Vencimiento'], item['Cantidad']), iid=i)

    # --- Generación de PDF e Impresión ---
    def generate_and_print_vale(self):
        """Genera el PDF y lo abre en el navegador para imprimir manualmente."""
        if not self.current_vale:
            messagebox.showwarning("Advertencia", "El vale de consumo está vacío. Agregue productos antes de generar.")
            return

    # Generar el nombre del archivo
        now = datetime.now()
        timestamp = now.strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(HISTORY_DIR, f"Vale_Consumo_Bioplates_{timestamp}.pdf")

        try:
            build_vale_pdf(filename, self.current_vale, now)
            # Abrir en navegador predeterminado
            import webbrowser
            from pathlib import Path
            url = Path(filename).resolve().as_uri()
            webbrowser.open(url)
            messagebox.showinfo("Vale Generado", "Vale generado y abierto en el navegador para impresion.")
            # Limpiar el vale después de la generación/impresión
            self.current_vale = []
            self.update_vale_treeview()

        except Exception as e:
            messagebox.showerror("Error de PDF/Impresión", f"Ocurrió un error al generar o imprimir el vale: {e}")


    def create_pdf(self, filename, vale_data, emission_time):
        """Obsoleto: usar pdf_utils.build_vale_pdf."""
        raise NotImplementedError("Usar pdf_utils.build_vale_pdf")


# --- Punto de entrada de la aplicación ---
def run_app() -> None:
    """Inicializa Tk y ejecuta la app."""
    root = tk.Tk()
    _ = ValeConsumoApp(root)
    root.mainloop()


if __name__ == "__main__":
    run_app()
