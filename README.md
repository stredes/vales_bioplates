# Sistema de Solicitud de Productos (Uso Bodega) - Bioplates

Sistema de gestión de solicitudes de productos para bodega con control de inventario, generación de PDFs e impresión directa.

## Características Principales

### ✅ Nuevas Funcionalidades (Actualización 2026)

- **Nomenclatura actualizada**: "Solicitud de Productos (Uso Bodega)" en lugar de "Vale de Consumo"
- **Resaltado de vencimientos**: Productos con fechas próximas de vencimiento (30 días o menos) resaltados en **rojo**
- **Gestión de usuarios**:
  - Lista de solicitantes personalizable
  - Lista de usuarios de bodega personalizable
  - Selección obligatoria al generar solicitud
- **Numeración correlativa**: Solicitudes numeradas desde 000 (solicitud_001, solicitud_002, etc.)
- **Control de ubicaciones**: Checkbox para mostrar/ocultar columna de ubicaciones
- **Cantidad de impresiones**: Selector de 1-10 copias
- **Impresión directa**: Envío directo a impresora sin abrir visor de PDF
- **Historial mejorado**:
  - Ordenado por fecha (más reciente primero) por defecto
  - Clic en encabezados de columna para ordenar
  - Doble clic sobre item para abrir PDF
- **PDF mejorado**: Línea punteada para "Doc. Asociado" en pie de página

### Gestión de Inventario

- Carga archivos Excel (.xlsx, .xls) con inventario
- Normalización automática de columnas
- **Alerta visual de vencimientos**: Productos que vencen en 30 días o menos resaltados en **rojo**
- Filtros múltiples:
  - Búsqueda por texto en producto
  - Subfamilia
  - Lote
  - Ubicación
  - Rango de fechas de vencimiento
  - Solo productos con stock

### Creación de Solicitudes

- Selección de productos del inventario
- Validación de stock disponible
- Consolidación automática de productos duplicados
- Descuento automático del stock
- Selección obligatoria de solicitante y usuario de bodega

### Generación de PDFs

- Estructura profesional:
  - Número correlativo de solicitud
  - Fecha y hora de emisión
  - Solicitante y usuario de bodega
  - Tabla con: Producto, Lote, Ubicación, Vencimiento, Cantidad
  - Líneas de firma (Entregado por / Recibido por)
  - Línea punteada para "Doc. Asociado"
- Guardado automático en carpeta `Vales_Historial/`
- JSON complementario con datos estructurados

### Historial

- Listado completo de solicitudes generadas
- Búsqueda por texto
- Ordenamiento por columnas (clic en encabezado)
- Doble clic para abrir PDF
- Acciones disponibles:
  - Abrir PDF
  - Reimprimir (con cantidad seleccionable)
  - Unificar seleccionadas

### Impresión

- Impresión directa a impresora predeterminada
- Soporte para SumatraPDF (impresión silenciosa)
- Cantidad configurable de copias (1-10)
- No abre visor de PDF

## Instalación

### Requisitos

- Windows 10/11
- Python 3.8 o superior
- Lector de PDF (recomendado: SumatraPDF para impresión silenciosa)

### Instalación de dependencias

```bash
pip install -r requirements.txt
```

### Ejecución

```bash
python run_app.py
```

O ejecutar el `.exe` compilado con PyInstaller:

```bash
# Compilar
python build_utf8.ps1

# Ejecutar
dist/ValeConsumoBioplates.exe
```

## Uso

### Primer Uso

1. **Configurar usuarios**:
   - Menú `Configuración` → `Gestionar Solicitantes...`
   - Agregar nombres de personas que pueden solicitar productos
   - Menú `Configuración` → `Gestionar Usuarios Bodega...`
   - Agregar nombres de usuarios que preparan las solicitudes

2. **Cargar inventario**:
   - Clic en `Seleccionar archivo de inventario...`
   - Seleccionar archivo Excel con inventario

### Crear Solicitud

1. **Filtrar productos** (opcional):
   - Usar búsqueda por texto
   - Aplicar filtros de subfamilia, lote, ubicación, vencimiento
   - Activar "Solo con stock" para ocultar productos sin disponibilidad

2. **Seleccionar usuarios**:
   - Elegir **Solicitante** (quien pide los productos)
   - Elegir **Usuario Bodega** (quien prepara la solicitud)

3. **Agregar productos**:
   - Seleccionar producto de la tabla
   - Especificar cantidad a retirar
   - Clic en `Agregar a Solicitud`
   - Repetir para cada producto necesario

4. **Generar e imprimir**:
   - Seleccionar cantidad de copias (1-10)
   - Clic en `Generar e Imprimir Solicitud`
   - Se imprime directamente sin abrir PDF

### Gestionar Historial

- **Buscar**: Escribir en campo de búsqueda
- **Ordenar**: Hacer clic en encabezado de columna
- **Abrir PDF**: Doble clic sobre la solicitud
- **Reimprimir**: Seleccionar solicitud → Clic en `Reimprimir`
- **Unificar**: Seleccionar múltiples solicitudes → Clic en `Unificar seleccionados`

### Atajos de Teclado

- `Ctrl+F`: Enfocar búsqueda de productos
- `Ctrl+H`: Ir a pestaña Historial
- `Ctrl+M`: Ir a pestaña Manager Solicitudes
- `Ctrl+G`: Generar e imprimir solicitud
- `Supr`: Eliminar item seleccionado de la solicitud

## Estructura de Archivos

```
proyecto/
├── run_app.py                     # Punto de entrada
├── vale_consumo_bioplates.py      # Interfaz gráfica principal
├── vale_manager.py                # Lógica de negocio
├── data_loader.py                 # Carga y normalización de Excel
├── pdf_utils.py                   # Generación de PDFs
├── filters.py                     # Sistema de filtros
├── vale_registry.py               # Registro de solicitudes
├── user_manager.py                # Gestión de usuarios y solicitantes
├── printing_utils.py              # Utilidades de impresión
├── settings_store.py              # Persistencia de configuración
├── config.py                      # Configuración global
├── requirements.txt               # Dependencias Python
├── build_utf8.ps1                 # Script de compilación
├── ValeConsumoBioplates.spec      # Configuración PyInstaller
├── app_settings.json              # Configuración de aplicación
├── Vales_Historial/               # PDFs y JSONs generados
│   ├── solicitud_001_*.pdf
│   ├── solicitud_001_*.json
│   ├── solicitantes.json
│   ├── usuarios_bodega.json
│   └── vales_index.json
└── instrucciones.txt              # Instrucciones detalladas
```

## Configuración

### Archivo `config.py`

- `INVENTORY_FILE`: Archivo Excel por defecto
- `AREA_FILTER`: Filtro de área ("Bioplates")
- `HISTORY_DIR`: Carpeta de historial
- `SUMATRA_PDF_PATH`: Ruta a SumatraPDF (opcional)

### Archivo `app_settings.json`

Generado automáticamente, contiene:
- Último directorio de inventario usado
- Configuración de autoimpresión
- Recordatorios

## Resolución de Problemas

### No imprime

**Problema**: Error 31 o no imprime  
**Solución**: 
1. Instalar SumatraPDF: https://www.sumatrapdfreader.org/
2. Configurar en `config.py`: `SUMATRA_PDF_PATH = r"C:\Program Files\SumatraPDF\SumatraPDF.exe"`

### No ve productos después de cargar Excel

**Problema**: Tabla vacía tras cargar archivo  
**Solución**: Verificar que el Excel tenga columnas: Producto, Lote, Fecha de Vencimiento, Cantidad/Stock, Ubicación

### No puede generar solicitud

**Problema**: Error al generar  
**Solución**: Verificar que haya:
1. Solicitante seleccionado
2. Usuario de bodega seleccionado
3. Al menos un producto agregado

## Tecnologías

- **Python 3.8+**
- **Tkinter**: Interfaz gráfica
- **Pandas**: Procesamiento de datos
- **ReportLab**: Generación de PDFs
- **PyWin32**: Impresión en Windows
- **PyPDF**: Unificación de PDFs

## Licencia

Uso interno - Bioplates

## Soporte

Para reportar problemas o sugerencias, contactar al administrador del sistema.
