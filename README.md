# Consolidado de Inventarios por Bodega

**Una aplicación de escritorio para unificar y centralizar datos de inventario de múltiples bodegas.**

---

## 📋 Descripción

Esta herramienta permite a **negocio** consolidar inventarios procedentes de hasta 6 bodegas diferentes en un único reporte centralizado. Combina un catálogo maestro de productos con inventarios individuales de cada bodega, generando automáticamente un consolidado con totales por producto y bodega.

### Beneficios Principales
- **Consolidación automática** de múltiples inventarios
- **Reconocimiento inteligente** de variaciones en nombres de columnas
- **Reporte unificado** con existencias por bodega y totales
- **Exportación flexible** (Excel con resumen o CSV)
- **Interfaz gráfica intuitiva** con vista previa en tiempo real

---

## 🚀 Requisitos del Sistema

- **Python**: 3.8 o superior
- **Sistema Operativo**: Windows, macOS o Linux
- **Dependencias**:
  - `pandas` - Manipulación y análisis de datos
  - `openpyxl` - Lectura/escritura de archivos Excel

---

## 📦 Instalación

### 1. Clonar o descargar el proyecto
```bash
git clone <url-del-repositorio>
cd inventarios
```

### 2. Crear un entorno virtual (recomendado)
```bash
python -m venv venv
# En Windows:
venv\Scripts\activate
# En macOS/Linux:
source venv/bin/activate
```

### 3. Instalar dependencias
```bash
pip install pandas openpyxl
```

### 4. Ejecutar la aplicación
```bash
python app_consolidado_inventarios.py
```

---

## 📖 Cómo Usar

### Paso 1: Preparar el Catálogo (Opcional)
1. Abre la pestaña **"Catálogo"**
2. Haz clic en **"Exportar plantilla catálogo"** para descargar un ejemplo
3. Completa con tus productos:
   - `codigo` - Código único del producto (requerido)
   - `descripcion` - Nombre del producto
   - `unidad` - Unidad de medida (UND, KG, LB, etc.)
   - `proveedor` - Nombre del proveedor (requerido)
4. Guarda como `.xlsx` o `.csv`
5. Carga el archivo con **"Cargar catálogo"**

### Paso 2: Cargar Inventarios de Bodegas
1. Abre la pestaña **"Inventarios"**
2. Haz clic en **"Agregar inventario"** (repite hasta 6 veces)
3. Cada archivo debe contener:
   - `codigo` - Código del producto (requerido)
   - `descripcion` - Nombre del producto (requerido)
   - `unidad` - Unidad de medida (requerido)
   - `existencias bodega` - Cantidad en stock (requerido)
4. (Opcional) Renombra cada bodega en el campo **"Nombre de bodega"**
5. Usa **"Exportar plantilla inventario"** para un ejemplo

### Paso 3: Generar Consolidado
1. Abre la pestaña **"Inventarios"**
2. Haz clic en **"Consolidar archivos"**
3. El sistema procesará automáticamente todos los datos
4. Ve a la pestaña **"Consolidado"** para ver el resultado

### Paso 4: Exportar Resultado
- **"Exportar consolidado a Excel"** - Genera archivo con 2 hojas:
  - `Consolidado` - Reporte principal con existencias por bodega
  - `Resumen` - Estadísticas de generación
- **"Guardar CSV"** - Exporta solo el consolidado en formato CSV

---

## 🏗️ Estructura del Proyecto

```
inventarios/
├── app_consolidado_inventarios.py    # Aplicación principal
├── README.md                          # Esta documentación
└── .git/                              # Control de versiones
```

---

## 🔧 Características Principales

### Reconocimiento Inteligente de Columnas
La aplicación detecta automáticamente las columnas independientemente de cómo se nombren:

| Columna Lógica | Alias Reconocidos |
|---|---|
| **codigo** | cod, codigo producto, codigo articulo, item, sku, cod item |
| **descripcion** | descripcion producto, producto, articulo, detalle, descrip |
| **unidad** | und, u m, u m medida, unidad medida, unidad de medida, um |
| **existencias** | existencias bodega, existencia bodega, stock, saldo, cantidad, inventario, disponible |
| **proveedor** | suplidor, supplier, nombre proveedor |

### Procesamiento de Datos
- ✅ Normalización de códigos (elimina decimales `.0`, espacios)
- ✅ Limpieza de valores nulos y texturas inválidas
- ✅ Agrupación automática por código (suma de existencias duplicadas)
- ✅ Conversión flexible de números (acepta comas como separadores)
- ✅ Fusión inteligente de catálogo con inventarios

### Consolidación
El consolidado resultante contiene:
- **codigo** - Identificador único del producto
- **proveedor** - Origen del producto
- **descripcion** - Nombre del producto
- **unidad** - Unidad de medida
- **existencia_[bodega]** - Existencias en cada bodega cargada
- **total_existencias** - Suma total de todas las bodegas

Los resultados se ordenan por: **Proveedor → Descripción → Código**

---

## 📊 Interfaz Gráfica

### Componentes Principales
- **Tarjetas de Resumen** - Contadores en tiempo real:
  - Registros en catálogo
  - Archivos de bodega cargados
  - Filas consolidadas
  - Columnas en resultado

- **3 Pestañas Principales**:
  - 📑 **Catálogo** - Carga y visualización del maestro de productos
  - 📦 **Inventarios** - Gestión de archivos de bodega y vista previa
  - 📋 **Consolidado** - Reporte final y exportación

- **Barra de Estado** - Mensajes de progreso y confirmación

---

## 💡 Ejemplos de Uso

### Ejemplo 1: Consolidar 3 Bodegas
```
1. Cargar catálogo_productos.xlsx
2. Cargar bodega_central.xlsx → Renombrar a "Bodega Central"
3. Cargar bodega_norte.xlsx → Renombrar a "Bodega Norte"
4. Cargar bodega_sur.xlsx → Renombrar a "Bodega Sur"
5. Consolidar
6. Exportar a Excel
→ Resultado: consolidado_inventarios_20260417_143022.xlsx
```

### Ejemplo 2: Sin Catálogo
```
1. Cargar solo inventarios (sin catálogo)
2. Consolidar
→ Resultado: consolidado con existencias por bodega
   (sin información de proveedor)
```

---

## ⚙️ Configuración

### Constantes Editables
En la parte superior del archivo `app_consolidado_inventarios.py`:

```python
APP_TITLE = "Consolidado de Inventarios por Bodega"
APP_GEOMETRY = "1280x780"        # Dimensiones de ventana
MAX_BODEGAS = 6                  # Número máximo de bodegas
PREVIEW_ROWS = 300               # Filas a mostrar en vista previa
```

---

## 🐛 Solución de Problemas

| Problema | Solución |
|---|---|
| **"Formato no soportado"** | Usa archivos `.xlsx`, `.xls` o `.csv` |
| **Columnas no reconocidas** | Verifica los alias en la tabla de características; renombra columnas si es necesario |
| **Códigos con decimales** | Se limpian automáticamente (no es necesario hacer nada) |
| **Caracteres especiales** | Se normalizan automáticamente (acentos, ñ, etc.) |
| **Archivos muy grandes** | La vista previa muestra max 300 filas; los datos completos se procesan |

---

## 📤 Formatos de Entrada Soportados

- **Catálogo**: `.xlsx`, `.xls`, `.csv`
- **Inventarios**: `.xlsx`, `.xls`, `.csv`
- **Codificación**: UTF-8 o Latin-1 (detecta automáticamente)

## 📥 Formatos de Salida

- **Excel** (`.xlsx`): Consolidad + Resumen
- **CSV** (`.csv`): Consolidado con separadores de coma

---

## 📝 Requisitos Mínimos de Datos

### Catálogo
```
codigo (requerido) | proveedor (requerido) | descripcion | unidad
1001              | Proveedor A          | Arroz 1 lb  | UND
1002              | Proveedor B          | Frijol 1 lb | UND
```

### Inventario de Bodega
```
codigo (requerido) | descripcion (requerido) | unidad (requerido) | existencias bodega (requerido)
1001              | Arroz 1 lb              | UND               | 150
1002              | Frijol 1 lb             | UND               | 90
```

---

## 🔐 Notas de Seguridad

- Los archivos se procesan **localmente** (sin conexión a internet)
- Los datos se mantienen en **memoria durante la sesión**
- Las exportaciones se guardan donde especifiques
- **Sin credenciales** ni autenticación requerida

---

## 🎨 Interfaz Visual

- **Tema**: Claro (Light Mode compatible)
- **Fuente**: Segoe UI (escalable)
- **Colores**: Azul profesional (#0f172a) con tonos neutrales
- **Responsive**: Tamaño mínimo 1100x680 px

---

## 📞 Contacto y Soporte

Para reportar problemas o sugerencias, contacta al equipo de desarrollo.

---

## 📄 Licencia

Desarrollado para **BANASUPRO** (Suplidora Nacional de Productos Básicos)

---

## 🔄 Historial de Versiones

| Versión | Fecha | Descripción |
|---|---|---|
| 1.0 | 17/04/2026 | Versión inicial - Consolidación de 6 bodegas |

---

## ✨ Características Futuras Potenciales

- [ ] Soporte para más de 6 bodegas
- [ ] Validación de datos con reporte de inconsistencias
- [ ] Gráficas de distribución de inventario
- [ ] Historial de consolidaciones anteriores
- [ ] Exportación a formatos adicionales (JSON, PDF)
- [ ] Sincronización con sistemas externos

---

**Última actualización**: 17 de abril de 2026
