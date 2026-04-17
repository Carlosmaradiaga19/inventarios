import os
import re
import unicodedata
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

APP_TITLE = "Consolidado de Inventarios por Bodega"
APP_GEOMETRY = "1280x780"
MAX_BODEGAS = 6
PREVIEW_ROWS = 300


def normalize_text(value: str) -> str:
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^a-z0-9]+", " ", text).strip()
    return text


COLUMN_ALIASES = {
    "codigo": {
        "codigo", "cod", "codigo producto", "codigo articulo", "item", "sku", "cod item"
    },
    "descripcion": {
        "descripcion", "descripcion producto", "producto", "articulo", "detalle", "descrip"
    },
    "unidad": {
        "unidad", "und", "u m", "u m medida", "unidad medida", "unidad de medida", "um"
    },
    "existencias": {
        "existencias bodega", "existencia bodega", "existencia", "existencias", "stock",
        "saldo", "cantidad", "inventario", "disponible"
    },
    "proveedor": {
        "proveedor", "suplidor", "supplier", "nombre proveedor"
    },
}


def read_table(file_path: str) -> pd.DataFrame:
    ext = Path(file_path).suffix.lower()
    if ext in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(file_path)
    if ext == ".csv":
        # Intento doble por compatibilidad regional
        try:
            return pd.read_csv(file_path, encoding="utf-8-sig")
        except Exception:
            try:
                return pd.read_csv(file_path, encoding="latin-1")
            except Exception as e:
                raise ValueError(f"No se pudo leer el archivo CSV. Verifica el formato y codificación: {str(e)}")
    raise ValueError("Formato no soportado. Usa archivos .xlsx, .xls o .csv")



def find_column(columns, logical_name):
    aliases = COLUMN_ALIASES.get(logical_name, set())
    normalized_map = {normalize_text(col): col for col in columns}

    # Coincidencia exacta por alias
    for alias in aliases:
        if alias in normalized_map:
            return normalized_map[alias]

    # Coincidencia parcial
    for norm_col, original in normalized_map.items():
        for alias in aliases:
            if alias in norm_col or norm_col in alias:
                return original
    return None



def clean_code(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .replace({"nan": "", "None": "", "<NA>": ""})
    )



def parse_numeric(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(r"[^0-9\.-]", "", regex=True)
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)



def first_non_empty(values):
    for value in values:
        if pd.notna(value):
            text = str(value).strip()
            if text and text.lower() not in {"nan", "none", "<na>"}:
                return text
    return ""


class InventoryConsolidatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(APP_GEOMETRY)
        self.minsize(1100, 680)

        self.catalog_df = None
        self.inventory_items = []
        self.consolidated_df = None

        self.style = ttk.Style(self)
        self.configure(bg="#f0f7ff")
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self._configure_styles()
        self._build_ui()

    def _create_button(self, parent, text, command, is_primary=True):
        """Crea un botón personalizado con azul eléctrico y efecto 3D"""
        if is_primary:
            btn = tk.Button(
                parent, text=text, command=command,
                bg="#0066ff", fg="#ffffff", font=("Segoe UI", 10, "bold"),
                relief="raised", borderwidth=2, padx=12, pady=8,
                activebackground="#0052cc", activeforeground="#ffffff",
                cursor="hand2"
            )
        else:
            btn = tk.Button(
                parent, text=text, command=command,
                bg="#e6f2ff", fg="#0052cc", font=("Segoe UI", 9),
                relief="raised", borderwidth=2, padx=10, pady=6,
                activebackground="#cce5ff", activeforeground="#0052cc",
                cursor="hand2"
            )
        return btn

    def _configure_styles(self):
        self.style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"), background="#f0f7ff", foreground="#001f5c")
        self.style.configure("Card.TFrame", background="#ffffff", borderwidth=2, relief="raised")
        self.style.configure("Muted.TLabel", background="#f0f7ff", foreground="#1e3a8a", font=("Segoe UI", 9))
        self.style.configure("CardTitle.TLabel", background="#ffffff", foreground="#0066ff", font=("Segoe UI", 10, "bold"))
        self.style.configure("CardValue.TLabel", background="#ffffff", foreground="#0052cc", font=("Segoe UI", 20, "bold"))
        self.style.configure("Separator.TFrame", background="#0066ff", height=2)
        self.style.configure("Section.TLabel", font=("Segoe UI", 10, "bold"), background="#f0f7ff", foreground="#001f5c")
        self.style.configure("Help.TLabel", background="#f0f7ff", foreground="#1e3a8a", font=("Segoe UI", 8))
        self.style.configure("Treeview", rowheight=28, font=("Segoe UI", 9), background="#ffffff", fieldbackground="#ffffff")
        self.style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#e6f2ff", foreground="#0052cc")
        self.style.map("Treeview", background=[("selected", "#0066ff")], foreground=[("selected", "#ffffff")])
        self.style.configure("TNotebook", background="#f0f7ff", borderwidth=0)
        self.style.configure("TNotebook.Tab", padding=(20, 12), font=("Segoe UI", 10, "bold"), background="#e6f2ff", foreground="#001f5c")
        self.style.map("TNotebook.Tab", background=[("selected", "#0066ff")], foreground=[("selected", "#ffffff")])
        self.style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=8, background="#0066ff", foreground="#ffffff", borderwidth=2, relief="raised")
        self.style.map("Primary.TButton", background=[("active", "#0052cc"), ("pressed", "#003d99"), ("!active", "#0066ff")], foreground=[("active", "#ffffff"), ("!active", "#ffffff")])
        self.style.configure("Secondary.TButton", font=("Segoe UI", 9), padding=6, background="#e6f2ff", foreground="#0052cc", borderwidth=2, relief="raised")
        self.style.map("Secondary.TButton", background=[("active", "#cce5ff"), ("pressed", "#99ccff"), ("!active", "#e6f2ff")], foreground=[("active", "#0052cc"), ("!active", "#0052cc")])

    def _build_ui(self):
        # ENCABEZADO TIPO DASHBOARD
        header = tk.Frame(self, bg="#0052cc", height=80)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        
        header_content = tk.Frame(header, bg="#0052cc")
        header_content.pack(fill="both", expand=True, padx=24, pady=16)
        
        tk.Label(header_content, text=APP_TITLE, font=("Segoe UI", 24, "bold"), bg="#0052cc", fg="#ffffff").pack(anchor="w")
        tk.Label(
            header_content,
            text="Carga catálogo y hasta 6 inventarios; consolida existencias por bodega y exporta a Excel.",
            font=("Segoe UI", 9),
            bg="#0052cc",
            fg="#e0e7ff"
        ).pack(anchor="w", pady=(4, 0))
        
        # CONTENEDOR PRINCIPAL
        top = tk.Frame(self, bg="#f0f7ff")
        top.pack(fill="both", expand=True)

        # TARJETAS DE RESUMEN
        self._build_summary_cards(top)

        # NOTEBOOK CON PESTANAS
        notebook_frame = tk.Frame(top, bg="#f0f7ff")
        notebook_frame.pack(fill="both", expand=True, padx=24, pady=(0, 24))
        
        self.notebook = ttk.Notebook(notebook_frame)
        self.notebook.pack(fill="both", expand=True)

        self.tab_catalog = tk.Frame(self.notebook, bg="#ffffff")
        self.tab_inventories = tk.Frame(self.notebook, bg="#ffffff")
        self.tab_result = tk.Frame(self.notebook, bg="#ffffff")

        self.notebook.add(self.tab_catalog, text="Catálogo")
        self.notebook.add(self.tab_inventories, text="Inventarios")
        self.notebook.add(self.tab_result, text="Consolidado")

        self._build_catalog_tab()
        self._build_inventories_tab()
        self._build_result_tab()

        # BARRA DE ESTADO
        bottom = tk.Frame(top, bg="#ffffff", height=40)
        bottom.pack(fill="x", side="bottom")
        bottom.pack_propagate(False)
        
        bottom_content = tk.Frame(bottom, bg="#ffffff")
        bottom_content.pack(fill="both", expand=True, padx=24, pady=8)
        
        self.status_var = tk.StringVar(value="Listo para cargar archivos.")
        tk.Label(bottom_content, textvariable=self.status_var, font=("Segoe UI", 9), bg="#ffffff", fg="#64748b").pack(side="left")

    def _build_summary_cards(self, parent):
        cards = tk.Frame(parent, bg="#f0f7ff")
        cards.pack(fill="x", padx=24, pady=(16, 20))

        self.catalog_count_var = tk.StringVar(value="0")
        self.files_count_var = tk.StringVar(value="0")
        self.rows_count_var = tk.StringVar(value="0")
        self.columns_count_var = tk.StringVar(value="0")

        card_data = [
            ("Registros en catálogo", self.catalog_count_var),
            ("Archivos de bodega", self.files_count_var),
            ("Filas consolidadas", self.rows_count_var),
            ("Columnas en resultado", self.columns_count_var),
        ]

        for idx, (title, value_var) in enumerate(card_data):
            card = tk.Frame(cards, bg="#ffffff", relief="flat", bd=0, highlightthickness=1, highlightbackground="#e0e7ff")
            card.grid(row=0, column=idx, sticky="nsew", padx=8, pady=0)
            cards.columnconfigure(idx, weight=1)
            
            inner = tk.Frame(card, bg="#ffffff")
            inner.pack(fill="both", expand=True, padx=16, pady=16)
            
            tk.Label(inner, text=title, font=("Segoe UI", 9, "bold"), bg="#ffffff", fg="#64748b").pack(anchor="w", pady=(0, 12))
            tk.Label(inner, textvariable=value_var, font=("Segoe UI", 28, "bold"), bg="#ffffff", fg="#0052cc").pack(anchor="w", pady=(8, 0))

    def _build_catalog_tab(self):
        # PANEL IZQUIERDO - ACCIONES
        left_panel = tk.Frame(self.tab_catalog, bg="#f8f9fa", width=280)
        left_panel.pack(side="left", fill="y", padx=12, pady=12)
        left_panel.pack_propagate(False)
        
        tk.Label(left_panel, text="Acciones", font=("Segoe UI", 12, "bold"), bg="#f8f9fa", fg="#001f5c").pack(anchor="w", padx=12, pady=(12, 16))
        
        btn_container = tk.Frame(left_panel, bg="#f8f9fa")
        btn_container.pack(fill="x", padx=12, pady=(0, 16))
        
        self._create_button(btn_container, "Cargar Catalogo", self.load_catalog, True).pack(fill="x", pady=6)
        self._create_button(btn_container, "Exportar Plantilla", self.export_catalog_template, False).pack(fill="x", pady=6)
        
        tk.Frame(left_panel, bg="#e0e7ff", height=1).pack(fill="x", padx=12, pady=12)
        
        tk.Label(left_panel, text="Requerimientos", font=("Segoe UI", 10, "bold"), bg="#f8f9fa", fg="#001f5c").pack(anchor="w", padx=12, pady=(0, 8))
        
        tip = (
            "El catalogo debe incluir:\n"
            "- Codigo (obligatorio)\n"
            "- Proveedor (obligatorio)\n"
            "- Descripcion (opcional)\n"
            "- Unidad (opcional)"
        )
        tk.Label(left_panel, text=tip, font=("Segoe UI", 8), bg="#f8f9fa", fg="#64748b", justify="left").pack(anchor="nw", padx=12)
        
        # PANEL DERECHO - TABLA
        right_panel = tk.Frame(self.tab_catalog, bg="#ffffff")
        right_panel.pack(side="right", fill="both", expand=True, padx=12, pady=12)
        
        header_info = tk.Frame(right_panel, bg="#ffffff")
        header_info.pack(fill="x", pady=(0, 12))
        tk.Label(header_info, text="Contenido del Catalogo", font=("Segoe UI", 12, "bold"), bg="#ffffff", fg="#001f5c").pack(anchor="w")
        
        self.catalog_info_var = tk.StringVar(value="No hay catalogo cargado.")
        tk.Label(header_info, textvariable=self.catalog_info_var, font=("Segoe UI", 9), bg="#ffffff", fg="#64748b").pack(anchor="w", pady=(4, 0))
        
        self.catalog_tree = self._create_tree(right_panel)
        self.catalog_tree.pack(fill="both", expand=True)

    def _build_inventories_tab(self):
        # PANEL IZQUIERDO - GESTION DE BODEGAS
        left_panel = tk.Frame(self.tab_inventories, bg="#f8f9fa", width=320)
        left_panel.pack(side="left", fill="both", padx=12, pady=12)
        left_panel.pack_propagate(False)
        
        tk.Label(left_panel, text="Gestion de Bodegas", font=("Segoe UI", 12, "bold"), bg="#f8f9fa", fg="#001f5c").pack(anchor="w", padx=12, pady=(12, 16))
        
        btn_container = tk.Frame(left_panel, bg="#f8f9fa")
        btn_container.pack(fill="x", padx=12, pady=(0, 12))
        
        self._create_button(btn_container, "Agregar Bodega", self.load_inventory, True).pack(fill="x", pady=6)
        self._create_button(btn_container, "Quitar Seleccionada", self.remove_selected_inventory, False).pack(fill="x", pady=6)
        self._create_button(btn_container, "Limpiar Lista", self.clear_inventories, False).pack(fill="x", pady=6)
        self._create_button(btn_container, "Exportar Plantilla", self.export_inventory_template, False).pack(fill="x", pady=6)
        
        tk.Frame(left_panel, bg="#e0e7ff", height=1).pack(fill="x", padx=12, pady=12)
        
        tk.Label(left_panel, text="Bodegas Cargadas", font=("Segoe UI", 10, "bold"), bg="#f8f9fa", fg="#001f5c").pack(anchor="w", padx=12, pady=(0, 8))
        
        columns = ("bodega", "archivo", "filas")
        self.inventories_tree = ttk.Treeview(left_panel, columns=columns, show="headings", height=8)
        for col, width in [("bodega", 120), ("archivo", 140), ("filas", 50)]:
            self.inventories_tree.heading(col, text=col.title())
            self.inventories_tree.column(col, width=width, anchor="w")
        self.inventories_tree.pack(fill="both", expand=True, pady=(0, 12))
        
        tk.Frame(left_panel, bg="#e0e7ff", height=1).pack(fill="x", padx=12, pady=12)
        
        tk.Label(left_panel, text="Editar Nombre", font=("Segoe UI", 10, "bold"), bg="#f8f9fa", fg="#001f5c").pack(anchor="w", padx=12, pady=(0, 8))
        
        rename_content = tk.Frame(left_panel, bg="#f8f9fa")
        rename_content.pack(fill="x", padx=12, pady=(0, 12))
        
        tk.Label(rename_content, text="Nombre:", font=("Segoe UI", 9), bg="#f8f9fa", fg="#001f5c").pack(anchor="w", pady=(0, 4))
        self.bodega_name_var = tk.StringVar()
        
        entry_style = tk.Entry(rename_content, textvariable=self.bodega_name_var, font=("Segoe UI", 10), bg="#ffffff", fg="#001f5c", relief="solid", bd=1)
        entry_style.pack(fill="x", pady=(0, 8))
        
        self._create_button(rename_content, "Actualizar Nombre", self.rename_selected_inventory, True).pack(fill="x")
        
        # PANEL DERECHO - VISTA PREVIA
        right_panel = tk.Frame(self.tab_inventories, bg="#ffffff")
        right_panel.pack(side="right", fill="both", expand=True, padx=12, pady=12)
        
        tk.Label(right_panel, text="Vista Previa del Inventario", font=("Segoe UI", 12, "bold"), bg="#ffffff", fg="#001f5c").pack(anchor="w", pady=(0, 12))
        
        self.inventory_preview_tree = self._create_tree(right_panel)
        self.inventory_preview_tree.pack(fill="both", expand=True)
        self.inventories_tree.bind("<<TreeviewSelect>>", self.on_inventory_select)

    def _build_result_tab(self):
        # PANEL IZQUIERDO - EXPORTACION
        left_panel = tk.Frame(self.tab_result, bg="#f8f9fa", width=280)
        left_panel.pack(side="left", fill="y", padx=12, pady=12)
        left_panel.pack_propagate(False)
        
        tk.Label(left_panel, text="Exportacion", font=("Segoe UI", 12, "bold"), bg="#f8f9fa", fg="#001f5c").pack(anchor="w", padx=12, pady=(12, 16))
        
        btn_container = tk.Frame(left_panel, bg="#f8f9fa")
        btn_container.pack(fill="x", padx=12, pady=(0, 16))
        
        self._create_button(btn_container, "Descargar Excel", self.export_consolidated, True).pack(fill="x", pady=6)
        self._create_button(btn_container, "Descargar CSV", self.export_consolidated_csv, False).pack(fill="x", pady=6)
        
        tk.Frame(left_panel, bg="#e0e7ff", height=1).pack(fill="x", padx=12, pady=12)
        
        tk.Label(left_panel, text="Acerca del Consolidado", font=("Segoe UI", 10, "bold"), bg="#f8f9fa", fg="#001f5c").pack(anchor="w", padx=12, pady=(0, 8))
        
        info_text = (
            "El consolidado incluye:\n\n"
            "- Codigos de producto\n"
            "- Proveedores\n"
            "- Descripciones\n"
            "- Unidades de medida\n"
            "- Existencias por bodega\n"
            "- Totales por producto"
        )
        tk.Label(left_panel, text=info_text, font=("Segoe UI", 8), bg="#f8f9fa", fg="#64748b", justify="left").pack(anchor="nw", padx=12)
        
        # PANEL DERECHO - TABLA DE RESULTADOS
        right_panel = tk.Frame(self.tab_result, bg="#ffffff")
        right_panel.pack(side="right", fill="both", expand=True, padx=12, pady=12)
        
        header_frame = tk.Frame(right_panel, bg="#ffffff")
        header_frame.pack(fill="x", pady=(0, 12))
        tk.Label(header_frame, text="Consolidado de Inventarios", font=("Segoe UI", 12, "bold"), bg="#ffffff", fg="#001f5c").pack(anchor="w")
        
        self.result_info_var = tk.StringVar(value="Aun no se ha generado consolidado.")
        tk.Label(header_frame, textvariable=self.result_info_var, font=("Segoe UI", 9), bg="#ffffff", fg="#64748b").pack(anchor="w", pady=(4, 0))
        
        self.result_tree = self._create_tree(right_panel)
        self.result_tree.pack(fill="both", expand=True)

    def _create_tree(self, parent):
        frame = ttk.Frame(parent)
        tree = ttk.Treeview(frame, show="headings")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return frame

    def set_status(self, text: str):
        self.status_var.set(text)
        self.update_idletasks()

    def refresh_cards(self):
        self.catalog_count_var.set("0" if self.catalog_df is None else f"{len(self.catalog_df):,}")
        self.files_count_var.set(str(len(self.inventory_items)))
        self.rows_count_var.set("0" if self.consolidated_df is None else f"{len(self.consolidated_df):,}")
        self.columns_count_var.set("0" if self.consolidated_df is None else str(len(self.consolidated_df.columns)))

    def fill_tree_from_df(self, tree_frame, df: pd.DataFrame):
        children = tree_frame.winfo_children()
        if not children:
            return
        tree = children[0]
        for item in tree.get_children():
            tree.delete(item)

        if df is None or df.empty:
            tree["columns"] = ()
            return

        preview = df.head(PREVIEW_ROWS).copy()
        preview = preview.fillna("")
        columns = list(preview.columns)
        tree["columns"] = columns

        for col in columns:
            tree.heading(col, text=str(col))
            width = 140
            if len(str(col)) > 18:
                width = 220
            tree.column(col, width=width, minwidth=100, anchor="w")

        for row in preview.astype(str).itertuples(index=False, name=None):
            tree.insert("", "end", values=row)

    def prepare_catalog(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]

        codigo_col = find_column(df.columns, "codigo")
        proveedor_col = find_column(df.columns, "proveedor")
        descripcion_col = find_column(df.columns, "descripcion")
        unidad_col = find_column(df.columns, "unidad")

        if not codigo_col or not proveedor_col:
            raise ValueError("El catálogo debe incluir al menos las columnas codigo y proveedor.")

        result = pd.DataFrame()
        result["codigo"] = clean_code(df[codigo_col])
        result["proveedor"] = df[proveedor_col].astype(str).fillna("").str.strip()
        result["descripcion_catalogo"] = df[descripcion_col].astype(str).fillna("").str.strip() if descripcion_col else ""
        result["unidad_catalogo"] = df[unidad_col].astype(str).fillna("").str.strip() if unidad_col else ""

        result = result[result["codigo"] != ""].copy()
        result = (
            result.groupby("codigo", as_index=False)
            .agg({
                "proveedor": first_non_empty,
                "descripcion_catalogo": first_non_empty,
                "unidad_catalogo": first_non_empty,
            })
        )
        return result

    def prepare_inventory(self, df: pd.DataFrame, bodega_name: str) -> pd.DataFrame:
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]

        codigo_col = find_column(df.columns, "codigo")
        descripcion_col = find_column(df.columns, "descripcion")
        unidad_col = find_column(df.columns, "unidad")
        existencias_col = find_column(df.columns, "existencias")

        if not all([codigo_col, descripcion_col, unidad_col, existencias_col]):
            raise ValueError(
                "Cada inventario debe incluir codigo, descripcion, unidad y existencias bodega."
            )

        result = pd.DataFrame()
        result["codigo"] = clean_code(df[codigo_col])
        result["descripcion"] = df[descripcion_col].astype(str).fillna("").str.strip()
        result["unidad"] = df[unidad_col].astype(str).fillna("").str.strip()
        result["existencias"] = parse_numeric(df[existencias_col])
        result = result[result["codigo"] != ""].copy()

        result = (
            result.groupby("codigo", as_index=False)
            .agg({
                "descripcion": first_non_empty,
                "unidad": first_non_empty,
                "existencias": "sum",
            })
        )
        result["bodega"] = bodega_name
        return result

    def load_catalog(self):
        file_path = filedialog.askopenfilename(
            title="Selecciona el catálogo",
            filetypes=[("Archivos Excel/CSV", "*.xlsx *.xls *.csv")],
        )
        if not file_path:
            return

        try:
            self.set_status("Cargando catálogo...")
            raw = read_table(file_path)
            self.catalog_df = self.prepare_catalog(raw)
            self.catalog_info_var.set(
                f"Catálogo cargado: {os.path.basename(file_path)} | Registros válidos: {len(self.catalog_df):,}"
            )
            self.fill_tree_from_df(self.catalog_tree, self.catalog_df)
            self.refresh_cards()
            self.set_status("Catálogo cargado correctamente.")
            self.notebook.select(self.tab_catalog)
        except Exception as e:
            messagebox.showerror("Error al cargar catálogo", str(e))
            self.set_status("No se pudo cargar el catálogo.")

    def load_inventory(self):
        if len(self.inventory_items) >= MAX_BODEGAS:
            messagebox.showwarning("Límite alcanzado", f"Solo se permiten {MAX_BODEGAS} archivos de bodega.")
            return

        file_path = filedialog.askopenfilename(
            title="Selecciona inventario de bodega",
            filetypes=[("Archivos Excel/CSV", "*.xlsx *.xls *.csv")],
        )
        if not file_path:
            return

        try:
            self.set_status("Cargando inventario...")
            raw = read_table(file_path)
            suggested_name = Path(file_path).stem
            prepared = self.prepare_inventory(raw, suggested_name)

            item = {
                "id": f"inv_{len(self.inventory_items) + 1}_{datetime.now().strftime('%H%M%S%f')}",
                "bodega": suggested_name,
                "archivo": file_path,
                "df": prepared,
                "source_columns": ", ".join(map(str, raw.columns.tolist())),
            }
            self.inventory_items.append(item)
            self.refresh_inventory_tree()
            self.refresh_cards()
            self.set_status(f"Archivo agregado: {os.path.basename(file_path)}")
            self.notebook.select(self.tab_inventories)
        except Exception as e:
            messagebox.showerror("Error al cargar inventario", str(e))
            self.set_status("No se pudo cargar el inventario.")

    def refresh_inventory_tree(self):
        for item_id in self.inventories_tree.get_children():
            self.inventories_tree.delete(item_id)

        for item in self.inventory_items:
            self.inventories_tree.insert(
                "",
                "end",
                iid=item["id"],
                values=(
                    item["bodega"],
                    os.path.basename(item["archivo"]),
                    len(item["df"]),
                    item["source_columns"],
                ),
            )

        if self.inventory_items:
            first_id = self.inventory_items[0]["id"]
            self.inventories_tree.selection_set(first_id)
            self.show_selected_preview(first_id)
        else:
            self.fill_tree_from_df(self.inventory_preview_tree, pd.DataFrame())

    def get_selected_inventory_item(self):
        selected = self.inventories_tree.selection()
        if not selected:
            return None
        selected_id = selected[0]
        for item in self.inventory_items:
            if item["id"] == selected_id:
                return item
        return None

    def on_inventory_select(self, _event=None):
        selected = self.inventories_tree.selection()
        if selected:
            self.show_selected_preview(selected[0])

    def show_selected_preview(self, item_id: str):
        item = next((x for x in self.inventory_items if x["id"] == item_id), None)
        if not item:
            return
        self.bodega_name_var.set(item["bodega"])
        self.fill_tree_from_df(self.inventory_preview_tree, item["df"])

    def rename_selected_inventory(self):
        item = self.get_selected_inventory_item()
        if not item:
            messagebox.showwarning("Selecciona un archivo", "Primero selecciona un inventario en la lista.")
            return

        new_name = self.bodega_name_var.get().strip()
        if not new_name:
            messagebox.showwarning("Nombre inválido", "Escribe un nombre para la bodega.")
            return

        item["bodega"] = new_name
        item["df"]["bodega"] = new_name
        self.refresh_inventory_tree()
        self.inventories_tree.selection_set(item["id"])
        self.show_selected_preview(item["id"])
        self.set_status("Nombre de bodega actualizado.")

    def remove_selected_inventory(self):
        item = self.get_selected_inventory_item()
        if not item:
            messagebox.showwarning("Selecciona un archivo", "No hay ningún inventario seleccionado.")
            return

        self.inventory_items = [x for x in self.inventory_items if x["id"] != item["id"]]
        self.refresh_inventory_tree()
        self.consolidated_df = None
        self.fill_tree_from_df(self.result_tree, pd.DataFrame())
        self.result_info_var.set("Aún no se ha generado consolidado.")
        self.refresh_cards()
        self.set_status("Archivo de bodega removido.")

    def clear_inventories(self):
        self.inventory_items.clear()
        self.consolidated_df = None
        self.refresh_inventory_tree()
        self.fill_tree_from_df(self.result_tree, pd.DataFrame())
        self.result_info_var.set("Aún no se ha generado consolidado.")
        self.refresh_cards()
        self.set_status("Se limpiaron los archivos de bodega.")

    def consolidate(self):
        if not self.inventory_items:
            messagebox.showwarning("Sin archivos", "Debes cargar al menos un inventario de bodega.")
            return

        try:
            self.set_status("Consolidando inventarios...")
            base_frames = []
            warehouse_cols = []

            for item in self.inventory_items:
                df = item["df"].copy()
                warehouse_col = f"existencia_{item['bodega']}"
                warehouse_cols.append(warehouse_col)
                df = df[["codigo", "descripcion", "unidad", "existencias"]].rename(columns={"existencias": warehouse_col})
                base_frames.append(df)

            all_codes = pd.concat([df[["codigo"]] for df in base_frames], ignore_index=True).drop_duplicates()
            consolidated = all_codes.copy()

            # Descripción y unidad tomando el primer valor no vacío entre todos los archivos
            descriptions = pd.concat([df[["codigo", "descripcion"]] for df in base_frames], ignore_index=True)
            descriptions = descriptions.groupby("codigo", as_index=False).agg({"descripcion": first_non_empty})
            units = pd.concat([df[["codigo", "unidad"]] for df in base_frames], ignore_index=True)
            units = units.groupby("codigo", as_index=False).agg({"unidad": first_non_empty})

            consolidated = consolidated.merge(descriptions, on="codigo", how="left")
            consolidated = consolidated.merge(units, on="codigo", how="left")

            for df in base_frames:
                consolidated = consolidated.merge(df, on=["codigo", "descripcion", "unidad"], how="left")

            for col in warehouse_cols:
                consolidated[col] = consolidated[col].fillna(0)

            if self.catalog_df is not None:
                consolidated = consolidated.merge(self.catalog_df, on="codigo", how="left")
                consolidated["proveedor"] = consolidated["proveedor"].fillna("")
                consolidated["descripcion"] = consolidated.apply(
                    lambda r: r["descripcion"] if str(r["descripcion"]).strip() else r.get("descripcion_catalogo", ""),
                    axis=1,
                )
                consolidated["unidad"] = consolidated.apply(
                    lambda r: r["unidad"] if str(r["unidad"]).strip() else r.get("unidad_catalogo", ""),
                    axis=1,
                )
                consolidated = consolidated.drop(columns=[c for c in ["descripcion_catalogo", "unidad_catalogo"] if c in consolidated.columns])
            else:
                consolidated["proveedor"] = ""

            consolidated["total_existencias"] = consolidated[warehouse_cols].sum(axis=1)

            ordered_cols = ["codigo", "proveedor", "descripcion", "unidad", *warehouse_cols, "total_existencias"]
            consolidated = consolidated[ordered_cols].sort_values(by=["proveedor", "descripcion", "codigo"]).reset_index(drop=True)

            self.consolidated_df = consolidated
            self.fill_tree_from_df(self.result_tree, self.consolidated_df)
            self.result_info_var.set(
                f"Consolidado generado correctamente | Productos: {len(self.consolidated_df):,} | Bodegas: {len(self.inventory_items)}"
            )
            self.refresh_cards()
            self.set_status("Consolidado listo.")
            self.notebook.select(self.tab_result)
        except Exception as e:
            messagebox.showerror("Error al consolidar", str(e))
            self.set_status("Ocurrió un error al consolidar.")

    def export_consolidated(self):
        if self.consolidated_df is None or self.consolidated_df.empty:
            messagebox.showwarning("Sin consolidado", "Primero debes generar el consolidado.")
            return

        default_name = f"consolidado_inventarios_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = filedialog.asksaveasfilename(
            title="Guardar consolidado en Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Archivo Excel", "*.xlsx")],
        )
        if not file_path:
            return

        try:
            self.set_status("Exportando a Excel...")
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                self.consolidated_df.to_excel(writer, index=False, sheet_name="Consolidado")

                resumen = pd.DataFrame(
                    {
                        "Indicador": [
                            "Fecha de generación",
                            "Archivos de bodega cargados",
                            "Productos consolidados",
                            "Registros en catálogo",
                        ],
                        "Valor": [
                            datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                            len(self.inventory_items),
                            len(self.consolidated_df),
                            0 if self.catalog_df is None else len(self.catalog_df),
                        ],
                    }
                )
                resumen.to_excel(writer, index=False, sheet_name="Resumen")

            self.set_status(f"Consolidado exportado: {file_path}")
            messagebox.showinfo("Exportación completada", f"Archivo guardado correctamente:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error al exportar", str(e))
            self.set_status("No se pudo exportar el consolidado.")

    def export_consolidated_csv(self):
        if self.consolidated_df is None or self.consolidated_df.empty:
            messagebox.showwarning("Sin consolidado", "Primero debes generar el consolidado.")
            return

        default_name = f"consolidado_inventarios_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        file_path = filedialog.asksaveasfilename(
            title="Guardar consolidado en CSV",
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("Archivo CSV", "*.csv")],
        )
        if not file_path:
            return

        try:
            self.consolidated_df.to_csv(file_path, index=False, encoding="utf-8-sig")
            self.set_status(f"CSV exportado: {file_path}")
            messagebox.showinfo("Exportación completada", f"Archivo guardado correctamente:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error al exportar CSV", str(e))
            self.set_status("No se pudo exportar el CSV.")

    def export_catalog_template(self):
        file_path = filedialog.asksaveasfilename(
            title="Guardar plantilla de catálogo",
            defaultextension=".xlsx",
            initialfile="plantilla_catalogo.xlsx",
            filetypes=[("Archivo Excel", "*.xlsx")],
        )
        if not file_path:
            return

        try:
            df = pd.DataFrame(
                {
                    "codigo": ["1001", "1002"],
                    "descripcion": ["Arroz 1 lb", "Frijol rojo 1 lb"],
                    "unidad": ["UND", "UND"],
                    "proveedor": ["Proveedor A", "Proveedor B"],
                }
            )
            df.to_excel(file_path, index=False)
            self.set_status("Plantilla de catálogo exportada.")
            messagebox.showinfo("Plantilla creada", f"Plantilla guardada en:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error al exportar plantilla", f"No se pudo guardar la plantilla: {str(e)}")
            self.set_status("Error al exportar plantilla de catálogo.")

    def export_inventory_template(self):
        file_path = filedialog.asksaveasfilename(
            title="Guardar plantilla de inventario",
            defaultextension=".xlsx",
            initialfile="plantilla_inventario_bodega.xlsx",
            filetypes=[("Archivo Excel", "*.xlsx")],
        )
        if not file_path:
            return

        try:
            df = pd.DataFrame(
                {
                    "codigo": ["1001", "1002"],
                    "descripcion": ["Arroz 1 lb", "Frijol rojo 1 lb"],
                    "unidad": ["UND", "UND"],
                    "existencias bodega": [150, 90],
                }
            )
            df.to_excel(file_path, index=False)
            self.set_status("Plantilla de inventario exportada.")
            messagebox.showinfo("Plantilla creada", f"Plantilla guardada en:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error al exportar plantilla", f"No se pudo guardar la plantilla: {str(e)}")
            self.set_status("Error al exportar plantilla de inventario.")


def main():
    app = InventoryConsolidatorApp()
    app.refresh_cards()
    app.mainloop()


if __name__ == "__main__":
    main()
