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
            return pd.read_csv(file_path, encoding="latin-1")
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
        self.configure(bg="#f3f6fb")
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self._configure_styles()
        self._build_ui()

    def _configure_styles(self):
        self.style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"), background="#f3f6fb", foreground="#0f172a")
        self.style.configure("Card.TFrame", background="#ffffff")
        self.style.configure("Muted.TLabel", background="#f3f6fb", foreground="#475569", font=("Segoe UI", 9))
        self.style.configure("CardTitle.TLabel", background="#ffffff", foreground="#0f172a", font=("Segoe UI", 11, "bold"))
        self.style.configure("CardValue.TLabel", background="#ffffff", foreground="#0f172a", font=("Segoe UI", 18, "bold"))
        self.style.configure("Treeview", rowheight=24, font=("Segoe UI", 9))
        self.style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        self.style.configure("TNotebook.Tab", padding=(12, 8), font=("Segoe UI", 10, "bold"))
        self.style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))

    def _build_ui(self):
        top = ttk.Frame(self, padding=14)
        top.pack(fill="both", expand=True)

        header = ttk.Frame(top)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text=APP_TITLE, style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Carga catálogo y hasta 6 inventarios; consolida existencias por bodega y exporta a Excel.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        self._build_summary_cards(top)

        self.notebook = ttk.Notebook(top)
        self.notebook.pack(fill="both", expand=True, pady=(10, 0))

        self.tab_catalog = ttk.Frame(self.notebook, padding=12)
        self.tab_inventories = ttk.Frame(self.notebook, padding=12)
        self.tab_result = ttk.Frame(self.notebook, padding=12)

        self.notebook.add(self.tab_catalog, text="Catálogo")
        self.notebook.add(self.tab_inventories, text="Inventarios")
        self.notebook.add(self.tab_result, text="Consolidado")

        self._build_catalog_tab()
        self._build_inventories_tab()
        self._build_result_tab()

        bottom = ttk.Frame(top)
        bottom.pack(fill="x", pady=(8, 0))
        self.status_var = tk.StringVar(value="Listo para cargar archivos.")
        ttk.Label(bottom, textvariable=self.status_var, style="Muted.TLabel").pack(side="left")

    def _build_summary_cards(self, parent):
        cards = ttk.Frame(parent)
        cards.pack(fill="x")

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
            card = ttk.Frame(cards, style="Card.TFrame", padding=14)
            card.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 8, 0))
            cards.columnconfigure(idx, weight=1)
            ttk.Label(card, text=title, style="CardTitle.TLabel").pack(anchor="w")
            ttk.Label(card, textvariable=value_var, style="CardValue.TLabel").pack(anchor="w", pady=(8, 0))

    def _build_catalog_tab(self):
        top_actions = ttk.Frame(self.tab_catalog)
        top_actions.pack(fill="x", pady=(0, 10))

        ttk.Button(top_actions, text="Cargar catálogo", style="Primary.TButton", command=self.load_catalog).pack(side="left")
        ttk.Button(top_actions, text="Exportar plantilla catálogo", command=self.export_catalog_template).pack(side="left", padx=8)

        tip = (
            "El catálogo debe incluir al menos: codigo y proveedor. "
            "También puede traer descripcion y unidad como apoyo."
        )
        ttk.Label(self.tab_catalog, text=tip, style="Muted.TLabel").pack(anchor="w", pady=(0, 8))

        self.catalog_info_var = tk.StringVar(value="No hay catálogo cargado.")
        ttk.Label(self.tab_catalog, textvariable=self.catalog_info_var).pack(anchor="w", pady=(0, 8))

        self.catalog_tree = self._create_tree(self.tab_catalog)
        self.catalog_tree.pack(fill="both", expand=True)

    def _build_inventories_tab(self):
        actions = ttk.Frame(self.tab_inventories)
        actions.pack(fill="x", pady=(0, 10))

        ttk.Button(actions, text="Agregar inventario", style="Primary.TButton", command=self.load_inventory).pack(side="left")
        ttk.Button(actions, text="Quitar seleccionado", command=self.remove_selected_inventory).pack(side="left", padx=8)
        ttk.Button(actions, text="Limpiar lista", command=self.clear_inventories).pack(side="left")
        ttk.Button(actions, text="Exportar plantilla inventario", command=self.export_inventory_template).pack(side="left", padx=8)
        ttk.Button(actions, text="Consolidar archivos", style="Primary.TButton", command=self.consolidate).pack(side="right")

        ttk.Label(
            self.tab_inventories,
            text="Cada archivo debe traer: codigo, descripcion, unidad y existencias bodega. Se permite .xlsx, .xls y .csv.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(0, 8))

        columns = ("bodega", "archivo", "filas", "columnas")
        self.inventories_tree = ttk.Treeview(self.tab_inventories, columns=columns, show="headings", height=8)
        for col, width in [("bodega", 200), ("archivo", 360), ("filas", 90), ("columnas", 240)]:
            self.inventories_tree.heading(col, text=col.title())
            self.inventories_tree.column(col, width=width, anchor="w")
        self.inventories_tree.pack(fill="x", pady=(0, 10))

        rename_frame = ttk.LabelFrame(self.tab_inventories, text="Nombre de bodega del archivo seleccionado", padding=10)
        rename_frame.pack(fill="x", pady=(0, 10))
        self.bodega_name_var = tk.StringVar()
        ttk.Entry(rename_frame, textvariable=self.bodega_name_var).pack(side="left", fill="x", expand=True)
        ttk.Button(rename_frame, text="Actualizar nombre", command=self.rename_selected_inventory).pack(side="left", padx=8)

        ttk.Label(self.tab_inventories, text="Vista previa del archivo seleccionado", style="Muted.TLabel").pack(anchor="w", pady=(0, 6))
        self.inventory_preview_tree = self._create_tree(self.tab_inventories)
        self.inventory_preview_tree.pack(fill="both", expand=True)
        self.inventories_tree.bind("<<TreeviewSelect>>", self.on_inventory_select)

    def _build_result_tab(self):
        actions = ttk.Frame(self.tab_result)
        actions.pack(fill="x", pady=(0, 10))
        ttk.Button(actions, text="Exportar consolidado a Excel", style="Primary.TButton", command=self.export_consolidated).pack(side="left")
        ttk.Button(actions, text="Guardar CSV", command=self.export_consolidated_csv).pack(side="left", padx=8)

        self.result_info_var = tk.StringVar(value="Aún no se ha generado consolidado.")
        ttk.Label(self.tab_result, textvariable=self.result_info_var).pack(anchor="w", pady=(0, 8))

        self.result_tree = self._create_tree(self.tab_result)
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
        tree = tree_frame.winfo_children()[0]
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
            messagebox.showinfo("Plantilla creada", f"Plantilla guardada en:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

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
            messagebox.showinfo("Plantilla creada", f"Plantilla guardada en:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))


def main():
    app = InventoryConsolidatorApp()
    app.refresh_cards()
    app.mainloop()


if __name__ == "__main__":
    main()
