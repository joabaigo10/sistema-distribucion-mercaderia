from __future__ import annotations

import sqlite3
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "distribucion_demo.db"
SAMPLE_XLSX = DATA_DIR / "remitos_demo.xlsx"

LOCALES_POR_BANNER = {
    "DEPORTE": ["LOCAL CENTRO", "LOCAL NORTE", "E-COMMERCE", "ALMACEN DEPORTE"],
    "MODA": ["MODA STORE", "LOCAL SUR", "ALMACEN MODA"],
    "JB": ["LOCAL JB", "ALMACEN JB"],
}


def conectar() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def inicializar_db() -> None:
    DATA_DIR.mkdir(exist_ok=True)
    with conectar() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS remitos (
                remito TEXT PRIMARY KEY,
                empresa TEXT,
                proveedor TEXT,
                factura TEXT,
                fecha TEXT,
                estado TEXT NOT NULL DEFAULT 'no_trabajado'
            );

            CREATE TABLE IF NOT EXISTS articulos (
                remito TEXT NOT NULL,
                articulo TEXT NOT NULL,
                descripcion TEXT,
                marca TEXT,
                estado TEXT NOT NULL DEFAULT 'no_trabajado',
                PRIMARY KEY (remito, articulo)
            );

            CREATE TABLE IF NOT EXISTS stock_inicial (
                remito TEXT NOT NULL,
                articulo TEXT NOT NULL,
                talle TEXT NOT NULL,
                cantidad INTEGER NOT NULL,
                PRIMARY KEY (remito, articulo, talle)
            );

            CREATE TABLE IF NOT EXISTS asignaciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                remito TEXT NOT NULL,
                articulo TEXT NOT NULL,
                local TEXT NOT NULL,
                talle TEXT NOT NULL,
                cantidad INTEGER NOT NULL,
                fecha TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            """
        )


class DistribucionDemo(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Sistema de Distribución de Mercadería - Demo Portfolio")
        self.geometry("1180x720")
        self.minsize(1000, 620)

        inicializar_db()
        self.df_actual = pd.DataFrame()
        self.remito_actual: str | None = None
        self.articulo_actual: str | None = None
        self.talles_actuales: list[str] = []
        self.stock_actual: dict[str, int] = {}
        self.entries: dict[tuple[str, str], tk.Entry] = {}

        self.crear_ui()
        self.cargar_demo_si_vacia()
        self.refrescar_todo()

    def crear_ui(self) -> None:
        barra = ttk.Frame(self, padding=8)
        barra.pack(fill="x")
        ttk.Button(barra, text="Cargar Excel", command=self.cargar_excel).pack(side="left", padx=4)
        ttk.Button(barra, text="Cargar archivo demo", command=self.cargar_archivo_demo).pack(side="left", padx=4)
        ttk.Button(barra, text="Exportar reporte", command=self.exportar_reporte).pack(side="left", padx=4)
        ttk.Label(barra, text="Demo pública: datos ficticios + SQLite local", foreground="#555").pack(side="right")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.tab_resumen = ttk.Frame(self.notebook, padding=8)
        self.tab_asignacion = ttk.Frame(self.notebook, padding=8)
        self.tab_distribucion = ttk.Frame(self.notebook, padding=8)
        self.tab_reportes = ttk.Frame(self.notebook, padding=8)

        self.notebook.add(self.tab_resumen, text="Resumen")
        self.notebook.add(self.tab_asignacion, text="Asignación")
        self.notebook.add(self.tab_distribucion, text="Distribución")
        self.notebook.add(self.tab_reportes, text="Reportes")

        self.crear_resumen()
        self.crear_asignacion()
        self.crear_distribucion()
        self.crear_reportes()

    def crear_resumen(self) -> None:
        columnas = ("remito", "articulo", "descripcion", "total", "estado", "empresa")
        self.tree_resumen = ttk.Treeview(self.tab_resumen, columns=columnas, show="headings")
        for col, ancho in zip(columnas, (120, 150, 300, 90, 130, 150)):
            self.tree_resumen.heading(col, text=col.upper())
            self.tree_resumen.column(col, width=ancho, anchor="center")
        self.tree_resumen.pack(fill="both", expand=True)
        self.tree_resumen.bind("<Double-1>", self.abrir_desde_resumen)
        ttk.Label(self.tab_resumen, text="Doble clic en un artículo para distribuirlo.").pack(anchor="w", pady=(6, 0))

    def crear_asignacion(self) -> None:
        superior = ttk.Frame(self.tab_asignacion)
        superior.pack(fill="x")
        self.lbl_articulo = ttk.Label(superior, text="Seleccione un artículo desde Resumen", font=("Arial", 14, "bold"))
        self.lbl_articulo.pack(side="left")

        ttk.Label(superior, text="Banner:").pack(side="left", padx=(30, 4))
        self.banner_var = tk.StringVar(value="OPEN")
        self.combo_banner = ttk.Combobox(superior, textvariable=self.banner_var, values=list(LOCALES_POR_BANNER), state="readonly", width=12)
        self.combo_banner.pack(side="left")
        self.combo_banner.bind("<<ComboboxSelected>>", lambda _e: self.construir_grilla())

        self.frame_grilla = ttk.Frame(self.tab_asignacion)
        self.frame_grilla.pack(fill="both", expand=True, pady=12)

        inferior = ttk.Frame(self.tab_asignacion)
        inferior.pack(fill="x")
        self.lbl_restante = ttk.Label(inferior, text="")
        self.lbl_restante.pack(side="left")
        ttk.Button(inferior, text="Guardar asignación", command=self.guardar_asignacion).pack(side="right")

    def crear_distribucion(self) -> None:
        columnas = ("remito", "articulo", "local", "talle", "cantidad")
        self.tree_distribucion = ttk.Treeview(self.tab_distribucion, columns=columnas, show="headings")
        for col, ancho in zip(columnas, (120, 150, 230, 90, 100)):
            self.tree_distribucion.heading(col, text=col.upper())
            self.tree_distribucion.column(col, width=ancho, anchor="center")
        self.tree_distribucion.pack(fill="both", expand=True)

    def crear_reportes(self) -> None:
        filtros = ttk.Frame(self.tab_reportes)
        filtros.pack(fill="x", pady=(0, 8))
        ttk.Label(filtros, text="Artículo:").pack(side="left")
        self.filtro_articulo = tk.StringVar()
        ttk.Entry(filtros, textvariable=self.filtro_articulo, width=18).pack(side="left", padx=4)
        ttk.Label(filtros, text="Local:").pack(side="left", padx=(12, 0))
        self.filtro_local = tk.StringVar()
        ttk.Entry(filtros, textvariable=self.filtro_local, width=18).pack(side="left", padx=4)
        ttk.Button(filtros, text="Buscar", command=self.refrescar_reportes).pack(side="left", padx=8)
        ttk.Button(filtros, text="Limpiar", command=self.limpiar_filtros).pack(side="left")

        columnas = ("remito", "articulo", "descripcion", "local", "cantidad", "estado")
        self.tree_reportes = ttk.Treeview(self.tab_reportes, columns=columnas, show="headings")
        for col, ancho in zip(columnas, (120, 150, 280, 220, 100, 130)):
            self.tree_reportes.heading(col, text=col.upper())
            self.tree_reportes.column(col, width=ancho, anchor="center")
        self.tree_reportes.pack(fill="both", expand=True)

    def cargar_demo_si_vacia(self) -> None:
        with conectar() as conn:
            cantidad = conn.execute("SELECT COUNT(*) FROM articulos").fetchone()[0]
        if cantidad == 0 and SAMPLE_XLSX.exists():
            self.importar_dataframe(pd.read_excel(SAMPLE_XLSX), mostrar_mensaje=False)

    def cargar_archivo_demo(self) -> None:
        if not SAMPLE_XLSX.exists():
            messagebox.showerror("Error", "No se encontró data/remitos_demo.xlsx")
            return
        self.importar_dataframe(pd.read_excel(SAMPLE_XLSX), mostrar_mensaje=True)

    def cargar_excel(self) -> None:
        ruta = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return
        try:
            self.importar_dataframe(pd.read_excel(ruta), mostrar_mensaje=True)
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo leer el archivo:\n{exc}")

    def importar_dataframe(self, df: pd.DataFrame, mostrar_mensaje: bool) -> None:
        requeridas = {"Remito", "Family", "Size", "Quantity", "Descripcion", "Empresa", "Proveedor", "Factura"}
        faltantes = requeridas - set(df.columns)
        if faltantes:
            raise ValueError("Faltan columnas: " + ", ".join(sorted(faltantes)))

        with conectar() as conn:
            for remito, grupo_remito in df.groupby("Remito"):
                primera = grupo_remito.iloc[0]
                conn.execute(
                    "INSERT OR REPLACE INTO remitos(remito, empresa, proveedor, factura, fecha, estado) VALUES (?, ?, ?, ?, date('now'), COALESCE((SELECT estado FROM remitos WHERE remito=?), 'no_trabajado'))",
                    (str(remito), str(primera["Empresa"]), str(primera["Proveedor"]), str(primera["Factura"]), str(remito)),
                )
                for articulo, grupo_art in grupo_remito.groupby("Family"):
                    descripcion = str(grupo_art.iloc[0]["Descripcion"])
                    marca = str(grupo_art.iloc[0].get("Marca", "DEMO"))
                    conn.execute(
                        "INSERT OR REPLACE INTO articulos(remito, articulo, descripcion, marca, estado) VALUES (?, ?, ?, ?, COALESCE((SELECT estado FROM articulos WHERE remito=? AND articulo=?), 'no_trabajado'))",
                        (str(remito), str(articulo), descripcion, marca, str(remito), str(articulo)),
                    )
                    conn.execute("DELETE FROM stock_inicial WHERE remito=? AND articulo=?", (str(remito), str(articulo)))
                    agrupado = grupo_art.groupby("Size", as_index=False)["Quantity"].sum()
                    for _, fila in agrupado.iterrows():
                        conn.execute(
                            "INSERT INTO stock_inicial(remito, articulo, talle, cantidad) VALUES (?, ?, ?, ?)",
                            (str(remito), str(articulo), str(fila["Size"]), int(fila["Quantity"])),
                        )
        self.refrescar_todo()
        if mostrar_mensaje:
            messagebox.showinfo("OK", "Archivo cargado correctamente en la base demo local.")

    def refrescar_todo(self) -> None:
        self.refrescar_resumen()
        self.refrescar_distribucion()
        self.refrescar_reportes()

    def refrescar_resumen(self) -> None:
        self.tree_resumen.delete(*self.tree_resumen.get_children())
        query = """
            SELECT a.remito, a.articulo, a.descripcion,
                   SUM(s.cantidad) AS total, a.estado, r.empresa
            FROM articulos a
            JOIN remitos r ON r.remito=a.remito
            JOIN stock_inicial s ON s.remito=a.remito AND s.articulo=a.articulo
            GROUP BY a.remito, a.articulo, a.descripcion, a.estado, r.empresa
            ORDER BY a.remito, a.articulo
        """
        with conectar() as conn:
            for fila in conn.execute(query):
                self.tree_resumen.insert("", "end", values=tuple(fila))

    def abrir_desde_resumen(self, _event=None) -> None:
        seleccion = self.tree_resumen.selection()
        if not seleccion:
            return
        valores = self.tree_resumen.item(seleccion[0], "values")
        self.remito_actual, self.articulo_actual = str(valores[0]), str(valores[1])
        self.notebook.select(self.tab_asignacion)
        self.construir_grilla()

    def construir_grilla(self) -> None:
        for widget in self.frame_grilla.winfo_children():
            widget.destroy()
        self.entries.clear()
        if not self.remito_actual or not self.articulo_actual:
            return

        with conectar() as conn:
            filas = conn.execute(
                "SELECT talle, cantidad FROM stock_inicial WHERE remito=? AND articulo=? ORDER BY talle",
                (self.remito_actual, self.articulo_actual),
            ).fetchall()
            descripcion = conn.execute(
                "SELECT descripcion FROM articulos WHERE remito=? AND articulo=?",
                (self.remito_actual, self.articulo_actual),
            ).fetchone()[0]
            previas = conn.execute(
                "SELECT local, talle, cantidad FROM asignaciones WHERE remito=? AND articulo=?",
                (self.remito_actual, self.articulo_actual),
            ).fetchall()

        self.talles_actuales = [str(f[0]) for f in filas]
        self.stock_actual = {str(f[0]): int(f[1]) for f in filas}
        mapa_previo = {(str(f[0]), str(f[1])): int(f[2]) for f in previas}
        locales = LOCALES_POR_BANNER[self.banner_var.get()]

        self.lbl_articulo.config(text=f"{self.remito_actual} · {self.articulo_actual} · {descripcion}")

        ttk.Label(self.frame_grilla, text="Local", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=4, pady=4)
        for col, talle in enumerate(self.talles_actuales, start=1):
            ttk.Label(self.frame_grilla, text=f"{talle}\nStock: {self.stock_actual[talle]}", font=("Arial", 9, "bold")).grid(row=0, column=col, padx=4, pady=4)

        for fila, local in enumerate(locales, start=1):
            ttk.Label(self.frame_grilla, text=local).grid(row=fila, column=0, sticky="w", padx=4, pady=3)
            for col, talle in enumerate(self.talles_actuales, start=1):
                entry = ttk.Entry(self.frame_grilla, width=8, justify="center")
                valor = mapa_previo.get((local, talle), 0)
                if valor:
                    entry.insert(0, str(valor))
                entry.grid(row=fila, column=col, padx=4, pady=3)
                entry.bind("<KeyRelease>", lambda _e: self.actualizar_restante())
                self.entries[(local, talle)] = entry
        self.actualizar_restante()

    def actualizar_restante(self) -> None:
        if not self.stock_actual:
            self.lbl_restante.config(text="")
            return
        textos = []
        for talle in self.talles_actuales:
            asignado = 0
            for (local, t), entry in self.entries.items():
                if t != talle:
                    continue
                valor = entry.get().strip()
                if valor.isdigit():
                    asignado += int(valor)
            restante = self.stock_actual[talle] - asignado
            textos.append(f"{talle}: {restante}")
        self.lbl_restante.config(text="Restante por talle · " + " | ".join(textos))

    def guardar_asignacion(self) -> None:
        if not self.remito_actual or not self.articulo_actual:
            return
        nuevas: list[tuple[str, str, int]] = []
        totales = {t: 0 for t in self.talles_actuales}

        for (local, talle), entry in self.entries.items():
            texto = entry.get().strip()
            if not texto:
                continue
            if not texto.isdigit() or int(texto) < 0:
                messagebox.showerror("Valor inválido", f"Revisá {local} / talle {talle}.")
                return
            cantidad = int(texto)
            if cantidad:
                nuevas.append((local, talle, cantidad))
                totales[talle] += cantidad

        excedidos = [t for t in self.talles_actuales if totales[t] > self.stock_actual[t]]
        if excedidos:
            messagebox.showerror("Stock excedido", "Se excedió el stock en: " + ", ".join(excedidos))
            return

        # El sobrante se envía automáticamente al almacén del banner.
        almacen = next((l for l in LOCALES_POR_BANNER[self.banner_var.get()] if "ALMACEN" in l), "ALMACEN")
        for talle in self.talles_actuales:
            restante = self.stock_actual[talle] - totales[talle]
            if restante > 0:
                nuevas.append((almacen, talle, restante))

        with conectar() as conn:
            conn.execute("DELETE FROM asignaciones WHERE remito=? AND articulo=?", (self.remito_actual, self.articulo_actual))
            conn.executemany(
                "INSERT INTO asignaciones(remito, articulo, local, talle, cantidad) VALUES (?, ?, ?, ?, ?)",
                [(self.remito_actual, self.articulo_actual, local, talle, cantidad) for local, talle, cantidad in nuevas],
            )
            conn.execute(
                "UPDATE articulos SET estado='asignado' WHERE remito=? AND articulo=?",
                (self.remito_actual, self.articulo_actual),
            )
            pendientes = conn.execute(
                "SELECT COUNT(*) FROM articulos WHERE remito=? AND estado!='asignado'",
                (self.remito_actual,),
            ).fetchone()[0]
            conn.execute(
                "UPDATE remitos SET estado=? WHERE remito=?",
                ("asignado" if pendientes == 0 else "parcial", self.remito_actual),
            )
        self.refrescar_todo()
        messagebox.showinfo("Guardado", "Asignación guardada en SQLite. El sobrante fue enviado al almacén.")

    def refrescar_distribucion(self) -> None:
        self.tree_distribucion.delete(*self.tree_distribucion.get_children())
        with conectar() as conn:
            for fila in conn.execute("SELECT remito, articulo, local, talle, cantidad FROM asignaciones ORDER BY remito, articulo, local, talle"):
                self.tree_distribucion.insert("", "end", values=tuple(fila))

    def refrescar_reportes(self) -> None:
        self.tree_reportes.delete(*self.tree_reportes.get_children())
        articulo = f"%{self.filtro_articulo.get().strip()}%"
        local = f"%{self.filtro_local.get().strip()}%"
        query = """
            SELECT a.remito, a.articulo, a.descripcion,
                   COALESCE(asig.local, 'Pendiente de separación') AS local,
                   COALESCE(SUM(asig.cantidad), 0) AS cantidad,
                   a.estado
            FROM articulos a
            LEFT JOIN asignaciones asig ON asig.remito=a.remito AND asig.articulo=a.articulo
            WHERE a.articulo LIKE ? AND COALESCE(asig.local, '') LIKE ?
            GROUP BY a.remito, a.articulo, a.descripcion, asig.local, a.estado
            ORDER BY a.remito, a.articulo, local
        """
        with conectar() as conn:
            for fila in conn.execute(query, (articulo, local)):
                self.tree_reportes.insert("", "end", values=tuple(fila))

    def limpiar_filtros(self) -> None:
        self.filtro_articulo.set("")
        self.filtro_local.set("")
        self.refrescar_reportes()

    def exportar_reporte(self) -> None:
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="reporte_distribucion_demo.xlsx")
        if not ruta:
            return
        query = """
            SELECT a.remito, a.articulo, a.descripcion, a.marca, asig.local, asig.talle, asig.cantidad, a.estado
            FROM articulos a
            LEFT JOIN asignaciones asig ON asig.remito=a.remito AND asig.articulo=a.articulo
            ORDER BY a.remito, a.articulo, asig.local, asig.talle
        """
        with conectar() as conn:
            df = pd.read_sql_query(query, conn)
        df.to_excel(ruta, index=False, sheet_name="Distribucion")
        messagebox.showinfo("Exportado", f"Reporte guardado en:\n{ruta}")


if __name__ == "__main__":
    DistribucionDemo().mainloop()
