from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Iterable

import pandas as pd

REQUIRED_COLUMNS = {
    "Remito", "Family", "Size", "Quantity", "Descripcion",
    "Empresa", "Proveedor", "Factura",
}


class DistributionStore:
    """Persistence and business rules for the portfolio demo."""

    def __init__(self, db_path: Path) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.initialize()

    def connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def initialize(self) -> None:
        with self.connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS remitos (
                    remito TEXT PRIMARY KEY,
                    empresa TEXT NOT NULL,
                    proveedor TEXT NOT NULL,
                    factura TEXT NOT NULL,
                    fecha TEXT NOT NULL DEFAULT CURRENT_DATE,
                    estado TEXT NOT NULL DEFAULT 'no_trabajado'
                );

                CREATE TABLE IF NOT EXISTS articulos (
                    remito TEXT NOT NULL,
                    articulo TEXT NOT NULL,
                    descripcion TEXT NOT NULL,
                    marca TEXT NOT NULL DEFAULT 'DEMO',
                    estado TEXT NOT NULL DEFAULT 'no_trabajado',
                    PRIMARY KEY (remito, articulo),
                    FOREIGN KEY (remito) REFERENCES remitos(remito) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS stock_inicial (
                    remito TEXT NOT NULL,
                    articulo TEXT NOT NULL,
                    talle TEXT NOT NULL,
                    cantidad INTEGER NOT NULL CHECK(cantidad >= 0),
                    PRIMARY KEY (remito, articulo, talle),
                    FOREIGN KEY (remito, articulo)
                        REFERENCES articulos(remito, articulo) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS asignaciones (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    remito TEXT NOT NULL,
                    articulo TEXT NOT NULL,
                    banner TEXT NOT NULL,
                    local TEXT NOT NULL,
                    talle TEXT NOT NULL,
                    cantidad INTEGER NOT NULL CHECK(cantidad > 0),
                    fecha TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (remito, articulo)
                        REFERENCES articulos(remito, articulo) ON DELETE CASCADE
                );

                CREATE INDEX IF NOT EXISTS idx_asignaciones_articulo
                    ON asignaciones(remito, articulo);
                CREATE INDEX IF NOT EXISTS idx_asignaciones_local
                    ON asignaciones(local);
                """
            )

    @staticmethod
    def validate_dataframe(df: pd.DataFrame) -> pd.DataFrame:
        missing = REQUIRED_COLUMNS - set(df.columns)
        if missing:
            raise ValueError("Faltan columnas obligatorias: " + ", ".join(sorted(missing)))
        if df.empty:
            raise ValueError("El archivo no contiene filas para importar.")

        clean = df.copy()
        for column in REQUIRED_COLUMNS - {"Quantity"}:
            clean[column] = clean[column].fillna("").astype(str).str.strip()
        clean["Marca"] = clean.get("Marca", "DEMO")
        clean["Marca"] = clean["Marca"].fillna("DEMO").astype(str).str.strip().replace("", "DEMO")
        clean["Quantity"] = pd.to_numeric(clean["Quantity"], errors="coerce")

        invalid = clean["Quantity"].isna() | (clean["Quantity"] < 0) | (clean["Quantity"] % 1 != 0)
        if invalid.any():
            rows = ", ".join(str(i + 2) for i in clean.index[invalid][:8])
            raise ValueError(f"Quantity debe contener enteros no negativos. Revisá las filas: {rows}.")
        if (clean[list(REQUIRED_COLUMNS - {"Quantity"})] == "").any(axis=None):
            raise ValueError("Hay celdas vacías en columnas obligatorias.")

        clean["Quantity"] = clean["Quantity"].astype(int)
        return clean

    def import_dataframe(self, df: pd.DataFrame) -> int:
        clean = self.validate_dataframe(df)
        imported_articles = 0
        with self.connect() as conn:
            for remito, receipt_group in clean.groupby("Remito", sort=False):
                first = receipt_group.iloc[0]
                conn.execute(
                    """
                    INSERT INTO remitos(remito, empresa, proveedor, factura)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(remito) DO UPDATE SET
                        empresa=excluded.empresa,
                        proveedor=excluded.proveedor,
                        factura=excluded.factura
                    """,
                    (remito, first["Empresa"], first["Proveedor"], first["Factura"]),
                )
                for article, article_group in receipt_group.groupby("Family", sort=False):
                    first_article = article_group.iloc[0]
                    # Reimporting an article intentionally replaces its stock and prior work.
                    conn.execute(
                        "DELETE FROM asignaciones WHERE remito=? AND articulo=?",
                        (remito, article),
                    )
                    conn.execute(
                        """
                        INSERT INTO articulos(remito, articulo, descripcion, marca, estado)
                        VALUES (?, ?, ?, ?, 'no_trabajado')
                        ON CONFLICT(remito, articulo) DO UPDATE SET
                            descripcion=excluded.descripcion,
                            marca=excluded.marca,
                            estado='no_trabajado'
                        """,
                        (remito, article, first_article["Descripcion"], first_article["Marca"]),
                    )
                    conn.execute(
                        "DELETE FROM stock_inicial WHERE remito=? AND articulo=?",
                        (remito, article),
                    )
                    totals = article_group.groupby("Size", as_index=False)["Quantity"].sum()
                    conn.executemany(
                        "INSERT INTO stock_inicial(remito, articulo, talle, cantidad) VALUES (?, ?, ?, ?)",
                        [(remito, article, str(row.Size), int(row.Quantity)) for row in totals.itertuples()],
                    )
                    imported_articles += 1
                self._update_receipt_status(conn, remito)
        return imported_articles

    def reset(self) -> None:
        with self.connect() as conn:
            conn.executescript("DELETE FROM asignaciones; DELETE FROM stock_inicial; DELETE FROM articulos; DELETE FROM remitos;")

    def summary(self) -> list[sqlite3.Row]:
        query = """
            SELECT a.remito, a.articulo, a.descripcion, a.marca,
                   SUM(s.cantidad) AS total, a.estado, r.empresa
            FROM articulos a
            JOIN remitos r ON r.remito=a.remito
            JOIN stock_inicial s ON s.remito=a.remito AND s.articulo=a.articulo
            GROUP BY a.remito, a.articulo, a.descripcion, a.marca, a.estado, r.empresa
            ORDER BY a.remito, a.articulo
        """
        with self.connect() as conn:
            return conn.execute(query).fetchall()

    def metrics(self) -> dict[str, int]:
        with self.connect() as conn:
            row = conn.execute(
                """
                SELECT
                    COUNT(DISTINCT r.remito) remitos,
                    COUNT(DISTINCT a.remito || '|' || a.articulo) articulos,
                    COALESCE(SUM(s.cantidad), 0) unidades,
                    COUNT(DISTINCT CASE WHEN a.estado='asignado' THEN a.remito || '|' || a.articulo END) asignados
                FROM articulos a
                JOIN remitos r ON r.remito=a.remito
                JOIN stock_inicial s ON s.remito=a.remito AND s.articulo=a.articulo
                """
            ).fetchone()
        return dict(row) if row else {"remitos": 0, "articulos": 0, "unidades": 0, "asignados": 0}

    def article_detail(self, remito: str, articulo: str) -> tuple[str, dict[str, int], list[sqlite3.Row]]:
        with self.connect() as conn:
            article = conn.execute(
                "SELECT descripcion FROM articulos WHERE remito=? AND articulo=?",
                (remito, articulo),
            ).fetchone()
            if article is None:
                raise LookupError("El artículo seleccionado ya no existe.")
            stock_rows = conn.execute(
                "SELECT talle, cantidad FROM stock_inicial WHERE remito=? AND articulo=? ORDER BY talle",
                (remito, articulo),
            ).fetchall()
            assignments = conn.execute(
                "SELECT banner, local, talle, cantidad FROM asignaciones WHERE remito=? AND articulo=?",
                (remito, articulo),
            ).fetchall()
        return article[0], {str(r[0]): int(r[1]) for r in stock_rows}, assignments

    def save_assignment(
        self,
        remito: str,
        articulo: str,
        banner: str,
        assignments: Iterable[tuple[str, str, int]],
        warehouse: str,
    ) -> None:
        assignments = list(assignments)
        _, stock, _ = self.article_detail(remito, articulo)
        totals = {size: 0 for size in stock}
        for local, size, quantity in assignments:
            if size not in totals:
                raise ValueError(f"Talle desconocido: {size}")
            if quantity < 0:
                raise ValueError("Las cantidades no pueden ser negativas.")
            totals[size] += quantity
        exceeded = [size for size, total in totals.items() if total > stock[size]]
        if exceeded:
            raise ValueError("Se excedió el stock en: " + ", ".join(exceeded))

        completed = [(local, size, qty) for local, size, qty in assignments if qty > 0]
        for size, total in totals.items():
            remainder = stock[size] - total
            if remainder:
                completed.append((warehouse, size, remainder))

        with self.connect() as conn:
            conn.execute("DELETE FROM asignaciones WHERE remito=? AND articulo=?", (remito, articulo))
            conn.executemany(
                "INSERT INTO asignaciones(remito, articulo, banner, local, talle, cantidad) VALUES (?, ?, ?, ?, ?, ?)",
                [(remito, articulo, banner, local, size, qty) for local, size, qty in completed],
            )
            conn.execute(
                "UPDATE articulos SET estado='asignado' WHERE remito=? AND articulo=?",
                (remito, articulo),
            )
            self._update_receipt_status(conn, remito)

    @staticmethod
    def _update_receipt_status(conn: sqlite3.Connection, remito: str) -> None:
        counts = conn.execute(
            """SELECT COUNT(*) total,
                      SUM(CASE WHEN estado='asignado' THEN 1 ELSE 0 END) done
               FROM articulos WHERE remito=?""",
            (remito,),
        ).fetchone()
        total, done = int(counts[0] or 0), int(counts[1] or 0)
        status = "no_trabajado" if done == 0 else ("asignado" if done == total else "parcial")
        conn.execute("UPDATE remitos SET estado=? WHERE remito=?", (status, remito))

    def distribution(self) -> list[sqlite3.Row]:
        with self.connect() as conn:
            return conn.execute(
                """SELECT remito, articulo, banner, local, talle, cantidad
                   FROM asignaciones ORDER BY remito, articulo, local, talle"""
            ).fetchall()

    def report(self, article: str = "", local: str = "") -> pd.DataFrame:
        query = """
            SELECT a.remito AS Remito, a.articulo AS Articulo, a.descripcion AS Descripcion,
                   a.marca AS Marca, COALESCE(x.banner, '-') AS Banner,
                   COALESCE(x.local, 'Pendiente de separación') AS Local,
                   COALESCE(x.talle, '-') AS Talle, COALESCE(x.cantidad, 0) AS Cantidad,
                   a.estado AS Estado
            FROM articulos a
            LEFT JOIN asignaciones x ON x.remito=a.remito AND x.articulo=a.articulo
            WHERE a.articulo LIKE ? AND COALESCE(x.local, '') LIKE ?
            ORDER BY a.remito, a.articulo, x.local, x.talle
        """
        with self.connect() as conn:
            return pd.read_sql_query(query, conn, params=(f"%{article.strip()}%", f"%{local.strip()}%"))
