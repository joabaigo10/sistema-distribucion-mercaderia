from pathlib import Path
import tempfile

import pandas as pd
import pytest

from data_store import DistributionStore


def demo_dataframe() -> pd.DataFrame:
    return pd.DataFrame([
        {"Remito": "R-1", "Family": "A-1", "Size": "S", "Quantity": 5, "Descripcion": "Remera", "Empresa": "DEMO", "Proveedor": "P-1", "Factura": "F-1", "Marca": "TEST"},
        {"Remito": "R-1", "Family": "A-1", "Size": "M", "Quantity": 7, "Descripcion": "Remera", "Empresa": "DEMO", "Proveedor": "P-1", "Factura": "F-1", "Marca": "TEST"},
    ])


def make_store() -> DistributionStore:
    directory = tempfile.TemporaryDirectory()
    store = DistributionStore(Path(directory.name) / "test.db")
    store._temporary_directory = directory  # type: ignore[attr-defined]
    return store


def test_import_and_metrics() -> None:
    store = make_store()
    assert store.import_dataframe(demo_dataframe()) == 1
    assert store.metrics() == {"remitos": 1, "articulos": 1, "unidades": 12, "asignados": 0}


def test_assignment_adds_remainder_to_warehouse() -> None:
    store = make_store()
    store.import_dataframe(demo_dataframe())
    store.save_assignment("R-1", "A-1", "DEPORTE", [("LOCAL CENTRO", "S", 3)], "ALMACÉN DEPORTE")
    rows = [tuple(row) for row in store.distribution()]
    assert ("R-1", "A-1", "DEPORTE", "LOCAL CENTRO", "S", 3) in rows
    assert ("R-1", "A-1", "DEPORTE", "ALMACÉN DEPORTE", "S", 2) in rows
    assert ("R-1", "A-1", "DEPORTE", "ALMACÉN DEPORTE", "M", 7) in rows


def test_cannot_exceed_stock() -> None:
    store = make_store()
    store.import_dataframe(demo_dataframe())
    with pytest.raises(ValueError, match="excedió"):
        store.save_assignment("R-1", "A-1", "DEPORTE", [("LOCAL CENTRO", "S", 6)], "ALMACÉN DEPORTE")


def test_reimport_clears_previous_assignment() -> None:
    store = make_store()
    df = demo_dataframe()
    store.import_dataframe(df)
    store.save_assignment("R-1", "A-1", "DEPORTE", [("LOCAL CENTRO", "S", 3)], "ALMACÉN DEPORTE")
    store.import_dataframe(df)
    assert store.distribution() == []
    assert store.summary()[0][5] == "no_trabajado"
