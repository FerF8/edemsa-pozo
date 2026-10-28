"""Helpers que alimentan la app Streamlit: lecturas, tabla y Excel en memoria."""

from decimal import Decimal
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook

from pozo.excel_out import Sesion, escribir_reparto
from pozo.factura import parse_factura_texto
from pozo.reparto import Lectura, repartir
from pozo.ui_state import (
    filas_reparto,
    lecturas_desde_sesiones,
    resumen_totales,
    workbook_a_bytes,
)
from tests.formula_eval import Libro

D = Decimal
FIXTURES = Path(__file__).resolve().parent / "fixtures"


def _noviembre():
    return parse_factura_texto((FIXTURES / "factura_11-2025.txt").read_text(encoding="utf-8"))


def test_suma_varias_sesiones_de_la_misma_persona():
    personas = ("Persona 1", "Persona 2", "Persona 3")
    sesiones = [
        Sesion("Persona 1", D("100"), D("110"), D("200"), D("220")),
        Sesion("Persona 1", D("110"), D("115"), D("220"), D("230")),
        Sesion("Persona 2", D("0"), D("5"), D("0"), D("0")),
    ]
    lecturas = lecturas_desde_sesiones(personas, sesiones)
    assert lecturas == [
        Lectura("Persona 1", D("15"), D("30")),
        Lectura("Persona 2", D("5"), D("0")),
        Lectura("Persona 3", D("0"), D("0")),
    ]


def test_varias_lecturas_solo_alta_o_solo_baja_se_suman():
    """Opción B: varias filas por persona; se puede dejar alta o baja en cero."""
    personas = ("Persona 1", "Persona 2", "Persona 3")
    sesiones = [
        Sesion("Persona 1", D("10"), D("20"), D("0"), D("0")),
        Sesion("Persona 1", D("20"), D("25"), D("0"), D("0")),
        Sesion("Persona 1", D("0"), D("0"), D("100"), D("140")),
        Sesion("Persona 1", D("0"), D("0"), D("140"), D("155")),
    ]
    lecturas = lecturas_desde_sesiones(personas, sesiones)
    assert lecturas[0] == Lectura("Persona 1", D("15"), D("55"))


def test_resumen_incluye_sin_asignar_y_cierra():
    factura = _noviembre()
    resultado = repartir(
        factura,
        [
            Lectura("Persona 1", D("134"), D("318")),
            Lectura("Persona 2", D("0"), D("0")),
            Lectura("Persona 3", D("0"), D("0")),
        ],
    )
    resumen = resumen_totales(resultado)
    assert resumen[-1]["persona"] == "Sin asignar"
    assert resumen[-1]["total"] == resultado.sin_asignar.total
    assert sum(fila["total"] for fila in resumen) == factura.total


def test_filas_reparto_tienen_todas_las_columnas():
    factura = _noviembre()
    resultado = repartir(
        factura,
        [
            Lectura("Persona 1", D("134"), D("318")),
            Lectura("Persona 2", D("0"), D("0")),
            Lectura("Persona 3", D("0"), D("0")),
        ],
    )
    filas = filas_reparto(resultado)
    assert filas[0]["concepto"] == "Energía alta"
    assert "Persona 1" in filas[0]
    assert filas[-1]["concepto"] == "Total a pagar"
    assert filas[-1]["Total"] == factura.total


def test_workbook_en_memoria_cierra_contra_la_factura(tmp_path):
    factura = _noviembre()
    destino = tmp_path / "tmp.xlsx"
    escribir_reparto(
        factura,
        destino,
        personas=("Persona 1", "Persona 2", "Persona 3"),
        sesiones=[Sesion("Persona 1", D("77140"), D("77274"), D("96892"), D("97210"))],
    )
    contenido = workbook_a_bytes(destino)
    assert isinstance(contenido, bytes)
    assert len(contenido) > 1000

    libro = Libro(load_workbook(BytesIO(contenido)))
    assert Decimal(str(libro.celda("Reparto", "B19"))) == D("0")
