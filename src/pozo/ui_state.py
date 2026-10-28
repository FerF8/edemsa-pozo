"""Estado y tablas para la app Streamlit.

Cargar PDF de EDEMSA, cargar varias lecturas de alta y baja por persona,
ver totales al instante y exportar el Excel con las mismas fórmulas.
"""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path

from pozo.excel_out import ETIQUETAS, Sesion
from pozo.factura import GRUPOS
from pozo.reparto import Lectura, Parte, Reparto

CERO = Decimal("0")


def lecturas_desde_sesiones(
    personas: tuple[str, str, str],
    sesiones: list[Sesion],
) -> list[Lectura]:
    lecturas = []
    for nombre in personas:
        alta = sum(
            (sesion.alta_fin - sesion.alta_inicio for sesion in sesiones if sesion.persona == nombre),
            CERO,
        )
        baja = sum(
            (sesion.baja_fin - sesion.baja_inicio for sesion in sesiones if sesion.persona == nombre),
            CERO,
        )
        lecturas.append(Lectura(nombre, alta, baja))
    return lecturas


def resumen_totales(resultado: Reparto) -> list[dict]:
    filas = [
        {
            "persona": parte.nombre,
            "total": parte.total,
            "kwh_alta": None,
            "kwh_baja": None,
        }
        for parte in resultado.personas
    ]
    filas.append(
        {
            "persona": resultado.sin_asignar.nombre,
            "total": resultado.sin_asignar.total,
            "kwh_alta": resultado.kwh_sin_asignar_alta,
            "kwh_baja": resultado.kwh_sin_asignar_baja,
        }
    )
    return filas


def filas_reparto(resultado: Reparto) -> list[dict]:
    columnas = [parte.nombre for parte in resultado.personas] + [resultado.sin_asignar.nombre]
    partes: list[Parte] = list(resultado.personas) + [resultado.sin_asignar]
    filas = []
    for grupo in GRUPOS:
        fila = {"concepto": ETIQUETAS[grupo]}
        for nombre, parte in zip(columnas, partes, strict=True):
            fila[nombre] = parte.importes[grupo]
        fila["Total"] = sum((parte.importes[grupo] for parte in partes), CERO)
        filas.append(fila)

    total = {"concepto": "Total a pagar"}
    for nombre, parte in zip(columnas, partes, strict=True):
        total[nombre] = parte.total
    total["Total"] = sum((parte.total for parte in partes), CERO)
    filas.append(total)
    return filas


def workbook_a_bytes(ruta: Path | str) -> bytes:
    return Path(ruta).read_bytes()
