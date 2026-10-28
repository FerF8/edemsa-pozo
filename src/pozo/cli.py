"""Comandos para armar el Excel de una factura o de todas las de la carpeta."""

from __future__ import annotations

import argparse
import csv
import sys
from decimal import Decimal
from pathlib import Path

from pozo.excel_out import Sesion, escribir_reparto
from pozo.factura import ParseError, parse_factura
from pozo.reparto import Lectura, repartir

PERSONAS = ("Persona 1", "Persona 2", "Persona 3")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Divide la factura de EDEMSA del pozo entre 3 personas")
    sub = parser.add_subparsers(dest="comando", required=True)

    uno = sub.add_parser("dividir", help="Arma el Excel de un PDF")
    uno.add_argument("pdf", type=Path)
    uno.add_argument("-o", "--salida", type=Path)
    uno.add_argument("--lecturas", type=Path, help="CSV con persona,alta_inicio,alta_fin,baja_inicio,baja_fin")

    lote = sub.add_parser("lote", help="Arma un Excel por cada PDF de la carpeta")
    lote.add_argument("--carpeta", type=Path, default=Path("."))
    lote.add_argument("--salida", type=Path, default=Path("repartos"))

    args = parser.parse_args(argv)
    try:
        if args.comando == "dividir":
            return _dividir(args.pdf, args.salida, args.lecturas)
        return _lote(args.carpeta, args.salida)
    except (ParseError, ValueError) as error:
        print(error, file=sys.stderr)
        return 1


def _dividir(pdf: Path, salida: Path | None, lecturas_csv: Path | None) -> int:
    factura = parse_factura(pdf)
    sesiones = _leer_csv(lecturas_csv) if lecturas_csv else []
    if salida is None:
        anio, mes = factura.mes.split("/")[1], factura.mes.split("/")[0]
        salida = Path("repartos") / f"{anio}-{mes}.xlsx"
    personas = _personas(sesiones)
    escribir_reparto(factura, salida, personas=personas, sesiones=sesiones)
    _imprimir(factura, personas, sesiones, salida)
    return 0


def _lote(carpeta: Path, salida: Path) -> int:
    pdfs = sorted(carpeta.glob("*.pdf"), key=_orden_pdf)
    if not pdfs:
        raise ParseError(f"No hay facturas en {carpeta}")
    for pdf in pdfs:
        mes = pdf.stem.split("_", 1)[1]  # 11-2025
        numero, anio = mes.split("-")
        destino = salida / f"{anio}-{int(numero):02d}.xlsx"
        csv_path = carpeta / "lecturas" / f"{anio}-{int(numero):02d}.csv"
        _dividir(pdf, destino, csv_path if csv_path.is_file() else None)
    return 0


def _orden_pdf(path: Path) -> tuple[int, int]:
    trozo = path.stem.split("_", 1)[-1]
    numero, anio = trozo.split("-", 1)
    return int(anio), int(numero)


def _leer_csv(path: Path) -> list[Sesion]:
    sesiones = []
    with path.open(newline="", encoding="utf-8") as archivo:
        lector = csv.DictReader(archivo)
        for fila in lector:
            if not (fila.get("persona") or "").strip():
                continue
            sesiones.append(
                Sesion(
                    persona=fila["persona"].strip(),
                    alta_inicio=Decimal(fila["alta_inicio"]),
                    alta_fin=Decimal(fila["alta_fin"]),
                    baja_inicio=Decimal(fila["baja_inicio"]),
                    baja_fin=Decimal(fila["baja_fin"]),
                )
            )
    return sesiones


def _personas(sesiones: list[Sesion]) -> tuple[str, str, str]:
    vistos = []
    for sesion in sesiones:
        if sesion.persona not in vistos:
            vistos.append(sesion.persona)
    nombres = list(PERSONAS)
    for indice, nombre in enumerate(vistos[:3]):
        nombres[indice] = nombre
    return nombres[0], nombres[1], nombres[2]


def _imprimir(factura, personas, sesiones, salida: Path) -> None:
    lecturas = []
    for nombre in personas:
        alta = sum(
            (s.alta_fin - s.alta_inicio for s in sesiones if s.persona == nombre),
            Decimal("0"),
        )
        baja = sum(
            (s.baja_fin - s.baja_inicio for s in sesiones if s.persona == nombre),
            Decimal("0"),
        )
        lecturas.append(Lectura(nombre, alta, baja))
    resultado = repartir(factura, lecturas)
    print(f"{factura.mes}  factura {factura.numero}  total ${factura.total}")
    for parte in resultado.personas:
        print(f"  {parte.nombre}: ${parte.total}")
    if resultado.sin_asignar.total != 0:
        print(f"  Sin asignar: ${resultado.sin_asignar.total}")
    if resultado.advertencia:
        print(f"  {resultado.advertencia}")
    print(f"  {salida}")


if __name__ == "__main__":
    raise SystemExit(main())
