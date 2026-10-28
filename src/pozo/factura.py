"""Lee una factura de EDEMSA (tarifa de riego agrícola) y clasifica cada concepto.

Criterio de reparto:
- La suma de las tres personas más lo que siga sin lectura tiene que ser el total de la factura.
- El IVA es el 27 % de responsable inscripto que figura en la factura.
- El subsidio nacional y la Ley 9083 se reparten con la energía.
- Mora, aviso de suspensión, reanudación, bonificaciones y devoluciones van en tres partes iguales.
- Los cargos fijos (comercialización, uso de red, exceso de potencia, CAP municipal,
  impuesto de sellos y redondeos) también van en tres partes iguales.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

import pdfplumber

# Orden usado por el reparto y por las fórmulas del Excel.
GRUPOS = (
    "energia_alta",
    "energia_baja",
    "subsidio_nacional",
    "fijo",
    "penal",
    "iibb",
    "ley9083",
    "iva",
    "tasa",
    "ccce",
    "otros",
)

GRUPOS_ELECTRICOS = {
    "energia_alta",
    "energia_baja",
    "subsidio_nacional",
    "fijo",
    "penal",
    "iibb",
    "ley9083",
}

MONEY_RE = re.compile(r"-?(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}(?!\d)")

# La primera regla que coincida se queda con la línea. Las más específicas van primero.
REGLAS: tuple[tuple[str, str, re.Pattern[str]], ...] = (
    ("energia_alta", "Energía alta excedida", re.compile(r"Energ[ií]a Alta Excedida", re.I)),
    ("energia_alta", "Energía alta", re.compile(r"Energ[ií]a Alta kWh", re.I)),
    ("energia_baja", "Energía baja excedida", re.compile(r"Energ[ií]a Baja Excedida", re.I)),
    ("energia_baja", "Energía baja", re.compile(r"Energ[ií]a Baja kWh", re.I)),
    ("subsidio_nacional", "Subsidio Estado Nacional", re.compile(r"Subsidio Estado Nacional", re.I)),
    ("ley9083", "Subsidio Ley 9083", re.compile(r"Subsidio Beneficio Ley 9083", re.I)),
    ("fijo", "Cargo de comercialización", re.compile(r"Cargo Comercializaci[oó]n", re.I)),
    ("fijo", "Uso de red", re.compile(r"Uso de Red kW", re.I)),
    ("fijo", "Exceso de potencia", re.compile(r"Exceso de Potencia", re.I)),
    ("iibb", "Ingresos brutos", re.compile(r"Impuestos Provinciales", re.I)),
    ("penal", "Recargo por mora", re.compile(r"Recargos por mora", re.I)),
    ("penal", "Aviso de suspensión", re.compile(r"Aviso de Suspensi[oó]n", re.I)),
    ("penal", "Reanudación del suministro", re.compile(r"Reanudaci[oó]n del Suministro", re.I)),
    ("otros", "Bonificación fuerza mayor", re.compile(r"Bonif\.[^\n]*fuerza mayor", re.I)),
    ("otros", "Bonificación deficiencias técnicas", re.compile(r"Bonif\.[^\n]*deficiencias", re.I)),
    ("otros", "Devolución por pago duplicado", re.compile(r"Devoluci[oó]n por Pago Duplicado", re.I)),
    ("otros", "Ajuste de redondeo", re.compile(r"Ajuste redondeo", re.I)),
    ("otros", "Redondeo", re.compile(r"(?<![A-Za-z])Redondeo\b", re.I)),
    ("iva", "IVA", re.compile(r"I\.V\.A\.", re.I)),
    ("tasa", "Tasa de fiscalización y control", re.compile(r"Tasa fisc", re.I)),
    ("ccce", "CCCE", re.compile(r"CCCE art", re.I)),
    ("otros", "Impuesto de sellos", re.compile(r"Imp\.sellos", re.I)),
    ("otros", "CAP municipal", re.compile(r"CAP Mun", re.I)),
)


class ParseError(Exception):
    """La factura no se pudo leer o no cierran los importes."""


@dataclass
class Concepto:
    nombre: str
    grupo: str
    importe: Decimal


@dataclass
class Factura:
    nic: str
    numero: str
    mes: str
    periodo_desde: str
    periodo_hasta: str
    vencimiento: str
    total: Decimal
    kwh_alta: Decimal
    kwh_baja: Decimal
    conceptos: list[Concepto]
    alta_inicio: Decimal
    alta_fin: Decimal
    baja_inicio: Decimal
    baja_fin: Decimal
    potencia_contratada: str = ""
    potencia_maxima: str = ""
    subtotal_electricos: Decimal = field(default_factory=lambda: Decimal("0"))


def parse_money(token: str) -> Decimal:
    return Decimal(token.replace(".", "").replace(",", ".")).quantize(Decimal("0.01"))


def last_money(line: str) -> Decimal | None:
    found = MONEY_RE.findall(line)
    if not found:
        return None
    return parse_money(found[-1])


def _first(pattern: str, text: str, label: str) -> str:
    match = re.search(pattern, text, flags=re.I)
    if not match:
        raise ParseError(f"No encontré {label} en la factura")
    return match.group(1)


def parse_factura(path: Path | str) -> Factura:
    pdf_path = Path(path)
    if not pdf_path.is_file():
        raise ParseError(f"No se encontró el PDF: {pdf_path}")

    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join((page.extract_text() or "") for page in pdf.pages)
    if not text.strip():
        raise ParseError(f"El PDF no tiene texto: {pdf_path.name}")
    return parse_factura_texto(text)


def parse_factura_texto(text: str) -> Factura:
    """Interpreta el texto ya extraído de una factura (PDF o fixture de prueba)."""
    if not text.strip():
        raise ParseError("El texto de la factura está vacío")

    plano = re.sub(r"\s+", " ", text)
    nic = _first(r"C[ÓO]DIGO PAGO ELECTR[ÓO]NICO\s+(\d+)", plano, "el NIC")
    numero = _first(r"FACTURA N[°º]\s*:\s*(\S+)", plano, "el número de factura")
    mes = _first(r"MES\s+(\d{2}/\d{4})", plano, "el mes")
    total = parse_money(_first(r"TOTAL\s+\$\s+([-0-9.]+,\d{2})", plano, "el total"))
    subtotal = parse_money(
        _first(r"1\.\s*Conceptos El[ée]ctricos\s+\$\s+([-0-9.]+,\d{2})", plano, "el subtotal eléctrico")
    )

    periodo = re.search(
        r"(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+\d+\s+Baja",
        plano,
    )
    if not periodo:
        raise ParseError("No encontré el período facturado")
    vencimiento_match = re.search(
        r"Fecha emisi[oó]n:\s*Vencimiento actual:\s*\d{2}/\d{2}/\d{4}\s+(\d{2}/\d{2}/\d{4})",
        plano,
        flags=re.I,
    )
    if not vencimiento_match:
        raise ParseError("No encontré el vencimiento")
    vence = vencimiento_match.group(1)

    alta = re.search(
        r"Alta 14 a 23hs\s+([-0-9.]+,\d{2})\s+([-0-9.]+,\d{2})\s+([-0-9.]+,\d{2})\s+([-0-9.]+,\d{2})",
        plano,
    )
    baja = re.search(
        r"Baja 23 a 14hs\s+([-0-9.]+,\d{2})\s+([-0-9.]+,\d{2})\s+([-0-9.]+,\d{2})\s+([-0-9.]+,\d{2})",
        plano,
    )
    if not alta or not baja:
        raise ParseError("No encontré las lecturas de alta y baja del medidor")

    potencia = re.search(
        r"Potencia Contratada:\s*([-0-9.]+)\s+M[aá]xima Registrada:\s*([-0-9.]+)",
        plano,
        flags=re.I,
    )

    conceptos: list[Concepto] = []
    for line in text.splitlines():
        for grupo, nombre, pattern in REGLAS:
            if pattern.search(line):
                importe = last_money(line)
                if importe is None:
                    raise ParseError(f"Encontré «{nombre}» pero no el importe: {line}")
                conceptos.append(Concepto(nombre, grupo, importe))
                break

    clasificado = sum((c.importe for c in conceptos), Decimal("0.00"))
    if clasificado != total:
        detalle = "\n".join(f"  {c.grupo}: {c.nombre} {c.importe}" for c in conceptos)
        raise ParseError(
            f"Los conceptos suman {clasificado} y la factura dice {total}. Falta o sobra algún renglón.\n{detalle}"
        )

    electrico = sum((c.importe for c in conceptos if c.grupo in GRUPOS_ELECTRICOS), Decimal("0.00"))
    if electrico != subtotal:
        detalle = "\n".join(
            f"  {c.grupo}: {c.nombre} {c.importe}" for c in conceptos if c.grupo in GRUPOS_ELECTRICOS
        )
        raise ParseError(
            f"Los conceptos eléctricos suman {electrico} y la factura dice {subtotal}.\n{detalle}"
        )

    return Factura(
        nic=nic,
        numero=numero,
        mes=mes,
        periodo_desde=periodo.group(1),
        periodo_hasta=periodo.group(2),
        vencimiento=vence,
        total=total,
        kwh_alta=parse_money(alta.group(4)),
        kwh_baja=parse_money(baja.group(4)),
        conceptos=conceptos,
        alta_inicio=parse_money(alta.group(1)),
        alta_fin=parse_money(alta.group(2)),
        baja_inicio=parse_money(baja.group(1)),
        baja_fin=parse_money(baja.group(2)),
        potencia_contratada=potencia.group(1) if potencia else "",
        potencia_maxima=potencia.group(2) if potencia else "",
        subtotal_electricos=subtotal,
    )
