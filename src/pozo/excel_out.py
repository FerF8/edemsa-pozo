"""Arma el Excel del reparto. Los resultados son fórmulas: si se corrige una lectura, se recalculan."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from pozo.factura import GRUPOS, Factura

FILA_CONCEPTOS = 32
FILA_GRUPOS = 19  # Factura!F19 es el total del primer grupo
FILAS_SESION = range(4, 19)
AZUL = Font(name="Arial", size=11, color="0000FF")
NEGRO = Font(name="Arial", size=11, color="000000")
NEGRO_NEGRITA = Font(name="Arial", size=11, color="000000", bold=True)
TITULO = Font(name="Arial", size=16, color="000000", bold=True)
AMARILLO = PatternFill("solid", fgColor="FFFF00")
GRIS = PatternFill("solid", fgColor="F2F2F2")
ENCABEZADO = PatternFill("solid", fgColor="1F4E79")
ENCABEZADO_FUENTE = Font(name="Arial", size=11, color="FFFFFF", bold=True)
PESOS = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
ENTERO = '#,##0.00'
FINO = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
ETIQUETAS = {
    "energia_alta": "Energía alta",
    "energia_baja": "Energía baja",
    "subsidio_nacional": "Subsidio nacional",
    "fijo": "Cargos fijos",
    "penal": "Mora, suspensión y reanudación",
    "iibb": "Ingresos brutos",
    "ley9083": "Subsidio Ley 9083",
    "iva": "IVA",
    "tasa": "Tasa de fiscalización y control",
    "ccce": "CCCE",
    "otros": "Otros (CAP, sellos, redondeos, bonificaciones)",
}


@dataclass
class Sesion:
    persona: str
    alta_inicio: Decimal
    alta_fin: Decimal
    baja_inicio: Decimal
    baja_fin: Decimal


def _num(value: Decimal) -> float:
    return float(value)


def _input(cell, value) -> None:
    cell.value = value
    cell.font = AZUL
    cell.fill = AMARILLO
    cell.border = FINO


def _formula(cell, formula: str, pesos: bool = False) -> None:
    cell.value = formula if formula.startswith("=") else f"={formula}"
    cell.font = NEGRO
    cell.border = FINO
    if pesos:
        cell.number_format = PESOS


def _texto(cell, value, negrita: bool = False) -> None:
    cell.value = value
    cell.font = NEGRO_NEGRITA if negrita else NEGRO
    cell.alignment = Alignment(wrap_text=True, vertical="center")


def _proporcional(fila: int, monto: str, pesos: list[str], suma: str, sin_activo: str) -> tuple[str, str, str, str]:
    """Cuatro fórmulas. Los centavos cierran en Sin asignar si hay consumo sin dueño; si no, en la tercera persona."""
    persona_1 = f'IF({suma}=0,0,ROUND(({monto})*({pesos[0]})/({suma}),2))'
    persona_2 = f'IF({suma}=0,0,ROUND(({monto})*({pesos[1]})/({suma}),2))'
    persona_3 = (
        f'IF({suma}=0,0,IF({sin_activo},ROUND(({monto})*({pesos[2]})/({suma}),2),({monto})-B{fila}-C{fila}))'
    )
    sin_asignar = f'IF({suma}=0,({monto}),IF({sin_activo},({monto})-B{fila}-C{fila}-D{fila},0))'
    return persona_1, persona_2, persona_3, sin_asignar


def _tercios(fila: int, monto: str) -> tuple[str, str, str, str]:
    return (
        f'ROUND(({monto})/3,2)',
        f'ROUND(({monto})/3,2)',
        f'({monto})-B{fila}-C{fila}',
        "0",
    )


def escribir_reparto(
    factura: Factura,
    destino: Path | str,
    personas: tuple[str, str, str] = ("Persona 1", "Persona 2", "Persona 3"),
    sesiones: list[Sesion] | None = None,
) -> Path:
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    sesiones = sesiones or []
    if len(sesiones) > len(list(FILAS_SESION)):
        raise ValueError("Hay más sesiones que filas en la hoja Lecturas")

    libro = Workbook()
    libro.calculation.calcMode = "auto"
    libro.calculation.fullCalcOnLoad = True

    factura_hoja = libro.active
    factura_hoja.title = "Factura"
    lecturas = libro.create_sheet("Lecturas")
    reparto = libro.create_sheet("Reparto")
    uso = libro.create_sheet("Uso", 0)

    _hoja_uso(uso)
    _hoja_factura(factura_hoja, factura)
    _hoja_lecturas(lecturas, personas, sesiones)
    _hoja_reparto(reparto)

    for hoja in libro.worksheets:
        hoja.sheet_view.showGridLines = False
        hoja.page_setup.orientation = "landscape"
        hoja.page_setup.fitToPage = True
        hoja.page_setup.fitToWidth = 1
        hoja.page_setup.fitToHeight = 1
        hoja.sheet_properties.pageSetUpPr.fitToPage = True
        hoja.oddHeader.left.text = "Pozo compartido · EDEMSA"
        hoja.oddFooter.left.text = "EDEMSA"

    libro.save(destino)
    return destino


def _hoja_uso(hoja) -> None:
    hoja["A1"] = "Cómo se divide la factura del pozo"
    hoja["A1"].font = TITULO
    hoja.merge_cells("A1:B1")
    lineas = [
        "1. En Lecturas, completá las celdas amarillas. Una fila por cada uso del pozo: la misma persona puede aparecer varias veces.",
        "2. Si en un tramo solo hubo alta o solo baja, dejá la otra banda en cero. Los tres nombres (filas 22 a 24) tienen que coincidir con los de cada sesión.",
        "3. El total de cada persona está en Reparto, fila Total a pagar. La diferencia contra la factura tiene que quedar en cero.",
        "4. Sin asignar no es una cuarta persona. Es el consumo del medidor que todavía nadie anotó. Los cargos fijos igual se dividen en tres.",
        "5. Una diferencia de hasta 2 kilovatios hora contra el medidor se toma como redondeo de la factura y no arma esa columna.",
        "6. El subsidio nacional se reparte según los pesos de energía de cada uno. La Ley 9083 sigue la misma base, sin la mora ni los avisos de corte.",
        "7. El IVA es el 27 % de la factura (responsable inscripto). Mora, suspensión, reanudación, bonificaciones y devoluciones van en tres partes iguales.",
        "8. Los cargos fijos (comercialización, uso de red, exceso de potencia, CAP municipal, sellos y redondeos) van en tres partes iguales. El medidor no dice quién provocó el pico de potencia.",
        "9. Los centavos de cierre quedan en la tercera persona. Si hay consumo sin asignar, esos centavos quedan en esa columna.",
        "10. Si corregís un importe amarillo de Factura, no vuelvas a generar este archivo: el programa lo pisa con el PDF.",
    ]
    for indice, linea in enumerate(lineas, start=3):
        hoja.cell(indice, 1, linea).font = NEGRO
        hoja.row_dimensions[indice].height = 32
        hoja.merge_cells(start_row=indice, start_column=1, end_row=indice, end_column=2)
    hoja.column_dimensions["A"].width = 120
    hoja.column_dimensions["B"].width = 20
    hoja.row_dimensions[1].height = 24
    hoja.freeze_panes = "A3"
    hoja.print_title_rows = "1:1"
    hoja.page_setup.orientation = "landscape"
    hoja.sheet_view.zoomScale = 120


def _hoja_factura(hoja, factura: Factura) -> None:
    campos = [
        (1, "NIC", factura.nic, False),
        (2, "Factura", factura.numero, False),
        (3, "Mes", factura.mes, False),
        (4, "Total a pagar", _num(factura.total), True),
        (5, "kWh alta del medidor", _num(factura.kwh_alta), False),
        (6, "kWh baja del medidor", _num(factura.kwh_baja), False),
        (7, "Período desde", factura.periodo_desde, False),
        (8, "Período hasta", factura.periodo_hasta, False),
        (9, "Vencimiento", factura.vencimiento, False),
        (10, "Lectura alta inicial", _num(factura.alta_inicio), False),
        (11, "Lectura alta final", _num(factura.alta_fin), False),
        (12, "Lectura baja inicial", _num(factura.baja_inicio), False),
        (13, "Lectura baja final", _num(factura.baja_fin), False),
        (14, "Potencia contratada (kW)", factura.potencia_contratada, False),
        (15, "Potencia máxima registrada (kW)", factura.potencia_maxima, False),
    ]
    for fila, etiqueta, valor, es_pesos in campos:
        _texto(hoja.cell(fila, 1), etiqueta, negrita=True)
        celda = hoja.cell(fila, 2, valor)
        celda.font = NEGRO
        celda.border = FINO
        if es_pesos:
            celda.number_format = PESOS
        elif isinstance(valor, float):
            celda.number_format = ENTERO
    hoja["B5"].fill = AMARILLO
    hoja["B5"].font = AZUL
    hoja["B6"].fill = AMARILLO
    hoja["B6"].font = AZUL
    hoja["B5"].comment = Comment(
        "Sale del PDF. Corregilo solo si la lectura automática no coincide con la factura.",
        "pozo",
    )

    _texto(hoja.cell(18, 5), "Clave", negrita=True)
    _texto(hoja.cell(18, 6), "Importe del grupo", negrita=True)
    for desplazamiento, grupo in enumerate(GRUPOS):
        fila = FILA_GRUPOS + desplazamiento
        hoja.cell(fila, 5, grupo).font = NEGRO
        _formula(
            hoja.cell(fila, 6),
            f'SUMIF($B${FILA_CONCEPTOS}:$B$80,E{fila},$C${FILA_CONCEPTOS}:$C$80)',
            pesos=True,
        )

    encabezados = ("Concepto", "Grupo", "Importe")
    for columna, titulo in enumerate(encabezados, start=1):
        celda = hoja.cell(31, columna, titulo)
        celda.font = ENCABEZADO_FUENTE
        celda.fill = ENCABEZADO
    for indice, concepto in enumerate(factura.conceptos):
        fila = FILA_CONCEPTOS + indice
        _texto(hoja.cell(fila, 1), concepto.nombre)
        _input(hoja.cell(fila, 2), concepto.grupo)
        _input(hoja.cell(fila, 3), _num(concepto.importe))
        hoja.cell(fila, 3).number_format = PESOS

    hoja.column_dimensions["A"].width = 42
    hoja.column_dimensions["B"].width = 28
    hoja.column_dimensions["C"].width = 18
    hoja.column_dimensions["E"].width = 24
    hoja.column_dimensions["F"].width = 20
    hoja.freeze_panes = "A32"
    hoja.auto_filter.ref = f"A31:C{FILA_CONCEPTOS + max(len(factura.conceptos), 1) - 1}"
    hoja.print_title_rows = "1:1"
    hoja.page_setup.orientation = "landscape"
    hoja.oddHeader.center.text = f"Factura {factura.numero}"


def _hoja_lecturas(hoja, personas: tuple[str, str, str], sesiones: list[Sesion]) -> None:
    hoja["A1"] = "Lecturas del medidor"
    hoja["A1"].font = TITULO
    hoja.merge_cells("A1:G1")
    _texto(
        hoja.cell(2, 1),
        "Una fila por cada vez que alguien usa el pozo. Si ya tenés los kilovatios hora, poné inicio 0 y fin igual a ese consumo.",
    )
    hoja.merge_cells("A2:G2")
    hoja.row_dimensions[2].height = 30

    titulos = ("Persona", "Alta inicio", "Alta fin", "Baja inicio", "Baja fin", "kWh alta", "kWh baja")
    for columna, titulo in enumerate(titulos, start=1):
        celda = hoja.cell(3, columna, titulo)
        celda.font = ENCABEZADO_FUENTE
        celda.fill = ENCABEZADO
    for fila in FILAS_SESION:
        for columna in range(1, 6):
            _input(hoja.cell(fila, columna), None)
            if columna > 1:
                hoja.cell(fila, columna).number_format = ENTERO
        _formula(hoja.cell(fila, 6), f'IF(OR(B{fila}="",C{fila}=""),0,C{fila}-B{fila})')
        _formula(hoja.cell(fila, 7), f'IF(OR(D{fila}="",E{fila}=""),0,E{fila}-D{fila})')
        hoja.cell(fila, 6).number_format = ENTERO
        hoja.cell(fila, 7).number_format = ENTERO

    for indice, sesion in enumerate(sesiones):
        fila = 4 + indice
        _input(hoja.cell(fila, 1), sesion.persona)
        _input(hoja.cell(fila, 2), _num(sesion.alta_inicio))
        _input(hoja.cell(fila, 3), _num(sesion.alta_fin))
        _input(hoja.cell(fila, 4), _num(sesion.baja_inicio))
        _input(hoja.cell(fila, 5), _num(sesion.baja_fin))
        for columna in range(2, 6):
            hoja.cell(fila, columna).number_format = ENTERO

    validacion = DataValidation(type="list", formula1="=$A$22:$A$24", allow_blank=True)
    validacion.error = "Elegí una de las tres personas"
    validacion.errorTitle = "Persona"
    validacion.add("A4:A18")
    hoja.add_data_validation(validacion)

    _texto(hoja.cell(20, 1), "Personas que comparten los cargos fijos", negrita=True)
    for columna, titulo in enumerate(("Persona", "kWh alta", "kWh baja"), start=1):
        celda = hoja.cell(21, columna, titulo)
        celda.font = ENCABEZADO_FUENTE
        celda.fill = ENCABEZADO
    for indice, nombre in enumerate(personas):
        fila = 22 + indice
        _input(hoja.cell(fila, 1), nombre)
        _formula(hoja.cell(fila, 2), f'SUMIF($A$4:$A$18,A{fila},$F$4:$F$18)')
        _formula(hoja.cell(fila, 3), f'SUMIF($A$4:$A$18,A{fila},$G$4:$G$18)')
        hoja.cell(fila, 2).number_format = ENTERO
        hoja.cell(fila, 3).number_format = ENTERO

    _texto(hoja.cell(26, 1), "Sin asignar", negrita=True)
    # Sin lecturas, todo el medidor queda sin dueño. Con lecturas, solo lo que supere 2 kWh.
    _formula(
        hoja.cell(26, 2),
        'IF(SUM(B22:B24)=0,IF(Factura!B5>0,Factura!B5,0),IF(Factura!B5-SUM(B22:B24)>2,Factura!B5-SUM(B22:B24),0))',
    )
    _formula(
        hoja.cell(26, 3),
        'IF(SUM(C22:C24)=0,IF(Factura!B6>0,Factura!B6,0),IF(Factura!B6-SUM(C22:C24)>2,Factura!B6-SUM(C22:C24),0))',
    )
    hoja.cell(26, 2).number_format = ENTERO
    hoja.cell(26, 3).number_format = ENTERO
    hoja["B26"].comment = Comment(
        "Consumo de alta que el medidor marcó y ninguna de las tres lecturas explica.",
        "pozo",
    )

    _formula(
        hoja.cell(28, 1),
        'IF(OR(SUM(B22:B24)>Factura!B5+2,SUM(C22:C24)>Factura!B6+2),'
        '"Las lecturas superan al medidor.",'
        'IF(OR(B26>0,C26>0),"Falta asignar consumo. Sin asignar no es una cuarta persona.",'
        '"Las lecturas cubren el medidor."))',
    )
    hoja.merge_cells("A28:G28")
    hoja["A28"].font = NEGRO_NEGRITA

    for columna, ancho in enumerate((28, 16, 16, 16, 16, 14, 14), start=1):
        hoja.column_dimensions[get_column_letter(columna)].width = ancho
    hoja.freeze_panes = "A4"
    hoja.print_title_rows = "3:3"
    hoja.page_setup.orientation = "landscape"
    hoja.auto_filter.ref = "A3:G18"


def _hoja_reparto(hoja) -> None:
    hoja["A1"] = "División de la factura"
    hoja["A1"].font = TITULO
    hoja.merge_cells("A1:F1")
    _formula(hoja.cell(2, 1), 'Factura!B7&" al "&Factura!B8&"  ·  vence "&Factura!B9')
    hoja.merge_cells("A2:F2")

    hoja["A4"] = "Concepto"
    for columna, referencia in enumerate(("Lecturas!A22", "Lecturas!A23", "Lecturas!A24"), start=2):
        _formula(hoja.cell(4, columna), referencia)
    hoja["E4"] = "Sin asignar"
    hoja["F4"] = "Total"
    for columna in range(1, 7):
        hoja.cell(4, columna).font = ENCABEZADO_FUENTE
        hoja.cell(4, columna).fill = ENCABEZADO

    montos = {grupo: f"Factura!F{FILA_GRUPOS + indice}" for indice, grupo in enumerate(GRUPOS)}
    suma_alta = "(Lecturas!B22+Lecturas!B23+Lecturas!B24+Lecturas!B26)"
    suma_baja = "(Lecturas!C22+Lecturas!C23+Lecturas!C24+Lecturas!C26)"
    hay_sin_asignar = "(Lecturas!B26+Lecturas!C26)>0"

    filas_formula = {
        5: _proporcional(
            5,
            montos["energia_alta"],
            ["Lecturas!B22", "Lecturas!B23", "Lecturas!B24", "Lecturas!B26"],
            suma_alta,
            "Lecturas!B26>0",
        ),
        6: _proporcional(
            6,
            montos["energia_baja"],
            ["Lecturas!C22", "Lecturas!C23", "Lecturas!C24", "Lecturas!C26"],
            suma_baja,
            "Lecturas!C26>0",
        ),
        7: _proporcional(
            7,
            montos["subsidio_nacional"],
            ["B5+B6", "C5+C6", "D5+D6", "E5+E6"],
            "(F5+F6)",
            hay_sin_asignar,
        ),
        8: _tercios(8, montos["fijo"]),
        9: _tercios(9, montos["penal"]),
        10: _proporcional(
            10,
            montos["iibb"],
            ["B5+B6+B7+B8+B9", "C5+C6+C7+C8+C9", "D5+D6+D7+D8+D9", "E5+E6+E7+E8+E9"],
            "(F5+F6+F7+F8+F9)",
            hay_sin_asignar,
        ),
        11: _proporcional(
            11,
            montos["ley9083"],
            ["B5+B6+B7+B8+B10", "C5+C6+C7+C8+C10", "D5+D6+D7+D8+D10", "E5+E6+E7+E8+E10"],
            "(F5+F6+F7+F8+F10)",
            hay_sin_asignar,
        ),
    }

    etiquetas_fila = {
        5: "energia_alta",
        6: "energia_baja",
        7: "subsidio_nacional",
        8: "fijo",
        9: "penal",
        10: "iibb",
        11: "ley9083",
    }
    for fila, grupo in etiquetas_fila.items():
        _texto(hoja.cell(fila, 1), ETIQUETAS[grupo])
        hoja.cell(fila, 1).fill = GRIS
        for columna, formula in enumerate(filas_formula[fila], start=2):
            _formula(hoja.cell(fila, columna), formula, pesos=True)
        _formula(hoja.cell(fila, 6), f"SUM(B{fila}:E{fila})", pesos=True)

    for columna in range(2, 6):
        letra = get_column_letter(columna)
        _formula(hoja.cell(12, columna), f"{letra}5+{letra}6+{letra}7+{letra}8+{letra}9+{letra}10+{letra}11", pesos=True)
    _texto(hoja.cell(12, 1), "Subtotal eléctrico", negrita=True)
    _formula(hoja.cell(12, 6), "SUM(B12:E12)", pesos=True)

    impuestos = {
        13: _proporcional(13, montos["iva"], ["B12", "C12", "D12", "E12"], "F12", hay_sin_asignar),
        14: _proporcional(
            14,
            montos["tasa"],
            ["B12-B7", "C12-C7", "D12-D7", "E12-E7"],
            "(F12-F7)",
            hay_sin_asignar,
        ),
        15: _proporcional(
            15,
            montos["ccce"],
            ["B12-B7", "C12-C7", "D12-D7", "E12-E7"],
            "(F12-F7)",
            hay_sin_asignar,
        ),
        16: _tercios(16, montos["otros"]),
    }
    for fila, grupo in ((13, "iva"), (14, "tasa"), (15, "ccce"), (16, "otros")):
        _texto(hoja.cell(fila, 1), ETIQUETAS[grupo])
        for columna, formula in enumerate(impuestos[fila], start=2):
            _formula(hoja.cell(fila, columna), formula, pesos=True)
        _formula(hoja.cell(fila, 6), f"SUM(B{fila}:E{fila})", pesos=True)

    _texto(hoja.cell(17, 1), "Total a pagar", negrita=True)
    for columna in range(2, 6):
        letra = get_column_letter(columna)
        _formula(hoja.cell(17, columna), f"{letra}12+{letra}13+{letra}14+{letra}15+{letra}16", pesos=True)
        hoja.cell(17, columna).font = NEGRO_NEGRITA
    _formula(hoja.cell(17, 6), "SUM(B17:E17)", pesos=True)
    hoja.cell(17, 6).font = NEGRO_NEGRITA
    for columna in range(1, 7):
        hoja.cell(17, columna).fill = PatternFill("solid", fgColor="D6EAF8")

    _texto(hoja.cell(19, 1), "Diferencia contra la factura", negrita=True)
    _formula(hoja.cell(19, 2), "F17-Factura!B4", pesos=True)
    _texto(hoja.cell(20, 1), "Suma de las 3 personas", negrita=True)
    _formula(hoja.cell(20, 2), "B17+C17+D17", pesos=True)
    _texto(hoja.cell(21, 1), "Pendiente sin asignar", negrita=True)
    _formula(hoja.cell(21, 2), "E17", pesos=True)

    hoja["A7"].comment = Comment(
        "Se reparte en proporción a los pesos de energía alta y baja de cada persona.",
        "pozo",
    )
    hoja["A8"].comment = Comment(
        "Comercialización, uso de red y exceso de potencia, en tres partes iguales.",
        "pozo",
    )
    hoja["A11"].comment = Comment(
        "Beneficio provincial. Sigue la base eléctrica salvo mora, aviso de suspensión y reanudación.",
        "pozo",
    )
    hoja["A16"].comment = Comment(
        "CAP municipal, impuesto de sellos, redondeos, bonificaciones y devoluciones, en tres partes iguales.",
        "pozo",
    )

    hoja.column_dimensions["A"].width = 52
    for columna in range(2, 7):
        hoja.column_dimensions[get_column_letter(columna)].width = 18
    hoja.freeze_panes = "B5"
    hoja.print_title_rows = "1:4"
    hoja.page_setup.orientation = "landscape"
    hoja.sheet_view.zoomScale = 120
