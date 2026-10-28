"""El reparto cierra al centavo contra la factura.

Los centavos de diferencia quedan en la tercera persona cuando las lecturas
cubren el medidor, y en la columna Sin asignar cuando falta consumo.
"""

from decimal import Decimal
from pathlib import Path

from pozo.factura import Concepto, Factura, parse_factura_texto
from pozo.reparto import Lectura, repartir

D = Decimal
FIXTURES = Path(__file__).resolve().parent / "fixtures"


def _factura() -> Factura:
    conceptos = [
        Concepto("Energía alta", "energia_alta", D("300.00")),
        Concepto("Energía baja", "energia_baja", D("100.00")),
        Concepto("Subsidio nacional", "subsidio_nacional", D("-40.00")),
        Concepto("Cargos fijos", "fijo", D("90.00")),
        Concepto("Ingresos brutos", "iibb", D("15.00")),
        Concepto("Ley 9083", "ley9083", D("-20.00")),
        Concepto("IVA", "iva", D("50.00")),
        Concepto("Tasa", "tasa", D("10.00")),
        Concepto("CCCE", "ccce", D("20.00")),
        Concepto("Otros", "otros", D("9.00")),
    ]
    return Factura(
        nic="0000001",
        numero="0003-1",
        mes="11/2025",
        periodo_desde="01/10/2025",
        periodo_hasta="01/11/2025",
        vencimiento="25/11/2025",
        total=sum((c.importe for c in conceptos), D("0")),
        kwh_alta=D("3"),
        kwh_baja=D("2"),
        conceptos=conceptos,
        alta_inicio=D("0"),
        alta_fin=D("3"),
        baja_inicio=D("0"),
        baja_fin=D("2"),
    )


def test_con_lecturas_completas_los_tres_suman_el_total():
    factura = _factura()
    assert factura.total == D("534.00")
    resultado = repartir(
        factura,
        [
            Lectura("A", D("2"), D("0")),
            Lectura("B", D("1"), D("2")),
            Lectura("C", D("0"), D("0")),
        ],
    )

    assert resultado.personas[0].total == D("248.08")
    assert resultado.personas[1].total == D("248.08")
    assert resultado.personas[2].total == D("37.84")
    assert resultado.sin_asignar.total == D("0.00")
    assert sum(p.total for p in resultado.personas) == factura.total


def test_el_consumo_que_falta_no_se_le_carga_a_los_que_anotaron():
    factura = _factura()
    factura.kwh_alta = D("10")
    factura.kwh_baja = D("10")
    resultado = repartir(
        factura,
        [
            Lectura("A", D("2"), D("0")),
            Lectura("B", D("0"), D("0")),
            Lectura("C", D("0"), D("0")),
        ],
    )

    assert resultado.kwh_sin_asignar_alta == D("8")
    assert resultado.kwh_sin_asignar_baja == D("10")
    assert resultado.personas[0].importes["energia_alta"] == D("60.00")
    assert resultado.personas[1].importes["energia_alta"] == D("0.00")
    assert resultado.sin_asignar.importes["energia_alta"] == D("240.00")
    assert resultado.sin_asignar.importes["energia_baja"] == D("100.00")
    assert resultado.personas[1].importes["fijo"] == D("30.00")
    assert resultado.sin_asignar.importes["fijo"] == D("0.00")
    totales = [p.total for p in resultado.personas] + [resultado.sin_asignar.total]
    assert sum(totales) == factura.total


def test_sin_lecturas_la_energia_queda_sin_asignar_y_el_total_cierra():
    factura = _factura()
    resultado = repartir(
        factura,
        [
            Lectura("A", D("0"), D("0")),
            Lectura("B", D("0"), D("0")),
            Lectura("C", D("0"), D("0")),
        ],
    )
    assert resultado.personas[0].importes["energia_alta"] == D("0.00")
    assert resultado.sin_asignar.importes["energia_alta"] == D("300.00")
    assert resultado.sin_asignar.importes["energia_baja"] == D("100.00")
    assert sum(p.total for p in resultado.personas) + resultado.sin_asignar.total == factura.total


def test_noviembre_con_lectura_parcial_cierra():
    texto = (FIXTURES / "factura_11-2025.txt").read_text(encoding="utf-8")
    factura = parse_factura_texto(texto)
    resultado = repartir(
        factura,
        [
            Lectura("Persona 1", D("134"), D("318")),
            Lectura("Persona 2", D("0"), D("0")),
            Lectura("Persona 3", D("0"), D("0")),
        ],
    )
    assert resultado.kwh_sin_asignar_alta == D("1967")
    assert resultado.kwh_sin_asignar_baja == D("2393")
    assert sum(p.total for p in resultado.personas) + resultado.sin_asignar.total == factura.total


def test_una_diferencia_de_hasta_2_kwh_no_arma_un_cuarto_socio():
    factura = _factura()
    factura.kwh_alta = D("5")
    resultado = repartir(
        factura,
        [
            Lectura("A", D("2"), D("0")),
            Lectura("B", D("1"), D("2")),
            Lectura("C", D("0"), D("0")),
        ],
    )
    assert resultado.kwh_sin_asignar_alta == D("0")
    assert sum(p.total for p in resultado.personas) == factura.total
