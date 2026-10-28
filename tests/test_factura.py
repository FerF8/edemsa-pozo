"""La factura de EDEMSA tiene que leerse completa: cada peso queda clasificado.

Las fixtures son texto anónimo (sin titular, CUIT, dirección ni NIC real).
"""

from decimal import Decimal
from pathlib import Path

import pytest

from pozo.factura import ParseError, parse_factura, parse_factura_texto

FIXTURES = Path(__file__).resolve().parent / "fixtures"

# Totales, kilovatios hora y subtotales de las facturas de ejemplo anonimizadas.
ESPERADO = {
    1: ("505118.00", "2360", "2654", "366673.73", "-365290.73"),
    2: ("174329.00", "1474", "1747", "272318.90", "-262758.68"),
    3: ("529484.00", "2342", "2874", "386130.21", "-376403.07"),
    4: ("24434.00", "125", "275", "108308.82", "-108308.82"),
    5: ("284058.00", "22", "200", "205080.83", None),
    6: ("265591.00", "37", "342", "189901.77", None),
    7: ("261522.00", "116", "28", "187301.93", None),
    8: ("255285.00", "26", "131", "182433.75", None),
    9: ("388468.00", "455", "483", "280696.16", None),
    10: ("291974.00", "130", "17", "213476.51", None),
    11: ("983968.00", "2101", "2711", "719117.10", None),
}

GRUPOS_ELECTRICOS = {
    "energia_alta",
    "energia_baja",
    "subsidio_nacional",
    "fijo",
    "penal",
    "iibb",
    "ley9083",
}


def _fixture(mes: int) -> str:
    return (FIXTURES / f"factura_{mes:02d}-2025.txt").read_text(encoding="utf-8")


@pytest.mark.parametrize("mes", sorted(ESPERADO))
def test_cada_factura_cierra_y_clasifica_todos_los_pesos(mes: int):
    total, alta, baja, electricos, ley = ESPERADO[mes]
    factura = parse_factura_texto(_fixture(mes))

    assert factura.total == Decimal(total)
    assert factura.kwh_alta == Decimal(alta)
    assert factura.kwh_baja == Decimal(baja)
    assert factura.nic == "0000001"

    assert sum(c.importe for c in factura.conceptos) == factura.total
    electrico = sum(
        c.importe for c in factura.conceptos if c.grupo in GRUPOS_ELECTRICOS
    )
    assert electrico == Decimal(electricos)

    leyes = [c.importe for c in factura.conceptos if c.grupo == "ley9083"]
    if ley is None:
        assert leyes == []
    else:
        assert sum(leyes) == Decimal(ley)


def test_noviembre_separa_energia_subsidio_e_impuestos():
    factura = parse_factura_texto(_fixture(11))
    por_grupo = {}
    for concepto in factura.conceptos:
        por_grupo[concepto.grupo] = por_grupo.get(concepto.grupo, Decimal("0")) + concepto.importe

    assert por_grupo["energia_alta"] == Decimal("354915.76")
    assert por_grupo["energia_baja"] == Decimal("281221.38")
    assert por_grupo["subsidio_nacional"] == Decimal("-89607.06")
    assert por_grupo["fijo"] == Decimal("28957.38") + Decimal("61028.06") + Decimal("61028.06")
    assert por_grupo["iibb"] == Decimal("21573.52")
    assert por_grupo["iva"] == Decimal("194161.62")
    assert por_grupo["tasa"] == Decimal("12130.86")
    assert por_grupo["ccce"] == Decimal("48523.45")
    assert "penal" not in por_grupo
    assert "ley9083" not in por_grupo


def test_febrero_deja_la_bonificacion_fuera_del_iva():
    factura = parse_factura_texto(_fixture(2))
    bonif = [c for c in factura.conceptos if "fuerza mayor" in c.nombre.lower() or c.grupo == "otros"]
    assert any(c.importe == Decimal("-204898.40") and c.grupo == "otros" for c in bonif)
    assert all(c.grupo != "fijo" for c in factura.conceptos if c.importe == Decimal("-204898.40"))


def test_pdf_desconocido_no_se_inventa():
    with pytest.raises(ParseError):
        parse_factura(Path(__file__).resolve().parents[1] / "no-existe.pdf")


def test_fixtures_sin_datos_personales():
    prohibido = ("AMBROGETTI", "FUCILLI", "PINCOLINI", "CARRIZAL", "VISTALBA", "1505416", "20174640336", "999203")
    for ruta in FIXTURES.glob("factura_*.txt"):
        texto = ruta.read_text(encoding="utf-8").upper()
        for palabra in prohibido:
            assert palabra not in texto, f"{ruta.name} todavía tiene {palabra}"
