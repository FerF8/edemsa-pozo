"""Reparte una factura ya leída entre tres personas.

Los centavos quedan en la tercera persona cuando las lecturas cubren el medidor.
Si faltan más de 2 kilovatios hora, esa diferencia va a «Sin asignar»: no se le
carga el consumo a quien sí anotó, y los cargos fijos siguen en tres partes.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP

from pozo.factura import GRUPOS, Factura

TOLERANCIA_KWH = Decimal("2")
CERO = Decimal("0.00")


@dataclass
class Lectura:
    nombre: str
    kwh_alta: Decimal
    kwh_baja: Decimal


@dataclass
class Parte:
    nombre: str
    importes: dict[str, Decimal] = field(default_factory=dict)

    @property
    def total(self) -> Decimal:
        return sum(self.importes.values(), CERO).quantize(Decimal("0.01"))


@dataclass
class Reparto:
    personas: list[Parte]
    sin_asignar: Parte
    kwh_sin_asignar_alta: Decimal
    kwh_sin_asignar_baja: Decimal
    advertencia: str


def _q(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def _suma(factura: Factura, grupo: str) -> Decimal:
    return sum((c.importe for c in factura.conceptos if c.grupo == grupo), CERO)


def _gap(medidor: Decimal, anotado: Decimal) -> Decimal:
    # Sin ninguna lectura, todo el medidor está sin dueño, aunque sea poco.
    if anotado == 0 and medidor > 0:
        return medidor
    diferencia = medidor - anotado
    if diferencia > TOLERANCIA_KWH:
        return diferencia
    return CERO


def _proporcional(monto: Decimal, pesos: list[Decimal], indice_cierre: int) -> list[Decimal]:
    """Reparte `monto` según `pesos`. El índice de cierre absorbe los centavos."""
    if monto == 0:
        return [CERO for _ in pesos]
    total_peso = sum(pesos, CERO)
    if total_peso == 0:
        resultado = [CERO for _ in pesos]
        resultado[-1] = _q(monto)
        return resultado

    resultado = [CERO for _ in pesos]
    asignado = CERO
    for indice, peso in enumerate(pesos):
        if indice == indice_cierre:
            continue
        parte = _q(monto * peso / total_peso)
        resultado[indice] = parte
        asignado += parte
    resultado[indice_cierre] = _q(monto - asignado)
    return resultado


def _tercios(monto: Decimal) -> list[Decimal]:
    primero = _q(monto / 3)
    segundo = _q(monto / 3)
    tercero = _q(monto - primero - segundo)
    return [primero, segundo, tercero, CERO]


def repartir(factura: Factura, lecturas: list[Lectura]) -> Reparto:
    if len(lecturas) != 3:
        raise ValueError("El pozo se reparte entre 3 personas")

    alta = [lectura.kwh_alta for lectura in lecturas]
    baja = [lectura.kwh_baja for lectura in lecturas]
    sin_alta = _gap(factura.kwh_alta, sum(alta, CERO))
    sin_baja = _gap(factura.kwh_baja, sum(baja, CERO))
    pesos_alta = alta + [sin_alta]
    pesos_baja = baja + [sin_baja]
    # Si hay consumo sin dueño, los centavos de las filas variables cierran ahí.
    cierre_variable = 3 if (sin_alta + sin_baja) > 0 else 2

    columnas = _proporcional(_suma(factura, "energia_alta"), pesos_alta, 3 if sin_alta > 0 else 2)
    energia_baja = _proporcional(_suma(factura, "energia_baja"), pesos_baja, 3 if sin_baja > 0 else 2)
    filas = {
        "energia_alta": columnas,
        "energia_baja": energia_baja,
    }

    pesos_subsidio = [filas["energia_alta"][i] + filas["energia_baja"][i] for i in range(4)]
    filas["subsidio_nacional"] = _proporcional(
        _suma(factura, "subsidio_nacional"), pesos_subsidio, cierre_variable
    )
    filas["fijo"] = _tercios(_suma(factura, "fijo"))
    filas["penal"] = _tercios(_suma(factura, "penal"))

    def base(indices_grupos: tuple[str, ...]) -> list[Decimal]:
        return [sum((filas[grupo][i] for grupo in indices_grupos), CERO) for i in range(4)]

    filas["iibb"] = _proporcional(
        _suma(factura, "iibb"),
        base(("energia_alta", "energia_baja", "subsidio_nacional", "fijo", "penal")),
        cierre_variable,
    )
    filas["ley9083"] = _proporcional(
        _suma(factura, "ley9083"),
        base(("energia_alta", "energia_baja", "subsidio_nacional", "fijo", "iibb")),
        cierre_variable,
    )

    electrico = base(("energia_alta", "energia_baja", "subsidio_nacional", "fijo", "penal", "iibb", "ley9083"))
    filas["iva"] = _proporcional(_suma(factura, "iva"), electrico, cierre_variable)

    base_impuestos = [electrico[i] - filas["subsidio_nacional"][i] for i in range(4)]
    filas["tasa"] = _proporcional(_suma(factura, "tasa"), base_impuestos, cierre_variable)
    filas["ccce"] = _proporcional(_suma(factura, "ccce"), base_impuestos, cierre_variable)
    filas["otros"] = _tercios(_suma(factura, "otros"))

    def parte(nombre: str, indice: int) -> Parte:
        return Parte(nombre, {grupo: filas[grupo][indice] for grupo in GRUPOS})

    personas = [parte(lectura.nombre, indice) for indice, lectura in enumerate(lecturas)]
    sin_asignar = parte("Sin asignar", 3)

    avisos = []
    if sin_alta > 0 or sin_baja > 0:
        avisos.append(
            "Faltan lecturas: "
            f"{sin_alta} kWh en alta y {sin_baja} kWh en baja quedan en Sin asignar. "
            "Esa columna no es una cuarta persona."
        )
    if sum(alta, CERO) > factura.kwh_alta + TOLERANCIA_KWH or sum(baja, CERO) > factura.kwh_baja + TOLERANCIA_KWH:
        avisos.append("Las lecturas superan lo que marcó el medidor en la factura.")

    return Reparto(
        personas=personas,
        sin_asignar=sin_asignar,
        kwh_sin_asignar_alta=sin_alta,
        kwh_sin_asignar_baja=sin_baja,
        advertencia=" ".join(avisos),
    )
