"""Panel Streamlit para repartir una factura de EDEMSA entre tres personas."""

from __future__ import annotations

import csv
import sys
import tempfile
from decimal import Decimal, InvalidOperation
from pathlib import Path

# En Streamlit Community Cloud el paquete se importa desde src/ sin instalar el wheel.
_SRC = Path(__file__).resolve().parent / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

import streamlit as st

from pozo.excel_out import Sesion, escribir_reparto
from pozo.factura import ParseError, parse_factura, parse_factura_texto
from pozo.reparto import repartir
from pozo.ui_state import (
    filas_reparto,
    lecturas_desde_sesiones,
    resumen_totales,
    workbook_a_bytes,
)

ROOT = Path(__file__).resolve().parent
DEFAULT_PERSONAS = ("Persona 1", "Persona 2", "Persona 3")
FIXTURA_EJEMPLO = ROOT / "tests" / "fixtures" / "factura_11-2025.txt"


def _dinero(valor: Decimal) -> str:
    return f"${valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _decimal(valor) -> Decimal:
    if valor is None or valor == "":
        return Decimal("0")
    try:
        return Decimal(str(valor))
    except (InvalidOperation, ValueError):
        return Decimal("0")


def _pdfs_locales() -> list[Path]:
    def clave(path: Path) -> tuple[int, int]:
        # Acepta nombres tipo 0000001_11-2025.pdf o factura_11-2025.pdf
        trozo = path.stem
        if "_" in trozo:
            trozo = trozo.split("_", 1)[1]
        if "-" not in trozo:
            return (0, 0)
        mes, anio = trozo.split("-", 1)
        try:
            return int(anio), int(mes)
        except ValueError:
            return (0, 0)

    return sorted(ROOT.glob("*.pdf"), key=clave)


def _cargar_factura(subido, elegido: str | None, usar_ejemplo: bool):
    if usar_ejemplo:
        return parse_factura_texto(FIXTURA_EJEMPLO.read_text(encoding="utf-8"))
    if subido is not None:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(subido.getvalue())
            ruta = Path(tmp.name)
        try:
            return parse_factura(ruta)
        finally:
            ruta.unlink(missing_ok=True)
    if elegido:
        return parse_factura(ROOT / elegido)
    return None


def _filas_editor(editor) -> list[dict]:
    if hasattr(editor, "to_dict"):
        return editor.to_dict("records")
    return list(editor)


def _sesiones_desde_editor(filas, personas: tuple[str, str, str]) -> list[Sesion]:
    sesiones = []
    for fila in filas:
        persona = str(fila.get("persona") or "").strip()
        if not persona or persona not in personas:
            continue
        alta_inicio = _decimal(fila.get("alta_inicio"))
        alta_fin = _decimal(fila.get("alta_fin"))
        baja_inicio = _decimal(fila.get("baja_inicio"))
        baja_fin = _decimal(fila.get("baja_fin"))
        # Se admite solo alta o solo baja: la otra banda queda en cero.
        if alta_fin == alta_inicio and baja_fin == baja_inicio:
            continue
        sesiones.append(
            Sesion(
                persona=persona,
                alta_inicio=alta_inicio,
                alta_fin=alta_fin,
                baja_inicio=baja_inicio,
                baja_fin=baja_fin,
            )
        )
    return sesiones


def _filas_iniciales(personas: tuple[str, str, str], factura_mes: str) -> list[dict]:
    # Varias filas vacías para que se vea que se pueden cargar múltiples lecturas.
    filas = []
    for persona in personas:
        for _ in range(2):
            filas.append(
                {
                    "persona": persona,
                    "alta_inicio": 0.0,
                    "alta_fin": 0.0,
                    "baja_inicio": 0.0,
                    "baja_fin": 0.0,
                }
            )
    anio, mes = factura_mes.split("/")[1], factura_mes.split("/")[0]
    csv_local = ROOT / "lecturas" / f"{anio}-{int(mes):02d}.csv"
    if csv_local.is_file():
        with csv_local.open(encoding="utf-8") as archivo:
            cargadas = list(csv.DictReader(archivo))
        if cargadas:
            return [
                {
                    "persona": fila["persona"],
                    "alta_inicio": float(fila["alta_inicio"]),
                    "alta_fin": float(fila["alta_fin"]),
                    "baja_inicio": float(fila["baja_inicio"]),
                    "baja_fin": float(fila["baja_fin"]),
                }
                for fila in cargadas
            ]
    return filas


def main() -> None:
    st.set_page_config(page_title="Pozo EDEMSA", page_icon="⚡", layout="wide")
    st.title("Pozo compartido · EDEMSA")
    st.caption(
        "La suma de las tres personas más «Sin asignar» tiene que coincidir con el total de la factura. "
        "Los PDF se procesan solo para calcular el reparto; no se guardan en este sitio."
    )

    with st.sidebar:
        st.header("Factura")
        usar_ejemplo = st.checkbox(
            "Usar factura de ejemplo (anónima)",
            value=False,
            help="Carga una factura de prueba sin datos personales para probar el reparto.",
        )
        subido = st.file_uploader(
            "Subí el PDF de EDEMSA",
            type=["pdf"],
            help="La factura tiene que ser el PDF con texto (no una foto escaneada).",
            disabled=usar_ejemplo,
        )
        locales = [p.name for p in _pdfs_locales()]
        elegido = st.selectbox(
            "O elegí un PDF de esta carpeta",
            options=[""] + locales,
            format_func=lambda x: "—" if x == "" else x,
            help="Los PDF locales no se suben a ningún servidor: se leen solo en tu máquina.",
            disabled=usar_ejemplo,
        )
        st.divider()
        st.header("Personas")
        st.caption("Los nombres tienen que coincidir con los de la tabla de lecturas.")
        p1 = st.text_input("Persona 1", value=DEFAULT_PERSONAS[0])
        p2 = st.text_input("Persona 2", value=DEFAULT_PERSONAS[1])
        p3 = st.text_input("Persona 3", value=DEFAULT_PERSONAS[2])
        personas = (
            p1.strip() or DEFAULT_PERSONAS[0],
            p2.strip() or DEFAULT_PERSONAS[1],
            p3.strip() or DEFAULT_PERSONAS[2],
        )

    try:
        factura = _cargar_factura(subido, elegido or None, usar_ejemplo)
    except ParseError as error:
        st.error(str(error))
        return

    if factura is None:
        st.info("Subí un PDF, elegí uno de la carpeta o activá la factura de ejemplo para empezar.")
        with st.expander("Cómo se reparte", expanded=True):
            st.markdown(
                """
1. **Energía alta y baja**: según los kilovatios hora de cada persona. Podés cargar **varias lecturas** por persona (una fila por cada uso del pozo).
2. Si en una fila solo usaste **alta** o solo **baja**, dejá la otra banda en cero.
3. **Cargos fijos** (comercialización, uso de red, exceso de potencia) y cargos excepcionales (mora, suspensión, etc.): en tres partes iguales.
4. **Subsidios** e impuestos: siguen la misma lógica que la factura (IVA del 27 % de responsable inscripto).
5. Si falta consumo del medidor, queda en **Sin asignar** (no es una cuarta persona).
                """
            )
        return

    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("Mes", factura.mes)
    col_b.metric("Total factura", _dinero(factura.total))
    col_c.metric("kWh alta", f"{factura.kwh_alta:,.0f}".replace(",", "."))
    col_d.metric("kWh baja", f"{factura.kwh_baja:,.0f}".replace(",", "."))
    st.caption(
        f"Factura {factura.numero} · NIC {factura.nic} · "
        f"{factura.periodo_desde} al {factura.periodo_hasta} · vence {factura.vencimiento}"
    )

    st.subheader("Lecturas del medidor")
    st.markdown(
        """
Agregá **una fila por cada uso** del pozo (botón **＋** al final de la tabla).
La misma persona puede aparecer muchas veces.

- Si ya tenés los kilovatios hora: poné **inicio 0** y **fin** igual a ese consumo.
- Si solo usaste **alta** o solo **baja** en ese tramo: dejá la otra banda en **0**.
- Los nombres tienen que coincidir con los de la barra lateral.
        """
    )

    anio, mes = factura.mes.split("/")[1], factura.mes.split("/")[0]
    editor = st.data_editor(
        _filas_iniciales(personas, factura.mes),
        num_rows="dynamic",
        width="stretch",
        column_config={
            "persona": st.column_config.SelectboxColumn(
                "Persona",
                options=list(personas),
                required=True,
                help="Quién usó el pozo en ese tramo.",
            ),
            "alta_inicio": st.column_config.NumberColumn(
                "Alta inicio",
                format="%.2f",
                help="Lectura de energía alta al empezar. Si ya tenés los kWh, usá 0.",
            ),
            "alta_fin": st.column_config.NumberColumn(
                "Alta fin",
                format="%.2f",
                help="Lectura de energía alta al terminar. Si no hubo consumo en alta, igual al inicio.",
            ),
            "baja_inicio": st.column_config.NumberColumn(
                "Baja inicio",
                format="%.2f",
                help="Lectura de energía baja al empezar. Si ya tenés los kWh, usá 0.",
            ),
            "baja_fin": st.column_config.NumberColumn(
                "Baja fin",
                format="%.2f",
                help="Lectura de energía baja al terminar. Si no hubo consumo en baja, igual al inicio.",
            ),
        },
        key=f"lecturas_{factura.numero}_{factura.mes}",
    )

    sesiones = _sesiones_desde_editor(_filas_editor(editor), personas)
    lecturas = lecturas_desde_sesiones(personas, sesiones)
    resultado = repartir(factura, lecturas)

    # Resumen de cuántas lecturas cargó cada uno
    conteo = {nombre: 0 for nombre in personas}
    for sesion in sesiones:
        conteo[sesion.persona] += 1
    st.caption(
        "Lecturas cargadas: "
        + " · ".join(f"{nombre}: {conteo[nombre]}" for nombre in personas)
    )

    if resultado.advertencia:
        st.warning(resultado.advertencia)

    st.subheader("Quién paga cuánto")
    resumen = resumen_totales(resultado)
    for fila, lectura in zip(resumen[:-1], lecturas, strict=True):
        fila["kwh_alta"] = lectura.kwh_alta
        fila["kwh_baja"] = lectura.kwh_baja

    cols = st.columns(len(resumen))
    for columna, fila in zip(cols, resumen, strict=True):
        with columna:
            st.metric(fila["persona"], _dinero(fila["total"]))
            if fila["kwh_alta"] is not None:
                st.caption(f"{fila['kwh_alta']} kWh alta · {fila['kwh_baja']} kWh baja")

    diferencia = sum((fila["total"] for fila in resumen), Decimal("0")) - factura.total
    if diferencia == 0:
        st.success("Cierra contra la factura: diferencia $0,00")
    else:
        st.error(f"No cierra: diferencia {_dinero(diferencia)}")

    with st.expander("Detalle por concepto", expanded=False):
        detalle = filas_reparto(resultado)
        st.dataframe(
            [
                {
                    clave: (_dinero(valor) if isinstance(valor, Decimal) else valor)
                    for clave, valor in fila.items()
                }
                for fila in detalle
            ],
            width="stretch",
            hide_index=True,
        )

    st.subheader("Exportar Excel")
    st.caption("El archivo usa las mismas fórmulas: si corregís una lectura en Excel, se recalcula el reparto.")
    with tempfile.TemporaryDirectory() as tmp:
        destino = Path(tmp) / f"{anio}-{int(mes):02d}.xlsx"
        escribir_reparto(factura, destino, personas=personas, sesiones=sesiones)
        contenido = workbook_a_bytes(destino)
    st.download_button(
        label="Descargar Excel del reparto",
        data=contenido,
        file_name=f"pozo-{anio}-{int(mes):02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Descarga el reparto en Excel para compartir o archivar.",
    )


if __name__ == "__main__":
    main()
