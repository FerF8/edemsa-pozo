# Pozo EDEMSA

Reparte una factura de [EDEMSA](https://www.edemsa.com/) (Mendoza) entre tres personas que comparten un medidor.

Lee el PDF de la factura, clasifica cada concepto (energía en alta y baja, cargos fijos, subsidios, impuestos) y calcula cuánto paga cada uno. La suma de las tres partes más lo que quede sin lectura tiene que coincidir con el total de la factura.

## Requisitos

- Python 3.11 o superior
- [uv](https://docs.astral.sh/uv/) (recomendado)

## Uso rápido (interfaz)

```bash
uv sync
uv run streamlit run app.py
```

Abrí la URL local que muestra la terminal (por defecto `http://localhost:8501`).

1. Subí el PDF de EDEMSA, o activá **Usar factura de ejemplo (anónima)** para probar.
2. Nombrá a las tres personas en la barra lateral.
3. Cargá las lecturas: **una fila por cada uso** del pozo (botón **＋**). La misma persona puede figurar muchas veces.
4. Si en un tramo solo usaste **alta** o solo **baja**, dejá la otra banda en cero.
5. Revisá los totales y, si querés, descargá el Excel.

Los PDF y las lecturas se procesan **solo en tu máquina**. Este proyecto no envía facturas a ningún servidor.

## Uso por línea de comandos

```bash
uv run pozo dividir factura.pdf -o reparto.xlsx
uv run pozo dividir factura.pdf --lecturas lecturas/ejemplo.csv -o reparto.xlsx
```

Formato del CSV de lecturas:

```csv
persona,alta_inicio,alta_fin,baja_inicio,baja_fin
Persona 1,100,110,200,220
Persona 1,110,115,0,0
```

## Cómo se reparte

| Concepto | Criterio |
|---|---|
| Energía alta / baja | Según los kilovatios hora de cada persona (sumando todas sus filas) |
| Subsidio nacional | En proporción a la energía de cada uno |
| Ley 9083 (si aparece) | Sobre la base eléctrica, sin mora ni avisos de corte |
| Comercialización, uso de red, exceso de potencia | En tres partes iguales |
| Mora, suspensión, reanudación, bonificaciones, devoluciones | En tres partes iguales |
| IVA (27 % responsable inscripto), tasa, CCCE, CAP, sellos, redondeos | Según la factura; los fijos de impuestos van en tres |

Si las lecturas no cubren el medidor (más de 2 kWh de diferencia), el faltante queda en **Sin asignar**. No es una cuarta persona: son los kilos que todavía nadie anotó. Los cargos fijos igual se dividen en tres.

## Datos personales

Este repositorio **no incluye** facturas PDF, fotos del medidor ni planillas con nombres reales.

- Las pruebas usan texto de facturas **anonimizado** en `tests/fixtures/`.
- Tus PDF, fotos y CSV reales deben quedar fuera de Git (están en `.gitignore`).
- No subas a GitHub archivos con NIC, CUIT, domicilio o lecturas reales.

## Pruebas

```bash
uv run pytest
```

## Publicar en internet (Streamlit Community Cloud)

Así otras personas entran con un link, sin instalar nada.

1. Entrá a [share.streamlit.io](https://share.streamlit.io) e iniciá sesión con GitHub.
2. **Create app** → elegí el repo `FerF8/edemsa-pozo`, rama `main`, archivo `app.py`.
3. Deploy. Cuando termine, vas a tener una URL tipo `https://….streamlit.app`.
4. Compartí ese link. Quien lo abra puede subir su PDF y cargar las lecturas.

Si cambiás el código, hacé `git push` y en Streamlit Cloud usá **Reboot** o esperá el redeploy automático.

## Licencia

Uso libre para quien comparta un medidor. No está afiliado a EDEMSA.
