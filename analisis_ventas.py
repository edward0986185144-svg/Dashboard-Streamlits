# -*- coding: utf-8 -*-
"""
Dashboard Comercial · Streamlit
===============================

Panel interactivo de ventas pensado para abrirse desde un enlace y usarse
sin escribir código: filtros, indicadores, gráficos y descarga de datos.

Ejecutar en local
-----------------
    pip install -r requirements.txt
    streamlit run analisis_ventas.py

Los datos se leen de ``ventas.csv`` (si tiene contenido) o, en su defecto,
de un conjunto de ejemplo. También se puede cargar un CSV o Excel desde
el botón "Cargar datos" del encabezado.

Columnas esperadas (los nombres se detectan de forma flexible):
    Fecha · Producto · Vendedor · Ventas · Cantidad
"""

from __future__ import annotations

import html
import io
import math
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.colors import sample_colorscale

# ---------------------------------------------------------------------------
# Configuración general
# ---------------------------------------------------------------------------

APP_DIR = Path(__file__).resolve().parent
CSV_PATH = APP_DIR / "ventas.csv"

INK = "#17181A"
MUTED = "#6E6E66"
LINE = "#E7E7E1"
ACCENT = "#0E7C66"
ACCENT_SOFT = "#EDF5F2"
BG = "#F5F5F2"

PALETTE = [
    "#0E7C66",
    "#7FB2A3",
    "#C9B27C",
    "#8C8C82",
    "#2F5D50",
    "#A9CDC2",
    "#D9C8A0",
    "#5F5F58",
]

MESES = ["ene", "feb", "mar", "abr", "may", "jun",
         "jul", "ago", "sep", "oct", "nov", "dic"]
MESES_LARGOS = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
                "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

PLOT_CONFIG = {
    "displayModeBar": False,
    "responsive": True,
    "scrollZoom": False,
}

DEMO = {
    "Fecha": [
        "2026-01-10", "2026-01-11", "2026-01-12", "2026-01-13",
        "2026-01-14", "2026-01-15", "2026-01-16", "2026-01-17",
    ],
    "Producto": [
        "Laptop", "Mouse", "Teclado", "Monitor",
        "Laptop", "Mouse", "Monitor", "Teclado",
    ],
    "Vendedor": ["Ana", "Luis", "Ana", "Carlos", "Luis", "Ana", "Carlos", "Luis"],
    "Ventas": [3500, 80, 150, 900, 4200, 100, 950, 100],
    "Cantidad": [1, 2, 1, 1, 1, 3, 1, 2],
}

try:
    import openpyxl  # noqa: F401  (necesario para exportar/leer Excel)

    HAS_OPENPYXL = True
except ImportError:  # pragma: no cover
    HAS_OPENPYXL = False


st.set_page_config(
    page_title="Dashboard Comercial",
    page_icon="◼",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# Estilos
# ---------------------------------------------------------------------------

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

:root {
    --bg: #F5F5F2;
    --card: #FFFFFF;
    --ink: #17181A;
    --muted: #6E6E66;
    --line: #E7E7E1;
    --accent: #0E7C66;
    --accent-soft: #EDF5F2;
}

html, body, .stApp, [class*="st-"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    -webkit-font-smoothing: antialiased;
}

.stApp { background: var(--bg); }

header[data-testid="stHeader"] {
    background: transparent;
    height: 0;
}
#MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] {
    display: none !important;
}

.block-container {
    padding: 2.5rem 2rem 3.5rem 2rem;
    max-width: 1280px;
}

/* ---------- Encabezado ---------- */
.hero-eyebrow {
    font-size: 11px;
    letter-spacing: .16em;
    text-transform: uppercase;
    font-weight: 600;
    color: var(--accent);
    margin-bottom: 8px;
}
.hero-title {
    font-size: clamp(30px, 4vw, 42px);
    line-height: 1.05;
    letter-spacing: -.035em;
    font-weight: 650;
    color: var(--ink);
    margin: 0 0 10px 0;
}
.hero-sub {
    font-size: 14px;
    color: var(--muted);
    margin: 0 0 16px 0;
    max-width: 620px;
}
.pill {
    display: inline-flex;
    align-items: center;
    gap: 7px;
    font-size: 11.5px;
    color: var(--muted);
    background: var(--card);
    border: 1px solid var(--line);
    padding: 5px 11px;
    border-radius: 999px;
    margin: 0 7px 6px 0;
    white-space: nowrap;
}
.pill b { color: var(--ink); font-weight: 600; }
.pill-accent {
    color: #0B5F4F;
    background: var(--accent-soft);
    border-color: #D5E7E1;
}
.dot {
    width: 6px;
    height: 6px;
    border-radius: 50%;
    background: var(--accent);
    display: inline-block;
    flex: 0 0 auto;
}

/* ---------- Tarjetas KPI ---------- */
.kpi {
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 18px 20px 16px 20px;
    height: 100%;
    box-shadow: 0 1px 2px rgba(23, 24, 26, .03);
}
.kpi-label {
    font-size: 10.5px;
    letter-spacing: .1em;
    text-transform: uppercase;
    font-weight: 600;
    color: #75756C;
}
.kpi-value {
    font-size: clamp(23px, 2.4vw, 29px);
    font-weight: 650;
    letter-spacing: -.03em;
    line-height: 1.15;
    color: var(--ink);
    margin: 7px 0 3px 0;
    font-variant-numeric: tabular-nums;
}
.kpi-value.accent { color: var(--accent); }
.kpi-foot {
    font-size: 11.5px;
    color: var(--muted);
    line-height: 1.35;
}

/* ---------- Contenedores con borde = tarjetas ---------- */
.card-marker { display: none; }

[data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .card-marker) {
    background: var(--card);
    border: 1px solid var(--line) !important;
    border-radius: 18px;
    padding: 18px 20px 12px 20px !important;
    box-shadow: 0 1px 2px rgba(23, 24, 26, .03);
}

/* ---------- Títulos de sección ---------- */
.sec { margin-bottom: 6px; }
.sec-title {
    font-size: 14.5px;
    font-weight: 600;
    letter-spacing: -.015em;
    color: var(--ink);
}
.sec-sub {
    font-size: 11.5px;
    color: var(--muted);
    margin-top: 2px;
}
.badge-3d {
    display: inline-block;
    margin-left: 8px;
    padding: 2px 7px;
    font-size: 9px;
    font-weight: 600;
    letter-spacing: .09em;
    color: var(--accent);
    background: var(--accent-soft);
    border: 1px solid #D6E7E1;
    border-radius: 999px;
    vertical-align: 2px;
}

/* ---------- Aviso de lectura rápida ---------- */
.insight {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 8px;
    padding: 12px 16px;
    background: var(--accent-soft);
    border: 1px solid #D6E7E1;
    border-radius: 12px;
    font-size: 12.5px;
    color: #34584E;
    margin: 16px 0 4px 0;
    line-height: 1.5;
}
.insight b { color: #0B5F4F; font-weight: 600; }

/* ---------- Estado vacío ---------- */
.empty {
    background: var(--card);
    border: 1px dashed #D8D8D1;
    border-radius: 16px;
    padding: 46px 24px;
    text-align: center;
    color: var(--ink);
    font-size: 15px;
    font-weight: 500;
    margin-top: 8px;
}
.empty span { display: block; color: var(--muted); font-weight: 400; font-size: 13px; margin-top: 6px; }

/* ---------- Pie ---------- */
.foot {
    margin-top: 34px;
    padding-top: 16px;
    border-top: 1px solid var(--line);
    font-size: 11.5px;
    color: #8A8A81;
    display: flex;
    justify-content: space-between;
    gap: 16px;
    flex-wrap: wrap;
}

/* ---------- Controles ---------- */
[data-testid="stWidgetLabel"] label,
[data-testid="stWidgetLabel"] label p,
[data-testid="stWidgetLabel"] p {
    font-size: 12px !important;
    font-weight: 500 !important;
    color: var(--muted) !important;
    opacity: 1 !important;
}
[data-baseweb="select"] > div {
    background: #FAFAF7 !important;
    border-color: var(--line) !important;
    border-radius: 10px !important;
    min-height: 41px;
    transition: border-color .15s ease;
}
[data-baseweb="select"] > div:hover { border-color: #C9C9C1 !important; }
[data-baseweb="input"] > div {
    background: #FAFAF7 !important;
    border-color: var(--line) !important;
    border-radius: 10px !important;
}
[data-baseweb="tag"] {
    background: #E5F1ED !important;
    border-radius: 6px !important;
}
[data-baseweb="tag"] span { color: #0B5F4F !important; font-weight: 500; }
[data-testid="stDateInput"] input { font-size: 13px; }

.stButton > button,
.stDownloadButton > button {
    border-radius: 10px;
    border: 1px solid var(--line);
    background: #FFFFFF;
    color: var(--ink);
    font-size: 12.5px;
    font-weight: 500;
    padding: .45rem .9rem;
    box-shadow: none;
    transition: all .15s ease;
}
.stButton > button:hover,
.stDownloadButton > button:hover {
    border-color: #BDBDB4;
    color: var(--ink);
    background: #FFFFFF;
}
.stDownloadButton > button[kind="primary"] {
    background: var(--ink);
    border-color: var(--ink);
    color: #FFFFFF;
}
.stDownloadButton > button[kind="primary"]:hover {
    background: var(--accent);
    border-color: var(--accent);
    color: #FFFFFF;
}

[data-testid="stPopover"] button {
    border-radius: 999px;
    border: 1px solid var(--line);
    background: #FFFFFF;
    font-size: 12.5px;
    font-weight: 500;
    color: var(--ink);
    padding: .4rem 1rem;
}
[data-testid="stPopover"] button:hover {
    border-color: #BDBDB4;
    color: var(--ink);
}
[data-testid="stPopoverBody"] {
    border-radius: 16px;
    border: 1px solid var(--line);
    box-shadow: 0 12px 32px rgba(23, 24, 26, .10);
}
[data-testid="stFileUploaderDropzone"] {
    background: #FAFAF7;
    border: 1px dashed #D8D8D1;
    border-radius: 12px;
    padding: 12px;
}
[data-testid="stFileUploaderDropzone"] button {
    border-radius: 8px;
    font-size: 12px;
}
[data-testid="stFileUploaderDropzoneInstructions"] span { font-size: 12px; }

div[data-testid="stDataFrame"] {
    border: 1px solid var(--line);
    border-radius: 12px;
    overflow: hidden;
}
[data-testid="stPlotlyChart"] { margin-top: 2px; }

@media (max-width: 640px) {
    .block-container { padding: 1.4rem 1rem 2.5rem 1rem; }
    .kpi { padding: 14px 16px; }
    .hero-sub { font-size: 13px; }
    /* En móvil ocultamos el menú de datos para no romper el encabezado */
    .hero-pills .pill { white-space: normal; }
}
</style>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Utilidades de datos
# ---------------------------------------------------------------------------

ALIASES = {
    "fecha": {"fecha", "date", "dia", "periodo", "fecha venta", "fecha de venta", "fecha operacion"},
    "producto": {"producto", "product", "item", "articulo", "descripcion", "producto vendido"},
    "vendedor": {"vendedor", "vendedora", "seller", "agente", "comercial", "vendedor a", "vendedor asignado"},
    "ventas": {"ventas", "venta", "ingresos", "ingreso", "total", "monto", "importe", "revenue", "valor", "total vendido"},
    "cantidad": {"cantidad", "unidades", "unidad", "units", "qty", "quantity", "cant", "piezas"},
    "precio": {"precio", "precio unitario", "price", "valor unitario", "p unitario", "precio de venta"},
}


def _norm(value: str) -> str:
    text = unicodedata.normalize("NFKD", str(value))
    text = text.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def _find_column(columns: dict, target: str):
    """Busca una columna por coincidencia exacta y luego parcial."""
    aliases = ALIASES[target]
    for original, normalized in columns.items():
        if normalized in aliases:
            return original
    for original, normalized in columns.items():
        if any(alias in normalized for alias in aliases):
            return original
    return None


def _read_table(payload: bytes, filename: str) -> pd.DataFrame:
    """Lee CSV (coma, punto y coma o tabulador) o Excel."""
    suffix = Path(filename).suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return pd.read_excel(io.BytesIO(payload))

    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return pd.read_csv(
                io.BytesIO(payload),
                sep=None,
                engine="python",
                encoding=encoding,
            )
        except UnicodeDecodeError as exc:
            last_error = exc
        except Exception as exc:  # pragma: no cover
            last_error = exc
            break
    raise ValueError(f"No se pudo leer el archivo: {last_error}")


def _parse_number(value) -> float:
    """Convierte '1.234,56', 'S/ 3.500' o '3500' a número."""
    if value is None:
        return math.nan
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)

    text = re.sub(r"[^0-9,.\-]", "", str(value)).strip()
    if not text or text in {"-", ".", ","}:
        return math.nan

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        if text.count(",") == 1 and len(text.split(",")[-1]) <= 2:
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif text.count(".") > 1:
        text = text.replace(".", "")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", text):
        text = text.replace(".", "")

    try:
        return float(text)
    except ValueError:
        return math.nan


def _pretty(value) -> str:
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "nat"}:
        return "Sin dato"
    return text.title() if text.islower() else text


def prepare_data(raw: pd.DataFrame) -> tuple[pd.DataFrame, bool, int]:
    """Normaliza columnas y tipos. Devuelve (df, tiene_fechas, filas_sin_fecha)."""
    if raw is None or raw.empty:
        raise ValueError("El archivo no contiene filas.")

    data = raw.copy()
    data.columns = [str(col).strip() for col in data.columns]
    normalized = {col: _norm(col) for col in data.columns}

    mapping = {}
    used_targets = set()
    for target in ALIASES:
        column = _find_column(normalized, target)
        if column is not None and target not in used_targets:
            mapping[column] = target
            used_targets.add(target)
    data = data.rename(columns=mapping)

    if "ventas" not in data.columns:
        if "precio" in data.columns and "cantidad" in data.columns:
            data["ventas"] = pd.to_numeric(data["precio"], errors="coerce") * pd.to_numeric(
                data["cantidad"], errors="coerce"
            )
        else:
            raise ValueError(
                "No encontré la columna de ventas. Se esperaba una columna como "
                "'Ventas', 'Ingresos', 'Total' o 'Monto'."
            )

    if "cantidad" not in data.columns:
        data["cantidad"] = 1
    if "producto" not in data.columns:
        data["producto"] = "General"
    if "vendedor" not in data.columns:
        data["vendedor"] = "Sin asignar"

    data["ventas"] = data["ventas"].map(_parse_number)
    data["cantidad"] = data["cantidad"].map(_parse_number)
    data = data.dropna(subset=["ventas"])
    data["cantidad"] = data["cantidad"].fillna(1)
    data.loc[data["cantidad"] < 0, "cantidad"] = 0

    data["producto"] = data["producto"].map(_pretty)
    data["vendedor"] = data["vendedor"].map(_pretty)

    has_dates = "fecha" in data.columns
    missing_dates = 0
    if has_dates:
        parsed = pd.to_datetime(data["fecha"], errors="coerce")
        if parsed.notna().sum() < max(3, len(parsed) * 0.4):
            alt = pd.to_datetime(data["fecha"], errors="coerce", dayfirst=True)
            if alt.notna().sum() > parsed.notna().sum():
                parsed = alt
        missing_dates = int(parsed.isna().sum())
        has_dates = bool(parsed.notna().any())
        data["fecha"] = parsed
    else:
        data["fecha"] = pd.NaT

    data = data.dropna(subset=["ventas"]).reset_index(drop=True)

    if has_dates:
        data = data.sort_values("fecha", na_position="last").reset_index(drop=True)

    data = data[["fecha", "producto", "vendedor", "ventas", "cantidad"]]

    if data.empty:
        raise ValueError("No quedaron filas válidas después de limpiar los datos.")

    return data, has_dates, missing_dates


def load_source(uploaded) -> tuple[pd.DataFrame, str, datetime | None]:
    """Devuelve (datos, etiqueta de fuente, última actualización)."""
    if uploaded is not None:
        raw = _read_table(uploaded.getvalue(), uploaded.name)
        file_id = getattr(uploaded, "file_id", None) or f"{uploaded.name}:{getattr(uploaded, 'size', 0)}"
        if st.session_state.get("_upload_id") != file_id:
            st.session_state["_upload_id"] = file_id
            st.session_state["_upload_stamp"] = datetime.now()
        return raw, uploaded.name, st.session_state.get("_upload_stamp", datetime.now())

    if CSV_PATH.exists() and CSV_PATH.stat().st_size > 0:
        raw = _read_table(CSV_PATH.read_bytes(), CSV_PATH.name)
        updated = datetime.fromtimestamp(CSV_PATH.stat().st_mtime)
        return raw, CSV_PATH.name, updated

    return pd.DataFrame(DEMO), "datos de ejemplo", None


# ---------------------------------------------------------------------------
# Formato
# ---------------------------------------------------------------------------


def money(value: float) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "—"
    if abs(value) >= 1_000_000:
        return f"S/ {value / 1_000_000:,.1f}M"
    return f"S/ {value:,.0f}"


def money_full(value: float) -> str:
    return f"S/ {value:,.0f}"


def units(value: float) -> str:
    value = float(value)
    if abs(value - round(value)) < 1e-9:
        return f"{int(round(value)):,}"
    return f"{value:,.1f}"


def _short_date(value) -> str:
    if pd.isna(value):
        return "—"
    return f"{value.day:02d} {MESES[value.month - 1]}"


def _long_date(value) -> str:
    if pd.isna(value):
        return "—"
    return f"{value.day} de {MESES_LARGOS[value.month - 1]} de {value.year}"


def _friendly_datetime(value: datetime | None) -> str:
    if value is None:
        return "Datos de ejemplo"
    if value.date() == date.today():
        return f"Actualizado hoy · {value:%H:%M}"
    return f"Actualizado {value.day} {MESES[value.month - 1]} · {value:%H:%M}"


def _safe(value) -> str:
    return html.escape(str(value))


def card_marker() -> None:
    """Marca invisible que convierte el contenedor de Streamlit en tarjeta."""
    st.markdown('<span class="card-marker"></span>', unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Gráficos
# ---------------------------------------------------------------------------


def _base_layout(fig: go.Figure, height: int = 310, showlegend: bool = False) -> go.Figure:
    fig.update_layout(
        height=height,
        margin=dict(l=6, r=6, t=10, b=18),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, system-ui, sans-serif", size=11.5, color=MUTED),
        showlegend=showlegend,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.24,
            xanchor="center",
            x=0.5,
            font=dict(size=11, color=MUTED),
            bgcolor="rgba(0,0,0,0)",
            title=None,
        ),
        hoverlabel=dict(
            bgcolor=INK,
            bordercolor=INK,
            font=dict(color="#FFFFFF", size=12, family="Inter, sans-serif"),
        ),
        hovermode="closest",
        separators=".,",
    )
    fig.update_xaxes(
        showgrid=False,
        zeroline=False,
        showline=False,
        ticks="",
        title=None,
        tickfont=dict(size=11, color=MUTED),
        automargin=True,
    )
    fig.update_yaxes(
        showgrid=True,
        gridcolor="#EEEEE9",
        gridwidth=1,
        zeroline=False,
        showline=False,
        ticks="",
        title=None,
        tickfont=dict(size=11, color=MUTED),
        automargin=True,
    )
    return fig


def chart_trend(daily: pd.DataFrame) -> go.Figure:
    labels = [f"{row.fecha.day:02d}/{row.fecha.month:02d}/{row.fecha.year}" for row in daily.itertuples()]
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=daily["fecha"],
            y=daily["ventas"],
            mode="lines+markers",
            line=dict(color=ACCENT, width=2.4),
            marker=dict(size=7, color="#FFFFFF", line=dict(color=ACCENT, width=2)),
            fill="tozeroy",
            fillcolor="rgba(14, 124, 102, 0.07)",
            customdata=list(zip(labels, daily["unidades"])),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Ingresos&nbsp;&nbsp;S/ %{y:,.0f}<br>"
                "Unidades&nbsp;&nbsp;%{customdata[1]}<extra></extra>"
            ),
        )
    )

    average = float(daily["ventas"].mean())
    fig.add_hline(
        y=average,
        line_dash="dot",
        line_color="#C9C9C2",
        line_width=1,
        annotation_text=f"promedio {money(average)}",
        annotation_position="top right",
        annotation_font=dict(size=10, color="#8A8A81"),
    )

    step = max(1, math.ceil(len(daily) / 6))
    ticks = daily.iloc[::step]
    if ticks["fecha"].iloc[-1] != daily["fecha"].iloc[-1]:
        ticks = pd.concat([ticks, daily.iloc[[-1]]])
    fig.update_xaxes(
        tickvals=ticks["fecha"],
        ticktext=[_short_date(value) for value in ticks["fecha"]],
    )

    if len(daily) > 1:
        span = daily["fecha"].max() - daily["fecha"].min()
    else:
        span = pd.Timedelta(days=1)
    pad = max(span * 0.03, pd.Timedelta(hours=6))
    fig.update_xaxes(range=[daily["fecha"].min() - pad, daily["fecha"].max() + pad])

    fig.update_yaxes(rangemode="tozero", tickprefix="S/ ", tickformat=",.0f")
    return _base_layout(fig, 315)


def chart_donut(products: pd.DataFrame, total: float) -> go.Figure:
    fig = px.pie(
        products,
        names="producto",
        values="ventas",
        hole=0.72,
        color="producto",
        color_discrete_sequence=PALETTE,
        category_orders={"producto": products["producto"].tolist()},
    )
    fig.update_traces(
        textinfo="none",
        sort=False,
        marker=dict(line=dict(color="#FFFFFF", width=3)),
        hovertemplate="<b>%{label}</b><br>S/ %{value:,.0f}<br>%{percent}<extra></extra>",
    )
    fig.add_annotation(
        x=0.5,
        y=0.5,
        text=(
            f"<span style='font-size:10.5px;letter-spacing:.08em;color:{MUTED}'>INGRESOS</span><br>"
            f"<b style='font-size:19px;color:{INK}'>{money(total)}</b>"
        ),
        showarrow=False,
        font=dict(family="Inter, sans-serif"),
    )
    return _base_layout(fig, 315, showlegend=True)


def chart_products(products: pd.DataFrame) -> go.Figure:
    data = products.sort_values("ventas").reset_index(drop=True)
    if len(data) <= 1:
        colors = [ACCENT]
    else:
        colors = sample_colorscale(
            [[0.0, "#DDEBE6"], [1.0, ACCENT]],
            [index / (len(data) - 1) for index in range(len(data))],
        )
    fig = go.Figure(
        go.Bar(
            x=data["ventas"],
            y=data["producto"],
            orientation="h",
            marker=dict(color=colors, line=dict(width=0)),
            text=[money_full(value) for value in data["ventas"]],
            textposition="outside",
            cliponaxis=False,
            textfont=dict(size=11, color=MUTED),
            customdata=data["unidades"],
            hovertemplate=(
                "<b>%{y}</b><br>"
                "Ingresos&nbsp;&nbsp;S/ %{x:,.0f}<br>"
                "Unidades&nbsp;&nbsp;%{customdata}<extra></extra>"
            ),
        )
    )
    top = float(data["ventas"].max()) if len(data) else 1
    fig.update_xaxes(visible=False, range=[0, top * 1.28])
    fig.update_yaxes(
        showgrid=False,
        tickfont=dict(size=12, color="#3F3F3B"),
        type="category",
    )
    fig.update_layout(bargap=0.45)
    return _base_layout(fig, 315)


def chart_sellers(sellers: pd.DataFrame, products: pd.DataFrame) -> go.Figure:
    totals = sellers.groupby("vendedor", as_index=False)["ventas"].sum()
    totals = totals.sort_values("ventas")
    order = totals["vendedor"].tolist()

    fig = px.bar(
        sellers,
        x="ventas",
        y="vendedor",
        color="producto",
        orientation="h",
        barmode="stack",
        color_discrete_sequence=PALETTE,
        category_orders={
            "vendedor": order,
            "producto": products["producto"].tolist(),
        },
    )
    fig.update_traces(
        marker_line_width=0,
        hovertemplate="<b>%{y}</b> · %{fullData.name}<br>S/ %{x:,.0f}<extra></extra>",
    )

    for row in totals.itertuples():
        fig.add_annotation(
            x=row.ventas,
            y=row.vendedor,
            text=f"  {money(row.ventas)}",
            showarrow=False,
            xanchor="left",
            yanchor="middle",
            font=dict(size=11, color=MUTED),
        )

    top = float(totals["ventas"].max()) if len(totals) else 1
    fig.update_xaxes(visible=False, range=[0, top * 1.28])
    fig.update_yaxes(
        showgrid=False,
        tickfont=dict(size=12, color="#3F3F3B"),
        type="category",
    )
    fig.update_layout(bargap=0.5)
    return _base_layout(fig, 315, showlegend=True)


def chart_3d(data: pd.DataFrame, has_dates: bool) -> go.Figure:
    """Dispersión 3D: fecha/operación, unidades e ingresos."""
    frame = data.copy()

    if has_dates and frame["fecha"].notna().any():
        frame = frame.dropna(subset=["fecha"]).copy()
        frame["dia"] = frame["fecha"].dt.normalize()
        origin = frame["dia"].min()
        frame["x"] = (frame["dia"] - origin).dt.days.astype(int)
        frame["fecha_label"] = frame["dia"].map(_short_date)

        unique_x = sorted(frame["x"].unique())
        step = max(1, math.ceil(len(unique_x) / 6))
        tick_x = unique_x[::step]
        if tick_x[-1] != unique_x[-1]:
            tick_x.append(unique_x[-1])
        tick_text = [
            _short_date(origin + pd.Timedelta(days=int(value))) for value in tick_x
        ]
        x_title = "Fecha"
    else:
        frame = frame.reset_index(drop=True)
        frame["x"] = frame.index
        frame["fecha_label"] = [f"Operación {index + 1}" for index in frame.index]

        unique_x = sorted(frame["x"].unique())
        step = max(1, math.ceil(len(unique_x) / 6))
        tick_x = unique_x[::step]
        if tick_x[-1] != unique_x[-1]:
            tick_x.append(unique_x[-1])
        tick_text = [f"#{int(value) + 1}" for value in tick_x]
        x_title = "Operación"

    order = (
        frame.groupby("producto")["ventas"]
        .sum()
        .sort_values(ascending=False)
        .index.tolist()
    )

    low, high = float(frame["ventas"].min()), float(frame["ventas"].max())
    if high > low:
        sizes = 7 + 9 * (frame["ventas"] - low) / (high - low)
    else:
        sizes = pd.Series(11.0, index=frame.index)
    frame["_size"] = sizes

    axis_style = dict(
        showbackground=True,
        backgroundcolor="#FAFAF7",
        gridcolor="#E9E9E3",
        gridwidth=1,
        zerolinecolor="#E1E1D9",
        showline=True,
        linecolor="#DCDCD4",
        tickfont=dict(size=9, color=MUTED),
    )

    fig = go.Figure()
    for position, product in enumerate(order):
        subset = frame[frame["producto"] == product]
        fig.add_trace(
            go.Scatter3d(
                x=subset["x"],
                y=subset["cantidad"],
                z=subset["ventas"],
                mode="markers",
                name=str(product),
                marker=dict(
                    size=subset["_size"],
                    color=PALETTE[position % len(PALETTE)],
                    opacity=0.9,
                    line=dict(color="#FFFFFF", width=1),
                ),
                customdata=list(
                    zip(subset["fecha_label"], subset["vendedor"], subset["producto"])
                ),
                hovertemplate=(
                    "<b>%{customdata[2]}</b> · %{customdata[0]}<br>"
                    "Vendedor · %{customdata[1]}<br>"
                    "Unidades · %{y}<br>"
                    "Ingresos · S/ %{z:,.0f}<extra></extra>"
                ),
            )
        )

    fig.update_layout(
        height=470,
        margin=dict(l=0, r=0, t=4, b=0),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, system-ui, sans-serif", size=11, color=MUTED),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=0.005,
            xanchor="center",
            x=0.5,
            font=dict(size=11, color=MUTED),
            bgcolor="rgba(0,0,0,0)",
            title=None,
        ),
        hoverlabel=dict(
            bgcolor=INK,
            bordercolor=INK,
            font=dict(color="#FFFFFF", size=12, family="Inter, sans-serif"),
        ),
        scene=dict(
            domain=dict(x=[0, 1], y=[0.09, 1]),
            xaxis={
                **axis_style,
                "title": dict(text=x_title, font=dict(size=10, color="#8A8A81")),
                "tickvals": tick_x,
                "ticktext": tick_text,
            },
            yaxis={
                **axis_style,
                "title": dict(text="Unidades", font=dict(size=10, color="#8A8A81")),
            },
            zaxis={
                **axis_style,
                "title": dict(text="Ingresos (S/)", font=dict(size=10, color="#8A8A81")),
                "tickprefix": "S/ ",
                "tickformat": ",.0f",
            },
            camera=dict(eye=dict(x=1.55, y=-1.55, z=0.95)),
            aspectmode="manual",
            aspectratio=dict(x=1.55, y=1.0, z=0.9),
            dragmode="orbit",
        ),
    )
    return fig


# ---------------------------------------------------------------------------
# Encabezado + carga de datos
# ---------------------------------------------------------------------------

header_left, header_right = st.columns([5.4, 1.5], vertical_alignment="bottom")

with header_right:
    with st.popover("Cargar datos", width="stretch"):
        uploaded_file = st.file_uploader(
            "Sube tu archivo",
            type=["csv", "xlsx"],
            label_visibility="collapsed",
        )
        st.caption(
            "CSV o Excel con columnas como **Fecha**, **Producto**, "
            "**Vendedor**, **Ventas** y **Cantidad**. Los nombres se detectan "
            "automáticamente y no se necesita código."
        )
        st.caption("Fuente actual: `ventas.csv` si tiene datos; si no, ejemplo.")

try:
    raw_data, source_name, updated_at = load_source(uploaded_file)
    data, has_dates, missing_dates = prepare_data(raw_data)
except Exception as exc:  # pragma: no cover
    with header_left:
        st.markdown(
            "<div class='hero-eyebrow'>Panel comercial</div>"
            "<div class='hero-title'>Rendimiento comercial</div>",
            unsafe_allow_html=True,
        )
    st.error(
        "No pudimos leer los datos. Revisa que el archivo tenga una columna de "
        "ventas (por ejemplo, **Ventas** o **Ingresos**) y que sea un CSV o Excel válido."
    )
    st.caption(f"Detalle técnico: {exc}")
    st.stop()

# Reinicia filtros cuando cambia el archivo de datos
signature = f"{source_name}|{len(raw_data)}|{raw_data.shape[1]}|{updated_at}"
if st.session_state.get("_data_signature") != signature:
    st.session_state["_data_signature"] = signature
    for key in ("flt_fechas", "flt_productos", "flt_vendedores"):
        st.session_state.pop(key, None)

source_label = "Datos de ejemplo" if updated_at is None else source_name
source_display = source_label if len(source_label) <= 34 else source_label[:31] + "..."

if has_dates and data["fecha"].notna().any():
    first_date = data["fecha"].min()
    last_date = data["fecha"].max()
    if first_date.date() == last_date.date():
        period_label = _long_date(last_date)
    elif first_date.year != last_date.year:
        period_label = f"{_short_date(first_date)} {first_date.year} – {_short_date(last_date)} {last_date.year}"
    elif first_date.month == last_date.month:
        period_label = f"{first_date.day} – {last_date.day} de {MESES_LARGOS[last_date.month - 1]} {last_date.year}"
    else:
        period_label = f"{_short_date(first_date)} – {_short_date(last_date)} {last_date.year}"
else:
    period_label = f"{len(data)} operaciones"

with header_left:
    st.markdown(
        f"""
<div>
  <div class="hero-eyebrow">Panel comercial</div>
  <div class="hero-title">Rendimiento comercial</div>
  <p class="hero-sub">Ingresos, productos y desempeño del equipo en una sola vista.
  Filtra y explora sin escribir código.</p>
  <div class="hero-pills">
    <span class="pill pill-accent" title="{_safe(source_label)}"><span class="dot"></span>Fuente · {_safe(source_display)}</span>
    <span class="pill">{_safe(period_label)}</span>
    <span class="pill">{_friendly_datetime(updated_at)}</span>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------------------------
# Filtros
# ---------------------------------------------------------------------------


def _clear_filters() -> None:
    for key in ("flt_fechas", "flt_productos", "flt_vendedores"):
        st.session_state.pop(key, None)


with st.container(border=True):
    card_marker()
    if has_dates:
        f1, f2, f3, f4 = st.columns([1.25, 1.1, 1.1, 0.55], vertical_alignment="bottom")
    else:
        f1, f2, f3, f4 = st.columns([1.15, 1.1, 1.1, 0.55], vertical_alignment="bottom")

    if has_dates:
        with f1:
            min_day = data["fecha"].min().date()
            max_day = data["fecha"].max().date()
            selected_range = st.date_input(
                "Periodo",
                value=(min_day, max_day),
                min_value=min_day,
                max_value=max_day,
                format="DD/MM/YYYY",
                key="flt_fechas",
            )
        with f2:
            selected_products = st.multiselect(
                "Producto",
                options=sorted(data["producto"].unique()),
                placeholder="Todos los productos",
                key="flt_productos",
            )
        with f3:
            selected_sellers = st.multiselect(
                "Vendedor",
                options=sorted(data["vendedor"].unique()),
                placeholder="Todo el equipo",
                key="flt_vendedores",
            )
        with f4:
            st.button("Limpiar", on_click=_clear_filters, width="stretch")
    else:
        with f1:
            st.markdown(
                "<div style='font-size:12px;color:#6E6E66;padding:10px 0'>"
                "El archivo no incluye fechas.<br>Los filtros de periodo no están disponibles.</div>",
                unsafe_allow_html=True,
            )
        with f2:
            selected_products = st.multiselect(
                "Producto",
                options=sorted(data["producto"].unique()),
                placeholder="Todos los productos",
                key="flt_productos",
            )
        with f3:
            selected_sellers = st.multiselect(
                "Vendedor",
                options=sorted(data["vendedor"].unique()),
                placeholder="Todo el equipo",
                key="flt_vendedores",
            )
        with f4:
            st.button("Limpiar", on_click=_clear_filters, width="stretch")
        selected_range = None


filtered = data.copy()

if has_dates and selected_range:
    if isinstance(selected_range, (tuple, list)):
        start, end = selected_range[0], selected_range[-1]
    else:
        start = end = selected_range
    in_range = filtered["fecha"].dt.date.between(start, end, inclusive="both")
    filtered = filtered[in_range | filtered["fecha"].isna()]

if selected_products:
    filtered = filtered[filtered["producto"].isin(selected_products)]
if selected_sellers:
    filtered = filtered[filtered["vendedor"].isin(selected_sellers)]

filtered = filtered.reset_index(drop=True)

if filtered.empty:
    st.markdown(
        "<div class='empty'>No hay operaciones con los filtros seleccionados."
        "<span>Prueba ampliando el periodo o limpiando los filtros.</span></div>",
        unsafe_allow_html=True,
    )
    st.stop()

# ---------------------------------------------------------------------------
# Indicadores
# ---------------------------------------------------------------------------

total_revenue = float(filtered["ventas"].sum())
total_units = float(filtered["cantidad"].sum())
operations = len(filtered)
ticket = total_revenue / operations if operations else 0.0

by_product = (
    filtered.groupby("producto", as_index=False)
    .agg(ventas=("ventas", "sum"), unidades=("cantidad", "sum"))
    .sort_values("ventas", ascending=False)
)
by_seller = (
    filtered.groupby("vendedor", as_index=False)
    .agg(ventas=("ventas", "sum"), unidades=("cantidad", "sum"))
    .sort_values("ventas", ascending=False)
)

top_product = by_product.iloc[0]
top_seller = by_seller.iloc[0]
top_product_share = (top_product["ventas"] / total_revenue * 100) if total_revenue else 0
top_seller_share = (top_seller["ventas"] / total_revenue * 100) if total_revenue else 0
units_leader = by_product.sort_values("unidades", ascending=False).iloc[0]
product_count = filtered["producto"].nunique()

kpi1, kpi2, kpi3, kpi4 = st.columns(4, gap="small")

with kpi1:
    st.markdown(
        f"""
<div class="kpi">
  <div class="kpi-label">Ingresos</div>
  <div class="kpi-value accent">{money(total_revenue)}</div>
  <div class="kpi-foot">{operations} operaciones · {product_count} productos</div>
</div>
""",
        unsafe_allow_html=True,
    )

with kpi2:
    st.markdown(
        f"""
<div class="kpi">
  <div class="kpi-label">Unidades vendidas</div>
  <div class="kpi-value">{units(total_units)}</div>
  <div class="kpi-foot">{units(total_units / operations) if operations else "0"} por operación</div>
</div>
""",
        unsafe_allow_html=True,
    )

with kpi3:
    st.markdown(
        f"""
<div class="kpi">
  <div class="kpi-label">Ticket promedio</div>
  <div class="kpi-value">{money_full(ticket)}</div>
  <div class="kpi-foot">ingreso medio por operación</div>
</div>
""",
        unsafe_allow_html=True,
    )

with kpi4:
    st.markdown(
        f"""
<div class="kpi">
  <div class="kpi-label">Mejor vendedor</div>
  <div class="kpi-value">{_safe(top_seller["vendedor"])}</div>
  <div class="kpi-foot">{money(top_seller["ventas"])} · {top_seller_share:.0f}% del total</div>
</div>
""",
        unsafe_allow_html=True,
    )

# Lectura rápida
daily = pd.DataFrame()
if has_dates and filtered["fecha"].notna().any():
    dated = filtered.dropna(subset=["fecha"]).copy()
    dated["dia"] = dated["fecha"].dt.normalize()
    daily = (
        dated.groupby("dia", as_index=False)
        .agg(ventas=("ventas", "sum"), unidades=("cantidad", "sum"))
        .rename(columns={"dia": "fecha"})
    )

insight_parts = []
if total_revenue:
    insight_parts.append(
        f"<b>{_safe(top_product['producto'])}</b> lidera los ingresos ({top_product_share:.0f}%)"
    )
insight_parts.append(
    f"<b>{_safe(units_leader['producto'])}</b> vende más unidades ({units(units_leader['unidades'])})"
)
if not daily.empty:
    best_day = daily.sort_values("ventas").iloc[-1]
    insight_parts.append(
        f"mejor día: <b>{_safe(_short_date(best_day['fecha']))}</b> ({money(best_day['ventas'])})"
    )
insight_parts.append(f"<b>{_safe(top_seller['vendedor'])}</b> concentra {top_seller_share:.0f}% del equipo")

st.markdown(
    "<div class='insight'>" + " · ".join(insight_parts) + "</div>",
    unsafe_allow_html=True,
)

if missing_dates:
    st.caption(f"Se omitieron {missing_dates} fila(s) sin fecha válida al calcular la tendencia.")

# ---------------------------------------------------------------------------
# Gráficos
# ---------------------------------------------------------------------------

row1_left, row1_right = st.columns([1.75, 1], gap="small")

with row1_left:
    with st.container(border=True):
        card_marker()
        st.markdown(
            "<div class='sec'><div class='sec-title'>Evolución de ingresos</div>"
            "<div class='sec-sub'>Ventas diarias del periodo seleccionado</div></div>",
            unsafe_allow_html=True,
        )
        if daily.empty:
            st.markdown(
                "<div style='height:315px;display:flex;align-items:center;justify-content:center;"
                "color:#8A8A81;font-size:13px;text-align:center'>"
                "La tendencia diaria necesita una columna de fechas válida.</div>",
                unsafe_allow_html=True,
            )
        else:
            st.plotly_chart(
                chart_trend(daily),
                theme=None,
                config=PLOT_CONFIG,
                key="chart_trend",
            )

with row1_right:
    with st.container(border=True):
        card_marker()
        st.markdown(
            "<div class='sec'><div class='sec-title'>Participación por producto</div>"
            "<div class='sec-sub'>Mezcla de ingresos</div></div>",
            unsafe_allow_html=True,
        )
        st.plotly_chart(
            chart_donut(by_product, total_revenue),
            theme=None,
            config=PLOT_CONFIG,
            key="chart_donut",
        )

row2_left, row2_right = st.columns(2, gap="small")

with row2_left:
    with st.container(border=True):
        card_marker()
        st.markdown(
            "<div class='sec'><div class='sec-title'>Ingresos por producto</div>"
            "<div class='sec-sub'>Ordenado de mayor a menor</div></div>",
            unsafe_allow_html=True,
        )
        st.plotly_chart(
            chart_products(by_product),
            theme=None,
            config=PLOT_CONFIG,
            key="chart_products",
        )

with row2_right:
    with st.container(border=True):
        card_marker()
        st.markdown(
            "<div class='sec'><div class='sec-title'>Ingresos por vendedor</div>"
            "<div class='sec-sub'>Composición por producto</div></div>",
            unsafe_allow_html=True,
        )
        st.plotly_chart(
            chart_sellers(filtered, by_product),
            theme=None,
            config=PLOT_CONFIG,
            key="chart_sellers",
        )

with st.container(border=True):
    card_marker()
    st.markdown(
        "<div class='sec'><div class='sec-title'>Vista 3D de operaciones"
        "<span class='badge-3d'>3D</span></div>"
        "<div class='sec-sub'>Un punto por operación · arrastra para girar · doble clic para restablecer</div></div>",
        unsafe_allow_html=True,
    )
    st.plotly_chart(
        chart_3d(filtered, has_dates),
        theme=None,
        config=PLOT_CONFIG,
        key="chart_3d",
    )

# ---------------------------------------------------------------------------
# Tabla + descargas
# ---------------------------------------------------------------------------

export = filtered.copy()
export.insert(0, "Fecha", export["fecha"].dt.strftime("%d/%m/%Y") if has_dates else "—")
export = export.rename(
    columns={
        "producto": "Producto",
        "vendedor": "Vendedor",
        "cantidad": "Cantidad",
        "ventas": "Ingresos",
    }
)
export["Precio unitario"] = export["Ingresos"] / export["Cantidad"].replace(0, pd.NA)
export = export[["Fecha", "Producto", "Vendedor", "Cantidad", "Precio unitario", "Ingresos"]]

with st.container(border=True):
    card_marker()

    head_left, head_right = st.columns([2.6, 1.4], vertical_alignment="center")

    with head_left:
        st.markdown(
            "<div class='sec'><div class='sec-title'>Detalle de operaciones</div>"
            f"<div class='sec-sub'>Mostrando {operations:,} de {len(data):,} registros</div></div>",
            unsafe_allow_html=True,
        )

    with head_right:
        dl1, dl2 = st.columns(2, gap="small")
        with dl1:
            if HAS_OPENPYXL:
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                    export.to_excel(writer, index=False, sheet_name="Ventas")
                st.download_button(
                    "Excel",
                    data=excel_buffer.getvalue(),
                    file_name="reporte_ventas.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                )
        with dl2:
            st.download_button(
                "CSV",
                data=export.to_csv(index=False).encode("utf-8-sig"),
                file_name="reporte_ventas.csv",
                mime="text/csv",
                type="primary",
                width="stretch",
            )

    if has_dates:
        display = export.copy()
        display["Fecha"] = pd.to_datetime(filtered["fecha"]).dt.date
        display.insert(2, "_orden", filtered["fecha"].values)
        display = display.sort_values("_orden", ascending=False).drop(columns="_orden")
    else:
        display = export.copy()

    fecha_config = (
        st.column_config.DateColumn("Fecha", format="DD/MM/YYYY", width="small")
        if has_dates
        else st.column_config.TextColumn("Fecha", width="small")
    )

    quantities_are_int = bool(
        pd.notna(display["Cantidad"]).all()
        and all(float(value).is_integer() for value in display["Cantidad"])
    )
    qty_format = "%d" if quantities_are_int else "%.2f"

    height = min(430, 70 + 35 * len(display))
    st.dataframe(
        display,
        width="stretch",
        hide_index=True,
        height=height,
        column_config={
            "Fecha": fecha_config,
            "Producto": st.column_config.TextColumn("Producto", width="medium"),
            "Vendedor": st.column_config.TextColumn("Vendedor", width="medium"),
            "Cantidad": st.column_config.NumberColumn("Cantidad", format=qty_format, width="small"),
            "Precio unitario": st.column_config.NumberColumn("Precio unitario", format="S/ %.0f", width="small"),
            "Ingresos": st.column_config.NumberColumn("Ingresos", format="S/ %d", width="small"),
        },
    )

# ---------------------------------------------------------------------------
# Pie
# ---------------------------------------------------------------------------

st.markdown(
    f"""
<div class="foot">
  <span>Dashboard comercial · {_safe(source_label)}</span>
  <span>Actualizado con {len(data):,} operaciones · Streamlit</span>
</div>
""",
    unsafe_allow_html=True,
)
