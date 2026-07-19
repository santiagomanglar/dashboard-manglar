"""
Dashboard Manglar — 2026
Tablero financiero para los socios.

Ejecutar en local:
    streamlit run app.py

Fuente de datos: por defecto lee el modelo en Excel.
Para conectarlo a Google Sheets en vivo, ver README.md
"""

import io
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ─────────────────────────── Configuración ───────────────────────────

st.set_page_config(
    page_title="Dashboard Manglar — 2026",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="collapsed",
)

ARCHIVO_MODELO = "Manglar_Modelo_Financiero.xlsx"
HOJA_MOV = "Extractos bancarios - Manglar"
HOJA_DATA = "Data"

MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre",
    12: "Diciembre",
}

# Paleta sobria
TINTA = "#1f2937"
GRIS = "#6b7280"
GRIS_SUAVE = "#e5e7eb"
VERDE = "#3f6f5b"
ROJO = "#a4553f"
AZUL = "#41607f"
ARENA = "#c9a227"

# ─────────────────────────── Estilos ───────────────────────────

st.markdown(
    """
    <style>
      #MainMenu, footer, header {visibility: hidden;}
      .block-container {padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1280px;}

      .titulo {font-size: 1.8rem; font-weight: 650; color: #1f2937;
               letter-spacing: -0.02em; margin-bottom: .15rem;}
      .subtitulo {font-size: .92rem; color: #6b7280; margin-bottom: 1.6rem;}

      .tarjeta {background: #ffffff; border: 1px solid #e5e7eb; border-radius: 10px;
                padding: 1.1rem 1.25rem; height: 100%;}
      .tarjeta-titulo {font-size: .78rem; font-weight: 600; color: #6b7280;
                       text-transform: uppercase; letter-spacing: .06em;
                       margin-bottom: .55rem;}
      .kpi-valor {font-size: 1.55rem; font-weight: 650; color: #1f2937;
                  letter-spacing: -0.02em; line-height: 1.15;}
      .kpi-nota {font-size: .78rem; color: #6b7280; margin-top: .25rem;}

      .seccion {font-size: 1.02rem; font-weight: 620; color: #1f2937;
                margin: 1.9rem 0 .7rem 0;}

      div[data-testid="stMetric"] {background: #fff; border: 1px solid #e5e7eb;
                                   border-radius: 10px; padding: .9rem 1.1rem;}
      .stDataFrame {border: 1px solid #e5e7eb; border-radius: 10px;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ─────────────────────────── Carga de datos ───────────────────────────

@st.cache_data(show_spinner=False)
def cargar(origen) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Lee movimientos y el mapeo de categorías del modelo."""
    mov = pd.read_excel(origen, sheet_name=HOJA_MOV)
    data = pd.read_excel(origen, sheet_name=HOJA_DATA)
    return mov, data


def preparar(mov: pd.DataFrame, data: pd.DataFrame) -> pd.DataFrame:
    """Limpia los movimientos y recalcula la cuenta desde el mapeo."""
    df = mov.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Nombres de columna del modelo -> nombres internos
    equivalencias = {
        "Fecha": "fecha",
        "Descripción banco": "descripcion",
        "Valor": "valor",
        "Categoría": "categoria",
        "Mes Caja": "mes_caja",
        "Cliente": "cliente",
        "Proyecto": "proyecto",
        "IVA?": "iva",
        "Mes (Causación P&G)": "mes_causacion",
    }
    df = df.rename(columns={k: v for k, v in equivalencias.items() if k in df.columns})

    faltan = [n for n in ("valor", "categoria", "mes_caja") if n not in df.columns]
    if faltan:
        st.error(
            "El archivo no tiene las columnas esperadas: "
            + ", ".join(faltan)
            + ". Verifique que sea el modelo financiero de Manglar."
        )
        st.stop()

    for col in ("cliente", "proyecto", "iva", "mes_causacion", "descripcion", "fecha"):
        if col not in df.columns:
            df[col] = np.nan

    df = df[pd.to_numeric(df["valor"], errors="coerce").notna()].copy()
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")

    # Mapeo categoría -> cuenta (más confiable que la fórmula en caché)
    mapa = (
        data.dropna(subset=[data.columns[0]])
        .drop_duplicates(subset=[data.columns[0]])
        .set_index(data.columns[0])[data.columns[1]]
        .to_dict()
    )
    df["cuenta"] = df["categoria"].map(mapa).fillna("Revisar")

    for c in ("mes_caja", "mes_causacion"):
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df["iva"] = df["iva"].astype(str).str.strip()
    df["cliente"] = df["cliente"].fillna("Sin asignar").astype(str).str.strip()
    df["proyecto"] = df["proyecto"].fillna("Sin asignar").astype(str).str.strip()
    df.loc[df["cliente"].isin(["nan", "NA", ""]), "cliente"] = "Sin asignar"
    df.loc[df["proyecto"].isin(["nan", "NA", ""]), "proyecto"] = "Sin asignar"

    # Valor neto de IVA (para P&G y rentabilidad)
    df["valor_neto"] = np.where(df["iva"].eq("Si"), df["valor"] / 1.19, df["valor"])
    return df


def money(x, corto: bool = False) -> str:
    if pd.isna(x):
        return "—"
    if corto:
        a = abs(x)
        if a >= 1_000_000_000:
            return f"${x/1_000_000_000:,.1f}B"
        if a >= 1_000_000:
            return f"${x/1_000_000:,.1f}M"
        if a >= 1_000:
            return f"${x/1_000:,.0f}K"
    return f"${x:,.0f}"


def tarjeta(titulo: str, valor: str, nota: str = "") -> str:
    return (
        f'<div class="tarjeta"><div class="tarjeta-titulo">{titulo}</div>'
        f'<div class="kpi-valor">{valor}</div>'
        f'<div class="kpi-nota">{nota}</div></div>'
    )


def base_layout(fig: go.Figure, alto: int = 300) -> go.Figure:
    fig.update_layout(
        height=alto,
        margin=dict(l=8, r=8, t=28, b=8),
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(family="ui-sans-serif, system-ui, sans-serif", size=12, color=TINTA),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=11)),
        hoverlabel=dict(bgcolor="white", font_size=12),
    )
    fig.update_xaxes(showgrid=False, linecolor=GRIS_SUAVE, tickfont=dict(size=11))
    fig.update_yaxes(gridcolor=GRIS_SUAVE, zerolinecolor=GRIS_SUAVE, tickfont=dict(size=11))
    return fig


# ─────────────────────────── Encabezado ───────────────────────────

st.markdown('<div class="titulo">Dashboard Manglar — 2026</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitulo">Agencia BTL, material POP y merchandising · Resumen financiero</div>',
    unsafe_allow_html=True,
)

# Origen de datos
try:
    mov_raw, data_raw = cargar(ARCHIVO_MODELO)
except Exception:
    st.info("Cargue el modelo financiero para ver el tablero.")
    subido = st.file_uploader("Modelo financiero (.xlsx)", type=["xlsx"])
    if subido is None:
        st.stop()
    mov_raw, data_raw = cargar(io.BytesIO(subido.getvalue()))

df = preparar(mov_raw, data_raw)

# ─────────────────────────── Filtros ───────────────────────────

meses_disp = sorted(int(m) for m in df["mes_caja"].dropna().unique())
clientes_disp = sorted(
    c for c in df["cliente"].unique() if c not in ("Sin asignar", "NA", "nan")
)

c1, c2 = st.columns([2, 2])
with c1:
    sel_meses = st.multiselect(
        "Mes", options=meses_disp, default=meses_disp,
        format_func=lambda m: MESES.get(m, str(m)),
    )
with c2:
    sel_clientes = st.multiselect("Cliente", options=clientes_disp, default=clientes_disp)

if not sel_meses:
    sel_meses = meses_disp
if not sel_clientes:
    sel_clientes = clientes_disp

f_caja = df[df["mes_caja"].isin(sel_meses)]
f_pyg = df[df["mes_causacion"].isin(sel_meses)]
f_cli = f_pyg[f_pyg["cliente"].isin(sel_clientes)]

# ─────────────────────────── Indicadores ───────────────────────────

ingresos = f_pyg.loc[f_pyg["cuenta"].eq("Ingresos"), "valor_neto"].sum()
costo_ventas = f_pyg.loc[f_pyg["cuenta"].eq("Costo Ventas"), "valor_neto"].sum()
gastos_op = f_pyg.loc[
    f_pyg["cuenta"].isin(["Costos Fijos", "Costos Operacionales", "Gastos Variables"]),
    "valor_neto",
].sum()

margen_bruto = ingresos + costo_ventas
resultado = margen_bruto + gastos_op
margen_pct = (margen_bruto / ingresos * 100) if ingresos else 0
saldo_caja = df["valor"].sum()  # caja total, sin filtrar

k1, k2, k3, k4 = st.columns(4)
k1.markdown(tarjeta("Saldo en caja", money(saldo_caja, True), "Acumulado a la fecha"),
            unsafe_allow_html=True)
k2.markdown(tarjeta("Ingresos", money(ingresos, True), "Netos de IVA"),
            unsafe_allow_html=True)
k3.markdown(tarjeta("Margen bruto", money(margen_bruto, True), f"{margen_pct:.1f}% sobre ingresos"),
            unsafe_allow_html=True)
k4.markdown(tarjeta("Resultado operativo", money(resultado, True), "Después de gastos"),
            unsafe_allow_html=True)

# ─────────────────────── Ingresos, costos y caja ───────────────────────

st.markdown('<div class="seccion">Resultado y caja</div>', unsafe_allow_html=True)
g1, g2 = st.columns(2)

with g1:
    pyg = (
        f_pyg.assign(
            grupo=np.select(
                [f_pyg["cuenta"].eq("Ingresos"), f_pyg["cuenta"].eq("Costo Ventas")],
                ["Ingresos", "Costo de ventas"],
                default="Gastos",
            )
        )
        .groupby(["mes_causacion", "grupo"])["valor_neto"].sum().reset_index()
    )
    fig = go.Figure()
    for nombre, color in [("Ingresos", VERDE), ("Costo de ventas", ROJO), ("Gastos", ARENA)]:
        sub = pyg[pyg["grupo"].eq(nombre)]
        fig.add_bar(
            x=[MESES.get(int(m), m) for m in sub["mes_causacion"]],
            y=sub["valor_neto"].abs(),
            name=nombre, marker_color=color,
        )
    fig.update_layout(barmode="group", title="Ingresos vs costos por mes")
    st.plotly_chart(base_layout(fig), use_container_width=True)

with g2:
    caja = f_caja.groupby("mes_caja")["valor"].sum().sort_index()
    acumulado = df[df["mes_caja"].notna()].groupby("mes_caja")["valor"].sum().sort_index().cumsum()
    acumulado = acumulado[acumulado.index.isin(sel_meses)]
    fig = go.Figure()
    fig.add_scatter(
        x=[MESES.get(int(m), m) for m in acumulado.index], y=acumulado.values,
        mode="lines+markers", name="Saldo acumulado",
        line=dict(color=AZUL, width=2.5), fill="tozeroy",
        fillcolor="rgba(65,96,127,0.08)",
    )
    fig.update_layout(title="Evolución del saldo en caja")
    st.plotly_chart(base_layout(fig), use_container_width=True)

# ─────────────────────── Rentabilidad por cliente ───────────────────────

st.markdown('<div class="seccion">Rentabilidad</div>', unsafe_allow_html=True)
r1, r2 = st.columns([1.15, 1])

rent = (
    f_cli.assign(
        tipo=np.select(
            [f_cli["cuenta"].eq("Ingresos"), f_cli["cuenta"].eq("Costo Ventas")],
            ["ingreso", "costo"], default="otro",
        )
    )
    .query("tipo != 'otro'")
    .pivot_table(index=["cliente", "proyecto"], columns="tipo",
                 values="valor_neto", aggfunc="sum", fill_value=0)
    .reset_index()
)
for col in ("ingreso", "costo"):
    if col not in rent.columns:
        rent[col] = 0.0
rent["neto"] = rent["ingreso"] + rent["costo"]
rent["margen"] = np.where(rent["ingreso"] > 0, rent["neto"] / rent["ingreso"] * 100, np.nan)
rent = rent.sort_values("ingreso", ascending=False)

with r1:
    por_cliente = rent.groupby("cliente")[["ingreso", "neto"]].sum().sort_values("ingreso")
    fig = go.Figure()
    fig.add_bar(y=por_cliente.index, x=por_cliente["ingreso"], name="Ingresos",
                orientation="h", marker_color=AZUL)
    fig.add_bar(y=por_cliente.index, x=por_cliente["neto"], name="Neto",
                orientation="h", marker_color=VERDE)
    fig.update_layout(barmode="group", title="Ingresos y neto por cliente")
    st.plotly_chart(base_layout(fig, 330), use_container_width=True)

with r2:
    conc = rent.groupby("cliente")["ingreso"].sum().sort_values(ascending=False)
    conc = conc[conc > 0]
    fig = go.Figure(
        go.Pie(
            labels=conc.index, values=conc.values, hole=0.58,
            marker=dict(colors=[AZUL, VERDE, ARENA, ROJO, GRIS, "#8fa8bd"]),
            textinfo="percent", sort=False,
        )
    )
    fig.update_layout(title="Concentración de ingresos")
    st.plotly_chart(base_layout(fig, 330), use_container_width=True)

# Detalle por proyecto
tabla = rent[["cliente", "proyecto", "ingreso", "costo", "neto", "margen"]].copy()
tabla.columns = ["Cliente", "Proyecto", "Ingresos", "Costos", "Neto", "Margen %"]
st.dataframe(
    tabla.style.format({
        "Ingresos": "${:,.0f}", "Costos": "${:,.0f}",
        "Neto": "${:,.0f}", "Margen %": "{:.1f}%",
    }),
    use_container_width=True, hide_index=True,
)

# ─────────────────────── Composición del gasto ───────────────────────

st.markdown('<div class="seccion">Composición del gasto</div>', unsafe_allow_html=True)
gasto = (
    f_pyg[f_pyg["valor_neto"] < 0]
    .groupby("cuenta")["valor_neto"].sum().abs().sort_values(ascending=True)
)
fig = go.Figure(go.Bar(y=gasto.index, x=gasto.values, orientation="h", marker_color=GRIS))
fig.update_layout(title="Salidas por cuenta")
st.plotly_chart(base_layout(fig, 260), use_container_width=True)

# Alerta de movimientos sin clasificar
sin_clasificar = df.loc[df["cuenta"].eq("Revisar"), "valor"].sum()
if abs(sin_clasificar) > 0:
    st.warning(
        f"Hay movimientos sin clasificar por {money(sin_clasificar)}. "
        "No entran al resultado ni al flujo hasta asignarles categoría."
    )

st.caption("Fuente: modelo financiero de Manglar. Uso interno.")
