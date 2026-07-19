"""
Dashboard Manglar — 2026
Tablero financiero para los socios.

Ejecutar en local:
    streamlit run app.py
"""

import io
import os

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
HOJA_CXP = "CxP"
HOJA_CXC = "CxC"

MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre",
    12: "Diciembre",
}

TINTA = "#111827"
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

      /* Fondo blanco siempre, sin importar el tema del navegador */
      [data-testid="stAppViewContainer"], [data-testid="stHeader"],
      .main, body {background-color: #ffffff !important;}
      [data-testid="stAppViewContainer"] p,
      [data-testid="stAppViewContainer"] span,
      [data-testid="stAppViewContainer"] div {color: #111827;}

      .block-container {padding-top: 2rem; padding-bottom: 3rem; max-width: 1320px;}

      .titulo {font-size: 1.85rem; font-weight: 680; color: #111827;
               letter-spacing: -0.02em; margin-bottom: .15rem;}
      .subtitulo {font-size: .92rem; color: #6b7280 !important; margin-bottom: 1.5rem;}

      .tarjeta {background: #ffffff; border: 1px solid #e5e7eb; border-radius: 10px;
                padding: 1.05rem 1.2rem;}
      .tarjeta-titulo {font-size: .74rem; font-weight: 600; color: #6b7280 !important;
                       text-transform: uppercase; letter-spacing: .07em;
                       margin-bottom: .5rem;}
      .kpi-valor {font-size: 1.6rem; font-weight: 680; color: #111827;
                  letter-spacing: -0.02em; line-height: 1.15;}
      .kpi-nota {font-size: .76rem; color: #6b7280 !important; margin-top: .2rem;}

      .seccion {font-size: 1.12rem; font-weight: 650; color: #111827 !important;
                margin: 2rem 0 .8rem 0; padding-bottom: .4rem;
                border-bottom: 1px solid #e5e7eb;}

      /* Chips de filtro en gris, no en rojo */
      span[data-baseweb="tag"] {background-color: #eef1f5 !important;
                                border-radius: 6px !important;}
      span[data-baseweb="tag"] span {color: #111827 !important;}

      div[data-testid="stDataFrame"] {border: 1px solid #e5e7eb; border-radius: 10px;}
      label {color: #6b7280 !important; font-size: .8rem !important;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ─────────────────────────── Carga ───────────────────────────

@st.cache_data(show_spinner=False)
def leer_hojas(origen):
    mov = pd.read_excel(origen, sheet_name=HOJA_MOV)
    data = pd.read_excel(origen, sheet_name=HOJA_DATA)
    try:
        cxp = pd.read_excel(origen, sheet_name=HOJA_CXP, header=4)
    except Exception:
        cxp = pd.DataFrame()
    try:
        cxc = pd.read_excel(origen, sheet_name=HOJA_CXC, header=2)
    except Exception:
        cxc = pd.DataFrame()
    return mov, data, cxp, cxc


def preparar(mov, data):
    df = mov.copy()
    df.columns = [str(c).strip() for c in df.columns]

    equivalencias = {
        "Fecha": "fecha", "Descripción banco": "descripcion", "Valor": "valor",
        "Categoría": "categoria", "Mes Caja": "mes_caja", "Cliente": "cliente",
        "Proyecto": "proyecto", "IVA?": "iva", "Mes (Causación P&G)": "mes_causacion",
    }
    df = df.rename(columns={k: v for k, v in equivalencias.items() if k in df.columns})

    faltan = [n for n in ("valor", "categoria", "mes_caja") if n not in df.columns]
    if faltan:
        st.error("Al archivo le faltan columnas: " + ", ".join(faltan))
        st.stop()

    for col in ("cliente", "proyecto", "iva", "mes_causacion", "descripcion", "fecha"):
        if col not in df.columns:
            df[col] = np.nan

    df = df[pd.to_numeric(df["valor"], errors="coerce").notna()].copy()
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")

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
    for c in ("cliente", "proyecto"):
        df[c] = df[c].fillna("Sin asignar").astype(str).str.strip()
        df.loc[df[c].isin(["nan", "NA", "", "None"]), c] = "Sin asignar"

    df["valor_neto"] = np.where(df["iva"].eq("Si"), df["valor"] / 1.19, df["valor"])
    return df


def money(x, corto=False):
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


def tarjeta(titulo, valor, nota=""):
    return (
        f'<div class="tarjeta"><div class="tarjeta-titulo">{titulo}</div>'
        f'<div class="kpi-valor">{valor}</div><div class="kpi-nota">{nota}</div></div>'
    )


def estilo(fig, alto=320, titulo="", leyenda="arriba"):
    if leyenda == "abajo":
        pos = dict(orientation="h", yanchor="top", y=-0.05, x=0.5, xanchor="center",
                   font=dict(size=11, color=TINTA), bgcolor="rgba(0,0,0,0)")
        margen = dict(l=10, r=40, t=52, b=60)
    elif leyenda == "no":
        pos = None
        margen = dict(l=10, r=40, t=52, b=10)
    else:
        pos = dict(orientation="h", yanchor="bottom", y=1.0, x=0,
                   font=dict(size=11, color=TINTA), bgcolor="rgba(0,0,0,0)")
        margen = dict(l=10, r=40, t=52, b=10)

    fig.update_layout(
        height=alto,
        title=dict(text=titulo, font=dict(size=14, color=TINTA), x=0, xanchor="left",
                   y=0.97, yanchor="top"),
        margin=margen,
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(family="ui-sans-serif, system-ui, sans-serif", size=12, color=TINTA),
        showlegend=pos is not None,
        legend=pos or {},
        hoverlabel=dict(bgcolor="white", font_size=12, bordercolor=GRIS_SUAVE),
    )
    fig.update_xaxes(showgrid=False, linecolor=GRIS_SUAVE,
                     tickfont=dict(size=11, color=GRIS))
    fig.update_yaxes(gridcolor=GRIS_SUAVE, zeroline=True, zerolinecolor=GRIS_SUAVE,
                     tickfont=dict(size=11, color=GRIS))
    return fig


# ─────────────────────────── Encabezado ───────────────────────────

st.markdown('<div class="titulo">Dashboard Manglar — 2026</div>', unsafe_allow_html=True)

origen = ARCHIVO_MODELO if os.path.exists(ARCHIVO_MODELO) else None
if origen is None:
    subido = st.file_uploader("Cargue el modelo financiero (.xlsx)", type=["xlsx"])
    if subido is None:
        st.stop()
    origen = io.BytesIO(subido.getvalue())

mov_raw, data_raw, cxp_raw, cxc_raw = leer_hojas(origen)
df = preparar(mov_raw, data_raw)

# ─────────────────────────── Filtro de mes ───────────────────────────

meses_disp = sorted(int(m) for m in df["mes_caja"].dropna().unique())
clientes_disp = sorted(c for c in df["cliente"].unique() if c != "Sin asignar")

sel_meses = st.multiselect("Mes", meses_disp, default=meses_disp,
                           format_func=lambda m: MESES.get(m, str(m)))
sel_meses = sel_meses or meses_disp

f_pyg = df[df["mes_causacion"].isin(sel_meses)]

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
saldo_caja = df["valor"].sum()

k = st.columns(4)
k[0].markdown(tarjeta("Saldo en caja", money(saldo_caja, True), "Acumulado a la fecha"),
              unsafe_allow_html=True)
k[1].markdown(tarjeta("Ingresos", money(ingresos, True), "Netos de IVA"),
              unsafe_allow_html=True)
k[2].markdown(tarjeta("Margen bruto", money(margen_bruto, True),
                      f"{margen_pct:.1f}% sobre ingresos"), unsafe_allow_html=True)
k[3].markdown(tarjeta("Resultado operativo", money(resultado, True), "Después de gastos"),
              unsafe_allow_html=True)

# ─────────────────────── Resultado y caja ───────────────────────

st.markdown('<div class="seccion">Resultado y caja</div>', unsafe_allow_html=True)
g1, g2 = st.columns(2)

with g1:
    tmp = f_pyg.assign(
        grupo=np.select(
            [f_pyg["cuenta"].eq("Ingresos"), f_pyg["cuenta"].eq("Costo Ventas")],
            ["Ingresos", "Costo de ventas"], default="Gastos",
        )
    )
    pyg = tmp.groupby(["mes_causacion", "grupo"])["valor_neto"].sum().reset_index()
    orden = sorted(pyg["mes_causacion"].dropna().unique())
    etiquetas = [MESES.get(int(m), str(m)) for m in orden]

    fig = go.Figure()
    for nombre, color in [("Ingresos", VERDE), ("Costo de ventas", ROJO), ("Gastos", ARENA)]:
        sub = (pyg[pyg["grupo"].eq(nombre)]
               .set_index("mes_causacion").reindex(orden)["valor_neto"].abs().fillna(0))
        fig.add_bar(x=etiquetas, y=sub.values, name=nombre, marker_color=color,
                    hovertemplate="%{x} · %{y:$,.0f}<extra>" + nombre + "</extra>")
    fig.update_layout(barmode="group")
    fig.update_yaxes(tickformat="$~s")
    st.plotly_chart(estilo(fig, 340, "Ingresos, costos y gastos por mes"),
                    use_container_width=True)

with g2:
    serie = (df[df["mes_caja"].notna()].groupby("mes_caja")["valor"]
             .sum().sort_index().cumsum())
    serie = serie[serie.index.isin(sel_meses)]
    fig = go.Figure()
    fig.add_scatter(
        x=[MESES.get(int(m), str(m)) for m in serie.index], y=serie.values,
        mode="lines+markers+text", name="Saldo",
        text=[money(v, True) for v in serie.values], textposition="top center",
        textfont=dict(size=10, color=GRIS),
        line=dict(color=AZUL, width=2.5), marker=dict(size=7),
        fill="tozeroy", fillcolor="rgba(65,96,127,0.07)",
        hovertemplate="%{x} · %{y:$,.0f}<extra></extra>",
    )
    fig.update_yaxes(tickformat="$~s")
    st.plotly_chart(estilo(fig, 340, "Evolución del saldo en caja"),
                    use_container_width=True)

# ─────────────────────── Rentabilidad ───────────────────────

st.markdown('<div class="seccion">Rentabilidad por cliente y proyecto</div>',
            unsafe_allow_html=True)

sel_clientes = st.multiselect("Cliente", clientes_disp, default=clientes_disp)
sel_clientes = sel_clientes or clientes_disp
f_cli = f_pyg[f_pyg["cliente"].isin(sel_clientes)]

rent = (
    f_cli.assign(
        tipo=np.select([f_cli["cuenta"].eq("Ingresos"), f_cli["cuenta"].eq("Costo Ventas")],
                       ["ingreso", "costo"], default="otro")
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

r1, r2 = st.columns([1.25, 1])
with r1:
    pc = rent.groupby("cliente")[["ingreso", "neto"]].sum()
    pc = pc[pc["ingreso"] > 0].sort_values("ingreso")
    fig = go.Figure()
    fig.add_bar(y=pc.index, x=pc["ingreso"], name="Ingresos", orientation="h",
                marker_color=AZUL, text=[money(v, True) for v in pc["ingreso"]],
                textposition="outside", textfont=dict(size=10),
                hovertemplate="%{y} · %{x:$,.0f}<extra></extra>")
    fig.add_bar(y=pc.index, x=pc["neto"], name="Neto", orientation="h",
                marker_color=VERDE, text=[money(v, True) for v in pc["neto"]],
                textposition="outside", textfont=dict(size=10),
                hovertemplate="%{y} · %{x:$,.0f}<extra></extra>")
    fig.update_layout(barmode="group")
    fig.update_xaxes(tickformat="$~s")
    st.plotly_chart(estilo(fig, 360, "Ingresos y neto por cliente"),
                    use_container_width=True)

with r2:
    conc = rent.groupby("cliente")["ingreso"].sum().sort_values(ascending=False)
    conc = conc[conc > 0]
    fig = go.Figure(go.Pie(
        labels=conc.index, values=conc.values, hole=0.6, sort=False,
        marker=dict(colors=[AZUL, VERDE, ARENA, ROJO, "#8fa8bd", GRIS]),
        textinfo="percent", textfont=dict(size=11, color="white"),
    ))
    st.plotly_chart(estilo(fig, 380, "Concentración de ingresos", "abajo"),
                    use_container_width=True)

tabla = rent[["cliente", "proyecto", "ingreso", "costo", "neto", "margen"]].copy()
tabla.columns = ["Cliente", "Proyecto", "Ingresos", "Costos", "Neto", "Margen %"]
tabla = tabla.sort_values("Ingresos", ascending=False)
st.dataframe(
    tabla.style.format({"Ingresos": "${:,.0f}", "Costos": "${:,.0f}",
                        "Neto": "${:,.0f}", "Margen %": "{:.1f}%"}),
    use_container_width=True, hide_index=True,
)

# ─────────────────────── Cuentas por cobrar ───────────────────────

if not cxc_raw.empty and "Cliente" in cxc_raw.columns:
    st.markdown('<div class="seccion">Cuentas por cobrar</div>', unsafe_allow_html=True)
    col_cli = "Cliente"
    cxc = cxc_raw.dropna(subset=[col_cli]).copy()
    cxc = cxc[~cxc[col_cli].astype(str).str.upper().str.startswith("TOTAL")]

    col_fact = "Valor Facturado" if "Valor Facturado" in cxc.columns else cxc.columns[2]
    col_cob = "Cobrado" if "Cobrado" in cxc.columns else cxc.columns[3]
    col_pend = "Saldo Pendiente" if "Saldo Pendiente" in cxc.columns else cxc.columns[4]
    for c in (col_fact, col_cob, col_pend):
        cxc[c] = pd.to_numeric(cxc[c], errors="coerce").fillna(0)

    kc = st.columns(3)
    kc[0].markdown(tarjeta("Facturado", money(cxc[col_fact].sum(), True), "Con IVA"),
                   unsafe_allow_html=True)
    kc[1].markdown(tarjeta("Cobrado", money(cxc[col_cob].sum(), True), "Recibido en caja"),
                   unsafe_allow_html=True)
    kc[2].markdown(tarjeta("Cartera pendiente", money(cxc[col_pend].sum(), True),
                           "Por recaudar"), unsafe_allow_html=True)

    cc1, cc2 = st.columns([1.25, 1])
    with cc1:
        pend = cxc[cxc[col_pend] > 0].sort_values(col_pend)
        if not pend.empty:
            fig = go.Figure(go.Bar(
                y=pend[col_cli].astype(str), x=pend[col_pend], orientation="h",
                marker_color=ARENA, text=[money(v, True) for v in pend[col_pend]],
                textposition="outside", textfont=dict(size=10),
                hovertemplate="%{y} · %{x:$,.0f}<extra></extra>",
            ))
            fig.update_xaxes(tickformat="$~s")
            st.plotly_chart(estilo(fig, 300, "Cartera pendiente por cliente"),
                            use_container_width=True)
    with cc2:
        if "Antigüedad" in cxc.columns:
            ant = cxc[cxc[col_pend] > 0].groupby("Antigüedad")[col_pend].sum()
            if not ant.empty:
                fig = go.Figure(go.Bar(x=ant.index.astype(str), y=ant.values,
                                       marker_color=ROJO,
                                       text=[money(v, True) for v in ant.values],
                                       textposition="outside", textfont=dict(size=10),
                hovertemplate="%{y} · %{x:$,.0f}<extra></extra>"))
                fig.update_yaxes(tickformat="$~s")
                st.plotly_chart(estilo(fig, 300, "Cartera por antigüedad (días)"),
                                use_container_width=True)

    cols_cxc = [c for c in [col_cli, "Proyecto", col_fact, col_cob, col_pend,
                            "Fecha Vence", "Antigüedad", "Estado", "Notas"]
                if c in cxc.columns]
    vista = cxc[cols_cxc].copy()
    if "Fecha Vence" in vista.columns:
        vista["Fecha Vence"] = pd.to_datetime(vista["Fecha Vence"],
                                              errors="coerce").dt.strftime("%d/%m/%Y")
    st.dataframe(
        vista.style.format({col_fact: "${:,.0f}", col_cob: "${:,.0f}",
                            col_pend: "${:,.0f}"}),
        use_container_width=True, hide_index=True,
    )

# ─────────────────────── Cuentas por pagar ───────────────────────

if not cxp_raw.empty and "Proveedor" in cxp_raw.columns:
    st.markdown('<div class="seccion">Cuentas por pagar</div>', unsafe_allow_html=True)
    cxp = cxp_raw.dropna(subset=["Proveedor"]).copy()

    col_tot = "Valor Total" if "Valor Total" in cxp.columns else cxp.columns[10]
    cxp[col_tot] = pd.to_numeric(cxp[col_tot], errors="coerce").fillna(0)
    estado = (cxp["Estado"].astype(str).str.strip().str.lower()
              if "Estado" in cxp.columns else pd.Series("", index=cxp.index))
    es_pagado = estado.eq("pagado")
    pagado = cxp.loc[es_pagado, col_tot].sum()
    pendiente = cxp.loc[~es_pagado, col_tot].sum()

    kp = st.columns(3)
    kp[0].markdown(tarjeta("Total facturado", money(cxp[col_tot].sum(), True),
                           "Proveedores"), unsafe_allow_html=True)
    kp[1].markdown(tarjeta("Pagado", money(pagado, True), "Ya desembolsado"),
                   unsafe_allow_html=True)
    kp[2].markdown(tarjeta("Pendiente de pago", money(pendiente, True),
                           "Compromiso de caja"), unsafe_allow_html=True)

    pp1, pp2 = st.columns([1.25, 1])
    with pp1:
        pend_p = cxp[~es_pagado]
        if not pend_p.empty:
            por_prov = pend_p.groupby("Proveedor")[col_tot].sum().sort_values()
            fig = go.Figure(go.Bar(
                y=por_prov.index.astype(str), x=por_prov.values, orientation="h",
                marker_color=ROJO, text=[money(v, True) for v in por_prov.values],
                textposition="outside", textfont=dict(size=10),
                hovertemplate="%{y} · %{x:$,.0f}<extra></extra>",
            ))
            fig.update_xaxes(tickformat="$~s")
            st.plotly_chart(estilo(fig, 300, "Pendiente por proveedor"),
                            use_container_width=True)
    with pp2:
        resumen = pd.Series({"Pagado": pagado, "Pendiente": pendiente})
        fig = go.Figure(go.Pie(labels=resumen.index, values=resumen.values, hole=0.6,
                               marker=dict(colors=[VERDE, ROJO]), sort=False,
                               textinfo="percent", textfont=dict(size=11, color="white")))
        st.plotly_chart(estilo(fig, 320, "Estado de pagos", "abajo"),
                        use_container_width=True)

    cols_cxp = [c for c in ["Fecha Recepción", "Fecha Pago", "Estado", "Proveedor",
                            "Cliente", "Proyecto", "Concepto/Rubro", col_tot]
                if c in cxp.columns]
    vista = cxp[cols_cxp].copy()
    for c in ("Fecha Recepción", "Fecha Pago"):
        if c in vista.columns:
            vista[c] = pd.to_datetime(vista[c], errors="coerce").dt.strftime("%d/%m/%Y")
    st.dataframe(vista.style.format({col_tot: "${:,.0f}"}),
                 use_container_width=True, hide_index=True)

# ─────────────────────── Composición del gasto ───────────────────────

st.markdown('<div class="seccion">Composición del gasto</div>', unsafe_allow_html=True)
gasto = (f_pyg[f_pyg["valor_neto"] < 0].groupby("cuenta")["valor_neto"]
         .sum().abs().sort_values())
fig = go.Figure(go.Bar(y=gasto.index, x=gasto.values, orientation="h", marker_color=GRIS,
                       text=[money(v, True) for v in gasto.values],
                       textposition="outside", textfont=dict(size=10),
                hovertemplate="%{y} · %{x:$,.0f}<extra></extra>"))
fig.update_xaxes(tickformat="$~s")
st.plotly_chart(estilo(fig, 280, "Salidas por cuenta"), use_container_width=True)

sin_clasificar = df.loc[df["cuenta"].eq("Revisar"), "valor"].sum()
if abs(sin_clasificar) > 0:
    st.warning(
        f"Hay movimientos sin clasificar por {money(sin_clasificar)}. "
        "No entran al resultado hasta asignarles categoría."
    )

