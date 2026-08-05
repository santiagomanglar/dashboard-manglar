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
HOJA_PROY = "Proyeccion Flujo de Caja"
HOJA_IVA = "Planeacion IVA"

# Fila donde arranca el encabezado de cada tabla (base 0, como lo espera pandas)
CABECERA_CXC = 9    # fila 10 de la hoja
CABECERA_CXP = 4    # fila 5 de la hoja

MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre",
    12: "Diciembre",
}

# Valores que significan "todavia no se asigno"
SIN_ASIGNAR = {"", "nan", "none", "na", "revisar", "por asignar", "sin asignar"}

TINTA = "#111827"
GRIS = "#6b7280"
GRIS_SUAVE = "#e5e7eb"
VERDE = "#3f6f5b"
ROJO = "#a4553f"
AZUL = "#41607f"
ARENA = "#c9a227"

PLOTLY_CONF = {"displayModeBar": False, "responsive": True}

# ─────────────────────────── Estilos ───────────────────────────

st.markdown(
    """
    <style>
      #MainMenu, footer, header {visibility: hidden;}

      [data-testid="stAppViewContainer"], [data-testid="stHeader"],
      .main, body {background-color: #ffffff !important;}
      [data-testid="stAppViewContainer"] p,
      [data-testid="stAppViewContainer"] span,
      [data-testid="stAppViewContainer"] div {color: #111827;}

      .block-container {padding-top: 2rem; padding-bottom: 3rem; max-width: 1320px;}

      .titulo {font-size: 1.85rem; font-weight: 680; color: #111827;
               letter-spacing: -0.02em; margin-bottom: 1.4rem;}

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

      span[data-baseweb="tag"] {background-color: #eef1f5 !important;
                                border-radius: 6px !important;}
      span[data-baseweb="tag"] span {color: #111827 !important;}

      div[data-testid="stDataFrame"] {border: 1px solid #e5e7eb; border-radius: 10px;}
      label {color: #6b7280 !important; font-size: .8rem !important;}

      /* Botones de filtro: sobrios, sin llenar la pantalla */
      div[data-testid="stButtonGroup"] button {
        border: 1px solid #e5e7eb !important; background: #ffffff !important;
        border-radius: 8px !important; font-size: .82rem !important;
        color: #374151 !important; padding: .25rem .7rem !important;}
      div[data-testid="stButtonGroup"] button[aria-checked="true"],
      div[data-testid="stButtonGroup"] button[aria-pressed="true"] {
        background: #41607f !important; border-color: #41607f !important;}
      div[data-testid="stButtonGroup"] button[aria-checked="true"] *,
      div[data-testid="stButtonGroup"] button[aria-pressed="true"] * {
        color: #ffffff !important;}

      div[data-testid="stPopover"] button {
        border: 1px solid #e5e7eb !important; background: #ffffff !important;
        border-radius: 8px !important; font-size: .82rem !important;
        color: #374151 !important; justify-content: space-between !important;}

      @media (max-width: 640px) {
        .block-container {padding-left: .75rem !important; padding-right: .75rem !important;
                          padding-top: 1.2rem !important;}
        .titulo {font-size: 1.32rem;}
        .seccion {font-size: 1rem; margin: 1.5rem 0 .6rem 0;}
        .tarjeta {padding: .8rem .9rem;}
        .tarjeta-titulo {font-size: .66rem; letter-spacing: .05em;}
        .kpi-valor {font-size: 1.28rem;}
        .kpi-nota {font-size: .7rem;}

        [data-testid="stHorizontalBlock"] {flex-wrap: wrap !important; gap: .55rem !important;}
        [data-testid="stHorizontalBlock"] > div[data-testid="column"] {
          flex: 1 1 calc(50% - .55rem) !important;
          min-width: calc(50% - .55rem) !important;
        }
        [data-testid="stHorizontalBlock"] > div[data-testid="column"]:has(.js-plotly-plot),
        [data-testid="stHorizontalBlock"] > div[data-testid="column"]:has([data-testid="stDataFrame"]) {
          flex: 1 1 100% !important; min-width: 100% !important;
        }
      }
    </style>
    """,
    unsafe_allow_html=True,
)


# ─────────────────────────── Utilidades ───────────────────────────

def cortar_en_total(df: pd.DataFrame, columna: str) -> pd.DataFrame:
    """Devuelve solo las filas de datos: corta en la fila TOTAL y descarta vacias.

    Asi la tabla puede crecer sin tener que tocar el codigo: mientras la fila
    TOTAL siga cerrando el bloque, se leen todas las filas que haya encima.
    """
    if columna not in df.columns:
        return pd.DataFrame()
    marca = df[columna].astype(str).str.strip().str.upper().eq("TOTAL")
    if marca.any():
        corte = marca.idxmax()
        df = df.loc[: corte - 1] if corte > 0 else df.iloc[0:0]
    return df.dropna(subset=[columna]).copy()


def numerico(df: pd.DataFrame, columnas) -> pd.DataFrame:
    for c in columnas:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        else:
            df[c] = 0.0
    return df


def texto(serie: pd.Series) -> pd.Series:
    return serie.fillna("Sin asignar").astype(str).str.strip()


def esta_asignado(serie: pd.Series) -> pd.Series:
    return ~texto(serie).str.lower().isin(SIN_ASIGNAR)


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


def barras_horizontal(etiquetas, valores, color, titulo, alto=300):
    fig = go.Figure(go.Bar(
        y=etiquetas, x=valores, orientation="h", marker_color=color,
        text=[money(v, True) for v in valores], textposition="auto",
        textfont=dict(size=10), cliponaxis=False,
        hovertemplate="%{y} · %{x:$,.0f}<extra></extra>",
    ))
    fig.update_xaxes(tickformat="$~s")
    return estilo(fig, alto, titulo, "no")


AYUDA_FILTRO = "Si no selecciona ninguno, se muestran todos."


def filtro(etiqueta, opciones, key, formato=None, umbral=8):
    """Filtro compacto. Sin seleccion significa todos, para no llenar la
    pantalla de etiquetas.

    Listas cortas se muestran como botones. Listas largas se colapsan en un
    boton que abre el buscador.
    """
    opciones = list(opciones)
    formato = formato or (lambda x: str(x))

    if not opciones:
        return []

    if len(opciones) <= umbral and hasattr(st, "pills"):
        sel = st.pills(etiqueta, opciones, selection_mode="multi", key=key,
                       format_func=formato, help=AYUDA_FILTRO)
        return list(sel) if sel else opciones

    actuales = st.session_state.get(key) or []
    resumen = (f"{etiqueta}: {len(actuales)} de {len(opciones)}" if actuales
               else f"{etiqueta}: todos ({len(opciones)})")

    if hasattr(st, "popover"):
        with st.popover(resumen, use_container_width=True):
            sel = st.multiselect(etiqueta, opciones, key=key, format_func=formato,
                                 placeholder="Escriba para buscar",
                                 label_visibility="collapsed")
    else:
        sel = st.multiselect(etiqueta, opciones, key=key, format_func=formato,
                             placeholder=f"Todos ({len(opciones)})",
                             help=AYUDA_FILTRO)
    return list(sel) if sel else opciones


# ─────────────────────────── Carga ───────────────────────────

def leer_tabla(origen, hoja, claves, fila_por_defecto):
    """Lee una tabla buscando sola la fila de encabezado.

    `claves` son los nombres posibles de la columna que identifica cada fila.
    Se aceptan varios para que un cambio de nombre en el modelo no rompa nada.
    """
    if isinstance(claves, str):
        claves = [claves]
    try:
        crudo = pd.read_excel(origen, sheet_name=hoja, header=None, nrows=30)
        fila = fila_por_defecto
        for i in range(len(crudo)):
            valores = [str(v).strip() for v in crudo.iloc[i].tolist()]
            if any(k in valores for k in claves):
                fila = i
                break
        if hasattr(origen, "seek"):
            origen.seek(0)
        return pd.read_excel(origen, sheet_name=hoja, header=fila)
    except Exception:
        return pd.DataFrame()


def columna_presente(df, opciones):
    """Devuelve el primer nombre de columna que exista en el marco."""
    for o in opciones:
        if o in df.columns:
            return o
    return None


@st.cache_data(show_spinner=False)
def leer_hojas(origen):
    mov = pd.read_excel(origen, sheet_name=HOJA_MOV)
    if hasattr(origen, "seek"):
        origen.seek(0)
    data = pd.read_excel(origen, sheet_name=HOJA_DATA)
    if hasattr(origen, "seek"):
        origen.seek(0)
    cxp = leer_tabla(origen, HOJA_CXP, ["Factura", "Numero"], CABECERA_CXP)
    cxc = leer_tabla(origen, HOJA_CXC, ["Folio"], CABECERA_CXC)
    if hasattr(origen, "seek"):
        origen.seek(0)
    proy = leer_cruda(origen, HOJA_PROY)
    if hasattr(origen, "seek"):
        origen.seek(0)
    iva = leer_cruda(origen, HOJA_IVA)
    return mov, data, cxp, cxc, proy, iva


def leer_cruda(origen, hoja):
    """Lee una hoja sin encabezado, para ubicar las filas por su etiqueta."""
    try:
        return pd.read_excel(origen, sheet_name=hoja, header=None)
    except Exception:
        return pd.DataFrame()


def fila_por_etiqueta(df, etiqueta, columna=1):
    """Devuelve el indice de la fila cuya etiqueta coincide.

    Se busca por texto y no por posicion, para que el tablero siga funcionando
    aunque se agreguen o quiten filas en el modelo.
    """
    if df.empty or columna >= df.shape[1]:
        return None
    col = df[columna].astype(str).str.strip().str.lower()
    coincide = col.eq(etiqueta.strip().lower())
    return int(coincide.idxmax()) if coincide.any() else None


def valores_fila(df, fila, desde=2, hasta=15):
    if fila is None:
        return []
    hasta = min(hasta, df.shape[1])
    vals = []
    for c in range(desde, hasta):
        v = df.iat[fila, c]
        vals.append(v if isinstance(v, (int, float)) and not pd.isna(v) else 0.0)
    return vals


def preparar_movimientos(mov, data):
    df = mov.copy()
    df.columns = [str(c).strip() for c in df.columns]

    equivalencias = {
        "Fecha": "fecha", "Descripción banco": "descripcion", "Valor": "valor",
        "Categoría": "categoria", "Mes Caja": "mes_caja", "Cliente": "cliente",
        "Proyecto": "proyecto", "IVA?": "iva", "Mes (Causación P&G)": "mes_causacion",
        "Factura": "factura",
    }
    df = df.rename(columns={k: v for k, v in equivalencias.items() if k in df.columns})

    faltan = [n for n in ("valor", "categoria", "mes_caja") if n not in df.columns]
    if faltan:
        st.error("Al archivo le faltan columnas: " + ", ".join(faltan))
        st.stop()

    for col in ("cliente", "proyecto", "iva", "mes_causacion", "descripcion",
                "fecha", "factura"):
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
        df[c] = texto(df[c])
        df.loc[df[c].str.lower().isin(SIN_ASIGNAR), c] = "Sin asignar"

    df["valor_neto"] = np.where(df["iva"].eq("Si"), df["valor"] / 1.19, df["valor"])
    return df


def preparar_cxc(bruto):
    if bruto.empty:
        return pd.DataFrame()
    df = cortar_en_total(bruto, "Folio")
    if df.empty:
        return df
    df = numerico(df, ["Base sin IVA", "IVA 19%", "Total Facturado", "Retenciones",
                       "Neto a Cobrar", "Cobrado", "Saldo Pendiente"])
    df["Mes"] = pd.to_numeric(df.get("Mes"), errors="coerce")
    for c in ("Cliente", "Proyecto", "Estado", "Antiguedad"):
        if c in df.columns:
            df[c] = texto(df[c])
    return df


def preparar_cxp(bruto):
    if bruto.empty:
        return pd.DataFrame(), 0.0, 0
    clave = columna_presente(bruto, ["Factura", "Numero"])
    if clave is None:
        return pd.DataFrame(), 0.0, 0
    df = cortar_en_total(bruto, clave)
    if df.empty:
        return df, 0.0, 0
    df = numerico(df, ["Valor Producto", "IVA", "Valor Total", "Pagado", "Saldo"])
    for c in ("Proveedor", "Cliente", "Proyecto", "Estado", "Alerta", "Dashboard"):
        if c in df.columns:
            df[c] = texto(df[c])
    df["asignado"] = esta_asignado(df["Cliente"]) & esta_asignado(df["Proyecto"])

    # La columna Dashboard del modelo decide que facturas se muestran.
    # Las marcadas No siguen en el CxP, solo no aparecen aqui.
    excluido, n_excluidas = 0.0, 0
    if "Dashboard" in df.columns:
        fuera = df["Dashboard"].str.strip().str.lower().eq("no")
        excluido = float(df.loc[fuera, "Valor Total"].sum())
        n_excluidas = int(fuera.sum())
        df = df[~fuera].copy()
    return df, excluido, n_excluidas


# ─────────────────────────── Encabezado ───────────────────────────

st.markdown('<div class="titulo">Dashboard Manglar — 2026</div>', unsafe_allow_html=True)

origen = ARCHIVO_MODELO if os.path.exists(ARCHIVO_MODELO) else None
if origen is None:
    subido = st.file_uploader("Cargue el modelo financiero (.xlsx)", type=["xlsx"])
    if subido is None:
        st.stop()
    origen = io.BytesIO(subido.getvalue())

mov_raw, data_raw, cxp_raw, cxc_raw, proy_raw, iva_raw = leer_hojas(origen)
df = preparar_movimientos(mov_raw, data_raw)
cxc = preparar_cxc(cxc_raw)
cxp, cxp_excluido, cxp_n_excl = preparar_cxp(cxp_raw)

# ─────────────────────────── Filtro de mes ───────────────────────────

meses_mov = {int(m) for m in df["mes_caja"].dropna().unique()}
meses_fac = {int(m) for m in cxc["Mes"].dropna().unique()} if not cxc.empty else set()
meses_disp = sorted(meses_mov | meses_fac)

sel_meses = filtro("Mes", meses_disp, "f_mes",
                   formato=lambda m: MESES.get(m, str(m)), umbral=12)

f_pyg = df[df["mes_causacion"].isin(sel_meses)]
cxc_mes = cxc[cxc["Mes"].isin(sel_meses)] if not cxc.empty else cxc

# ─────────────────────────── Indicadores ───────────────────────────

# El ingreso del P&G sale de las facturas emitidas, igual que en el modelo
ingresos = cxc_mes["Base sin IVA"].sum() if not cxc_mes.empty else 0.0
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
k[1].markdown(tarjeta("Ingresos", money(ingresos, True), "Facturado, neto de IVA"),
              unsafe_allow_html=True)
k[2].markdown(tarjeta("Margen bruto", money(margen_bruto, True),
                      f"{margen_pct:.1f}% sobre ingresos"), unsafe_allow_html=True)
k[3].markdown(tarjeta("Resultado operativo", money(resultado, True), "Después de gastos"),
              unsafe_allow_html=True)

# ─────────────────────── Resultado y caja ───────────────────────

st.markdown('<div class="seccion">Resultado y caja</div>', unsafe_allow_html=True)
g1, g2 = st.columns(2)

with g1:
    orden = sorted(sel_meses)
    etiquetas = [MESES.get(int(m), str(m)) for m in orden]

    serie_ing = (cxc_mes.groupby("Mes")["Base sin IVA"].sum().reindex(orden).fillna(0)
                 if not cxc_mes.empty else pd.Series(0.0, index=orden))
    tmp = f_pyg.assign(
        grupo=np.where(f_pyg["cuenta"].eq("Costo Ventas"), "Costo de ventas", "Gastos")
    )
    tmp = tmp[tmp["cuenta"].isin(
        ["Costo Ventas", "Costos Fijos", "Costos Operacionales", "Gastos Variables"])]
    pyg = tmp.groupby(["mes_causacion", "grupo"])["valor_neto"].sum().reset_index()

    fig = go.Figure()
    fig.add_bar(x=etiquetas, y=serie_ing.values, name="Ingresos facturados",
                marker_color=VERDE,
                hovertemplate="%{x} · %{y:$,.0f}<extra>Ingresos</extra>")
    for nombre, color in [("Costo de ventas", ROJO), ("Gastos", ARENA)]:
        sub = (pyg[pyg["grupo"].eq(nombre)]
               .set_index("mes_causacion").reindex(orden)["valor_neto"].abs().fillna(0))
        fig.add_bar(x=etiquetas, y=sub.values, name=nombre, marker_color=color,
                    hovertemplate="%{x} · %{y:$,.0f}<extra>" + nombre + "</extra>")
    fig.update_layout(barmode="group")
    fig.update_yaxes(tickformat="$~s")
    st.plotly_chart(estilo(fig, 340, "Ingresos facturados y costos por mes"),
                    use_container_width=True, config=PLOTLY_CONF)

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
                    use_container_width=True, config=PLOTLY_CONF)

# ─────────────────────── Rentabilidad ───────────────────────

st.markdown('<div class="seccion">Rentabilidad por cliente y proyecto</div>',
            unsafe_allow_html=True)

fr1, fr2 = st.columns([1, 2])
with fr1:
    opciones_base = ["Facturas", "Movimientos"]
    if hasattr(st, "segmented_control"):
        enfoque = st.segmented_control("Base de cálculo", opciones_base,
                                       default="Facturas", key="enfoque")
        enfoque = enfoque or "Facturas"
    else:
        enfoque = st.radio("Base de cálculo", opciones_base, horizontal=True,
                           key="enfoque")


def rentabilidad_facturas():
    if cxc.empty:
        return pd.DataFrame(columns=["cliente", "proyecto", "ingresos", "costos"])
    ing = (cxc.groupby(["Cliente", "Proyecto"])["Base sin IVA"].sum()
           .rename("ingresos"))
    if not cxp.empty:
        cos = (cxp[cxp["asignado"]].groupby(["Cliente", "Proyecto"])["Valor Producto"]
               .sum().rename("costos"))
    else:
        cos = pd.Series(dtype=float, name="costos")
    r = pd.concat([ing, cos], axis=1).fillna(0).reset_index()
    r.columns = ["cliente", "proyecto", "ingresos", "costos"]
    r["costos"] = -r["costos"]
    return r


def rentabilidad_movimientos():
    base = df[df["mes_causacion"].notna()]
    r = (base.assign(
            tipo=np.select([base["cuenta"].eq("Ingresos"),
                            base["cuenta"].eq("Costo Ventas")],
                           ["ingresos", "costos"], default="otro"))
         .query("tipo != 'otro'")
         .pivot_table(index=["cliente", "proyecto"], columns="tipo",
                      values="valor_neto", aggfunc="sum", fill_value=0)
         .reset_index())
    for c in ("ingresos", "costos"):
        if c not in r.columns:
            r[c] = 0.0
    return r[["cliente", "proyecto", "ingresos", "costos"]]


rent = (rentabilidad_facturas() if enfoque == "Facturas"
        else rentabilidad_movimientos())
rent["neto"] = rent["ingresos"] + rent["costos"]
rent["margen"] = np.where(rent["ingresos"] > 0,
                          rent["neto"] / rent["ingresos"] * 100, np.nan)

clientes_disp = sorted(c for c in rent["cliente"].unique()
                       if str(c).lower() not in SIN_ASIGNAR)
with fr2:
    sel_clientes = filtro("Cliente", clientes_disp, "f_rent_cli")
rent = rent[rent["cliente"].isin(sel_clientes)]

r1, r2 = st.columns([1.25, 1])
with r1:
    pc = rent.groupby("cliente")[["ingresos", "neto"]].sum()
    pc = pc[pc["ingresos"] > 0].sort_values("ingresos")
    fig = go.Figure()
    fig.add_bar(y=pc.index, x=pc["ingresos"], name="Ingresos", orientation="h",
                marker_color=AZUL, text=[money(v, True) for v in pc["ingresos"]],
                textposition="auto", textfont=dict(size=10), cliponaxis=False,
                hovertemplate="%{y} · %{x:$,.0f}<extra>Ingresos</extra>")
    fig.add_bar(y=pc.index, x=pc["neto"], name="Neto", orientation="h",
                marker_color=VERDE, text=[money(v, True) for v in pc["neto"]],
                textposition="auto", textfont=dict(size=10), cliponaxis=False,
                hovertemplate="%{y} · %{x:$,.0f}<extra>Neto</extra>")
    fig.update_layout(barmode="group")
    fig.update_xaxes(tickformat="$~s")
    st.plotly_chart(estilo(fig, 360, "Ingresos y neto por cliente"),
                    use_container_width=True, config=PLOTLY_CONF)

with r2:
    conc = rent.groupby("cliente")["ingresos"].sum().sort_values(ascending=False)
    conc = conc[conc > 0]
    fig = go.Figure(go.Pie(
        labels=conc.index, values=conc.values, hole=0.6, sort=False,
        marker=dict(colors=[AZUL, VERDE, ARENA, ROJO, "#8fa8bd", GRIS, "#b08968"]),
        textinfo="percent", textfont=dict(size=11, color="white"),
    ))
    st.plotly_chart(estilo(fig, 380, "Concentración de ingresos", "abajo"),
                    use_container_width=True, config=PLOTLY_CONF)

tabla = rent[["cliente", "proyecto", "ingresos", "costos", "neto", "margen"]].copy()
tabla.columns = ["Cliente", "Proyecto", "Ingresos", "Costos", "Neto", "Margen %"]
tabla = tabla.sort_values("Ingresos", ascending=False)
st.dataframe(
    tabla.style.format({"Ingresos": "${:,.0f}", "Costos": "${:,.0f}",
                        "Neto": "${:,.0f}", "Margen %": "{:.1f}%"}),
    use_container_width=True, hide_index=True,
)

# Aviso honesto: si faltan costos por asignar, el margen esta sobrestimado
if enfoque == "Facturas" and not cxp.empty:
    sin_asignar = cxp.loc[~cxp["asignado"], "Valor Producto"].sum()
    total_costo = cxp["Valor Producto"].sum()
    if sin_asignar > 0:
        pct = sin_asignar / total_costo * 100 if total_costo else 0
        st.warning(
            f"Faltan {money(sin_asignar)} de costos por asignar a un proyecto "
            f"({pct:.0f}% del total). Hasta completarlos, el margen de esta vista "
            "queda por encima del real."
        )

# ─────────────────────── Cuentas por cobrar ───────────────────────

if not cxc.empty:
    st.markdown('<div class="seccion">Cuentas por cobrar</div>', unsafe_allow_html=True)

    fc1, fc2 = st.columns(2)
    with fc1:
        op_cli = sorted(cxc["Cliente"].unique())
        sel_cxc_cli = filtro("Cliente", op_cli, "f_cxc_cli")
    with fc2:
        op_est = sorted(cxc["Estado"].unique())
        sel_cxc_est = filtro("Estado", op_est, "f_cxc_est")

    v = cxc[cxc["Cliente"].isin(sel_cxc_cli) & cxc["Estado"].isin(sel_cxc_est)]

    if v.empty:
        st.info("No hay facturas con los filtros seleccionados.")
    else:
        kc = st.columns(4)
        kc[0].markdown(tarjeta("Facturado", money(v["Total Facturado"].sum(), True),
                               "Con IVA"), unsafe_allow_html=True)
        kc[1].markdown(tarjeta("Retenciones", money(v["Retenciones"].sum(), True),
                               "Crédito tributario"), unsafe_allow_html=True)
        kc[2].markdown(tarjeta("Cobrado", money(v["Cobrado"].sum(), True),
                               "Recibido en banco"), unsafe_allow_html=True)
        kc[3].markdown(tarjeta("Cartera", money(v["Saldo Pendiente"].sum(), True),
                               "Por recaudar"), unsafe_allow_html=True)

        cc1, cc2 = st.columns([1.25, 1])
        with cc1:
            pend = (v[v["Saldo Pendiente"] > 0].groupby("Cliente")["Saldo Pendiente"]
                    .sum().sort_values())
            if not pend.empty:
                st.plotly_chart(
                    barras_horizontal(pend.index, pend.values, ARENA,
                                      "Cartera pendiente por cliente"),
                    use_container_width=True, config=PLOTLY_CONF)
            else:
                st.info("No hay cartera pendiente con estos filtros.")
        with cc2:
            if "Antiguedad" in v.columns:
                ant = (v[v["Saldo Pendiente"] > 0]
                       .groupby("Antiguedad")["Saldo Pendiente"].sum())
                if not ant.empty:
                    orden_ant = ["Vigente", "1-30", "31-60", "61-90", ">90"]
                    ordenadas = [o for o in orden_ant if o in ant.index]
                    resto = [o for o in ant.index if o not in orden_ant]
                    ant = ant.reindex(ordenadas + resto)
                    fig = go.Figure(go.Bar(
                        x=ant.index.astype(str), y=ant.values, marker_color=ROJO,
                        text=[money(x, True) for x in ant.values],
                        textposition="auto", textfont=dict(size=10), cliponaxis=False,
                        hovertemplate="%{x} · %{y:$,.0f}<extra></extra>"))
                    fig.update_yaxes(tickformat="$~s")
                    st.plotly_chart(estilo(fig, 300, "Cartera por antigüedad", "no"),
                                    use_container_width=True, config=PLOTLY_CONF)

        cols = [c for c in ["Folio", "Fecha", "Cliente", "Proyecto", "Total Facturado",
                            "Retenciones", "Cobrado", "Saldo Pendiente", "Fecha Vence",
                            "Antiguedad", "Estado"] if c in v.columns]
        vista = v[cols].copy()
        for c in ("Fecha", "Fecha Vence"):
            if c in vista.columns:
                vista[c] = pd.to_datetime(vista[c], errors="coerce").dt.strftime("%d/%m/%Y")
        st.dataframe(
            vista.style.format({"Total Facturado": "${:,.0f}", "Retenciones": "${:,.0f}",
                                "Cobrado": "${:,.0f}", "Saldo Pendiente": "${:,.0f}"}),
            use_container_width=True, hide_index=True)

# ─────────────────────── Cuentas por pagar ───────────────────────

if not cxp.empty:
    st.markdown('<div class="seccion">Cuentas por pagar</div>', unsafe_allow_html=True)

    fp1, fp2 = st.columns(2)
    with fp1:
        op_prov = sorted(cxp["Proveedor"].unique())
        sel_prov = filtro("Proveedor", op_prov, "f_cxp_prov")
    with fp2:
        col_alerta = "Alerta" if "Alerta" in cxp.columns else "Estado"
        op_est_p = sorted(cxp[col_alerta].astype(str).unique())
        sel_est_p = filtro(col_alerta, op_est_p, "f_cxp_est")

    p = cxp[cxp["Proveedor"].isin(sel_prov)
            & cxp[col_alerta].astype(str).isin(sel_est_p)]

    if p.empty:
        st.info("No hay documentos con los filtros seleccionados.")
    else:
        kp = st.columns(4)
        kp[0].markdown(tarjeta("Facturado", money(p["Valor Total"].sum(), True),
                               "Proveedores"), unsafe_allow_html=True)
        kp[1].markdown(tarjeta("IVA descontable", money(p["IVA"].sum(), True),
                               "Crédito fiscal"), unsafe_allow_html=True)
        kp[2].markdown(tarjeta("Pagado", money(p["Pagado"].sum(), True),
                               "Ya desembolsado"), unsafe_allow_html=True)
        kp[3].markdown(tarjeta("Por pagar", money(p["Saldo"].sum(), True),
                               "Compromiso de caja"), unsafe_allow_html=True)

        pp1, pp2 = st.columns([1.25, 1])
        with pp1:
            por_prov = (p[p["Saldo"] > 0].groupby("Proveedor")["Saldo"]
                        .sum().sort_values())
            if not por_prov.empty:
                st.plotly_chart(
                    barras_horizontal(por_prov.index, por_prov.values, ROJO,
                                      "Pendiente por proveedor"),
                    use_container_width=True, config=PLOTLY_CONF)
            else:
                st.info("No hay saldos pendientes con estos filtros.")
        with pp2:
            if "Alerta" in p.columns:
                orden_al = ["Vencida", "Por vencer", "Vigente", "Sin fecha"]
                al = (p[p["Saldo"] > 0].groupby("Alerta")["Saldo"].sum())
                al = al.reindex([o for o in orden_al if o in al.index]).dropna()
                if not al.empty:
                    colores = {"Vencida": ROJO, "Por vencer": ARENA,
                               "Vigente": VERDE, "Sin fecha": GRIS}
                    fig = go.Figure(go.Bar(
                        x=al.index.astype(str), y=al.values,
                        marker_color=[colores.get(i, GRIS) for i in al.index],
                        text=[money(v, True) for v in al.values],
                        textposition="auto", textfont=dict(size=10), cliponaxis=False,
                        hovertemplate="%{x} · %{y:$,.0f}<extra></extra>"))
                    fig.update_yaxes(tickformat="$~s")
                    st.plotly_chart(estilo(fig, 320, "Saldo por estado de vencimiento", "no"),
                                    use_container_width=True, config=PLOTLY_CONF)

        # Aviso de lo vencido: es el dato que hay que mirar primero
        if "Alerta" in p.columns:
            vencido = p.loc[p["Alerta"].eq("Vencida"), "Saldo"].sum()
            n_venc = int((p["Alerta"].eq("Vencida") & (p["Saldo"] > 0)).sum())
            if vencido > 0:
                st.warning(
                    f"Hay {money(vencido)} vencidos en {n_venc} facturas. "
                    "Aparecen en la primera semana de la proyección de caja."
                )

        if cxp_n_excl:
            st.caption(
                f"No se muestran {cxp_n_excl} documentos marcados como No en el modelo "
                f"({money(cxp_excluido)}). Siguen contando en el CxP."
            )

        cols = [c for c in ["Fecha Recepción", "Factura", "Numero", "Proveedor", "Cliente",
                            "Proyecto", "Valor Total", "Pagado", "Saldo",
                            "Vence", "Días", "Alerta"] if c in p.columns]
        vista = p[cols].copy()
        for c in ("Fecha Recepción", "Vence"):
            if c in vista.columns:
                vista[c] = pd.to_datetime(vista[c], errors="coerce").dt.strftime("%d/%m/%Y")
        st.dataframe(
            vista.style.format({"Valor Total": "${:,.0f}", "Pagado": "${:,.0f}",
                                "Saldo": "${:,.0f}", "Días": "{:.0f}"}),
            use_container_width=True, hide_index=True)

# ─────────────────────── Proyección de caja ───────────────────────

if not proy_raw.empty:
    fila_desde = fila_por_etiqueta(proy_raw, "Desde")
    fila_saldo = fila_por_etiqueta(proy_raw, "Saldo final")
    if fila_desde is not None and fila_saldo is not None:
        st.markdown('<div class="seccion">Proyección de caja — 13 semanas</div>',
                    unsafe_allow_html=True)

        fechas = []
        for c in range(2, min(16, proy_raw.shape[1])):
            v = proy_raw.iat[fila_desde, c]
            fechas.append(pd.to_datetime(v, errors="coerce"))
        etiquetas_sem = [d.strftime("%d/%m") if pd.notna(d) else "" for d in fechas]

        saldos = valores_fila(proy_raw, fila_saldo)
        cobros = valores_fila(proy_raw, fila_por_etiqueta(proy_raw, "Cobros de cartera (CxC)"))
        f_pagos = fila_por_etiqueta(proy_raw, "Pagos a proveedores (por vencimiento)")
        if f_pagos is None:
            f_pagos = fila_por_etiqueta(proy_raw, "Pagos a proveedores (CxP)")
        pagos = valores_fila(proy_raw, f_pagos)
        egresos = valores_fila(proy_raw, fila_por_etiqueta(proy_raw, "Total egresos"))

        n = min(len(etiquetas_sem), len(saldos))
        etiquetas_sem, saldos = etiquetas_sem[:n], saldos[:n]
        cobros = (cobros + [0] * n)[:n]
        pagos = (pagos + [0] * n)[:n]
        egresos = (egresos + [0] * n)[:n]

        minimo = min(saldos) if saldos else 0
        semana_min = etiquetas_sem[saldos.index(minimo)] if saldos else "—"
        kp2 = st.columns(4)
        kp2[0].markdown(tarjeta("Saldo hoy", money(saldo_caja, True),
                                "Según el extracto"), unsafe_allow_html=True)
        kp2[1].markdown(tarjeta("Cobros proyectados", money(sum(cobros), True),
                                "13 semanas"), unsafe_allow_html=True)
        kp2[2].markdown(tarjeta("Egresos proyectados", money(sum(egresos), True),
                                "13 semanas"), unsafe_allow_html=True)
        kp2[3].markdown(tarjeta("Saldo más bajo", money(minimo, True),
                                f"Semana del {semana_min}"), unsafe_allow_html=True)

        fig = go.Figure()
        fig.add_bar(x=etiquetas_sem, y=cobros, name="Cobros", marker_color=VERDE,
                    hovertemplate="%{x} · %{y:$,.0f}<extra>Cobros</extra>")
        fig.add_bar(x=etiquetas_sem, y=[abs(v) for v in pagos], name="Pagos a proveedores",
                    marker_color=ROJO,
                    hovertemplate="%{x} · %{y:$,.0f}<extra>Pagos</extra>")
        fig.add_scatter(x=etiquetas_sem, y=saldos, name="Saldo proyectado",
                        mode="lines+markers", line=dict(color=AZUL, width=2.5),
                        marker=dict(size=6), yaxis="y",
                        hovertemplate="%{x} · %{y:$,.0f}<extra>Saldo</extra>")
        fig.update_layout(barmode="group")
        fig.update_yaxes(tickformat="$~s")
        st.plotly_chart(estilo(fig, 360, "Cobros, pagos y saldo semana a semana"),
                        use_container_width=True, config=PLOTLY_CONF)

        if minimo < 0:
            st.warning(
                f"La proyección queda en negativo la semana del {semana_min} "
                f"({money(minimo)}). Conviene revisar el calendario de pagos."
            )

# ─────────────────────── IVA ───────────────────────

if not iva_raw.empty:
    fila_bim1 = fila_por_etiqueta(iva_raw, "Bim 1")
    if fila_bim1 is not None:
        st.markdown('<div class="seccion">IVA</div>', unsafe_allow_html=True)

        filas = []
        for i in range(6):
            fila = fila_bim1 + i
            if fila >= len(iva_raw):
                break

            def num(col):
                v = iva_raw.iat[fila, col]
                return float(v) if isinstance(v, (int, float)) and not pd.isna(v) else 0.0

            filas.append({
                "Bimestre": str(iva_raw.iat[fila, 1]),
                "Periodo": str(iva_raw.iat[fila, 2]),
                "IVA generado": num(3),
                "IVA descontable": num(4),
                "ReteIVA": num(5),
                "A pagar": num(6),
                "Ya pagado": num(7),
                "Diferencia": num(8),
            })
        tabla_iva = pd.DataFrame(filas)
        activos = tabla_iva[(tabla_iva["IVA generado"] != 0)
                            | (tabla_iva["IVA descontable"] != 0)]

        if not activos.empty:
            ki = st.columns(3)
            ki[0].markdown(tarjeta("IVA generado", money(activos["IVA generado"].sum(), True),
                                   "Facturas de venta"), unsafe_allow_html=True)
            ki[1].markdown(tarjeta("IVA descontable",
                                   money(activos["IVA descontable"].sum(), True),
                                   "Facturas de compra"), unsafe_allow_html=True)
            ki[2].markdown(tarjeta("Pendiente con la DIAN",
                                   money(activos["Diferencia"].sum(), True),
                                   "Calculado menos girado"), unsafe_allow_html=True)

            fig = go.Figure()
            fig.add_bar(x=activos["Periodo"], y=activos["IVA generado"],
                        name="Generado", marker_color=VERDE,
                        hovertemplate="%{x} · %{y:$,.0f}<extra>Generado</extra>")
            fig.add_bar(x=activos["Periodo"], y=activos["IVA descontable"],
                        name="Descontable", marker_color=AZUL,
                        hovertemplate="%{x} · %{y:$,.0f}<extra>Descontable</extra>")
            fig.add_bar(x=activos["Periodo"], y=activos["A pagar"],
                        name="A pagar", marker_color=ARENA,
                        hovertemplate="%{x} · %{y:$,.0f}<extra>A pagar</extra>")
            fig.update_layout(barmode="group")
            fig.update_yaxes(tickformat="$~s")
            st.plotly_chart(estilo(fig, 320, "IVA por bimestre"),
                            use_container_width=True, config=PLOTLY_CONF)

            st.dataframe(
                activos.style.format({
                    "IVA generado": "${:,.0f}", "IVA descontable": "${:,.0f}",
                    "ReteIVA": "${:,.0f}", "A pagar": "${:,.0f}",
                    "Ya pagado": "${:,.0f}", "Diferencia": "${:,.0f}"}),
                use_container_width=True, hide_index=True)

# ─────────────────────── Composición del gasto ───────────────────────

st.markdown('<div class="seccion">Composición del gasto</div>', unsafe_allow_html=True)
gasto = (f_pyg[f_pyg["valor_neto"] < 0].groupby("cuenta")["valor_neto"]
         .sum().abs().sort_values())
if not gasto.empty:
    st.plotly_chart(barras_horizontal(gasto.index, gasto.values, GRIS,
                                      "Salidas por cuenta", 280),
                    use_container_width=True, config=PLOTLY_CONF)

sin_clasificar = df.loc[df["cuenta"].eq("Revisar"), "valor"].sum()
if abs(sin_clasificar) > 0:
    st.warning(
        f"Hay movimientos sin clasificar por {money(sin_clasificar)}. "
        "No entran al resultado hasta asignarles categoría."
    )
