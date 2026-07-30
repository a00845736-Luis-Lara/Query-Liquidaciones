# -*- coding: utf-8 -*-
"""
Tablero Ejecutivo - Data Query (Control de Salidas)
====================================================
Lee EN VIVO el archivo 'Data Query - estados.xlsx' publicado en SharePoint
(Excel Web) y muestra un tablero tipo "sala de control".

- No requiere subir archivos manualmente: jala los datos desde el link.
- Se refresca solo cada N minutos (configurable en la barra lateral).
- Soporta que se agreguen nuevas filas (nuevos pedidos) sin tocar el codigo.

Autor: generado para el proyecto de dashboards de Operadora Logistica.
"""

import io
import time
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st

# streamlit-autorefresh es opcional; si no esta instalado, seguimos sin auto-refresh.
try:
    from streamlit_autorefresh import st_autorefresh
    HAS_AUTOREFRESH = True
except Exception:
    HAS_AUTOREFRESH = False


# =============================================================================
# CONFIGURACION
# =============================================================================
st.set_page_config(
    page_title="Tablero Ejecutivo · Data Query",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---- Link del archivo en SharePoint -----------------------------------------
# Se lee primero de los "Secrets" de Streamlit (recomendado). Si no existe,
# usa el valor por defecto de abajo. Asi puedes cambiar el link sin tocar codigo.
DEFAULT_SHARE_URL = (
    "https://diszasa365-my.sharepoint.com/:x:/g/personal/"
    "jaguilar_operadoralogistica_com/"
    "IQCoGmfBr55DS4E019WnIjZDAQx_bn-W09KKi1_nCjMx-hY"
)
SHARE_URL = st.secrets.get("DATA_QUERY_URL", DEFAULT_SHARE_URL)

SHEET_NAME = "ControlSalidas_Live"   # hoja con los datos
TTL_SEGUNDOS = 300                   # cache: cuanto vive la descarga antes de re-jalar

# ---- Paleta (tema oscuro "sala de control") ---------------------------------
CLR = {
    "bg":        "#0a0e1a",
    "panel":     "#131a2b",
    "panel2":    "#1b2438",
    "border":    "#243049",
    "text":      "#e6ecf5",
    "muted":     "#8a96ad",
    "verde":     "#2ecc71",
    "amarillo":  "#f1c40f",
    "rojo":      "#e74c3c",
    "azul":      "#3498db",
    "morado":    "#9b59b6",
    "cyan":      "#1abc9c",
}


# =============================================================================
# CARGA DE DATOS
# =============================================================================
def _build_download_url(share_url: str) -> str:
    """Convierte un link de 'Compartir' de SharePoint en link de descarga directa."""
    share_url = share_url.strip()
    if "download=1" in share_url:
        return share_url
    sep = "&" if "?" in share_url else "?"
    return f"{share_url}{sep}download=1"


@st.cache_data(ttl=TTL_SEGUNDOS, show_spinner=False)
def descargar_excel(share_url: str) -> bytes:
    """Descarga los bytes del .xlsx desde SharePoint. Cacheado por TTL_SEGUNDOS."""
    url = _build_download_url(share_url)
    headers = {"User-Agent": "Mozilla/5.0 (compatible; DashboardBot/1.0)"}
    resp = requests.get(url, headers=headers, allow_redirects=True, timeout=60)
    resp.raise_for_status()
    contenido = resp.content
    # Un .xlsx es un ZIP: debe empezar con los bytes magicos 'PK'.
    if contenido[:2] != b"PK":
        raise ValueError(
            "El link no devolvio un archivo Excel. Lo mas probable es que "
            "no este publico ('Cualquier persona con el enlace') o que haya "
            "expirado. Revisa los permisos del archivo en SharePoint."
        )
    return contenido


def excel_serial_a_fecha(serie: pd.Series) -> pd.Series:
    """Convierte fechas seriales de Excel (numeros) a datetime real."""
    s = pd.to_numeric(serie, errors="coerce")
    return pd.to_datetime(s, unit="D", origin="1899-12-30")


def limpiar_estado(serie: pd.Series) -> pd.Series:
    """Normaliza el campo Estado (une variantes 'W/w (sin catalogo)')."""
    s = serie.astype(str).str.strip()
    s = s.str.replace(r"^[Ww]\s*\(sin catalogo\)$", "Sin catálogo", regex=True)
    return s


@st.cache_data(ttl=TTL_SEGUNDOS, show_spinner=False)
def cargar_datos(share_url: str) -> pd.DataFrame:
    """Descarga, lee y transforma la hoja ControlSalidas_Live en un DataFrame limpio."""
    contenido = descargar_excel(share_url)
    df = pd.read_excel(io.BytesIO(contenido), sheet_name=SHEET_NAME)

    # Tipos y campos derivados (defensivo: solo si la columna existe)
    if "Valor" in df:
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0.0)
    if "Avance %" in df:
        df["Avance %"] = pd.to_numeric(df["Avance %"], errors="coerce")
    if "Estado" in df:
        df["Estado"] = limpiar_estado(df["Estado"])
    if "Fecha de creacion" in df:
        df["f_creacion"] = excel_serial_a_fecha(df["Fecha de creacion"])
    if "Fecha cierre" in df:
        df["f_cierre"] = excel_serial_a_fecha(df["Fecha cierre"])

    # Cuenta como texto (son categorias: 1, 2, 3) para que no se sume ni promedie
    if "Cuenta" in df:
        df["Cuenta"] = df["Cuenta"].astype(str).str.strip()

    return df


# =============================================================================
# ESTILOS (CSS)
# =============================================================================
def inyectar_css():
    st.markdown(
        f"""
        <style>
        .stApp {{ background-color: {CLR['bg']}; color: {CLR['text']}; }}
        section[data-testid="stSidebar"] {{ background-color: {CLR['panel']}; }}
        #MainMenu, footer {{ visibility: hidden; }}

        .titulo-app {{
            font-size: 26px; font-weight: 800; letter-spacing: .5px;
            color: {CLR['text']}; margin: 0;
        }}
        .subtitulo-app {{ color: {CLR['muted']}; font-size: 13px; margin-top: 2px; }}

        .kpi {{
            background: {CLR['panel']};
            border: 1px solid {CLR['border']};
            border-radius: 14px; padding: 16px 18px; height: 100%;
        }}
        .kpi .etq {{ color: {CLR['muted']}; font-size: 12px; text-transform: uppercase;
                     letter-spacing: .6px; margin-bottom: 6px; }}
        .kpi .val {{ font-size: 30px; font-weight: 800; line-height: 1.1; }}
        .kpi .sub {{ color: {CLR['muted']}; font-size: 12px; margin-top: 4px; }}

        .seccion {{
            font-size: 15px; font-weight: 700; color: {CLR['text']};
            border-left: 4px solid {CLR['azul']}; padding-left: 10px;
            margin: 8px 0 4px 0;
        }}
        .alerta {{
            background: {CLR['panel']}; border: 1px solid {CLR['border']};
            border-left: 4px solid {CLR['rojo']};
            border-radius: 10px; padding: 10px 14px; margin-bottom: 8px;
            font-size: 13px;
        }}
        .badge {{ display:inline-block; padding:2px 10px; border-radius:20px;
                  font-size:11px; font-weight:700; }}
        div[data-testid="stMetricValue"] {{ font-size: 26px; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def tarjeta_kpi(etiqueta, valor, sub="", color=None):
    color = color or CLR["text"]
    st.markdown(
        f"""
        <div class="kpi">
            <div class="etq">{etiqueta}</div>
            <div class="val" style="color:{color}">{valor}</div>
            <div class="sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def fmt_money(x):
    return f"${x:,.0f}"


def fmt_num(x):
    return f"{x:,}"


# =============================================================================
# APP
# =============================================================================
def main():
    inyectar_css()

    # ---------------- Barra lateral ----------------
    with st.sidebar:
        st.markdown("### ⚙️ Controles")
        minutos = st.slider("Auto-actualizar cada (min)", 1, 30, 5)
        if st.button("🔄 Actualizar ahora", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
        st.divider()

    # Auto-refresh
    if HAS_AUTOREFRESH:
        st_autorefresh(interval=minutos * 60 * 1000, key="auto_refresh")

    # ---------------- Cabecera ----------------
    izq, der = st.columns([3, 1])
    with izq:
        st.markdown('<p class="titulo-app">📦 Tablero Ejecutivo · Control de Salidas</p>',
                    unsafe_allow_html=True)
        st.markdown('<p class="subtitulo-app">Fuente: Data Query — SharePoint (en vivo)</p>',
                    unsafe_allow_html=True)
    with der:
        st.markdown(
            f'<p class="subtitulo-app" style="text-align:right">Actualizado<br>'
            f'<b style="color:{CLR["text"]}">{datetime.now().strftime("%d/%m/%Y %H:%M:%S")}</b></p>',
            unsafe_allow_html=True,
        )

    # ---------------- Carga ----------------
    try:
        df = cargar_datos(SHARE_URL)
    except Exception as e:
        st.error("No se pudieron cargar los datos desde SharePoint.")
        st.exception(e)
        st.info(
            "Verifica que el archivo este compartido como **'Cualquier persona "
            "con el enlace'** y que el link en los *Secrets* sea correcto."
        )
        st.stop()

    if df.empty:
        st.warning("El archivo se leyo pero no tiene filas.")
        st.stop()

    # ---------------- Filtros (sidebar) ----------------
    with st.sidebar:
        st.markdown("### 🔍 Filtros")
        cuentas = sorted(df["Cuenta"].dropna().unique()) if "Cuenta" in df else []
        f_cuenta = st.multiselect("Cuenta", cuentas, default=cuentas)

        estados = sorted(df["Estado"].dropna().unique()) if "Estado" in df else []
        f_estado = st.multiselect("Estado", estados, default=estados)

        if "f_creacion" in df and df["f_creacion"].notna().any():
            fmin = df["f_creacion"].min().date()
            fmax = df["f_creacion"].max().date()
            rango = st.date_input("Rango de fecha (creación)", (fmin, fmax),
                                  min_value=fmin, max_value=fmax)
        else:
            rango = None

    # Aplicar filtros
    d = df.copy()
    if f_cuenta and "Cuenta" in d:
        d = d[d["Cuenta"].isin(f_cuenta)]
    if f_estado and "Estado" in d:
        d = d[d["Estado"].isin(f_estado)]
    if rango and isinstance(rango, (list, tuple)) and len(rango) == 2 and "f_creacion" in d:
        ini, fin = pd.Timestamp(rango[0]), pd.Timestamp(rango[1]) + pd.Timedelta(days=1)
        d = d[(d["f_creacion"] >= ini) & (d["f_creacion"] < fin)]

    if d.empty:
        st.warning("Ningún registro coincide con los filtros seleccionados.")
        st.stop()

    # ---------------- KPIs ----------------
    total_docs = len(d)
    valor_total = float(d["Valor"].sum()) if "Valor" in d else 0.0
    despachados = int((d["Estado"] == "DESPACHADO").sum()) if "Estado" in d else 0
    sin_cerrar = int(d["f_cierre"].isna().sum()) if "f_cierre" in d else 0
    avance_prom = float(d["Avance %"].mean()) if "Avance %" in d else 0.0
    pct_desp = (despachados / total_docs * 100) if total_docs else 0

    st.markdown("")  # espacio
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        tarjeta_kpi("Documentos", fmt_num(total_docs), "en el periodo filtrado", CLR["azul"])
    with c2:
        tarjeta_kpi("Valor total", fmt_money(valor_total), "suma del campo Valor", CLR["verde"])
    with c3:
        col = CLR["verde"] if pct_desp >= 90 else CLR["amarillo"] if pct_desp >= 70 else CLR["rojo"]
        tarjeta_kpi("Despachados", fmt_num(despachados), f"{pct_desp:.1f}% del total", col)
    with c4:
        col = CLR["verde"] if sin_cerrar == 0 else CLR["amarillo"] if sin_cerrar < 50 else CLR["rojo"]
        tarjeta_kpi("Sin cerrar", fmt_num(sin_cerrar), "documentos sin fecha de cierre", col)
    with c5:
        col = CLR["verde"] if avance_prom >= 90 else CLR["amarillo"] if avance_prom >= 70 else CLR["rojo"]
        tarjeta_kpi("Avance prom.", f"{avance_prom:.1f}%", "promedio de Avance %", col)

    st.markdown("<br>", unsafe_allow_html=True)

    # ---------------- Graficas fila 1 ----------------
    g1, g2 = st.columns([1.3, 1])

    with g1:
        st.markdown('<div class="seccion">Documentos por Estado</div>', unsafe_allow_html=True)
        est = (d["Estado"].value_counts().rename_axis("Estado")
               .reset_index(name="Documentos").sort_values("Documentos"))
        fig = px.bar(est, x="Documentos", y="Estado", orientation="h", text="Documentos")
        fig.update_traces(marker_color=CLR["azul"], textposition="outside")
        _tema_plotly(fig, alto=max(260, 40 * len(est)))
        st.plotly_chart(fig, use_container_width=True)

    with g2:
        st.markdown('<div class="seccion">Valor por Cuenta</div>', unsafe_allow_html=True)
        if "Cuenta" in d and "Valor" in d:
            val_cta = (d.groupby("Cuenta")["Valor"].sum()
                       .reset_index().sort_values("Valor", ascending=False))
            fig = go.Figure(go.Pie(
                labels=[f"Cuenta {c}" for c in val_cta["Cuenta"]],
                values=val_cta["Valor"], hole=0.55,
                marker=dict(colors=[CLR["azul"], CLR["cyan"], CLR["morado"], CLR["amarillo"]]),
                textinfo="label+percent",
            ))
            fig.update_layout(annotations=[dict(
                text=fmt_money(val_cta["Valor"].sum()), x=0.5, y=0.5,
                font=dict(size=16, color=CLR["text"]), showarrow=False)])
            _tema_plotly(fig, alto=320)
            st.plotly_chart(fig, use_container_width=True)

    # ---------------- Grafica tendencia ----------------
    st.markdown('<div class="seccion">Tendencia diaria (documentos y valor)</div>',
                unsafe_allow_html=True)
    if "f_creacion" in d and d["f_creacion"].notna().any():
        tmp = d.dropna(subset=["f_creacion"]).copy()
        tmp["dia"] = tmp["f_creacion"].dt.date
        por_dia = tmp.groupby("dia").agg(
            Documentos=("Id", "count") if "Id" in tmp else ("Estado", "count"),
            Valor=("Valor", "sum"),
        ).reset_index()

        fig = go.Figure()
        fig.add_bar(x=por_dia["dia"], y=por_dia["Documentos"],
                    name="Documentos", marker_color=CLR["azul"], opacity=0.65)
        fig.add_trace(go.Scatter(
            x=por_dia["dia"], y=por_dia["Valor"], name="Valor",
            yaxis="y2", mode="lines", line=dict(color=CLR["verde"], width=2)))
        fig.update_layout(
            yaxis=dict(title="Documentos"),
            yaxis2=dict(title="Valor", overlaying="y", side="right", showgrid=False),
            legend=dict(orientation="h", y=1.12, x=0),
        )
        _tema_plotly(fig, alto=320)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.caption("Sin fechas válidas para graficar tendencia.")

    # ---------------- Alertas ----------------
    st.markdown('<div class="seccion">🚨 Alertas operativas</div>', unsafe_allow_html=True)
    avance_incompleto = int((d["Avance %"] < 100).sum()) if "Avance %" in d else 0
    sin_catalogo = int((d["Estado"] == "Sin catálogo").sum()) if "Estado" in d else 0

    a1, a2, a3 = st.columns(3)
    with a1:
        _alerta("Documentos sin cerrar", sin_cerrar,
                "sin fecha de cierre registrada", CLR["rojo"] if sin_cerrar else CLR["verde"])
    with a2:
        _alerta("Avance incompleto (<100%)", avance_incompleto,
                "aún en proceso", CLR["amarillo"] if avance_incompleto else CLR["verde"])
    with a3:
        _alerta("Estados sin catalogar", sin_catalogo,
                "requieren revisión de captura", CLR["rojo"] if sin_catalogo else CLR["verde"])

    # ---------------- Tabla detalle ----------------
    st.markdown('<div class="seccion">📋 Detalle de documentos</div>', unsafe_allow_html=True)
    buscar = st.text_input("Buscar por ID o Documento", placeholder="Escribe un ID o número de documento…")

    cols_mostrar = [c for c in
                    ["Id", "Documento", "Cuenta", "Estado", "Valor", "Avance %",
                     "Usuario", "Ubicacion", "f_creacion", "f_cierre"]
                    if c in d.columns]
    tabla = d[cols_mostrar].copy()

    if buscar:
        b = buscar.strip().lower()
        mask = pd.Series(False, index=tabla.index)
        if "Id" in tabla:
            mask |= tabla["Id"].astype(str).str.lower().str.contains(b, na=False)
        if "Documento" in tabla:
            mask |= tabla["Documento"].astype(str).str.lower().str.contains(b, na=False)
        tabla = tabla[mask]

    # Formato de fechas para lectura
    for c in ["f_creacion", "f_cierre"]:
        if c in tabla:
            tabla[c] = pd.to_datetime(tabla[c]).dt.strftime("%d/%m/%Y %H:%M")
    tabla = tabla.rename(columns={"f_creacion": "Creación", "f_cierre": "Cierre"})

    st.dataframe(
        tabla, use_container_width=True, height=420, hide_index=True,
        column_config={
            "Valor": st.column_config.NumberColumn("Valor", format="$%.2f"),
            "Avance %": st.column_config.NumberColumn("Avance %", format="%.1f%%"),
        },
    )
    st.caption(f"Mostrando {len(tabla):,} de {total_docs:,} documentos filtrados.")

    # Descargar lo filtrado
    st.download_button(
        "⬇️ Descargar tabla (CSV)",
        data=tabla.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"data_query_{datetime.now():%Y%m%d_%H%M}.csv",
        mime="text/csv",
    )


# =============================================================================
# HELPERS DE PRESENTACION
# =============================================================================
def _tema_plotly(fig, alto=320):
    fig.update_layout(
        height=alto,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=CLR["text"], size=12),
        margin=dict(l=10, r=10, t=10, b=10),
    )
    fig.update_xaxes(gridcolor=CLR["border"], zerolinecolor=CLR["border"])
    fig.update_yaxes(gridcolor=CLR["border"], zerolinecolor=CLR["border"])


def _alerta(titulo, valor, sub, color):
    st.markdown(
        f"""
        <div class="alerta" style="border-left-color:{color}">
            <div style="color:{CLR['muted']};font-size:12px">{titulo}</div>
            <div style="font-size:24px;font-weight:800;color:{color}">{valor:,}</div>
            <div style="color:{CLR['muted']};font-size:11px">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
