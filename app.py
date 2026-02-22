"""
app.py — FarmaAnalytics · Aplicación principal
================================================
Panel web profesional de análisis de ventas para farmacias.

Estructura:
  · Autenticación con streamlit-authenticator
  · Pantalla de bienvenida + carga de Excel si no hay datos
  · Dashboard con sidebar de filtros globales y 7 páginas de análisis
"""

import warnings
from datetime import datetime, timedelta

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# Debe ir ANTES de cualquier otra llamada a st.*
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FarmaAnalytics",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CSS — TEMA OSCURO PROFESIONAL
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
    /* ── Fondo general ──────────────────────────────── */
    .stApp { background-color: #1a1a2e; }

    /* ── Sidebar ────────────────────────────────────── */
    [data-testid="stSidebar"] {
        background-color: #16213e;
        border-right: 1px solid #2ECC71;
    }

    /* ── Texto general ──────────────────────────────── */
    .stMarkdown p, .stMarkdown li, label, .stText { color: #ECF0F1; }

    /* ── Títulos ────────────────────────────────────── */
    h1 { color: #2ECC71 !important; }
    h2 { color: #3498DB !important; }
    h3 { color: #ECF0F1 !important; }

    /* ── Botones primarios ──────────────────────────── */
    .stButton > button {
        background-color: #2ECC71;
        color: #1a1a2e;
        border: none;
        border-radius: 6px;
        font-weight: 700;
        transition: background 0.2s;
    }
    .stButton > button:hover {
        background-color: #27AE60;
        color: #ffffff;
    }

    /* ── File uploader ──────────────────────────────── */
    [data-testid="stFileUploader"] {
        background-color: #16213e;
        border: 2px dashed #2ECC71;
        border-radius: 10px;
        padding: 1rem;
    }

    /* ── Métricas ───────────────────────────────────── */
    [data-testid="stMetric"] {
        background-color: #16213e;
        border-left: 4px solid #2ECC71;
        border-radius: 8px;
        padding: 0.8rem 1rem;
    }
    [data-testid="stMetricLabel"] { color: #BDC3C7 !important; }
    [data-testid="stMetricValue"] { color: #2ECC71 !important; }

    /* ── DataFrames ─────────────────────────────────── */
    [data-testid="stDataFrame"] {
        border: 1px solid #2ECC71;
        border-radius: 8px;
    }

    /* ── Alertas / Info ─────────────────────────────── */
    .stAlert { background-color: #16213e; border-radius: 8px; }

    /* ── Separador ──────────────────────────────────── */
    hr { border-color: #2ECC71 !important; opacity: 0.3; }

    /* ── Selectbox / Multiselect / Slider ───────────── */
    [data-testid="stSelectbox"] > div > div,
    [data-testid="stMultiSelect"] > div > div {
        background-color: #16213e;
        border-color: #2ECC71;
    }

    /* ── Radio buttons ──────────────────────────────── */
    [data-testid="stRadio"] label { color: #ECF0F1 !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────────────────────────────────
# IMPORTACIONES LOCALES
# ─────────────────────────────────────────────────────────────────────────────
from auth import create_authenticator, get_user_role, render_login_ui
from data_processor import apply_filters, process_excel


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS GLOBALES
# ─────────────────────────────────────────────────────────────────────────────

def euros(value, decimals: int = 2) -> str:
    """
    Formatea un número como moneda en euros con formato español.
    Ejemplo: 1234567.89  →  "1.234.567,89 €"
    """
    if pd.isna(value) or value is None:
        return "0,00 €"
    sign = "-" if value < 0 else ""
    formatted = (
        f"{abs(value):,.{decimals}f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )
    return f"{sign}{formatted} €"


def render_footer(df: pd.DataFrame):
    """Muestra el footer con estadísticas del dataset en todas las páginas."""
    st.markdown("---")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.caption(f"📊 Total registros cargados: **{len(df):,}**")
    with c2:
        if "Fecha_ES" in df.columns and not df["Fecha_ES"].isna().all():
            min_d = df["Fecha_ES"].min().strftime("%d/%m/%Y")
            max_d = df["Fecha_ES"].max().strftime("%d/%m/%Y")
            st.caption(f"📅 Rango del dataset: **{min_d}** → **{max_d}**")
    with c3:
        st.caption("💊 FarmaAnalytics © 2024")


def no_data_message():
    """Mensaje cuando no hay datos con los filtros aplicados."""
    st.markdown(
        """
        <div style="text-align:center; padding:3rem 2rem;
                    background:#16213e; border-radius:12px; margin:2rem 0;">
            <h3 style="color:#E74C3C;">⚠️ Sin datos disponibles</h3>
            <p style="color:#BDC3C7;">
                No hay registros que coincidan con los filtros aplicados.<br>
                Prueba a ampliar el rango de fechas o modifica los filtros del sidebar.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# Paleta de colores para organismos (consistente en todos los gráficos)
_ORG_COLORS = [
    "#2ECC71", "#3498DB", "#E74C3C", "#F39C12",
    "#9B59B6", "#1ABC9C", "#E67E22", "#E91E63",
]

# Grupos considerados "receta"
_RECETA_GRUPOS = {"RECETA XXI", "MUFACE", "ISFAS", "MUGEJU"}


def _plot_layout(title: str = "", height: int = 420, margin_b: int = 50) -> dict:
    """
    Devuelve un dict de parámetros de layout Plotly con el tema oscuro del proyecto.
    Usar como: fig.update_layout(**_plot_layout("Mi título"))
    """
    return dict(
        template="plotly_dark",
        paper_bgcolor="#1a1a2e",
        plot_bgcolor="#16213e",
        font=dict(color="#ECF0F1", family="sans-serif"),
        title=dict(text=title, font=dict(color="#ECF0F1", size=15)) if title else {},
        height=height,
        margin=dict(l=10, r=10, t=45 if title else 20, b=margin_b),
        xaxis=dict(gridcolor="#2c2c4e", linecolor="#2c2c4e"),
        yaxis=dict(gridcolor="#2c2c4e", linecolor="#2c2c4e"),
    )


# ─────────────────────────────────────────────────────────────────────────────
# PÁGINAS  (estructura base — contenido completo en siguientes versiones)
# ─────────────────────────────────────────────────────────────────────────────

def page_resumen(df: pd.DataFrame, df_filtered: pd.DataFrame):
    """Página 1 — Resumen Ejecutivo: 8 KPIs + gráfico diario + gráfico mensual."""
    st.title("📊 Resumen Ejecutivo")

    if df_filtered.empty:
        no_data_message()
        render_footer(df)
        return

    # ── Cálculo de KPIs ───────────────────────────────────────────────────────
    facturacion   = df_filtered["Facturación"].sum()       if "Facturación"        in df_filtered.columns else 0.0
    pvp_facturado = df_filtered["PVP Facturado"].sum()     if "PVP Facturado"      in df_filtered.columns else 0.0
    importe_neto  = df_filtered["Importe Neto"].sum()      if "Importe Neto"       in df_filtered.columns else 0.0
    aport_gob     = df_filtered["Diferencia_Gobierno"].sum() if "Diferencia_Gobierno" in df_filtered.columns else 0.0
    n_ops         = len(df_filtered)
    ticket_medio  = facturacion / n_ops if n_ops > 0 else 0.0
    n_productos   = df_filtered["Código"].nunique() if "Código" in df_filtered.columns else 0

    # % Receta: facturación procedente de organismos de receta vs total
    if "Organismo_Grupo" in df_filtered.columns and facturacion > 0:
        fact_receta = df_filtered.loc[
            df_filtered["Organismo_Grupo"].isin(_RECETA_GRUPOS), "Facturación"
        ].sum()
        pct_receta = fact_receta / facturacion * 100
    else:
        pct_receta = 0.0

    # ── Fila 1: 4 KPIs financieros ────────────────────────────────────────────
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("💶 Total Facturación",    euros(facturacion))
    with col2:
        st.metric("🧾 PVP Facturado",        euros(pvp_facturado))
    with col3:
        st.metric("💳 Importe Neto",         euros(importe_neto))
    with col4:
        st.metric("🏛️ Aportación Gobierno",  euros(aport_gob))

    # ── Fila 2: 4 KPIs operativos ─────────────────────────────────────────────
    col5, col6, col7, col8 = st.columns(4)
    with col5:
        st.metric("🔢 Nº Operaciones",       f"{n_ops:,}")
    with col6:
        st.metric("🎫 Ticket Medio",          euros(ticket_medio))
    with col7:
        st.metric(
            "💊 % Receta vs V.Libre",
            f"{pct_receta:.1f} %",
            help="% de facturación procedente de receta (RECETA XXI, MUFACE, ISFAS, MUGEJU)",
        )
    with col8:
        st.metric("📦 Productos distintos",   f"{n_productos:,}")

    st.markdown("---")

    # ── Gráfico 1: facturación diaria + media móvil 7 días ────────────────────
    st.subheader("📈 Facturación Diaria")

    if "Fecha_ES" in df_filtered.columns and "Facturación" in df_filtered.columns:
        daily = (
            df_filtered
            .groupby(df_filtered["Fecha_ES"].dt.date)["Facturación"]
            .sum()
            .reset_index()
        )
        daily.columns = ["Fecha", "Facturación"]
        daily = daily.sort_values("Fecha")
        daily["MM7"] = daily["Facturación"].rolling(window=7, min_periods=1).mean()

        fig_line = go.Figure()

        # Área rellena para la facturación diaria
        fig_line.add_trace(go.Scatter(
            x=daily["Fecha"],
            y=daily["Facturación"],
            name="Facturación diaria",
            mode="lines",
            line=dict(color="#3498DB", width=1.5),
            fill="tozeroy",
            fillcolor="rgba(52,152,219,0.12)",
            hovertemplate="%{x|%d/%m/%Y}<br><b>%{y:,.2f} €</b><extra>Diario</extra>",
        ))

        # Línea de media móvil destacada
        fig_line.add_trace(go.Scatter(
            x=daily["Fecha"],
            y=daily["MM7"],
            name="Media móvil 7 días",
            mode="lines",
            line=dict(color="#2ECC71", width=2.5),
            hovertemplate="%{x|%d/%m/%Y}<br><b>%{y:,.2f} €</b><extra>MM7</extra>",
        ))

        fig_line.update_layout(
            **_plot_layout(height=380),
            xaxis_title="Fecha",
            yaxis_title="Facturación (€)",
            hovermode="x unified",
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02,
                xanchor="right", x=1, bgcolor="rgba(0,0,0,0)",
            ),
        )
        st.plotly_chart(fig_line, use_container_width=True)

    st.markdown("---")

    # ── Gráfico 2: PVP Facturado vs Importe Neto por mes (barras superpuestas) ─
    st.subheader("📊 Comparativa Mensual: PVP Facturado vs Importe Neto")

    cols_needed = {"Mes_Num", "Mes", "Año", "PVP Facturado", "Importe Neto"}
    if cols_needed.issubset(df_filtered.columns):
        monthly = (
            df_filtered
            .groupby(["Año", "Mes_Num", "Mes"], as_index=False)
            .agg(
                PVP_Facturado=("PVP Facturado", "sum"),
                Importe_Neto=("Importe Neto", "sum"),
            )
            .sort_values(["Año", "Mes_Num"])
        )
        # Etiqueta legible: "Ene 2024"
        _MES_ABREV = {
            "January":"Ene","February":"Feb","March":"Mar","April":"Abr",
            "May":"May","June":"Jun","July":"Jul","August":"Ago",
            "September":"Sep","October":"Oct","November":"Nov","December":"Dic",
        }
        monthly["Periodo"] = (
            monthly["Mes"].map(_MES_ABREV).fillna(monthly["Mes"])
            + " "
            + monthly["Año"].astype(str)
        )

        fig_bar = go.Figure()

        # Barra exterior: PVP Facturado (azul, semi-transparente)
        fig_bar.add_trace(go.Bar(
            x=monthly["Periodo"],
            y=monthly["PVP_Facturado"],
            name="PVP Facturado",
            marker_color="rgba(52,152,219,0.75)",
            hovertemplate="<b>%{x}</b><br>PVP Facturado: %{y:,.2f} €<extra></extra>",
        ))

        # Barra superpuesta: Importe Neto (verde, semi-transparente)
        fig_bar.add_trace(go.Bar(
            x=monthly["Periodo"],
            y=monthly["Importe_Neto"],
            name="Importe Neto",
            marker_color="rgba(46,204,113,0.85)",
            hovertemplate="<b>%{x}</b><br>Importe Neto: %{y:,.2f} €<extra></extra>",
        ))

        fig_bar.update_layout(
            **_plot_layout(height=380),
            barmode="overlay",          # barras superpuestas
            xaxis_title="Mes",
            yaxis_title="Importe (€)",
            bargap=0.15,
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02,
                xanchor="right", x=1, bgcolor="rgba(0,0,0,0)",
            ),
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    render_footer(df)


def page_organismos(df: pd.DataFrame, df_filtered: pd.DataFrame):
    """
    Página 2 — Organismos.
    · Barras horizontales por organismo (orden descendente)
    · Donut de distribución % + barras apiladas Genérico vs Marca
    · Tabla interactiva con todas las métricas
    """
    st.title("🏥 Organismos")

    if df_filtered.empty:
        no_data_message()
        render_footer(df)
        return

    # ── Agregar por Organismo_Grupo ───────────────────────────────────────────
    agg_cols = {"Facturación": ("Facturación", "sum")}
    if "Cantidad (Unidades)" in df_filtered.columns:
        agg_cols["Unidades"] = ("Cantidad (Unidades)", "sum")
    if "PVP Facturado" in df_filtered.columns:
        agg_cols["PVP_Facturado"] = ("PVP Facturado", "sum")
    if "Importe Neto" in df_filtered.columns:
        agg_cols["Importe_Neto"] = ("Importe Neto", "sum")
    if "Diferencia_Gobierno" in df_filtered.columns:
        agg_cols["Diferencia_Gobierno"] = ("Diferencia_Gobierno", "sum")

    grp = (
        df_filtered
        .groupby("Organismo_Grupo", as_index=False)
        .agg(Operaciones=("Facturación", "count"), **agg_cols)
        .sort_values("Facturación", ascending=False)
    )

    total_fact = grp["Facturación"].sum()
    grp["Ticket_Medio"] = grp["Facturación"] / grp["Operaciones"].replace(0, 1)
    grp["Pct_Total"]    = grp["Facturación"] / total_fact * 100 if total_fact > 0 else 0.0

    org_list  = grp["Organismo_Grupo"].tolist()
    colores   = _ORG_COLORS[: len(org_list)]
    color_map = dict(zip(org_list, colores))

    # ── Barras horizontales: Facturación por Organismo ────────────────────────
    st.subheader("💶 Facturación por Organismo")

    fig_barh = go.Figure(go.Bar(
        x=grp["Facturación"],
        y=grp["Organismo_Grupo"],
        orientation="h",
        marker=dict(
            color=[color_map.get(o, "#2ECC71") for o in org_list],
        ),
        text=[euros(v) for v in grp["Facturación"]],
        textposition="outside",
        textfont=dict(color="#ECF0F1", size=12),
        hovertemplate=(
            "<b>%{y}</b><br>"
            "Facturación: <b>%{x:,.2f} €</b><extra></extra>"
        ),
        cliponaxis=False,
    ))
    fig_barh.update_layout(
        **_plot_layout(height=max(300, len(grp) * 52 + 60), margin_b=40),
        xaxis_title="Facturación (€)",
        yaxis=dict(autorange="reversed", tickfont=dict(size=12)),
        margin=dict(l=10, r=140, t=20, b=40),
    )
    st.plotly_chart(fig_barh, use_container_width=True)

    st.markdown("---")

    # ── Donut + Barras apiladas (lado a lado) ─────────────────────────────────
    col_donut, col_stack = st.columns(2)

    # — Donut: distribución % de facturación —
    with col_donut:
        st.subheader("🍩 Distribución de Facturación")
        fig_donut = go.Figure(go.Pie(
            labels=grp["Organismo_Grupo"],
            values=grp["Facturación"],
            hole=0.52,
            marker=dict(colors=colores, line=dict(color="#1a1a2e", width=2)),
            textinfo="percent+label",
            textfont=dict(size=11),
            hovertemplate=(
                "<b>%{label}</b><br>"
                "%{percent:.1%}<br>"
                "%{value:,.2f} €<extra></extra>"
            ),
            direction="clockwise",
            sort=False,
        ))
        # Anotación central
        fig_donut.add_annotation(
            text=f"<b>{euros(total_fact)}</b>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=13, color="#2ECC71"),
        )
        fig_donut.update_layout(
            **_plot_layout(height=380),
            showlegend=True,
            legend=dict(
                orientation="v", yanchor="middle", y=0.5,
                xanchor="left", x=1.02,
                font=dict(size=10), bgcolor="rgba(0,0,0,0)",
            ),
            margin=dict(l=0, r=120, t=20, b=20),
        )
        st.plotly_chart(fig_donut, use_container_width=True)

    # — Barras apiladas: Genérico vs Marca por organismo —
    with col_stack:
        st.subheader("💊 Genérico vs Marca")

        if "Es_Generico" in df_filtered.columns:
            stacked = (
                df_filtered
                .groupby(["Organismo_Grupo", "Es_Generico"], as_index=False)["Facturación"]
                .sum()
            )
            gen_s   = stacked[stacked["Es_Generico"] == True].set_index("Organismo_Grupo")["Facturación"]
            marca_s = stacked[stacked["Es_Generico"] == False].set_index("Organismo_Grupo")["Facturación"]

            fig_stack = go.Figure()
            fig_stack.add_trace(go.Bar(
                name="Genérico",
                x=org_list,
                y=[gen_s.get(o, 0) for o in org_list],
                marker_color="#2ECC71",
                hovertemplate="<b>%{x}</b><br>Genérico: %{y:,.2f} €<extra></extra>",
            ))
            fig_stack.add_trace(go.Bar(
                name="Marca",
                x=org_list,
                y=[marca_s.get(o, 0) for o in org_list],
                marker_color="#3498DB",
                hovertemplate="<b>%{x}</b><br>Marca: %{y:,.2f} €<extra></extra>",
            ))
            fig_stack.update_layout(
                **_plot_layout(height=380, margin_b=90),
                barmode="stack",
                xaxis=dict(
                    tickangle=-35,
                    tickfont=dict(size=10),
                    gridcolor="#2c2c4e",
                ),
                yaxis_title="Facturación (€)",
                legend=dict(
                    orientation="h", yanchor="bottom", y=1.02,
                    xanchor="right", x=1, bgcolor="rgba(0,0,0,0)",
                ),
                margin=dict(l=10, r=10, t=20, b=90),
            )
            st.plotly_chart(fig_stack, use_container_width=True)
        else:
            st.info("No se encontró la columna 'Es_Generico' en los datos.")

    st.markdown("---")

    # ── Tabla interactiva ─────────────────────────────────────────────────────
    st.subheader("📋 Detalle por Organismo")

    tabla = grp.copy()

    # Construir columnas de visualización con formato
    tabla_display = pd.DataFrame()
    tabla_display["Organismo"]           = tabla["Organismo_Grupo"]
    tabla_display["Operaciones"]         = tabla["Operaciones"].apply(lambda x: f"{x:,}")
    if "Unidades" in tabla.columns:
        tabla_display["Unidades"]        = tabla["Unidades"].apply(lambda x: f"{x:,.0f}")
    tabla_display["Facturación"]         = tabla["Facturación"].apply(euros)
    if "PVP_Facturado" in tabla.columns:
        tabla_display["PVP Facturado"]   = tabla["PVP_Facturado"].apply(euros)
    if "Importe_Neto" in tabla.columns:
        tabla_display["Importe Neto"]    = tabla["Importe_Neto"].apply(euros)
    if "Diferencia_Gobierno" in tabla.columns:
        tabla_display["Dif. Gobierno"]   = tabla["Diferencia_Gobierno"].apply(euros)
    tabla_display["Ticket Medio"]        = tabla["Ticket_Medio"].apply(euros)
    tabla_display["% del Total"]         = tabla["Pct_Total"].apply(lambda x: f"{x:.1f}%")

    st.dataframe(
        tabla_display,
        use_container_width=True,
        hide_index=True,
    )

    render_footer(df)


def page_vendedores(df: pd.DataFrame, df_filtered: pd.DataFrame):
    """
    Página 3 — Vendedores.
    · Tabla: operaciones, facturación, ticket medio, unidades, horas activas, fact/hora
    · Barras: facturación por vendedor
    · Barras: ticket medio por vendedor
    · Heatmap: Vendedor × Hora del día (nº operaciones)
    · Línea: evolución semanal de facturación por vendedor
    """
    st.title("👥 Vendedores")

    if df_filtered.empty:
        no_data_message()
        render_footer(df)
        return

    if "Vendedor" not in df_filtered.columns:
        st.warning("No se encontró la columna 'Vendedor' en los datos.")
        render_footer(df)
        return

    # ── Métricas base por vendedor ────────────────────────────────────────────
    grp = (
        df_filtered
        .groupby("Vendedor", as_index=False)
        .agg(Operaciones=("Facturación", "count"), Facturación=("Facturación", "sum"))
    )

    if "Cantidad (Unidades)" in df_filtered.columns:
        uni = df_filtered.groupby("Vendedor")["Cantidad (Unidades)"].sum().rename("Unidades")
        grp = grp.merge(uni, on="Vendedor", how="left")

    if "Hora_Int" in df_filtered.columns:
        hact = df_filtered.groupby("Vendedor")["Hora_Int"].nunique().rename("Horas_Activas")
        grp  = grp.merge(hact, on="Vendedor", how="left")
    grp["Horas_Activas"] = grp.get("Horas_Activas", 1).fillna(1).clip(lower=1)

    grp["Ticket_Medio"]  = grp["Facturación"] / grp["Operaciones"].replace(0, 1)
    grp["Fact_por_Hora"] = grp["Facturación"]  / grp["Horas_Activas"]
    grp = grp.sort_values("Facturación", ascending=False).reset_index(drop=True)

    # Paleta de colores por vendedor (consistente en todos los gráficos de esta página)
    n_vend      = len(grp)
    colores_v   = (_ORG_COLORS * ((n_vend // len(_ORG_COLORS)) + 1))[:n_vend]
    color_vmap  = dict(zip(grp["Vendedor"], colores_v))
    vend_order  = grp["Vendedor"].tolist()

    # ── Tabla de rendimiento ──────────────────────────────────────────────────
    st.subheader("📋 Rendimiento por Vendedor")

    tbl = pd.DataFrame()
    tbl["Vendedor"]     = grp["Vendedor"]
    tbl["Operaciones"]  = grp["Operaciones"].apply(lambda x: f"{x:,}")
    tbl["Facturación"]  = grp["Facturación"].apply(euros)
    tbl["Ticket Medio"] = grp["Ticket_Medio"].apply(euros)
    if "Unidades" in grp.columns:
        tbl["Unidades"] = grp["Unidades"].apply(lambda x: f"{x:,.0f}")
    tbl["Horas Activas"] = grp["Horas_Activas"].apply(lambda x: f"{int(x)}")
    tbl["Fact. / Hora"]  = grp["Fact_por_Hora"].apply(euros)

    st.dataframe(tbl, use_container_width=True, hide_index=True)

    st.markdown("---")

    # ── Barras: facturación + ticket medio (lado a lado) ─────────────────────
    col_l, col_r = st.columns(2)

    with col_l:
        st.subheader("💶 Facturación por Vendedor")
        fig_fv = go.Figure(go.Bar(
            x=vend_order,
            y=[grp.loc[grp["Vendedor"] == v, "Facturación"].values[0] for v in vend_order],
            marker_color=[color_vmap[v] for v in vend_order],
            text=[euros(grp.loc[grp["Vendedor"] == v, "Facturación"].values[0]) for v in vend_order],
            textposition="outside",
            textfont=dict(color="#ECF0F1", size=10),
            hovertemplate="<b>%{x}</b><br>%{y:,.2f} €<extra></extra>",
            cliponaxis=False,
        ))
        fig_fv.update_layout(
            **_plot_layout(height=360, margin_b=80),
            yaxis_title="Facturación (€)",
            xaxis=dict(tickangle=-30, tickfont=dict(size=10)),
            margin=dict(l=10, r=10, t=20, b=90),
        )
        st.plotly_chart(fig_fv, use_container_width=True)

    with col_r:
        st.subheader("🎫 Ticket Medio por Vendedor")
        fig_tv = go.Figure(go.Bar(
            x=vend_order,
            y=[grp.loc[grp["Vendedor"] == v, "Ticket_Medio"].values[0] for v in vend_order],
            marker_color=[color_vmap[v] for v in vend_order],
            text=[euros(grp.loc[grp["Vendedor"] == v, "Ticket_Medio"].values[0]) for v in vend_order],
            textposition="outside",
            textfont=dict(color="#ECF0F1", size=10),
            hovertemplate="<b>%{x}</b><br>%{y:,.2f} €<extra></extra>",
            cliponaxis=False,
        ))
        fig_tv.update_layout(
            **_plot_layout(height=360, margin_b=80),
            yaxis_title="Ticket medio (€)",
            xaxis=dict(tickangle=-30, tickfont=dict(size=10)),
            margin=dict(l=10, r=10, t=20, b=90),
        )
        st.plotly_chart(fig_tv, use_container_width=True)

    st.markdown("---")

    # ── Heatmap: Vendedor × Hora del día ─────────────────────────────────────
    st.subheader("🕐 Actividad Horaria por Vendedor (nº operaciones)")

    if "Hora_Int" in df_filtered.columns:
        import numpy as np

        # Pivot: índice=Vendedor, columnas=Hora 0–23
        pivot_h = (
            df_filtered
            .groupby(["Vendedor", "Hora_Int"])
            .size()
            .unstack(fill_value=0)
            .reindex(columns=range(24), fill_value=0)
            .reindex(vend_order)          # orden por facturación desc
        )

        fig_hm = go.Figure(go.Heatmap(
            z=pivot_h.values,
            x=[f"{h:02d}h" for h in range(24)],
            y=pivot_h.index.tolist(),
            colorscale=[
                [0.0,  "#16213e"],
                [0.01, "#0e4429"],
                [0.35, "#2ECC71"],
                [1.0,  "#F39C12"],
            ],
            showscale=True,
            colorbar=dict(
                title=dict(text="Ops.", font=dict(color="#ECF0F1")),
                tickfont=dict(color="#ECF0F1"),
            ),
            hovertemplate="<b>%{y}</b> · %{x}<br><b>%{z}</b> operaciones<extra></extra>",
        ))
        fig_hm.update_layout(
            **_plot_layout(height=max(280, n_vend * 46 + 80), margin_b=50),
            xaxis=dict(side="bottom", tickangle=0, tickfont=dict(size=9),
                       gridcolor="rgba(0,0,0,0)"),
            yaxis=dict(tickfont=dict(size=11), gridcolor="rgba(0,0,0,0)"),
            margin=dict(l=10, r=90, t=20, b=50),
        )
        st.plotly_chart(fig_hm, use_container_width=True)

    st.markdown("---")

    # ── Línea: evolución semanal de facturación por vendedor ─────────────────
    st.subheader("📅 Evolución Semanal de Facturación por Vendedor")

    if "Fecha_ES" in df_filtered.columns:
        df_sem = df_filtered.copy()
        # Inicio de semana (lunes) como etiqueta temporal
        df_sem["Semana"] = df_sem["Fecha_ES"] - pd.to_timedelta(
            df_sem["Fecha_ES"].dt.dayofweek, unit="D"
        )
        df_sem["Semana"] = df_sem["Semana"].dt.normalize()

        weekly = (
            df_sem
            .groupby(["Semana", "Vendedor"], as_index=False)["Facturación"]
            .sum()
        )

        fig_sem = go.Figure()
        for i, vend in enumerate(vend_order):
            dv = weekly[weekly["Vendedor"] == vend].sort_values("Semana")
            if dv.empty:
                continue
            fig_sem.add_trace(go.Scatter(
                x=dv["Semana"],
                y=dv["Facturación"],
                name=vend,
                mode="lines+markers",
                line=dict(width=2, color=colores_v[i]),
                marker=dict(size=5, color=colores_v[i]),
                hovertemplate=(
                    f"<b>{vend}</b><br>"
                    "%{x|%d/%m/%Y}<br>"
                    "%{y:,.2f} €<extra></extra>"
                ),
            ))

        fig_sem.update_layout(
            **_plot_layout(height=390),
            xaxis_title="Semana (inicio lunes)",
            yaxis_title="Facturación (€)",
            hovermode="x unified",
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02,
                xanchor="right", x=1, bgcolor="rgba(0,0,0,0)",
            ),
        )
        st.plotly_chart(fig_sem, use_container_width=True)

    render_footer(df)


def page_temporal(df: pd.DataFrame, df_filtered: pd.DataFrame):
    """
    Página 4 — Análisis Temporal.
    · Heatmap calendario estilo GitHub
    · Barras: operaciones por hora (0-23)
    · Barras: facturación por día de la semana
    · Línea+barras: facturación mensual acumulada
    · Top 10 días por facturación y por operaciones
    """
    st.title("📅 Análisis Temporal")

    if df_filtered.empty:
        no_data_message()
        render_footer(df)
        return

    if "Fecha_ES" not in df_filtered.columns:
        st.warning("No se encontró la columna 'Fecha_ES' en los datos.")
        render_footer(df)
        return

    import numpy as np

    # ── Tabla diaria base (usada en múltiples gráficos) ───────────────────────
    daily = (
        df_filtered
        .groupby(df_filtered["Fecha_ES"].dt.date, as_index=False)
        .agg(Facturación=("Facturación", "sum"), Operaciones=("Facturación", "count"))
    )
    daily.columns = ["Fecha", "Facturación", "Operaciones"]
    daily["Fecha"] = pd.to_datetime(daily["Fecha"])
    daily = daily.sort_values("Fecha").reset_index(drop=True)

    # ── Heatmap calendario estilo GitHub ─────────────────────────────────────
    st.subheader("📆 Heatmap Calendario — Facturación Diaria")

    min_d     = daily["Fecha"].min()
    max_d     = daily["Fecha"].max()
    all_dates = pd.date_range(min_d, max_d, freq="D")

    # DataFrame con todos los días (0 en los que no hay ventas)
    cal = pd.DataFrame({"Fecha": all_dates})
    cal = cal.merge(daily[["Fecha", "Facturación"]], on="Fecha", how="left").fillna(0)
    cal["dow"]        = cal["Fecha"].dt.dayofweek          # 0 = lunes
    cal["week_start"] = cal["Fecha"] - pd.to_timedelta(cal["dow"], unit="D")

    weeks_sorted = sorted(cal["week_start"].unique())
    week_map     = {w: i for i, w in enumerate(weeks_sorted)}
    cal["wi"]    = cal["week_start"].map(week_map)
    n_weeks      = len(weeks_sorted)

    # Matriz 7 (días) × N (semanas)
    z_cal     = np.zeros((7, n_weeks))
    hover_cal = [[""] * n_weeks for _ in range(7)]

    for _, row in cal.iterrows():
        d, w = int(row["dow"]), int(row["wi"])
        z_cal[d, w]     = row["Facturación"]
        hover_cal[d][w] = (
            f"{pd.Timestamp(row['Fecha']).strftime('%d/%m/%Y')}"
            f"<br>{euros(row['Facturación'])}"
        )

    # Etiquetas de mes en el eje X
    xtick_vals, xtick_text, seen_m = [], [], set()
    for w_date in weeks_sorted:
        ts    = pd.Timestamp(w_date)
        m_key = ts.strftime("%Y-%m")
        if m_key not in seen_m:
            seen_m.add(m_key)
            xtick_vals.append(week_map[w_date])
            xtick_text.append(ts.strftime("%b %Y"))

    _DIAS_CORTOS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]

    fig_cal = go.Figure(go.Heatmap(
        z=z_cal,
        x=list(range(n_weeks)),
        y=_DIAS_CORTOS,
        text=hover_cal,
        hovertemplate="%{text}<extra></extra>",
        colorscale=[
            [0.00, "#161b22"],
            [0.01, "#0e4429"],
            [0.25, "#006d32"],
            [0.60, "#26a641"],
            [1.00, "#39d353"],
        ],
        showscale=True,
        colorbar=dict(
            title=dict(text="€", font=dict(color="#ECF0F1")),
            tickfont=dict(color="#ECF0F1"),
            len=0.9,
        ),
        xgap=2,
        ygap=2,
    ))
    fig_cal.update_layout(
        **_plot_layout(height=235, margin_b=45),
        xaxis=dict(
            tickvals=xtick_vals,
            ticktext=xtick_text,
            tickfont=dict(size=10),
            gridcolor="rgba(0,0,0,0)",
            linecolor="rgba(0,0,0,0)",
        ),
        yaxis=dict(
            autorange="reversed",       # lunes arriba
            tickfont=dict(size=10),
            gridcolor="rgba(0,0,0,0)",
            linecolor="rgba(0,0,0,0)",
        ),
        margin=dict(l=45, r=90, t=20, b=45),
    )
    st.plotly_chart(fig_cal, use_container_width=True)

    st.markdown("---")

    # ── Barras: hora del día + día de la semana (lado a lado) ─────────────────
    _DIA_ES = {
        0: "Lunes", 1: "Martes", 2: "Miércoles", 3: "Jueves",
        4: "Viernes", 5: "Sábado", 6: "Domingo",
    }
    col_hora, col_dia = st.columns(2)

    # — Operaciones por hora —
    with col_hora:
        st.subheader("🕐 Operaciones por Hora del Día")

        if "Hora_Int" in df_filtered.columns:
            hora_grp = df_filtered.groupby("Hora_Int").size().rename("Operaciones").reset_index()
            hora_full = pd.DataFrame({"Hora_Int": range(24)})
            hora_grp  = hora_full.merge(hora_grp, on="Hora_Int", how="left").fillna(0)

            fig_hora = go.Figure(go.Bar(
                x=hora_grp["Hora_Int"].apply(lambda h: f"{int(h):02d}h"),
                y=hora_grp["Operaciones"],
                marker=dict(
                    color=hora_grp["Operaciones"],
                    colorscale=[[0, "#16213e"], [1, "#2ECC71"]],
                    showscale=False,
                ),
                hovertemplate="<b>%{x}</b><br><b>%{y:.0f}</b> operaciones<extra></extra>",
            ))
            fig_hora.update_layout(
                **_plot_layout(height=350, margin_b=40),
                xaxis_title="Hora",
                yaxis_title="Nº Operaciones",
                bargap=0.08,
            )
            st.plotly_chart(fig_hora, use_container_width=True)
        else:
            st.info("Columna 'Hora_Int' no disponible.")

    # — Facturación por día de la semana —
    with col_dia:
        st.subheader("📅 Facturación por Día de la Semana")

        df_dia = df_filtered.copy()
        df_dia["dow_num"]  = df_dia["Fecha_ES"].dt.dayofweek
        df_dia["dow_name"] = df_dia["dow_num"].map(_DIA_ES)

        dia_grp = (
            df_dia
            .groupby(["dow_num", "dow_name"], as_index=False)["Facturación"]
            .sum()
        )
        # Rellenar los 7 días aunque no haya ventas
        dias_base = pd.DataFrame({
            "dow_num":  range(7),
            "dow_name": [_DIA_ES[i] for i in range(7)],
        })
        dia_grp = dias_base.merge(dia_grp[["dow_num", "Facturación"]], on="dow_num", how="left").fillna(0)

        fig_dia = go.Figure(go.Bar(
            x=dia_grp["dow_name"],
            y=dia_grp["Facturación"],
            marker=dict(
                color=dia_grp["Facturación"],
                colorscale=[[0, "#16213e"], [1, "#3498DB"]],
                showscale=False,
            ),
            text=[euros(v) if v > 0 else "" for v in dia_grp["Facturación"]],
            textposition="outside",
            textfont=dict(color="#ECF0F1", size=9),
            hovertemplate="<b>%{x}</b><br>%{y:,.2f} €<extra></extra>",
            cliponaxis=False,
        ))
        fig_dia.update_layout(
            **_plot_layout(height=350, margin_b=40),
            yaxis_title="Facturación (€)",
            bargap=0.15,
            margin=dict(l=10, r=10, t=20, b=60),
        )
        st.plotly_chart(fig_dia, use_container_width=True)

    st.markdown("---")

    # ── Facturación mensual + acumulado (eje doble) ───────────────────────────
    st.subheader("📈 Facturación Mensual Acumulada")

    _MES_ABREV = {
        "January": "Ene", "February": "Feb", "March": "Mar",  "April":    "Abr",
        "May":     "May", "June":     "Jun", "July":  "Jul",  "August":   "Ago",
        "September":"Sep","October":  "Oct", "November":"Nov", "December": "Dic",
    }
    cols_m = {"Año", "Mes_Num", "Mes"}
    if cols_m.issubset(df_filtered.columns):
        monthly = (
            df_filtered
            .groupby(["Año", "Mes_Num", "Mes"], as_index=False)["Facturación"]
            .sum()
            .sort_values(["Año", "Mes_Num"])
        )
        monthly["Periodo"]   = (
            monthly["Mes"].map(_MES_ABREV).fillna(monthly["Mes"])
            + " " + monthly["Año"].astype(str)
        )
        monthly["Acumulado"] = monthly["Facturación"].cumsum()

        fig_cum = go.Figure()
        fig_cum.add_trace(go.Bar(
            x=monthly["Periodo"],
            y=monthly["Facturación"],
            name="Facturación mensual",
            marker_color="rgba(52,152,219,0.65)",
            yaxis="y",
            hovertemplate="<b>%{x}</b><br>Mensual: %{y:,.2f} €<extra></extra>",
        ))
        fig_cum.add_trace(go.Scatter(
            x=monthly["Periodo"],
            y=monthly["Acumulado"],
            name="Acumulado",
            mode="lines+markers",
            line=dict(color="#2ECC71", width=2.5),
            marker=dict(size=7, color="#2ECC71"),
            yaxis="y2",
            hovertemplate="<b>%{x}</b><br>Acumulado: %{y:,.2f} €<extra></extra>",
        ))
        fig_cum.update_layout(
            **_plot_layout(height=390),
            xaxis_title="Mes",
            yaxis=dict(
                title="Facturación mensual (€)",
                gridcolor="#2c2c4e",
            ),
            yaxis2=dict(
                title="Acumulado (€)",
                overlaying="y",
                side="right",
                gridcolor="rgba(0,0,0,0)",
                tickfont=dict(color="#2ECC71"),
                title_font=dict(color="#2ECC71"),
                showgrid=False,
            ),
            barmode="group",
            bargap=0.15,
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02,
                xanchor="right", x=1, bgcolor="rgba(0,0,0,0)",
            ),
            hovermode="x unified",
        )
        st.plotly_chart(fig_cum, use_container_width=True)

    st.markdown("---")

    # ── Top 10 días ───────────────────────────────────────────────────────────
    st.subheader("🏆 Top 10 Días")

    top10_fact = daily.nlargest(10, "Facturación").copy()
    top10_ops  = daily.nlargest(10, "Operaciones").copy()

    col_tf, col_to = st.columns(2)

    with col_tf:
        st.markdown("**Por Facturación**")
        t1 = pd.DataFrame()
        t1["Fecha"]       = top10_fact["Fecha"].dt.strftime("%d/%m/%Y")
        t1["Facturación"] = top10_fact["Facturación"].apply(euros)
        t1["Operaciones"] = top10_fact["Operaciones"].apply(lambda x: f"{int(x):,}")
        st.dataframe(t1, use_container_width=True, hide_index=True)

    with col_to:
        st.markdown("**Por Nº Operaciones**")
        t2 = pd.DataFrame()
        t2["Fecha"]       = top10_ops["Fecha"].dt.strftime("%d/%m/%Y")
        t2["Operaciones"] = top10_ops["Operaciones"].apply(lambda x: f"{int(x):,}")
        t2["Facturación"] = top10_ops["Facturación"].apply(euros)
        st.dataframe(t2, use_container_width=True, hide_index=True)

    render_footer(df)


def page_productos(df: pd.DataFrame, df_filtered: pd.DataFrame):
    """
    Página 5 — Productos.
    · Métricas % genérico vs marca (unidades e importe)
    · Tabla Top 50: código, denominación, tipo, organismo principal, unidades, facturación, precio medio
    · Barras horiz: Top 20 por facturación y Top 20 por unidades
    · Barras agrupadas: Top 10 genéricos vs Top 10 marca (por rank)
    """
    st.title("💊 Productos")

    if df_filtered.empty:
        no_data_message()
        render_footer(df)
        return

    if "Código" not in df_filtered.columns or "Denominación" not in df_filtered.columns:
        st.warning("No se encontraron las columnas 'Código' o 'Denominación' en los datos.")
        render_footer(df)
        return

    # Helper local para moda segura
    def _mode(s):
        m = s.dropna().mode()
        return m.iloc[0] if len(m) > 0 else "N/D"

    # ── Agregar por producto ──────────────────────────────────────────────────
    agg_d = {"Facturación": ("Facturación", "sum")}
    if "Cantidad (Unidades)" in df_filtered.columns:
        agg_d["Unidades"] = ("Cantidad (Unidades)", "sum")
    if "Pvp" in df_filtered.columns:
        agg_d["Precio_Medio"] = ("Pvp", "mean")

    prod = (
        df_filtered
        .groupby(["Código", "Denominación"], as_index=False)
        .agg(**agg_d)
        .sort_values("Facturación", ascending=False)
        .reset_index(drop=True)
    )

    # Organismo más frecuente por producto
    if "Organismo_Grupo" in df_filtered.columns:
        org_freq = (
            df_filtered.groupby("Código")["Organismo_Grupo"]
            .agg(_mode).rename("Org_Frecuente")
        )
        prod = prod.merge(org_freq, on="Código", how="left")
    else:
        prod["Org_Frecuente"] = "N/D"

    # Tipo: Genérico / Marca
    if "Es_Generico" in df_filtered.columns:
        gen_first = df_filtered.groupby("Código")["Es_Generico"].first().rename("Es_Generico")
        prod = prod.merge(gen_first, on="Código", how="left")
        prod["Tipo"] = prod["Es_Generico"].map({True: "Genérico", False: "Marca"}).fillna("N/D")
    else:
        prod["Es_Generico"] = False
        prod["Tipo"] = "N/D"

    # ── Métricas genérico vs marca ────────────────────────────────────────────
    st.subheader("💊 Genérico vs Marca — Resumen")

    total_units = prod["Unidades"].sum() if "Unidades" in prod.columns else 1
    total_fact  = prod["Facturación"].sum() or 1
    gen_units   = prod.loc[prod["Tipo"] == "Genérico", "Unidades"].sum() if "Unidades" in prod.columns else 0
    gen_fact    = prod.loc[prod["Tipo"] == "Genérico", "Facturación"].sum()

    pct_gu = gen_units / total_units * 100
    pct_gf = gen_fact  / total_fact  * 100

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("💊 Genérico — Unidades",  f"{pct_gu:.1f} %",
                  help="% de unidades dispensadas que son genéricos (código > 600.000)")
    with c2:
        st.metric("🏷️ Marca — Unidades",     f"{100 - pct_gu:.1f} %")
    with c3:
        st.metric("💶 Genérico — Importe",   f"{pct_gf:.1f} %",
                  help="% de facturación correspondiente a genéricos")
    with c4:
        st.metric("💰 Marca — Importe",      f"{100 - pct_gf:.1f} %")

    st.markdown("---")

    # ── Tabla Top 50 ──────────────────────────────────────────────────────────
    st.subheader("📋 Top 50 Productos por Facturación")

    top50 = prod.head(50).copy()
    tbl_p = pd.DataFrame()
    tbl_p["Código"]              = top50["Código"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "N/D")
    tbl_p["Denominación"]        = top50["Denominación"]
    tbl_p["Tipo"]                = top50["Tipo"]
    tbl_p["Organismo Principal"] = top50["Org_Frecuente"]
    if "Unidades" in top50.columns:
        tbl_p["Unidades"]        = top50["Unidades"].apply(lambda x: f"{x:,.0f}")
    tbl_p["Facturación"]         = top50["Facturación"].apply(euros)
    if "Precio_Medio" in top50.columns:
        tbl_p["Precio Medio"]    = top50["Precio_Medio"].apply(euros)

    st.dataframe(tbl_p, use_container_width=True, hide_index=True)

    st.markdown("---")

    # ── Barras Top 20 por facturación + por unidades (lado a lado) ───────────
    col_bf, col_bu = st.columns(2)

    # Verde = Genérico, Azul = Marca
    def _tipo_color(tipo):
        return "#2ECC71" if tipo == "Genérico" else "#3498DB"

    with col_bf:
        st.subheader("💶 Top 20 por Facturación")
        t20f = prod.head(20).copy()
        t20f["Nombre"] = t20f["Denominación"].str.slice(0, 30)
        fig_tf = go.Figure()
        fig_tf.add_trace(go.Bar(
            x=t20f["Facturación"],
            y=t20f["Nombre"],
            orientation="h",
            marker_color=[_tipo_color(t) for t in t20f["Tipo"]],
            text=[euros(v) for v in t20f["Facturación"]],
            textposition="outside",
            textfont=dict(color="#ECF0F1", size=9),
            hovertemplate="<b>%{y}</b><br>%{x:,.2f} €<extra></extra>",
            cliponaxis=False,
        ))
        # Leyenda manual de tipo
        for tipo, col in [("Genérico", "#2ECC71"), ("Marca", "#3498DB")]:
            fig_tf.add_trace(go.Bar(
                x=[None], y=[None], name=tipo,
                marker_color=col, showlegend=True,
            ))
        fig_tf.update_layout(
            **_plot_layout(height=530, margin_b=30),
            xaxis_title="Facturación (€)",
            yaxis=dict(autorange="reversed", tickfont=dict(size=9)),
            legend=dict(orientation="h", yanchor="bottom", y=1.01,
                        xanchor="right", x=1, bgcolor="rgba(0,0,0,0)"),
            margin=dict(l=10, r=115, t=30, b=30),
        )
        st.plotly_chart(fig_tf, use_container_width=True)

    with col_bu:
        st.subheader("📦 Top 20 por Unidades")
        if "Unidades" in prod.columns:
            t20u = prod.nlargest(20, "Unidades").reset_index(drop=True)
            t20u["Nombre"] = t20u["Denominación"].str.slice(0, 30)
            fig_tu = go.Figure(go.Bar(
                x=t20u["Unidades"],
                y=t20u["Nombre"],
                orientation="h",
                marker_color=[_tipo_color(t) for t in t20u["Tipo"]],
                text=[f"{int(v):,}" for v in t20u["Unidades"]],
                textposition="outside",
                textfont=dict(color="#ECF0F1", size=9),
                hovertemplate="<b>%{y}</b><br>%{x:,.0f} uds<extra></extra>",
                cliponaxis=False,
            ))
            fig_tu.update_layout(
                **_plot_layout(height=530, margin_b=30),
                xaxis_title="Unidades",
                yaxis=dict(autorange="reversed", tickfont=dict(size=9)),
                margin=dict(l=10, r=80, t=30, b=30),
            )
            st.plotly_chart(fig_tu, use_container_width=True)
        else:
            st.info("Columna 'Cantidad (Unidades)' no disponible.")

    st.markdown("---")

    # ── Barras agrupadas: Top 10 genéricos vs Top 10 marca (por rank) ─────────
    st.subheader("🔬 Top 10 Genéricos vs Top 10 Marca")

    top10g = prod[prod["Tipo"] == "Genérico"].head(10).reset_index(drop=True)
    top10m = prod[prod["Tipo"] == "Marca"].head(10).reset_index(drop=True)

    if top10g.empty and top10m.empty:
        st.info("No hay datos suficientes para la comparativa genérico / marca.")
    else:
        n_ranks = max(len(top10g), len(top10m))
        ranks   = [f"#{i + 1}" for i in range(n_ranks)]

        fig_gm = go.Figure()
        if not top10g.empty:
            fig_gm.add_trace(go.Bar(
                name="Genérico",
                x=ranks[: len(top10g)],
                y=top10g["Facturación"],
                customdata=top10g["Denominación"].str.slice(0, 30),
                marker_color="#2ECC71",
                text=[euros(v) for v in top10g["Facturación"]],
                textposition="outside",
                textfont=dict(color="#ECF0F1", size=9),
                hovertemplate="<b>Genérico %{x}</b><br>%{customdata}<br>%{y:,.2f} €<extra></extra>",
                cliponaxis=False,
            ))
        if not top10m.empty:
            fig_gm.add_trace(go.Bar(
                name="Marca",
                x=ranks[: len(top10m)],
                y=top10m["Facturación"],
                customdata=top10m["Denominación"].str.slice(0, 30),
                marker_color="#3498DB",
                text=[euros(v) for v in top10m["Facturación"]],
                textposition="outside",
                textfont=dict(color="#ECF0F1", size=9),
                hovertemplate="<b>Marca %{x}</b><br>%{customdata}<br>%{y:,.2f} €<extra></extra>",
                cliponaxis=False,
            ))

        fig_gm.update_layout(
            **_plot_layout(height=420, margin_b=50),
            barmode="group",
            xaxis_title="Posición en ranking (por facturación dentro de cada tipo)",
            yaxis_title="Facturación (€)",
            bargap=0.20,
            bargroupgap=0.05,
            legend=dict(orientation="h", yanchor="bottom", y=1.02,
                        xanchor="right", x=1, bgcolor="rgba(0,0,0,0)"),
        )
        st.plotly_chart(fig_gm, use_container_width=True)

    render_footer(df)


def page_clientes(df: pd.DataFrame, df_filtered: pd.DataFrame):
    """Página 6 — Análisis de Clientes."""
    st.title("👤 Clientes")

    if df_filtered.empty:
        no_data_message()
        render_footer(df)
        return

    if "Cliente" not in df_filtered.columns:
        st.warning("La columna 'Cliente' no está disponible en los datos.")
        render_footer(df)
        return

    df_cli = df_filtered[
        df_filtered["Cliente"].notna() &
        (df_filtered["Cliente"].astype(str).str.strip() != "")
    ].copy()

    if df_cli.empty:
        st.warning("No hay datos de clientes en el rango seleccionado.")
        render_footer(df)
        return

    # ── Agregación por cliente ────────────────────────────────────────────
    cli_agg = (
        df_cli.groupby("Cliente")
        .agg(
            Operaciones=("Facturación", "count"),
            Facturación=("Facturación", "sum"),
            Último_Día=("Fecha_ES", "max"),
        )
        .reset_index()
    )
    cli_agg["Ticket_Medio"] = cli_agg["Facturación"] / cli_agg["Operaciones"].replace(0, 1)

    if "Organismo_Grupo" in df_cli.columns:
        org_principal = (
            df_cli.groupby(["Cliente", "Organismo_Grupo"])
            .size()
            .reset_index(name="cnt")
            .sort_values("cnt", ascending=False)
            .drop_duplicates("Cliente")
            .set_index("Cliente")["Organismo_Grupo"]
        )
        cli_agg["Organismo_Principal"] = cli_agg["Cliente"].map(org_principal).fillna("-")
    else:
        cli_agg["Organismo_Principal"] = "-"

    cli_agg = cli_agg.sort_values("Facturación", ascending=False).reset_index(drop=True)

    total_facturacion = cli_agg["Facturación"].sum()
    n_clientes = len(cli_agg)

    # ── Regla del 80/20 ──────────────────────────────────────────────────
    cli_agg["Facturación_Acum"] = cli_agg["Facturación"].cumsum()
    cli_agg["Pct_Acum"] = (
        cli_agg["Facturación_Acum"] / total_facturacion
        if total_facturacion > 0
        else 0.0
    )
    n_80 = int((cli_agg["Pct_Acum"] < 0.80).sum()) + 1
    n_80 = min(n_80, n_clientes)
    pct_clientes_80 = n_80 / n_clientes * 100 if n_clientes > 0 else 0.0

    # ── KPI cards ─────────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Clientes únicos", f"{n_clientes:,}")
    c2.metric("Facturación total", euros(total_facturacion))
    c3.metric(
        "Ticket medio global",
        euros(total_facturacion / max(int(cli_agg["Operaciones"].sum()), 1)),
    )
    c4.metric(
        "Clientes → 80% facturación",
        str(n_80),
        delta=f"{pct_clientes_80:.1f}% del total",
        delta_color="off",
    )

    # ── Top 50 tabla ──────────────────────────────────────────────────────
    st.markdown("### 📋 Top 50 Clientes por Facturación")
    top50 = cli_agg.head(50).copy()
    top50.index = range(1, len(top50) + 1)

    st.dataframe(
        top50.assign(
            **{
                "Facturación €": top50["Facturación"].apply(euros),
                "Ticket Medio €": top50["Ticket_Medio"].apply(euros),
                "Último Día": top50["Último_Día"].dt.strftime("%d/%m/%Y"),
                "% Total": (top50["Facturación"] / total_facturacion * 100).map(
                    "{:.1f}%".format
                ),
            }
        )[[
            "Cliente", "Operaciones", "Facturación €",
            "Ticket Medio €", "Último Día", "Organismo_Principal", "% Total",
        ]].rename(columns={"Organismo_Principal": "Organismo Principal"}),
        use_container_width=True,
        height=380,
    )

    # ── Top 20 barras horizontales ────────────────────────────────────────
    st.markdown("### 📊 Top 20 Clientes — Facturación")
    top20 = cli_agg.head(20).copy()
    labels = (
        top20["Cliente"].astype(str)
        .apply(lambda x: x if len(x) <= 14 else x[:13] + "…")
        .values
    )
    fig_bar = go.Figure(go.Bar(
        x=top20["Facturación"].values[::-1],
        y=labels[::-1],
        orientation="h",
        marker_color=_ORG_COLORS[1],
        text=[euros(v) for v in top20["Facturación"].values[::-1]],
        textposition="outside",
        hovertemplate="<b>%{y}</b><br>Facturación: %{x:,.2f} €<extra></extra>",
    ))
    fig_bar.update_layout(
        **_plot_layout("Top 20 Clientes por Facturación", height=540, margin_b=50),
        xaxis_title="Facturación (€)",
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    # ── Curva de Pareto ───────────────────────────────────────────────────
    st.markdown("### 📈 Curva de Pareto — Concentración de Facturación")
    pareto = cli_agg.copy()
    pareto["Pct_Clientes"] = (pareto.index + 1) / n_clientes * 100

    fig_pareto = go.Figure()
    fig_pareto.add_trace(go.Scatter(
        x=pareto["Pct_Clientes"],
        y=pareto["Pct_Acum"] * 100,
        mode="lines",
        name="% Facturación acum.",
        line=dict(color=_ORG_COLORS[0], width=2.5),
        fill="tozeroy",
        fillcolor="rgba(46,204,113,0.10)",
        hovertemplate="Top %{x:.1f}% clientes → %{y:.1f}% facturación<extra></extra>",
    ))
    fig_pareto.add_hline(
        y=80, line_dash="dash", line_color="#E74C3C", line_width=1.5,
        annotation_text="80% Facturación", annotation_position="right",
    )
    fig_pareto.add_vline(
        x=pct_clientes_80, line_dash="dash", line_color="#F39C12", line_width=1.5,
        annotation_text=f"{pct_clientes_80:.1f}% clientes",
        annotation_position="top left",
    )
    layout_pareto = _plot_layout("Curva de Pareto — Clientes vs Facturación", height=430)
    layout_pareto["xaxis"] = dict(
        title="% Clientes acumulado", gridcolor="#2c2c4e",
        linecolor="#2c2c4e", ticksuffix="%",
    )
    layout_pareto["yaxis"] = dict(
        title="% Facturación acumulada", gridcolor="#2c2c4e",
        linecolor="#2c2c4e", ticksuffix="%",
    )
    layout_pareto["showlegend"] = False
    fig_pareto.update_layout(**layout_pareto)
    st.plotly_chart(fig_pareto, use_container_width=True)

    render_footer(df)


def page_credito(df: pd.DataFrame, df_filtered: pd.DataFrame):
    """Página 7 — Crédito y Descuentos."""
    st.title("💳 Crédito y Descuentos")

    if df_filtered.empty:
        no_data_message()
        render_footer(df)
        return

    _CRED_KW = ("CRED", "CRÉD", "A CUENTA", "PENDIENTE", "FIADO", "CRÉDITO", "CREDITO")

    # ── Separar crédito / contado ─────────────────────────────────────────
    if "Tipo de Operación" in df_filtered.columns:
        mask_cred = df_filtered["Tipo de Operación"].astype(str).str.upper().apply(
            lambda x: any(k in x for k in _CRED_KW)
        )
    else:
        mask_cred = pd.Series(False, index=df_filtered.index)

    df_cred    = df_filtered[mask_cred].copy()
    df_contado = df_filtered[~mask_cred].copy()

    total_contado = df_contado["Facturación"].sum()
    total_credito = df_cred["Facturación"].sum()
    total_general = total_contado + total_credito
    pct_credito   = total_credito / total_general * 100 if total_general > 0 else 0.0

    # ── KPI cards ─────────────────────────────────────────────────────────
    st.markdown("### 💰 Resumen Crédito vs Contado")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Ventas",      euros(total_general))
    c2.metric("Ventas Contado",    euros(total_contado))
    c3.metric("Ventas Crédito",    euros(total_credito))
    c4.metric("% Crédito / Total", f"{pct_credito:.1f}%")

    # ── Tabla crédito por cliente ─────────────────────────────────────────
    st.markdown("### 📋 Crédito por Cliente")
    if not df_cred.empty and "Cliente" in df_cred.columns:
        cred_cli = (
            df_cred.groupby("Cliente")
            .agg(
                Operaciones=("Facturación", "count"),
                Facturación=("Facturación", "sum"),
                Último_Día=("Fecha_ES", "max"),
            )
            .reset_index()
            .sort_values("Facturación", ascending=False)
        )
        cred_cli["Ticket_Medio"] = (
            cred_cli["Facturación"] / cred_cli["Operaciones"].replace(0, 1)
        )
        cred_cli.index = range(1, len(cred_cli) + 1)
        st.dataframe(
            cred_cli.assign(
                **{
                    "Facturación €": cred_cli["Facturación"].apply(euros),
                    "Ticket Medio €": cred_cli["Ticket_Medio"].apply(euros),
                    "Último Día": cred_cli["Último_Día"].dt.strftime("%d/%m/%Y"),
                }
            )[["Cliente", "Operaciones", "Facturación €", "Ticket Medio €", "Último Día"]],
            use_container_width=True,
            height=300,
        )
    else:
        st.info("ℹ️ No se detectaron operaciones a crédito con los filtros actuales.")

    # ── Evolución mensual crédito vs contado ──────────────────────────────
    st.markdown("### 📅 Evolución Mensual — Crédito vs Contado")
    if "Mes_Num" in df_filtered.columns and "Año" in df_filtered.columns:
        df_m = df_filtered.copy()
        df_m["_Periodo"] = (
            df_m["Año"].astype(str) + "-" +
            df_m["Mes_Num"].astype(str).str.zfill(2)
        )
        if "Tipo de Operación" in df_m.columns:
            df_m["_EsCred"] = df_m["Tipo de Operación"].astype(str).str.upper().apply(
                lambda x: any(k in x for k in _CRED_KW)
            )
        else:
            df_m["_EsCred"] = False

        all_periods  = sorted(df_m["_Periodo"].unique())
        men_contado  = df_m[~df_m["_EsCred"]].groupby("_Periodo")["Facturación"].sum()
        men_credito  = df_m[df_m["_EsCred"]].groupby("_Periodo")["Facturación"].sum()

        if len(all_periods) > 0:
            fig_evol = go.Figure()
            fig_evol.add_trace(go.Bar(
                name="Contado",
                x=all_periods,
                y=[men_contado.get(p, 0) for p in all_periods],
                marker_color=_ORG_COLORS[0],
                hovertemplate="%{x}<br>Contado: %{y:,.2f} €<extra></extra>",
            ))
            fig_evol.add_trace(go.Bar(
                name="Crédito",
                x=all_periods,
                y=[men_credito.get(p, 0) for p in all_periods],
                marker_color=_ORG_COLORS[2],
                hovertemplate="%{x}<br>Crédito: %{y:,.2f} €<extra></extra>",
            ))
            evol_layout = _plot_layout(
                "Evolución mensual Contado vs Crédito", height=400
            )
            evol_layout["barmode"]    = "stack"
            evol_layout["xaxis_title"] = "Período"
            evol_layout["yaxis_title"] = "Facturación (€)"
            evol_layout["legend"]      = dict(
                orientation="h", y=1.08, x=0.5, xanchor="center"
            )
            fig_evol.update_layout(**evol_layout)
            st.plotly_chart(fig_evol, use_container_width=True)

    # ── Análisis de descuentos ────────────────────────────────────────────
    st.markdown("### 🏷️ Análisis de Descuentos")

    if "Descuento" not in df_filtered.columns:
        st.info("ℹ️ La columna 'Descuento' no está disponible en los datos.")
        render_footer(df)
        return

    total_dto     = df_filtered["Descuento"].sum()
    n_ops_dto     = int((df_filtered["Descuento"] > 0).sum())
    pct_ops_dto   = n_ops_dto / len(df_filtered) * 100 if len(df_filtered) > 0 else 0.0
    df_desc       = df_filtered[df_filtered["Descuento"] > 0]
    dto_medio     = df_desc["Descuento"].mean() if not df_desc.empty else 0.0

    d1, d2, d3 = st.columns(3)
    d1.metric("Descuento total aplicado",     euros(total_dto))
    d2.metric("Operaciones con descuento",    f"{n_ops_dto:,}  ({pct_ops_dto:.1f}%)")
    d3.metric("Descuento medio / operación",  euros(dto_medio))

    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("#### Top 10 Productos con Mayor Descuento")
        if "Denominación" in df_filtered.columns and not df_desc.empty:
            top_prod_dto = (
                df_desc.groupby("Denominación")["Descuento"]
                .sum()
                .nlargest(10)
                .reset_index()
            )
            fig_dp = go.Figure(go.Bar(
                x=top_prod_dto["Descuento"].values[::-1],
                y=top_prod_dto["Denominación"].values[::-1],
                orientation="h",
                marker_color=_ORG_COLORS[3],
                text=[euros(v) for v in top_prod_dto["Descuento"].values[::-1]],
                textposition="outside",
                hovertemplate="<b>%{y}</b><br>Descuento: %{x:,.2f} €<extra></extra>",
            ))
            fig_dp.update_layout(
                **_plot_layout("", height=380, margin_b=40),
                xaxis_title="Descuento total (€)",
            )
            st.plotly_chart(fig_dp, use_container_width=True)
        else:
            st.info("Sin datos de descuento por producto.")

    with col_b:
        st.markdown("#### Descuento por Organismo")
        if "Organismo_Grupo" in df_filtered.columns:
            dto_org = (
                df_filtered.groupby("Organismo_Grupo")["Descuento"]
                .sum()
                .sort_values(ascending=False)
                .reset_index()
            )
            colors_org = [
                _ORG_COLORS[i % len(_ORG_COLORS)] for i in range(len(dto_org))
            ]
            fig_do = go.Figure(go.Bar(
                x=dto_org["Descuento"].values[::-1],
                y=dto_org["Organismo_Grupo"].values[::-1],
                orientation="h",
                marker_color=colors_org[::-1],
                text=[euros(v) for v in dto_org["Descuento"].values[::-1]],
                textposition="outside",
                hovertemplate="<b>%{y}</b><br>Descuento: %{x:,.2f} €<extra></extra>",
            ))
            fig_do.update_layout(
                **_plot_layout("", height=380, margin_b=40),
                xaxis_title="Descuento total (€)",
            )
            st.plotly_chart(fig_do, use_container_width=True)
        else:
            st.info("Sin datos de organismo.")

    # ── Tabla resumen descuentos por organismo ────────────────────────────
    st.markdown("#### Resumen de Descuentos por Organismo")
    if "Organismo_Grupo" in df_filtered.columns:
        dto_tabla = (
            df_filtered.groupby("Organismo_Grupo")
            .agg(
                Operaciones=("Descuento", "count"),
                Ops_con_Dto=("Descuento", lambda x: int((x > 0).sum())),
                Descuento_Total=("Descuento", "sum"),
                Descuento_Medio=("Descuento", lambda x: x[x > 0].mean() if (x > 0).any() else 0.0),
            )
            .reset_index()
            .sort_values("Descuento_Total", ascending=False)
        )
        dto_tabla.index = range(1, len(dto_tabla) + 1)
        st.dataframe(
            dto_tabla.assign(
                **{
                    "Descuento Total €": dto_tabla["Descuento_Total"].apply(euros),
                    "Descuento Medio €": dto_tabla["Descuento_Medio"].apply(euros),
                    "% Ops c/ Dto": (
                        dto_tabla["Ops_con_Dto"]
                        / dto_tabla["Operaciones"].replace(0, 1)
                        * 100
                    ).map("{:.1f}%".format),
                }
            )[[
                "Organismo_Grupo", "Operaciones", "Ops_con_Dto",
                "% Ops c/ Dto", "Descuento Total €", "Descuento Medio €",
            ]].rename(columns={"Organismo_Grupo": "Organismo", "Ops_con_Dto": "Ops con Dto"}),
            use_container_width=True,
            height=280,
        )

    render_footer(df)


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — FILTROS GLOBALES + NAVEGACIÓN
# ─────────────────────────────────────────────────────────────────────────────

def render_sidebar(df: pd.DataFrame, authenticator):
    """
    Renderiza el sidebar con:
      · Información del usuario + botón de logout
      · Menú de navegación entre páginas
      · Filtros globales (fechas, hora, organismo, operación, vendedor, producto)
      · Contador de registros mostrados vs totales

    Args:
        df:             DataFrame completo (sin filtrar)
        authenticator:  Objeto Authenticate de streamlit-authenticator

    Returns:
        tuple: (df_filtered, pagina_seleccionada)
    """
    with st.sidebar:

        # ── Logo y usuario ────────────────────────────────────────────────
        user_name = st.session_state.get("name", "Usuario")
        st.markdown(
            f"""
            <div style="text-align:center; padding:1.2rem 0 0.8rem 0;">
                <h2 style="color:#2ECC71; margin:0; font-size:1.6rem;">💊 FarmaAnalytics</h2>
                <p style="color:#BDC3C7; font-size:0.85rem; margin:0.4rem 0 0 0;">
                    👤 <strong>{user_name}</strong>
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # ── Botón de logout ───────────────────────────────────────────────
        try:
            authenticator.logout(button_name="🚪 Cerrar sesión", location="sidebar")
        except TypeError:
            authenticator.logout("🚪 Cerrar sesión", "sidebar")

        st.markdown("---")

        # ── Navegación ────────────────────────────────────────────────────
        st.markdown("### 📋 Navegación")
        pagina = st.radio(
            label="Página:",
            options=[
                "📊 Resumen Ejecutivo",
                "🏥 Organismos",
                "👥 Vendedores",
                "📅 Análisis Temporal",
                "💊 Productos",
                "👤 Clientes",
                "💳 Crédito y Descuentos",
            ],
            label_visibility="collapsed",
        )

        st.markdown("---")

        # ── Filtros globales ──────────────────────────────────────────────
        st.markdown("### 🔍 Filtros")

        # Rango de fechas
        if "Fecha_ES" in df.columns and not df["Fecha_ES"].isna().all():
            min_date = df["Fecha_ES"].min().date()
            max_date = df["Fecha_ES"].max().date()
        else:
            min_date = datetime.today().date() - timedelta(days=365)
            max_date = datetime.today().date()

        date_start = st.date_input(
            "📅 Fecha inicio", value=min_date,
            min_value=min_date, max_value=max_date,
        )
        date_end = st.date_input(
            "📅 Fecha fin", value=max_date,
            min_value=min_date, max_value=max_date,
        )

        # Rango horario
        hour_range = st.slider(
            "🕐 Rango horario (h)",
            min_value=0, max_value=23, value=(0, 23),
        )

        # Organismo_Grupo
        sel_organismos = []
        if "Organismo_Grupo" in df.columns:
            all_organismos = sorted(df["Organismo_Grupo"].dropna().unique().tolist())
            sel_organismos = st.multiselect(
                "🏥 Organismo", options=all_organismos, default=[]
            )

        # Tipo de Operación
        sel_tipos = []
        if "Tipo de Operación" in df.columns:
            all_tipos = sorted(df["Tipo de Operación"].dropna().unique().tolist())
            sel_tipos = st.multiselect(
                "📋 Tipo Operación", options=all_tipos, default=[]
            )

        # Vendedor
        sel_vendedores = []
        if "Vendedor" in df.columns:
            all_vendedores = sorted(df["Vendedor"].dropna().unique().tolist())
            sel_vendedores = st.multiselect(
                "👤 Vendedor", options=all_vendedores, default=[]
            )

        # Genérico vs Marca
        product_type = st.radio(
            "💊 Tipo producto",
            options=["Todos", "Solo Genéricos", "Solo Marca"],
            index=0,
        )

        # Botón limpiar filtros — recarga la página (reset de widgets)
        if st.button("🧹 Limpiar filtros", use_container_width=True):
            st.rerun()

        st.markdown("---")

    # ── Aplicar filtros ───────────────────────────────────────────────────
    df_filtered = apply_filters(
        df,
        date_start   = date_start,
        date_end     = date_end,
        hour_min     = hour_range[0],
        hour_max     = hour_range[1],
        organismos   = sel_organismos if sel_organismos else None,
        tipos_op     = sel_tipos      if sel_tipos      else None,
        vendedores   = sel_vendedores if sel_vendedores else None,
        product_type = product_type,
    )

    # Contador de registros en el sidebar
    with st.sidebar:
        total   = len(df)
        visible = len(df_filtered)
        pct     = (visible / total * 100) if total > 0 else 0.0
        color   = "#2ECC71" if visible > 0 else "#E74C3C"
        st.markdown(
            f'<p style="color:{color}; font-size:0.85rem;">'
            f'📌 Mostrando <strong>{visible:,}</strong> de '
            f'<strong>{total:,}</strong> registros ({pct:.1f}%)</p>',
            unsafe_allow_html=True,
        )

    return df_filtered, pagina


# ─────────────────────────────────────────────────────────────────────────────
# PANTALLA DE BIENVENIDA (sin archivo cargado)
# ─────────────────────────────────────────────────────────────────────────────

def render_welcome_screen():
    """
    Muestra la pantalla de bienvenida con instrucciones y el uploader de Excel.

    Returns:
        UploadedFile | None: El archivo subido por el usuario, o None.
    """
    # ── Cabecera ──────────────────────────────────────────────────────────
    st.markdown(
        """
        <div style="text-align:center; padding:2rem 0 1.5rem 0;">
            <h1 style="color:#2ECC71; font-size:2.8rem;">💊 FarmaAnalytics</h1>
            <p style="color:#BDC3C7; font-size:1.1rem;">
                Panel profesional de análisis de ventas — bienvenido,
                <strong style="color:#3498DB;">{name}</strong>
            </p>
        </div>
        """.format(name=st.session_state.get("name", "usuario")),
        unsafe_allow_html=True,
    )

    # ── Card central con uploader ─────────────────────────────────────────
    col_l, col_c, col_r = st.columns([1, 2.5, 1])
    with col_c:
        st.markdown(
            """
            <div style="background:#16213e; border:2px solid #2ECC71;
                        border-radius:14px; padding:2rem 2.5rem; margin-bottom:1.5rem;">
                <h3 style="color:#2ECC71; text-align:center; margin-top:0;">
                    📂 Carga tu archivo de ventas
                </h3>
                <p style="color:#BDC3C7; text-align:center; margin-bottom:1.5rem;">
                    Sube el Excel exportado desde tu sistema de gestión de farmacia.<br>
                    Los datos se procesan <strong>en memoria</strong> y
                    <strong>no se guardan en ningún servidor</strong>.
                </p>
                <hr style="border-color:#2ECC71; opacity:0.3;">
                <p style="color:#95A5A6; font-size:0.85rem; margin:1rem 0 0 0;">
                    <strong>📋 Columnas requeridas en el Excel:</strong><br>
                    Fecha · Fecha_ES · Hora · Tipo de Operación · Empresa · Código<br>
                    Denominación · Organismo · Cantidad (Unidades) · Pvp<br>
                    PVP Facturado · Importe Bruto · Descuento · Importe Neto<br>
                    Cliente · Vendedor
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        uploaded_file = st.file_uploader(
            "Sube tu archivo de ventas (.xlsx)",
            type=["xlsx"],
            help="Formato: Excel (.xlsx). Tamaño máximo: 200 MB.",
        )

        st.markdown(
            """
            <div style="background:#0f3460; border-radius:8px;
                        padding:0.8rem 1rem; margin-top:1rem;">
                <p style="color:#BDC3C7; font-size:0.82rem; margin:0;">
                    ℹ️ <strong>Privacidad:</strong> tu archivo nunca sale de tu sesión de navegador.
                    Al cerrar sesión o recargar la página, los datos se eliminan automáticamente.
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    return uploaded_file


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIÓN PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

def main():
    """
    Punto de entrada de la aplicación.

    Flujo:
      1. Crear authenticator y renderizar login
      2. Si no autenticado → mostrar pantalla de login y parar
      3. Si autenticado sin archivo → mostrar welcome screen y uploader
      4. Si autenticado con archivo → mostrar dashboard completo
    """
    # ── 1. Autenticación ──────────────────────────────────────────────────
    try:
        authenticator, config = create_authenticator()
    except FileNotFoundError as e:
        st.error(f"⚠️ Error de configuración: {e}")
        st.stop()

    name, auth_status, username = render_login_ui(authenticator)

    # ── 2. Gestión del estado de autenticación ────────────────────────────
    if auth_status is False:
        st.error("❌ Usuario o contraseña incorrectos. Inténtalo de nuevo.")
        st.stop()

    if auth_status is None:
        # El formulario ya se renderizó en render_login_ui; sólo detenemos
        st.stop()

    # Usuario autenticado correctamente
    user_role = get_user_role(config, username)

    # ── 3. Sin archivo → pantalla de bienvenida ───────────────────────────
    if st.session_state.get("df") is None:

        # Logout visible en la esquina superior derecha
        col_space, col_logout = st.columns([8, 1])
        with col_logout:
            try:
                authenticator.logout(button_name="🚪 Salir", location="main")
            except TypeError:
                authenticator.logout("🚪 Salir", "main")

        uploaded_file = render_welcome_screen()

        if uploaded_file is not None:
            with st.spinner("⏳ Procesando tu archivo de ventas…"):
                try:
                    file_bytes = uploaded_file.read()
                    df = process_excel(file_bytes)
                    st.session_state["df"] = df
                    st.success(
                        f"✅ Archivo cargado: **{len(df):,} registros** procesados correctamente."
                    )
                    st.balloons()
                    st.rerun()
                except ValueError as e:
                    st.error(f"❌ Error en el formato del archivo: {e}")
                    st.info(
                        "Verifica que el Excel tiene el formato correcto "
                        "con todas las columnas requeridas."
                    )
        return

    # ── 4. Con archivo → dashboard completo ──────────────────────────────
    df: pd.DataFrame = st.session_state["df"]

    # Sidebar: filtros y navegación
    df_filtered, pagina = render_sidebar(df, authenticator)

    # Botón para cambiar archivo (en el sidebar)
    with st.sidebar:
        st.markdown("---")
        if st.button("📂 Cambiar archivo", use_container_width=True):
            # Limpiar datos y volver a welcome screen
            st.session_state.pop("df", None)
            # Limpiar cache del procesador para liberar memoria
            process_excel.clear()
            st.rerun()

    # ── Router de páginas ─────────────────────────────────────────────────
    page_router = {
        "📊 Resumen Ejecutivo"    : page_resumen,
        "🏥 Organismos"           : page_organismos,
        "👥 Vendedores"           : page_vendedores,
        "📅 Análisis Temporal"    : page_temporal,
        "💊 Productos"            : page_productos,
        "👤 Clientes"             : page_clientes,
        "💳 Crédito y Descuentos" : page_credito,
    }

    page_fn = page_router.get(pagina, page_resumen)
    page_fn(df, df_filtered)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    main()
