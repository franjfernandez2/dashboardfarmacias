"""
data_processor.py — Procesamiento del Excel de ventas
======================================================
Carga el archivo Excel, limpia los datos y genera todas las columnas
calculadas que necesita el dashboard.

Todas las funciones "pesadas" usan @st.cache_data para que Streamlit
no reprocese el mismo archivo en cada interacción del usuario.
"""

import io
from typing import Optional

import numpy as np
import pandas as pd
import streamlit as st


# ─────────────────────────────────────────────────────────────────────────────
# Procesamiento principal del Excel
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner="⏳ Procesando archivo Excel…")
def process_excel(file_bytes: bytes) -> pd.DataFrame:
    """
    Carga y procesa el Excel de ventas de farmacia.

    Columnas esperadas en el archivo:
        Fecha, Fecha_ES, Hora, Tipo de Operación, Empresa, Código,
        Denominación, Organismo, Cantidad (Unidades), Pvp, PVP Facturado,
        Importe Bruto, Descuento, Importe Neto, Cliente, Vendedor,
        Existencias Anteriores, Existencias Posteriores → usadas para stock y roturas

    Columnas calculadas que añade esta función:
        Facturación         = Cantidad (Unidades) × Pvp
        Diferencia_Gobierno = PVP Facturado - Importe Neto
        Es_Generico         = True si Código > 600 000
        Hora_Int            = Hora como entero (0-23)
        Dia_Semana          = nombre del día de la semana (de Fecha_ES)
        Mes                 = nombre del mes (de Fecha_ES)
        Mes_Num             = número del mes 1-12
        Año                 = año de la venta
        Organismo_Grupo     = clasificación del organismo en grupos predefinidos

    Args:
        file_bytes: Contenido binario del archivo .xlsx

    Returns:
        pd.DataFrame: DataFrame limpio con todas las columnas calculadas.

    Raises:
        ValueError: Si el archivo no tiene el formato esperado.
    """
    # ── 1. Leer el Excel ──────────────────────────────────────────────────
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
    except Exception as e:
        raise ValueError(f"No se pudo leer el archivo Excel: {e}") from e

    # ── 2. Limpiar nombres de columna (quitar espacios extra) ─────────────
    df.columns = df.columns.str.strip()

    # ── 3. Convertir columnas de existencias a numérico (NO se eliminan)  ─
    for col in ["Existencias Anteriores", "Existencias Posteriores"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # ── 4. Convertir Fecha_ES a datetime ──────────────────────────────────
    if "Fecha_ES" in df.columns:
        df["Fecha_ES"] = pd.to_datetime(df["Fecha_ES"], errors="coerce", dayfirst=True)
    elif "Fecha" in df.columns:
        # Fallback: usar la columna Fecha si no existe Fecha_ES
        df["Fecha_ES"] = pd.to_datetime(df["Fecha"], errors="coerce", dayfirst=True)
    else:
        raise ValueError(
            "El archivo no contiene la columna 'Fecha_ES' ni 'Fecha'. "
            "Verifica el formato del Excel."
        )

    # ── 5. Limpiar y convertir columnas numéricas ─────────────────────────
    numeric_cols = [
        "Cantidad (Unidades)", "Pvp", "PVP Facturado",
        "Importe Bruto", "Descuento", "Importe Neto", "Código",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # ── 6. Columnas calculadas ────────────────────────────────────────────

    # Facturación = Cantidad × PVP unitario
    if "Cantidad (Unidades)" in df.columns and "Pvp" in df.columns:
        df["Facturación"] = df["Cantidad (Unidades)"] * df["Pvp"]
    else:
        df["Facturación"] = 0.0

    # Diferencia_Gobierno = PVP Facturado - Importe Neto
    if "PVP Facturado" in df.columns and "Importe Neto" in df.columns:
        df["Diferencia_Gobierno"] = df["PVP Facturado"] - df["Importe Neto"]
    else:
        df["Diferencia_Gobierno"] = 0.0

    # Es_Generico: código > 600 000
    if "Código" in df.columns:
        df["Es_Generico"] = df["Código"] > 600_000
    else:
        df["Es_Generico"] = False

    # Hora_Int: extraer la hora como entero desde "HH:MM:SS"
    if "Hora" in df.columns:
        df["Hora_Int"] = (
            df["Hora"]
            .astype(str)
            .str.split(":")
            .str[0]
            .apply(lambda x: int(x) if str(x).strip().isdigit() else 0)
        )
    else:
        df["Hora_Int"] = 0

    # Dia_Semana, Mes, Mes_Num y Año desde Fecha_ES
    df["Dia_Semana"] = df["Fecha_ES"].dt.day_name()
    df["Mes"]        = df["Fecha_ES"].dt.month_name()
    df["Mes_Num"]    = df["Fecha_ES"].dt.month
    df["Año"]        = df["Fecha_ES"].dt.year

    # Organismo_Grupo: clasificación en grupos predefinidos
    if "Organismo" in df.columns:
        df["Organismo_Grupo"] = df["Organismo"].apply(_classify_organismo)
    else:
        df["Organismo_Grupo"] = "OTRAS ENTIDADES"

    # ── 7. Columnas derivadas de existencias (stock y roturas) ───────────
    if "Existencias Posteriores" in df.columns and "Código" in df.columns:
        # Rotura de stock: existencias posteriores negativas
        df["Rotura_Stock"] = df["Existencias Posteriores"] < 0
        # Stock bajo: entre 0 y 2 unidades (sin ser rotura)
        df["Stock_Bajo"] = (
            (df["Existencias Posteriores"] >= 0) &
            (df["Existencias Posteriores"] <= 2)
        )
        # Stock actual = última existencia posterior conocida por producto
        ultimo_stock = (
            df.sort_values("Fecha_ES")
            .groupby("Código")["Existencias Posteriores"]
            .last()
            .rename("Stock_Actual")
        )
        df = df.merge(ultimo_stock, on="Código", how="left")
    else:
        df["Rotura_Stock"] = False
        df["Stock_Bajo"]   = False
        df["Stock_Actual"] = 0

    # ── 8. Eliminar filas sin fecha válida ────────────────────────────────
    df = df.dropna(subset=["Fecha_ES"]).reset_index(drop=True)

    return df


# ─────────────────────────────────────────────────────────────────────────────
# Clasificación de organismos
# ─────────────────────────────────────────────────────────────────────────────

def _classify_organismo(organismo) -> str:
    """
    Clasifica el valor del campo 'Organismo' en uno de los grupos predefinidos.

    Reglas (en orden de prioridad):
        PXXI / GXXI / MAN  → RECETA XXI
        MUFACE / MFC        → MUFACE
        ISFAS / ISF         → ISFAS
        MUGEJU / MGJ        → MUGEJU
        001 - VTA. LIBRE    → VENTA LIBRE   (coincidencia exacta)
        DEPÓSITO            → DEPÓSITOS
        Resto               → OTRAS ENTIDADES

    Args:
        organismo: Valor del campo Organismo (puede ser str, float o None)

    Returns:
        str: Nombre del grupo.
    """
    if pd.isna(organismo) or str(organismo).strip() == "":
        return "OTRAS ENTIDADES"

    org_upper = str(organismo).upper().strip()

    if any(k in org_upper for k in ("PXXI", "GXXI", "MAN")):
        return "RECETA XXI"
    if any(k in org_upper for k in ("MUFACE", "MFC")):
        return "MUFACE"
    if any(k in org_upper for k in ("ISFAS", "ISF")):
        return "ISFAS"
    if any(k in org_upper for k in ("MUGEJU", "MGJ")):
        return "MUGEJU"
    if org_upper == "001 - VTA. LIBRE":
        return "VENTA LIBRE"
    if "DEP" in org_upper and ("SITO" in org_upper or "ÓSITO" in org_upper):
        return "DEPÓSITOS"

    return "OTRAS ENTIDADES"


# ─────────────────────────────────────────────────────────────────────────────
# Aplicación de filtros globales
# ─────────────────────────────────────────────────────────────────────────────

def apply_filters(
    df: pd.DataFrame,
    date_start=None,
    date_end=None,
    hour_min: int = 0,
    hour_max: int = 23,
    organismos: Optional[list] = None,
    tipos_op: Optional[list] = None,
    vendedores: Optional[list] = None,
    product_type: str = "Todos",
) -> pd.DataFrame:
    """
    Aplica los filtros globales del sidebar al DataFrame procesado.

    Args:
        df:           DataFrame devuelto por process_excel()
        date_start:   Fecha de inicio (datetime.date o None para sin límite)
        date_end:     Fecha de fin    (datetime.date o None para sin límite)
        hour_min:     Hora mínima del rango (0-23)
        hour_max:     Hora máxima del rango (0-23)
        organismos:   Lista de Organismo_Grupo a incluir (None = todos)
        tipos_op:     Lista de Tipo de Operación a incluir (None = todos)
        vendedores:   Lista de vendedores a incluir (None = todos)
        product_type: 'Todos' | 'Solo Genéricos' | 'Solo Marca'

    Returns:
        pd.DataFrame: Subconjunto del DataFrame original con los filtros aplicados.
    """
    filtered = df.copy()

    # Filtro por rango de fechas
    if date_start is not None and "Fecha_ES" in filtered.columns:
        filtered = filtered[filtered["Fecha_ES"].dt.date >= date_start]
    if date_end is not None and "Fecha_ES" in filtered.columns:
        filtered = filtered[filtered["Fecha_ES"].dt.date <= date_end]

    # Filtro por rango horario
    if "Hora_Int" in filtered.columns:
        filtered = filtered[
            (filtered["Hora_Int"] >= hour_min) & (filtered["Hora_Int"] <= hour_max)
        ]

    # Filtro por Organismo_Grupo
    if organismos and "Organismo_Grupo" in filtered.columns:
        filtered = filtered[filtered["Organismo_Grupo"].isin(organismos)]

    # Filtro por Tipo de Operación
    if tipos_op and "Tipo de Operación" in filtered.columns:
        filtered = filtered[filtered["Tipo de Operación"].isin(tipos_op)]

    # Filtro por Vendedor
    if vendedores and "Vendedor" in filtered.columns:
        filtered = filtered[filtered["Vendedor"].isin(vendedores)]

    # Filtro por tipo de producto (genérico vs marca)
    if "Es_Generico" in filtered.columns:
        if product_type == "Solo Genéricos":
            filtered = filtered[filtered["Es_Generico"] == True]
        elif product_type == "Solo Marca":
            filtered = filtered[filtered["Es_Generico"] == False]

    return filtered.reset_index(drop=True)
