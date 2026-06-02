import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Analizador INSTAPAYOUTS",
    page_icon="💸",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; }

    [data-testid="metric-container"] {
        background: white;
        border-radius: 15px;
        padding: 15px;
        box-shadow: 0px 2px 8px rgba(0,0,0,0.08);
        border: 1px solid #e2e8f0;
        transition: box-shadow 0.2s ease;
    }
    [data-testid="metric-container"]:hover {
        box-shadow: 0px 4px 16px rgba(0,0,0,0.12);
    }

    [data-testid="stDownloadButton"] button {
        background-color: #16a34a;
        color: white;
        border-radius: 10px;
        border: none;
        font-weight: 600;
        width: 100%;
        transition: background-color 0.2s ease;
    }
    [data-testid="stDownloadButton"] button:hover {
        background-color: #15803d;
    }

    [data-testid="stFileUploader"] {
        border: 2px dashed #cbd5e1;
        border-radius: 12px;
        padding: 1rem;
        background-color: #f8fafc;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.title("💸 Analizador INSTAPAYOUTS")
st.caption("Carga tu Excel y calcula automáticamente las comisiones por rango")

# ─────────────────────────────────────────────────────────────────────────────
# COLUMNAS QUE SE USAN DEL EXCEL
# ─────────────────────────────────────────────────────────────────────────────
COLUMNAS_REQUERIDAS = {
    "creacion_deuda_fecha_peru",
    "operador_dispersion",
    "moneda",
    "total",
}

COLUMNAS_OPCIONALES = [
    "creacion_deuda_fecha_peru",
    "cus_name",
    "operador_dispersion",
    "referencia",
    "moneda",
    "total",
    "tipo_de_documento",
    "numero_documento",
    "cliente",
    "cuenta",
    "cci",
    "tipo_de_cuenta",
    "estado",
    "fecha_pagado_rechazado_peru",
]

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — CONFIGURACIÓN DE TARIFAS Y RANGOS
# ─────────────────────────────────────────────────────────────────────────────
st.sidebar.header("⚙️ Configuración de comisiones")

st.sidebar.subheader("💰 Tarifas fijas")
tarifa_fija   = st.sidebar.number_input("Tarifa fija (por operación)",  min_value=0.0, value=0.00, step=0.10, help="Se suma a cada operación que no sea GMONEY exclusivo")
tarifa_gmoney = st.sidebar.number_input("Tarifa GMONEY (por operación)", min_value=0.0, value=0.00, step=0.10, help="Se aplica solo a operaciones GMONEY cuando no usan misma comisión")

st.sidebar.subheader("🔧 Opciones")
gmoney_misma_comision = st.sidebar.checkbox("GMONEY usa misma comisión que el resto", value=True,  help="Si está activo, GMONEY paga la misma comisión porcentual/fija que los demás operadores")
aplicar_igv           = st.sidebar.checkbox("Aplicar IGV (18%)",                       value=True,  help="Multiplica el total de comisión × 1.18")

st.sidebar.subheader("📌 Rangos de comisión")
st.sidebar.caption("Se aplican en orden: primero el rango, luego el porcentaje o monto fijo.")

lim_bajo        = st.sidebar.number_input("Rango 1 — hasta monto (≤)",        min_value=0.0,   value=500.0,  step=100.0,  help="Montos menores o iguales a este valor usan comisión %")
pct_bajo        = st.sidebar.number_input("Rango 1 — comisión %",              min_value=0.0,   value=0.72,   step=0.01,   help="Porcentaje aplicado sobre el monto (ej: 0.72 = 0.72%)")
lim_medio_desde = st.sidebar.number_input("Rango 2 — desde monto (≥)",         min_value=0.0,   value=500.01, step=100.0,  help="Inicio del rango con comisión fija")
lim_medio_hasta = st.sidebar.number_input("Rango 2 — hasta monto (≤)",         min_value=0.0,   value=3000.0, step=100.0,  help="Fin del rango con comisión fija")
comision_fija   = st.sidebar.number_input("Rango 2 — comisión fija (importe)", min_value=0.0,   value=3.80,   step=0.10,   help="Monto fijo independiente del importe de la operación")
pct_alto        = st.sidebar.number_input("Rango 3 — comisión % (> rango 2)",  min_value=0.0,   value=1.00,   step=0.01,   help="Porcentaje aplicado sobre montos que superan el límite del rango 2")

# ─────────────────────────────────────────────────────────────────────────────
# SUBIR ARCHIVO
# ─────────────────────────────────────────────────────────────────────────────
st.divider()

archivo = st.file_uploader(
    "📂 Subir archivo Excel (.xlsx / .xls)",
    type=["xlsx", "xls"],
    help="El archivo debe tener las columnas: creacion_deuda_fecha_peru, operador_dispersion, moneda, total",
)

# ─────────────────────────────────────────────────────────────────────────────
# PROCESAMIENTO
# ─────────────────────────────────────────────────────────────────────────────
if archivo is not None:

    # ── Leer Excel ────────────────────────────────────────────────────────────
    try:
        df = pd.read_excel(archivo, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ Error leyendo el archivo: {e}")
        st.stop()

    # ── Normalizar nombres de columnas ────────────────────────────────────────
    df.columns = df.columns.str.strip().str.lower()

    # ── Validar columnas mínimas necesarias ───────────────────────────────────
    faltantes = COLUMNAS_REQUERIDAS - set(df.columns)
    if faltantes:
        st.error(
            f"⚠️ El archivo no tiene estas columnas requeridas: `{'`, `'.join(sorted(faltantes))}`\n\n"
            "Verifica que el Excel tenga el formato correcto."
        )
        st.stop()

    # ── Seleccionar solo las columnas que existen ─────────────────────────────
    columnas_presentes = [c for c in COLUMNAS_OPCIONALES if c in df.columns]
    df = df[columnas_presentes].copy()

    # ── Parsear fechas ────────────────────────────────────────────────────────
    df["creacion_deuda_fecha_peru"] = pd.to_datetime(
        df["creacion_deuda_fecha_peru"], errors="coerce"
    )

    # ── Asegurar que total sea numérico ───────────────────────────────────────
    df["total"] = pd.to_numeric(df["total"], errors="coerce").fillna(0)

    # ── Crear columna de mes para filtrar ─────────────────────────────────────
    df["mes"] = df["creacion_deuda_fecha_peru"].dt.strftime("%Y-%m")

    meses_disponibles = sorted(df["mes"].dropna().unique())

    if not meses_disponibles:
        st.error("❌ No se encontraron fechas válidas en `creacion_deuda_fecha_peru`.")
        st.stop()

    col_ok1, col_ok2 = st.columns([3, 1])
    col_ok1.success(f"✅ Archivo cargado — **{len(df):,} filas** · **{len(df.columns)} columnas**")
    col_ok2.caption(f"Meses detectados: {len(meses_disponibles)}")

    # ─────────────────────────────────────────────────────────────────────────
    # FILTROS DE MES Y MONEDA
    # ─────────────────────────────────────────────────────────────────────────
    st.divider()
    col_f1, col_f2, col_f3 = st.columns([2, 2, 3])

    mes_sel = col_f1.selectbox("📅 Mes", meses_disponibles)

    df_mes = df[df["mes"] == mes_sel].copy()

    monedas_disponibles = sorted(df_mes["moneda"].dropna().astype(str).unique())
    if not monedas_disponibles:
        st.warning("⚠️ No hay monedas disponibles para el mes seleccionado.")
        st.stop()

    moneda_sel = col_f2.selectbox("💵 Moneda", monedas_disponibles)

    df_filtrado = df_mes[df_mes["moneda"] == moneda_sel].copy()

    col_f3.metric(
        "📊 Registros en este filtro",
        f"{len(df_filtrado):,}",
        help="Total de filas para el mes y moneda seleccionados",
    )

    if df_filtrado.empty:
        st.warning("⚠️ No hay datos para el mes y moneda seleccionados.")
        st.stop()

    # ─────────────────────────────────────────────────────────────────────────
    # DETECCIÓN DE GMONEY
    # ─────────────────────────────────────────────────────────────────────────
    es_gmoney = (
        df_filtrado["operador_dispersion"]
        .astype(str)
        .str.upper()
        .str.contains("GMONEY", na=False)
    )

    # ─────────────────────────────────────────────────────────────────────────
    # FUNCIÓN DE COMISIÓN POR RANGO
    #
    # Lógica:
    #   total ≤ lim_bajo                        → total × (pct_bajo / 100)
    #   lim_medio_desde ≤ total ≤ lim_medio_hasta → comision_fija (importe fijo)
    #   total > lim_medio_hasta                 → total × (pct_alto / 100)
    # ─────────────────────────────────────────────────────────────────────────
    def calcular_comision_por_rango(total: float) -> float:
        if total <= lim_bajo:
            return total * (pct_bajo / 100)
        elif lim_medio_desde <= total <= lim_medio_hasta:
            return comision_fija
        else:
            return total * (pct_alto / 100)

    # ─────────────────────────────────────────────────────────────────────────
    # CÁLCULO DE COMISIONES COLUMNA A COLUMNA
    # ─────────────────────────────────────────────────────────────────────────

    # 1) Comisión por rango (según el monto de cada operación)
    df_filtrado["comision_por_rango"] = df_filtrado["total"].apply(calcular_comision_por_rango)

    # Si GMONEY NO usa misma comisión → sus operaciones no pagan comisión por rango
    if not gmoney_misma_comision:
        df_filtrado.loc[es_gmoney, "comision_por_rango"] = 0.0

    # 2) Tarifa fija por operación
    #    - Si GMONEY usa misma comisión → todos pagan tarifa_fija
    #    - Si NO → GMONEY paga 0 en tarifa fija
    if gmoney_misma_comision:
        df_filtrado["tarifa_fija_op"] = tarifa_fija
    else:
        df_filtrado["tarifa_fija_op"] = np.where(es_gmoney, 0.0, tarifa_fija)

    # 3) Fee exclusivo GMONEY (solo aplica si GMONEY NO usa misma comisión)
    if gmoney_misma_comision:
        df_filtrado["fee_gmoney"] = 0.0
    else:
        df_filtrado["fee_gmoney"] = np.where(es_gmoney, tarifa_gmoney, 0.0)

    # 4) Subtotal de comisión antes de IGV
    df_filtrado["subtotal_comision"] = (
        df_filtrado["comision_por_rango"]
        + df_filtrado["tarifa_fija_op"]
        + df_filtrado["fee_gmoney"]
    )

    # 5) IGV (18%) sobre el subtotal — si está activado
    if aplicar_igv:
        df_filtrado["igv_monto"]        = (df_filtrado["subtotal_comision"] * 0.18).round(2)
        df_filtrado["total_comision"]   = (df_filtrado["subtotal_comision"] * 1.18).round(2)
    else:
        df_filtrado["igv_monto"]        = 0.0
        df_filtrado["total_comision"]   = df_filtrado["subtotal_comision"].round(2)

    # 6) Neto que recibe el cliente
    df_filtrado["neto"] = (df_filtrado["total"] - df_filtrado["total_comision"]).round(2)

    # ── Redondeo final de todas las columnas calculadas ───────────────────────
    cols_calculadas = [
        "comision_por_rango", "tarifa_fija_op", "fee_gmoney",
        "subtotal_comision", "igv_monto", "total_comision", "neto",
    ]
    for col in cols_calculadas:
        df_filtrado[col] = df_filtrado[col].round(2)

    # ─────────────────────────────────────────────────────────────────────────
    # TABLA DE RESULTADOS
    # ─────────────────────────────────────────────────────────────────────────
    st.divider()
    st.subheader("📊 Resultado por operación")

    # Construir DataFrame de visualización con tipos seguros (evita error PyArrow)
    df_vista = df_filtrado.copy()
    for col in df_vista.select_dtypes(include=["datetime64[ns]", "datetime64[ns, UTC]"]).columns:
        df_vista[col] = df_vista[col].astype(str)

    if len(df_vista) > 500:
        st.info(f"ℹ️ Mostrando las primeras 500 filas de {len(df_vista):,}. Descarga el Excel para verlas todas.")

    st.dataframe(
        df_vista.head(500),
        use_container_width=True,
        hide_index=True,
        height=500,
    )

    # ─────────────────────────────────────────────────────────────────────────
    # DASHBOARD RESUMEN
    # ─────────────────────────────────────────────────────────────────────────
    st.divider()
    st.subheader("📈 Resumen del período")

    sim = moneda_sel  # símbolo / código de moneda

    total_ops        = len(df_filtrado)
    total_procesado  = df_filtrado["total"].sum()
    total_comision   = df_filtrado["total_comision"].sum()
    total_igv        = df_filtrado["igv_monto"].sum()
    total_neto       = df_filtrado["neto"].sum()

    total_gmoney_monto = df_filtrado.loc[es_gmoney, "total"].sum()
    total_bcp_monto    = df_filtrado[df_filtrado["operador_dispersion"].astype(str).str.upper().str.contains("BCP",   na=False)]["total"].sum()
    total_bbva_monto   = df_filtrado[df_filtrado["operador_dispersion"].astype(str).str.upper().str.contains("BBVA",  na=False)]["total"].sum()

    # Fila 1
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🔢 Operaciones",        f"{total_ops:,}")
    c2.metric(f"💰 Total {sim}",       f"{sim} {total_procesado:,.2f}")
    c3.metric("💸 Comisión total",     f"{sim} {total_comision:,.2f}", help="Incluye IGV si está activado")
    c4.metric("🧮 Neto",              f"{sim} {total_neto:,.2f}")

    # Fila 2
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("🏛 IGV aplicado",       f"{sim} {total_igv:,.2f}"      if aplicar_igv else "No aplicado")
    c6.metric("🏧 GMONEY",            f"{sim} {total_gmoney_monto:,.2f}")
    c7.metric("🏦 BCP",               f"{sim} {total_bcp_monto:,.2f}")
    c8.metric("🏦 BBVA",              f"{sim} {total_bbva_monto:,.2f}")

    # ─────────────────────────────────────────────────────────────────────────
    # RESUMEN POR OPERADOR
    # ─────────────────────────────────────────────────────────────────────────
    st.divider()
    st.subheader("🏪 Resumen por operador de dispersión")

    resumen_operadores = (
        df_filtrado
        .groupby("operador_dispersion", as_index=False)
        .agg(
            operaciones      = ("total",           "count"),
            total_procesado  = ("total",           "sum"),
            total_comision   = ("total_comision",  "sum"),
            total_neto       = ("neto",            "sum"),
        )
        .sort_values("total_procesado", ascending=False)
    )
    for col in ["total_procesado", "total_comision", "total_neto"]:
        resumen_operadores[col] = resumen_operadores[col].round(2)

    st.dataframe(
        resumen_operadores,
        use_container_width=True,
        hide_index=True,
    )

    # ─────────────────────────────────────────────────────────────────────────
    # DESCARGAR EXCEL
    # ─────────────────────────────────────────────────────────────────────────
    st.divider()
    st.subheader("⬇️ Descargar reporte")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Hoja principal con el detalle por operación
        df_filtrado.to_excel(writer, index=False, sheet_name="Detalle")
        # Hoja de resumen por operador
        resumen_operadores.to_excel(writer, index=False, sheet_name="Por Operador")

    col_dl1, col_dl2 = st.columns([1, 2])
    col_dl1.download_button(
        label="📥 Descargar Excel completo",
        data=output.getvalue(),
        file_name=f"instapayouts_{mes_sel}_{moneda_sel}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    col_dl2.caption(
        f"El Excel incluye 2 hojas: **Detalle** ({len(df_filtrado):,} filas) y **Por Operador** ({len(resumen_operadores)} filas)"
    )

# ─────────────────────────────────────────────────────────────────────────────
# SIN ARCHIVO
# ─────────────────────────────────────────────────────────────────────────────
else:
    st.info("👆 Sube un archivo Excel para comenzar el análisis.")

    with st.expander("ℹ️ ¿Cómo funciona el cálculo de comisiones?"):
        st.markdown(
            """
            El cálculo se hace **operación por operación** en 3 rangos configurables en el sidebar:

            | Rango | Condición | Comisión |
            |-------|-----------|----------|
            | Bajo  | total ≤ límite 1 | total × % configurado |
            | Medio | límite 2 ≤ total ≤ límite 3 | monto fijo |
            | Alto  | total > límite 3 | total × % configurado |

            Luego se suman:
            - **Comisión por rango** (calculada arriba)
            - **Tarifa fija** por operación (configurable)
            - **Fee GMONEY** si aplica

            Si IGV está activo, el subtotal × 1.18 = **comisión total**.

            **Neto = Total operación − Comisión total**
            """
        )
