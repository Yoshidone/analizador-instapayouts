import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# =====================================================
# CONFIGURACIÓN
# =====================================================
st.set_page_config(
    page_title="Analizador INSTAPAYOUTS",
    page_icon="💸",
    layout="wide"
)

# =====================================================
# ESTILOS
# =====================================================
st.markdown("""
<style>

.main {
    background-color: #f8fafc;
}

.block-container {
    padding-top: 2rem;
}

h1, h2, h3 {
    color: #0f172a;
}

.stButton>button {
    background-color: #2563eb;
    color: white;
    border-radius: 10px;
    border: none;
    padding: 0.5rem 1rem;
    font-weight: 600;
}

.stDownloadButton>button {
    background-color: #16a34a;
    color: white;
    border-radius: 10px;
    border: none;
    padding: 0.5rem 1rem;
    font-weight: 600;
}

[data-testid="metric-container"] {
    background-color: white;
    border-radius: 15px;
    padding: 15px;
    box-shadow: 0px 2px 8px rgba(0,0,0,0.08);
}

</style>
""", unsafe_allow_html=True)

# =====================================================
# TÍTULO
# =====================================================
st.title("💸 Analizador INSTAPAYOUTS")
st.caption("Carga tu Excel y calcula comisiones automáticamente")

# =====================================================
# SUBIR ARCHIVO
# =====================================================
archivo = st.file_uploader(
    "📂 Subir archivo Excel",
    type=["xlsx", "xls"]
)

# =====================================================
# VALIDACIÓN
# =====================================================
if archivo is not None:

    # =====================================================
    # LEER EXCEL
    # =====================================================
    try:
        df = pd.read_excel(archivo)

        st.success("✅ Archivo cargado correctamente")

    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        st.stop()

    # =====================================================
    # LIMPIAR COLUMNAS
    # =====================================================
    df.columns = df.columns.str.strip().str.lower()

    # =====================================================
    # PREVIEW
    # =====================================================
    st.subheader("📋 Vista previa")

    st.dataframe(
        df.head(20),
        use_container_width=True
    )

    # =====================================================
    # SIDEBAR
    # =====================================================
    st.sidebar.header("⚙️ Configuración")

    porcentaje = st.sidebar.number_input(
        "Porcentaje comisión (%)",
        min_value=0.0,
        value=3.50,
        step=0.1
    )

    fee_fijo = st.sidebar.number_input(
        "Fee fijo",
        min_value=0.0,
        value=1.00,
        step=0.1
    )

    fee_gmoney = st.sidebar.number_input(
        "Fee adicional GMONEY",
        min_value=0.0,
        value=0.50,
        step=0.1
    )

    aplicar_igv = st.sidebar.checkbox(
        "Aplicar IGV (18%)",
        value=True
    )

    # =====================================================
    # COLUMNAS
    # =====================================================
    columnas = df.columns.tolist()

    st.subheader("🔎 Configuración de columnas")

    c1, c2, c3 = st.columns(3)

    with c1:
        columna_monto = st.selectbox(
            "Columna monto",
            columnas
        )

    with c2:
        columna_fecha = st.selectbox(
            "Columna fecha",
            columnas
        )

    with c3:
        columna_gmoney = st.selectbox(
            "Columna fee_gmoney",
            columnas,
            index=columnas.index("fee_gmoney") if "fee_gmoney" in columnas else 0
        )

    # =====================================================
    # FECHA
    # =====================================================
    try:

        df[columna_fecha] = pd.to_datetime(
            df[columna_fecha],
            errors="coerce"
        )

        df["mes"] = df[columna_fecha].dt.strftime("%Y-%m")

    except:
        st.error("❌ No se pudo convertir la columna fecha")
        st.stop()

    # =====================================================
    # FILTRO MES
    # =====================================================
    meses = sorted(df["mes"].dropna().unique())

    mes_seleccionado = st.selectbox(
        "📅 Selecciona el mes",
        meses
    )

    df_filtrado = df[df["mes"] == mes_seleccionado].copy()

    # =====================================================
    # MONTO NUMÉRICO
    # =====================================================
    df_filtrado[columna_monto] = pd.to_numeric(
        df_filtrado[columna_monto],
        errors="coerce"
    ).fillna(0)

    # =====================================================
    # COMISIONES
    # =====================================================

    # Comisión %
    df_filtrado["comision_porcentaje"] = (
        df_filtrado[columna_monto] * porcentaje / 100
    )

    # Fee fijo
    df_filtrado["fee_fijo_resultado"] = fee_fijo

    # Comisión GMONEY
    df_filtrado["comision_gmoney"] = np.where(
        df_filtrado[columna_gmoney].notna(),
        fee_gmoney,
        0
    )

    # Total sin IGV
    df_filtrado["total_sin_igv"] = (
        df_filtrado["comision_porcentaje"] +
        df_filtrado["fee_fijo_resultado"] +
        df_filtrado["comision_gmoney"]
    )

    # IGV
    if aplicar_igv:

        df_filtrado["igv"] = (
            df_filtrado["total_sin_igv"] * 0.18
        )

    else:

        df_filtrado["igv"] = 0

    # Total con IGV
    df_filtrado["total_con_igv"] = (
        df_filtrado["total_sin_igv"] +
        df_filtrado["igv"]
    )

    # =====================================================
    # REDONDEO
    # =====================================================
    columnas_redondeo = [
        "comision_porcentaje",
        "fee_fijo_resultado",
        "comision_gmoney",
        "total_sin_igv",
        "igv",
        "total_con_igv"
    ]

    for col in columnas_redondeo:
        df_filtrado[col] = df_filtrado[col].round(2)

    # =====================================================
    # RESULTADO
    # =====================================================
    st.subheader("📊 Resultado Final")

    st.dataframe(
        df_filtrado,
        use_container_width=True,
        height=500
    )

    # =====================================================
    # DASHBOARD
    # =====================================================
    st.subheader("📈 Dashboard")

    total_operaciones = len(df_filtrado)

    monto_total = df_filtrado[columna_monto].sum()

    total_sin_igv = df_filtrado["total_sin_igv"].sum()

    total_igv = df_filtrado["igv"].sum()

    total_con_igv = df_filtrado["total_con_igv"].sum()

    d1, d2, d3, d4, d5 = st.columns(5)

    with d1:
        st.metric(
            "Operaciones",
            f"{total_operaciones:,}"
        )

    with d2:
        st.metric(
            "Monto Total",
            f"S/ {monto_total:,.2f}"
        )

    with d3:
        st.metric(
            "Total Sin IGV",
            f"S/ {total_sin_igv:,.2f}"
        )

    with d4:
        st.metric(
            "IGV",
            f"S/ {total_igv:,.2f}"
        )

    with d5:
        st.metric(
            "Total Con IGV",
            f"S/ {total_con_igv:,.2f}"
        )

    # =====================================================
    # RESUMEN
    # =====================================================
    st.subheader("📌 Resumen por Mes")

    resumen = df.groupby("mes").agg({
        columna_monto: "sum"
    }).reset_index()

    resumen.columns = [
        "Mes",
        "Monto Total"
    ]

    st.dataframe(
        resumen,
        use_container_width=True
    )

    # =====================================================
    # DESCARGA EXCEL
    # =====================================================
    st.subheader("⬇️ Descargar Reporte")

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df_filtrado.to_excel(
            writer,
            index=False,
            sheet_name="Reporte"
        )

        resumen.to_excel(
            writer,
            index=False,
            sheet_name="Resumen"
        )

    excel_data = output.getvalue()

    st.download_button(
        label="📥 Descargar Excel Final",
        data=excel_data,
        file_name=f"reporte_instapayouts_{mes_seleccionado}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
