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
    padding: 0.5rem 1rem;
    border: none;
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
st.caption("Carga tu Excel y calcula automáticamente las comisiones")

# =====================================================
# UPLOADER
# =====================================================

archivo = st.file_uploader(
    "📂 Subir archivo Excel",
    type=["xlsx", "xls"]
)

# =====================================================
# SI EXISTE ARCHIVO
# =====================================================

if archivo is not None:

    # =====================================================
    # LEER EXCEL
    # =====================================================

    try:

        df = pd.read_excel(archivo)

        st.success("✅ Archivo cargado correctamente")

    except Exception as e:

        st.error(f"❌ Error al leer archivo: {e}")
        st.stop()

    # =====================================================
    # LIMPIAR COLUMNAS
    # =====================================================

    df.columns = df.columns.str.strip().str.lower()

    # =====================================================
    # COLUMNAS IMPORTANTES
    # =====================================================

    columnas_importantes = [
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
        "itf",
        "comision_destino",
        "comision_origen",
        "yape_id",
        "fee_gmoney"
    ]

    # =====================================================
    # SOLO COLUMNAS IMPORTANTES
    # =====================================================

    columnas_existentes = [
        col for col in columnas_importantes
        if col in df.columns
    ]

    df = df[columnas_existentes].copy()

    # =====================================================
    # CONVERTIR COLUMNAS PROBLEMÁTICAS
    # =====================================================

    for col in df.columns:

        try:
            df[col] = df[col].astype(str)
        except:
            pass

    # =====================================================
    # VISTA PREVIA
    # =====================================================

    st.subheader("📋 Vista previa")

    st.dataframe(
        df.head(100),
        use_container_width=True,
        height=500
    )

    # =====================================================
    # SIDEBAR
    # =====================================================

    st.sidebar.header("⚙️ Configuración")

    porcentaje_comision = st.sidebar.number_input(
        "Porcentaje comisión (%)",
        min_value=0.0,
        value=3.50,
        step=0.1
    )

    tarifa_fija = st.sidebar.number_input(
        "Tarifa fija",
        min_value=0.0,
        value=1.00,
        step=0.1
    )

    comision_destino_input = st.sidebar.number_input(
        "Comisión destino",
        min_value=0.0,
        value=0.00,
        step=0.1
    )

    comision_origen_input = st.sidebar.number_input(
        "Comisión origen",
        min_value=0.0,
        value=0.00,
        step=0.1
    )

    tarifa_gmoney = st.sidebar.number_input(
        "Tarifa adicional GMONEY",
        min_value=0.0,
        value=0.50,
        step=0.1
    )

    aplicar_igv = st.sidebar.checkbox(
        "Aplicar IGV (18%)",
        value=True
    )

    # =====================================================
    # CONVERTIR FECHA
    # =====================================================

    df["creacion_deuda_fecha_peru"] = pd.to_datetime(
        df["creacion_deuda_fecha_peru"],
        errors="coerce"
    )

    # =====================================================
    # MES
    # =====================================================

    df["mes"] = df["creacion_deuda_fecha_peru"].dt.strftime("%Y-%m")

    # =====================================================
    # FILTRO MES
    # =====================================================

    meses = sorted(
        df["mes"].dropna().unique()
    )

    mes_seleccionado = st.selectbox(
        "📅 Selecciona el mes",
        meses
    )

    df_filtrado = df[
        df["mes"] == mes_seleccionado
    ].copy()

    # =====================================================
    # TOTAL NUMÉRICO
    # =====================================================

    df_filtrado["total"] = pd.to_numeric(
        df_filtrado["total"],
        errors="coerce"
    ).fillna(0)

    # =====================================================
    # COMISIÓN %
    # =====================================================

    df_filtrado["comision_porcentaje"] = (
        df_filtrado["total"] *
        porcentaje_comision / 100
    )

    # =====================================================
    # TARIFA FIJA
    # =====================================================

    df_filtrado["tarifa_fija"] = tarifa_fija

    # =====================================================
    # COMISIONES MANUALES
    # =====================================================

    df_filtrado["comision_destino"] = (
        comision_destino_input
    )

    df_filtrado["comision_origen"] = (
        comision_origen_input
    )

    # =====================================================
    # GMONEY
    # =====================================================

    df_filtrado["fee_gmoney"] = np.where(

        df_filtrado["operador_dispersion"]
        .astype(str)
        .str.upper()
        .str.contains("GMONEY"),

        tarifa_gmoney,

        0
    )

    # =====================================================
    # TOTAL SIN IGV
    # =====================================================

    df_filtrado["total_sin_igv"] = (

        df_filtrado["comision_porcentaje"] +

        df_filtrado["tarifa_fija"] +

        df_filtrado["comision_destino"] +

        df_filtrado["comision_origen"] +

        df_filtrado["fee_gmoney"]
    )

    # =====================================================
    # IGV
    # =====================================================

    if aplicar_igv:

        df_filtrado["igv"] = (
            df_filtrado["total_sin_igv"] * 0.18
        )

    else:

        df_filtrado["igv"] = 0

    # =====================================================
    # TOTAL CON IGV
    # =====================================================

    df_filtrado["total_con_igv"] = (

        df_filtrado["total_sin_igv"] +

        df_filtrado["igv"]
    )

    # =====================================================
    # REDONDEO
    # =====================================================

    columnas_redondeo = [
        "comision_porcentaje",
        "tarifa_fija",
        "comision_destino",
        "comision_origen",
        "fee_gmoney",
        "total_sin_igv",
        "igv",
        "total_con_igv"
    ]

    for col in columnas_redondeo:

        df_filtrado[col] = (
            pd.to_numeric(
                df_filtrado[col],
                errors="coerce"
            )
            .fillna(0)
            .round(2)
        )

    # =====================================================
    # RESULTADO STRING
    # =====================================================

    df_resultado = df_filtrado.copy()

    for col in df_resultado.columns:

        try:
            df_resultado[col] = (
                df_resultado[col]
                .astype(str)
            )
        except:
            pass

    # =====================================================
    # RESULTADO FINAL
    # =====================================================

    st.subheader("📊 Resultado Final")

    st.dataframe(
        df_resultado,
        use_container_width=True,
        height=600
    )

    # =====================================================
    # DASHBOARD
    # =====================================================

    st.subheader("📈 Dashboard")

    total_operaciones = len(df_filtrado)

    total_procesado = (
        df_filtrado["total"].sum()
    )

    total_comisiones = (
        df_filtrado["total_sin_igv"].sum()
    )

    total_igv = (
        df_filtrado["igv"].sum()
    )

    total_con_igv = (
        df_filtrado["total_con_igv"].sum()
    )

    total_gmoney = (

        df_filtrado[
            df_filtrado["operador_dispersion"]
            .astype(str)
            .str.upper()
            .str.contains("GMONEY")
        ]["total"]

        .sum()
    )

    total_bcp = (

        df_filtrado[
            df_filtrado["operador_dispersion"]
            .astype(str)
            .str.upper()
            .str.contains("BCP")
        ]["total"]

        .sum()
    )

    total_bbva = (

        df_filtrado[
            df_filtrado["operador_dispersion"]
            .astype(str)
            .str.upper()
            .str.contains("BBVA")
        ]["total"]

        .sum()
    )

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.metric(
            "Operaciones",
            f"{total_operaciones:,}"
        )

    with c2:
        st.metric(
            "Total Procesado",
            f"S/ {total_procesado:,.2f}"
        )

    with c3:
        st.metric(
            "Total Comisión",
            f"S/ {total_comisiones:,.2f}"
        )

    with c4:
        st.metric(
            "Total con IGV",
            f"S/ {total_con_igv:,.2f}"
        )

    c5, c6, c7, c8 = st.columns(4)

    with c5:
        st.metric(
            "IGV",
            f"S/ {total_igv:,.2f}"
        )

    with c6:
        st.metric(
            "GMONEY",
            f"S/ {total_gmoney:,.2f}"
        )

    with c7:
        st.metric(
            "BCP",
            f"S/ {total_bcp:,.2f}"
        )

    with c8:
        st.metric(
            "BBVA",
            f"S/ {total_bbva:,.2f}"
        )

    # =====================================================
    # DESCARGAR EXCEL
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

    excel_data = output.getvalue()

    st.download_button(
        label="📥 Descargar Excel Final",
        data=excel_data,
        file_name=f"reporte_instapayouts_{mes_seleccionado}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =====================================================
# SIN ARCHIVO
# =====================================================

else:

    st.info(
        "👆 Sube un archivo Excel para comenzar"
    )
