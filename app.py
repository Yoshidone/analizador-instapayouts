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
# SUBIR ARCHIVO
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
        "fee_gmoney"
    ]

    # =====================================================
    # FILTRAR COLUMNAS
    # =====================================================

    columnas_existentes = [
        col for col in columnas_importantes
        if col in df.columns
    ]

    df = df[columnas_existentes].copy()

    # =====================================================
    # FECHAS
    # =====================================================

    df["creacion_deuda_fecha_peru"] = pd.to_datetime(
        df["creacion_deuda_fecha_peru"],
        errors="coerce"
    )

    # =====================================================
    # TOTAL NUMÉRICO
    # =====================================================

    df["total"] = pd.to_numeric(
        df["total"],
        errors="coerce"
    ).fillna(0)

    # =====================================================
    # MES
    # =====================================================

    df["mes"] = (
        df["creacion_deuda_fecha_peru"]
        .dt.strftime("%Y-%m")
    )

    # =====================================================
    # SIDEBAR
    # =====================================================

    st.sidebar.header("⚙️ Configuración")

    porcentaje_comision = st.sidebar.number_input(
        "Porcentaje Comisión (%)",
        min_value=0.0,
        value=3.50,
        step=0.1
    )

    tarifa_fija = st.sidebar.number_input(
        "Tarifa Fija",
        min_value=0.0,
        value=1.00,
        step=0.1
    )

    tarifa_gmoney = st.sidebar.number_input(
        "Tarifa GMONEY",
        min_value=0.0,
        value=0.50,
        step=0.1
    )

    aplicar_igv = st.sidebar.checkbox(
        "Aplicar IGV (18%)",
        value=True
    )

    # =====================================================
    # FILTRO MES
    # =====================================================

    meses = sorted(
        df["mes"]
        .dropna()
        .unique()
    )

    mes_seleccionado = st.selectbox(
        "📅 Selecciona el mes",
        meses
    )

    df_filtrado = df[
        df["mes"] == mes_seleccionado
    ].copy()

    # =====================================================
    # FILTRO MONEDA
    # =====================================================

    monedas = sorted(
        df_filtrado["moneda"]
        .dropna()
        .astype(str)
        .unique()
    )

    moneda_seleccionada = st.selectbox(
        "💵 Selecciona moneda",
        monedas
    )

    df_filtrado = df_filtrado[
        df_filtrado["moneda"]
        == moneda_seleccionada
    ]

    # =====================================================
    # PORCENTAJE COMISION
    # SOLO SI NO ES GMONEY
    # =====================================================

    df_filtrado["porcentaje_comision"] = np.where(

        df_filtrado["operador_dispersion"]
        .astype(str)
        .str.upper()
        .str.contains("GMONEY"),

        0,

        (
            df_filtrado["total"] *
            porcentaje_comision / 100
        )
    )

    # =====================================================
    # TARIFA FIJA
    # SOLO SI NO ES GMONEY
    # =====================================================

    df_filtrado["tarifa_fija"] = np.where(

        df_filtrado["operador_dispersion"]
        .astype(str)
        .str.upper()
        .str.contains("GMONEY"),

        0,

        tarifa_fija
    )

    # =====================================================
    # FEE GMONEY
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
    # IGV
    # =====================================================

    df_filtrado["igv"] = (

        df_filtrado["porcentaje_comision"] +

        df_filtrado["tarifa_fija"] +

        df_filtrado["fee_gmoney"]
    )

    # =====================================================
    # APLICAR IGV
    # =====================================================

    if aplicar_igv:

        df_filtrado["igv"] = (
            df_filtrado["igv"] * 1.18
        )

    # =====================================================
    # NETO
    # =====================================================

    df_filtrado["neto"] = (

        df_filtrado["total"] -

        df_filtrado["igv"]
    )

    # =====================================================
    # REDONDEO
    # =====================================================

    columnas_redondeo = [
        "porcentaje_comision",
        "tarifa_fija",
        "fee_gmoney",
        "igv",
        "neto"
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
    # RESULTADO FINAL
    # =====================================================

    st.subheader("📊 Resultado Final")

    st.dataframe(
        df_filtrado,
        use_container_width=True,
        height=650
    )

    # =====================================================
    # DASHBOARD
    # =====================================================

    st.subheader("📈 Dashboard")

    total_operaciones = len(df_filtrado)

    total_procesado = (
        df_filtrado["total"].sum()
    )

    total_comision = (
        df_filtrado["igv"].sum()
    )

    total_neto = (
        df_filtrado["neto"].sum()
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
            f"Total {moneda_seleccionada}",
            f"{moneda_seleccionada} {total_procesado:,.2f}"
        )

    with c3:
        st.metric(
            "Total Comisión",
            f"{moneda_seleccionada} {total_comision:,.2f}"
        )

    with c4:
        st.metric(
            "NETO",
            f"{moneda_seleccionada} {total_neto:,.2f}"
        )

    c5, c6, c7 = st.columns(3)

    with c5:
        st.metric(
            "GMONEY",
            f"{moneda_seleccionada} {total_gmoney:,.2f}"
        )

    with c6:
        st.metric(
            "BCP",
            f"{moneda_seleccionada} {total_bcp:,.2f}"
        )

    with c7:
        st.metric(
            "BBVA",
            f"{moneda_seleccionada} {total_bbva:,.2f}"
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
        file_name=f"reporte_instapayouts_{mes_seleccionado}_{moneda_seleccionada}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =====================================================
# SIN ARCHIVO
# =====================================================

else:

    st.info(
        "👆 Sube un archivo Excel para comenzar"
    )
