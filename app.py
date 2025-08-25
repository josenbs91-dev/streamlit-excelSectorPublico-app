import streamlit as st
from reporte_flujo_caja import procesar_flujo_caja
from reporte_oa2 import procesar_oa2  # Función que necesita 2 archivos

st.title("Reportes del Sector Público")

# =======================
# REPORTE FLUJO DE CAJA
# =======================
st.header("Reporte Flujo de Caja")
file_formato_a = st.file_uploader("Subir Formato A", type=["xls","xlsx"], key="fc_a")
file_formato_b = st.file_uploader("Subir Formato B", type=["xls","xlsx"], key="fc_b")
file_formato_c = st.file_uploader("Subir Formato C", type=["xls","xlsx"], key="fc_c")

if file_formato_a and file_formato_b and file_formato_c:
    # La función ahora devuelve df_c_s adicional y df_estructura
    output_fc, df_a, df_b, df_c, df_c_s, df_estructura = procesar_flujo_caja(
        file_formato_a, file_formato_b, file_formato_c
    )

    if output_fc:
        st.success("Flujo de Caja procesado correctamente ✅")

        st.subheader("Resumen Formato A")
        st.dataframe(df_a)

        st.subheader("Resumen Formato B")
        st.dataframe(df_b)

        st.subheader("Resumen Formato C")
        st.dataframe(df_c)

        st.subheader("Resumen Formato C_s (Totales sin criterios 003-002 y 068-005)")
        st.dataframe(df_c_s)

        st.subheader("Estructura Consolidada")
        st.dataframe(df_estructura)

        st.download_button(
            label="📥 Descargar Flujo de Caja",
            data=output_fc.getvalue(),
            file_name="Reporte_Flujo_Caja.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# =======================
# REPORTE OA-2
# =======================
st.header("Reporte OA-2")
file_antes = st.file_uploader("Subir archivo ANTES", type=["xls","xlsx"], key="oa2_antes")
file_nuevo = st.file_uploader("Subir archivo MES NUEVO", type=["xls","xlsx"], key="oa2_nuevo")

if file_antes and file_nuevo:
    output_oa2, comparaciones = procesar_oa2(file_antes, file_nuevo)
    if output_oa2:
        st.success("OA-2 procesado correctamente ✅")

        # Mostrar cada hoja de comparaciones
        for nombre, df_comp in comparaciones.items():
            st.subheader(nombre)
            st.dataframe(df_comp)

        st.download_button(
            label="📥 Descargar OA-2",
            data=output_oa2.getvalue(),
            file_name="Reporte_OA2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
