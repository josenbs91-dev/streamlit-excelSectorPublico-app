import streamlit as st
import pandas as pd
from io import BytesIO

def procesar_flujo_caja(file_formato_a, file_formato_b, file_formato_c):
    """
    Procesa los archivos Formato A, B y C y genera un resumen consolidado
    con columna adicional 'Ubicación en el Flujo' y hoja 'Estructura Consolidada'.
    """
    try:
        # ===== Leer archivos =====
        df_a = pd.read_excel(file_formato_a, dtype=str)
        df_b = pd.read_excel(file_formato_b, dtype=str)
        df_c = pd.read_excel(file_formato_c, dtype=str)

        # Convertir montos a numérico
        df_a["monto_nacional"] = pd.to_numeric(df_a["monto_nacional"], errors="coerce").fillna(0)
        df_b["monto_nacional"] = pd.to_numeric(df_b["monto_nacional"], errors="coerce").fillna(0)
        df_c["monto_nacional"] = pd.to_numeric(df_c["monto_nacional"], errors="coerce").fillna(0)

        # ===== Mapeos de Ubicación en el Flujo =====
        flujo_a = {
            "2.1. 1": "2.1. Personal y Obligaciones Sociales",
            "2.1. 3": "2.1. Personal y Obligaciones Sociales",
            "2.1. 5": "2.4.2 Otros Gastos Corrientes Sentencias Judiciales",
            "2.3. 1": "2.3. Bienes y Servicios",
            "2.3. 2": "2.3. Bienes y Servicios",
            "2.5. 4": "2.4.3 Otros Gastos Corrientes P. Iimp, D. Adm. y Multas Gub.",
            "2.6. 3": "4.3 Otros Gastos de Capital"
        }
        flujo_b = {
            "1.1. 4": "1.1. Impuestos y Contribuciones Obligatorias",
            "1.3. 1": "1.2. Venta de Bienes",
            "1.3. 2": "1.3. Prestación de Servicios",
            "1.3. 3": "1.3. Prestación de Servicios",
            "1.5. 1": "1.4.1 De la Propiedad Financiera",
            "1.5. 2": "1.5.1 Otros Ingresos Corrientes",
            "1.5. 5": "1.5.1 Otros Ingresos Corrientes",
            "1.9. 1": "1.5.1 Otros Ingresos Corrientes"
        }

        # ===== Formato A =====
        dict_a = {}
        for _, row in df_a[df_a["fase"]=="G"].iterrows():
            clave = str(row["clasificador"])[:6]
            dict_a[clave] = dict_a.get(clave, 0) + row["monto_nacional"]
        df_resumen_a = pd.DataFrame({"Clasificador": list(dict_a.keys()), 
                                     "Monto Nacional": list(dict_a.values())})
        df_resumen_a["Ubicación en el Flujo"] = df_resumen_a["Clasificador"].map(flujo_a)

        # ===== Formato B =====
        dict_b = {}
        for _, row in df_b[df_b["fase"]=="R"].iterrows():
            clave = str(row["clasificador"])[:6]
            dict_b[clave] = dict_b.get(clave, 0) + row["monto_nacional"]
        df_resumen_b = pd.DataFrame({"Clasificador": list(dict_b.keys()), 
                                     "Monto Nacional": list(dict_b.values())})
        df_resumen_b["Ubicación en el Flujo"] = df_resumen_b["Clasificador"].map(flujo_b)

        # ===== Formato C =====
        df_c["criterios"] = df_c["banco"].astype(str) + "-" + df_c["cta_cte"].astype(str)
        dict_c = {}
        for _, row in df_c.iterrows():
            if row["tipo_operacion"] != "TC" and row["fase"] in ["G", "R"] and row["criterios"] in ["003-002","068-005"]:
                clave = row["fase"] + "|" + row["criterios"]
                dict_c[clave] = dict_c.get(clave, 0) + row["monto_nacional"]

        resumen_c_rows = []
        for fase in ["G", "R"]:
            subtotal = 0
            for key, val in dict_c.items():
                if key.startswith(fase):
                    resumen_c_rows.append({"Fase": fase, "Criterios": key.split("|")[1], "Monto Nacional": val})
                    subtotal += val
            resumen_c_rows.append({"Fase": "", "Criterios": f"TOTAL {fase}", "Monto Nacional": subtotal})
        df_resumen_c = pd.DataFrame(resumen_c_rows)
        df_resumen_c["Ubicación en el Flujo"] = ""
        df_resumen_c.loc[df_resumen_c["Criterios"]=="TOTAL G", "Ubicación en el Flujo"] = "2.4.1 Otros Gastos Corrientes Arancel Distribuido"
        df_resumen_c.loc[df_resumen_c["Criterios"]=="TOTAL R", "Ubicación en el Flujo"] = "1.1. Impuestos y Contribuciones Obligatorias por Distribuir"

        # ===== Formato C_S =====
        dict_c_s = {}
        for _, row in df_c.iterrows():
            if row["fase"] in ["G", "R"] and row["criterios"] not in ["003-002","068-005"] and row["tipo_operacion"] != "TC":
                dict_c_s[row["fase"]] = dict_c_s.get(row["fase"], 0) + row["monto_nacional"]

        resumen_c_s_rows = []
        for fase, val in dict_c_s.items():
            resumen_c_s_rows.append({"Fase": "", "Criterios": f"TOTAL {fase}", "Monto Nacional": val})
        df_resumen_c_s = pd.DataFrame(resumen_c_s_rows)
        df_resumen_c_s["Ubicación en el Flujo"] = ""
        df_resumen_c_s.loc[df_resumen_c_s["Criterios"]=="TOTAL G", "Ubicación en el Flujo"] = "2.4.6 Otros Gastos Corrientes"
        df_resumen_c_s.loc[df_resumen_c_s["Criterios"]=="TOTAL R", "Ubicación en el Flujo"] = "1.5.1 Otros Ingresos Corrientes"

        # ===== Estructura Consolidada =====
        dfs_para_consolidar = []

        df_a_consol = df_resumen_a[["Ubicación en el Flujo", "Monto Nacional"]]
        df_b_consol = df_resumen_b[["Ubicación en el Flujo", "Monto Nacional"]]
        df_c_consol = df_resumen_c[["Ubicación en el Flujo", "Monto Nacional"]]
        df_c_s_consol = df_resumen_c_s[["Ubicación en el Flujo", "Monto Nacional"]]

        dfs_para_consolidar.extend([df_a_consol, df_b_consol, df_c_consol, df_c_s_consol])

        df_estructura = pd.concat(dfs_para_consolidar)
        df_estructura = df_estructura.groupby("Ubicación en el Flujo", as_index=False).sum()

        # ===== Exportar todo a Excel =====
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_resumen_a.to_excel(writer, index=False, sheet_name="FormatoA")
            df_resumen_b.to_excel(writer, index=False, sheet_name="FormatoB")
            df_resumen_c.to_excel(writer, index=False, sheet_name="FormatoC")
            df_resumen_c_s.to_excel(writer, index=False, sheet_name="FormatoC_S")
            df_estructura.to_excel(writer, index=False, sheet_name="Estructura Consolidada")
        output.seek(0)

        return output, df_resumen_a, df_resumen_b, df_resumen_c, df_resumen_c_s, df_estructura

    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
        return None, None, None, None, None, None
