import pandas as pd
from io import BytesIO
from openpyxl.worksheet.table import Table, TableStyleInfo

def procesar_oa2(file_antes, file_nuevo):
    """
    Procesa los archivos ANTES y MES NUEVO y genera un reporte OA-2.
    Devuelve un BytesIO con el Excel listo para descargar y un diccionario con las comparaciones.
    """
    try:
        # Leer archivos Excel
        tabla_antes = pd.read_excel(file_antes, dtype=str)
        tabla_nuevo = pd.read_excel(file_nuevo, dtype=str)

        # Convertir montos a numérico
        tabla_antes["MONTO"] = pd.to_numeric(tabla_antes.get("MONTO", 0), errors="coerce").fillna(0)
        tabla_nuevo["MONTO"] = pd.to_numeric(tabla_nuevo.get("MONTO", 0), errors="coerce").fillna(0)

        # Crear datounico y cuenta
        tabla_antes["datounico"] = (
            tabla_antes["EXPEDIENTE / CASO"].astype(str) + "-" +
            tabla_antes["NUM_DOC_DEMANDANTE"].astype(str) + "-" +
            tabla_antes["DEMANDANTE_NOMBRE"].astype(str)
        )
        tabla_antes["cuenta"] = tabla_antes["MAYOR"].astype(str) + "-" + tabla_antes["SUB_CTA"].astype(str)

        tabla_nuevo["datounico"] = (
            tabla_nuevo["EXPEDIENTE / CASO"].astype(str) + "-" +
            tabla_nuevo["NUM_DOC_DEMANDANTE"].astype(str) + "-" +
            tabla_nuevo["DEMANDANTE_NOMBRE"].astype(str)
        )
        tabla_nuevo["cuenta"] = tabla_nuevo["MAYOR"].astype(str) + "-" + tabla_nuevo["SUB_CTA"].astype(str)

        # Función para comparar cuentas según prefijo
        def comparar(tabla_antes, tabla_nuevo, prefijo):
            resultados = []
            antes_pref = tabla_antes[tabla_antes["cuenta"].str.startswith(prefijo)]
            nuevo_pref = tabla_nuevo[tabla_nuevo["cuenta"].str.startswith(prefijo)]
            
            for _, row in antes_pref.iterrows():
                datounico = row["datounico"]
                cuenta_antes = row["cuenta"]
                monto_antes = row["MONTO"]
                match = nuevo_pref[nuevo_pref["datounico"] == datounico]

                if not match.empty:
                    cuentas_nuevas = match["cuenta"].unique()
                    if cuenta_antes in cuentas_nuevas:
                        monto_nuevo = match.loc[match["cuenta"] == cuenta_antes, "MONTO"].sum()
                        diferencia = monto_nuevo - monto_antes
                        resultados.append([datounico, cuenta_antes, cuenta_antes, monto_antes, monto_nuevo, diferencia, "Misma cuenta"])
                    else:
                        monto_total = match["MONTO"].sum()
                        resultados.append([datounico, cuenta_antes, ", ".join(cuentas_nuevas), monto_antes, monto_total, None, "Cuenta diferente"])
                else:
                    resultados.append([datounico, cuenta_antes, "-", monto_antes, 0, -monto_antes, "Solo en ANTES"])

            for _, row in nuevo_pref.iterrows():
                datounico = row["datounico"]
                cuenta_nueva = row["cuenta"]
                monto_nuevo = row["MONTO"]
                if datounico not in antes_pref["datounico"].values:
                    resultados.append([datounico, "-", cuenta_nueva, 0, monto_nuevo, monto_nuevo, "Solo en MES NUEVO"])

            df_comp = pd.DataFrame(resultados, columns=["datounico", "Cuenta_ANTES", "Cuenta_MES_NUEVO", "MONTO_ANTES", "MONTO_MES_NUEVO", "Diferencia", "Resultado"])
            return df_comp

        # Comparaciones por prefijo
        comparaciones = {
            "Comparación 1202": comparar(tabla_antes, tabla_nuevo, "1202"),
            "Comparación 9110": comparar(tabla_antes, tabla_nuevo, "9110"),
            "Comparación 2401": comparar(tabla_antes, tabla_nuevo, "2401"),
            "Comparación 2103": comparar(tabla_antes, tabla_nuevo, "2103")
        }

        # Exportar a Excel en memoria
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            tabla_antes.to_excel(writer, index=False, sheet_name="ANTES")
            tabla_nuevo.to_excel(writer, index=False, sheet_name="MES NUEVO")
            for nombre, df_comp in comparaciones.items():
                df_comp.to_excel(writer, index=False, sheet_name=nombre)
        output.seek(0)

        return output, comparaciones

    except Exception as e:
        print(f"Error al procesar OA-2: {e}")
        return None, None
