import pandas as pd
from io import BytesIO

def procesar_oa2(file_antes, file_nuevo):
    """
    Procesa los archivos ANTES y MES NUEVO según pasos indicados.
    Devuelve un BytesIO con el Excel listo para descargar y un diccionario con las comparaciones.
    """
    try:
        # ===== Paso 1: Leer archivos Excel =====
        tabla_antes = pd.read_excel(file_antes, dtype=str)
        tabla_nuevo = pd.read_excel(file_nuevo, dtype=str)

        # Asegurarse que MONTO es numérico
        tabla_antes["MONTO"] = pd.to_numeric(tabla_antes.get("MONTO", 0), errors="coerce").fillna(0)
        tabla_nuevo["MONTO"] = pd.to_numeric(tabla_nuevo.get("MONTO", 0), errors="coerce").fillna(0)

        # ===== Paso 2 y 3: Crear datounico y cuenta, sumar MONTO si se repite =====
        def crear_tabla(df):
            df["datounico"] = (
                df["EXPEDIENTE / CASO"].astype(str) + "-" +
                df["NUM_DOC_DEMANDANTE"].astype(str) + "-" +
                df["DEMANDANTE_NOMBRE"].astype(str)
            )
            df["cuenta"] = df["MAYOR"].astype(str) + "-" + df["SUB_CTA"].astype(str)
            # Agrupar por datounico y cuenta, sumar MONTO
            df_agrupado = df.groupby(["datounico", "cuenta"], as_index=False)["MONTO"].sum()
            return df_agrupado

        df_antes = crear_tabla(tabla_antes)
        df_nuevo = crear_tabla(tabla_nuevo)

        # ===== Paso 4 a 7: Comparaciones por prefijo =====
        def comparar_por_prefijo(df_antes, df_nuevo, prefijo):
            resultados = []
            antes_pref = df_antes[df_antes["cuenta"].str.startswith(prefijo)]
            nuevo_pref = df_nuevo[df_nuevo["cuenta"].str.startswith(prefijo)]

            # Comparar filas existentes en ANTES
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

            # Comparar filas que solo existen en MES NUEVO
            for _, row in nuevo_pref.iterrows():
                datounico = row["datounico"]
                cuenta_nueva = row["cuenta"]
                monto_nuevo = row["MONTO"]
                if datounico not in antes_pref["datounico"].values:
                    resultados.append([datounico, "-", cuenta_nueva, 0, monto_nuevo, monto_nuevo, "Solo en MES NUEVO"])

            return pd.DataFrame(resultados, columns=["datounico", "Cuenta_ANTES", "Cuenta_MES_NUEVO",
                                                     "MONTO_ANTES", "MONTO_MES_NUEVO", "Diferencia", "Resultado"])

        # Prefijos indicados
        prefijos = ["1202", "9110", "2401", "2103"]
        comparaciones = {f"Comparación {p}": comparar_por_prefijo(df_antes, df_nuevo, p) for p in prefijos}

        # ===== Exportar a Excel =====
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_antes.to_excel(writer, index=False, sheet_name="ANTES")
            df_nuevo.to_excel(writer, index=False, sheet_name="MES NUEVO")
            for nombre, df_comp in comparaciones.items():
                df_comp.to_excel(writer, index=False, sheet_name=nombre)
        output.seek(0)

        return output, comparaciones

    except Exception as e:
        print(f"Error al procesar OA-2: {e}")
        return None, None
