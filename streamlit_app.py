import streamlit as st
import pandas as pd
import datetime

st.set_page_config(page_title="Revisión de Desvíos", page_icon="🚍", layout="wide")

st.title("📊 Revisión de Desvíos Activos e Inactivos")
st.markdown("Sube el archivo Excel con los datos de desvíos para generar el informe consolidado.")

archivo = st.file_uploader("📁 Cargar archivo Excel", type=["xlsx"])

if archivo is not None:
    df = pd.read_excel(archivo)

    # Convertir Hora Instante a datetime
    df["Hora Instante"] = pd.to_datetime(df["Hora Instante"], errors="coerce")
    df["Hora Instante"] = df["Hora Instante"].dt.time

    # Extraer Fecha
    df["Fecha Instante"] = pd.to_datetime(df["Fecha Instante"], errors="coerce").dt.date

    # Determinar el estado más reciente por Código Desvío
    df_sorted = df.sort_values(by=["Código de Desvío", "Hora Instante"], ascending=[True, False])
    df_sorted["Estado Reciente"] = df_sorted.groupby("Código de Desvío")["Estado"].transform("first")

    # Clasificación final
    def definir_estado_final(grupo):
        estados = grupo["Estado"].unique()
        if len(estados) > 1:
            return "Modificado"
        else:
            return estados[0]

    df_estado = df.groupby("Código de Desvío").apply(definir_estado_final).reset_index(name="Estado Final")

    # Revisión (uno por código con estado más reciente)
    df_revision = df_sorted.drop_duplicates(subset=["Código de Desvío"], keep="first")
    df_revision["Revisión"] = df_revision["Estado Reciente"].replace({"Activo": "Revisar", "Inactivo": "No Revisar"})

    # Contar repeticiones
    df_conteo = df["Código de Desvío"].value_counts().reset_index()
    df_conteo.columns = ["Código de Desvío", "Cantidad"]

    # Unir todo
    final = df_sorted.merge(df_estado, on="Código de Desvío", how="left")
    final = final.merge(df_conteo, on="Código de Desvío", how="left")
    final = final.merge(df_revision[["Código de Desvío", "Revisión"]], on="Código de Desvío", how="left")

    # Reordenar columnas
    columnas_orden = ["Fecha Instante", "Hora Instante", "Nombre del Usuario", "Código de Desvío",
                      "Estado", "Estado Final", "Cantidad", "Revisión"]
    columnas_existentes = [col for col in columnas_orden if col in final.columns]
    columnas_restantes = [col for col in final.columns if col not in columnas_existentes]
    final = final[columnas_existentes + columnas_restantes]

    # Mostrar tabla
    st.success("✅ Archivo procesado correctamente.")
    st.dataframe(final, use_container_width=True)

    # Descargar
    nombre = f"Revision de desvios {datetime.datetime.now().date()}.xlsx"
    final_excel = final.copy()
    final_excel.to_excel(nombre, index=False)

    with open(nombre, "rb") as f:
        st.download_button(
            label="📥 Descargar archivo procesado",
            data=f,
            file_name=nombre,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
 
