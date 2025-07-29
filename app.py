
import streamlit as st
import pandas as pd
import re
from datetime import datetime, date
from pytz import timezone
from io import BytesIO

st.set_page_config(page_title="Revisión de Desvíos", layout="wide")
st.title("🚍 Revisión de Desvíos Operativos")

uploaded_desvios = st.file_uploader("📂 Cargar archivo de Desvíos", type=["xlsx"])
uploaded_pmt = st.file_uploader("📂 Cargar archivo de PMT", type=["xlsx"])

if uploaded_desvios and uploaded_pmt:
    df = pd.read_excel(uploaded_desvios, skiprows=1)
    pmt_df = pd.read_excel(uploaded_pmt)

    df.columns = [
       'Fecha', 'Instante', 'Línea', 'Coche', 'Código Bus', 'Nº SAE Bus',
       'Acción', 'Descripción Acción', 'Usuario', 'Nombre Usuario', 'Puesto',
       'Parámetros', 'Motivo', 'Descripción Motivo', 'Otra Columna', 'RUTA', 'ZONA'
    ]
    df = df[df["Descripción Acción"] == "Desvio"]
    df["Ruta"] = df["RUTA"].astype(str).str.strip()
    df["Zona"] = df["ZONA"].astype(str).str.strip()

    df["Estado Desvío"] = df["Parámetros"].apply(
        lambda x: "Activo" if isinstance(x, str) and ('Activar="SI"' in x or 'Activo="SI"' in x) else "Inactivo"
    )

    def extraer_codigo(param):
        if isinstance(param, str):
            match = re.search(r'Desvio="(\d+)"', param)
            if match:
                return match.group(1)
        return None

    df["Código Desvío"] = df["Parámetros"].apply(extraer_codigo)

    df["Instante"] = pd.to_datetime(df["Fecha"].astype(str) + " " + df["Instante"].astype(str), errors='coerce')
    df["Fecha Instante"] = df["Instante"].dt.date
    df["Hora Instante"] = df["Instante"].dt.strftime("%H:%M:%S")

    def evaluar_estado(grupo):
        cantidad = len(grupo)
        estados = grupo["Estado Desvío"].unique()
        if cantidad == 1:
            return grupo.iloc[0]["Estado Desvío"]
        elif cantidad == 2:
            return grupo.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]
        else:
            if "Activo" in estados and "Inactivo" in estados:
                return "Modificado"
            elif "Activo" in estados:
                return "Activo"
            else:
                return "Inactivo"

    estado_final = df.groupby("Código Desvío", group_keys=False).apply(evaluar_estado).reset_index()
    estado_final.columns = ["Código Desvío", "Estado Final"]

    conteo = df["Código Desvío"].value_counts().reset_index()
    conteo.columns = ["Código Desvío", "Cantidad"]

    df_final = pd.merge(df, estado_final, on="Código Desvío", how="left")
    df_final = pd.merge(df_final, conteo, on="Código Desvío", how="left")

    pmt_df["ID"] = pmt_df["ID"].astype(str).str.strip()
    pmt_id = pmt_df["ID"].tolist()
    df_final["Pmt o Desvíos Nuevos"] = df_final["Código Desvío"].apply(
        lambda x: "PMT" if str(x) in pmt_id else "Desvío Nuevo"
    )

    def ultimo_estado(grupo):
        return grupo.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]

    estado_mas_reciente = df_final.groupby("Código Desvío").apply(ultimo_estado).reset_index()
    estado_mas_reciente.columns = ["Código Desvío", "Estados"]

    df_final = pd.merge(df_final, estado_mas_reciente, on="Código Desvío", how="left")
    df_final["Revisión"] = df_final["Estados"].replace({
        "Activo": "Revisar",
        "Inactivo": "No Revisar"
    })

    def calcular_duracion_fila(instante):
        if pd.notnull(instante):
            ahora = datetime.now(timezone("America/Bogota")).replace(tzinfo=None)
            return ahora - instante
        return pd.NaT

    df_final["Duración Activo"] = df_final["Instante"].apply(calcular_duracion_fila)

    def formato_duracion(tiempo):
        if pd.isnull(tiempo):
            return ""
        total_segundos = int(tiempo.total_seconds())
        horas = total_segundos // 3600
        minutos = (total_segundos % 3600) // 60
        if horas > 0 and minutos > 0:
            return f"{horas} horas {minutos} minutos"
        elif horas > 0:
            return f"{horas} horas"
        elif minutos > 0:
            return f"{minutos} minutos"
        else:
            return "Menos de 1 minuto"

    df_final["Duración Activo"] = df_final["Duración Activo"].apply(formato_duracion)

    df_final = df_final[[  
        "Fecha Instante", "Hora Instante", "Nombre Usuario", "Código Desvío",
        "Estado Desvío", "Estado Final", "Cantidad", "Ruta", "Zona",
        "Pmt o Desvíos Nuevos", "Estados", "Revisión", "Duración Activo"
    ]]

    fecha_actual = date.today().strftime("%Y-%m-%d")
    output = BytesIO()
    df_final.to_excel(output, index=False)
    output.seek(0)

    st.success("✅ Archivo procesado correctamente. Puedes descargarlo abajo.")
    st.download_button(
        label="📥 Descargar archivo Excel",
        data=output,
        file_name=f"Revision de desvios {fecha_actual}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
