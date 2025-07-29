# streamlit_app.py

import streamlit as st
import pandas as pd
import re
from datetime import datetime, date
from pytz import timezone
from io import BytesIO
import plotly.express as px
import smtplib
from email.message import EmailMessage

st.set_page_config(page_title="Revisión de Desvíos", page_icon="🚍", layout="wide")
st.title("🚍 Revisión de Desvíos Operativos")

st.markdown("Sube el archivo **de desvíos (acciones)** y la **base PMT**. El sistema detecta el formato y ajusta encabezados automáticamente.")

col1, col2 = st.columns(2)
with col1:
    f_desv = st.file_uploader("📂 Archivo de Desvíos (acciones .xlsx)", type=["xlsx"], key="desv")
with col2:
    f_pmt = st.file_uploader("📂 Base PMT (.xlsx)", type=["xlsx"], key="pmt")

if not f_desv:
    st.info("👈 Sube al menos el **archivo de desvíos** para comenzar.")
    st.stop()

def leer_desvios(file):
    try_orders = [{"skiprows": 1}, {"skiprows": 0}]
    raw = None
    for opts in try_orders:
        try:
            raw = pd.read_excel(file, engine="openpyxl", **opts)
            break
        except:
            continue

    if raw.shape[1] in (16, 17):
        cols16 = ['Fecha', 'Instante', 'Línea', 'Coche', 'Código Bus', 'Nº SAE Bus', 'Acción', 'Descripción Acción', 'Usuario', 'Nombre Usuario', 'Puesto', 'Parámetros', 'Motivo', 'Descripción Motivo', 'Otra Columna', 'RUTA']
        cols17 = cols16 + ['ZONA']
        raw.columns = cols16 if raw.shape[1] == 16 else cols17
        if raw.shape[1] == 16:
            raw["ZONA"] = ""
    return raw

try:
    df_raw = leer_desvios(f_desv)
except Exception as e:
    st.error("Error al leer archivo de desvíos")
    st.exception(e)
    st.stop()

if "Descripción Acción" in df_raw.columns:
    df = df_raw[df_raw["Descripción Acción"].astype(str).str.strip().str.lower() == "desvio"].copy()
else:
    df = df_raw.copy()

df["Ruta"] = df.get("RUTA", "")
df["Zona"] = df.get("ZONA", "")

df["Estado Desvío"] = df["Parámetros"].apply(
    lambda x: "Activo" if isinstance(x, str) and any(k in x for k in ['Activar="SI"', 'Activo="SI"', 'ACTIVAR="SI"', 'ACTIVO="SI"']) else "Inactivo"
)

def extraer_codigo(param):
    if isinstance(param, str):
        m = re.search(r'Desvio="(\d+)"', param)
        if m:
            return m.group(1)
    return None

df["Código Desvío"] = df["Parámetros"].apply(extraer_codigo)

df["Instante"] = pd.to_datetime(df["Fecha"].astype(str) + " " + df["Instante"].astype(str), errors="coerce")
df["Fecha Instante"] = df["Instante"].dt.date
df["Hora Instante"] = df["Instante"].dt.strftime("%H:%M:%S")

# Agrupar por Código Desvío
estado_final = df.groupby("Código Desvío", group_keys=False).apply(
    lambda g: g.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]
    if len(g) == 2 else (
        "Modificado" if ("Activo" in g["Estado Desvío"].values and "Inactivo" in g["Estado Desvío"].values)
        else g.iloc[0]["Estado Desvío"]
    )
)
estado_final = estado_final.reset_index(name="Estado Final")

conteo = df["Código Desvío"].value_counts().reset_index().rename(columns={"index": "Código Desvío", "Código Desvío": "Cantidad"})
df_final = pd.merge(df, estado_final, on="Código Desvío", how="left")
df_final = pd.merge(df_final, conteo, on="Código Desvío", how="left")

estado_mas_reciente = df_final.groupby("Código Desvío").apply(
    lambda g: g.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]
).reset_index(name="Estados")
df_final = pd.merge(df_final, estado_mas_reciente, on="Código Desvío", how="left")
df_final["Revisión"] = df_final["Estados"].replace({"Activo": "Revisar", "Inactivo": "No Revisar"})

df_final["Duración Activo"] = df_final["Instante"].apply(
    lambda x: datetime.now(timezone("America/Bogota")).replace(tzinfo=None) - x if pd.notnull(x) else pd.NaT
)
df_final["Duración Activo"] = df_final["Duración Activo"].apply(
    lambda td: f"{td.seconds // 3600} horas {((td.seconds % 3600) // 60)} minutos" if pd.notnull(td) else ""
)

if f_pmt:
    try:
        pmt_df = pd.read_excel(f_pmt, engine="openpyxl")
        pmt_ids = pmt_df["ID"].astype(str).str.strip().tolist() if "ID" in pmt_df.columns else []
        df_final["Pmt o Desvíos Nuevos"] = df_final["Código Desvío"].apply(lambda x: "PMT" if str(x) in pmt_ids else "Desvío Nuevo")
    except:
        df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"
else:
    df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"

cols_finales = [
    "Fecha Instante", "Hora Instante", "Nombre Usuario", "Código Desvío", "Estado Desvío",
    "Estado Final", "Cantidad", "Ruta", "Zona", "Pmt o Desvíos Nuevos", "Estados", "Revisión", "Duración Activo"
]
df_final = df_final[[c for c in cols_finales if c in df_final.columns]]

with st.expander("🔎 Filtros"):
    rutas = df_final["Ruta"].unique().tolist()
    zonas = df_final["Zona"].unique().tolist()
    estados = df_final["Estado Final"].unique().tolist()

    sel_ruta = st.multiselect("Filtrar por Ruta", rutas, default=rutas)
    sel_zona = st.multiselect("Filtrar por Zona", zonas, default=zonas)
    sel_estado = st.multiselect("Filtrar por Estado Final", estados, default=estados)

    df_final = df_final[
        df_final["Ruta"].isin(sel_ruta) &
        df_final["Zona"].isin(sel_zona) &
        df_final["Estado Final"].isin(sel_estado)
    ]

st.success("✅ Procesado con éxito. Vista previa:")
st.dataframe(df_final, use_container_width=True)

# Gráficas
col1, col2 = st.columns(2)
with col1:
    fig1 = px.pie(df_final, names="Estado Final", title="Distribución de Estado Final")
    st.plotly_chart(fig1, use_container_width=True)
with col2:
    fig2 = px.bar(df_final["Revisión"].value_counts().reset_index(),
                  x="index", y="Revisión", title="Revisión por Estado")
    st.plotly_chart(fig2, use_container_width=True)

# Descargar Excel
buffer = BytesIO()
df_final.to_excel(buffer, index=False)
buffer.seek(0)
st.download_button(
    "📥 Descargar Excel final",
    data=buffer,
    file_name=f"Revision de desvios {date.today().strftime('%Y-%m-%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Enviar por correo
with st.expander("📧 Enviar por correo"):
    correo = st.text_input("Correo destino")
    if st.button("Enviar reporte"):
        if correo:
            msg = EmailMessage()
            msg['Subject'] = 'Reporte de Desvíos Operativos'
            msg['From'] = st.secrets["HOTMAIL_USER"]
            msg['To'] = correo
            msg.set_content("Adjunto encontrará el reporte generado desde la aplicación Streamlit.")
            msg.add_attachment(buffer.getvalue(), maintype='application',
                               subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               filename="Reporte_Desvios.xlsx")
            try:
                with smtplib.SMTP('smtp.office365.com', 587) as smtp:
                    smtp.starttls()
                    smtp.login(st.secrets["HOTMAIL_USER"], st.secrets["HOTMAIL_PASS"])
                    smtp.send_message(msg)
                st.success("📨 Correo enviado con éxito")
            except Exception as e:
                st.error("Error al enviar correo")
                st.exception(e)
        else:
            st.warning("Debe ingresar un correo válido")

