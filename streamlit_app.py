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
    try:
        df = pd.read_excel(file, skiprows=1)
    except:
        df = pd.read_excel(file)
    return df

try:
    df = leer_desvios(f_desv)
except Exception as e:
    st.error("❌ Error al leer archivo de desvíos")
    st.exception(e)
    st.stop()

# Ignorar columnas no deseadas
if "Unnamed: 14" in df.columns:
    df = df.drop(columns=["Unnamed: 14"])

# Normalizar columnas
df["Ruta"] = df.get("RUTA", "")
df["Zona"] = df.get("ZONA", "")

# Estado del desvío
df["Estado Desvío"] = df["Parámetros"].apply(
    lambda x: "Activo" if isinstance(x, str) and any(k in x for k in ['Activar="SI"', 'Activo="SI"', 'ACTIVAR="SI"', 'ACTIVO="SI"']) else "Inactivo"
)

# Extraer Código de Desvío
def extraer_codigo(param):
    if isinstance(param, str):
        m = re.search(r'Desvio="(\d+)"', param)
        if m:
            return m.group(1)
    return None

df["Código Desvío"] = df["Parámetros"].apply(extraer_codigo)
df = df[df["Código Desvío"].notna()].copy()

# Instante como datetime
df["Instante"] = pd.to_datetime(df["Fecha"].astype(str) + " " + df["Instante"].astype(str), errors="coerce")
df["Fecha Instante"] = df["Instante"].dt.date
df["Hora Instante"] = df["Instante"].dt.strftime("%H:%M:%S")

# Estado Final
def estado_final(grupo):
    if len(grupo) == 1:
        return grupo.iloc[0]["Estado Desvío"]
    elif len(grupo) == 2:
        return grupo.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]
    else:
        estados = grupo["Estado Desvío"].unique()
        if "Activo" in estados and "Inactivo" in estados:
            return "Modificado"
        elif "Activo" in estados:
            return "Activo"
        else:
            return "Inactivo"

df_estado = df.groupby("Código Desvío", group_keys=False).apply(estado_final).reset_index()
df_estado.columns = ["Código Desvío", "Estado Final"]

# Conteo de repeticiones
df_conteo = df["Código Desvío"].value_counts().rename_axis("Código Desvío").reset_index(name="Cantidad")

# Unión
df_final = pd.merge(df, df_estado, on="Código Desvío", how="left")
df_final = pd.merge(df_final, df_conteo, on="Código Desvío", how="left")

# Estado reciente y revisión
estado_mas_reciente = df_final.groupby("Código Desvío").apply(
    lambda g: g.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]
).reset_index(name="Estados")
df_final = pd.merge(df_final, estado_mas_reciente, on="Código Desvío", how="left")
df_final["Revisión"] = df_final["Estados"].replace({"Activo": "Revisar", "Inactivo": "No Revisar"})

# Duración
df_final["Duración Activo"] = df_final["Instante"].apply(
    lambda x: datetime.now(timezone("America/Bogota")).replace(tzinfo=None) - x if pd.notnull(x) else pd.NaT
)
df_final["Duración Activo"] = df_final["Duración Activo"].apply(
    lambda td: f"{td.seconds // 3600} horas {((td.seconds % 3600) // 60)} minutos" if pd.notnull(td) else ""
)

# Cruce con PMT
if f_pmt:
    try:
        df_pmt = pd.read_excel(f_pmt, engine="openpyxl")
        if "ID" in df_pmt.columns:
            pmt_ids = df_pmt["ID"].astype(str).str.strip().tolist()
            df_final["Pmt o Desvíos Nuevos"] = df_final["Código Desvío"].apply(lambda x: "PMT" if str(x) in pmt_ids else "Desvío Nuevo")
        else:
            df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"
    except:
        df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"
else:
    df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"

# Columnas finales
cols_finales = [
    "Fecha Instante", "Hora Instante", "Nombre Usuario", "Código Desvío", "Estado Desvío",
    "Estado Final", "Cantidad", "Ruta", "Zona", "Pmt o Desvíos Nuevos", "Estados", "Revisión", "Duración Activo"
]
df_final = df_final[[c for c in cols_finales if c in df_final.columns]]

# Filtros

with st.expander("🔎 Filtros"):
    rutas = df_final["Ruta"].dropna().unique().tolist()
    zonas = df_final["Zona"].dropna().unique().tolist()
    estados = df_final["Estado Final"].dropna().unique().tolist()
    tipos_desvio = df_final["Pmt o Desvíos Nuevos"].dropna().unique().tolist()
    rev_opciones = df_final["Revisión"].dropna().unique().tolist()

    sel_ruta = st.multiselect("Filtrar por Ruta", rutas, default=rutas, key="filtro_ruta")
    sel_zona = st.multiselect("Filtrar por Zona", zonas, default=zonas, key="filtro_zona")
    sel_estado = st.multiselect("Filtrar por Estado Final", estados, default=estados, key="filtro_estado")
    sel_tipo = st.multiselect("Filtrar por Tipo de Desvío", tipos_desvio, default=tipos_desvio, key="filtro_tipo")
    sel_rev = st.multiselect("Filtrar por Revisión", rev_opciones, default=rev_opciones, key="filtro_rev")

    df_final = df_final[
        df_final["Ruta"].isin(sel_ruta) &
        df_final["Zona"].isin(sel_zona) &
        df_final["Estado Final"].isin(sel_estado) &
        df_final["Pmt o Desvíos Nuevos"].isin(sel_tipo) &
        df_final["Revisión"].isin(sel_rev)
    ]

    rutas = df_final["Ruta"].unique().tolist()
    zonas = df_final["Zona"].unique().tolist()
    estados = df_final["Estado Final"].unique().tolist()

    sel_ruta = st.multiselect("Filtrar por Ruta", rutas, default=rutas, key="filtro_ruta")
    sel_zona = st.multiselect("Filtrar por Zona", zonas, default=zonas, key="filtro_zona")
    sel_estado = st.multiselect("Filtrar por Estado Final", estados, default=estados, key="filtro_estado")

    df_final = df_final[
        df_final["Ruta"].isin(sel_ruta) &
        df_final["Zona"].isin(sel_zona) &
        df_final["Estado Final"].isin(sel_estado)
    ]

# Vista previa
st.success("✅ Procesado con éxito. Vista previa:")
st.dataframe(df_final, use_container_width=True)

# Gráficas
col1, col2 = st.columns(2)
with col1:
    fig1 = px.pie(df_final, names="Estado Final", title="Distribución de Estado Final")
    st.plotly_chart(fig1, use_container_width=True)
with col2:
    conteo_revision = df_final["Revisión"].value_counts(dropna=True).reset_index()
conteo_revision.columns = ["Estado Revisión", "Cantidad"]
fig2 = px.bar(conteo_revision, x="Estado Revisión", y="Cantidad", title="Revisión por Estado")
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


