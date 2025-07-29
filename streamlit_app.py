# streamlit_app.py
import streamlit as st
import pandas as pd
import re
from datetime import datetime, date
from pytz import timezone
from io import BytesIO
import plotly.express as px

st.set_page_config(page_title="Revisión de Desvíos", page_icon="🚍", layout="wide")
st.title("🚍 Revisión de Desvíos Operativos")
st.markdown("Sube el archivo **de desvíos (acciones)** y la **base PMT**. El sistema detecta el formato y ajusta encabezados automáticamente.")

# ---------- CARGA DE ARCHIVOS ----------
col1, col2 = st.columns(2)
with col1:
    f_desv = st.file_uploader("📂 Archivo de Desvíos (acciones .xlsx)", type=["xlsx"], key="desv")
with col2:
    f_pmt = st.file_uploader("📂 Base PMT (.xlsx)", type=["xlsx"], key="pmt")

if not f_desv:
    st.info("👈 Sube al menos el **archivo de desvíos** para comenzar.")
    st.stop()

# ---------- LECTURA Y LIMPIEZA ----------
def leer_desvios(file):
    for opts in [{"skiprows": 1}, {"skiprows": 0}]:
        try:
            df = pd.read_excel(file, engine="openpyxl", **opts)
            break
        except: continue
    if df.shape[1] in (16, 17):
        columnas = ['Fecha', 'Instante', 'Línea', 'Coche', 'Código Bus', 'Nº SAE Bus',
                    'Acción', 'Descripción Acción', 'Usuario', 'Nombre Usuario', 'Puesto',
                    'Parámetros', 'Motivo', 'Descripción Motivo', 'Otra Columna', 'RUTA']
        if df.shape[1] == 17:
            columnas.append('ZONA')
        df.columns = columnas[:df.shape[1]]
        if 'ZONA' not in df.columns: df['ZONA'] = ""
    return df

df_raw = leer_desvios(f_desv)
if "Descripción Acción" in df_raw.columns:
    df = df_raw[df_raw["Descripción Acción"].str.lower().str.strip() == "desvio"].copy()
else:
    df = df_raw.copy()

df["Ruta"] = df.get("RUTA", "")
df["Zona"] = df.get("ZONA", "")
df["Estado Desvío"] = df["Parámetros"].apply(lambda x: "Activo" if isinstance(x, str) and any(s in x for s in ['Activar="SI"','Activo="SI"','ACTIVAR="SI"','ACTIVO="SI"']) else "Inactivo")

def extraer_codigo(p):
    if isinstance(p, str):
        m = re.search(r'Desvio="(\d+)"', p)
        return m.group(1) if m else None
    return None

df["Código Desvío"] = df["Parámetros"].apply(extraer_codigo)
df["Instante"] = pd.to_datetime(df["Fecha"].astype(str) + " " + df["Instante"].astype(str), errors="coerce")
df["Fecha Instante"] = df["Instante"].dt.date
df["Hora Instante"] = df["Instante"].dt.strftime("%H:%M:%S")

# ---------- ESTADOS Y REVISIÓN ----------
def evaluar_estado(grupo):
    cantidad = len(grupo)
    estados = grupo["Estado Desvío"].unique()
    if cantidad == 1: return grupo.iloc[0]["Estado Desvío"]
    elif cantidad == 2: return grupo.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]
    elif "Activo" in estados and "Inactivo" in estados: return "Modificado"
    return "Activo" if "Activo" in estados else "Inactivo"

estado_final = df.groupby("Código Desvío", group_keys=False).apply(evaluar_estado).reset_index()
estado_final.columns = ["Código Desvío", "Estado Final"]
conteo = df["Código Desvío"].value_counts().reset_index()
conteo.columns = ["Código Desvío", "Cantidad"]

estado_reciente = df.groupby("Código Desvío").apply(lambda g: g.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]).reset_index()
estado_reciente.columns = ["Código Desvío", "Estados"]

# ---------- DURACIÓN ----------
def calc_duracion(instante):
    if pd.notnull(instante):
        ahora = datetime.now(timezone("America/Bogota")).replace(tzinfo=None)
        dur = ahora - instante
        h, m = divmod(int(dur.total_seconds())//60, 60)
        return f"{h} horas {m} minutos" if h or m else "Menos de 1 minuto"
    return ""

df = df.merge(estado_final, on="Código Desvío", how="left")
df = df.merge(conteo, on="Código Desvío", how="left")
df = df.merge(estado_reciente, on="Código Desvío", how="left")
df["Revisión"] = df["Estados"].replace({"Activo": "Revisar", "Inactivo": "No Revisar"})
df["Duración Activo"] = df["Instante"].apply(calc_duracion)

# ---------- CRUCE PMT ----------
if f_pmt:
    try:
        pmt = pd.read_excel(f_pmt, engine="openpyxl")
        pmt_ids = pmt["ID"].astype(str).str.strip().tolist() if "ID" in pmt.columns else []
        df["Pmt o Desvíos Nuevos"] = df["Código Desvío"].apply(lambda x: "PMT" if str(x) in pmt_ids else "Desvío Nuevo")
    except:
        df["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"
else:
    df["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"

# ---------- FILTROS INTERACTIVOS ----------
with st.sidebar:
    st.header("🔍 Filtros")
    rutas = st.multiselect("Ruta", sorted(df["Ruta"].dropna().unique().tolist()))
    zonas = st.multiselect("Zona", sorted(df["Zona"].dropna().unique().tolist()))
    estados = st.multiselect("Estado Final", sorted(df["Estado Final"].dropna().unique().tolist()))

filtro = (
    (df["Ruta"].isin(rutas) if rutas else True) &
    (df["Zona"].isin(zonas) if zonas else True) &
    (df["Estado Final"].isin(estados) if estados else True)
)
df_filtrado = df[filtro]

# ---------- GRÁFICAS ----------
g1 = px.histogram(df_filtrado, x="Ruta", color="Estado Final", barmode="group", title="Desvíos por Ruta")
g2 = px.pie(df_filtrado, names="Zona", title="Distribución por Zona")

# ---------- RESULTADOS ----------
cols_final = [
    "Fecha Instante", "Hora Instante", "Nombre Usuario", "Código Desvío", "Estado Desvío",
    "Estado Final", "Cantidad", "Ruta", "Zona", "Pmt o Desvíos Nuevos", "Estados", "Revisión", "Duración Activo"
]
res = df_filtrado[[c for c in cols_final if c in df_filtrado.columns]].copy()
st.success(f"✅ {len(res)} registros filtrados.")
st.dataframe(res, use_container_width=True)
st.plotly_chart(g1, use_container_width=True)
st.plotly_chart(g2, use_container_width=True)

# ---------- DESCARGA ----------
buffer = BytesIO()
res.to_excel(buffer, index=False)
buffer.seek(0)
st.download_button(
    "📅 Descargar Excel filtrado",
    data=buffer,
    file_name=f"Revision de desvios {date.today()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


