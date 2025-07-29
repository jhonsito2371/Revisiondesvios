import streamlit as st
import pandas as pd
import re
from datetime import datetime, date
from pytz import timezone
from io import BytesIO
import matplotlib.pyplot as plt
import seaborn as sns

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

# ---------- LECTURA ----------
def leer_desvios(file):
    for skip in [1, 0]:
        try:
            raw = pd.read_excel(file, skiprows=skip, engine="openpyxl")
            break
        except:
            continue
    if raw.shape[1] in (16, 17):
        cols = [
            'Fecha', 'Instante', 'Línea', 'Coche', 'Código Bus', 'Nº SAE Bus', 'Acción',
            'Descripción Acción', 'Usuario', 'Nombre Usuario', 'Puesto', 'Parámetros',
            'Motivo', 'Descripción Motivo', 'Otra Columna', 'RUTA']
        if raw.shape[1] == 17:
            cols += ['ZONA']
        raw.columns = cols
        if 'ZONA' not in raw.columns:
            raw['ZONA'] = ""
    return raw

try:
    df_raw = leer_desvios(f_desv)
except Exception as e:
    st.error("No se pudo leer el archivo de desvíos.")
    st.stop()

if "Descripción Acción" in df_raw.columns:
    df = df_raw[df_raw["Descripción Acción"].str.lower().str.strip() == "desvio"].copy()
else:
    df = df_raw.copy()

if "RUTA" in df.columns: df["Ruta"] = df["RUTA"].astype(str).str.strip()
if "ZONA" in df.columns: df["Zona"] = df["ZONA"].astype(str).str.strip()
if "Ruta" not in df.columns: df["Ruta"] = ""
if "Zona" not in df.columns: df["Zona"] = ""

df["Estado Desvío"] = df["Parámetros"].apply(lambda x:
    "Activo" if isinstance(x, str) and any(k in x.upper() for k in ["ACTIVAR=\"SI\"", "ACTIVO=\"SI\""]) else "Inactivo")

def extraer_codigo(param):
    if isinstance(param, str):
        m = re.search(r'Desvio="(\d+)?"', param)
        if m: return m.group(1)
    return None

df["Código Desvío"] = df["Parámetros"].apply(extraer_codigo)
df["Instante"] = pd.to_datetime(df["Fecha"].astype(str) + " " + df["Instante"].astype(str), errors="coerce")
df["Fecha Instante"] = df["Instante"].dt.date
df["Hora Instante"] = df["Instante"].dt.strftime("%H:%M:%S")

# --------- ESTADO FINAL ---------
def evaluar_estado(grupo):
    estados = grupo["Estado Desvío"].unique()
    if len(grupo) == 1:
        return grupo.iloc[0]["Estado Desvío"]
    if len(grupo) == 2:
        return grupo.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]
    if "Activo" in estados and "Inactivo" in estados:
        return "Modificado"
    return estados[0]

df_final = df.copy()
df_final["Estado Final"] = df.groupby("Código Desvío", group_keys=False).apply(evaluar_estado).reindex(df.index).values
df_final["Cantidad"] = df_final.groupby("Código Desvío")["Código Desvío"].transform("count")

# -------- ESTADO RECIENTE --------
def ultimo_estado(grupo):
    return grupo.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]
df_final["Estados"] = df_final.groupby("Código Desvío")["Estado Desvío"].transform(ultimo_estado)
df_final["Revisión"] = df_final["Estados"].replace({"Activo": "Revisar", "Inactivo": "No Revisar"})

# -------- DURACIÓN --------
def calc_duracion_fila(instante):
    ahora = datetime.now(timezone("America/Bogota")).replace(tzinfo=None)
    return ahora - instante if pd.notnull(instante) else pd.NaT

def formato_duracion(td):
    if pd.isnull(td): return ""
    total = int(td.total_seconds())
    h, m = divmod(total // 60, 60)
    return f"{h} horas {m} minutos" if h or m else "Menos de 1 minuto"

df_final["Duración Activo"] = df_final["Instante"].apply(calc_duracion_fila).apply(formato_duracion)

# -------- CRUCE PMT --------
if f_pmt:
    try:
        pmt_df = pd.read_excel(f_pmt, engine="openpyxl")
        if "ID" in pmt_df.columns:
            pmt_ids = pmt_df["ID"].astype(str).str.strip().tolist()
            df_final["Pmt o Desvíos Nuevos"] = df_final["Código Desvío"].apply(
                lambda x: "PMT" if str(x) in pmt_ids else "Desvío Nuevo")
        else:
            df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"
    except:
        df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"
else:
    df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"

# ----------- FILTROS -----------
with st.expander("🔍 Filtros:"):
    zona = st.multiselect("Zona", options=sorted(df_final["Zona"].dropna().unique()))
    ruta = st.multiselect("Ruta", options=sorted(df_final["Ruta"].dropna().unique()))
    estado = st.multiselect("Estado Final", options=sorted(df_final["Estado Final"].dropna().unique()))

    mask = (
        (df_final["Zona"].isin(zona) if zona else True) &
        (df_final["Ruta"].isin(ruta) if ruta else True) &
        (df_final["Estado Final"].isin(estado) if estado else True)
    )
    df_final = df_final[mask]

# ----------- GRÁFICOS -----------
st.subheader("📊 Resumen Visual")
col1, col2 = st.columns(2)
with col1:
    fig1, ax1 = plt.subplots()
    df_final["Estado Final"].value_counts().plot(kind='bar', ax=ax1, color="skyblue")
    ax1.set_title("Desvíos por Estado Final")
    st.pyplot(fig1)

with col2:
    fig2, ax2 = plt.subplots()
    df_final["Zona"].value_counts().plot(kind='barh', ax=ax2, color="lightgreen")
    ax2.set_title("Cantidad por Zona")
    st.pyplot(fig2)

# ----------- EXPORTAR -----------
cols_finales = [
    "Fecha Instante", "Hora Instante", "Nombre Usuario", "Código Desvío",
    "Estado Desvío", "Estado Final", "Cantidad", "Ruta", "Zona",
    "Pmt o Desvíos Nuevos", "Estados", "Revisión", "Duración Activo"]

df_final = df_final[[c for c in cols_finales if c in df_final.columns]]
st.success("✅ Procesado con éxito. Vista previa:")
st.dataframe(df_final, use_container_width=True)

buffer = BytesIO()
df_final.to_excel(buffer, index=False)
buffer.seek(0)
st.download_button(
    "📅 Descargar Excel final",
    data=buffer,
    file_name=f"Revision de desvios {date.today().strftime('%Y-%m-%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
