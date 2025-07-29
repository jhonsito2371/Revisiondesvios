import streamlit as st
import pandas as pd
import re
from datetime import datetime, date
from pytz import timezone
from io import BytesIO
import matplotlib.pyplot as plt

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
    raw, last_err = None, None
    for opts in try_orders:
        try:
            raw = pd.read_excel(file, engine="openpyxl", **opts)
            break
        except Exception as e:
            last_err = e
    if raw is None:
        raise last_err

    if raw.shape[1] in (16, 17):
        cols16 = ['Fecha', 'Instante', 'Línea', 'Coche', 'Código Bus', 'Nº SAE Bus',
                  'Acción', 'Descripción Acción', 'Usuario', 'Nombre Usuario', 'Puesto',
                  'Parámetros', 'Motivo', 'Descripción Motivo', 'Otra Columna', 'RUTA']
        cols17 = cols16 + ['ZONA']
        raw.columns = cols16 if raw.shape[1] == 16 else cols17
        if raw.shape[1] == 16:
            raw["ZONA"] = ""
    return raw

try:
    df_raw = leer_desvios(f_desv)
except Exception as e:
    st.error("❌ No se pudo leer el archivo de desvíos. Verifica que sea un .xlsx válido.")
    st.exception(e)
    st.stop()

if "Descripción Acción" in df_raw.columns:
    df = df_raw[df_raw["Descripción Acción"].astype(str).str.strip().str.lower() == "desvio"].copy()
else:
    df = df_raw.copy()

df["Ruta"] = df.get("RUTA", "").astype(str).str.strip()
df["Zona"] = df.get("ZONA", "").astype(str).str.strip()

df["Estado Desvío"] = df["Parámetros"].apply(
    lambda x: "Activo" if isinstance(x, str) and ('Activar="SI"' in x or 'Activo="SI"' in x or 'ACTIVAR="SI"' in x or 'ACTIVO="SI"' in x) else "Inactivo"
)

def extraer_codigo(param):
    if isinstance(param, str):
        m = re.search(r'Desvio="(\d+)"', param)
        return m.group(1) if m else None
    return None

df["Código Desvío"] = df["Parámetros"].apply(extraer_codigo)

df["Instante"] = pd.to_datetime(df["Fecha"].astype(str) + " " + df["Instante"].astype(str), errors="coerce")
df["Fecha Instante"] = df["Instante"].dt.date

def evaluar_estado(grupo):
    estados = grupo["Estado Desvío"].unique()
    if len(grupo) == 1:
        return grupo.iloc[0]["Estado Desvío"]
    elif len(grupo) == 2:
        return grupo.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]
    if "Activo" in estados and "Inactivo" in estados:
        return "Modificado"
    return "Activo" if "Activo" in estados else "Inactivo"

estado_final = df.groupby("Código Desvío", group_keys=False).apply(evaluar_estado).reset_index()
estado_final.columns = ["Código Desvío", "Estado Final"]
conteo = df["Código Desvío"].value_counts().reset_index()
conteo.columns = ["Código Desvío", "Cantidad"]

ultimo_estado = df.groupby("Código Desvío").apply(lambda g: g.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]).reset_index()
ultimo_estado.columns = ["Código Desvío", "Estados"]

revisiones = ultimo_estado["Estados"].replace({"Activo": "Revisar", "Inactivo": "No Revisar"})
ultimo_estado["Revisión"] = revisiones

ahora = datetime.now(timezone("America/Bogota")).replace(tzinfo=None)
df["Duración Activo"] = df["Instante"].apply(lambda x: ahora - x if pd.notnull(x) else None)

def formato_duracion(td):
    if pd.isnull(td): return ""
    total = int(td.total_seconds())
    h, m = divmod(total // 60, 60)
    return f"{h}h {m}m" if h else f"{m} minutos"

df["Duración Activo"] = df["Duración Activo"].apply(formato_duracion)

df_final = df.merge(estado_final, on="Código Desvío", how="left")
df_final = df_final.merge(conteo, on="Código Desvío", how="left")
df_final = df_final.merge(ultimo_estado, on="Código Desvío", how="left")

if f_pmt:
    try:
        pmt_df = pd.read_excel(f_pmt, engine="openpyxl")
        pmt_ids = pmt_df["ID"].astype(str).str.strip().tolist() if "ID" in pmt_df.columns else []
        df_final["Pmt o Desvíos Nuevos"] = df_final["Código Desvío"].apply(
            lambda x: "PMT" if str(x) in pmt_ids else "Desvío Nuevo")
    except:
        df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"
else:
    df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"

# -------- FILTROS INTERACTIVOS --------
st.sidebar.header("🔍 Filtros")
rutas = st.sidebar.multiselect("Filtrar por Ruta:", sorted(df_final["Ruta"].dropna().unique()))
zonas = st.sidebar.multiselect("Filtrar por Zona:", sorted(df_final["Zona"].dropna().unique()))
estados = st.sidebar.multiselect("Filtrar por Estado Final:", sorted(df_final["Estado Final"].dropna().unique()))

df_filtrado = df_final.copy()
if rutas: df_filtrado = df_filtrado[df_filtrado["Ruta"].isin(rutas)]
if zonas: df_filtrado = df_filtrado[df_filtrado["Zona"].isin(zonas)]
if estados: df_filtrado = df_filtrado[df_filtrado["Estado Final"].isin(estados)]

# -------- VISUALIZACIÓN --------
st.success("✅ Procesado con éxito. Vista previa:")
st.dataframe(df_filtrado, use_container_width=True)

# -------- GRÁFICAS --------
st.subheader("📊 Resumen visual")
g1 = df_filtrado["Estado Final"].value_counts()
fig1, ax1 = plt.subplots()
g1.plot(kind="bar", ax=ax1, color="#4e79a7")
ax1.set_title("Cantidad por Estado Final")
st.pyplot(fig1)

if "Ruta" in df_filtrado.columns:
    g2 = df_filtrado["Ruta"].value_counts().head(10)
    fig2, ax2 = plt.subplots()
    g2.plot(kind="barh", ax=ax2, color="#f28e2c")
    ax2.set_title("Top 10 Rutas con más desvíos")
    st.pyplot(fig2)

# -------- DESCARGA --------
buffer = BytesIO()
df_filtrado.to_excel(buffer, index=False)
buffer.seek(0)
st.download_button(
    "📅 Descargar Excel final",
    data=buffer,
    file_name=f"Revision de desvios {date.today().strftime('%Y-%m-%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
