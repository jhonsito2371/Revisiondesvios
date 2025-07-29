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

# ---------- CARGA DE ARCHIVOS ----------
col1, col2 = st.columns(2)
with col1:
    f_desv = st.file_uploader("📂 Archivo de Desvíos (acciones .xlsx)", type=["xlsx"], key="desv")
with col2:
    f_pmt = st.file_uploader("📂 Base PMT (.xlsx)", type=["xlsx"], key="pmt")

if not f_desv:
    st.info("👈 Sube al menos el **archivo de desvíos** para comenzar.")
    st.stop()

# ---------- LECTURA Y PROCESAMIENTO ----------
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
        base_cols = ['Fecha', 'Instante', 'Línea', 'Coche', 'Código Bus', 'Nº SAE Bus',
                     'Acción', 'Descripción Acción', 'Usuario', 'Nombre Usuario', 'Puesto',
                     'Parámetros', 'Motivo', 'Descripción Motivo', 'Otra Columna', 'RUTA']
        if raw.shape[1] == 17:
            raw.columns = base_cols + ['ZONA']
        else:
            raw.columns = base_cols
            raw['ZONA'] = ""
    return raw

try:
    df_raw = leer_desvios(f_desv)
except Exception as e:
    st.error("❌ Error al leer el archivo de desvíos.")
    st.exception(e)
    st.stop()

# ---------- FILTRADO Y NORMALIZACIÓN ----------
df = df_raw[df_raw.get("Descripción Acción", "").astype(str).str.lower().str.strip() == "desvio"].copy() if "Descripción Acción" in df_raw.columns else df_raw.copy()

df["Ruta"] = df.get("RUTA", "").astype(str).str.strip()
df["Zona"] = df.get("ZONA", "").astype(str).str.strip()

df["Estado Desvío"] = df["Parámetros"].apply(lambda x: "Activo" if isinstance(x, str) and any(k in x for k in ["Activar=\"SI\"", "Activo=\"SI\"", "ACTIVAR=\"SI\"", "ACTIVO=\"SI\""]) else "Inactivo")

def extraer_codigo(param):
    if isinstance(param, str):
        m = re.search(r'Desvio="(\d+)"', param)
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
    if cantidad == 1:
        return grupo.iloc[0]["Estado Desvío"]
    elif cantidad == 2:
        return grupo.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]
    else:
        if "Activo" in estados and "Inactivo" in estados:
            return "Modificado"
        return "Activo" if "Activo" in estados else "Inactivo"

estado_final = df.groupby("Código Desvío", group_keys=False).apply(evaluar_estado).reset_index()
estado_final.columns = ["Código Desvío", "Estado Final"]
conteo = df["Código Desvío"].value_counts().reset_index()
conteo.columns = ["Código Desvío", "Cantidad"]

ultimo_estado = df.groupby("Código Desvío")["Estado Desvío"].last().reset_index()
ultimo_estado.columns = ["Código Desvío", "Estados"]

# ---------- MERGE FINAL ----------
df_final = df.merge(estado_final, on="Código Desvío", how="left")
df_final = df_final.merge(conteo, on="Código Desvío", how="left")
df_final = df_final.merge(ultimo_estado, on="Código Desvío", how="left")
df_final["Revisión"] = df_final["Estados"].replace({"Activo": "Revisar", "Inactivo": "No Revisar"})

# ---------- CRUCE PMT ----------
if f_pmt:
    try:
        pmt_df = pd.read_excel(f_pmt, engine="openpyxl")
        ids = pmt_df["ID"].astype(str).str.strip().tolist() if "ID" in pmt_df.columns else []
        df_final["Pmt o Desvíos Nuevos"] = df_final["Código Desvío"].apply(lambda x: "PMT" if str(x) in ids else "Desvío Nuevo")
    except Exception as e:
        df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"
else:
    df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"

# ---------- FILTROS INTERACTIVOS ----------
st.sidebar.title("Filtros")
ruta_sel = st.sidebar.multiselect("Filtrar por Ruta", sorted(df_final["Ruta"].dropna().unique()))
zona_sel = st.sidebar.multiselect("Filtrar por Zona", sorted(df_final["Zona"].dropna().unique()))
estado_sel = st.sidebar.multiselect("Filtrar por Estado", sorted(df_final["Estado Final"].dropna().unique()))

filtros = (df_final["Ruta"].isin(ruta_sel) if ruta_sel else True) & \
          (df_final["Zona"].isin(zona_sel) if zona_sel else True) & \
          (df_final["Estado Final"].isin(estado_sel) if estado_sel else True)

df_filtrado = df_final[filtros].copy()

# ---------- VISUALIZACIÓN ----------
st.success("✅ Procesado con éxito. Vista previa:")
st.dataframe(df_filtrado, use_container_width=True)

# ---------- GRÁFICA ----------
st.subheader("🌐 Distribución por Estado Final")
fig, ax = plt.subplots()
df_filtrado["Estado Final"].value_counts().plot(kind="bar", color="skyblue", ax=ax)
ax.set_ylabel("Cantidad")
ax.set_xlabel("Estado Final")
ax.set_title("Cantidad de desvíos por estado")
st.pyplot(fig)

# ---------- DESCARGA ----------
cols_finales = ["Fecha Instante", "Hora Instante", "Nombre Usuario", "Código Desvío",
                "Estado Desvío", "Estado Final", "Cantidad", "Ruta", "Zona",
                "Pmt o Desvíos Nuevos", "Estados", "Revisión"]
cols_exist = [c for c in cols_finales if c in df_filtrado.columns]
excel_buffer = BytesIO()
df_filtrado[cols_exist].to_excel(excel_buffer, index=False)
excel_buffer.seek(0)

st.download_button("📥 Descargar Excel filtrado", data=excel_buffer,
                   file_name=f"Revision de desvios {date.today()}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")



