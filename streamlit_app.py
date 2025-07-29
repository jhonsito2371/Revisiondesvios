import streamlit as st
import pandas as pd
import re
from datetime import datetime, date
from pytz import timezone
from io import BytesIO

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

# ---------- LECTURA ROBUSTA DEL ARCHIVO DE DESVÍOS ----------
def leer_desvios(file):
    """
    Intenta leer el archivo de 'acciones' crudo (con encabezados en la segunda fila).
    Si no cuadra, intenta sin skiprows.
    Luego normaliza a columnas estándar esperadas.
    """
    try_orders = [
        {"skiprows": 1},   # formato más común en el archivo "acciones"
        {"skiprows": 0},   # por si ya viene con encabezados en la primera fila
    ]

    raw = None
    last_err = None
    for opts in try_orders:
        try:
            raw = pd.read_excel(file, engine="openpyxl", **opts)
            break
        except Exception as e:
            last_err = e
            continue
    if raw is None:
        raise last_err

    # Si el archivo viene crudo, suele traer 16/17 columnas; renombramos a estándar
    # Permitimos con o sin ZONA
    if raw.shape[1] in (16, 17):
        # Aseguramos longitud
        cols_objetivo_16 = [
            'Fecha', 'Instante', 'Línea', 'Coche', 'Código Bus', 'Nº SAE Bus',
            'Acción', 'Descripción Acción', 'Usuario', 'Nombre Usuario', 'Puesto',
            'Parámetros', 'Motivo', 'Descripción Motivo', 'Otra Columna', 'RUTA'
        ]
        cols_objetivo_17 = cols_objetivo_16 + ['ZONA']

        if raw.shape[1] == 16:
            raw.columns = cols_objetivo_16
            raw["ZONA"] = ""  # si no viene ZONA, la creamos vacía
        else:
            raw.columns = cols_objetivo_17
    else:
        # Si trae otros encabezados (por ejemplo un archivo ya procesado),
        # simplemente devolvemos lo leído para que más abajo validemos columnas.
        pass

    return raw

try:
    df_raw = leer_desvios(f_desv)
except Exception as e:
    st.error("❌ No se pudo leer el archivo de desvíos. Verifica que sea un .xlsx válido.")
    st.exception(e)
    st.stop()

# ---------- NORMALIZACIÓN Y VALIDACIONES ----------
# Si existe "Descripción Acción", filtramos a "Desvio"
if "Descripción Acción" in df_raw.columns:
    df = df_raw[df_raw["Descripción Acción"].astype(str).str.strip().str.lower() == "desvio"].copy()
else:
    df = df_raw.copy()

# Crear 'Ruta' y 'Zona' si existen o vacías
if "RUTA" in df.columns and "Ruta" not in df.columns:
    df["Ruta"] = df["RUTA"].astype(str).str.strip()
if "ZONA" in df.columns and "Zona" not in df.columns:
    df["Zona"] = df["ZONA"].astype(str).str.strip()
if "Ruta" not in df.columns:
    df["Ruta"] = ""
if "Zona" not in df.columns:
    df["Zona"] = ""

# Estado Activo/Inactivo desde 'Parámetros'
if "Parámetros" in df.columns:
    df["Estado Desvío"] = df["Parámetros"].apply(
        lambda x: "Activo" if isinstance(x, str) and (
            'Activar="SI"' in x or 'Activo="SI"' in x or 'ACTIVAR="SI"' in x or 'ACTIVO="SI"' in x
        ) else "Inactivo"
    )
else:
    st.error("❌ Falta la columna **Parámetros** en el archivo de desvíos.")
    st.write("Columnas detectadas:", list(df.columns))
    st.stop()

# Extraer Código Desvío desde 'Parámetros'
def extraer_codigo(param):
    if isinstance(param, str):
        m = re.search(r'Desvio="(\d+)"', param)
        if m:
            return m.group(1)
    return None

df["Código Desvío"] = df["Parámetros"].apply(extraer_codigo)

# Instante = Fecha + Hora (cuando vienen separadas)
if "Fecha" in df.columns and "Instante" in df.columns:
    df["Instante"] = pd.to_datetime(df["Fecha"].astype(str) + " " + df["Instante"].astype(str), errors="coerce")
elif "Instante" in df.columns:
    df["Instante"] = pd.to_datetime(df["Instante"], errors="coerce")
# Derivados de fecha/hora
df["Fecha Instante"] = df["Instante"].dt.date
df["Hora Instante"] = df["Instante"].dt.strftime("%H:%M:%S")

# Validación mínima para continuar
requeridas = ["Instante", "Código Desvío", "Estado Desvío", "Ruta", "Zona"]
faltantes = [c for c in requeridas if c not in df.columns]
if faltantes:
    st.error("❌ Faltan columnas necesarias para procesar:")
    st.write(faltantes)
    st.caption("Columnas detectadas:")
    st.write(list(df.columns))
    st.stop()

# ---------- ESTADO FINAL POR CÓDIGO ----------
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

# ---------- ESTADO RECIENTE + REVISIÓN ----------
def ultimo_estado(grupo):
    return grupo.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]

estado_mas_reciente = df_final.groupby("Código Desvío").apply(ultimo_estado).reset_index()
estado_mas_reciente.columns = ["Código Desvío", "Estados"]
df_final = pd.merge(df_final, estado_mas_reciente, on="Código Desvío", how="left")

df_final["Revisión"] = df_final["Estados"].replace({"Activo": "Revisar", "Inactivo": "No Revisar"})

# ---------- DURACIÓN: AHORA (CO) - INSTANTE (POR FILA) ----------
def calc_duracion_fila(instante):
    if pd.notnull(instante):
        ahora = datetime.now(timezone("America/Bogota")).replace(tzinfo=None)
        return ahora - instante
    return pd.NaT

df_final["Duración Activo"] = df_final["Instante"].apply(calc_duracion_fila)

def formato_duracion(td):
    if pd.isnull(td):
        return ""
    total = int(td.total_seconds())
    h = total // 3600
    m = (total % 3600) // 60
    if h > 0 and m > 0:
        return f"{h} horas {m} minutos"
    elif h > 0:
        return f"{h} horas"
    elif m > 0:
        return f"{m} minutos"
    else:
        return "Menos de 1 minuto"

df_final["Duración Activo"] = df_final["Duración Activo"].apply(formato_duracion)

# ---------- CRUCE PMT (OPCIONAL) ----------
if f_pmt:
    try:
        pmt_df = pd.read_excel(f_pmt, engine="openpyxl")
        if "ID" in pmt_df.columns:
            pmt_ids = pmt_df["ID"].astype(str).str.strip().tolist()
            df_final["Pmt o Desvíos Nuevos"] = df_final["Código Desvío"].apply(
                lambda x: "PMT" if str(x) in pmt_ids else "Desvío Nuevo"
            )
        else:
            df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"
            st.warning("⚠️ La base PMT no tiene columna 'ID'. Se marca todo como 'Desvío Nuevo'.")
    except Exception as e:
        st.warning("⚠️ No se pudo leer la base PMT. Se continúa sin cruce.")
        st.exception(e)
        df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"
else:
    df_final["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"

# ---------- ORDEN FINAL Y DESCARGA ----------
cols_finales = [
    "Fecha Instante", "Hora Instante", "Nombre Usuario", "Código Desvío",
    "Estado Desvío", "Estado Final", "Cantidad", "Ruta", "Zona",
    "Pmt o Desvíos Nuevos", "Estados", "Revisión", "Duración Activo"
]
# Algunas columnas podrían no existir (p.ej. Nombre Usuario). Mostramos las que haya.
cols_exist = [c for c in cols_finales if c in df_final.columns]
df_final = df_final[cols_exist].copy()

st.success("✅ Procesado con éxito. Vista previa:")
st.dataframe(df_final, use_container_width=True)

# Descargar Excel
buffer = BytesIO()
df_final.to_excel(buffer, index=False)
buffer.seek(0)
st.download_button(
    "📥 Descargar Excel final",
    data=buffer,
    file_name=f"Revision de desvios {date.today().strftime('%Y-%m-%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
