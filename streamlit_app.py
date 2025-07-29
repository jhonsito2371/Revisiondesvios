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

# ---------- CARGA DE ARCHIVOS ----------
col1, col2 = st.columns(2)
with col1:
    f_desv = st.file_uploader("📂 Archivo de Desvíos (acciones .xlsx)", type=["xlsx"], key="desv")
with col2:
    f_pmt = st.file_uploader("📂 Base PMT (.xlsx)", type=["xlsx"], key="pmt")

if not f_desv:
    st.info("👈 Sube al menos el **archivo de desvíos** para comenzar.")
    st.stop()

# ---------- LECTURA DEL ARCHIVO ----------
def leer_desvios(file):
    try:
        df = pd.read_excel(file, skiprows=1)
    except:
        df = pd.read_excel(file)

    if df.shape[1] == 16:
        df.columns = [
            'Fecha', 'Instante', 'Línea', 'Coche', 'Código Bus', 'Nº SAE Bus',
            'Acción', 'Descripción Acción', 'Usuario', 'Nombre Usuario', 'Puesto',
            'Parámetros', 'Motivo', 'Descripción Motivo', 'Otra Columna', 'RUTA'
        ]
        df["ZONA"] = ""
    elif df.shape[1] == 17:
        df.columns = [
            'Fecha', 'Instante', 'Línea', 'Coche', 'Código Bus', 'Nº SAE Bus',
            'Acción', 'Descripción Acción', 'Usuario', 'Nombre Usuario', 'Puesto',
            'Parámetros', 'Motivo', 'Descripción Motivo', 'Otra Columna', 'RUTA', 'ZONA'
        ]
    return df

try:
    df_raw = leer_desvios(f_desv)
except Exception as e:
    st.error("❌ No se pudo leer el archivo de desvíos. Verifica que sea un .xlsx válido.")
    st.exception(e)
    st.stop()

if "Descripción Acción" in df_raw.columns:
    df = df_raw[df_raw["Descripción Acción"].astype(str).str.lower() == "desvio"].copy()
else:
    df = df_raw.copy()

df["Ruta"] = df["RUTA"].astype(str).str.strip()
df["Zona"] = df["ZONA"].astype(str).str.strip()

# Estado Desvío
if "Parámetros" in df.columns:
    df["Estado Desvío"] = df["Parámetros"].apply(
        lambda x: "Activo" if isinstance(x, str) and ('Activar="SI"' in x or 'Activo="SI"' in x) else "Inactivo")
else:
    st.stop()

# Código Desvío
def extraer_codigo(param):
    if isinstance(param, str):
        m = re.search(r'Desvio="(\d+)"', param)
        if m:
            return m.group(1)
    return None

df["Código Desvío"] = df["Parámetros"].apply(extraer_codigo)

# Instante
df["Instante"] = pd.to_datetime(df["Fecha"].astype(str) + " " + df["Instante"].astype(str), errors="coerce")
df["Fecha Instante"] = df["Instante"].dt.date
df["Hora Instante"] = df["Instante"].dt.strftime("%H:%M:%S")

# Estado Final y Estados
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

def ultimo_estado(grupo):
    return grupo.sort_values("Instante", ascending=False).iloc[0]["Estado Desvío"]

df["Cantidad"] = df.groupby("Código Desvío")["Código Desvío"].transform("count")
df["Estado Final"] = df.groupby("Código Desvío", group_keys=False).apply(evaluar_estado)
df["Estados"] = df.groupby("Código Desvío", group_keys=False).apply(ultimo_estado)
df["Revisión"] = df["Estados"].replace({"Activo": "Revisar", "Inactivo": "No Revisar"})

# Duración
ahora = datetime.now(timezone("America/Bogota")).replace(tzinfo=None)
df["Duración Activo"] = df["Instante"].apply(lambda x: ahora - x if pd.notnull(x) else pd.NaT)

def formato_duracion(td):
    if pd.isnull(td): return ""
    total = int(td.total_seconds())
    h = total // 3600
    m = (total % 3600) // 60
    if h > 0 and m > 0: return f"{h} horas {m} minutos"
    elif h > 0: return f"{h} horas"
    elif m > 0: return f"{m} minutos"
    else: return "Menos de 1 minuto"

df["Duración Activo"] = df["Duración Activo"].apply(formato_duracion)

# PMT
if f_pmt:
    try:
        pmt_df = pd.read_excel(f_pmt)
        if "ID" in pmt_df.columns:
            ids = pmt_df["ID"].astype(str).tolist()
            df["Pmt o Desvíos Nuevos"] = df["Código Desvío"].apply(lambda x: "PMT" if str(x) in ids else "Desvío Nuevo")
        else:
            df["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"
    except:
        df["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"
else:
    df["Pmt o Desvíos Nuevos"] = "Desvío Nuevo"

# ---------- FILTROS ----------
st.sidebar.header("🔍 Filtros")
rutas = st.sidebar.multiselect("Ruta", sorted(df["Ruta"].dropna().unique()))
zonas = st.sidebar.multiselect("Zona", sorted(df["Zona"].dropna().unique()))
estados = st.sidebar.multiselect("Estado Final", sorted(df["Estado Final"].dropna().unique()))

filtro_df = df.copy()
if rutas:
    filtro_df = filtro_df[filtro_df["Ruta"].isin(rutas)]
if zonas:
    filtro_df = filtro_df[filtro_df["Zona"].isin(zonas)]
if estados:
    filtro_df = filtro_df[filtro_df["Estado Final"].isin(estados)]

st.success("✅ Procesado con éxito. Vista previa:")
st.dataframe(filtro_df)

# ---------- GRAFICAS ----------
st.subheader("📊 Visualización de Datos")
col1, col2 = st.columns(2)

with col1:
    fig1 = px.histogram(filtro_df, x="Ruta", color="Estado Final", title="Cantidad por Ruta y Estado")
    st.plotly_chart(fig1, use_container_width=True)

with col2:
    fig2 = px.pie(filtro_df, names="Zona", title="Distribución por Zona")
    st.plotly_chart(fig2, use_container_width=True)

# ---------- DESCARGA ----------
st.subheader("📥 Exportar")
buffer = BytesIO()
filtro_df.to_excel(buffer, index=False)
buffer.seek(0)
st.download_button(
    "📅 Descargar Excel",
    data=buffer,
    file_name=f"Resumen Desvios {date.today().strftime('%Y-%m-%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ---------- ENVÍO POR CORREO ----------
st.subheader("📧 Enviar por correo")
correo_destino = st.text_input("Correo de destino")

if st.button("📤 Enviar resumen por correo"):
    try:
        emisor = "tucorreo@hotmail.com"
        clave = "tu_contraseña"

        mensaje = EmailMessage()
        mensaje["Subject"] = f"Resumen de Desvíos Operativos - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        mensaje["From"] = emisor
        mensaje["To"] = correo_destino
        mensaje.set_content("Adjunto el resumen de desvíos operativos filtrado.")

        mensaje.add_attachment(
            buffer.getvalue(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="Resumen Desvios.xlsx"
        )

        with smtplib.SMTP("smtp.office365.com", 587) as smtp:
            smtp.starttls()
            smtp.login(emisor, clave)
            smtp.send_message(mensaje)

        st.success("✅ Correo enviado con éxito")
    except Exception as e:
        st.error(f"❌ Error al enviar el correo: {e}")

