import io
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Carga masiva | Aula Virtual", layout="wide")
st.title("📥 Carga masiva de videoconferencias (Aula Virtual) — Validación")

# -----------------------
# Definición de plantilla
# -----------------------
COLUMNAS_REQUERIDAS = [
    "CORREO",   # host (cuenta en el AV)
    "TEMA",     # título de la reunión
    "PERIODO",  # ej. 20242
    "FACULTAD",
    "ESCUELA",
    "CURSO",
    "GRUPO",
    "INICIO",   # ej. 2025-08-15 07:40 (hora local Lima)
    "FIN",      # ej. 2025-08-15 09:20
    "DURACION", # minutos (si está vacío, lo calcularemos)
    "DIAS"      # como lo espera el AV (se procesará en el Paso 2)
]

st.subheader("1) Descargar plantilla")

plantilla = pd.DataFrame(columns=COLUMNAS_REQUERIDAS)
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as w:
    plantilla.to_excel(w, index=False, sheet_name="Plantilla")
    ayuda = pd.DataFrame({
        "Campo": COLUMNAS_REQUERIDAS,
        "Notas": [
            "Correo del host (cuenta en el AV).",
            "Tema de la reunión.",
            "Periodo académico (ej. 20242).",
            "Nombre de la facultad tal como aparece en el AV.",
            "Nombre de la escuela tal como aparece en el AV.",
            "Nombre del curso tal como aparece en el AV.",
            "Grupo/sección tal como aparece en el AV.",
            "Fecha y hora de inicio (ej. 2025-08-15 07:40).",
            "Fecha y hora de fin.",
            "Minutos de duración. Si lo dejas vacío, lo calcularemos con INICIO y FIN.",
            "Días tal como espera el AV (lo interpretaremos en el Paso 2)."
        ]
    })
    ayuda.to_excel(w, index=False, sheet_name="AYUDA")

st.download_button(
    "📄 Descargar plantilla (Excel)",
    data=buf.getvalue(),
    file_name="plantilla_cargamasiva_av.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# -----------------------
# Subir y validar archivo
# -----------------------
st.subheader("2) Subir y validar tu archivo")
archivo = st.file_uploader("Sube el Excel (.xlsx) con tus videoconferencias", type=["xlsx"])

def _a_dt(x):
    """Intenta convertir a datetime; devuelve None si no puede."""
    try:
        v = pd.to_datetime(x)
        if pd.isna(v):
            return None
        return pd.to_datetime(v).to_pydatetime()
    except Exception:
        return None

def _duracion_min(inicio, fin):
    """Minutos entre inicio y fin; si fin < inicio, asume cruce de medianoche."""
    if not inicio or not fin:
        return None
    delta = (fin - inicio).total_seconds() / 60
    if delta < 0:
        delta += 24 * 60
    return int(round(delta))

if archivo is not None:
    df = pd.read_excel(archivo)
    df.columns = [c.upper().strip() for c in df.columns]

    faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
    if faltantes:
        st.error("Faltan columnas obligatorias: " + ", ".join(faltantes))
        st.stop()

    prev = df.copy()
    prev["_INICIO_DT"] = prev["INICIO"].apply(_a_dt)
    prev["_FIN_DT"]    = prev["FIN"].apply(_a_dt)

    def _dur_preview(row):
        val = row.get("DURACION")
        try:
            if pd.isna(val) or str(val).strip() == "":
                return _duracion_min(row["_INICIO_DT"], row["_FIN_DT"])
            return int(val)
        except Exception:
            return _duracion_min(row["_INICIO_DT"], row["_FIN_DT"])

    prev["DURACION_PREVIEW"] = prev.apply(_dur_preview, axis=1)
    prev["OK_INICIO"]   = prev["_INICIO_DT"].apply(lambda x: x is not None)
    prev["OK_FIN"]      = prev["_FIN_DT"].apply(lambda x: x is not None)
    prev["OK_DURACION"] = prev["DURACION_PREVIEW"].apply(lambda x: isinstance(x, int) and x > 0)

    st.success(f"Archivo cargado: {len(prev)} filas • {len(prev.columns)} columnas")
    st.caption("Se muestran las primeras 20 filas. DURACION_PREVIEW es solo para verificación (se autocalcula si falta).")
    st.dataframe(prev.head(20), use_container_width=True)

    problemas = []
    if not prev["OK_INICIO"].all():   problemas.append("Hay filas con INICIO inválido/no parseable.")
    if not prev["OK_FIN"].all():      problemas.append("Hay filas con FIN inválido/no parseable.")
    if not prev["OK_DURACION"].all(): problemas.append("Hay filas con DURACION vacía o no válida (se puede autocalcular).")

    if problemas:
        st.warning("⚠️ Observaciones:\n- " + "\n- ".join(problemas))
    else:
        st.success("✅ Listo para el siguiente paso (ejecución).")

st.divider()
st.subheader("3) ¿Qué sigue?")
st.markdown(
"""
**Paso 2** (cuando confirmes): integramos el botón **Ejecutar** con **MODO PRUEBA/PRODUCCIÓN**,  
credenciales por **variables de entorno** y generación de **reporte** (CSV) por fila.
"""
)
