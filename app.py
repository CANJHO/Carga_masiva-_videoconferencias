# app.py
import io
from datetime import datetime
import pandas as pd
import streamlit as st

from runner_av import run_batch

st.set_page_config(page_title="Carga masiva | Aula Virtual", layout="wide")
st.title("📥 Carga masiva de videoconferencias (Aula Virtual)")

st.caption("Sube tu Excel, elige el modo y ejecuta. En **PRUEBA VISUAL** el robot llena el formulario en el navegador **sin guardar**.")

# -----------------------
# Definición de plantilla
# -----------------------
COLUMNAS_REQUERIDAS = [
    "CORREO",   # host (cuenta en el AV)
    "TEMA",     # título de la reunión
    "PERIODO",
    "FACULTAD",
    "ESCUELA",
    "CURSO",
    "GRUPO",
    "INICIO",   # 'YYYY-mm-dd HH:MM' (hora local Lima)
    "FIN",      # 'YYYY-mm-dd HH:MM'
    "DURACION", # minutos (si está vacío, se calcula)
    "DIAS"      # como lo espera el AV (ej. LU,MA,MI ... o 1,2,3 ...)
]

with st.expander("📄 Descargar plantilla (requerida)"):
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
                "Fecha y hora de inicio (YYYY-mm-dd HH:MM).",
                "Fecha y hora de fin (YYYY-mm-dd HH:MM).",
                "Minutos de duración. Si lo dejas vacío, se calcula.",
                "Días (LU,MA,MI,JU,VI,SA,DO o 1..7)."
            ]
        })
        ayuda.to_excel(w, index=False, sheet_name="AYUDA")

    st.download_button(
        "⬇️ Descargar plantilla (Excel)",
        data=buf.getvalue(),
        file_name="plantilla_cargamasiva_av.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -----------------------
# Subir y validar archivo
# -----------------------
st.subheader("1) Sube y valida tu archivo")
archivo = st.file_uploader("Sube el Excel (.xlsx) con tus videoconferencias", type=["xlsx"])

def _a_dt(x):
    try:
        v = pd.to_datetime(x)
        if pd.isna(v):
            return None
        return pd.to_datetime(v).to_pydatetime()
    except Exception:
        return None

def _duracion_min(inicio, fin):
    if not inicio or not fin:
        return None
    delta = (fin - inicio).total_seconds() / 60
    if delta < 0:
        delta += 24 * 60
    return int(round(delta))

df_validado = None
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

    df_validado = prev

    st.success(f"Archivo cargado: {len(prev)} filas")
    st.caption("Vista previa (primeras 20 filas). La duración se calcula si no viene en el archivo.")
    st.dataframe(prev.head(20), use_container_width=True)

    problemas = []
    if not prev["OK_INICIO"].all():   problemas.append("Hay filas con INICIO inválido.")
    if not prev["OK_FIN"].all():      problemas.append("Hay filas con FIN inválido.")
    if not prev["OK_DURACION"].all(): problemas.append("Hay filas con DURACION vacía o inválida (se puede autocalcular).")

    if problemas:
        st.warning("⚠️ Observaciones:\n- " + "\n- ".join(problemas))
    else:
        st.success("✅ Listo para el siguiente paso.")

st.divider()

# ========================
# 2) Ejecutar
# ========================
st.subheader("2) Ejecutar lote")

col1, col2 = st.columns([2,1])
with col1:
    modo = st.selectbox(
        "Modo",
        ["PRUEBA VISUAL (navegador, sin guardar)", "PRODUCCIÓN"],
        index=0,
        help="En PRUEBA VISUAL verás al robot llenar el formulario sin guardar; en PRODUCCIÓN guardará realmente."
    )
with col2:
    # En PRUEBA VISUAL, forzamos navegador visible (headless=False)
    headless = st.checkbox("Headless (oculto)", value=(modo == "PRODUCCIÓN"),
                           help="Déjalo desmarcado para ver el navegador. En PRUEBA VISUAL se ignora y se muestra siempre.")

ejecutar = st.button("🚀 Ejecutar ahora", type="primary", use_container_width=True)

if ejecutar:
    if df_validado is None:
        st.error("Primero sube tu Excel válido.")
        st.stop()

    # Para runner: usar el df original, no la preview extendida
    df_to_run = df.copy()

    # Corre
    st.info("Iniciando…")
    resumen = run_batch(
        df_to_run,
        modo=("PRUEBA VISUAL (navegador, sin guardar)" if modo.startswith("PRUEBA") else "PRODUCCIÓN"),
        headless=(False if modo.startswith("PRUEBA") else headless)
    )

    st.success(f"✅ Lote terminado • Total: {resumen['total']} • OK: {resumen['ok']} • Fallas: {resumen['fail']}")
    st.write(f"📄 Log TXT: {resumen.get('log_txt', '')}")
    st.write(f"📊 Log CSV: {resumen.get('log_csv', '')}")
    if 'screenshots_dir' in resumen:
        st.write(f"🖼️ Capturas: {resumen['screenshots_dir']}")
