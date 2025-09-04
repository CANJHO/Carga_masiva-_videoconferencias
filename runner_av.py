# runner_av.py
import os
import time
import csv
from datetime import datetime
from typing import Dict, Any, List, Tuple

import pandas as pd
from dotenv import load_dotenv

# Playwright
from playwright.sync_api import sync_playwright

load_dotenv()

AV_URL = os.getenv("AV_URL")
AV_USER = os.getenv("AV_USER")
AV_PASS = os.getenv("AV_PASS")
TZ     = os.getenv("TZ", "America/Lima")

LOG_DIR = "logs"
SS_DIR  = "screenshots"

os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(SS_DIR, exist_ok=True)

def _now_tag() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def _duracion_min(inicio_dt, fin_dt) -> int:
    if pd.isna(inicio_dt) or pd.isna(fin_dt):
        return None
    delta = (pd.to_datetime(fin_dt) - pd.to_datetime(inicio_dt)).total_seconds() / 60
    if delta < 0:
        delta += 24*60
    return int(round(delta))

def _prep_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza mínimamente: fechas a datetime y calcula DURACION si falta/incorrecta."""
    t = df.copy()
    t.columns = [c.upper().strip() for c in t.columns]

    # Parse de INICIO/FIN
    t["_INICIO_DT"] = pd.to_datetime(t["INICIO"], errors="coerce")
    t["_FIN_DT"]    = pd.to_datetime(t["FIN"], errors="coerce")

    # Duración (si falta o inválida)
    def _dur(row):
        v = row.get("DURACION")
        try:
            if pd.isna(v) or str(v).strip() == "":
                return _duracion_min(row["_INICIO_DT"], row["_FIN_DT"])
            return int(v)
        except Exception:
            return _duracion_min(row["_INICIO_DT"], row["_FIN_DT"])

    t["DURACION_CALC"] = t.apply(_dur, axis=1)

    return t

def _write_logs(base_name: str, rows: List[Dict[str, Any]]) -> Tuple[str, str]:
    """Escribe TXT y CSV con el resultado del lote."""
    ts = _now_tag()
    txt_path = os.path.join(LOG_DIR, f"{base_name}_{ts}.txt")
    csv_path = os.path.join(LOG_DIR, f"{base_name}_{ts}.csv")

    # TXT
    with open(txt_path, "w", encoding="utf-8") as f:
        for r in rows:
            f.write(
                f"[{r['timestamp']}] {r['status']} | {r['correo']} | {r['tema']} | "
                f"{r['inicio']} -> {r['fin']} | {r['mensaje']}\n"
            )

    # CSV
    fieldnames = ["timestamp","status","correo","tema","periodo","facultad","escuela","curso",
                  "grupo","inicio","fin","duracion","dias","mensaje","meeting_url"]
    with open(csv_path, "w", encoding="utf-8", newline="") as c:
        w = csv.DictWriter(c, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k,"") for k in fieldnames})

    return txt_path, csv_path

# ---------- Playwright helpers (estructura) ----------
def _login(page):
    # En el Paso 3 insertamos selectores reales de login
    page.goto(AV_URL, wait_until="domcontentloaded")
    # ejemplo:
    # page.fill("input[name='username']", AV_USER)
    # page.fill("input[name='password']", AV_PASS)
    # page.click("button:has-text('Ingresar')")
    # page.wait_for_load_state("networkidle")
    pass

def create_in_av(page, row: Dict[str, Any], save: bool = True, screenshot_path: str = "") -> Tuple[bool, str, str]:
    """
    Crea o simula la creación de una videoconferencia.
    - save=False  → solo rellena y toma captura (PRUEBA VISUAL).
    - save=True   → hace clic en Guardar (PRODUCCIÓN).

    Devuelve: (ok, message, meeting_url)
    """

    # 1) Abrir modal "Nueva videoconferencia"
    # page.click("text= Nueva videoconferencia")  # ← TU SELECTOR

    # 2) Seleccionar PERIODO / FACULTAD / ESCUELA / CURSO / GRUPO
    # seleccionar_opcion_unica(page, "selector-periodo", row["PERIODO"])
    # seleccionar_opcion_unica(page, "selector-facultad", row["FACULTAD"])
    # seleccionar_opcion_unica(page, "selector-escuela", row["ESCUELA"])
    # seleccionar_opcion_unica(page, "selector-curso", row["CURSO"])
    # seleccionar_opcion_unica(page, "selector-grupo", row["GRUPO"])

    # 3) Host/correo
    # page.fill("input[name='correo']", row["CORREO"])

    # 4) Título / tema
    # page.fill("input[name='tema']", row["TEMA"])

    # 5) Fechas y duración (usa row["_INICIO_DT"], row["_FIN_DT"], row["DURACION_CALC"])
    # page.fill("input[name='inicio']", formatea(row["_INICIO_DT"]))
    # page.fill("input[name='fin']",    formatea(row["_FIN_DT"]))
    # page.fill("input[name='duracion']", str(row["DURACION_CALC"]))

    # 6) Días (si corresponde en tu UI)
    # marcar_dias(page, row["DIAS"])

    # ----- PRUEBA VISUAL: NO GUARDAR -----
    if not save:
        if screenshot_path:
            page.screenshot(path=screenshot_path, full_page=True)
        return True, "Simulación visual: formulario rellenado (no se guardó).", ""

    # 7) Guardar
    # page.click("button:has-text('Guardar')")

    # 8) Leer SweetAlert / mensaje y (si es posible) URL/ID de la reunión
    # msg = page.locator(".swal-text,.swal2-html-container").inner_text()
    # url = extraer_url(page)  # si tu UI la muestra
    msg = "Guardado (implementa lectura de mensaje)"
    url = ""
    return True, msg, url


# ---------- API principal del runner ----------
def run_batch(df: pd.DataFrame, modo: str, headless: bool) -> Dict[str, Any]:
    """
    modo: "PRUEBA (sin navegador)" | "PRUEBA VISUAL (navegador, sin guardar)" | "PRODUCCIÓN"
    """
    if not AV_URL or not AV_USER or not AV_PASS:
        raise RuntimeError("Faltan variables de entorno AV_URL/AV_USER/AV_PASS en .env")

    t = _prep_dataframe(df)

    resultados: List[Dict[str, Any]] = []
    base_log_name = "cargamasiva_av"
    screenshots_dir = SS_DIR

    # ---------------- PRUEBA (sin navegador) ----------------
    if modo.startswith("PRUEBA (sin navegador)"):
        for _, r in t.iterrows():
            resultados.append({
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "status": "SIMULADO",
                "correo": str(r.get("CORREO","")),
                "tema": str(r.get("TEMA","")),
                "periodo": str(r.get("PERIODO","")),
                "facultad": str(r.get("FACULTAD","")),
                "escuela": str(r.get("ESCUELA","")),
                "curso": str(r.get("CURSO","")),
                "grupo": str(r.get("GRUPO","")),
                "inicio": str(r.get("_INICIO_DT","")),
                "fin": str(r.get("_FIN_DT","")),
                "duracion": str(r.get("DURACION_CALC","")),
                "dias": str(r.get("DIAS","")),
                "mensaje": "OK (modo prueba sin navegador)",
                "meeting_url": ""
            })
        txt, csv = _write_logs(base_log_name+"_PRUEBA", resultados)
        return {
            "total": len(resultados),
            "ok": len(resultados),
            "fail": 0,
            "log_txt": txt,
            "log_csv": csv
        }

    # --------------- VISUAL y PRODUCCIÓN (con navegador) ---------------
    visual = modo.startswith("PRUEBA VISUAL")
    produccion = modo.startswith("PRODUCCIÓN")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(locale="es-PE", timezone_id=TZ)
        page = context.new_page()
        try:
            _login(page)  # ← Implementa tus selectores reales de login aquí

            for i, r in t.iterrows():
                try:
                    # ruta de captura por fila
                    ss_path = os.path.join(screenshots_dir, f"{'visual' if visual else 'prod'}_row{i+1}_{_now_tag()}.png")
                    ok, msg, url = create_in_av(
                        page,
                        r.to_dict(),
                        save=not visual,                 # ← en VISUAL no guardamos
                        screenshot_path=ss_path          # ← capturamos formulario lleno
                    )
                    status = "SIMULADO_VISUAL" if visual else ("GUARDADO" if ok else "ERROR")
                    resultados.append({
                        "timestamp": datetime.now().isoformat(timespec="seconds"),
                        "status": status,
                        "correo": str(r.get("CORREO","")),
                        "tema": str(r.get("TEMA","")),
                        "periodo": str(r.get("PERIODO","")),
                        "facultad": str(r.get("FACULTAD","")),
                        "escuela": str(r.get("ESCUELA","")),
                        "curso": str(r.get("CURSO","")),
                        "grupo": str(r.get("GRUPO","")),
                        "inicio": str(r.get("_INICIO_DT","")),
                        "fin": str(r.get("_FIN_DT","")),
                        "duracion": str(r.get("DURACION_CALC","")),
                        "dias": str(r.get("DIAS","")),
                        "mensaje": msg,
                        "meeting_url": url
                    })
                except Exception as e:
                    ss_path = os.path.join(screenshots_dir, f"error_row{i+1}_{_now_tag()}.png")
                    try: page.screenshot(path=ss_path, full_page=True)
                    except: pass
                    resultados.append({
                        "timestamp": datetime.now().isoformat(timespec="seconds"),
                        "status": "ERROR",
                        "correo": str(r.get("CORREO","")),
                        "tema": str(r.get("TEMA","")),
                        "periodo": str(r.get("PERIODO","")),
                        "facultad": str(r.get("FACULTAD","")),
                        "escuela": str(r.get("ESCUELA","")),
                        "curso": str(r.get("CURSO","")),
                        "grupo": str(r.get("GRUPO","")),
                        "inicio": str(r.get("_INICIO_DT","")),
                        "fin": str(r.get("_FIN_DT","")),
                        "duracion": str(r.get("DURACION_CALC","")),
                        "dias": str(r.get("DIAS","")),
                        "mensaje": f"Excepción: {e}",
                        "meeting_url": ""
                    })
        finally:
            try:
                context.close()
                browser.close()
            except:
                pass

    suf = "_VISUAL" if visual else ""
    txt, csv = _write_logs(base_log_name+suf, resultados)
    return {
        "total": len(resultados),
        "ok": len([r for r in resultados if r["status"] in ("SIMULADO_VISUAL","GUARDADO")]),
        "fail": len([r for r in resultados if r["status"] == "ERROR"]),
        "log_txt": txt,
        "log_csv": csv,
        "screenshots_dir": screenshots_dir
    }
