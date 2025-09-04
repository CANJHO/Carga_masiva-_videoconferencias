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

def create_in_av(page, row: Dict[str, Any]) -> Tuple[bool, str, str]:
    """
    Crea una videoconferencia en el AV a partir de 'row'.
    DEVUELVE: (ok, message, meeting_url)
    En el Paso 3 pegamos aquí los selectores reales (Nueva videoconferencia, selects, guardar, etc.)
    """
    # --- SELECTORES REALES AQUÍ EN EL PASO 3 ---
    # Debe usar: row["CORREO"], row["TEMA"], row["PERIODO"], row["FACULTAD"],
    #            row["ESCUELA"], row["CURSO"], row["GRUPO"], row["_INICIO_DT"],
    #            row["_FIN_DT"], row["DURACION_CALC"], row["DIAS"]
    ok = False
    msg = "No implementado aún"
    url = ""
    return ok, msg, url

# ---------- API principal del runner ----------
def run_batch(df: pd.DataFrame, modo_prueba: bool, headless: bool) -> Dict[str, Any]:
    if not AV_URL or not AV_USER or not AV_PASS:
        raise RuntimeError("Faltan variables de entorno AV_URL/AV_USER/AV_PASS en .env")

    t = _prep_dataframe(df)

    resultados: List[Dict[str, Any]] = []
    base_log_name = "cargamasiva_av"

    if modo_prueba:
        # Simular sin abrir navegador
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
                "mensaje": "OK (modo prueba)",
                "meeting_url": ""
            })
        txt, csv = _write_logs(base_log_name+"_PRUEBA", resultados)
        summary = {
            "total": len(resultados),
            "ok": len([r for r in resultados if r["status"]=="SIMULADO"]),
            "fail": 0,
            "log_txt": txt,
            "log_csv": csv
        }
        return summary

    # Producción: Playwright
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(locale="es-PE", timezone_id=TZ)
        page = context.new_page()
        try:
            _login(page)

            for i, r in t.iterrows():
                try:
                    ok, msg, url = create_in_av(page, r.to_dict())
                    status = "GUARDADO" if ok else "ERROR"
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
                    # screenshot por fila fallida
                    ss_path = os.path.join(SS_DIR, f"error_row{i+1}_{_now_tag()}.png")
                    try:
                        page.screenshot(path=ss_path, full_page=True)
                    except:
                        pass
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

    txt, csv = _write_logs(base_log_name, resultados)
    summary = {
        "total": len(resultados),
        "ok": len([r for r in resultados if r["status"] in ("GUARDADO","SIMULADO")]),
        "fail": len([r for r in resultados if r["status"] == "ERROR"]),
        "log_txt": txt,
        "log_csv": csv
    }
    return summary
