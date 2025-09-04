# runner_av.py
import os
import csv
from datetime import datetime
from typing import Dict, Any, List, Tuple

import pandas as pd
from dotenv import load_dotenv

# ===== FIX para Windows / Streamlit: event loop correcto =====
import sys, asyncio
if sys.platform.startswith("win"):
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    except Exception:
        pass
# ============================================================

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
    t = df.copy()
    t.columns = [c.upper().strip() for c in t.columns]
    t["_INICIO_DT"] = pd.to_datetime(t["INICIO"], errors="coerce")
    t["_FIN_DT"]    = pd.to_datetime(t["FIN"], errors="coerce")

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
    ts = _now_tag()
    txt_path = os.path.join(LOG_DIR, f"{base_name}_{ts}.txt")
    csv_path = os.path.join(LOG_DIR, f"{base_name}_{ts}.csv")

    with open(txt_path, "w", encoding="utf-8") as f:
        for r in rows:
            f.write(
                f"[{r['timestamp']}] {r['status']} | {r['correo']} | {r['tema']} | "
                f"{r['inicio']} -> {r['fin']} | {r['mensaje']}\n"
            )

    fieldnames = ["timestamp","status","correo","tema","periodo","facultad","escuela","curso",
                  "grupo","inicio","fin","duracion","dias","mensaje","meeting_url"]
    with open(csv_path, "w", encoding="utf-8", newline="") as c:
        w = csv.DictWriter(c, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k,"") for k in fieldnames})

    return txt_path, csv_path

# ----------- Login (coloca aquí tus selectores reales) -----------
def _login(page):
    # 1) Ir al login
    page.goto(AV_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(800)

    # (debug) verifica que las credenciales EXISTEN
    if not AV_USER or not AV_PASS:
        raise RuntimeError("AV_USER o AV_PASS están vacíos. Revisa tu .env y reinicia la app.")

    # 2) USERNAME: intentos por id/name/placeholder/text input
    username_filled = False
    for sel in [
        "input[name='username']",
        "input#username",
        "input[name*='user' i]",
        "input[placeholder*='Usuario' i]",
        "input[placeholder*='Correo' i]",
    ]:
        try:
            page.locator(sel).first.wait_for(state="visible", timeout=3000)
            page.locator(sel).first.fill(AV_USER)
            username_filled = True
            break
        except:
            continue
    if not username_filled:
        # último intento genérico
        try:
            page.locator("input[type='text']").first.fill(AV_USER)
            username_filled = True
        except:
            pass

    # 3) PASSWORD: por type/name/placeholder
    pwd_filled = False
    for sel in [
        "input[type='password']",
        "input[name='password']",
        "input[placeholder*='Contraseña' i]",
    ]:
        try:
            page.locator(sel).first.wait_for(state="visible", timeout=3000)
            page.locator(sel).first.fill(AV_PASS)
            pwd_filled = True
            break
        except:
            continue

    # 4) Clic en el botón Ingresar (varios textos posibles)
    clicked = False
    for txt in ["Ingresar", "Acceder", "Entrar", "Iniciar sesión", "Iniciar sesion", "Login", "Sign in"]:
        try:
            page.get_by_role("button", name=txt, exact=False).click(timeout=2000)
            clicked = True
            break
        except:
            try:
                page.get_by_text(txt, exact=False).first.click(timeout=2000)
                clicked = True
                break
            except:
                continue

    # 5) Espera a que complete el login (o al menos navegar)
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except:
        page.wait_for_timeout(1000)

    # 6) Captura para ver en qué quedó (carpeta screenshots/)
    try:
        page.screenshot(path=os.path.join(SS_DIR, f"login_state_{_now_tag()}.png"), full_page=True)
    except:
        pass

    # 7) Pausa breve para que lo veas en PRUEBA VISUAL
    page.wait_for_timeout(1200)


# ----------- Crear en AV (gancho para PRUEBA VISUAL / PROD) -----------
def create_in_av(page, row: Dict[str, Any], save: bool = True, screenshot_path: str = "") -> Tuple[bool, str, str]:
    """
    Rellena el formulario de videoconferencia.
    - save=False: solo deja el formulario lleno y toma captura (PRUEBA VISUAL)
    - save=True : presiona Guardar (PRODUCCIÓN)
    Devuelve (ok, mensaje, meeting_url)
    """

    # Aquí va tu flujo real: abrir modal, seleccionar PERIODO/FACULTAD/ESCUELA/CURSO/GRUPO,
    # completar CORREO/TEMA/INICIO/FIN/DURACION, marcar DIAS, etc.
    # Ejemplos (coloca tus selectores):
    #
    # page.click("text='Nueva videoconferencia'")
    # seleccionar(page, "selector-periodo", row["PERIODO"])
    # seleccionar(page, "selector-facultad", row["FACULTAD"])
    # seleccionar(page, "selector-escuela", row["ESCUELA"])
    # seleccionar(page, "selector-curso", row["CURSO"])
    # seleccionar(page, "selector-grupo", row["GRUPO"])
    # page.fill("input[name='correo']", row["CORREO"])
    # page.fill("input[name='tema']", row["TEMA"])
    # page.fill("input[name='inicio']",  formatea(row["_INICIO_DT"]))
    # page.fill("input[name='fin']",     formatea(row["_FIN_DT"]))
    # page.fill("input[name='duracion']", str(row["DURACION_CALC"]))
    # marcar_dias(page, row["DIAS"])

    # Captura del estado (sea visual o prod)
    if screenshot_path:
        try:
            page.screenshot(path=screenshot_path, full_page=True)
        except:
            pass

    if not save:
        return True, "Simulación visual: formulario rellenado (no se guardó).", ""

    # Guardar (completa con tu selector) y leer mensaje
    # page.click("button:has-text('Guardar')")
    # msg = page.locator(".swal-text,.swal2-html-container").inner_text()
    msg = "Guardado (implementar lectura del mensaje)"
    meeting_url = ""  # si tu UI la muestra, extrae aquí
    return True, msg, meeting_url

# ------------------- Batch runner con MODOS -------------------
def run_batch(df: pd.DataFrame, modo: str, headless: bool) -> Dict[str, Any]:
    """
    modo:
      - "PRUEBA (sin navegador)"
      - "PRUEBA VISUAL (navegador, sin guardar)"
      - "PRODUCCIÓN"
    """
    if not AV_URL or not AV_USER or not AV_PASS:
        raise RuntimeError("Faltan variables de entorno AV_URL/AV_USER/AV_PASS en .env")

    t = _prep_dataframe(df)

    resultados: List[Dict[str, Any]] = []
    base_log_name = "cargamasiva_av"
    screenshots_dir = SS_DIR

    # ------- PRUEBA sin navegador -------
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
        return {"total": len(resultados), "ok": len(resultados), "fail": 0,
                "log_txt": txt, "log_csv": csv}

    # ------- PRUEBA VISUAL / PRODUCCIÓN (con navegador) -------
    visual = modo.startswith("PRUEBA VISUAL")
    produccion = modo.startswith("PRODUCCIÓN")

    with sync_playwright() as p:
        # Abrimos con ventana visible, más lenta y maximizada (cuando headless=False)
        browser = p.chromium.launch(
            headless=headless,
            slow_mo=400,                          # retrasa acciones para que las veas
            args=["--start-maximized"]            # abre max
        )
        # no_viewport=True deja que la ventana use el tamaño del SO
        context = browser.new_context(
            no_viewport=True,
            locale="es-PE",
            timezone_id=TZ
        )
        page = context.new_page()
        try:
            _login(page)

            for i, r in t.iterrows():
                try:
                    ss_path = os.path.join(screenshots_dir,
                        f"{'visual' if visual else 'prod'}_row{i+1}_{_now_tag()}.png")
                    ok, msg, url = create_in_av(
                        page,
                        r.to_dict(),
                        save=not visual,           # en visual NO guarda
                        screenshot_path=ss_path
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
                    err_ss = os.path.join(screenshots_dir, f"error_row{i+1}_{_now_tag()}.png")
                    try:
                        page.screenshot(path=err_ss, full_page=True)
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

    suf = "_VISUAL" if visual else ""
    txt, csv = _write_logs(base_log_name+suf, resultados)
    return {
        "total": len(resultados),
        "ok": len([r for r in resultados if r["status"] in ("SIMULADO_VISUAL","GUARDADO")]),
        "fail": len([r for r in resultados if r["status"] == "ERROR"]),
        "log_txt": txt,
        "log_csv": csv,
        "screenshots_dir": SS_DIR
    }
