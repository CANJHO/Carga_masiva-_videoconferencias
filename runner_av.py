# cargamasiva_before.py
# ------------------------------------------------------------
# Versión simple (antes): solo TXT y capturas de error.
# Mantiene MODO_PRUEBA para no guardar cuando quieras simular.
# ------------------------------------------------------------

import os
from datetime import datetime
from typing import Any, Dict, List

import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()

MODO_PRUEBA = True  # <-- CAMBIA a False para guardar

EXCEL_PATH = os.getenv("EXCEL_PATH", "videoconferencias.xlsx")
AV_URL     = os.getenv("AV_URL",    "https://aulavirtual2.autonomadeica.edu.pe/login?ReturnUrl=%2F")
AV_VC_URL  = os.getenv("AV_VC_URL", "https://aulavirtual2.autonomadeica.edu.pe/web/conference/videoconferencias")
USERNAME   = os.getenv("AV_USER", "superadmin")
PASSWORD   = os.getenv("AV_PASS", "tju.uzq!pgu7XGU0xrm")
TZ         = os.getenv("TZ", "America/Lima")

LOG_DIR = "logs"
SS_DIR  = "screenshots"
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(SS_DIR,  exist_ok=True)

TXT_LOG = os.path.join(LOG_DIR, "log_resultados.txt")

def now_tag() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def a_dt(x):
    try:
        v = pd.to_datetime(x)
        if pd.isna(v): return None
        return v
    except Exception:
        return None

def duracion_min(inicio, fin) -> int:
    if inicio is None or fin is None:
        return None
    delta = (fin - inicio).total_seconds() / 60
    if delta < 0:
        delta += 24*60
    return int(round(delta))

def prep_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    t = df.copy()
    t.columns = [c.upper().strip() for c in t.columns]
    req = ["CORREO","TEMA","PERIODO","FACULTAD","ESCUELA","CURSO","GRUPO",
           "INICIO","FIN","DURACION","DIAS"]
    faltan = [c for c in req if c not in t.columns]
    if faltan:
        raise RuntimeError("Faltan columnas obligatorias: " + ", ".join(faltan))
    t["_INICIO_DT"] = t["INICIO"].apply(a_dt)
    t["_FIN_DT"]    = t["FIN"].apply(a_dt)
    def _dur(row):
        v = row.get("DURACION")
        try:
            if pd.isna(v) or str(v).strip()=="":
                return duracion_min(row["_INICIO_DT"], row["_FIN_DT"])
            return int(v)
        except Exception:
            return duracion_min(row["_INICIO_DT"], row["_FIN_DT"])
    t["DURACION_CALC"] = t.apply(_dur, axis=1)
    return t

def login(page):
    page.goto(AV_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(300)
    user_loc = page.locator(
        "input[ng-model='username'], input[placeholder='USUARIO'], input[name='username']"
    ).first
    pass_loc = page.locator(
        "input[type='password'], input[placeholder='CONTRASEÑA'], input[name='password']"
    ).first
    user_loc.wait_for(state="visible", timeout=10000)
    pass_loc.wait_for(state="visible", timeout=10000)
    user_loc.fill("")
    try:
        user_loc.type(USERNAME, delay=30)
    except:
        user_loc.click(); page.keyboard.insert_text(USERNAME)
    pass_loc.evaluate(
        """(el, v) => { el.value=v; el.dispatchEvent(new Event('input',{bubbles:true})); el.dispatchEvent(new Event('change',{bubbles:true})); }""",
        PASSWORD
    )
    # Ingresar
    try:
        page.get_by_role("button", name="INGRESAR", exact=False).click(timeout=1500)
    except:
        pass_loc.press("Enter")
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except:
        page.wait_for_timeout(800)
    try:
        page.goto(AV_VC_URL, wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle", timeout=10000)
    except:
        pass

def abrir_nueva(page):
    for txt in ["Nueva videoconferencia","Nueva conferencia","Crear videoconferencia","Nueva","Crear"]:
        try:
            page.get_by_role("button", name=txt, exact=False).first.click(timeout=1500)
            return True
        except:
            try:
                page.get_by_text(txt, exact=False).first.click(timeout=1500)
                return True
            except:
                continue
    try:
        page.locator("button:has(svg)").filter(has_text="").first.click(timeout=1500)
        return True
    except:
        return False

def safe_fill(page, label_text: str, value: Any):
    if value is None or str(value).strip() == "": return
    value = str(value)
    for try_fn in [
        lambda: page.get_by_label(label_text, exact=False).fill(value),
        lambda: page.locator(f"input[placeholder*='{label_text}' i]").first.fill(value),
        lambda: page.locator(f"input[name*='{label_text.lower()}']").first.fill(value),
        lambda: page.locator(f"textarea[placeholder*='{label_text}' i]").first.fill(value),
    ]:
        try:
            try_fn(); return
        except: continue

def safe_select(page, label_text: str, value: Any):
    if value is None or str(value).strip() == "": return
    value = str(value)
    try:
        page.get_by_label(label_text, exact=False).select_option(label=value); return
    except: pass
    for try_fn in [
        lambda: page.get_by_label(label_text, exact=False).fill(value),
        lambda: page.locator(f"[role='combobox'][aria-label*='{label_text}' i]").first.fill(value),
        lambda: page.locator(f"input[placeholder*='{label_text}' i]").first.fill(value),
    ]:
        try:
            try_fn(); page.keyboard.press("Enter"); return
        except: continue

def marcar_dias(page, dias: str):
    if not dias: return
    piezas = [d.strip() for d in str(dias).replace("|",",").split(",") if d.strip()]
    for d in piezas:
        try:
            page.get_by_label(d, exact=False).check()
        except:
            try:
                page.get_by_text(d, exact=False).first.click()
            except:
                pass

def main():
    df = pd.read_excel(EXCEL_PATH)
    t  = prep_dataframe(df)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=400, args=["--start-maximized"])
        context = browser.new_context(no_viewport=True, locale="es-PE", timezone_id=TZ)
        page = context.new_page()
        try:
            login(page)

            with open(TXT_LOG, "a", encoding="utf-8") as f:
                for i, r in t.iterrows():
                    fila = r.to_dict()
                    try:
                        if not abrir_nueva(page):
                            raise RuntimeError("No pudo abrir 'Nueva videoconferencia'.")

                        def fmt(dt):
                            try: return pd.to_datetime(dt).strftime("%Y-%m-%d %H:%M")
                            except: return ""
                        safe_select(page, "Periodo",  fila.get("PERIODO",""))
                        safe_select(page, "Facultad", fila.get("FACULTAD",""))
                        safe_select(page, "Escuela",  fila.get("ESCUELA",""))
                        safe_select(page, "Curso",    fila.get("CURSO",""))
                        safe_select(page, "Grupo",    fila.get("GRUPO",""))
                        safe_fill(page, "Correo", fila.get("CORREO",""))
                        safe_fill(page, "Tema",   fila.get("TEMA",""))
                        safe_fill(page, "Inicio", fmt(fila.get("_INICIO_DT")))
                        safe_fill(page, "Fin",    fmt(fila.get("_FIN_DT")))
                        dur = fila.get("DURACION_CALC","") or fila.get("DURACION","")
                        safe_fill(page, "Duración", str(dur))
                        marcar_dias(page, fila.get("DIAS",""))

                        # Captura
                        try:
                            page.screenshot(path=os.path.join(SS_DIR, f"fila{i+1}_{'preview' if MODO_PRUEBA else 'prod'}_{now_tag()}.png"), full_page=True)
                        except: pass

                        if MODO_PRUEBA:
                            # Cerrar sin guardar
                            for txt in ["Cerrar","Cancelar","Cancelar cambios"]:
                                try: page.get_by_role("button", name=txt, exact=False).first.click(timeout=800); break
                                except:
                                    try: page.get_by_text(txt, exact=False).first.click(timeout=800); break
                                    except: pass
                            estado="SIMULADO_VISUAL"; msg="NO guardado"
                        else:
                            # Guardar
                            for txt in ["Guardar","Crear","Guardar cambios","Save"]:
                                try: page.get_by_role("button", name=txt, exact=False).first.click(timeout=1500); break
                                except:
                                    try: page.get_by_text(txt, exact=False).first.click(timeout=1500); break
                                    except: pass
                            estado="GUARDADO"; msg="Guardado"

                        f.write(f"[{datetime.now().isoformat(sep=' ', timespec='seconds')}] {estado} | {fila.get('CORREO','')} | {fila.get('TEMA','')} | {fila.get('_INICIO_DT','')} -> {fila.get('_FIN_DT','')} | {msg}\n")

                    except Exception as e:
                        try:
                            page.screenshot(path=os.path.join(SS_DIR, f"error_row{i+1}_{now_tag()}.png"), full_page=True)
                        except: pass
                        f.write(f"[{datetime.now().isoformat(sep=' ', timespec='seconds')}] ERROR | {fila.get('CORREO','')} | {fila.get('TEMA','')} | Excepción: {e}\n")

            page.wait_for_timeout(1200)
        finally:
            try: context.close(); browser.close()
            except: pass

if __name__ == "__main__":
    main()
