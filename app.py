# cargamasiva.py
# ------------------------------------------------------------
# Carga masiva de videoconferencias en el Aula Virtual (Zoom)
# - Lee "videoconferencias.xlsx"
# - Login con Playwright
# - Entra al módulo de Videoconferencias
# - Por cada fila: abre "Nueva videoconferencia", completa todo
# - Si MODO_PRUEBA=True: NO guarda (simulación visual)
# - Si MODO_PRUEBA=False: guarda y registra el mensaje
# - Logs: logs/log_resultados.txt y logs/log_resultados.csv
# - Capturas: screenshots/ (errores o prueba)
# ------------------------------------------------------------

import os
import csv
from datetime import datetime
from typing import Any, Dict, List

import pandas as pd
from dotenv import load_dotenv

from playwright.sync_api import sync_playwright

# ============= CONFIG =============
load_dotenv()

# Activa/desactiva simulación (NO guarda si True)
MODO_PRUEBA = True  # <-- CAMBIA a False para producción

# Excel de entrada
EXCEL_PATH = os.getenv("EXCEL_PATH", "videoconferencias.xlsx")

# URL (login te redirige a videoconferencias)
AV_URL = os.getenv(
    "AV_URL",
    "https://aulavirtual2.autonomadeica.edu.pe/login?ReturnUrl=%2F"
)
AV_VC_URL = os.getenv(
    "AV_VC_URL",
    "https://aulavirtual2.autonomadeica.edu.pe/web/conference/videoconferencias"
)

USERNAME = os.getenv("AV_USER", "superadmin")
PASSWORD = os.getenv("AV_PASS", "tju.uzq!pgu7XGU0xrm")

TZ = os.getenv("TZ", "America/Lima")

LOG_DIR = "logs"
SS_DIR  = "screenshots"
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(SS_DIR,  exist_ok=True)

TXT_LOG = os.path.join(LOG_DIR, "log_resultados.txt")

# ============= UTILIDADES =============
def now_tag() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def a_dt(x):
    try:
        v = pd.to_datetime(x)
        if pd.isna(v):
            return None
        return v
    except Exception:
        return None

def duracion_min(inicio, fin) -> int:
    if inicio is None or fin is None:
        return None
    delta = (fin - inicio).total_seconds() / 60
    if delta < 0:
        delta += 24 * 60
    return int(round(delta))

def write_logs(rows: List[Dict[str, Any]]) -> str:
    ts = now_tag()
    txt_path = os.path.join(LOG_DIR, f"log_resultados_{ts}.txt")
    csv_path = os.path.join(LOG_DIR, f"log_resultados_{ts}.csv")

    with open(txt_path, "w", encoding="utf-8") as f:
        for r in rows:
            f.write(
                f"[{r['timestamp']}] {r['status']} | {r['correo']} | TEMA: {r['tema']} | "
                f"{r['inicio']} -> {r['fin']} | {r['mensaje']}\n"
            )

    fieldnames = [
        "timestamp","status","correo","tema","periodo","facultad","escuela","curso",
        "grupo","inicio","fin","duracion","dias","mensaje","meeting_url"
    ]
    with open(csv_path, "w", encoding="utf-8", newline="") as c:
        w = csv.DictWriter(c, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fieldnames})

    return f"TXT: {txt_path} | CSV: {csv_path}"

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
            if pd.isna(v) or str(v).strip() == "":
                return duracion_min(row["_INICIO_DT"], row["_FIN_DT"])
            return int(v)
        except Exception:
            return duracion_min(row["_INICIO_DT"], row["_FIN_DT"])

    t["DURACION_CALC"] = t.apply(_dur, axis=1)
    return t

# ============= PLAYWRIGHT HELPERS =============
def login(page):
    page.goto(AV_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(300)

    # Inputs de login
    user_loc = page.locator(
        "input[ng-model='username'], input[placeholder='USUARIO'], input[name='username']"
    ).first
    pass_loc = page.locator(
        "input[type='password'], input[placeholder='CONTRASEÑA'], input[name='password']"
    ).first

    user_loc.wait_for(state="visible", timeout=10000)
    pass_loc.wait_for(state="visible", timeout=10000)

    # Usuario: tecleo para eventos
    user_loc.fill("")
    try:
        user_loc.type(USERNAME, delay=30)
    except:
        user_loc.click()
        page.keyboard.insert_text(USERNAME)

    # Contraseña: fijar EXACTO por JS (evita mayúscula inicial)
    pass_loc.evaluate(
        """(el, v) => {
            try { el.setAttribute('type','password'); } catch(e){}
            try { el.setAttribute('autocapitalize','off'); } catch(e){}
            try { el.setAttribute('autocorrect','off'); } catch(e){}
            try { el.setAttribute('autocomplete','off'); } catch(e){}
            try { el.setAttribute('spellcheck','false'); } catch(e){}
            try { el.style.textTransform = 'none'; } catch(e){}
            el.value = v;
            el.dispatchEvent(new Event('input',  { bubbles:true }));
            el.dispatchEvent(new Event('change', { bubbles:true }));
        }""",
        PASSWORD
    )

    # Botón INGRESAR
    clicked = False
    for txt in ["INGRESAR","Ingresar","Entrar","Acceder","Iniciar sesión","Login"]:
        try:
            page.get_by_role("button", name=txt, exact=False).click(timeout=1500)
            clicked = True
            break
        except:
            try:
                page.locator(f"button:has-text('{txt}')").first.click(timeout=1500)
                clicked = True
                break
            except:
                continue
    if not clicked:
        try:
            pass_loc.press("Enter")
        except:
            pass

    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except:
        page.wait_for_timeout(800)

    # Ir directo al módulo de videoconferencias
    try:
        page.goto(AV_VC_URL, wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle", timeout=10000)
    except:
        pass

def click_swal_ok(page, timeout=20000):
    # Soporta sweetalert/sweetalert2
    try:
        page.locator(".swal-button--confirm, .swal2-confirm").first.click(timeout=timeout)
    except:
        pass

def abrir_nueva_vc(page) -> bool:
    # Busca distintas variantes del botón
    for txt in [
        "Nueva videoconferencia","Nueva Videoconferencia","Nueva conferencia",
        "Crear videoconferencia","Agregar videoconferencia","Nueva","Crear"
    ]:
        try:
            page.get_by_role("button", name=txt, exact=False).first.click(timeout=1500)
            return True
        except:
            try:
                page.get_by_text(txt, exact=False).first.click(timeout=1500)
                return True
            except:
                continue
    # Fallback: ícono (+)
    try:
        page.locator("button:has(svg)").filter(has_text="").first.click(timeout=1500)
        return True
    except:
        return False

def safe_fill(page, label_text: str, value: Any):
    if value is None or str(value).strip() == "":
        return
    value = str(value)
    for try_fn in [
        lambda: page.get_by_label(label_text, exact=False).fill(value),
        lambda: page.locator(f"input[placeholder*='{label_text}' i]").first.fill(value),
        lambda: page.locator(f"input[name*='{label_text.lower()}']").first.fill(value),
        lambda: page.locator(f"textarea[placeholder*='{label_text}' i]").first.fill(value),
    ]:
        try:
            try_fn()
            return
        except:
            continue

def safe_select(page, label_text: str, value: Any):
    if value is None or str(value).strip() == "":
        return
    value = str(value)
    # <select>
    try:
        page.get_by_label(label_text, exact=False).select_option(label=value)
        return
    except:
        pass
    # combobox/autocomplete
    for try_fn in [
        lambda: page.get_by_label(label_text, exact=False).fill(value),
        lambda: page.locator(f"[role='combobox'][aria-label*='{label_text}' i]").first.fill(value),
        lambda: page.locator(f"input[aria-label*='{label_text}' i]").first.fill(value),
        lambda: page.locator(f"input[placeholder*='{label_text}' i]").first.fill(value),
    ]:
        try:
            try_fn()
            page.wait_for_timeout(120)
            page.keyboard.press("Enter")
            return
        except:
            continue

def marcar_dias(page, dias_str: str):
    if not dias_str:
        return
    dias = [d.strip() for d in str(dias_str).replace("|", ",").split(",") if d.strip()]
    DIA_MAP = {
        "1":"LUNES","2":"MARTES","3":"MIÉRCOLES","4":"JUEVES","5":"VIERNES","6":"SÁBADO","7":"DOMINGO",
        "LU":"LUNES","MA":"MARTES","MI":"MIÉRCOLES","JU":"JUEVES","VI":"VIERNES","SA":"SÁBADO","DO":"DOMINGO",
        "LUNES":"LUNES","MARTES":"MARTES","MIERCOLES":"MIÉRCOLES","MIÉRCOLES":"MIÉRCOLES","JUEVES":"JUEVES",
        "VIERNES":"VIERNES","SABADO":"SÁBADO","SÁBADO":"SÁBADO","DOMINGO":"DOMINGO",
    }
    for d in dias:
        dd = DIA_MAP.get(d.upper(), d)
        ok = False
        try:
            page.get_by_label(dd, exact=False).check()
            ok = True
        except:
            try:
                page.get_by_text(dd, exact=False).first.click()
                ok = True
            except:
                pass
        if not ok:
            try:
                page.locator(f"input[type='checkbox'][value*='{dd}' i]").first.check()
            except:
                pass

def llenar_formulario(page, row: Dict[str, Any]):
    # Selects
    safe_select(page, "Periodo",  row.get("PERIODO", ""))
    safe_select(page, "Facultad", row.get("FACULTAD", ""))
    safe_select(page, "Escuela",  row.get("ESCUELA", ""))
    safe_select(page, "Curso",    row.get("CURSO", ""))
    safe_select(page, "Grupo",    row.get("GRUPO", ""))

    # Inputs
    safe_fill(page, "Correo", row.get("CORREO", ""))
    safe_fill(page, "Tema",   row.get("TEMA", ""))

    def fmt(dt):
        try:
            return pd.to_datetime(dt).strftime("%Y-%m-%d %H:%M")
        except:
            return ""
    safe_fill(page, "Inicio", fmt(row.get("_INICIO_DT")))
    safe_fill(page, "Fin",    fmt(row.get("_FIN_DT")))

    dur = row.get("DURACION_CALC", "") or row.get("DURACION", "")
    safe_fill(page, "Duración", str(dur))
    safe_fill(page, "Duracion", str(dur))  # variante sin tilde

    # Días
    marcar_dias(page, row.get("DIAS", ""))

# ============= MAIN =============
def main():
    df = pd.read_excel(EXCEL_PATH)
    t  = prep_dataframe(df)

    resultados: List[Dict[str, Any]] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=400, args=["--start-maximized"])
        context = browser.new_context(no_viewport=True, locale="es-PE", timezone_id=TZ)
        page = context.new_page()

        try:
            # 1) Login + ir al módulo
            login(page)

            # 2) Iterar filas
            for i, r in t.iterrows():
                fila = r.to_dict()
                correo = str(fila.get("CORREO",""))
                tema   = str(fila.get("TEMA",""))

                try:
                    # 2.1) Abrir “Nueva videoconferencia”
                    ok = abrir_nueva_vc(page)
                    if not ok:
                        raise RuntimeError("No se pudo abrir el formulario 'Nueva videoconferencia'.")

                    page.wait_for_timeout(600)

                    # 2.2) Llenar
                    llenar_formulario(page, fila)

                    # 2.3) Captura del estado
                    ss_path = os.path.join(SS_DIR, f"fila{i+1}_{'preview' if MODO_PRUEBA else 'prod'}_{now_tag()}.png")
                    try:
                        page.screenshot(path=ss_path, full_page=True)
                    except:
                        pass

                    # 2.4) Guardar / o cerrar sin guardar
                    if MODO_PRUEBA:
                        # Cerrar modal sin guardar (intenta botón "Cerrar"/"Cancelar" o X)
                        cerrados = False
                        for txt in ["Cerrar","Cancelar","Cancelar cambios"]:
                            try:
                                page.get_by_role("button", name=txt, exact=False).first.click(timeout=800)
                                cerrados = True
                                break
                            except:
                                try:
                                    page.get_by_text(txt, exact=False).first.click(timeout=800)
                                    cerrados = True
                                    break
                                except:
                                    continue
                        if not cerrados:
                            try:
                                page.locator("button.close, .modal-header button:has(svg), .modal-header button.close").first.click(timeout=800)
                            except:
                                pass
                        status  = "SIMULADO_VISUAL"
                        mensaje = "Formulario llenado (NO guardado)."
                        meeting = ""
                    else:
                        # Guardar
                        guardado = False
                        for txt in ["Guardar","Crear","Crear videoconferencia","Guardar cambios","Save"]:
                            try:
                                page.get_by_role("button", name=txt, exact=False).first.click(timeout=1500)
                                guardado = True
                                break
                            except:
                                try:
                                    page.get_by_text(txt, exact=False).first.click(timeout=1500)
                                    guardado = True
                                    break
                                except:
                                    continue

                        msg = ""
                        if guardado:
                            try:
                                # Espera al popup y hace OK
                                click_swal_ok(page, timeout=20000)
                                msg = "Guardado"
                            except:
                                msg = "Guardado (sin confirmación)"
                        status  = "GUARDADO"
                        mensaje = msg
                        meeting = ""

                    resultados.append({
                        "timestamp": datetime.now().isoformat(timespec="seconds"),
                        "status": status,
                        "correo": correo,
                        "tema": tema,
                        "periodo": str(fila.get("PERIODO","")),
                        "facultad": str(fila.get("FACULTAD","")),
                        "escuela": str(fila.get("ESCUELA","")),
                        "curso": str(fila.get("CURSO","")),
                        "grupo": str(fila.get("GRUPO","")),
                        "inicio": str(fila.get("_INICIO_DT","")),
                        "fin": str(fila.get("_FIN_DT","")),
                        "duracion": str(fila.get("DURACION_CALC","")),
                        "dias": str(fila.get("DIAS","")),
                        "mensaje": mensaje,
                        "meeting_url": meeting
                    })

                except Exception as e:
                    # Captura de error
                    err_ss = os.path.join(SS_DIR, f"error_row{i+1}_{now_tag()}.png")
                    try:
                        page.screenshot(path=err_ss, full_page=True)
                    except:
                        pass

                    resultados.append({
                        "timestamp": datetime.now().isoformat(timespec="seconds"),
                        "status": "ERROR",
                        "correo": correo,
                        "tema": tema,
                        "periodo": str(fila.get("PERIODO","")),
                        "facultad": str(fila.get("FACULTAD","")),
                        "escuela": str(fila.get("ESCUELA","")),
                        "curso": str(fila.get("CURSO","")),
                        "grupo": str(fila.get("GRUPO","")),
                        "inicio": str(fila.get("_INICIO_DT","")),
                        "fin": str(fila.get("_FIN_DT","")),
                        "duracion": str(fila.get("DURACION_CALC","")),
                        "dias": str(fila.get("DIAS","")),
                        "mensaje": f"Excepción: {e}",
                        "meeting_url": ""
                    })

            # Espera final para ver la última pantalla
            page.wait_for_timeout(1200)

        finally:
            try:
                context.close()
                browser.close()
            except:
                pass

    # Escribir logs
    resumen_paths = write_logs(resultados)
    print("==== PROCESO TERMINADO ====")
    print(f"Total: {len(resultados)} | OK: {len([r for r in resultados if r['status']!='ERROR'])} | "
          f"Errores: {len([r for r in resultados if r['status']=='ERROR'])}")
    print(resumen_paths)

if __name__ == "__main__":
    main()
