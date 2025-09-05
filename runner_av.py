# runner_av.py
# Motor Playwright llamado desde Streamlit (app.py)
# - MODOS:
#   * "PRUEBA VISUAL (navegador, sin guardar)"  -> selecciona Aula, clic en Agregar, llena modal, NO guarda
#   * "PRODUCCIÓN"                               -> llena y guarda
#
# .env:
#   AV_URL=https://aulavirtual2.autonomadeica.edu.pe/login?ReturnUrl=%2F
#   AV_VC_URL=https://aulavirtual2.autonomadeica.edu.pe/web/conference/videoconferencias
#   AV_USER=Superadmin
#   AV_PASS=tju.uzq!pgu7XGU0xrm
#   TZ=America/Lima

import os
import csv
from datetime import datetime
from typing import Dict, Any, List, Tuple

import pandas as pd
from dotenv import load_dotenv

# ===== FIX Windows loop =====
import sys, asyncio
if sys.platform.startswith("win"):
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    except Exception:
        pass

from playwright.sync_api import sync_playwright

load_dotenv()

AV_URL   = os.getenv("AV_URL",   "https://aulavirtual2.autonomadeica.edu.pe/login?ReturnUrl=%2F")
AV_VC_URL= os.getenv("AV_VC_URL","https://aulavirtual2.autonomadeica.edu.pe/web/conference/videoconferencias")
AV_USER  = os.getenv("AV_USER",  "")
AV_PASS  = os.getenv("AV_PASS",  "")
TZ       = os.getenv("TZ",       "America/Lima")

LOG_DIR = "logs"
SS_DIR  = "screenshots"
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(SS_DIR,  exist_ok=True)

# ---------------- Utils ----------------
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

    requeridas = ["CORREO","TEMA","PERIODO","FACULTAD","ESCUELA","CURSO","GRUPO",
                  "INICIO","FIN","DURACION","DIAS"]
    faltan = [c for c in requeridas if c not in t.columns]
    if faltan:
        raise RuntimeError("Faltan columnas obligatorias: " + ", ".join(faltan))

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
                f"[{r['timestamp']}] {r['status']} | {r['correo']} | TEMA: {r['tema']} | "
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

# ---------------- Login ----------------
def _login(page):
    page.goto(AV_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(300)

    # Campos login
    user_loc = page.locator(
        "input[ng-model='username'], input[placeholder='USUARIO'], input[name='username']"
    ).first
    pass_loc = page.locator(
        "input[type='password'], input[placeholder='CONTRASEÑA'], input[name='password']"
    ).first

    user_loc.wait_for(state="visible", timeout=10000)
    pass_loc.wait_for(state="visible", timeout=10000)

    # Usuario
    user_loc.fill("")
    try:
        user_loc.type(AV_USER, delay=30)
    except:
        user_loc.click()
        page.keyboard.insert_text(AV_USER)

    # Contraseña EXACTA por JS (evita mayúscula inicial / autocapitalize)
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
        AV_PASS
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

    # Ir al módulo de videoconferencias (la página de tu captura)
    try:
        page.goto(AV_VC_URL, wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle", timeout=10000)
    except:
        pass

# ---------------- Helpers página lista (Aula + Agregar) ----------------
def _select_aula(page, correo: str) -> bool:
    """
    Selecciona el 'Aula' (combo superior). Busca por label/aria/placeholder
    y soporta select2/combobox. Usa el CORREO como valor a buscar.
    """
    correo = (correo or "").strip()
    if not correo:
        return False

    # 1) select clásico asociado a label 'Aula'
    try:
        page.get_by_label("Aula", exact=False).select_option(label=correo)
        page.wait_for_timeout(150)
        return True
    except:
        pass

    # 2) combobox por aria/placeholder
    for sel in [
        "[role='combobox'][aria-label*='Aula' i]",
        "input[aria-label*='Aula' i]",
        "input[placeholder*='Aula' i]",
    ]:
        try:
            box = page.locator(sel).first
            box.click(timeout=800)
            page.keyboard.type(correo, delay=15)
            page.wait_for_timeout(150)
            page.keyboard.press("Enter")
            page.wait_for_timeout(150)
            return True
        except:
            continue

    # 3) select2 cercano a label 'Aula'
    try:
        cb = page.locator(
            "xpath=//label[contains(translate(.,'a','A'),'AULA')]/following::*"
            "[self::span[contains(@class,'select2-selection')] or self::input or self::select][1]"
        ).first
        cb.click(timeout=800)
        page.keyboard.type(correo, delay=15)
        page.wait_for_timeout(150)
        page.keyboard.press("Enter")
        page.wait_for_timeout(150)
        return True
    except:
        pass

    # 4) último intento: primer combobox visible
    try:
        cb = page.locator("[role='combobox'], .select2-selection, select").first
        cb.click(timeout=800)
        page.keyboard.type(correo, delay=15)
        page.wait_for_timeout(150)
        page.keyboard.press("Enter")
        page.wait_for_timeout(150)
        return True
    except:
        return False

def _click_agregar(page) -> bool:
    """
    Clic en el botón 'Agregar' de la lista de Videoconferencias.
    """
    for sel in [
        "button:has-text('Agregar')",
        "button:has-text('AGREGAR')",
        "[role='button']:has-text('Agregar')"
    ]:
        try:
            page.locator(sel).first.click(timeout=1500)
            return True
        except:
            continue
    try:
        page.locator("button:has(svg)").filter(has_text="").first.click(timeout=1000)
        return True
    except:
        return False

def _wait_modal(page):
    """
    Espera que aparezca un modal/form para crear la videoconferencia.
    """
    for sel in [".modal.show", ".modal-dialog", "form", "[role='dialog']"]:
        try:
            page.locator(sel).first.wait_for(state="visible", timeout=4000)
            return
        except:
            continue
    page.wait_for_timeout(400)

# ---------------- Helpers del formulario (modal) ----------------
def _safe_fill(page, label_text: str, value: Any):
    if value is None or str(value).strip() == "":
        return
    value = str(value)
    probes = [
        lambda: page.get_by_label(label_text, exact=False).fill(value),
        lambda: page.locator(f"input[placeholder*='{label_text}' i]").first.fill(value),
        lambda: page.locator(f"input[name*='{label_text.lower()}']").first.fill(value),
        lambda: page.locator(f"textarea[placeholder*='{label_text}' i]").first.fill(value),
        lambda: page.locator(f"textarea[name*='{label_text.lower()}']").first.fill(value),
    ]
    for f in probes:
        try:
            f()
            return
        except:
            continue

def _select2_like(page, root_sel: str, value: str) -> bool:
    try:
        root = page.locator(root_sel).first
        root.click(timeout=800)
        page.keyboard.type(value, delay=20)
        page.wait_for_timeout(150)
        page.keyboard.press("Enter")
        return True
    except:
        return False

def _safe_select(page, label_text: str, value: Any):
    if value is None or str(value).strip() == "":
        return
    value = str(value)
    try:
        page.get_by_label(label_text, exact=False).select_option(label=value)
        return
    except:
        pass
    for sel in [
        f"[role='combobox'][aria-label*='{label_text}' i]",
        f"input[aria-label*='{label_text}' i]",
        f"input[placeholder*='{label_text}' i]",
    ]:
        if _select2_like(page, sel, value): return
    for sel in [
        f".select2:has(label:has-text('{label_text}'))",
        f"div:has(> label:has-text('{label_text}')) .select2-selection",
        f"div:has(> label:has-text('{label_text}')) [role='combobox']",
    ]:
        if _select2_like(page, sel, value): return

def _marcar_dias(page, dias_str: str):
    if not dias_str:
        return
    partes = [d.strip() for d in str(dias_str).replace("|", ",").split(",") if d.strip()]
    DIA_MAP = {
        "1":"LUNES","2":"MARTES","3":"MIÉRCOLES","4":"JUEVES","5":"VIERNES","6":"SÁBADO","7":"DOMINGO",
        "LU":"LUNES","MA":"MARTES","MI":"MIÉRCOLES","JU":"JUEVES","VI":"VIERNES","SA":"SÁBADO","DO":"DOMINGO",
        "LUNES":"LUNES","MARTES":"MARTES","MIERCOLES":"MIÉRCOLES","MIÉRCOLES":"MIÉRCOLES",
        "JUEVES":"JUEVES","VIERNES":"VIERNES","SABADO":"SÁBADO","SÁBADO":"SÁBADO","DOMINGO":"DOMINGO",
    }
    for d in partes:
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

def _llenar_formulario(page, row: Dict[str, Any]):
    # Selects
    _safe_select(page, "Periodo",  row.get("PERIODO", ""))
    _safe_select(page, "Facultad", row.get("FACULTAD", ""))
    _safe_select(page, "Escuela",  row.get("ESCUELA", ""))
    _safe_select(page, "Curso",    row.get("CURSO", ""))
    _safe_select(page, "Grupo",    row.get("GRUPO", ""))

    # Inputs básicos
    for (label, col) in [
        ("Correo", "CORREO"),
        ("Usuario", "CORREO"),
        ("Host", "CORREO"),
        ("Tema", "TEMA"),
        ("Título", "TEMA"),
    ]:
        _safe_fill(page, label, row.get(col, ""))

    # Fechas / horas
    def fmt(dt):
        try:
            return pd.to_datetime(dt).strftime("%Y-%m-%d %H:%M")
        except:
            return ""
    _safe_fill(page, "Inicio", fmt(row.get("_INICIO_DT")))
    _safe_fill(page, "Fin",    fmt(row.get("_FIN_DT")))

    # Duración
    dur = row.get("DURACION_CALC", "") or row.get("DURACION", "")
    for label in ["Duración", "Duracion", "Minutos"]:
        _safe_fill(page, label, str(dur))

    # Días (si el formulario lo usa)
    _marcar_dias(page, row.get("DIAS", ""))

# ---------------- Runner principal ----------------
def run_batch(df: pd.DataFrame, modo: str, headless: bool) -> Dict[str, Any]:
    """
    modo:
      - "PRUEBA VISUAL (navegador, sin guardar)"
      - "PRODUCCIÓN"
    """
    if not AV_URL or not AV_USER or not AV_PASS:
        raise RuntimeError("Faltan variables de entorno AV_URL/AV_USER/AV_PASS en .env")

    t = _prep_dataframe(df)
    resultados: List[Dict[str, Any]] = []
    visual = modo.startswith("PRUEBA VISUAL")
    base_log_name = "cargamasiva_av"

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=(False if visual else headless),
            slow_mo=400,
            args=["--start-maximized"]
        )
        context = browser.new_context(
            no_viewport=True,
            locale="es-PE",
            timezone_id=TZ
        )
        page = context.new_page()
        try:
            _login(page)

            for i, r in t.iterrows():
                fila = r.to_dict()
                correo = str(fila.get("CORREO",""))
                tema   = str(fila.get("TEMA",""))

                try:
                    # 0) Seleccionar AULA (combo superior con el correo)
                    aula_ok = _select_aula(page, correo)
                    if not aula_ok:
                        # no bloqueamos, pero dejamos constancia
                        msg_aula = "No se pudo seleccionar Aula; se continúa."
                    else:
                        msg_aula = "Aula seleccionada."

                    # 1) Clic en Agregar
                    ok_add = _click_agregar(page)
                    if not ok_add:
                        raise RuntimeError("No se pudo hacer clic en 'Agregar'.")

                    # 2) Esperar modal
                    _wait_modal(page)
                    page.wait_for_timeout(300)

                    # 3) Llenar formulario
                    _llenar_formulario(page, fila)

                    # 4) Captura (siempre)
                    ss_path = os.path.join(
                        SS_DIR,
                        f"{'visual' if visual else 'prod'}_row{i+1}_{_now_tag()}.png"
                    )
                    try:
                        page.screenshot(path=ss_path, full_page=True)
                    except:
                        pass

                    if visual:
                        # 5) Cerrar modal SIN guardar
                        cerrado = False
                        for txt in ["Cerrar","Cancelar","Cancelar cambios","Salir"]:
                            try:
                                page.get_by_role("button", name=txt, exact=False).first.click(timeout=800)
                                cerrado = True
                                break
                            except:
                                try:
                                    page.get_by_text(txt, exact=False).first.click(timeout=800)
                                    cerrado = True
                                    break
                                except:
                                    continue
                        if not cerrado:
                            try:
                                page.locator("button.close, .modal-header button:has(svg), .modal-header button.close").first.click(timeout=800)
                            except:
                                pass
                        status  = "SIMULADO_VISUAL"
                        mensaje = f"Formulario llenado (NO guardado). {msg_aula}"
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
                        if guardado:
                            try:
                                page.locator(".swal-button--confirm, .swal2-confirm").first.click(timeout=20000)
                            except:
                                pass
                        status  = "GUARDADO"
                        mensaje = f"Guardado. {msg_aula}"
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
                    err_ss = os.path.join(SS_DIR, f"error_row{i+1}_{_now_tag()}.png")
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
                        "fin": str(fila.get("_INICIO_DT","")),
                        "duracion": str(fila.get("DURACION_CALC","")),
                        "dias": str(fila.get("DIAS","")),
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
    txt, csv = _write_logs("cargamasiva_av"+suf, resultados)
    return {
        "total": len(resultados),
        "ok": len([r for r in resultados if r["status"] in ("SIMULADO_VISUAL","GUARDADO")]),
        "fail": len([r for r in resultados if r["status"] == "ERROR"]),
        "log_txt": txt,
        "log_csv": csv,
        "screenshots_dir": SS_DIR
    }
