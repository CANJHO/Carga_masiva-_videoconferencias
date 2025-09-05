# runner_av.py
import os
import csv
from datetime import datetime
from typing import Dict, Any, List, Tuple

import pandas as pd
from dotenv import load_dotenv

# ===== FIX Windows/Streamlit: event loop correcto =====
import sys, asyncio
if sys.platform.startswith("win"):
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    except Exception:
        pass
# =====================================================

from playwright.sync_api import sync_playwright

load_dotenv()

AV_URL   = os.getenv("AV_URL")
AV_USER  = os.getenv("AV_USER")
AV_PASS  = os.getenv("AV_PASS")
AV_VC_URL= os.getenv("AV_VC_URL", "").strip()  # opcional: URL directa a videoconferencias
TZ       = os.getenv("TZ", "America/Lima")

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

# ----------- LOGIN (usuario tecleado; contraseña exacta por JS) -----------
def _login(page):
    # 1) Navegar al login
    page.goto(AV_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(300)

    # 2) Credenciales limpias
    user = (AV_USER or "").strip()
    pwd  = (AV_PASS or "").strip()
    if not user or not pwd:
        raise RuntimeError("AV_URL/AV_USER/AV_PASS faltan. Revisa tu .env (sin comillas ni espacios).")

    # 3) Localizadores
    user_loc = page.locator(
        "input[ng-model='username'], input[placeholder='USUARIO'], input[name='username']"
    ).first
    pass_loc = page.locator(
        "input[type='password'], input[placeholder='CONTRASEÑA'], input[name='password']"
    ).first

    user_loc.wait_for(state="visible", timeout=7000)
    pass_loc.wait_for(state="visible", timeout=7000)

    # 4) Usuario: tecleo real
    user_loc.fill("")
    try:
        user_loc.type(user, delay=40)
    except:
        user_loc.click()
        page.keyboard.insert_text(user)

    # 5) Contraseña: NO teclear; fijar valor exacto por JS (evita mayúsculas automáticas)
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
        pwd
    )

    # Debug seguro
    try:
        pv = pass_loc.evaluate("el => el.value")
        eq = (pv == pwd)
        print(f"[debug] user_len={len(user)} pwd_equal={eq} pwd_len={len(pv)}")
    except Exception as e:
        print(f"[debug] no se pudo leer el input password: {e}")

    # 6) Clic en INGRESAR
    clicked = False
    for txt in ["INGRESAR", "Ingresar", "Acceder", "Entrar", "Iniciar sesión", "Iniciar sesion", "Login", "Sign in"]:
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

    # 7) Esperar navegación / mensaje
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except:
        page.wait_for_timeout(800)

    # Captura del estado
    try:
        page.screenshot(path=os.path.join(SS_DIR, f"login_state_{_now_tag()}.png"), full_page=True)
    except:
        pass

    page.wait_for_timeout(400)

# -------------------------------------------------------------------
# Ir a la pantalla de videoconferencias (después del login)
# Usa AV_VC_URL si existe; si no, intenta por texto de menú.
# -------------------------------------------------------------------
def go_to_videoconferencias(page):
    if AV_VC_URL:
        try:
            page.goto(AV_VC_URL, wait_until="domcontentloaded")
            page.wait_for_load_state("networkidle", timeout=10000)
            return
        except:
            pass
    # Fallback por navegación con textos visibles
    for txt in ["Videoconferencias", "Conferencias", "Zoom", "Reuniones"]:
        try:
            page.get_by_role("link", name=txt, exact=False).first.click(timeout=2000)
            page.wait_for_load_state("networkidle", timeout=10000)
            return
        except:
            try:
                page.get_by_text(txt, exact=False).first.click(timeout=2000)
                page.wait_for_load_state("networkidle", timeout=10000)
                return
            except:
                continue
    page.wait_for_timeout(400)

# -------------------------------------------------------------------
# Crear/Simular creación de videoconferencia
# save=False => PRUEBA VISUAL (no guarda), save=True => PRODUCCIÓN (guarda)
# row: PERIODO, FACULTAD, ESCUELA, CURSO, GRUPO, CORREO, TEMA, _INICIO_DT, _FIN_DT, DURACION_CALC, DIAS
# -------------------------------------------------------------------
def create_in_av(page, row: Dict[str, Any], save: bool = True, screenshot_path: str = "") -> Tuple[bool, str, str]:
    def fmt(dt):
        try:
            d = pd.to_datetime(dt)
            return d.strftime("%Y-%m-%d %H:%M")
        except:
            return ""

    # Ir al módulo
    try:
        go_to_videoconferencias(page)
    except:
        pass

    # 1) Abrir formulario "Nueva videoconferencia"
    abierto = False
    for txt in [
        "Nueva videoconferencia", "Nueva Videoconferencia",
        "Nueva conferencia", "Crear videoconferencia",
        "Agregar videoconferencia", "Nueva", "Crear"
    ]:
        try:
            page.get_by_role("button", name=txt, exact=False).first.click(timeout=1500)
            abierto = True
            break
        except:
            try:
                page.get_by_text(txt, exact=False).first.click(timeout=1500)
                abierto = True
                break
            except:
                continue
    if not abierto:
        try:
            page.locator("button:has(svg)").filter(has_text="").first.click(timeout=1500)
            abierto = True
        except:
            pass
    page.wait_for_timeout(600)

    # Helpers de escritura
    def safe_fill(label_text: str, value: Any):
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

    def safe_select(label_text: str, value: Any):
        if value is None or str(value).strip() == "":
            return
        value = str(value)
        try:
            page.get_by_label(label_text, exact=False).select_option(label=value)
            return
        except:
            pass
        for try_fn in [
            lambda: page.get_by_label(label_text, exact=False).fill(value),
            lambda: page.locator(f"[role='combobox'][aria-label*='{label_text}' i]").first.fill(value),
            lambda: page.locator(f"input[aria-label*='{label_text}' i]").first.fill(value),
            lambda: page.locator(f"input[placeholder*='{label_text}' i]").first.fill(value),
        ]:
            try:
                try_fn()
                page.wait_for_timeout(150)
                page.keyboard.press("Enter")
                return
            except:
                continue

    # 2) Selecciones principales
    safe_select("Periodo",  row.get("PERIODO", ""))
    safe_select("Facultad", row.get("FACULTAD", ""))
    safe_select("Escuela",  row.get("ESCUELA", ""))
    safe_select("Curso",    row.get("CURSO", ""))
    safe_select("Grupo",    row.get("GRUPO", ""))

    # 3) Datos básicos
    safe_fill("Correo", row.get("CORREO", ""))
    safe_fill("Tema",   row.get("TEMA", ""))

    # 4) Fechas / horas / duración
    safe_fill("Inicio", fmt(row.get("_INICIO_DT")))
    safe_fill("Fin",    fmt(row.get("_FIN_DT")))
    dur = row.get("DURACION_CALC", "") or row.get("DURACION", "")
    safe_fill("Duración", str(dur))
    safe_fill("Duracion", str(dur))  # por si no usan tilde

    # 5) Días
    try:
        dias_raw = str(row.get("DIAS", "")).strip()
        if dias_raw:
            dias = [d.strip() for d in dias_raw.replace("|", ",").split(",") if d.strip()]
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
    except:
        pass

    # 6) Captura del estado
    page.wait_for_timeout(600)
    if screenshot_path:
        try:
            page.screenshot(path=screenshot_path, full_page=True)
        except:
            pass

    # 7) PRUEBA VISUAL: no guardamos
    if not save:
        page.wait_for_timeout(1000)
        return True, "Simulación visual: formulario rellenado (no se guardó).", ""

    # 8) Guardar (PRODUCCIÓN)
    for txt in ["Guardar", "Crear", "Crear videoconferencia", "Guardar cambios", "Save"]:
        try:
            page.get_by_role("button", name=txt, exact=False).click(timeout=2000)
            break
        except:
            try:
                page.get_by_text(txt, exact=False).first.click(timeout=2000)
                break
            except:
                pass

    # 9) Leer mensaje / URL
    msg = ""
    try:
        msg = page.locator(".swal-text,.swal2-html-container,.alert-success").first.inner_text(timeout=8000)
    except:
        pass

    meeting_url = ""
    try:
        meeting_url = page.locator("a[target='_blank']:has-text('http')").first.get_attribute("href")
    except:
        pass

    return True, (msg or "Guardado"), (meeting_url or "")

# ------------------- Batch runner con MODOS -------------------
def run_batch(df: pd.DataFrame, modo: str, headless: bool) -> Dict[str, Any]:
    """
    modo:
      - "PRUEBA (sin navegador)"
      - "PRUEBA VISUAL (navegador, sin guardar)"
      - "PRODUCCIÓN"
    """
    if not AV_URL or not AV_USER or not AV_PASS:
        raise RuntimeError("Faltan AV_URL/AV_USER/AV_PASS en .env")

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

    # ------- PRUEBA VISUAL / PRODUCCIÓN -------
    visual = modo.startswith("PRUEBA VISUAL")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=headless,
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
                try:
                    ss_path = os.path.join(screenshots_dir, f"{'visual' if visual else 'prod'}_row{i+1}_{_now_tag()}.png")
                    ok, msg, url = create_in_av(
                        page,
                        r.to_dict(),
                        save=not visual,
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
            if visual:
                page.wait_for_timeout(5000)
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
