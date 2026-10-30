"""
verificar_credenciales.py — CHEQUEO DE SOLO LECTURA antes de un cyber.

Prueba, con los secretos reales de GitHub, todo lo que el modo cyber necesita:
  1. Yandex      -> token valido + acceso a los 3 contadores (incluye JUNTOZ)
                    + las mismas consultas que hace pipeline.py.
  2. Google      -> Sheet mensual (metas / neto), Sheet BACKEND y Sheet
                    META HORARIA: que existan las pestanias y que la cuenta de
                    servicio tenga permiso de EDICION (sin escribir nada).
  3. Excel       -> archivo de proyecciones con sus pestanias y valores.
  4. Gmail       -> login SMTP (NO manda ningun correo).
  5. Playwright  -> genera un PNG de prueba (se sube como artifact).

NO escribe data.json / historico.json / neto.json, NO toca los Sheets y NO
manda correos. Sale con codigo 1 si algo falla.
"""
import json
import os
import traceback
from datetime import date, timedelta

import requests

RESULTADOS = []


def chequeo(nombre):
    def deco(fn):
        def run():
            try:
                detalle = fn()
                RESULTADOS.append(("OK", nombre, detalle or ""))
            except Exception as e:
                RESULTADOS.append(("FALLA", nombre, f"{type(e).__name__}: {e}"))
                traceback.print_exc()
        return run
    return deco


import pipeline as p          # lee mi_token.txt (igual que en produccion)
import proy_diaria as pd

AYER = date.today() - timedelta(days=1)


# ---------------------------------------------------------------------------
# 1. YANDEX
# ---------------------------------------------------------------------------
@chequeo("Yandex: token valido y acceso a los 3 contadores")
def yandex_contadores():
    if not p.TOKEN:
        raise RuntimeError("mi_token.txt vacio: el secreto YANDEX_TOKEN no existe o esta vacio")
    r = requests.get("https://api-metrika.yandex.net/management/v1/counters",
                     headers=p.HEADERS, params={"per_page": 1000}, timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"HTTP {r.status_code}: {r.text[:300]}")
    ids = {str(c["id"]) for c in r.json().get("counters", [])}
    faltan = [f"{m} ({c['contador']})" for m, c in p.MARCAS.items()
              if c["contador"] not in ids]
    if faltan:
        raise RuntimeError("el token NO tiene acceso a: " + ", ".join(faltan))
    return f"{len(ids)} contadores visibles, los 3 del dashboard incluidos"


def _yandex_marca(marca):
    cont = p.MARCAS[marca]["contador"]
    serie, corte = p.traer_sesiones_hora(cont)              # hoy, por hora
    f = p.traer_funnel(cont, str(AYER))
    camp = p.traer_campanas(cont, fecha=str(AYER))
    prod = p.traer_top_productos(cont, fecha=str(AYER))
    ing = p.traer_ingresos_hora(cont, fecha=str(AYER))
    ses_ayer = sum(p.traer_sesiones_hora_fecha(cont, str(AYER)))
    if ses_ayer == 0:
        raise RuntimeError(f"0 sesiones el {AYER}: el contador responde pero sin datos")
    return (f"hoy {sum(serie)} ses (corte {corte}h) | ayer {ses_ayer} ses, "
            f"S/ {f['ingresos']:,.0f}, {f['compras']} compras, "
            f"{len(camp)} campanias, {len(prod)} productos, "
            f"ingresos/hora S/ {sum(ing):,.0f}")




# ---------------------------------------------------------------------------
# 2. GOOGLE
# ---------------------------------------------------------------------------
def _email_cuenta():
    with open("credenciales.json", encoding="utf-8") as fh:
        return json.load(fh)["client_email"]


def _puede_editar(file_id):
    """Pregunta a Drive si la cuenta de servicio puede editar (no escribe)."""
    from google.oauth2 import service_account
    from google.auth.transport.requests import AuthorizedSession
    creds = service_account.Credentials.from_service_account_file(
        "credenciales.json",
        scopes=["https://www.googleapis.com/auth/drive.metadata.readonly"])
    r = AuthorizedSession(creds).get(
        f"https://www.googleapis.com/drive/v3/files/{file_id}",
        params={"fields": "name,capabilities/canEdit", "supportsAllDrives": "true"},
        timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"Drive HTTP {r.status_code}: {r.text[:200]}")
    return r.json().get("capabilities", {}).get("canEdit", False)


def _gc():
    import gspread
    return gspread.service_account(filename="credenciales.json")


@chequeo("Google: credenciales.json valido")
def google_cuenta():
    _gc()
    return f"cuenta de servicio: {_email_cuenta()}"


@chequeo("Google: Sheet mensual (metas + neto)")
def google_mensual():
    partes = []
    for marca in ("LA CURACAO", "TIENDAS EFE"):
        meta = pd.leer_proyeccion_diaria(marca, AYER)
        if not meta:
            raise RuntimeError(f"{marca}: meta del {AYER} vacia o 0 (col F)")
        neto = pd.leer_neto_diario(marca, AYER)
        partes.append(f"{marca} meta {round(meta)}"
                      + (f", venta neta {neto['venta_neta']:,.0f}" if neto else ", sin neto aun"))
    return " | ".join(partes)


@chequeo(f"Google: Sheet BACKEND '{p.NOMBRE_SHEET}'")
def google_backend():
    ss = _gc().open(p.NOMBRE_SHEET)
    tabs = [w.title for w in ss.worksheets()]
    faltan = [t for t in ("Cortes", "Semana") if t not in tabs]
    if faltan:
        raise RuntimeError(f"faltan pestanias {faltan} (hay: {tabs})")
    if not _puede_editar(ss.id):
        raise RuntimeError(f"la cuenta {_email_cuenta()} solo puede LEER: darle Editor")
    filas = len(ss.worksheet("Cortes").get_all_values())
    return f"pestanias OK, permiso de edicion OK, Cortes tiene {filas} filas"


@chequeo("Google: Sheet META HORARIA")
def google_meta_horaria():
    ss = _gc().open_by_key(p.META_SHEET_KEY)
    tabs = [w.title for w in ss.worksheets()]
    faltan = [t for t in p.META_TAB.values() if t not in tabs]
    if faltan:
        raise RuntimeError(f"faltan pestanias {faltan} (hay: {tabs})")
    if not _puede_editar(p.META_SHEET_KEY):
        raise RuntimeError(f"la cuenta {_email_cuenta()} solo puede LEER: darle Editor")
    return f"'{ss.title}': pestanias OK, permiso de edicion OK"


# ---------------------------------------------------------------------------
# 3. EXCEL DE PROYECCIONES
# ---------------------------------------------------------------------------
@chequeo(f"Excel: {p.ARCHIVO_PROY}")
def excel():
    from openpyxl import load_workbook
    wb = load_workbook(p.ARCHIVO_PROY, data_only=True)
    partes = []
    for hoja in ("CYBER DAYS", "CYBER WOW"):
        if hoja not in wb.sheetnames:
            raise RuntimeError(f"falta la pestania '{hoja}' (hay: {wb.sheetnames})")
        ws = wb[hoja]
        for marca, base in p.COL_BASE.items():
            total = sum(float(ws.cell(row=3 + h, column=base + d).value or 0)
                        for h in range(24) for d in range(7))
            if total <= 0:
                raise RuntimeError(f"{hoja} / {marca}: proyeccion vacia")
        partes.append(hoja)
    return "pestanias " + ", ".join(partes) + " con valores para las 3 marcas"


# ---------------------------------------------------------------------------
# 4. GMAIL (solo login)
# ---------------------------------------------------------------------------
@chequeo("Gmail: login SMTP (no envia nada)")
def gmail():
    import smtplib, ssl
    from correo import GMAIL_USER, GMAIL_PASS, DESTINATARIOS
    with smtplib.SMTP_SSL("smtp.gmail.com", 465,
                          context=ssl.create_default_context(), timeout=60) as s:
        s.login(GMAIL_USER, GMAIL_PASS)
    return f"login OK como {GMAIL_USER}; destinatarios: {', '.join(DESTINATARIOS)}"


# ---------------------------------------------------------------------------
# 5. PLAYWRIGHT (PNG del correo)
# ---------------------------------------------------------------------------
@chequeo("Playwright: PNG del reporte (LA CURACAO)")
def png():
    from correo import construir_html, render_png
    r = p.armar_reporte("LA CURACAO")
    render_png(construir_html(r), "prueba_reporte.png")
    tam = os.path.getsize("prueba_reporte.png")
    if tam < 1000:
        raise RuntimeError(f"PNG sospechosamente chico ({tam} bytes)")
    return f"prueba_reporte.png generado ({tam // 1024} KB) -> ver en Artifacts"


if __name__ == "__main__":
    yandex_contadores()
    for marca in p.MARCAS:
        chequeo(f"Yandex: consultas del pipeline - {marca}")(
            lambda m=marca: _yandex_marca(m))()
    for fn in (google_cuenta, google_mensual, google_backend,
               google_meta_horaria, excel, gmail, png):
        fn()
    print("\n" + "=" * 70)
    print("RESUMEN DE VERIFICACION")
    print("=" * 70)
    for estado, nombre, detalle in RESULTADOS:
        print(f"[{estado:5}] {nombre}")
        if detalle:
            print(f"          {detalle}")
    fallas = sum(1 for e, *_ in RESULTADOS if e == "FALLA")
    print("=" * 70)
    print("TODO OK" if not fallas else f"{fallas} CHEQUEO(S) FALLARON")
    raise SystemExit(1 if fallas else 0)
