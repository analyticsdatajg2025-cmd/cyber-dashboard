"""
pipeline.py — EL CEREBRO. Es lo único que corre el cron.

Fusiona lo que antes estaba repartido en:
  reporte_datos.py + extras.py + build_data.py + escribir_sheet.py + main.py

Flujo de una corrida:
  1. Extrae todo de Yandex (una sola vez por marca): funnel, sesiones/hora,
     ingresos/hora, campañas, top productos.
  2. Lee las proyecciones del Excel (solo hojas CYBER DAYS / CYBER WOW).
  3. Escribe data.json      -> lo consume dashboard.html.
  4. Escribe Google Sheet   -> pestañas Cortes + Semana (gerencia).
  5. Genera PNG + correo    -> flujo WhatsApp (vive en correo.py, no cambia).

Si algo falla: manda correo de alerta y sale con código != 0
(para que GitHub Actions marque el run como fallido).

Fuera de las dos semanas de cyber no hay proyecciones horarias: el pipeline
corre igual, solo que la meta queda en 0 (el dashboard lo maneja sin romperse).
"""
import json
import sys
import traceback
import requests
from datetime import datetime, date
from openpyxl import load_workbook

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
with open("mi_token.txt", "r", encoding="utf-8") as f:
    TOKEN = f.read().strip()

HEADERS = {"Authorization": f"OAuth {TOKEN}"}
URL_DATA   = "https://api-metrika.yandex.net/stat/v1/data"
URL_BYTIME = "https://api-metrika.yandex.net/stat/v1/data/bytime"

SALIDA_JSON  = "data.json"
ARCHIVO_PROY = "PROYECCIONES_CYBER_JULIO_2026.xlsx"

# Modelo de atribución que ve gerencia por defecto en Yandex
ATTRIBUTION = "lastsign"

MARCAS = {
    "LA CURACAO":  {"contador": "98373248"},
    "TIENDAS EFE": {"contador": "98373144"},
}
DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

# --- Ventanas de los cybers de julio 2026 (editables) ---
CYBER_DAYS = (date(2026, 7, 6),  date(2026, 7, 12))   # 6 al 12 julio
CYBER_WOW  = (date(2026, 7, 13), date(2026, 7, 19))   # 13 al 19 julio

# Columna base por marca en las hojas CYBER (weekday Lun=0 -> primera col de datos):
#   LA CURACAO : HORA=A, LUNES=B(2) ... DOMINGO=H(8)
#   TIENDAS EFE: HORA=J, LUNES=K(11) ... DOMINGO=Q(17)
COL_BASE = {"LA CURACAO": 2, "TIENDAS EFE": 11}


# ---------------------------------------------------------------------------
# PROYECCIONES (solo CYBER DAYS / CYBER WOW)
# ---------------------------------------------------------------------------
def hoja_evento(hoy=None):
    """Devuelve el nombre de la hoja según la fecha, o None si no es semana cyber."""
    hoy = hoy or date.today()
    if CYBER_DAYS[0] <= hoy <= CYBER_DAYS[1]:
        return "CYBER DAYS"
    if CYBER_WOW[0] <= hoy <= CYBER_WOW[1]:
        return "CYBER WOW"
    return None


def leer_proyecciones(marca, hoy=None):
    """Lista de 24 proyecciones (hora 0..23) para la marca y el día de hoy.
    Fuera de semana cyber devuelve 24 ceros."""
    hoy = hoy or date.today()
    hoja = hoja_evento(hoy)
    if hoja is None:
        return [0.0] * 24
    col = COL_BASE[marca] + hoy.weekday()      # Lun=0 -> col base
    wb = load_workbook(ARCHIVO_PROY, data_only=True)
    ws = wb[hoja]
    return [float(ws.cell(row=3 + h, column=col).value or 0) for h in range(24)]


# ---------------------------------------------------------------------------
# EXTRACCIÓN YANDEX
# ---------------------------------------------------------------------------
def traer_sesiones_hora(contador):
    p = {"ids": contador, "metrics": "ym:s:visits",
         "date1": "today", "date2": "today", "group": "hour"}
    d = requests.get(URL_BYTIME, headers=HEADERS, params=p).json()
    serie = [int(v or 0) for v in d["data"][0]["metrics"][0]]
    corte = d.get("last_period_index", 23)
    return serie, corte


def traer_funnel(contador):
    metrics = ["ym:s:visits", "ym:s:ecommerceRevenue", "ym:s:ecommercePurchases",
               "ym:s:productImpressions", "ym:s:productBasketsQuantity",
               "ym:s:productPurchasedQuantity"]
    p = {"ids": contador, "metrics": ",".join(metrics),
         "date1": "today", "date2": "today"}
    t = requests.get(URL_DATA, headers=HEADERS, params=p).json()["totals"]
    sesiones, ingresos, compras, vistos, carrito, comprados = t
    return {
        "ingresos": ingresos,
        "compras": int(compras),
        "vistos": int(vistos),
        "carrito": int(carrito),
        "comprados": int(comprados),
        "cr": (compras / sesiones) if sesiones else 0,
        "ticket": (ingresos / compras) if compras else 0,
    }


def traer_campanas(contador, top=100):
    p = {
        "ids": contador,
        "dimensions": "ym:s:UTMCampaign",
        "metrics": "ym:s:visits,ym:s:ecommercePurchases,ym:s:ecommerceRevenue",
        "date1": "today", "date2": "today",
        "sort": "-ym:s:ecommerceRevenue",
        "limit": top,
        "attribution": ATTRIBUTION,
    }
    d = requests.get(URL_DATA, headers=HEADERS, params=p).json()
    out = []
    for fila in d.get("data", []):
        nombre = fila["dimensions"][0].get("name") or "(sin campaña)"
        ses, compras, ingresos = fila["metrics"]
        out.append({"nombre": nombre, "sesiones": int(ses),
                    "compras": int(compras), "ingresos": round(ingresos, 2)})
    return out


def traer_top_productos(contador, top=100):
    p = {
        "ids": contador,
        "dimensions": "ym:s:productName,ym:s:productID",
        "metrics": ("ym:s:productImpressions,ym:s:productBasketsQuantity,"
                    "ym:s:productPurchasedQuantity,ym:s:visits,"
                    "ym:s:productPurchasedPrice"),
        "date1": "today", "date2": "today",
        "sort": "-ym:s:productPurchasedQuantity",
        "limit": top,
    }
    d = requests.get(URL_DATA, headers=HEADERS, params=p).json()
    out = []
    for fila in d.get("data", []):
        nombre = fila["dimensions"][0].get("name") or "(sin nombre)"
        sku = fila["dimensions"][1].get("name") or ""
        vistos, carrito, comprados, ses, ingresos = fila["metrics"]
        out.append({"nombre": nombre, "sku": sku, "vistos": int(vistos),
                    "carrito": int(carrito), "comprados": int(comprados),
                    "sesiones": int(ses), "ingresos": round(ingresos, 2)})
    return out


def traer_ingresos_hora(contador):
    p = {"ids": contador, "metrics": "ym:s:ecommerceRevenue",
         "date1": "today", "date2": "today", "group": "hour"}
    d = requests.get(URL_BYTIME, headers=HEADERS, params=p).json()
    if not d.get("data"):
        return [0.0] * 24
    serie = d["data"][0]["metrics"][0]
    return [round(float(v or 0), 2) for v in serie]


# ---------------------------------------------------------------------------
# REPORTE (crudo) — lo consumen el PNG/correo y el payload del dashboard
# ---------------------------------------------------------------------------
def armar_reporte(marca):
    cfg = MARCAS[marca]
    dia_idx = date.today().weekday()
    proy = leer_proyecciones(marca)
    real, corte = traer_sesiones_hora(cfg["contador"])
    funnel = traer_funnel(cfg["contador"])
    real_acum = sum(real[:corte + 1])
    meta = sum(proy)
    return {
        "marca": marca, "dia": DIAS[dia_idx], "fecha": str(date.today()),
        "corte": corte, "proy": proy, "real": real,
        "meta": meta, "real_acum": real_acum,
        "avance": (real_acum / meta) if meta else 0,
        "funnel": funnel,
    }


# ---------------------------------------------------------------------------
# PAYLOAD data.json (lo lee dashboard.html)
# ---------------------------------------------------------------------------
def construir_marca(marca):
    """Devuelve (reporte_crudo, bloque_json). El crudo se reusa para el PNG."""
    r = armar_reporte(marca)
    contador = MARCAS[marca]["contador"]

    campanas  = traer_campanas(contador)
    productos = traer_top_productos(contador)
    ing_hora  = traer_ingresos_hora(contador)

    corte = r["corte"]
    por_hora = [{
        "h": h,
        "ses_proy": round(r["proy"][h]),
        "ses_real": r["real"][h] if h <= corte else None,
        "ing_real": ing_hora[h]   if h <= corte else None,
    } for h in range(24)]

    f = r["funnel"]
    bloque = {
        "meta_sesiones_dia":   round(r["meta"]),
        "meta_sesiones_corte": round(sum(r["proy"][:corte + 1])),
        "real": {
            "sesiones":  r["real_acum"],
            "ingresos":  round(f["ingresos"], 2),
            "cr":        round(f["cr"], 4),
            "ticket":    round(f["ticket"], 2),
            "compras":   f["compras"],
            "vistos":    f["vistos"],
            "carrito":   f["carrito"],
            "comprados": f["comprados"],
        },
        "por_hora": por_hora,
        "campanas": campanas,
        "top_productos": productos,
    }
    return r, bloque


def construir_payload():
    """Recorre todas las marcas una sola vez. Devuelve (payload, reportes_crudos)."""
    payload = {
        "generado": datetime.now().astimezone().isoformat(timespec="seconds"),
        "corte": None,
        "attribution": ATTRIBUTION,
        "marcas": {},
        "semana": {},
    }
    reportes = {}
    for marca in MARCAS:
        r, bloque = construir_marca(marca)
        payload["marcas"][marca] = bloque
        reportes[marca] = r
        payload["corte"] = r["corte"]
    return payload, reportes


# ---------------------------------------------------------------------------
# GOOGLE SHEET (gspread) — pestañas Cortes + Semana
# ---------------------------------------------------------------------------
NOMBRE_SHEET = "CYBER 2026 - BACKEND"    # <-- ajusta al nombre real de tu Sheet
DIAS_CORTOS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]


def _abrir_sheet():
    import gspread
    gc = gspread.service_account(filename="credenciales.json")
    return gc.open(NOMBRE_SHEET)


def escribir_corte(reportes):
    """Escribe/sobrescribe la fila del corte por marca (idempotente marca+fecha+hora)."""
    ss = _abrir_sheet()
    sh = ss.worksheet("Cortes")
    filas = sh.get_all_values()
    hoy = str(date.today())
    for marca, r in reportes.items():
        hora = f"{r['corte']:02d}:00"
        f = r["funnel"]
        nueva = [marca, hoy, hora, r["real_acum"], round(f["ingresos"], 2),
                 round(f["cr"], 4), round(f["ticket"], 2), round(r["avance"], 4)]
        idx = None
        for i, row in enumerate(filas[1:], start=2):
            if len(row) >= 3 and row[0] == marca and row[1] == hoy and row[2] == hora:
                idx = i
                break
        if idx:
            sh.update(f"A{idx}:H{idx}", [nueva])
        else:
            sh.append_row(nueva, value_input_option="USER_ENTERED")


def escribir_semana(reportes):
    """Actualiza el total del día por marca (idempotente marca+fecha)."""
    ss = _abrir_sheet()
    sh = ss.worksheet("Semana")
    filas = sh.get_all_values()
    hoy = str(date.today())
    dia = DIAS_CORTOS[date.today().weekday()]
    for marca, r in reportes.items():
        proy_corte = round(sum(r["proy"][:r["corte"] + 1]))
        nueva = [marca, hoy, dia, round(r["meta"]), r["real_acum"],
                 round(r["real_acum"] / proy_corte, 4) if proy_corte else 0]
        idx = None
        for i, row in enumerate(filas[1:], start=2):
            if len(row) >= 2 and row[0] == marca and row[1] == hoy:
                idx = i
                break
        if idx:
            sh.update(f"A{idx}:F{idx}", [nueva])
        else:
            sh.append_row(nueva, value_input_option="USER_ENTERED")


def leer_semana():
    """Lee la pestaña Semana para inyectarla en data.json (vista Semana del dashboard)."""
    ss = _abrir_sheet()
    sh = ss.worksheet("Semana")
    filas = sh.get_all_values()[1:]
    out = {}
    for row in filas:
        if len(row) < 6:
            continue
        marca, fecha, dia, proy, real, avance = row[:6]
        def num(x):
            x = str(x).strip().replace(",", ".")
            return float(x) if x else 0.0
        out.setdefault(marca, []).append({
            "fecha": fecha, "dia": dia,
            "proy": int(num(proy)),
            "real": int(num(real)),
            "avance": num(avance),
        })
    return out


# ---------------------------------------------------------------------------
# ALERTA DE FALLO
# ---------------------------------------------------------------------------
def alerta(msg):
    """Correo simple de alerta si el pipeline truena."""
    try:
        from email.message import EmailMessage
        import smtplib, ssl
        from correo import GMAIL_USER, GMAIL_PASS, DESTINATARIOS
        m = EmailMessage()
        m["Subject"] = "FALLO pipeline Cyber - revisar"
        m["From"] = GMAIL_USER
        m["To"] = ", ".join(DESTINATARIOS)
        m.set_content("El pipeline de Cyber falló:\n\n" + msg)
        ctx = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ctx) as s:
            s.login(GMAIL_USER, GMAIL_PASS)
            s.send_message(m)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    print("Extrayendo de Yandex...")
    payload, reportes = construir_payload()

    # Google Sheet (si falla, el JSON igual se escribe)
    try:
        escribir_corte(reportes)
        escribir_semana(reportes)
        payload["semana"] = leer_semana()
        print("Sheet actualizado.")
    except Exception as e:
        print(f"Aviso: Sheet no actualizado ({e}). El JSON sigue.")

    # data.json
    with open(SALIDA_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"{SALIDA_JSON} escrito.")

    # PNG + correo (flujo WhatsApp, vive en correo.py)
    from correo import construir_html, render_png, construir_correo, enviar
    reps_lista = []
    for marca, r in reportes.items():
        html = construir_html(r)
        render_png(html, f"reporte_{marca.replace(' ', '_')}.png")
        reps_lista.append(r)
    enviar(construir_correo(reps_lista))
    print("PNG + correo enviados.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        err = traceback.format_exc()
        print(err, file=sys.stderr)
        alerta(err)
        sys.exit(1)
