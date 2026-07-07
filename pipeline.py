"""
pipeline.py — EL CEREBRO. Es lo único que corre el cron.

Flujo de una corrida:
  1. Extrae de Yandex (por marca): funnel, sesiones/hora, ingresos/hora,
     campañas, top productos.
  2. Lee proyecciones del Excel (hojas CYBER DAYS / CYBER WOW).
  3. Escribe data.json          -> lo consume dashboard.html.
  4. Escribe Google Sheet BACKEND (Cortes + Semana, totales por día).
  5. Escribe Google Sheet META HORARIA (columnas Real por día/hora) y arma
     el bloque semana_horas del data.json (tabla por hora del dashboard).
  6. Genera PNG + correo (flujo WhatsApp, vive en correo.py).

Si algo falla: correo de alerta y sale con código != 0.
Fuera de las semanas cyber no hay proyección horaria: corre igual con meta 0.
"""
import json
import sys
import traceback
import requests
from datetime import datetime, date, timedelta
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
with open("mi_token.txt", "r", encoding="utf-8") as f:
    TOKEN = f.read().strip()

HEADERS = {"Authorization": f"OAuth {TOKEN}"}
URL_DATA   = "https://api-metrika.yandex.net/stat/v1/data"
URL_BYTIME = "https://api-metrika.yandex.net/stat/v1/data/bytime"

SALIDA_JSON  = "data.json"
SALIDA_HISTORICO = "historico.json"
ARCHIVO_PROY = "PROYECCIONES_CYBER_JULIO_2026.xlsx"
ATTRIBUTION  = "lastsign"

MARCAS = {
    "LA CURACAO":  {"contador": "98373248"},
    "TIENDAS EFE": {"contador": "98373144"},
    "JUNTOZ": {"contador": "98373308"}
}
DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
DIAS_CORTOS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]

# --- Ventanas de los cybers de julio 2026 (editables) ---
CYBER_DAYS = (date(2026, 7, 6),  date(2026, 7, 12))   # 6 al 12 julio
CYBER_WOW  = (date(2026, 7, 13), date(2026, 7, 19))   # 13 al 19 julio

# Columna base por marca en el Excel de proyecciones:
#   LA CURACAO : Lunes=B(2) ... Domingo=H(8)
#   TIENDAS EFE: Lunes=K(11) ... Domingo=Q(17)
COL_BASE = {"LA CURACAO": 2, "TIENDAS EFE": 11, "JUNTOZ": 20}


# ---------------------------------------------------------------------------
# PROYECCIONES (Excel, solo CYBER DAYS / CYBER WOW)
# ---------------------------------------------------------------------------
def hoja_evento(hoy=None):
    hoy = hoy or date.today()
    if CYBER_DAYS[0] <= hoy <= CYBER_DAYS[1]:
        return "CYBER DAYS"
    if CYBER_WOW[0] <= hoy <= CYBER_WOW[1]:
        return "CYBER WOW"
    return None


def leer_proyecciones(marca, hoy=None):
    """24 proyecciones (hora 0..23) para la marca y el día dado. Fuera de cyber: ceros."""
    hoy = hoy or date.today()
    hoja = hoja_evento(hoy)
    if hoja is None:
        return [0.0] * 24
    col = COL_BASE[marca] + hoy.weekday()
    wb = load_workbook(ARCHIVO_PROY, data_only=True)
    ws = wb[hoja]
    return [float(ws.cell(row=3 + h, column=col).value or 0) for h in range(24)]


# ---------------------------------------------------------------------------
# EXTRACCIÓN YANDEX
# ---------------------------------------------------------------------------
def traer_sesiones_hora(contador):
    """Sesiones/hora de HOY. Devuelve (serie24, corte)."""
    p = {"ids": contador, "metrics": "ym:s:visits",
         "date1": "today", "date2": "today", "group": "hour"}
    d = requests.get(URL_BYTIME, headers=HEADERS, params=p).json()
    serie = [int(v or 0) for v in d["data"][0]["metrics"][0]]
    corte = d.get("last_period_index", 23)
    return serie, corte


def traer_sesiones_hora_fecha(contador, fecha):
    """Sesiones/hora de una fecha específica (YYYY-MM-DD). Devuelve serie24."""
    p = {"ids": contador, "metrics": "ym:s:visits",
         "date1": fecha, "date2": fecha, "group": "hour"}
    d = requests.get(URL_BYTIME, headers=HEADERS, params=p).json()
    if not d.get("data"):
        return [0] * 24
    return [int(v or 0) for v in d["data"][0]["metrics"][0]]


def traer_funnel(contador, fecha="today"):
    metrics = ["ym:s:visits", "ym:s:ecommerceRevenue", "ym:s:ecommercePurchases",
               "ym:s:productImpressions", "ym:s:productBasketsQuantity",
               "ym:s:productPurchasedQuantity"]
    p = {"ids": contador, "metrics": ",".join(metrics),
         "date1": fecha, "date2": fecha}
    t = requests.get(URL_DATA, headers=HEADERS, params=p).json()["totals"]
    sesiones, ingresos, compras, vistos, carrito, comprados = t
    return {
        "ingresos": ingresos, "compras": int(compras), "vistos": int(vistos),
        "carrito": int(carrito), "comprados": int(comprados),
        "cr": (compras / sesiones) if sesiones else 0,
        "ticket": (ingresos / compras) if compras else 0,
    }


def traer_campanas(contador, top=100, fecha="today"):
    p = {"ids": contador, "dimensions": "ym:s:UTMCampaign",
         "metrics": "ym:s:visits,ym:s:ecommercePurchases,ym:s:ecommerceRevenue",
         "date1": fecha, "date2": fecha, "sort": "-ym:s:ecommerceRevenue",
         "limit": top, "attribution": ATTRIBUTION}
    d = requests.get(URL_DATA, headers=HEADERS, params=p).json()
    out = []
    for fila in d.get("data", []):
        nombre = fila["dimensions"][0].get("name") or "(sin campaña)"
        ses, compras, ingresos = fila["metrics"]
        out.append({"nombre": nombre, "sesiones": int(ses),
                    "compras": int(compras), "ingresos": round(ingresos, 2)})
    return out


def traer_top_productos(contador, top=100, fecha="today"):
    p = {"ids": contador, "dimensions": "ym:s:productName,ym:s:productID",
         "metrics": ("ym:s:productImpressions,ym:s:productBasketsQuantity,"
                     "ym:s:productPurchasedQuantity,ym:s:visits,"
                     "ym:s:productPurchasedPrice"),
         "date1": fecha, "date2": fecha,
         "sort": "-ym:s:productPurchasedQuantity", "limit": top}
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


def traer_ingresos_hora(contador, fecha="today"):
    p = {"ids": contador, "metrics": "ym:s:ecommerceRevenue",
         "date1": fecha, "date2": fecha, "group": "hour"}
    d = requests.get(URL_BYTIME, headers=HEADERS, params=p).json()
    if not d.get("data"):
        return [0.0] * 24
    serie = d["data"][0]["metrics"][0]
    return [round(float(v or 0), 2) for v in serie]


# ---------------------------------------------------------------------------
# REPORTE (crudo)
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
# PAYLOAD data.json
# ---------------------------------------------------------------------------
def construir_marca(marca):
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
            "sesiones":  r["real_acum"], "ingresos": round(f["ingresos"], 2),
            "cr": round(f["cr"], 4), "ticket": round(f["ticket"], 2),
            "compras": f["compras"], "vistos": f["vistos"],
            "carrito": f["carrito"], "comprados": f["comprados"],
        },
        "por_hora": por_hora, "campanas": campanas, "top_productos": productos,
    }
    return r, bloque


def construir_payload():
    payload = {
        "generado": datetime.now().astimezone().isoformat(timespec="seconds"),
        "corte": None, "attribution": ATTRIBUTION,
        "marcas": {}, "semana": {}, "semana_horas": {},
    }
    reportes = {}
    for marca in MARCAS:
        r, bloque = construir_marca(marca)
        payload["marcas"][marca] = bloque
        reportes[marca] = r
        payload["corte"] = r["corte"]
    return payload, reportes


# ---------------------------------------------------------------------------
# SEMANA HORAS — detalle proy/real por hora, por día del cyber, por marca
# ---------------------------------------------------------------------------
def dias_del_evento(hoy=None):
    """(nombre_hoja, [fechas desde el inicio del evento hasta hoy]) o (None, [])."""
    hoy = hoy or date.today()
    hoja = hoja_evento(hoy)
    if hoja is None:
        return None, []
    inicio = CYBER_DAYS[0] if hoja == "CYBER DAYS" else CYBER_WOW[0]
    dias, d = [], inicio
    while d <= hoy:
        dias.append(d)
        d += timedelta(days=1)
    return hoja, dias


def construir_semana_horas(reportes, hoy=None):
    """Para cada marca y cada día del evento: 24 filas {h, proy, real}.
    Hoy usa el real ya extraído (limitado al corte); días pasados se re-consultan
    a Yandex (día completo). Devuelve (nombre_hoja, dict)."""
    hoy = hoy or date.today()
    hoja, dias = dias_del_evento(hoy)
    out = {m: {} for m in MARCAS}
    if hoja is None:
        return None, out
    for marca, cfg in MARCAS.items():
        cont = cfg["contador"]
        for d in dias:
            proy = leer_proyecciones(marca, d)
            if d == hoy:
                serie, corte = reportes[marca]["real"], reportes[marca]["corte"]
            else:
                serie, corte = traer_sesiones_hora_fecha(cont, str(d)), 23
            filas = [{
                "h": h,
                "proy": round(proy[h]),
                "real": (serie[h] if h <= corte else None),
            } for h in range(24)]
            out[marca][str(d)] = {"dia": DIAS_CORTOS[d.weekday()], "horas": filas}
    return hoja, out


# ---------------------------------------------------------------------------
# HISTÓRICO (historico.json) — días YA CERRADOS, mismo shape que marcas[x]
# ---------------------------------------------------------------------------
def construir_bloque_fecha(marca, fecha):
    """Bloque de un día CERRADO (corte 23, 24h reales) con el MISMO shape que
    construir_marca(). `fecha` es un objeto date."""
    cont = MARCAS[marca]["contador"]
    fstr = str(fecha)
    proy      = leer_proyecciones(marca, fecha)
    real      = traer_sesiones_hora_fecha(cont, fstr)     # 24h completas
    funnel    = traer_funnel(cont, fstr)
    campanas  = traer_campanas(cont, fecha=fstr)
    productos = traer_top_productos(cont, fecha=fstr)
    ing_hora  = traer_ingresos_hora(cont, fecha=fstr)
    por_hora = [{
        "h": h,
        "ses_proy": round(proy[h]),
        "ses_real": real[h],
        "ing_real": ing_hora[h],
    } for h in range(24)]
    meta = round(sum(proy))
    return {
        "dia": DIAS_CORTOS[fecha.weekday()],
        "fecha": fstr,
        "meta_sesiones_dia":   meta,
        "meta_sesiones_corte": meta,          # día cerrado: corte = día completo
        "real": {
            "sesiones": sum(real), "ingresos": round(funnel["ingresos"], 2),
            "cr": round(funnel["cr"], 4), "ticket": round(funnel["ticket"], 2),
            "compras": funnel["compras"], "vistos": funnel["vistos"],
            "carrito": funnel["carrito"], "comprados": funnel["comprados"],
        },
        "por_hora": por_hora, "campanas": campanas, "top_productos": productos,
    }


def actualizar_historico(hoy=None):
    """Mantiene historico.json con los días del evento YA CERRADOS (desde el
    inicio del evento hasta AYER). Rellena los que falten y refresca solo 'ayer'
    (por si Yandex asentó tarde). Los días más antiguos ya guardados no se
    vuelven a consultar. El primer día del evento aún no cierra: no hace nada.
    Es un flujo aditivo; si falla, main() lo captura y el resto sigue igual."""
    hoy = hoy or date.today()
    hoja, dias = dias_del_evento(hoy)
    cerrados = [d for d in dias if d < hoy]

    hist = {"marcas": {m: {} for m in MARCAS}}
    try:
        with open(SALIDA_HISTORICO, "r", encoding="utf-8") as f:
            prev = json.load(f)
        if isinstance(prev.get("marcas"), dict):
            for m in MARCAS:
                hist["marcas"][m] = prev["marcas"].get(m, {})
    except FileNotFoundError:
        pass

    if cerrados:
        ayer = hoy - timedelta(days=1)
        for marca in MARCAS:
            for d in cerrados:
                key = str(d)
                if key in hist["marcas"][marca] and d != ayer:
                    continue   # ya guardado y no es 'ayer' -> no re-consultar
                hist["marcas"][marca][key] = construir_bloque_fecha(marca, d)

    hist["generado"] = datetime.now().astimezone().isoformat(timespec="seconds")
    with open(SALIDA_HISTORICO, "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=2)
    return hist


# ---------------------------------------------------------------------------
# GOOGLE SHEET "META HORARIA" — escribe la columna Real por día/hora
# ---------------------------------------------------------------------------
# Sheet: "Meta horaria Cyber Julio26"
META_SHEET_KEY = "1kucg2oRhJGWrNET_OdS-AVO9Kdc5NE7jeFJO01mT6oE"
META_TAB = {"CYBER DAYS": "Cyber days", "CYBER WOW": "Cyber wow"}
# Fila de la hora 0 de cada bloque de marca:
META_BASE_ROW = {"LA CURACAO": 3, "TIENDAS EFE": 33, "JUNTOZ": 63}


def escribir_meta_horaria(nombre_evento, semana_horas):
    """Escribe la columna Real de cada día (por marca) en el Sheet Meta horaria.
    Real col = 3 + weekday*4 (C,G,K,O,S,W,AA). VAR/%VAR son fórmulas del Sheet."""
    import gspread
    if not nombre_evento:
        return
    gc = gspread.service_account(filename="credenciales.json")
    ws = gc.open_by_key(META_SHEET_KEY).worksheet(META_TAB[nombre_evento])
    reqs = []
    for marca, dias in semana_horas.items():
        base = META_BASE_ROW[marca]
        for fecha, info in dias.items():
            w = date.fromisoformat(fecha).weekday()
            col = get_column_letter(3 + w * 4)
            valores = [[(fila["real"] if fila["real"] is not None else "")]
                       for fila in info["horas"]]
            reqs.append({"range": f"{col}{base}:{col}{base + 23}", "values": valores})
    if reqs:
        ws.batch_update(reqs, value_input_option="USER_ENTERED")


# ---------------------------------------------------------------------------
# GOOGLE SHEET "BACKEND" (gspread) — pestañas Cortes + Semana (totales/día)
# ---------------------------------------------------------------------------
NOMBRE_SHEET = "CYBER 2026 - BACKEND"


def _abrir_sheet():
    import gspread
    gc = gspread.service_account(filename="credenciales.json")
    return gc.open(NOMBRE_SHEET)


def escribir_corte(reportes):
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
    ss = _abrir_sheet()
    sh = ss.worksheet("Semana")
    filas = sh.get_all_values()[1:]
    out = {}

    def num(x):
        x = str(x).strip().replace(",", ".")
        return float(x) if x else 0.0

    for row in filas:
        if len(row) < 6:
            continue
        marca, fecha, dia, proy, real, avance = row[:6]
        out.setdefault(marca, []).append({
            "fecha": fecha, "dia": dia,
            "proy": int(num(proy)), "real": int(num(real)), "avance": num(avance),
        })
    return out


# ---------------------------------------------------------------------------
# ALERTA DE FALLO
# ---------------------------------------------------------------------------
def alerta(msg):
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

    # Sheet BACKEND (Cortes + Semana totales)
    try:
        escribir_corte(reportes)
        escribir_semana(reportes)
        payload["semana"] = leer_semana()
        print("Sheet BACKEND actualizado.")
    except Exception as e:
        print(f"Aviso: Sheet BACKEND no actualizado ({e}). El JSON sigue.")

    # Detalle por hora (semana_horas) -> JSON del dashboard
    try:
        nombre_ev, semana_horas = construir_semana_horas(reportes)
        payload["semana_horas"] = semana_horas
    except Exception as e:
        print(f"Aviso: semana_horas no construida ({e}).")
        nombre_ev, semana_horas = None, {}

    # Sheet META HORARIA (columnas Real)
    try:
        if nombre_ev:
            escribir_meta_horaria(nombre_ev, semana_horas)
            print("Sheet META HORARIA actualizado.")
    except Exception as e:
        print(f"Aviso: Meta horaria no escrita ({e}). El JSON sigue.")

    # data.json
    with open(SALIDA_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"{SALIDA_JSON} escrito.")

    # historico.json (días cerrados) — aditivo, NO bloquea el flujo en vivo
    try:
        actualizar_historico()
        print(f"{SALIDA_HISTORICO} actualizado.")
    except Exception as e:
        print(f"Aviso: {SALIDA_HISTORICO} no actualizado ({e}). El resto sigue.")

    # PNG + correo
    from correo import construir_html, render_png, construir_correo, enviar
    reps_lista = []
    for marca, r in reportes.items():
        render_png(construir_html(r), f"reporte_{marca.replace(' ', '_')}.png")
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
