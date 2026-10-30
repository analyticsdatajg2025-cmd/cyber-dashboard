"""
panel.py — CORRIDA LIVIANA para el "Panel en vivo" (pantalla TV).

Pensado para correr cada 5 min solo la noche de lanzamiento del cyber
(21:00 -> 03:00). Escribe UNICAMENTE panel.json: no toca data.json,
historico.json, neto.json, Sheets ni manda correos. El pipeline horario
(pipeline.py) sigue funcionando igual, por separado.

Por marca (las 3 por separado, gross de Yandex, sin metas):
  · KPIs del dia en curso: ventas, sesiones, transacciones, CR, ticket.
  · Funnel por sesiones (como GA4): sesiones -> con vista de producto ->
    con carrito -> con compra. Respaldo: unidades de producto.
  · Ventas, sesiones y transacciones por canal: Meta, Google, TikTok,
    Organico, Directo, Otros.
  · Top productos del dia por ventas.

Si una marca falla, se conserva su ultimo dato bueno (marcado ok=false) para
que la pantalla nunca quede en cero. Solo sale con error si fallan las 3.

Uso: python panel.py   (necesita mi_token.txt, igual que pipeline.py)
"""
import json
import re
import sys
from datetime import datetime, date, timedelta

import requests

from pipeline import (MARCAS, HEADERS, URL_BYTIME, URL_DATA, ATTRIBUTION,
                      traer_funnel, traer_top_productos)

SALIDA_PANEL = "panel.json"
HORAS_PANEL = 8          # horas por hora que se guardan (no se grafican hoy)
TOP_PANEL = 10          # productos por marca (el panel muestra 5)


def traer_horas(contador, ahora):
    """Sesiones e ingresos por hora de ayer + hoy. Devuelve (filas de las
    ultimas HORAS_PANEL horas hasta la hora actual, sesiones totales de hoy)."""
    ayer = (ahora.date() - timedelta(days=1)).isoformat()
    hoy = ahora.date().isoformat()
    p = {"ids": contador, "metrics": "ym:s:visits,ym:s:ecommerceRevenue",
         "date1": ayer, "date2": hoy, "group": "hour"}
    r = requests.get(URL_BYTIME, headers=HEADERS, params=p, timeout=60)
    d = r.json()
    if r.status_code != 200 or "data" not in d:
        raise RuntimeError(f"Yandex HTTP {r.status_code}: "
                           f"{d.get('message') or d.get('errors') or d}")
    intervalos = d.get("time_intervals") or []
    if d["data"]:
        ses, ing = d["data"][0]["metrics"][0], d["data"][0]["metrics"][1]
    else:
        ses = ing = [0] * len(intervalos)
    tope = ahora.replace(minute=0, second=0, microsecond=0)
    filas = []
    for (ini, _), s, i in zip(intervalos, ses, ing):
        t = datetime.strptime(ini[:19], "%Y-%m-%d %H:%M:%S")
        if t <= tope:
            filas.append({"fecha": t.date().isoformat(), "h": t.hour,
                          "ses": int(s or 0), "ing": round(float(i or 0), 2)})
    ses_hoy = sum(x["ses"] for x in filas if x["fecha"] == hoy)
    return filas[-HORAS_PANEL:], ses_hoy


# ---------------------------------------------------------------------------
# FUENTES DE TRAFICO -> Meta / Google / TikTok / Organico / Directo / Otros
#
# Se usan las MISMAS reglas que el "FX CANAL DETALLE" de GA4 (fuente / medio /
# campana de la sesion) para que el panel cuadre con los reportes de GA4.
# Cuando la visita no trae UTM, la fuente/medio se arman con lo que Yandex
# detecta (tipo de trafico + buscador / red social / dominio).
# ---------------------------------------------------------------------------
CANALES = ["Meta", "Google", "TikTok", "Orgánico", "Directo", "Otros"]


def _rx(patron, texto):
    return re.search(patron, texto) is not None


def canal_detalle(fuente, medio, campana):
    """Traduccion 1:1 del CASE 'FX CANAL DETALLE' de GA4 (sin la rama APP:
    el contador web de Yandex no ve la app)."""
    f, m, c = (fuente or "").lower(), (medio or "").lower(), (campana or "").lower()
    google = _rx(r"google|adwords", f)
    if google and (_rx(r"pmax|p-max|perf-?max|performance ?max", c)
                   or _rx(r"cross-network|cross_network", m)):
        return "GOOGLE ADS - PMAX"
    if google and (_rx(r"demandgen|demand[\s-]gen|discovery", c)
                   or _rx(r"discovery|demand[\s-]gen", m)):
        return "GOOGLE ADS - DEMAND GEN"
    if google and _rx(r"shopping|merchant", c) and _rx(r"cpc|ppc|paid|shopping", m):
        return "GOOGLE ADS - SHOPPING"
    if (_rx(r"google|adwords|youtube", f)
            and (_rx(r"youtube|_video_|video_views|influence", c) or _rx(r"video", m))
            and _rx(r"cpc|ppc|paid|video", m)):
        return "GOOGLE ADS - VIDEO"
    if ((google and (_rx(r"display|banner|gdn|prioridades", c)
                     or _rx(r"display|banner|banners|programatica|cpm|gd(n|s)", m)))
            or (_rx(r"social[\s-]media", f) and _rx(r"google-paid|google_paid", m))
            or _rx(r"^gd(n|s)$", m)):
        return "GOOGLE DISPLAY"
    if _rx(r"google|adwords|yandex", f) and _rx(r"cpc|ppc|paid|search-paid", m):
        return "GOOGLE SEM"
    if google and _rx(r"cpc|ppc|paid", m):
        return "GOOGLE ADS - OTROS (revisar)"
    if ((_rx(r"facebook|instagram|meta", f) or f in ("fb", "ig", "an", "msg", "fb_ig")
         or _rx(r"^fb", f)
         or (_rx(r"social[\s-]media", f) and _rx(r"facebook|instagram|fb", m)))
            and _rx(r"cpc|paid|paidsocial|paid_social|paid-social|dpa", m)):
        return "FACEBOOK ADS"
    if _rx(r"facebook-paid|facebook_paid|dpa", m) and not _rx(r"tiktok|google|adwords|yandex", f):
        return "FACEBOOK ADS"
    if ((_rx(r"tiktok", f) or _rx(r"tiktok-paid|tiktok_paid", m)
         or (_rx(r"social[\s-]media", f) and _rx(r"tiktok", m)))
            and _rx(r"cpc|paid", m)):
        return "TIKTOK ADS"
    if (_rx(r"email|mailing|correo|newsletter|edm|retention", m)
            or _rx(r"sendgrid|mailchimp|emarsys|klaviyo|hubspot|vtex|salesforce|sfmc|"
                   r"exacttarget|brevo|sendinblue|emblue|^blue$|^blue ", f)):
        return "EMAIL"
    if (_rx(r"sms|text|mms|sms_mkt|sms_marketing|mensajes", m)
            or _rx(r"sms|infobip|twilio|masiv|aldamo|masivmovil|smsmkt|smsmasivo|smsup", f)):
        return "SMS"
    if (_rx(r"wsp|whatsapp|wa\.me|whatsapp\.com|web\.whatsapp|api\.whatsapp|"
            r"business\.whatsapp", f)
            or _rx(r"wsp_mkt|whatsapp_mkt|wsp_biz|whatsapp_business|wsp|whatsapp", m)):
        return "WSP"
    if _rx(r"organic", m) and _rx(r"google|bing|yahoo|baidu|duckduckgo|ecosia|yandex|ask|seznam", f):
        return "SEO / BUSQUEDA ORG"
    if (f == "(direct)" and m == "(none)") or f == "direct":
        return "DIRECTO"
    if (_rx(r"facebook|meta", f) or f == "fb"
            or (_rx(r"social[\s-]media", f) and _rx(r"facebook|fb", m) and not _rx(r"paid", m))):
        return "FACEBOOK ORG"
    if (_rx(r"instagram", f) or f == "ig"
            or (_rx(r"social[\s-]media", f) and _rx(r"instagram", m) and not _rx(r"paid", m))):
        return "INSTAGRAM ORG"
    if (_rx(r"tiktok", f) or _rx(r"tiktok", m)
            or (_rx(r"social[\s-]media", f) and _rx(r"tiktok", m) and not _rx(r"paid", m))):
        return "TIKTOK ORG"
    if ((_rx(r"linkedin|twitter|t\.co|pinterest|youtube|pangleglobal", f) or f == "an"
         or _rx(r"social|referral", m))
            and not _rx(r"tiktok|facebook|instagram|fb", m)):
        return "OTROS SOCIAL ORG"
    return "REFERRAL"


def canal_panel(detalle):
    """Agrupa el canal detalle en los 6 canales del panel."""
    if detalle.startswith("GOOGLE"):
        return "Google"
    if detalle in ("FACEBOOK ADS", "FACEBOOK ORG", "INSTAGRAM ORG"):
        return "Meta"
    if detalle in ("TIKTOK ADS", "TIKTOK ORG"):
        return "TikTok"
    if detalle == "SEO / BUSQUEDA ORG":
        return "Orgánico"
    if detalle == "DIRECTO":
        return "Directo"
    return "Otros"


# Medio equivalente cuando la visita no trae UTM (tipo de trafico de Yandex)
_MEDIO_SIN_UTM = {"organic": "organic", "ad": "cpc", "social": "social",
                  "referral": "referral", "email": "email", "messenger": "referral",
                  "internal": "referral", "saved": "referral",
                  "recommend": "referral", "qrcode": "referral"}


def fuente_medio(tipo, motor, utm_fuente, utm_medio):
    """(fuente, medio) al estilo GA4 a partir de una fila de Yandex."""
    if utm_fuente:
        # sin utm_medium (p.ej. autoetiquetado de Google) se usa el tipo de
        # trafico que detecto Yandex: "ad" -> cpc, "organic" -> organic, etc.
        return utm_fuente, (utm_medio or _MEDIO_SIN_UTM.get(tipo, ""))
    if tipo == "direct":
        return "(direct)", "(none)"
    if tipo == "internal":
        return "(internal)", "referral"
    return (motor or "(not set)"), _MEDIO_SIN_UTM.get(tipo, "referral")


def _consulta(params):
    r = requests.get(URL_DATA, headers=HEADERS, params=params, timeout=60)
    d = r.json()
    if r.status_code != 200 or not ("data" in d or "totals" in d):
        raise RuntimeError(f"Yandex HTTP {r.status_code}: "
                           f"{d.get('message') or d.get('errors') or d}")
    return d


def traer_fuentes(contador, fecha="today"):
    """Sesiones, transacciones y ventas por canal. Devuelve
    (lista por canal del panel, lista por canal detalle, filas crudas)."""
    base = {"ids": contador, "date1": fecha, "date2": fecha,
            "attribution": ATTRIBUTION, "limit": 5000,
            "metrics": "ym:s:visits,ym:s:ecommercePurchases,ym:s:ecommerceRevenue",
            "sort": "-ym:s:visits"}
    dims = ["ym:s:<attribution>TrafficSource", "ym:s:<attribution>SourceEngine",
            "ym:s:<attribution>UTMSource", "ym:s:<attribution>UTMMedium",
            "ym:s:<attribution>UTMCampaign"]
    d = None
    for n in (5, 3, 1):          # si Yandex no acepta alguna dimension, se reduce
        try:
            d = _consulta({**base, "dimensions": ",".join(dims[:n])})
            break
        except RuntimeError as e:
            print(f"  Aviso fuentes con {n} dimensiones ({contador}): {e}",
                  file=sys.stderr)
    if d is None:
        raise RuntimeError("Yandex no acepto la consulta de fuentes")
    agg = {c: [0, 0, 0.0] for c in CANALES}
    det = {}
    crudo = []
    for fila in d.get("data") or []:
        dd = [(x or {}) for x in fila["dimensions"]] + [{}] * 5
        tipo = dd[0].get("id") or ""
        motor, utm_f, utm_m, utm_c = (dd[i].get("name") or "" for i in (1, 2, 3, 4))
        fuente, medio = fuente_medio(tipo, motor, utm_f, utm_m)
        detalle = canal_detalle(fuente, medio, utm_c)
        canal = canal_panel(detalle)
        ses, compras, ingresos = (fila["metrics"] + [0, 0, 0])[:3]
        for a in (agg[canal], det.setdefault(detalle, [0, 0, 0.0])):
            a[0] += int(ses or 0); a[1] += int(compras or 0); a[2] += float(ingresos or 0)
        if len(crudo) < 60:
            crudo.append({"tipo": tipo, "motor": motor, "utm": utm_f, "medio": utm_m,
                          "campana": utm_c, "sesiones": int(ses or 0),
                          "detalle": detalle, "canal": canal})
    out = [{"fuente": c, "sesiones": v[0], "compras": v[1],
            "ingresos": round(v[2], 2)} for c, v in agg.items()]
    detalle = sorted(({"detalle": k, "sesiones": v[0], "compras": v[1],
                       "ingresos": round(v[2], 2)} for k, v in det.items()),
                     key=lambda x: -x["sesiones"])
    return out, detalle, crudo


# ---------------------------------------------------------------------------
# FUNNEL POR SESIONES (como GA4): se cuentan SESIONES, nunca supera al paso
# anterior. Validado contra GA4 (LA CURACAO, 1 oct 2026):
#   · Yandex IGNORA los filtros por metrica (productImpressions>0 devolvia
#     todas las sesiones), por eso no se usan.
#   · Vieron producto -> sesiones con algun producto registrado
#                        (filtro de dimension ym:s:productID no nulo).
#   · Carrito / compra -> metas automaticas de ecommerce del contador
#                        (tipos e_cart / e_purchase): "sesiones con la meta".
# ---------------------------------------------------------------------------
URL_GOALS = "https://api-metrika.yandex.net/management/v1/counter/{c}/goals"
_METAS = {}                     # cache por contador: {"e_cart": id, "e_purchase": id}


def metas_ecommerce(contador):
    if contador not in _METAS:
        r = requests.get(URL_GOALS.format(c=contador), headers=HEADERS, timeout=60)
        if r.status_code != 200:
            raise RuntimeError(f"metas HTTP {r.status_code}: {r.text[:150]}")
        ids = {}
        for g in r.json().get("goals", []):
            if g.get("type") in ("e_cart", "e_purchase") and g.get("type") not in ids:
                ids[g["type"]] = g["id"]
        _METAS[contador] = ids
    return _METAS[contador]


def traer_funnel_sesiones(contador, sesiones, fecha="today"):
    metas = metas_ecommerce(contador)
    if "e_cart" not in metas or "e_purchase" not in metas:
        raise RuntimeError(f"el contador no tiene metas de ecommerce: {metas}")
    base = {"ids": contador, "date1": fecha, "date2": fecha}
    d = _consulta({**base, "metrics": f"ym:s:goal{metas['e_cart']}visits,"
                                      f"ym:s:goal{metas['e_purchase']}visits"})
    carrito, compra = [int(x or 0) for x in (d.get("totals") or [0, 0])[:2]]
    d = _consulta({**base, "metrics": "ym:s:visits", "filters": "ym:s:productID!n"})
    vistos = int((d.get("totals") or [0])[0] or 0)
    out = {"tipo": "sesiones", "sesiones": sesiones, "vistos": vistos,
           "carrito": carrito, "comprados": compra}
    # control: en un funnel de sesiones cada paso es <= al anterior. Si no se
    # cumple (p.ej. Yandex ignoro el filtro y devolvio todas las sesiones),
    # mejor caer al respaldo que mostrar un funnel incoherente.
    v = [out["sesiones"], out["vistos"], out["carrito"], out["comprados"]]
    if any(b > a * 1.02 + 5 for a, b in zip(v, v[1:])) or \
       (v[0] > 50 and v[1] == v[0]):
        raise RuntimeError(f"funnel incoherente {v}")
    return out


def construir_marca(contador, ahora):
    f = traer_funnel(contador)                       # hoy
    horas, sesiones = traer_horas(contador, ahora)
    hoy = ahora.date().isoformat()
    productos = traer_top_productos(contador)        # top 100 por unidades
    top = sorted(productos, key=lambda p: p["ingresos"], reverse=True)[:TOP_PANEL]
    try:
        fuentes, _, fuentes_crudo = traer_fuentes(contador)
    except Exception as e:                       # no tumba la marca entera
        print(f"  Aviso fuentes ({contador}): {e}", file=sys.stderr)
        fuentes, fuentes_crudo = None, None
    try:
        funnel = traer_funnel_sesiones(contador, sesiones)
        print(f"  funnel por sesiones OK ({contador})")
    except Exception as e:                       # respaldo: unidades de producto
        print(f"  Aviso funnel por sesiones ({contador}): {e} -> uso unidades",
              file=sys.stderr)
        funnel = {"tipo": "unidades", "sesiones": sesiones, "vistos": f["vistos"],
                  "carrito": f["carrito"], "comprados": f["comprados"]}
    return {
        "ok": True,
        "fecha": hoy,
        "actualizado": ahora.isoformat(timespec="seconds"),
        "kpis": {
            "ingresos": round(f["ingresos"], 2),
            "sesiones": sesiones,
            "compras": f["compras"],
            "cr": round(f["compras"] / sesiones, 4) if sesiones else 0,
            "ticket": round(f["ticket"], 2),
        },
        "funnel": funnel,
        "horas": horas,
        "fuentes": fuentes,
        "fuentes_crudo": fuentes_crudo,   # para revisar la clasificacion
        "top": [{"nombre": p["nombre"], "sku": p["sku"],
                 "comprados": p["comprados"], "ingresos": p["ingresos"]}
                for p in top],
    }


def main():
    ahora = datetime.now()
    try:
        with open(SALIDA_PANEL, "r", encoding="utf-8") as fh:
            prev = json.load(fh).get("marcas", {})
    except (FileNotFoundError, json.JSONDecodeError):
        prev = {}

    out = {"generado": ahora.astimezone().isoformat(timespec="seconds"),
           "marcas": {}}
    fallas = 0
    for marca, cfg in MARCAS.items():
        try:
            out["marcas"][marca] = construir_marca(cfg["contador"], ahora)
            k = out["marcas"][marca]["kpis"]
            print(f"  OK {marca}: S/ {k['ingresos']:,.0f} · {k['sesiones']} ses "
                  f"· {k['compras']} trx")
        except Exception as e:
            fallas += 1
            print(f"  FALLA {marca}: {type(e).__name__}: {e}", file=sys.stderr)
            viejo = prev.get(marca)
            if viejo:
                viejo["ok"] = False
                viejo["error"] = str(e)[:200]
                out["marcas"][marca] = viejo     # ultimo dato bueno

    if fallas == len(MARCAS):
        raise SystemExit("Fallaron las 3 marcas: panel.json no se actualiza.")
    with open(SALIDA_PANEL, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=2)
    print(f"{SALIDA_PANEL} escrito ({fallas} marca(s) con falla).")


if __name__ == "__main__":
    main()
