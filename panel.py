"""
panel.py — CORRIDA LIVIANA para el "Panel en vivo" (pantalla TV).

Pensado para correr cada ~10 min solo la noche de lanzamiento del cyber
(aprox. 23:00 -> 02:00). Escribe UNICAMENTE panel.json: no toca data.json,
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
# ---------------------------------------------------------------------------
CANALES = ["Meta", "Google", "TikTok", "Orgánico", "Directo", "Otros"]
_PATRONES = [                       # se busca en utm_source + motor/red social
    ("Meta",   ("facebook", "instagram", "meta", "fb", "ig", "messenger")),
    ("Google", ("google", "youtube", "gclid", "pmax", "adwords")),
    ("TikTok", ("tiktok",)),
]


def _canal(tipo, motor, utm):
    """Clasifica una fila de Yandex en un canal del panel.
    tipo  = id de TrafficSource (ad, organic, direct, social, referral, ...)
    motor = SourceEngine (Google, Facebook, TikTok, ...)
    utm   = utm_source."""
    texto = f"{motor} {utm}".lower()
    pagado = tipo == "ad" or bool(utm)
    if tipo == "direct" and not utm:
        return "Directo"
    if tipo == "organic" and not utm:
        return "Orgánico"            # SEO: Google organico va aqui, no en Google
    palabras = set(re.split(r"[^a-z0-9]+", texto))
    for canal, claves in _PATRONES:
        # claves cortas (fb, ig) solo como palabra completa; largas, en cualquier parte
        if any(k in palabras or (len(k) >= 5 and k in texto) for k in claves):
            if canal == "Google" and not pagado:
                return "Orgánico"
            return canal
    return "Otros"


def _consulta(params):
    r = requests.get(URL_DATA, headers=HEADERS, params=params, timeout=60)
    d = r.json()
    if r.status_code != 200 or not ("data" in d or "totals" in d):
        raise RuntimeError(f"Yandex HTTP {r.status_code}: "
                           f"{d.get('message') or d.get('errors') or d}")
    return d


def traer_fuentes(contador, fecha="today"):
    """Sesiones, transacciones y ventas de hoy por canal. Devuelve
    (lista por canal, filas crudas para revisar la clasificacion)."""
    base = {"ids": contador, "date1": fecha, "date2": fecha,
            "attribution": ATTRIBUTION, "limit": 500,
            "metrics": "ym:s:visits,ym:s:ecommercePurchases,ym:s:ecommerceRevenue",
            "sort": "-ym:s:visits"}
    dims = ["ym:s:<attribution>TrafficSource", "ym:s:<attribution>SourceEngine",
            "ym:s:<attribution>UTMSource"]
    try:
        d = _consulta({**base, "dimensions": ",".join(dims)})
    except RuntimeError as e:
        # si Yandex no acepta alguna dimension, al menos por tipo de trafico
        print(f"  Aviso fuentes detalladas ({contador}): {e}", file=sys.stderr)
        d = _consulta({**base, "dimensions": dims[0]})
    agg = {c: [0, 0, 0.0] for c in CANALES}
    crudo = []
    for fila in d.get("data") or []:
        dd = fila["dimensions"] + [{}, {}]
        tipo = (dd[0] or {}).get("id") or ""
        motor = (dd[1] or {}).get("name") or ""
        utm = (dd[2] or {}).get("name") or ""
        ses, compras, ingresos = fila["metrics"]
        canal = _canal(tipo, motor, utm)
        a = agg[canal]
        a[0] += int(ses or 0); a[1] += int(compras or 0); a[2] += float(ingresos or 0)
        if len(crudo) < 60:
            crudo.append({"tipo": tipo, "motor": motor, "utm": utm,
                          "sesiones": int(ses or 0), "canal": canal})
    out = [{"fuente": c, "sesiones": v[0], "compras": v[1],
            "ingresos": round(v[2], 2)} for c, v in agg.items()]
    return out, crudo


# ---------------------------------------------------------------------------
# FUNNEL POR SESIONES (como GA4): sesiones que vieron producto / agregaron al
# carrito / compraron. Se cuentan sesiones, no fichas: nunca supera a sesiones.
# ---------------------------------------------------------------------------
_FILTROS_FUNNEL = {
    "vistos":    "ym:s:productImpressions>0",
    "carrito":   "ym:s:productBasketsQuantity>0",
    "comprados": "ym:s:ecommercePurchases>0",
}


def traer_funnel_sesiones(contador, sesiones, fecha="today"):
    out = {"tipo": "sesiones", "sesiones": sesiones}
    for paso, filtro in _FILTROS_FUNNEL.items():
        d = _consulta({"ids": contador, "date1": fecha, "date2": fecha,
                       "metrics": "ym:s:visits", "filters": filtro})
        out[paso] = int((d.get("totals") or [0])[0] or 0)
    # control: en un funnel de sesiones cada paso es <= al anterior. Si no se
    # cumple (p.ej. Yandex ignoro el filtro y devolvio todas las sesiones),
    # mejor caer al respaldo que mostrar un funnel incoherente.
    v = [out["sesiones"], out["vistos"], out["carrito"], out["comprados"]]
    if any(b > a * 1.02 + 5 for a, b in zip(v, v[1:])) or \
       (v[0] > 50 and v[1] == v[2] == v[3] == v[0]):
        raise RuntimeError(f"funnel incoherente {v}")
    return out


def construir_marca(contador, ahora):
    f = traer_funnel(contador)                       # hoy
    horas, sesiones = traer_horas(contador, ahora)
    hoy = ahora.date().isoformat()
    productos = traer_top_productos(contador)        # top 100 por unidades
    top = sorted(productos, key=lambda p: p["ingresos"], reverse=True)[:TOP_PANEL]
    try:
        fuentes, fuentes_crudo = traer_fuentes(contador)
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
