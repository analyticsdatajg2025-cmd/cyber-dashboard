"""
panel.py — CORRIDA LIVIANA para el "Panel en vivo" (pantalla TV).

Pensado para correr cada ~10 min solo la noche de lanzamiento del cyber
(aprox. 23:00 -> 02:00). Escribe UNICAMENTE panel.json: no toca data.json,
historico.json, neto.json, Sheets ni manda correos. El pipeline horario
(pipeline.py) sigue funcionando igual, por separado.

Por marca (las 3 por separado, gross de Yandex, sin metas):
  · KPIs del dia en curso: ventas, sesiones, transacciones, CR, ticket.
  · Funnel en unidades de producto: vistos -> al carrito -> comprados.
  · Ventas, sesiones y transacciones por tipo de fuente de trafico.
  · Top productos del dia por ventas.

Si una marca falla, se conserva su ultimo dato bueno (marcado ok=false) para
que la pantalla nunca quede en cero. Solo sale con error si fallan las 3.

Uso: python panel.py   (necesita mi_token.txt, igual que pipeline.py)
"""
import json
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


# Tipos de trafico de Yandex (id de la dimension TrafficSource) -> etiqueta.
FUENTES = {
    "ad": "Pagado (anuncios)", "organic": "Orgánico (buscadores)",
    "direct": "Directo", "social": "Redes sociales", "referral": "Otros sitios",
    "email": "Email", "messenger": "Mensajería", "internal": "Interno",
    "saved": "Guardados", "recommend": "Recomendaciones", "qrcode": "Código QR",
    "undefined": "Sin definir",
}


def traer_fuentes(contador):
    """Sesiones, transacciones y ventas de hoy por tipo de trafico."""
    p = {"ids": contador, "dimensions": "ym:s:<attribution>TrafficSource",
         "metrics": "ym:s:visits,ym:s:ecommercePurchases,ym:s:ecommerceRevenue",
         "date1": "today", "date2": "today", "attribution": ATTRIBUTION,
         "sort": "-ym:s:ecommerceRevenue", "limit": 20}
    r = requests.get(URL_DATA, headers=HEADERS, params=p, timeout=60)
    d = r.json()
    if r.status_code != 200 or "data" not in d:
        raise RuntimeError(f"Yandex HTTP {r.status_code}: "
                           f"{d.get('message') or d.get('errors') or d}")
    out = []
    for fila in d["data"]:
        dim = fila["dimensions"][0] or {}
        ses, compras, ingresos = fila["metrics"]
        nombre = FUENTES.get(dim.get("id"), dim.get("name") or "Sin definir")
        out.append({"fuente": nombre, "sesiones": int(ses or 0),
                    "compras": int(compras or 0),
                    "ingresos": round(float(ingresos or 0), 2)})
    # orden: ventas y, a igualdad, sesiones (de madrugada puede no haber ventas)
    out.sort(key=lambda x: (x["ingresos"], x["sesiones"]), reverse=True)
    return out


def construir_marca(contador, ahora):
    f = traer_funnel(contador)                       # hoy
    horas, sesiones = traer_horas(contador, ahora)
    hoy = ahora.date().isoformat()
    productos = traer_top_productos(contador)        # top 100 por unidades
    top = sorted(productos, key=lambda p: p["ingresos"], reverse=True)[:TOP_PANEL]
    try:
        fuentes = traer_fuentes(contador)
    except Exception as e:                       # no tumba la marca entera
        print(f"  Aviso fuentes ({contador}): {e}", file=sys.stderr)
        fuentes = None
    return {
        "ok": True,
        "fecha": hoy,
        "actualizado": ahora.isoformat(timespec="seconds"),
        "kpis": {
            "ingresos": round(f["ingresos"], 2),
            "sesiones": sesiones,
            "compras": f["compras"],
            "cr": round(f["cr"], 4),
            "ticket": round(f["ticket"], 2),
        },
        "funnel": {"sesiones": sesiones, "vistos": f["vistos"],
                   "carrito": f["carrito"], "comprados": f["comprados"]},
        "horas": horas,
        "fuentes": fuentes,
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
