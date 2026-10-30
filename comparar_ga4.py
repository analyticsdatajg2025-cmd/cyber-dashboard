"""
comparar_ga4.py — VALIDACION del panel contra GA4. Solo lectura.

Consulta en Yandex un dia y una marca con las MISMAS funciones que usa
panel.py y lo pone al lado de los numeros de referencia de GA4 (tomados de
los dashboards de GA4), con la diferencia en %. No escribe ningun archivo
del repo ni toca Sheets.

Uso:  FECHA=2026-10-01 MARCA="LA CURACAO" python comparar_ga4.py
En Actions: workflow "Comparar con GA4" (el resultado sale en el resumen).

Que esperar: Yandex y GA4 son herramientas distintas (cuentan sesiones con
reglas propias, filtran bots distinto, atribuyen distinto). Diferencias de
hasta ~15% son normales; mas de ~30% indica que algo se mide distinto.
"""
import os
import sys

import requests

from pipeline import MARCAS, HEADERS, traer_funnel
from panel import (CANALES, traer_fuentes, traer_funnel_sesiones, _consulta)

FECHA = os.environ.get("FECHA", "2026-10-01").strip()
MARCA = os.environ.get("MARCA", "LA CURACAO").strip().upper()

# Referencia GA4 — LA CURACAO, 1 oct 2026 (dashboards "Resumen" y "FX canal").
# Canales GA4 agrupados como el panel:
#   Google  = Google Ads PMAX + Google SEM + Google Ads Shopping
#   Meta    = Facebook Ads + Facebook org + Instagram org
#   TikTok  = TikTok Ads + TikTok org
#   Organico= SEO / busqueda org ; Directo = Directo
#   Otros   = el resto (otros social, referral, APP, email)
GA4 = {
    ("2026-10-01", "LA CURACAO"): {
        "total": {"sesiones": 30345, "compras": 137, "ingresos": 145821.1},
        "funnel": {"sesiones": 30345, "vistos": 27962, "carrito": 1438,
                   "comprados": 137},
        "canales": {   # sesiones, transacciones, ingresos
            "Meta":     (1832, 9, 8404.9),
            "Google":   (10587, 69, 61095.9),
            "TikTok":   (7870, 0, 0.0),
            "Orgánico": (4947, 43, 50644.5),
            "Directo":  (3979, 4, 3268.8),
            "Otros":    (1130, 12, 22407.0),
        },
        # tabla "FX CANAL DETALLE" (sesiones, transacciones, ingresos)
        "detalle": {
            "GOOGLE ADS - PMAX": (8265, 60, 42072.9),
            "FACEBOOK ADS": (1536, 7, 7500.9),
            "GOOGLE SEM": (1460, 4, 6074.0),
            "GOOGLE ADS - SHOPPING": (1130, 5, 12949.0),
            "TIKTOK ADS": (8544, 0, 0.0),
            "GOOGLE DISPLAY": (104, 0, 0.0),
            "SEO / BUSQUEDA ORG": (5420, 43, 50644.5),
            "DIRECTO": (3582, 4, 3268.8),
            "OTROS SOCIAL ORG": (550, 3, 2906.0),
            "REFERRAL": (507, 3, 6147.0),
            "FACEBOOK ORG": (312, 2, 904.0),
            "TIKTOK ORG": (14, 0, 0.0),
            "INSTAGRAM ORG": (9, 0, 0.0),
            "APP LA CURACAO": (540, 6, 13354.0),
            "EMAIL": (29, 0, 0.0),
        },
    },
}

URL_GOALS = "https://api-metrika.yandex.net/management/v1/counter/{c}/goals"


def diagnostico_funnel(cont, sesiones):
    """Prueba formas alternativas de contar SESIONES por paso del funnel y
    lista las metas (goals) del contador. Solo lectura."""
    out = ["## Diagnóstico: cómo contar sesiones por paso", "",
           f"Sesiones totales del día: **{sesiones:,}**. Una opción sirve si da "
           "un número MENOR a las sesiones y distinto entre pasos.", "",
           "| Intento | Resultado |", "|---|---:|"]
    intentos = [
        ("filtro productImpressions>0", {"metrics": "ym:s:visits",
                                         "filters": "ym:s:productImpressions>0"}),
        ("filtro productImpressions>=1", {"metrics": "ym:s:visits",
                                          "filters": "ym:s:productImpressions>=1"}),
        ("filtro productBasketsQuantity>0", {"metrics": "ym:s:visits",
                                             "filters": "ym:s:productBasketsQuantity>0"}),
        ("filtro ecommercePurchases>0", {"metrics": "ym:s:visits",
                                         "filters": "ym:s:ecommercePurchases>0"}),
        ("filtro productID no nulo", {"metrics": "ym:s:visits",
                                      "filters": "ym:s:productID!n"}),
        ("filtro productName no nulo", {"metrics": "ym:s:visits",
                                        "filters": "ym:s:productName!n"}),
        ("metrica productImpressionsUniq", {"metrics": "ym:s:productImpressionsUniq"}),
        ("metrica productBasketsUniq", {"metrics": "ym:s:productBasketsUniq"}),
    ]
    for nombre, extra in intentos:
        try:
            d = _consulta({"ids": cont, "date1": FECHA, "date2": FECHA, **extra})
            out.append(f"| {nombre} | {int((d.get('totals') or [0])[0] or 0):,} |")
        except Exception as e:
            out.append(f"| {nombre} | error: {str(e)[:90]} |")
    out.append("")
    # metas del contador: 'visitas con meta' es una metrica por sesion
    try:
        r = requests.get(URL_GOALS.format(c=cont), headers=HEADERS, timeout=60)
        goals = r.json().get("goals", []) if r.status_code == 200 else []
        if r.status_code != 200:
            out.append(f"No pude listar metas: HTTP {r.status_code} {r.text[:150]}")
    except Exception as e:
        goals = []
        out.append(f"No pude listar metas: {e}")
    if goals:
        out += [f"Metas del contador ({len(goals)}):", "",
                "| ID | Nombre | Tipo | Sesiones con la meta |", "|---|---|---|---:|"]
        for g in goals[:30]:
            try:
                d = _consulta({"ids": cont, "date1": FECHA, "date2": FECHA,
                               "metrics": f"ym:s:goal{g['id']}visits"})
                v = f"{int((d.get('totals') or [0])[0] or 0):,}"
            except Exception as e:
                v = f"error: {str(e)[:60]}"
            out.append(f"| {g.get('id')} | {g.get('name', '')} | {g.get('type', '')} | {v} |")
    out.append("")
    return out


def dif(y, g):
    if not g:
        return "—"
    p = (y - g) / g * 100
    marca = "✅" if abs(p) <= 15 else ("⚠️" if abs(p) <= 30 else "❌")
    return f"{p:+.1f}% {marca}"


def fila(nombre, g, y, dinero=False):
    f = (lambda v: f"S/ {v:,.0f}") if dinero else (lambda v: f"{v:,.0f}")
    gs = f(g) if g is not None else "—"
    return f"| {nombre} | {gs} | {f(y)} | {dif(y, g) if g is not None else '—'} |"


def main():
    if MARCA not in MARCAS:
        sys.exit(f"Marca desconocida: {MARCA}. Opciones: {', '.join(MARCAS)}")
    cont = MARCAS[MARCA]["contador"]
    ref = GA4.get((FECHA, MARCA))
    out = [f"# Yandex vs GA4 — {MARCA} — {FECHA}", ""]
    if not ref:
        out += ["_No hay referencia GA4 cargada para esta fecha/marca: solo se "
                "muestran los valores de Yandex._", ""]

    # --- Totales -----------------------------------------------------------
    tot = _consulta({"ids": cont, "date1": FECHA, "date2": FECHA,
                     "metrics": "ym:s:visits"})
    sesiones = int((tot.get("totals") or [0])[0] or 0)
    f = traer_funnel(cont, FECHA)
    g = (ref or {}).get("total", {})
    out += ["## Totales del día", "", "| Métrica | GA4 | Yandex | Diferencia |",
            "|---|---:|---:|---:|",
            fila("Sesiones", g.get("sesiones"), sesiones),
            fila("Transacciones", g.get("compras"), f["compras"]),
            fila("Ventas gross", g.get("ingresos"), f["ingresos"], dinero=True)]
    cr_y = f["compras"] / sesiones * 100 if sesiones else 0
    cr_g = g["compras"] / g["sesiones"] * 100 if g.get("sesiones") else None
    out += [f"| CR | {f'{cr_g:.2f}%' if cr_g is not None else '—'} | {cr_y:.2f}% | |", ""]

    # --- Funnel ------------------------------------------------------------
    gf = (ref or {}).get("funnel", {})
    out += ["## Funnel por sesiones (lo que muestra el panel)", ""]
    try:
        fs = traer_funnel_sesiones(cont, sesiones, FECHA)
        out += ["| Paso | GA4 | Yandex | Diferencia |", "|---|---:|---:|---:|"]
        for k, n in [("sesiones", "Sesiones"), ("vistos", "Vieron producto"),
                     ("carrito", "Agregaron al carrito"), ("comprados", "Compraron")]:
            out.append(fila(n, gf.get(k), fs[k]))
        out.append("")
        # tasas de paso: comparan "forma" del funnel aunque el volumen difiera
        out += ["| Tasa de paso | GA4 | Yandex |", "|---|---:|---:|"]
        for a, b, n in [("sesiones", "vistos", "Sesión → ve producto"),
                        ("vistos", "carrito", "Ve producto → carrito"),
                        ("carrito", "comprados", "Carrito → compra")]:
            tg = f"{gf[b]/gf[a]*100:.1f}%" if gf.get(a) else "—"
            ty = f"{fs[b]/fs[a]*100:.1f}%" if fs[a] else "—"
            out.append(f"| {n} | {tg} | {ty} |")
    except Exception as e:
        out.append(f"❌ El funnel por sesiones NO funcionó: `{e}` → el panel "
                   "usaría el funnel por unidades.")
    out += ["", f"_Referencia: funnel por unidades (fichas vistas / unidades al "
                f"carrito / unidades compradas): {f['vistos']:,} / {f['carrito']:,} "
                f"/ {f['comprados']:,}_", ""]

    # --- Canales -----------------------------------------------------------
    gc = (ref or {}).get("canales", {})
    canales, detalle, crudo = traer_fuentes(cont, FECHA)
    out += ["## Canales", "",
            "| Canal | Ses. GA4 | Ses. Yandex | Dif. | Trx GA4 | Trx Yandex | "
            "Ventas GA4 | Ventas Yandex |",
            "|---|---:|---:|---:|---:|---:|---:|---:|"]
    for c in canales:
        gs, gt, gi = gc.get(c["fuente"], (None, None, None))
        n = lambda v, pre="": "—" if v is None else f"{pre}{v:,.0f}"
        out.append(f"| {c['fuente']} | {n(gs)} | {c['sesiones']:,} | "
                   f"{dif(c['sesiones'], gs) if gs else '—'} | {n(gt)} | "
                   f"{c['compras']} | {n(gi, 'S/ ')} | S/ {c['ingresos']:,.0f} |")
    out += ["", "_Ojo: GA4 incluye tráfico de la APP (en 'Otros'); el contador "
                "web de Yandex probablemente no lo ve._", ""]

    # --- Canal detalle (mismas reglas que el FX CANAL DETALLE de GA4) -------
    gd = (ref or {}).get("detalle", {})
    yd = {x["detalle"]: x for x in detalle}
    out += ["## Canal detalle (reglas FX CANAL DETALLE)", "",
            "| Canal detalle | Ses. GA4 | Ses. Yandex | Dif. | Trx GA4 | Trx Yandex | "
            "Ventas GA4 | Ventas Yandex |", "|---|---:|---:|---:|---:|---:|---:|---:|"]
    n = lambda v, pre="": "—" if v is None else f"{pre}{v:,.0f}"
    for k in list(gd) + [k for k in yd if k not in gd]:
        gs, gt, gi = gd.get(k, (None, None, None))
        y = yd.get(k, {"sesiones": 0, "compras": 0, "ingresos": 0})
        out.append(f"| {k} | {n(gs)} | {y['sesiones']:,} | "
                   f"{dif(y['sesiones'], gs) if gs else '—'} | {n(gt)} | {y['compras']} | "
                   f"{n(gi, 'S/ ')} | S/ {y['ingresos']:,.0f} |")
    out.append("")

    # --- Diagnostico del funnel ------------------------------------------
    out += diagnostico_funnel(cont, sesiones)

    # --- Detalle crudo para afinar la clasificacion -----------------------
    out += ["## Detalle crudo de Yandex (para revisar la clasificación)", "",
            "| Tipo | Motor / red | utm_source | utm_medium | utm_campaign | "
            "Sesiones | → Canal detalle |", "|---|---|---|---|---|---:|---|"]
    for r in crudo[:40]:
        out.append(f"| {r['tipo']} | {r['motor'] or '—'} | {r['utm'] or '—'} | "
                   f"{r['medio'] or '—'} | {(r['campana'] or '—')[:40]} | "
                   f"{r['sesiones']:,} | {r['detalle']} |")

    texto = "\n".join(out)
    print(texto)
    resumen = os.environ.get("GITHUB_STEP_SUMMARY")
    if resumen:
        with open(resumen, "a", encoding="utf-8") as fh:
            fh.write(texto + "\n")


if __name__ == "__main__":
    main()
