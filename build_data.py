"""
Ensambla el diccionario único que alimenta el dashboard (data.json)
y la pestaña de gerencia. Misma data, dos destinos.

Combina:
  - armar_reporte(marca)  -> tu reporte_datos.py (funnel, sesiones/hora, proy, meta, corte)
  - extras.py             -> campañas, top productos, ingresos/hora
"""
from datetime import datetime
from reporte_datos import armar_reporte, MARCAS
from extras import (traer_campanas, traer_top_productos,
                    traer_ingresos_hora, ATTRIBUTION)


def construir_marca(marca):
    """Devuelve (reporte_crudo, bloque_para_json). El crudo se reusa para el PNG."""
    r = armar_reporte(marca)               # ya pega a Yandex 2 veces
    contador = MARCAS[marca]["contador"]

    campanas  = traer_campanas(contador)
    productos = traer_top_productos(contador)
    ing_hora  = traer_ingresos_hora(contador)

    corte = r["corte"]
    por_hora = []
    for h in range(24):
        por_hora.append({
            "h": h,
            "ses_proy": round(r["proy"][h]),
            "ses_real": r["real"][h] if h <= corte else None,
            "ing_real": ing_hora[h]   if h <= corte else None,
        })

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
        "semana": {},   # se llena leyendo el Sheet (escribir_sheet.leer_semana)
    }
    reportes = {}
    for marca in MARCAS:
        r, bloque = construir_marca(marca)
        payload["marcas"][marca] = bloque
        reportes[marca] = r
        payload["corte"] = r["corte"]   # mismo corte para ambas marcas
    return payload, reportes
