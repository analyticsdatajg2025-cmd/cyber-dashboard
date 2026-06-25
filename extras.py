"""
Extracciones nuevas validadas contra los contadores reales (status 200).
Se suman a tu reporte_datos.py existente (funnel + sesiones/hora).

Nombres confirmados en el probe del 2026:
  - Campañas:      ym:s:UTMCampaign
  - Top productos: ym:s:productName / ym:s:productID + productPurchasedPrice (= ingreso)
  - Ingresos/hora: bytime + ym:s:ecommerceRevenue
"""
import requests

with open("mi_token.txt", "r", encoding="utf-8") as f:
    TOKEN = f.read().strip()

HEADERS = {"Authorization": f"OAuth {TOKEN}"}
URL_DATA   = "https://api-metrika.yandex.net/stat/v1/data"
URL_BYTIME = "https://api-metrika.yandex.net/stat/v1/data/bytime"

# Modelo de atribución que ve gerencia por defecto en Yandex
ATTRIBUTION = "lastsign"


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
        out.append({
            "nombre": nombre,
            "sesiones": int(ses),
            "compras": int(compras),
            "ingresos": round(ingresos, 2),
        })
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
        out.append({
            "nombre": nombre,
            "sku": sku,
            "vistos": int(vistos),
            "carrito": int(carrito),
            "comprados": int(comprados),
            "sesiones": int(ses),
            "ingresos": round(ingresos, 2),
        })
    return out


def traer_ingresos_hora(contador):
    """bytime devuelve las métricas como lista anidada: data[0]['metrics'][0] = array de 24h."""
    p = {
        "ids": contador,
        "metrics": "ym:s:ecommerceRevenue",
        "date1": "today", "date2": "today",
        "group": "hour",
    }
    d = requests.get(URL_BYTIME, headers=HEADERS, params=p).json()
    if not d.get("data"):
        return [0.0] * 24
    serie = d["data"][0]["metrics"][0]   # <-- el array por hora
    return [round(float(v or 0), 2) for v in serie]


if __name__ == "__main__":
    CONT = "98373248"  # LA CURACAO
    print("Campañas:", traer_campanas(CONT)[:3])
    print("Productos:", traer_top_productos(CONT)[:3])
    print("Ingresos/hora:", traer_ingresos_hora(CONT))
