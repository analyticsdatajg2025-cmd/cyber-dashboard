"""
diag_busquedas.py — DIAGNOSTICO de solo lectura: ¿Yandex tiene las palabras
buscadas en el buscador de cada web?

No escribe archivos del repo, no toca Sheets, no manda correos.
Para cada marca imprime:
  1) los parametros de URL mas vistos del dia (ahi deberia aparecer el de
     busqueda: q, _q, ft, text, query, search...);
  2) las URLs de paginas vistas que parecen de resultados de busqueda, con
     el termino que se extrae de cada una;
  3) un top 15 de terminos, como saldria en el correo.

Uso:  FECHA=today python diag_busquedas.py
En Actions: workflow "Diagnostico busquedas" (resultado en el log).
"""
import os
import requests

from pipeline import (MARCAS, HEADERS, URL_DATA, RE_URL_BUSQUEDA,
                      termino_de_url)

FECHA = os.environ.get("FECHA", "today").strip() or "today"

def consulta(params):
    d = requests.get(URL_DATA, headers=HEADERS, params=params, timeout=60).json()
    if "errors" in d:
        print(f"    ERROR Yandex: {d.get('message') or d['errors']}")
        return []
    return d.get("data", [])


def main():
    for marca, cfg in MARCAS.items():
        cont = cfg["contador"]
        print("=" * 72)
        print(f"{marca}  (contador {cont}, fecha {FECHA})")

        print("  1) Parametros de URL mas vistos:")
        filas = consulta({"ids": cont, "dimensions": "ym:pv:URLParamName",
                          "metrics": "ym:pv:pageviews", "date1": FECHA,
                          "date2": FECHA, "sort": "-ym:pv:pageviews",
                          "limit": 40})
        for f in filas:
            print(f"     {f['dimensions'][0].get('name')!s:30} "
                  f"{int(f['metrics'][0]):>8,} vistas")

        print("  2) URLs que parecen de busqueda (top 25):")
        filas = consulta({"ids": cont, "dimensions": "ym:pv:URL",
                          "metrics": "ym:pv:pageviews",
                          "filters": f"ym:pv:URL=~'{RE_URL_BUSQUEDA}'",
                          "date1": FECHA, "date2": FECHA,
                          "sort": "-ym:pv:pageviews", "limit": 5000})
        conteo = {}
        for i, f in enumerate(filas):
            url = f["dimensions"][0].get("name") or ""
            n = int(f["metrics"][0])
            t = termino_de_url(url)
            if i < 25:
                print(f"     {n:>6,}  {url[:110]}  ->  {t!r}")
            if t:
                conteo[t] = conteo.get(t, 0) + n
        print(f"     ({len(filas)} URLs de busqueda, {len(conteo)} terminos distintos)")

        print("  3) Top 15 terminos:")
        for t, n in sorted(conteo.items(), key=lambda kv: -kv[1])[:15]:
            print(f"     {n:>6,}  {t}")


if __name__ == "__main__":
    main()
