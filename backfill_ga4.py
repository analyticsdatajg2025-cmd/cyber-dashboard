"""
backfill_ga4.py — CORRIDA ONE-SHOT. Re-trae desde GA4 los ultimos N dias
cerrados (por defecto 10) y los sobreescribe en historico.json.

Por que: al pasar de Yandex a GA4, los dias recientes del historico quedaron
con numeros de Yandex. actualizar_historico() no vuelve a consultar dias ya
guardados, asi que este script los reemplaza de una vez con GA4.

Usa exactamente las mismas funciones que pipeline.py (construir_bloque_fecha),
asi que el shape de cada dia es identico. Solo toca historico.json.
Despues conviene correr "Monitor Diario" a mano para que data.json
(Semana / Mes) se reconstruya con este historico.

Uso:  python backfill_ga4.py          (10 dias)
      DIAS=15 python backfill_ga4.py
"""
import json
import os
from datetime import date, datetime, timedelta

from pipeline import (MARCAS, MARCAS_ACTIVAS, SALIDA_HISTORICO,
                      construir_bloque_fecha, refrescar_metas_historico)


def main():
    dias = int(os.environ.get("DIAS", "10") or 10)
    ayer = date.today() - timedelta(days=1)
    fechas = [ayer - timedelta(days=i) for i in range(dias)]

    hist = {"marcas": {m: {} for m in MARCAS}}
    try:
        with open(SALIDA_HISTORICO, "r", encoding="utf-8") as f:
            prev = json.load(f)
        if isinstance(prev.get("marcas"), dict):
            for m in MARCAS:
                hist["marcas"][m] = prev["marcas"].get(m, {})
    except (FileNotFoundError, json.JSONDecodeError):
        pass

    print(f"Backfill GA4: {fechas[-1]} -> {fechas[0]} "
          f"· marcas: {', '.join(MARCAS_ACTIVAS)}")
    # Se arma todo primero y se escribe al final: si GA4 falla a mitad de
    # camino, historico.json queda como estaba.
    nuevos = {}
    for marca in MARCAS_ACTIVAS:
        for d in fechas:
            b = construir_bloque_fecha(marca, d)
            nuevos.setdefault(marca, {})[str(d)] = b
            print(f"  {marca} {d}: sesiones {b['real']['sesiones']}, "
                  f"ingresos {b['real']['ingresos']}")
    for marca, dd in nuevos.items():
        hist["marcas"].setdefault(marca, {}).update(dd)

    try:
        n = refrescar_metas_historico(hist)
        if n:
            print(f"  metas corregidas desde el Sheet: {n}")
    except Exception as e:
        print(f"  Aviso: no pude refrescar metas ({e}).")

    hist["generado"] = datetime.now().astimezone().isoformat(timespec="seconds")
    with open(SALIDA_HISTORICO, "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=2)
    print(f"{SALIDA_HISTORICO} escrito.")


if __name__ == "__main__":
    main()
