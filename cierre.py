"""
cierre.py — CORREO DE CIERRE de un dia YA TERMINADO (24 h completas).

El correo horario (pipeline.py) nunca manda el dia cerrado: el ultimo corte
sale a las 23:xx con datos hasta las 22:59. Este script arma el mismo correo
(2 imagenes por marca: trafico + top productos) pero con el dia completo
00:00-23:59 de la fecha indicada.

Uso:
    python cierre.py               -> cierre de AYER (hora Lima)
    python cierre.py 2026-11-03    -> cierre de esa fecha
    FECHA=2026-11-03 python cierre.py

Lo corre el workflow cierre.yml (manual o por cron-job.org a primera hora).
No escribe data.json / historico.json / Sheets: solo PNG + correo.
"""
import os
import sys
from datetime import date, timedelta

from pipeline import (MARCAS, DIAS, dia_pertenece_a_evento, meta_diaria,
                      leer_proyecciones, USAR_PROY_HORARIA,
                      traer_sesiones_hora_fecha, traer_funnel,
                      traer_top_productos, traer_busquedas, alerta)
from proy_diaria import leer_neto_diario


def fecha_objetivo():
    txt = (sys.argv[1] if len(sys.argv) > 1 else os.environ.get("FECHA", "")).strip()
    if txt:
        return date.fromisoformat(txt)
    return date.today() - timedelta(days=1)


def proy_del_dia(marca, fecha):
    """Curva horaria del Excel si la fecha es de cyber (sin depender de si
    HOY es cyber); ceros si no hay curva."""
    if USAR_PROY_HORARIA and dia_pertenece_a_evento(fecha):
        proy = leer_proyecciones(marca, fecha)
        if sum(proy) > 0:
            return proy
    return [0.0] * 24


def top_productos_fecha(contador, fstr, top=100):
    """Top por unidades UNIDO al top por ingresos (igual que el corte en vivo)."""
    por_und = traer_top_productos(contador, top, fecha=fstr)
    try:
        por_ing = traer_top_productos(contador, top, fecha=fstr,
                                      orden="ym:s:productPurchasedPrice")
    except Exception as e:
        print(f"Aviso: top por ingresos no traido ({e}).")
        por_ing = []
    vistos, out = set(), []
    for p in por_und + por_ing:
        clave = (p["sku"], p["nombre"])
        if clave not in vistos:
            vistos.add(clave)
            out.append(p)
    return out


def neto_cierre(marca, fecha):
    """Neto del dia cerrado leido DIRECTO del Sheet al momento de correr (asi
    no depende de que pipeline.py ya haya refrescado neto.json). None si el
    Sheet aun no lo tiene o falla: correo.py prueba neto.json y si tampoco
    hay, muestra gross."""
    try:
        n = leer_neto_diario(marca, fecha)
    except Exception as e:
        print(f"Aviso: neto de {marca} no leido del Sheet ({e}).")
        return None
    if not n:
        return None
    return {k: n[k] for k in ("venta_neta", "cr_neto", "ticket_neto")}


def armar_reporte_cierre(marca, fecha):
    cont = MARCAS[marca]["contador"]
    fstr = str(fecha)
    proy = proy_del_dia(marca, fecha)
    real = traer_sesiones_hora_fecha(cont, fstr)          # 24 h
    real_acum = sum(real)
    meta = sum(proy) if any(proy) else meta_diaria(marca, fecha)
    return {
        "marca": marca, "dia": DIAS[fecha.weekday()], "fecha": fstr,
        "cierre": True,
        "corte": 23, "hora_corte": 24, "proy": proy, "real": real,
        "meta": meta, "real_acum": real_acum,
        "avance": (real_acum / meta) if meta else 0,
        "funnel": traer_funnel(cont, fstr),
        "top_productos": top_productos_fecha(cont, fstr),
        "busquedas": traer_busquedas(cont, fecha=fstr) or [],
        "neto": neto_cierre(marca, fecha),
    }


def main():
    fecha = fecha_objetivo()
    # JUNTOZ solo entra si esa fecha fue de cyber (mismo criterio que el diario)
    marcas = [m for m in MARCAS if dia_pertenece_a_evento(fecha) or m != "JUNTOZ"]
    print(f"Cierre del {fecha} · marcas: {', '.join(marcas)}")

    from correo import (construir_html, construir_html_top, archivos_png,
                        render_png, construir_correo, enviar)
    reportes = []
    for marca in marcas:
        r = armar_reporte_cierre(marca, fecha)
        png_trafico, png_top = archivos_png(marca)
        render_png(construir_html(r), png_trafico)
        try:
            render_png(construir_html_top(r), png_top)
        except Exception as e:
            print(f"Aviso: imagen top de {marca} no generada ({e}).")
        reportes.append(r)
        print(f"  OK {marca}: {r['real_acum']:,} sesiones, avance {r['avance']:.1%}, "
              f"neto: {'Sheet' if r['neto'] else 'neto.json o gross'}")
    enviar(construir_correo(reportes))
    print("Correo de cierre enviado.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        err = traceback.format_exc()
        print(err, file=sys.stderr)
        alerta(err)
        sys.exit(1)
