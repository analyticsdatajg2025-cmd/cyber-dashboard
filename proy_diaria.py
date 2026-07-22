"""
proy_diaria.py — Proyección diaria de sesiones + NETO diario (modo 'diario').

De cada pestaña mensual del Sheet de reporte diario ("LC & EFE & SC - JUL26")
lee, por marca (LA CURACAO / TIENDAS EFE) y por día:
    col F  = Sesiones Proy   -> meta del día  (leer_proyeccion_diaria)
    col O  = % CR Real        -> cr_neto        \\
    col R  = Venta Real       -> venta_neta      >  (leer_neto_diario)
    col AL = Ticket Prom      -> ticket_neto     /
    col J  = Trx Real         -> trx_netas      /

JUNTOZ no está en estas hojas -> devuelve None/0.0 (monitor en vivo puro).

Estructura: bloques apilados con banner en col C ("La Curacao", "Tiendas Efe";
el "CONSOLIDADO CURACAO + EFE" se ignora). Dentro de cada bloque, una fila por
día con la fecha en col D.

El modo se controla con la variable de entorno MODO ('diario' | 'cyber').
"""
import os
from datetime import date, datetime, timedelta

MODO = os.environ.get("MODO", "cyber").strip().lower()
ES_DIARIO = (MODO == "diario")

# === EDITA: key del Google Sheet de reporte diario ===
#   URL: https://docs.google.com/spreadsheets/d/<KEY>/edit  -> pega solo <KEY>
PROY_DIARIA_SHEET_KEY = "15Nj4rQjh4uAmAsT42VP2bcXTwmfMqcpw7qYsugdUo1Q"

# Se lee con el MISMO service account que el resto (credenciales.json).
CREDENCIALES = "credenciales.json"

_MES = {1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR", 5: "MAY", 6: "JUN",
        7: "JUL", 8: "AGO", 9: "SET", 10: "OCT", 11: "NOV", 12: "DIC"}

# Overrides para meses cuyo tab NO sigue el patrón "LC & EFE & SC - {MES}{YY}".
_TAB_OVERRIDE = {
    # (2026, 3): "LC & EFE & SKULL - MAR26",
}

_BANNERS = {"la curacao": "LA CURACAO", "tiendas efe": "TIENDAS EFE"}
_EPOCH = date(1899, 12, 30)   # serial de fecha de Sheets/Excel

# Índices de columna (0-based) dentro de la fila que devuelve gspread (col A = 0)
_COL = {"fecha": 3, "proy": 5, "trx": 9, "cr": 14, "venta": 17, "ticket": 37}

_ss = None
_cache = {}   # {tab: {(marca, 'YYYY-MM-DD'): {proy, venta_neta, cr_neto, ticket_neto, trx_netas}}}


def _norm(s):
    return str(s or "").strip().lower()


def _abrir():
    global _ss
    if _ss is None:
        import gspread
        gc = gspread.service_account(filename=CREDENCIALES)
        _ss = gc.open_by_key(PROY_DIARIA_SHEET_KEY)
    return _ss


def _tab_para(fecha, titulos):
    if (fecha.year, fecha.month) in _TAB_OVERRIDE:
        return _TAB_OVERRIDE[(fecha.year, fecha.month)]
    suf = f"{_MES[fecha.month]}{fecha.strftime('%y')}"
    exacto = f"LC & EFE & SC - {suf}"
    if exacto in titulos:
        return exacto
    for t in titulos:                       # fallback: variantes SKULL / SC / plano
        if t.startswith("LC & EFE") and t.replace(" ", "").endswith(suf):
            return t
    return None


def _a_fecha(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):         # serial (UNFORMATTED_VALUE)
        try:
            return _EPOCH + timedelta(days=int(v))
        except (ValueError, OverflowError):
            return None
    if isinstance(v, str):                  # fallback: texto
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%Y"):
            try:
                return datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


def _num(v):
    if isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or "").strip().replace(" ", "")
    if not s or s.startswith("#"):          # celdas vacías o error de fórmula (#VALUE!)
        return 0.0
    s = s.replace(".", "").replace(",", ".")   # miles con punto (locale PE)
    try:
        return float(s)
    except ValueError:
        return 0.0


def _cargar_tab(tab):
    ws = _abrir().worksheet(tab)
    try:
        vals = ws.get_values(value_render_option="UNFORMATTED_VALUE")
    except TypeError:                        # gspread viejo sin ese kwarg
        vals = ws.get_all_values()
    tabla, bloque = {}, None
    for row in vals:
        c = _norm(row[2]) if len(row) > 2 else ""
        if c in _BANNERS:
            bloque = _BANNERS[c]
            continue
        if "consolidado" in c:
            bloque = None
            continue
        if bloque is None:
            continue

        def cel(campo):
            i = _COL[campo]
            return row[i] if len(row) > i else None

        f = _a_fecha(cel("fecha"))
        if f is None:
            continue
        tabla[(bloque, f.isoformat())] = {
            "proy":        _num(cel("proy")),
            "venta_neta":  round(_num(cel("venta")), 2),
            "cr_neto":     round(_num(cel("cr")), 4),
            "ticket_neto": round(_num(cel("ticket")), 2),
            "trx_netas":   int(_num(cel("trx"))),
        }
    return tabla


def _entrada(marca, fecha):
    fecha = fecha or date.today()
    if marca not in _BANNERS.values():
        return None
    if PROY_DIARIA_SHEET_KEY == "PEGA_LA_KEY_AQUI":
        return None
    tab = _tab_para(fecha, [w.title for w in _abrir().worksheets()])
    if tab is None:
        return None
    if tab not in _cache:
        _cache[tab] = _cargar_tab(tab)
    return _cache[tab].get((marca, fecha.isoformat()))


def leer_proyeccion_diaria(marca, fecha=None):
    """Meta de sesiones del día (col F) para LA CURACAO / TIENDAS EFE. 0.0 si no aplica."""
    e = _entrada(marca, fecha)
    return e["proy"] if e else 0.0


def leer_neto_diario(marca, fecha=None):
    """Neto del día desde el Sheet: {venta_neta, cr_neto, ticket_neto, trx_netas}.
    Devuelve None si la marca no aplica, falta la key, o no hay venta ese día."""
    e = _entrada(marca, fecha)
    if not e or e["venta_neta"] <= 0:
        return None
    return {
        "venta_neta":  e["venta_neta"],
        "cr_neto":     e["cr_neto"],
        "ticket_neto": e["ticket_neto"],
        "trx_netas":   e["trx_netas"],
    }
