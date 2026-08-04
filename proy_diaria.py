"""
proy_diaria.py — Lectura del Sheet mensual de reporte diario.

De cada pestania mensual ("LC & EFE & SC - AGO26") lee, por marca
(LA CURACAO / TIENDAS EFE) y por dia:
    col D  = Fecha
    col F  = Sesiones Proy   -> meta del dia   (leer_proyeccion_diaria)
                             -> meta del mes   (leer_meta_mensual, suma del mes)
    col O  = % CR Real       -> cr_neto        \\
    col R  = Venta Real      -> venta_neta      >  (leer_neto_diario)
    col AL = Ticket Real     -> ticket_neto     /
    col J  = Trx Real        -> solo calculo interno, no se publica

Los valores se devuelven TAL CUAL estan en el Sheet: no se divide entre IGV
ni se recalcula nada. Lo que esta en la col R es lo que sale en el dashboard.

Estructura de la hoja: bloques apilados, cada uno con un banner en col C
    fila   2 -> "CONSOLIDADO CURACAO + EFE"   (se ignora)
    fila  38 -> "La Curacao"                  -> LA CURACAO
    fila  74 -> "Tiendas Efe"                 -> TIENDAS EFE
    fila 110 -> "SkullCandy"                  (se ignora)
Dentro de cada bloque hay una fila por dia con la fecha en col D. La fila de
encabezado del bloque ("Dia" / "Fecha") tambien trae los totales del mes, pero
como su col D no es una fecha valida, se descarta sola.

JUNTOZ no esta en estas hojas -> devuelve None / 0.0.
"""
from datetime import date, datetime, timedelta

# === EDITA: key del Google Sheet de reporte diario ===
#   URL: https://docs.google.com/spreadsheets/d/<KEY>/edit  -> pega solo <KEY>
PROY_DIARIA_SHEET_KEY = "15Nj4rQjh4uAmAsT42VP2bcXTwmfMqcpw7qYsugdUo1Q"

# Se lee con el MISMO service account que el resto (credenciales.json).
CREDENCIALES = "credenciales.json"

_MES = {1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR", 5: "MAY", 6: "JUN",
        7: "JUL", 8: "AGO", 9: "SET", 10: "OCT", 11: "NOV", 12: "DIC"}

# Overrides para meses cuyo tab NO sigue el patron "LC & EFE & SC - {MES}{YY}".
_TAB_OVERRIDE = {
    # (2026, 3): "LC & EFE & SKULL - MAR26",
}

# Banners que SI son un bloque de marca.
_BANNERS = {"la curacao": "LA CURACAO", "tiendas efe": "TIENDAS EFE"}

# Banners conocidos que cierran el bloque anterior pero NO son marca nuestra.
# Cualquier otro texto en col C se trata como ruido y NO corta el bloque:
# hay filas de dia sin fecha (ej. el ultimo dia del mes en algunas hojas) que
# antes se confundian con un banner y mataban el resto del bloque.
_BANNERS_IGNORAR = {"consolidado curacao + efe", "skullcandy", "skull candy",
                    "consolidado", "juntoz"}

_EPOCH = date(1899, 12, 30)   # serial de fecha de Sheets/Excel

# Indices de columna (0-based) dentro de la fila que devuelve gspread (col A = 0)
_COL = {"fecha": 3, "proy": 5, "trx": 9, "cr": 14, "venta": 17, "ticket": 37}

_ss = None
_cache = {}      # {tab: {(marca, 'YYYY-MM-DD'): {...}}}
_titulos = None  # lista de hojas, cacheada: se consulta ~60 veces por corrida
                 # (una por dia del mes y del historico) y cada llamada a
                 # worksheets() es un hit contra la API de Sheets.


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
    """Numero desde la celda. Con UNFORMATTED_VALUE gspread ya devuelve floats;
    el parseo de texto es solo un salvavidas para gspread viejo."""
    if isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or "").strip()
    if not s or s.startswith("#"):          # vacias o error de formula (#VALUE!)
        return 0.0
    # formato peruano: 'S/.93.104' -> 93104 ; '0,32%' -> 0.0032
    es_pct = s.endswith("%")
    s = (s.replace("S/.", "").replace("S/", "").replace("$", "")
          .replace(" ", "").replace("%", ""))
    s = s.replace(".", "").replace(",", ".")   # '.' miles fuera, ',' decimal
    try:
        n = float(s)
    except ValueError:
        return 0.0
    return n / 100.0 if es_pct else n


def _vacia(x):
    return x is None or (isinstance(x, str) and x.strip() == "")


def _cargar_tab(tab):
    ws = _abrir().worksheet(tab)
    try:
        vals = ws.get_values(value_render_option="UNFORMATTED_VALUE")
    except TypeError:                        # gspread viejo sin ese kwarg
        vals = ws.get_all_values()
    tabla, bloque = {}, None
    for row in vals:
        def cel(campo):
            i = _COL[campo]
            return row[i] if len(row) > i else None

        c = _norm(row[2]) if len(row) > 2 else ""
        # Banner = col C con texto conocido + col D (fecha) y col F (proy) vacias.
        # Solo los nombres catalogados cambian de bloque; el resto se ignora
        # para no cortar el bloque por una fila de dia mal formada.
        if c and _vacia(cel("fecha")) and _vacia(cel("proy")):
            if c in _BANNERS:
                bloque = _BANNERS[c]
                continue
            if c in _BANNERS_IGNORAR:
                bloque = None
                continue
            continue          # texto desconocido -> ruido, no corta el bloque
        if bloque is None:
            continue

        f = _a_fecha(cel("fecha"))
        if f is None:
            continue
        tabla[(bloque, f.isoformat())] = {
            "proy":        _num(cel("proy")),
            "venta_neta":  round(_num(cel("venta")), 2),
            "cr_neto":     round(_num(cel("cr")), 6),
            "ticket_neto": round(_num(cel("ticket")), 2),
            "trx_netas":   int(_num(cel("trx"))),
        }
    return tabla


def _titulos_hojas():
    global _titulos
    if _titulos is None:
        _titulos = [w.title for w in _abrir().worksheets()]
    return _titulos


def _tabla(fecha):
    """Tabla completa del mes de `fecha`, o None si no hay hoja."""
    if PROY_DIARIA_SHEET_KEY == "PEGA_LA_KEY_AQUI":
        return None
    tab = _tab_para(fecha, _titulos_hojas())
    if tab is None:
        return None
    if tab not in _cache:
        _cache[tab] = _cargar_tab(tab)
    return _cache[tab]


def _entrada(marca, fecha):
    fecha = fecha or date.today()
    if marca not in _BANNERS.values():
        return None
    t = _tabla(fecha)
    return t.get((marca, fecha.isoformat())) if t else None


def leer_proyeccion_diaria(marca, fecha=None):
    """Meta de sesiones del dia (col F). 0.0 si la marca no aplica."""
    e = _entrada(marca, fecha)
    return e["proy"] if e else 0.0


def leer_meta_mensual(marca, fecha=None):
    """Meta de sesiones del MES = suma de la col F de todos los dias del mes
    al que pertenece `fecha` (incluye los dias que aun no ocurren)."""
    fecha = fecha or date.today()
    if marca not in _BANNERS.values():
        return 0.0
    t = _tabla(fecha)
    if not t:
        return 0.0
    pref = fecha.strftime("%Y-%m")
    return sum(v["proy"] for (m, f), v in t.items()
               if m == marca and f.startswith(pref))


def leer_neto_diario(marca, fecha=None):
    """Neto del dia desde el Sheet: {venta_neta, cr_neto, ticket_neto, trx_netas}.
    None si la marca no aplica, falta la key, o no hay venta ese dia."""
    e = _entrada(marca, fecha)
    if not e or e["venta_neta"] <= 0:
        return None
    return {
        "venta_neta":  e["venta_neta"],
        "cr_neto":     e["cr_neto"],
        "ticket_neto": e["ticket_neto"],
        "trx_netas":   e["trx_netas"],
    }
