"""
Escribe al Google Sheet de gerencia (gspread) y lee el histórico de semana.

Pestañas esperadas en el Sheet:
  - "Cortes":  marca | fecha | hora | sesiones | ingresos | cr | ticket | avance
  - "Semana":  marca | fecha | dia | proy | real | avance

Credenciales: service account JSON. En GitHub Actions va como secret
GOOGLE_CREDS y se escribe a credenciales.json antes de correr.
"""
import gspread
from datetime import date

NOMBRE_SHEET = "CYBER 2026 - BACKEND"   # <-- ajusta al nombre real
DIAS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]


def _abrir():
    gc = gspread.service_account(filename="credenciales.json")
    return gc.open(NOMBRE_SHEET)


def escribir_corte(reportes):
    """Escribe/sobrescribe la fila del corte por marca (idempotente por marca+fecha+hora)."""
    ss = _abrir()
    sh = ss.worksheet("Cortes")
    filas = sh.get_all_values()
    hoy = str(date.today())

    for marca, r in reportes.items():
        hora = f"{r['corte']:02d}:00"
        f = r["funnel"]
        nueva = [marca, hoy, hora, r["real_acum"], round(f["ingresos"], 2),
                 round(f["cr"], 4), round(f["ticket"], 2), round(r["avance"], 4)]

        # Idempotencia: buscar fila con misma marca+fecha+hora
        idx = None
        for i, row in enumerate(filas[1:], start=2):
            if len(row) >= 3 and row[0] == marca and row[1] == hoy and row[2] == hora:
                idx = i
                break
        if idx:
            sh.update(f"A{idx}:H{idx}", [nueva])
        else:
            sh.append_row(nueva, value_input_option="USER_ENTERED")


def escribir_semana(reportes):
    """Actualiza el total del día por marca (idempotente por marca+fecha)."""
    ss = _abrir()
    sh = ss.worksheet("Semana")
    filas = sh.get_all_values()
    hoy = str(date.today())
    dia = DIAS[date.today().weekday()]

    for marca, r in reportes.items():
        proy_corte = round(sum(r["proy"][:r["corte"] + 1]))
        nueva = [marca, hoy, dia, round(r["meta"]), r["real_acum"],
                 round(r["real_acum"] / proy_corte, 4) if proy_corte else 0]
        idx = None
        for i, row in enumerate(filas[1:], start=2):
            if len(row) >= 2 and row[0] == marca and row[1] == hoy:
                idx = i
                break
        if idx:
            sh.update(f"A{idx}:F{idx}", [nueva])
        else:
            sh.append_row(nueva, value_input_option="USER_ENTERED")


def leer_semana():
    """Lee la pestaña Semana para inyectarla en el data.json (vista Semana del dashboard)."""
    ss = _abrir()
    sh = ss.worksheet("Semana")
    filas = sh.get_all_values()[1:]   # salta header
    out = {}
    for row in filas:
        if len(row) < 6:
            continue
        marca, fecha, dia, proy, real, avance = row[:6]
        out.setdefault(marca, []).append({
            "fecha": fecha, "dia": dia,
            "proy": int(float(proy or 0)),
            "real": int(float(real or 0)),
            "avance": float(avance or 0),
        })
    return out
