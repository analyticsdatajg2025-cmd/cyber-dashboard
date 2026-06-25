"""
main.py — el ÚNICO archivo que corre el cron.

Flujo de una corrida:
  1. Extrae todo de Yandex (una sola vez por marca).
  2. Escribe data.json   -> lo consume el dashboard HTML.
  3. Escribe Google Sheet -> lo ve gerencia (corte + histórico semana).
  4. Genera PNG + correo  -> tu flujo actual de WhatsApp, intacto.

Si algo falla, manda correo de alerta y termina con código ≠ 0
(para que GitHub Actions marque el run como fallido).
"""
import json
import sys
import traceback
from build_data import construir_payload
from render import construir_html, render_png
from enviar_correo import construir_correo, enviar, GMAIL_USER

SALIDA_JSON = "data.json"


def alerta(msg):
    """Correo simple de alerta si el pipeline truena."""
    try:
        from email.message import EmailMessage
        import smtplib, ssl
        from enviar_correo import GMAIL_PASS, DESTINATARIOS
        m = EmailMessage()
        m["Subject"] = "FALLO pipeline Cyber - revisar"
        m["From"] = GMAIL_USER
        m["To"] = ", ".join(DESTINATARIOS)
        m.set_content("El pipeline de Cyber falló:\n\n" + msg)
        ctx = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ctx) as s:
            s.login(GMAIL_USER, GMAIL_PASS)
            s.send_message(m)
    except Exception:
        pass  # si hasta la alerta falla, igual salimos con error abajo


def main():
    # 1 + parte de 2/3: extracción única
    print("Extrayendo de Yandex...")
    payload, reportes = construir_payload()

    # 3b: inyectar histórico de semana leído del Sheet (si está disponible)
    try:
        from escribir_sheet import leer_semana, escribir_corte, escribir_semana
        escribir_corte(reportes)
        escribir_semana(reportes)
        payload["semana"] = leer_semana()
        print("Sheet actualizado.")
    except Exception as e:
        print(f"Aviso: Sheet no actualizado ({e}). El JSON sigue.")

    # 2: data.json
    with open(SALIDA_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"{SALIDA_JSON} escrito.")

    # 4: PNG + correo (flujo WhatsApp actual)
    reps_lista = []
    for marca, r in reportes.items():
        html = construir_html(r)
        render_png(html, f"reporte_{marca.replace(' ', '_')}.png")
        reps_lista.append(r)
    enviar(construir_correo(reps_lista))
    print("PNG + correo enviados.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        err = traceback.format_exc()
        print(err, file=sys.stderr)
        alerta(err)
        sys.exit(1)
