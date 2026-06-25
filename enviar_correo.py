import smtplib, ssl, mimetypes
from email.message import EmailMessage
from datetime import date
from reporte_datos import armar_reporte, MARCAS
from render import construir_html, render_png, texto_corte, fmt_k

# === Configuración ===
with open("gmail.txt", "r", encoding="utf-8") as f:
    lineas = [ln.strip() for ln in f if ln.strip()]
GMAIL_USER = lineas[0]
GMAIL_PASS = lineas[1].replace(" ", "")    # quita espacios por si los tiene

# A quién mandar el reporte (tú mismo, para que lo reenvíes por WhatsApp).
# Puedes poner varios separados por coma.
DESTINATARIOS = ["brandonveraf@gmail.com"]

# === Construcción del mensaje ===
def construir_correo(reportes):
    """Arma un solo correo con AMBAS marcas: 2 imágenes + 2 bloques de texto."""
    msg = EmailMessage()
    hora_corte = reportes[0]["corte"]
    msg["Subject"] = f"Corte sesiones {hora_corte}:00 - {date.today()}"
    msg["From"] = GMAIL_USER
    msg["To"] = ", ".join(DESTINATARIOS)

    # Cuerpo: texto plano con los bloques listos para copy/paste a WhatsApp
    cuerpo = "\n\n".join(texto_corte(r) + f"  ({r['marca']})" for r in reportes)
    cuerpo += "\n\n---\nLas imágenes están adjuntas. Reenvía cada una por WhatsApp."
    msg.set_content(cuerpo)

    # Adjuntar las dos PNG ya generadas
    for r in reportes:
        archivo = f"reporte_{r['marca'].replace(' ','_')}.png"
        with open(archivo, "rb") as f:
            data = f.read()
        msg.add_attachment(data, maintype="image", subtype="png", filename=archivo)
    return msg

def enviar(msg):
    contexto = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto) as smtp:
        smtp.login(GMAIL_USER, GMAIL_PASS)
        smtp.send_message(msg)

if __name__ == "__main__":
    print("📊 Generando reportes...")
    reportes = []
    for marca in MARCAS:
        r = armar_reporte(marca)
        html = construir_html(r)
        archivo = f"reporte_{marca.replace(' ','_')}.png"
        render_png(html, archivo)
        reportes.append(r)
        print(f"  ✅ {archivo}  (corte {r['corte']}:00, avance {r['avance']:.1%})")

    print("\n📧 Enviando correo...")
    msg = construir_correo(reportes)
    enviar(msg)
    print(f"  ✅ Enviado a: {', '.join(DESTINATARIOS)}")