from playwright.sync_api import sync_playwright
from reporte_datos import armar_reporte, MARCAS

# Paleta por marca (ajustable luego al hex exacto de tu branding)
PALETA = {
    "LA CURACAO":  {"primary": "#F5B800", "texto_header": "#1a1a1a"},
    "TIENDAS EFE": {"primary": "#1C6DD0", "texto_header": "#ffffff"},
}

def fmt_k(v):
    return f"{v/1000:.1f}k" if v >= 10000 else f"{v:,.0f}"

CSS = """
* { margin:0; padding:0; box-sizing:border-box; font-family: -apple-system, 'Segoe UI', Roboto, sans-serif; }
body { width:640px; background:#fff; padding:24px; }
.logo { font-size:26px; font-weight:800; margin-bottom:18px; }
.pill { color:#fff; text-align:center; font-weight:700; font-size:18px;
        padding:10px; border-radius:30px; margin:18px 0 14px; }
.cards { display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px; }
.card { border:1px solid #eee; border-radius:12px; padding:12px; }
.card .lbl { font-size:11px; color:#777; margin-bottom:6px; }
.card .val { font-size:20px; font-weight:700; color:#222; }
table { width:100%; border-collapse:collapse; font-size:12px; margin-top:4px; }
th { background:#111; color:#fff; padding:6px 4px; font-size:11px; }
td { padding:4px 6px; text-align:center; border-bottom:1px solid #f2f2f2; }
.hora { font-weight:700; background:#fafafa; }
.pend { color:#bbb; }
.total td { font-weight:800; background:#111; color:#fff; }
"""

def color_alpha(base_rgb, frac):
    a = 0.12 + 0.68 * max(0, min(1, frac))
    return f"rgba({base_rgb},{a:.2f})"

def construir_html(r):
    pal = PALETA[r["marca"]]
    prim, txt = pal["primary"], pal["texto_header"]
    f = r["funnel"]

    cards = f"""
    <div class="cards">
      <div class="card"><div class="lbl">Ingresos Gross</div><div class="val">{fmt_k(f['ingresos'])} S/.</div></div>
      <div class="card"><div class="lbl">CR Gross</div><div class="val">{f['cr']*100:.2f} %</div></div>
      <div class="card"><div class="lbl">Ticket Promedio</div><div class="val">{f['ticket']:,.2f} S/.</div></div>
    </div>
    <div class="cards" style="margin-top:10px;">
      <div class="card"><div class="lbl">Artículos vistos</div><div class="val">{fmt_k(f['vistos'])}</div></div>
      <div class="card"><div class="lbl">Añadidos al carrito</div><div class="val">{f['carrito']:,}</div></div>
      <div class="card"><div class="lbl">Artículos comprados</div><div class="val">{f['comprados']:,}</div></div>
    </div>
    """

    proy, real, corte = r["proy"], r["real"], r["corte"]
    max_p = max(proy) or 1
    max_r = max(real[:corte+1] or [1]) or 1
    filas = ""
    for h in range(24):
        p = proy[h]
        if h <= corte:
            rd = real[h]
            var = rd - p
            pvar = (rd / p - 1) if p else 0
            bg_p = color_alpha("46,160,107", p / max_p)
            bg_r = color_alpha("245,178,40", rd / max_r)
            col_v = "#1f9d57" if var >= 0 else "#c0392b"
            bg_pv = "rgba(46,160,107,0.18)" if pvar >= 0 else "rgba(192,57,43,0.16)"
            filas += f"""<tr>
              <td class="hora">{h}</td>
              <td style="background:{bg_p}">{p:,.0f}</td>
              <td style="background:{bg_r}">{rd:,}</td>
              <td style="color:{col_v}">{var:,.0f}</td>
              <td style="background:{bg_pv}">{pvar:.1%}</td></tr>"""
        else:
            filas += f"""<tr>
              <td class="hora">{h}</td>
              <td style="background:{color_alpha('46,160,107', p/max_p)}">{p:,.0f}</td>
              <td class="pend">—</td><td class="pend">—</td><td class="pend">pend.</td></tr>"""

    total_p = sum(proy); total_r = r["real_acum"]
    tvar = total_r - total_p
    tpvar = (total_r/total_p - 1) if total_p else 0
    filas += f"""<tr class="total">
      <td>TOTAL</td><td>{total_p:,.0f}</td><td>{total_r:,}</td>
      <td>{tvar:,.0f}</td><td>{tpvar:.1%}</td></tr>"""

    return f"""<html><head><meta charset="utf-8"><style>{CSS}</style></head><body>
      <div class="logo" style="color:{prim}">{r['marca']}</div>
      <div class="pill" style="background:{prim};color:{txt}">Funnel de conversión</div>
      {cards}
      <div class="pill" style="background:{prim};color:{txt}">Tráfico</div>
      <table>
        <tr><th>HORA</th><th>Proyección</th><th>Real</th><th>VAR</th><th>%VAR</th></tr>
        {filas}
      </table>
    </body></html>"""

def texto_corte(r):
    return (f"SESIONES | Corte {r['corte']}:00\n"
            f"• Meta del día: {fmt_k(r['meta'])}\n"
            f"• Real: {fmt_k(r['real_acum'])} - Avance al {r['avance']*100:.1f}%")

def render_png(html, salida):
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(viewport={"width": 640, "height": 900}, device_scale_factor=2)
        pg.set_content(html, wait_until="load")
        pg.screenshot(path=salida, full_page=True)
        b.close()

if __name__ == "__main__":
    for marca in MARCAS:
        r = armar_reporte(marca)
        html = construir_html(r)
        archivo = f"reporte_{marca.replace(' ','_')}.png"
        render_png(html, archivo)
        print(f"✅ {archivo}")
        print(texto_corte(r), "\n")