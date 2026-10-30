# CLAUDE.md — cyber-dashboard

Contexto para trabajar en este repo. Leer completo antes de cambiar algo.

## Qué es

Pipeline de datos del **dashboard de ventas web** de Grupo EFE / Conecta Retail
para tres marcas: **LA CURACAO**, **TIENDAS EFE** y **JUNTOZ** (marketplace).
Este repo NO tiene el frontend principal: genera JSONs que leen dos páginas
HTML alojadas en el cPanel de Banahosting (`cybertest.hablamas.pe`).

- Fuente de datos web: **Yandex Metrica** (Reporting API + Management API).
- Metas y venta neta: **Google Sheets** (cuenta de servicio).
- Corre en **GitHub Actions**; los resultados se commitean al repo y el front
  los lee de `raw.githubusercontent.com/.../main/*.json`.
- Idioma del código, comentarios y commits: **español**.

## Mapa de archivos

| Archivo | Rol |
|---|---|
| `pipeline.py` | "El cerebro". Corrida diaria / horaria-cyber. Extrae de Yandex, arma `data.json`, `historico.json`, `neto.json`; en cyber escribe Sheets, genera PNG y manda correo. |
| `proy_diaria.py` | Lee el Sheet mensual de reporte diario (metas de sesiones col F, venta neta col R, CR col O, ticket col AL). |
| `correo.py` | PNG por marca (Playwright) + correo por Gmail SMTP. También expone `GMAIL_USER/PASS/DESTINATARIOS` para la alerta de fallo. |
| `panel.py` | Corrida **liviana** del "Panel en vivo" (pantalla TV). Escribe SOLO `panel.json`. |
| `panel.html` | Página del Panel en vivo (se sube al cPanel junto a `dashboard_paid.html`). Lee `panel.json` cada 60 s. |
| `comparar_ga4.py` | Validación de solo lectura: Yandex vs números de GA4 cargados a mano (por fecha/marca). |
| `verificar_credenciales.py` | Chequeo de solo lectura de todos los secretos/accesos (Yandex, Google, Excel, Gmail, Playwright). |
| `PROYECCIONES_CYBER_JULIO_2026.xlsx` | Proyección horaria de sesiones para el modo cyber (pestañas `CYBER DAYS` / `CYBER WOW`). |
| `data.json` / `historico.json` / `neto.json` / `panel.json` | Salidas (las escribe el bot, no editar a mano). |

`dashboard_paid.html` (front principal) **no está en el repo**: vive en el cPanel.
Lee `data.json`, `historico.json`, `neto.json`. Tiene una pestaña
`<a class="vtab" href="panel.html" target="_blank">Panel en vivo ↗</a>`.

## Flujos

### 1. Dashboard principal (`pipeline.py`)

Dos modos, decididos **solo por la fecha** (`CYBER_DAYS` / `CYBER_WOW` en `pipeline.py`):

| | DIARIO | CYBER |
|---|---|---|
| Workflow | `Diario.yml` — cron `51 19 * * *` UTC (14:51 Lima) | `cyber.yml` — cada hora 09–21 Lima |
| Meta | Sheet mensual col F | Excel de proyección horaria |
| JUNTOZ | no se consulta | sí |
| Sheets BACKEND / META HORARIA | no | sí |
| PNG + correo | no | sí, cada hora |

`MODO` en el workflow dice qué workflow corre; si no le toca el día, se salta
solo (`job_le_toca()`). `MODO=auto` corre siempre (para pruebas locales).

Funciones Yandex en `pipeline.py`: `traer_sesiones_hora`, `traer_sesiones_hora_fecha`,
`traer_funnel`, `traer_campanas` (+ `evaluar_campanas_numericas` con el mapa
`CAMPANAS_META` de IDs de Meta), `traer_top_productos`, `traer_ingresos_hora`.
Atribución: `ATTRIBUTION = "lastsign"`.

`historico.json`: días cerrados desde `HIST_INICIO`, retención 60 días,
refresca los últimos 4 en diario, rellena huecos de a 7 por corrida.
`neto.json`: `pipeline.py` escribe solo `por_dia`; el bloque `hoy` lo escribe
`neto.py` desde la PC del usuario (no está en el repo) durante el cyber. Nunca
pisar `hoy`.

### 2. Panel en vivo (`panel.py` → `panel.json` → `panel.html`)

Pantalla TV para la noche de lanzamiento de cada cyber (21:00 → 03:00 Lima).

```
cron-job.org (cada 5 min, POST workflow_dispatch con token fine-grained)
  └─▶ GitHub Actions "Panel en vivo" (panel.yml) → python panel.py
        └─▶ commit "panel HH:MM Lima" con panel.json en main
              └─▶ panel.html (cPanel) lee raw.githubusercontent .../panel.json cada 60 s
```

- `panel.yml`: `workflow_dispatch` (disparo principal desde cron-job.org) +
  cron de respaldo `*/15 2-7 2,16 11 *` (21:00–02:59 Lima, noches 01→02 nov y
  15→16 nov). `concurrency: panel`. Respaldo cada 15 (no 5) para no duplicar
  consultas a Yandex.
- Por marca, `panel.json` trae: `kpis` (ventas gross, sesiones, compras, CR,
  ticket) del **día en curso** (se reinician a las 00:00), `funnel`, `fuentes`
  (6 canales), `fuentes_crudo` (filas sin clasificar, para depurar), `top`
  (productos por ventas), `horas` (no se grafican).
- Si una marca falla, se conserva su último dato con `ok:false`; solo falla
  el job si caen las 3.
- `panel.html`: 3 columnas, sin scroll; embudo SVG; tabla de canales sin
  "Otros" (% sesiones sobre el total de la marca); top productos ordenado por
  ventas, oculta filas que no entran; semáforo de frescura verde ≤12 min,
  ámbar ≤25, rojo más. No requiere refrescar la página.

**Funnel por sesiones (validado contra GA4):** Yandex **ignora los filtros por
métrica** (`ym:s:productImpressions>0` devuelve todas las sesiones). Se usa:
- Vieron producto → `ym:s:visits` con filtro de dimensión `ym:s:productID!n`.
- Carrito / compra → metas automáticas del contador de tipo `e_cart` /
  `e_purchase` (métrica `ym:s:goal<ID>visits`), buscadas por tipo vía
  Management API (`/management/v1/counter/{id}/goals`).
- Si no hay metas o el funnel no es monótono → respaldo en **unidades de
  producto** (`productImpressions`, `productBasketsQuantity`,
  `productPurchasedQuantity`). Ojo: unidades ≠ sesiones (una sesión ve varias fichas).

**Canales:** `canal_detalle()` en `panel.py` es la traducción 1:1 del CASE
**"FX CANAL DETALLE"** de GA4 (fuente / medio / campaña). Sin UTM, fuente y
medio se arman desde `TrafficSource` + `SourceEngine` de Yandex
(`fuente_medio()`); sin `utm_medium` se deriva del tipo de tráfico (`ad`→`cpc`).
`canal_panel()` agrupa en Meta / Google / TikTok / Orgánico / Directo / Otros.

### 3. Validación

- **Comparar con GA4** (`comparar_ga4.yml`, manual, inputs fecha + marca):
  totales, funnel + tasas de paso, canales, canal detalle, diagnóstico de
  formas de contar sesiones y detalle crudo. Resultado en el resumen del job.
  Referencias GA4 cargadas: LA CURACAO 2026-10-01.
- **Verificar credenciales** (`verificar.yml`, manual): sale con código 1 si
  algo falla.

Resultados conocidos (Yandex vs GA4): totales −5% a −13%; campañas pagadas
±15%; Yandex pone en **Directo** ~1,000 sesiones/día que GA4 pone en **SEO**
(la suma de ambos cuadra); funnel −11% a −29% con la misma forma; Yandex no
ve la APP; sesiones `internal` van a Otros.

## Secretos y configuración

| Secreto (GitHub Actions) | Uso | Se escribe a |
|---|---|---|
| `YANDEX_TOKEN` | OAuth Yandex Metrica | `mi_token.txt` |
| `GOOGLE_CREDS` | JSON de cuenta de servicio (`dash-bot@test-lacuracao.iam.gserviceaccount.com`) | `credenciales.json` |
| `GMAIL_USER`, `GMAIL_PASS` | Gmail + contraseña de aplicación | `gmail.txt` |

Los workflows escriben estos archivos al inicio y los borran al final.
**Nunca** usar `git add -A` / `git add .` en un workflow (subiría secretos):
siempre `git add <archivo>` explícito. `pipeline.py` lee `mi_token.txt` al importarse.

Contadores Yandex (`MARCAS` en `pipeline.py`): LA CURACAO `98373248`,
TIENDAS EFE `98373144`, JUNTOZ `98373308`.

Otros IDs: Sheet mensual (`PROY_DIARIA_SHEET_KEY` en `proy_diaria.py`),
Sheet "Meta horaria Cyber Julio26" (`META_SHEET_KEY`), Sheet
`"CYBER 2026 - BACKEND"` (`NOMBRE_SHEET`, abierto por nombre).

## Estado y pendientes (al 30 oct 2026)

- Fuente activa: **Yandex**. Hubo un incidente de seguridad con Yandex: se
  recomendó rotar `YANDEX_TOKEN`. La migración a **GA4** está en el PR #2
  (draft, sin mergear) — no mergear sin decidirlo con el usuario.
- **Modo cyber del pipeline NO está configurado** para noviembre: las fechas
  siguen en julio. Próximos eventos: **Cyber WOW 02–05 nov**, **Cyber Days
  16–18 nov**. Para activarlo hace falta: Excel de proyecciones nuevo, copia
  nueva del Sheet "Meta horaria" (si no, se sobrescribe el de julio) y
  decidir si se reutiliza el Sheet BACKEND.
- Panel en vivo: listo y probado con datos reales. Requiere activar
  cron-job.org cada 5 min la noche de cada evento (20:55 → 03:05) y que
  `panel.html` esté subido al cPanel.

## Cómo trabajar aquí

- No hay tests ni CI en PRs. Validar con: `python -m py_compile`, mocks de
  `requests.get` (simular respuestas de Yandex) y render del HTML con
  Playwright/Chromium a 1920×1080, 1912×958 y 1366×768.
- Para probar `panel.html` local: servir la carpeta y cambiar `FUENTES` a
  `["./panel.json"]`; Chart.js ya no se usa en el panel.
- Cambios por rama + PR draft a `main`; el usuario mergea. No forzar push.
- Workflows `workflow_dispatch` solo aparecen en Actions si el archivo está en `main`.
- Zona horaria: los workflows usan `TZ: America/Lima`; los crons de GitHub
  están en UTC (Lima = UTC−5) y son "best effort" (pueden atrasarse o saltarse).
- Comunicar con el usuario en español, sin jerga innecesaria; las métricas
  son **gross** (incluyen IGV); el número oficial es GA4 / Daily.
