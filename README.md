# Pokerstore — Sistema de Gestión Financiera

Sistema de gestión financiera y operativa para Pokerstore Argentina: libro diario,
stock, ventas, precios, presupuestos, fondo en dólares para importación, programa
de puntos de fidelización y escaneo de código de barras — todo en una sola app web
conectada a Google Sheets.

## Arquitectura (importante leer antes de tocar código)

No hay servidor propio ni base de datos tradicional. Son tres piezas separadas:

```
┌─────────────────┐       POST/GET        ┌──────────────────┐       lee/escribe      ┌───────────────┐
│  frontend/       │ ────────────────────> │  backend/         │ ─────────────────────> │  Google Sheet  │
│  index.html      │  fetch(API, {...})    │  Codigo.gs         │  SpreadsheetApp        │  (la "base de  │
│  (Netlify)       │ <──────────────────── │  (Google Apps      │ <───────────────────── │   datos")      │
└─────────────────┘      JSON              │   Script)           │                        └───────────────┘
                                            └──────────────────┘
```

- **`frontend/index.html`** — Todo el frontend en un solo archivo (HTML+CSS+JS
  embebido, sin build, sin dependencias externas salvo CDNs puntuales). Se
  despliega tal cual a **Netlify**. No hay React ni bundler: es un HTML que se
  edita directo y se sube.

- **`backend/Codigo.gs`** — El backend. Vive pegado en el editor de **Google
  Apps Script**, atado a la planilla de Google Sheets (Extensiones → Apps
  Script desde el Sheet). Recibe `doPost`/`doGet`, hace `SpreadsheetApp.getActiveSpreadsheet()`
  y lee/escribe filas. Se publica como "Aplicación web" con acceso "Cualquier
  usuario" para que el frontend le pueda pegar sin login.

- **Google Sheet** (no versionado acá, vive en Drive) — la base de datos real.
  Pestañas: `Libro Diario`, `Stock`, `Ventas`, `Precios`, `Configuracion`,
  `Clientes`, `Proveedores`, `Adjuntos`, `Arqueos`, `Historico`,
  `Historico Ventas`.

## Cómo desplegar un cambio

**Frontend:**
1. Editar `frontend/index.html`
2. Validar sintaxis del JS embebido antes de subir (ver `docs/gotchas.md`)
3. Subir el archivo a Netlify (deploy manual o conectar el repo a Netlify para
   deploy automático en push — pendiente de configurar)

**Backend:**
1. Editar `backend/Codigo.gs`
2. Pegar el contenido completo en el editor de Apps Script (Extensiones →
   Apps Script desde el Sheet — **no** un proyecto de Apps Script suelto,
   tiene que estar atado a la planilla)
3. **Implementar → Administrar implementaciones → lápiz ✏️ → Versión: "Nueva
   versión" → Implementar.** Este paso es obligatorio: guardar el código NO
   alcanza, si no se re-implementa la app web sigue sirviendo la versión vieja.
4. Verificar que "Quién tiene acceso" siga en "Cualquier usuario"

## Convenciones de datos que hay que conocer

- **Formato de fecha inconsistente**: las ventas y movimientos pueden venir en
  ISO (`2026-08-12T03:00:00.000Z`), ISO corto (`2026-08-12`) o argentino
  (`12/08/2026`), mezclados en la misma hoja. Cualquier código que compare o
  filtre por fecha **tiene que normalizar primero**. Ver `normFechaPuntos()`
  en el frontend como referencia — es la función a reusar, no reinventarla.

- **Namespacing por `quien` en el Libro Diario**: para features que no deben
  tocar las cajas en pesos de Guille/Juampi (fondo en dólares, puntos de
  fidelización, asociaciones de código de barras), se usa el mismo Libro
  Diario pero con `quien` en un valor especial (`'usd'`, `'puntos'`,
  `'codigos'`). El cálculo de saldo de caja filtra por `quien==='guille'` o
  `'juampi'`, así que estos movimientos quedan invisibles para las cajas y
  para los totales del mes. Ver `docs/gotchas.md` para el detalle de cada uno.

- **Formato de período para el arqueo**: el frontend manda `periodo` como
  `"AAAA/MM"` (ej: `"2026/08"`). El backend arma el mes/año de cada venta y
  compara contra ese string — tienen que coincidir en formato exacto.

Ver `docs/gotchas.md` para el historial de bugs ya resueltos y por qué.
