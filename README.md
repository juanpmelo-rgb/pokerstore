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
  tocar las cajas en pesos de Guille/Juampi (fondo en dólares, canjes de
  puntos de fidelización), se usa el mismo Libro Diario pero con `quien` en
  un valor especial (`'usd'`, `'puntos'`). El cálculo de saldo de caja filtra
  por `quien==='guille'` o `'juampi'`, así que estos movimientos quedan
  invisibles para las cajas y para los totales del mes.

- **Códigos de barras**: viven en la hoja `Stock`, columna G (`Codigo`), como
  texto y separados por coma si un producto tiene varios. Se asocian solos:
  al escanear un código desconocido (en Ventas o en el escaneo de Stock) la
  app pide elegir el producto y llama a `asociarCodigo`. El escáner USB
  funciona como teclado (tipea el código + Enter); si no hay ningún campo
  activo, la app igual captura el código en las pestañas Ventas y Stock.

- **Período del arqueo**: `"AAAA/MM"` como texto, elegido en el modal (el
  cierre se hace a principio del mes siguiente). Las ventas no se archivan:
  los puntos de La Liga se calculan desde la hoja `Ventas`. Ver
  `docs/gotchas.md` #8 y #9.

Ver `docs/gotchas.md` para el historial de bugs ya resueltos y por qué.
