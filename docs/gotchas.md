# Gotchas y bugs ya resueltos

Historial de problemas reales que costó diagnosticar. Antes de "arreglar" algo
que se parezca a esto, revisar acá primero.

## 1. Comparación de fechas sin normalizar → falsos positivos/negativos

**Síntoma:** una auditoría de ventas vs. Libro Diario marcaba como "faltante"
un cobro que ya estaba registrado, porque comparaba `"2026-07-27"` contra
`"2026-07-27T03:00:00.000Z"` como strings exactos. Correr la acción de
"corregir todo" sobre ese falso positivo generó **cobros duplicados reales**
en producción.

**Fix:** toda comparación de fechas pasa primero por una función de
normalización a `YYYY-MM-DD` (`normFechaPuntos()` en el frontend). Nunca
comparar `v.fecha === m.fecha` directo ni usar `.split('/')` asumiendo un
formato fijo.

**Bug relacionado:** el conteo de "ventas a archivar" antes del arqueo usaba
`fecha.split('/')` esperando `DD/MM/AAAA`. Con fechas ISO, el split no tira
error pero da un array que no matchea nada → siempre contaba 0 ventas, sin
ningún mensaje de error. Bug silencioso.

## 2. Backend y frontend usando formatos de período distintos

El frontend arma el período como `"AAAA/MM"` (`hoy.getFullYear()+'/'+mes`).
El Apps Script, al decidir qué ventas borrar de la hoja activa tras
copiarlas al histórico, armaba el período de cada venta como `"MM/AAAA"`.
Nunca coincidían → las ventas se copiaban a `Historico Ventas` pero **nunca
se borraban** de `Ventas`, generando duplicados silenciosos mes a mes.

**Fix:** ambos lados usan `"AAAA/MM"` como formato canónico de período.

## 3. Apps Script con errores de sintaxis que rompen TODO el backend

Un archivo `.gs` con una sola línea rota (ej: `/ comentario` con una sola
barra en vez de `//`, o una función sin cerrar al final del archivo) hace
que **ningún** endpoint funcione — ni `doGet` ni `doPost` — pero Google
Apps Script no siempre lo deja evidente: puede seguir sirviendo una
implementación vieja publicada, mientras el código nuevo (roto) queda sin
compilar en el editor.

**Síntoma confuso:** `doGet` (sin lógica, solo devuelve texto fijo) puede
funcionar bien mientras `doPost` (que sí toca `SpreadsheetApp`) falla, dando
la falsa impresión de un problema de permisos cuando en realidad es un
error de sintaxis.

**Cómo verificar rápido:** correr `node --check archivo.gs` (Apps Script es
JS válido en su mayoría) antes de pegar cualquier cambio en el editor de
Google.

## 4. "Proyecto sin título" no significa que esté desvinculado del Sheet

Un Apps Script atado a una planilla puede perfectamente decir "Proyecto sin
título" arriba si nunca se le puso nombre — no es indicador de que sea un
proyecto suelto. El indicador real de cuál Apps Script corresponde a la
planilla es entrar **siempre** por Extensiones → Apps Script desde dentro
del Sheet, nunca desde `script.google.com/home` directo.

## 5. Guardar el código en Apps Script no alcanza para que la app lo use

Hay que **re-implementar** explícitamente (Implementar → Administrar
implementaciones → lápiz → Versión: "Nueva versión" → Implementar) después
de cada cambio de código. Si no, la app web pública sigue sirviendo la
versión anterior indefinidamente, aunque el editor muestre el código nuevo
guardado.

## 6. Cartel de error que no se auto-limpia

Un mensaje de error de conexión mostrado en el DOM quedaba pegado en
pantalla aunque la siguiente carga de datos fuera exitosa — porque nunca se
programó la lógica de removerlo en el `try` exitoso, solo de mostrarlo en el
`catch`. Cualquier UI de error tiene que tener su contraparte de limpieza en
el camino feliz.

**Relacionado:** Google Apps Script suele fallar la primera llamada después
de un deploy nuevo ("cold start"). Conviene reintentar 2-3 veces con backoff
antes de mostrarle cualquier error al usuario.

## 7. Llamar a una función con el nombre equivocado (typo entre versiones)

Al portar lógica de una sesión/versión del archivo a otra, se llamó a
`normFecha()` cuando en ese archivo la función se llamaba
`normFechaPuntos()`. Como es un `ReferenceError` dentro de un `onclick`, el
navegador no muestra nada visible en la UI — el botón "no hace nada" a
simple vista. Solo aparece en la consola del navegador.

**Lección:** cualquier función helper reusada entre features (fechas, plata,
etc.) debería tener un nombre único y estable, documentado, para evitar este
tipo de mismatch al copiar/pegar bloques de lógica entre versiones del
archivo.
