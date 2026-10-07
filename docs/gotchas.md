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

**Ojo:** ese fix (v13) resolvió solo la mitad del problema. Ver el punto 2b.

## 2b. `getValues()` devuelve Date reales, no texto — el bug de duplicados siguió vivo hasta v14

**Síntoma:** idéntico al del punto 2 (ventas copiadas a `Historico Ventas`
pero nunca borradas de `Ventas`), pero *después* de haber aplicado el fix de
la v13. Parecía que el fix no había tomado.

**Causa:** la v13 hacía `String(celda)` y matcheaba contra dos regex, ISO y
`DD/MM/AAAA`. Pero una celda formateada como fecha **no devuelve texto**:
`getValues()` devuelve un objeto `Date`. Y eso pasa siempre, porque cuando el
frontend escribe `"12/08/2026"` con `appendRow`, Sheets la convierte a `Date`
sola. `String(unDate)` da:

```
Wed Aug 12 2026 00:00:00 GMT-0300 (Argentina Standard Time)
```

que no matchea ninguno de los dos regex → el período quedaba vacío → la fila
nunca se borraba. Silencioso, sin error.

**La pista que lo delata:** si el frontend recibe fechas como
`"2026-08-12T03:00:00.000Z"`, es porque la celda es un `Date` real y
`JSON.stringify` lo serializó a ISO. O sea: ver ISO-con-hora en el frontend
es la prueba de que el backend está recibiendo `Date`, no strings.

**Fix (v14):** `periodoDeFecha_()` en el backend resuelve el caso `Date`
**antes** de intentar cualquier regex. Toda lectura de fecha desde la
planilla tiene que pasar por ahí — nunca `String(celda).match(...)` directo.

**Regla general:** en el backend, asumir siempre `Date`; en el frontend,
asumir siempre string (llega por JSON) y normalizar con `normFechaPuntos()`.

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

## 8. El cierre de mes se hace en el mes siguiente → el período no puede salir de "hoy"

**Síntoma:** en `Historico`, lo archivado como "septiembre" eran movimientos
de agosto (lo mismo con julio y agosto), y `Historico Ventas` nunca llegó a
crearse aunque ya se habían hecho tres cierres.

**Causa:** el arqueo calculaba el período con la fecha del día. Como el
cierre se hace los primeros días del mes siguiente (1/7, 2/8, 2/9), cerrar
junio quedaba etiquetado "julio", y las ventas que buscaba archivar eran las
del mes recién empezado (casi ninguna), así que las del mes cerrado quedaban
en `Ventas` para siempre.

**Agravante:** `appendRow` con `"2026/07"` hace que Sheets lo convierta en la
fecha 01/07/2026. El frontend después recibía `"2026-07-01T03:00:00.000Z"` y
`periodoLabel()` lo mostraba roto.

**Fix (v15 + frontend):** el modal de arqueo tiene un selector de mes (hasta
el día 15 sugiere el mes anterior) y avisa si ese mes ya tiene un arqueo.
El backend valida `AAAA/MM` y escribe el período con formato texto vía
`agregarFilas_()`. Los períodos viejos se corrigieron con
`backend/CorreccionPeriodos.gs` (se corre una sola vez), y el frontend igual
normaliza con `normPeriodo()` por si aparece alguno guardado como fecha.

## 9. Las ventas NO se archivan: de ahí salen los puntos de La Liga

El cierre de mes llegó a tener un paso que movía las ventas del mes a
`Historico Ventas`. Nunca funcionó en la práctica (ver 2b y 8), y cuando se
arregló se vio el problema de fondo: **los puntos de fidelización, los
resúmenes por canal/socio y el historial por cliente se calculan solo desde la
hoja `Ventas`**, y el frontend nunca lee `Historico Ventas`. Los puntos tienen
12 meses de vigencia, así que archivar ventas les borraba puntos a los
clientes (y los canjes, que viven en el Libro Diario, los dejaban en negativo).

**Decisión (v16):** el cierre solo archiva el Libro Diario; `Ventas` guarda
todas las ventas. Son ~200 por mes, Sheets lo banca bien por años. El backend
ignora `ventasData` si algún frontend viejo lo manda.

Si algún día `Ventas` crece demasiado: antes de archivar, el frontend tiene
que leer también el histórico para puntos y reportes — no al revés.

## 10. Cuentas de La Liga partidas por cómo se escribe el nombre

Hasta octubre 2026 la cuenta de puntos era el texto del campo Cliente de cada
venta, en minúsculas. "Héctor MDQ" y "hector mdq" eran dos cuentas, y la venta
creaba un cliente nuevo sola cada vez que el nombre no coincidía exacto, así
que la hoja Clientes juntó repetidos. También sumaban puntos los canales
("mercado libre") y las ventas canceladas.

Se reinició La Liga el 1/10/2026 (`LIGA_INICIO`: las ventas anteriores no
suman) y ahora:
- los nombres se comparan con `claveNombre()` / `claveNombre_()` (sin tildes,
  sin mayúsculas, sin espacios de más) — tienen que seguir siendo iguales en
  frontend y backend;
- cada cliente que suma puntos tiene DNI, y nombre y DNI son únicos (el
  backend lo valida también, por si Juampi y Guille cargan a la vez);
- la venta ya no crea clientes sola: si el nombre es nuevo pide el DNI, y si
  el DNI ya existe ofrece usar ese cliente;
- solo suman ventas Pagadas/Entregadas.

Los repetidos viejos se unieron una sola vez con `backend/UnificarClientes.gs`.

## 11. Ventas y caja desincronizadas

En octubre 2026 los saldos de caja "no daban". La cuenta del sistema estaba
bien (cada cierre = saldo anterior + cobros − gastos − retiros, diferencia 0);
lo que fallaba era la carga:
- una venta cargada como "Abierto" y después pasada a "Pagado" nunca sumaba
  su cobro (el cobro solo se creaba al registrar la venta);
- borrar una venta dejaba su cobro en el Libro Diario;
- editar el monto de una venta no tocaba el cobro;
- con un filtro de canal activo, ✏️/🗑 usaban el índice de la lista filtrada
  y editaban/borraban OTRA venta.

Ahora toda venta nueva se registra como Pagado con su cobro, y editar o
borrar una venta corrige su cobro si sigue en el Libro Diario (se busca con
`idxCobroDeVenta()`: misma fecha, monto, quien y "Venta <canal>:"). Si ya se
archivó en un cierre, la app avisa que hay que cargar el ajuste a mano.
Ventas y movimientos guardan `fila` (número de fila en la hoja) y se editan
por fila, no por posición en la lista.
