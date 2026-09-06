// ============================================================
//  POKERSTORE — APPS SCRIPT BACKEND v14 — el archivado de ventas ahora reconoce
//  celdas con Date real (SISTEMA NÚCLEO)
//  v14: periodoDeFecha_() maneja Date | ISO | DD/MM/AAAA. La v13 solo matcheaba
//       texto, y como las celdas de fecha son Date reales, las ventas se copiaban
//       a "Historico Ventas" pero nunca se borraban de "Ventas" → duplicados.
//  v13 corrigió dos errores de sintaxis de la v11:
//    1) La primera línea tenía UNA sola barra "/" en vez de "//"
//    2) La función obtenerCarpeta_() quedó sin cerrar al final
//  Libro Diario col G = Proveedor · Historico col H = Proveedor
//  Hojas: Libro Diario | Stock | Configuracion | Ventas
//         Historico | Arqueos | Clientes | Proveedores | Adjuntos | Precios | Historico Ventas
// ============================================================

const SHEET_DIARIO   = 'Libro Diario';
const SHEET_STOCK    = 'Stock';
const SHEET_CONFIG   = 'Configuracion';
const SHEET_VENTAS   = 'Ventas';
const SHEET_HISTORICO= 'Historico';
const SHEET_ARQUEOS  = 'Arqueos';
const SHEET_CLIENTES = 'Clientes';
const SHEET_PROVEED  = 'Proveedores';
const SHEET_ADJUNTOS = 'Adjuntos';
const SHEET_PRECIOS  = 'Precios';

// Carpeta de Drive donde se guardan los adjuntos (se crea sola si no existe)
const CARPETA_ADJUNTOS = 'Pokerstore - Adjuntos';

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var result;
  try { result = handleAction(body); }
  catch(err) { result = { ok: false, error: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, msg: 'Pokerstore API v14 activa' })).setMimeType(ContentService.MimeType.JSON);
}

function handleAction(b) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── GET ALL ────────────────────────────────────────────────
  if (b.action === 'getAll') {
    return {
      ok: true,
      diario:      getSheetData(ss, SHEET_DIARIO),
      stock:       getSheetData(ss, SHEET_STOCK),
      config:      getSheetData(ss, SHEET_CONFIG),
      ventas:      getSheetData(ss, SHEET_VENTAS),
      historico:   getSheetData(ss, SHEET_HISTORICO),
      arqueos:     getSheetData(ss, SHEET_ARQUEOS),
      clientes:    getSheetData(ss, SHEET_CLIENTES),
      proveedores: getSheetData(ss, SHEET_PROVEED),
      adjuntos:    getSheetData(ss, SHEET_ADJUNTOS),
      precios:     getSheetData(ss, SHEET_PRECIOS)
    };
  }

  // ── LIBRO DIARIO ───────────────────────────────────────────
  if (b.action === 'addMovimiento') {
    ss.getSheetByName(SHEET_DIARIO).appendRow([b.fecha, b.desc, b.cat, b.quien, b.tipo, b.monto, b.proveedor || '']);
    return { ok: true };
  }
  if (b.action === 'deleteMovimiento') {
    ss.getSheetByName(SHEET_DIARIO).deleteRow(Number(b.row) + 2);
    return { ok: true };
  }
  if (b.action === 'updateMovimiento') {
    ss.getSheetByName(SHEET_DIARIO).getRange(Number(b.row) + 2, 1, 1, 7)
      .setValues([[b.fecha, b.desc, b.cat, b.quien, b.tipo, b.monto, b.proveedor || '']]);
    return { ok: true };
  }

  // ── STOCK ──────────────────────────────────────────────────
  if (b.action === 'addProducto') {
    ss.getSheetByName(SHEET_STOCK).appendRow([b.nombre, b.cat, b.stock, b.costo, b.precio, b.minstock]);
    return { ok: true };
  }
  if (b.action === 'deleteProducto') {
    ss.getSheetByName(SHEET_STOCK).deleteRow(Number(b.row) + 2);
    return { ok: true };
  }
  if (b.action === 'updateStock') {
    var sheet = ss.getSheetByName(SHEET_STOCK);
    var data = sheet.getDataRange().getValues();
    var nombre = String(b.nombre).trim();
    var delta = Number(b.delta);
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === nombre) {
        var nuevo = Math.max(0, (Number(data[i][2])||0) + delta);
        sheet.getRange(i + 1, 3).setValue(nuevo);
        if (b.nuevoCosto && Number(b.nuevoCosto) > 0) sheet.getRange(i + 1, 4).setValue(Number(b.nuevoCosto));
        return { ok: true, stockNuevo: nuevo };
      }
    }
    return { ok: false, error: 'Producto no encontrado: ' + nombre };
  }

  // ── VENTAS ─────────────────────────────────────────────────
  if (b.action === 'addVenta') {
    ss.getSheetByName(SHEET_VENTAS).appendRow([b.fecha, b.canal, b.quien, b.cliente, b.producto, b.cantidad, b.monto, b.pago, b.estado, b.orden, b.evento]);
    return { ok: true };
  }
  if (b.action === 'deleteVenta') {
    ss.getSheetByName(SHEET_VENTAS).deleteRow(Number(b.row) + 2);
    return { ok: true };
  }
  if (b.action === 'updateVentaEstado') {
    ss.getSheetByName(SHEET_VENTAS).getRange(Number(b.row) + 2, 9).setValue(b.estado);
    return { ok: true };
  }
  if (b.action === 'updateVenta') {
    ss.getSheetByName(SHEET_VENTAS).getRange(Number(b.row) + 2, 1, 1, 11)
      .setValues([[b.fecha, b.canal, b.quien, b.cliente, b.producto, b.cantidad, b.monto, b.pago, b.estado, b.orden, b.evento]]);
    return { ok: true };
  }

  // ── CONFIG / MARKUPS ───────────────────────────────────────
  if (b.action === 'saveConfig') {
    var sheet = ss.getSheetByName(SHEET_CONFIG);
    var cfg = b.config;
    var data = sheet.getDataRange().getValues();
    var updates = { cajaGuille: cfg.cajaGuille, cajaJuampi: cfg.cajaJuampi, mayvol: cfg.mayvol, mayest: cfg.mayest, retail: cfg.retail, web: cfg.web };
    for (var i = 1; i < data.length; i++) {
      var key = String(data[i][0]);
      if (key in updates) { sheet.getRange(i + 1, 2).setValue(updates[key]); delete updates[key]; }
    }
    Object.keys(updates).forEach(function(k){ sheet.appendRow([k, updates[k]]); });
    return { ok: true };
  }

  // ── ARQUEO DE CAJA / CIERRE DE MES ─────────────────────────
  if (b.action === 'arqueoMes') {
    var diario = ss.getSheetByName(SHEET_DIARIO);
    var histo  = ss.getSheetByName(SHEET_HISTORICO);
    var arq    = ss.getSheetByName(SHEET_ARQUEOS);
    var config = ss.getSheetByName(SHEET_CONFIG);

    // 1. Mover movimientos actuales a Histórico con el período
    var movs = diario.getDataRange().getValues();
    for (var i = 1; i < movs.length; i++) {
      if (movs[i][1]) { // tiene descripción
        histo.appendRow([movs[i][0], movs[i][1], movs[i][2], movs[i][3], movs[i][4], movs[i][5], b.periodo, movs[i][6] || '']);
      }
    }

    // 2. Registrar el arqueo
    arq.appendRow([b.fecha, b.periodo, b.cajaG, b.cajaJ, b.total, b.valorStock, b.movs]);

    // 3. Actualizar saldos en Config (saldo final = saldo inicial del mes nuevo)
    var cfgData = config.getDataRange().getValues();
    var saldoUpdates = { cajaGuille: b.cajaG, cajaJuampi: b.cajaJ };
    for (var j = 1; j < cfgData.length; j++) {
      var k = String(cfgData[j][0]);
      if (k in saldoUpdates) { config.getRange(j + 1, 2).setValue(saldoUpdates[k]); delete saldoUpdates[k]; }
    }
    Object.keys(saldoUpdates).forEach(function(k){ config.appendRow([k, saldoUpdates[k]]); });

    // 4. Limpiar Libro Diario (dejar solo headers)
    var lastRow = diario.getLastRow();
    if (lastRow > 1) diario.deleteRows(2, lastRow - 1);

    // 5. ARCHIVADO DE VENTAS
    if (b.ventasData && b.ventasData.length > 0) {
      var ventasSheet = ss.getSheetByName(SHEET_VENTAS);
      var historicoVentasSheet = ss.getSheetByName('Historico Ventas');

      // Crear hoja "Historico Ventas" si no existe
      if (!historicoVentasSheet) {
        historicoVentasSheet = ss.insertSheet('Historico Ventas');
        var headerVentas = ventasSheet.getRange(1, 1, 1, ventasSheet.getLastColumn()).getValues()[0];
        headerVentas.push('periodo');
        historicoVentasSheet.appendRow(headerVentas);
      }

      // Agregar las ventas a Historico Ventas
      var datosVentas = [];
      for (var v = 0; v < b.ventasData.length; v++) {
        var venta = b.ventasData[v];
        datosVentas.push([
          venta.fecha, venta.canal, venta.quien, venta.cliente, venta.producto,
          venta.cantidad, venta.monto, venta.pago, venta.estado, venta.orden, venta.evento, b.periodo
        ]);
      }
      if (datosVentas.length > 0) {
        historicoVentasSheet.getRange(historicoVentasSheet.getLastRow() + 1, 1, datosVentas.length, datosVentas[0].length).setValues(datosVentas);
      }

      // Borrar las ventas del mes de la hoja Ventas activa (de abajo hacia arriba)
      // El período de cada venta se calcula con periodoDeFecha_(), que maneja
      // los tres formatos que puede devolver la hoja (Date real, ISO, argentino).
      var ventasActuales = ventasSheet.getDataRange().getValues();
      for (var w = ventasActuales.length - 1; w >= 1; w--) {
        var mesAnioVenta = periodoDeFecha_(ventasActuales[w][0]);
        if (mesAnioVenta && mesAnioVenta === b.periodo) {
          ventasSheet.deleteRow(w + 1);
        }
      }
    }

    return { ok: true };
  }

  // ── CLIENTES ───────────────────────────────────────────────
  if (b.action === 'addCliente') {
    ss.getSheetByName(SHEET_CLIENTES).appendRow([b.nombre, b.tipo, b.telefono, b.email, b.localidad, b.ultconsulta, b.notas, b.lista]);
    return { ok: true };
  }
  if (b.action === 'updateCliente') {
    var sh = ss.getSheetByName(SHEET_CLIENTES);
    var fila = Number(b.row) + 2;
    sh.getRange(fila, 1, 1, 8).setValues([[b.nombre, b.tipo, b.telefono, b.email, b.localidad, b.ultconsulta, b.notas, b.lista]]);
    return { ok: true };
  }
  if (b.action === 'deleteCliente') {
    ss.getSheetByName(SHEET_CLIENTES).deleteRow(Number(b.row) + 2);
    return { ok: true };
  }

  // ── PROVEEDORES ────────────────────────────────────────────
  if (b.action === 'addProveedor') {
    ss.getSheetByName(SHEET_PROVEED).appendRow([b.nombre, b.provee, b.contacto, b.telefono, b.email, b.frecuencia]);
    return { ok: true };
  }
  if (b.action === 'deleteProveedor') {
    ss.getSheetByName(SHEET_PROVEED).deleteRow(Number(b.row) + 2);
    return { ok: true };
  }

  // ── ADJUNTOS (subida a Drive) ──────────────────────────────
  if (b.action === 'uploadAdjunto') {
    var carpeta = obtenerCarpeta_();
    var blob = Utilities.newBlob(Utilities.base64Decode(b.fileData), b.mimeType, b.filename);
    var archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = archivo.getUrl();
    ss.getSheetByName(SHEET_ADJUNTOS).appendRow([b.fecha, b.tipo, b.desc, b.relacionado, url]);
    return { ok: true, url: url };
  }
  if (b.action === 'deleteAdjunto') {
    ss.getSheetByName(SHEET_ADJUNTOS).deleteRow(Number(b.row) + 2);
    return { ok: true };
  }

  return { ok: false, error: 'Acción desconocida: ' + b.action };
}

// Devuelve el período "AAAA/MM" de una fecha de la planilla, o '' si no se puede leer.
//
// IMPORTANTE: getValues() devuelve un objeto Date real (no texto) cuando la celda
// está formateada como fecha, que es el caso normal — al escribir "12/08/2026" con
// appendRow, Sheets la convierte en Date automáticamente. Ese caso hay que
// resolverlo ANTES de intentar cualquier regex: String(unDate) da
// "Wed Aug 12 2026 00:00:00 GMT-0300 (...)", que no matchea ni ISO ni DD/MM/AAAA
// y hacía que la venta nunca se borrara de la hoja activa (duplicados mes a mes).
function periodoDeFecha_(valor) {
  if (!valor && valor !== 0) return '';

  // Caso 1: la celda es un Date real (el más común)
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    if (isNaN(valor.getTime())) return '';
    var mes = valor.getMonth() + 1;
    return valor.getFullYear() + '/' + (mes < 10 ? '0' + mes : String(mes));
  }

  var s = String(valor).trim();
  if (!s) return '';

  // Caso 2: texto ISO ("2026-08-12" o "2026-08-12T03:00:00.000Z")
  var m1 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m1) return m1[1] + '/' + m1[2];

  // Caso 3: texto argentino ("12/08/2026")
  var m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m2) return m2[3] + '/' + (m2[2].length === 1 ? '0' + m2[2] : m2[2]);

  return '';
}

function getSheetData(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  return sheet.getDataRange().getValues();
}

function obtenerCarpeta_() {
  var carpetas = DriveApp.getFoldersByName(CARPETA_ADJUNTOS);
  if (carpetas.hasNext()) return carpetas.next();
  return DriveApp.createFolder(CARPETA_ADJUNTOS);
}
