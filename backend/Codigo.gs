// ============================================================
//  POKERSTORE — APPS SCRIPT BACKEND v19 — clientes con DNI único
//  (SISTEMA NÚCLEO)
//  v19: Clientes col I = DNI y col J = SumaPuntos. addCliente/updateCliente
//       rechazan nombres (sin importar tildes/mayúsculas) y DNI repetidos.
//       Renombrar un cliente renombra sus ventas y canjes de La Liga.
//  v18: asociarCodigo responde ok si el código ya estaba en ese mismo producto,
//       para que el frontend pueda reintentar cuando Google devuelve un error.
//  v17: asociarCodigo / quitarCodigo guardan los códigos en Stock col G (texto,
//       varios separados por coma) en vez de filas 'codigos' en el Libro Diario.
//  v16: arqueoMes solo archiva el Libro Diario. Las ventas quedan en "Ventas"
//       porque de ahí salen los puntos de La Liga (vigencia 12 meses): moverlas
//       a "Historico Ventas" les borraba los puntos a los clientes.
//  v15: arqueoMes valida el período (AAAA/MM) y lo escribe con formato texto
//       vía agregarFilas_(). Antes, Sheets convertía "2026/09" en una fecha.
//  v13 corrigió dos errores de sintaxis de la v11:
//    1) La primera línea tenía UNA sola barra "/" en vez de "//"
//    2) La función obtenerCarpeta_() quedó sin cerrar al final
//  Libro Diario col G = Proveedor · Historico col H = Proveedor · Stock col G = Codigo
//  Hojas: Libro Diario | Stock | Configuracion | Ventas
//         Historico | Arqueos | Clientes | Proveedores | Adjuntos | Precios
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
  return ContentService.createTextOutput(JSON.stringify({ ok: true, msg: 'Pokerstore API v19 activa' })).setMimeType(ContentService.MimeType.JSON);
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

  // ── CÓDIGOS DE BARRAS (Stock col G, varios separados por coma) ──
  if (b.action === 'asociarCodigo') {
    var sheet = ss.getSheetByName(SHEET_STOCK);
    var data = sheet.getDataRange().getValues();
    var nombre = String(b.nombre).trim();
    var codigo = String(b.codigo || '').trim();
    if (!codigo || codigo.indexOf(',') >= 0) return { ok: false, error: 'Código inválido: ' + codigo };
    var fila = -1;
    for (var i = 1; i < data.length; i++) {
      var nom = String(data[i][0]).trim();
      if (codigosDeCelda_(data[i][6]).indexOf(codigo) >= 0) {
        // Mismo producto = un reintento de algo que ya se guardó
        if (nom === nombre) return { ok: true, yaEstaba: true };
        return { ok: false, error: 'El código ' + codigo + ' ya está asociado a "' + nom + '"' };
      }
      if (fila < 0 && nom === nombre) fila = i + 1;
    }
    if (fila < 0) return { ok: false, error: 'Producto no encontrado: ' + nombre };
    var actuales = codigosDeCelda_(data[fila - 1][6]);
    actuales.push(codigo);
    asegurarEncabezadoCodigo_(sheet, data);
    // Como texto: si no, Sheets muestra un EAN-13 como 7,79E+12 y se come los ceros iniciales
    sheet.getRange(fila, 7).setNumberFormat('@').setValue(actuales.join(', '));
    return { ok: true };
  }
  if (b.action === 'quitarCodigo') {
    var sheet = ss.getSheetByName(SHEET_STOCK);
    var data = sheet.getDataRange().getValues();
    var codigo = String(b.codigo || '').trim();
    for (var i = 1; i < data.length; i++) {
      var actuales = codigosDeCelda_(data[i][6]);
      var pos = actuales.indexOf(codigo);
      if (pos >= 0) {
        actuales.splice(pos, 1);
        sheet.getRange(i + 1, 7).setNumberFormat('@').setValue(actuales.join(', '));
        return { ok: true };
      }
    }
    return { ok: false, error: 'Código no encontrado: ' + codigo };
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

    if (!/^\d{4}\/\d{2}$/.test(String(b.periodo))) {
      return { ok: false, error: 'Período inválido: "' + b.periodo + '" (se espera AAAA/MM)' };
    }

    // 1. Mover movimientos actuales a Histórico con el período
    var movs = diario.getDataRange().getValues();
    var filasHisto = [];
    for (var i = 1; i < movs.length; i++) {
      if (movs[i][1]) { // tiene descripción
        filasHisto.push([movs[i][0], movs[i][1], movs[i][2], movs[i][3], movs[i][4], movs[i][5], b.periodo, movs[i][6] || '']);
      }
    }
    agregarFilas_(histo, filasHisto, [7]);

    // 2. Registrar el arqueo
    agregarFilas_(arq, [[b.fecha, b.periodo, b.cajaG, b.cajaJ, b.total, b.valorStock, b.movs]], [2]);

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

    // Las ventas NO se archivan: los puntos de La Liga, los resúmenes y el
    // historial por cliente se calculan desde la hoja Ventas. Si un frontend
    // viejo manda ventasData, se ignora a propósito.

    return { ok: true };
  }

  // ── CLIENTES ───────────────────────────────────────────────
  // Cols: A Nombre · B Tipo · C Telefono · D Email · E Localidad · F UltimaConsulta
  //       G Notas · H Lista · I DNI (texto) · J SumaPuntos ('no' = no suma en La Liga)
  // Ni el nombre (comparado con claveNombre_) ni el DNI pueden repetirse.
  if (b.action === 'addCliente') {
    var shA = ss.getSheetByName(SHEET_CLIENTES);
    var errA = validarCliente_(shA, b, -1);
    if (errA) return { ok: false, error: errA };
    asegurarEncabezadosClientes_(shA);
    agregarFilas_(shA, [filaCliente_(b)], [9]);
    return { ok: true };
  }
  if (b.action === 'updateCliente') {
    var sh = ss.getSheetByName(SHEET_CLIENTES);
    var fila = Number(b.row) + 2;
    var err = validarCliente_(sh, b, fila);
    if (err) return { ok: false, error: err };
    asegurarEncabezadosClientes_(sh);
    var nombreViejo = sh.getRange(fila, 1).getValue();
    sh.getRange(fila, 9).setNumberFormat('@');
    sh.getRange(fila, 1, 1, 10).setValues([filaCliente_(b)]);
    // Si cambió el nombre, se cambia también en sus ventas y canjes para que
    // no pierda los puntos (La Liga une ventas y clientes por el nombre).
    var renombradas = 0;
    if (claveNombre_(nombreViejo) && claveNombre_(nombreViejo) !== claveNombre_(b.nombre)) {
      renombradas = renombrarEnHoja_(ss.getSheetByName(SHEET_VENTAS), 4, null, nombreViejo, b.nombre)
                  + renombrarEnHoja_(ss.getSheetByName(SHEET_DIARIO), 7, 4, nombreViejo, b.nombre)
                  + renombrarEnHoja_(ss.getSheetByName(SHEET_HISTORICO), 8, 4, nombreViejo, b.nombre);
    }
    return { ok: true, renombradas: renombradas };
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

// Agrega filas al final de la hoja. Las columnas de colsTexto (1-based) se
// formatean como texto ANTES de escribir: si no, Sheets convierte "2026/09"
// en la fecha 01/09/2026 y el período se pierde.
function agregarFilas_(sheet, filas, colsTexto) {
  if (!filas.length) return;
  var desde = sheet.getLastRow() + 1;
  (colsTexto || []).forEach(function(c) {
    sheet.getRange(desde, c, filas.length, 1).setNumberFormat('@');
  });
  sheet.getRange(desde, 1, filas.length, filas[0].length).setValues(filas);
}

function codigosDeCelda_(valor) {
  return String(valor == null ? '' : valor).split(',')
    .map(function(s) { return s.trim(); })
    .filter(function(s) { return s; });
}

// En Stock la fila de títulos ("Nombre", "Categoria", ...) no es la primera:
// arriba hay una fila marcador. Se busca por contenido.
function asegurarEncabezadoCodigo_(sheet, data) {
  for (var i = 0; i < Math.min(3, data.length); i++) {
    if (String(data[i][0]).trim() === 'Nombre') {
      if (!String(data[i][6] || '').trim()) sheet.getRange(i + 1, 7).setValue('Codigo');
      return;
    }
  }
}

// Clave para comparar nombres de clientes: sin tildes, sin mayúsculas, sin
// espacios de más. Igual que claveNombre() en el frontend.
function claveNombre_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}
function soloDigitos_(v) {
  return String(v == null ? '' : v).replace(/\D/g, '');
}

function filaCliente_(b) {
  return [String(b.nombre || '').trim(), b.tipo || '', b.telefono || '', b.email || '', b.localidad || '',
          b.ultconsulta || '', b.notas || '', b.lista || '', soloDigitos_(b.dni),
          String(b.sumaPuntos || '').toLowerCase() === 'no' ? 'no' : ''];
}

// Devuelve el texto del error, o '' si el cliente se puede guardar.
// filaPropia = fila que se está editando (para no compararla consigo misma); -1 al agregar.
function validarCliente_(sheet, b, filaPropia) {
  var clave = claveNombre_(b.nombre);
  var dni = soloDigitos_(b.dni);
  if (!clave) return 'Falta el nombre del cliente';
  if (dni && !/^\d{7,11}$/.test(dni)) return 'DNI inválido: ' + b.dni;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (i + 1 === filaPropia) continue;
    if (claveNombre_(data[i][0]) === clave) return 'Ya existe un cliente llamado "' + data[i][0] + '"';
    if (dni && soloDigitos_(data[i][8]) === dni) return 'El DNI ' + dni + ' ya está registrado a nombre de "' + data[i][0] + '"';
  }
  return '';
}

function asegurarEncabezadosClientes_(sheet) {
  if (!String(sheet.getRange(1, 9).getValue()).trim()) sheet.getRange(1, 9).setValue('DNI');
  if (!String(sheet.getRange(1, 10).getValue()).trim()) sheet.getRange(1, 10).setValue('SumaPuntos');
}

// Reemplaza el nombre de un cliente en la columna colNombre (1-based). Si se
// pasa colQuien, solo en las filas con quien='puntos' (canjes de La Liga).
function renombrarEnHoja_(sheet, colNombre, colQuien, viejo, nuevo) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var data = sheet.getDataRange().getValues();
  var clave = claveNombre_(viejo), n = 0;
  for (var i = 1; i < data.length; i++) {
    if (colQuien && String(data[i][colQuien - 1]).trim() !== 'puntos') continue;
    if (claveNombre_(data[i][colNombre - 1]) === clave) {
      sheet.getRange(i + 1, colNombre).setValue(nuevo);
      n++;
    }
  }
  return n;
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
