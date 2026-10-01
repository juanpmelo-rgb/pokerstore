// ============================================================
//  UNIFICACIÓN ÚNICA DE CLIENTES (octubre 2026)
//
//  Antes de exigir nombre y DNI únicos, la hoja Clientes tenía el mismo
//  cliente cargado dos veces ("Nacho" / "nacho", "Héctor MDQ" / "hector mdq").
//  Este script:
//    · une los clientes repetidos (mismo nombre sin importar tildes ni
//      mayúsculas): queda la primera fila, completada con los datos que
//      falten de las otras, y se borran las demás
//    · agrega los títulos DNI (col I) y SumaPuntos (col J)
//    · marca "no suma puntos" a Mercado Libre, Tienda Nube y Héctor MDQ
//      (y crea el cliente Tienda Nube si no existe)
//  Las ventas no se tocan: se unen a su cliente por el nombre.
//
//  Uso, desde el editor de Apps Script:
//    1. Ejecutar revisarUnificacion() → solo muestra en el registro qué cambiaría.
//    2. Si está bien, ejecutar aplicarUnificacion().
//  Correrlo dos veces no hace nada. Una vez aplicado, este archivo se puede borrar.
//  Usa claveNombre_() y soloDigitos_() de Codigo.gs.
// ============================================================

var CLIENTES_SIN_PUNTOS = ['mercado libre', 'mercadolibre', 'tienda nube', 'tiendanube', 'hector mdq'];

function revisarUnificacion() {
  var plan = planUnificacion_();
  mostrarUnificacion_(plan);
  if (plan.unir.length || plan.sinPuntos.length || plan.crearTiendaNube || plan.titulos) {
    console.log('Si esto es correcto, ejecutá aplicarUnificacion().');
  }
}

function aplicarUnificacion() {
  var plan = planUnificacion_();
  mostrarUnificacion_(plan);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CLIENTES);

  asegurarEncabezadosClientes_(sh);
  plan.unir.forEach(function(u) {
    sh.getRange(u.fila, 1, 1, 10).setValues([u.datos]);
  });
  plan.sinPuntos.forEach(function(c) {
    sh.getRange(c.fila, 10).setValue('no');
  });
  // De abajo hacia arriba, para que no se corran los números de fila
  var borrar = [];
  plan.unir.forEach(function(u) { borrar = borrar.concat(u.borrar); });
  borrar.sort(function(a, b) { return b - a; }).forEach(function(f) { sh.deleteRow(f); });
  if (plan.crearTiendaNube) {
    agregarFilas_(sh, [['Tienda Nube', 'Corporativo', '', '', '', new Date(), 'Canal de venta: no suma puntos', 'tienda', '', 'no']], [9]);
  }

  console.log('✓ Unificación aplicada. Recargá el sistema para ver los cambios.');
}

function planUnificacion_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CLIENTES);
  var data = sh.getDataRange().getValues();
  var plan = { unir: [], sinPuntos: [], crearTiendaNube: false, titulos: false };
  plan.titulos = !String(data[0][8] || '').trim() || !String(data[0][9] || '').trim();

  // Agrupar por nombre
  var grupos = {}, orden = [];
  for (var i = 1; i < data.length; i++) {
    var k = claveNombre_(data[i][0]);
    if (!k) continue;
    if (!grupos[k]) { grupos[k] = []; orden.push(k); }
    grupos[k].push(i);
  }

  orden.forEach(function(k) {
    var filas = grupos[k];
    var datos = [];
    for (var c = 0; c < 10; c++) datos.push(data[filas[0]][c] === undefined ? '' : data[filas[0]][c]);
    var sinPuntos = CLIENTES_SIN_PUNTOS.indexOf(k) >= 0;

    if (filas.length > 1) {
      // La primera fila manda; de las otras solo se toman los datos que falten
      filas.slice(1).forEach(function(f) {
        for (var c = 0; c < 10; c++) {
          if (String(datos[c]).trim() === '' && data[f][c] !== undefined && String(data[f][c]).trim() !== '') datos[c] = data[f][c];
        }
      });
      datos[8] = soloDigitos_(datos[8]);
      if (sinPuntos) datos[9] = 'no';
      plan.unir.push({ fila: filas[0] + 1, nombre: datos[0], datos: datos,
                       borrar: filas.slice(1).map(function(f) { return f + 1; }),
                       nombres: filas.map(function(f) { return data[f][0]; }) });
    } else if (sinPuntos && String(datos[9]).trim().toLowerCase() !== 'no') {
      plan.sinPuntos.push({ fila: filas[0] + 1, nombre: datos[0] });
    }
  });

  plan.crearTiendaNube = !grupos['tienda nube'] && !grupos['tiendanube'];
  return plan;
}

function mostrarUnificacion_(plan) {
  if (!plan.unir.length && !plan.sinPuntos.length && !plan.crearTiendaNube && !plan.titulos) {
    console.log('Nada para hacer: no hay clientes repetidos y las cuentas sin puntos ya están marcadas.');
    return;
  }
  console.log('CLIENTES REPETIDOS a unir: ' + plan.unir.length);
  plan.unir.forEach(function(u) {
    console.log('  ' + u.nombres.join(' + ') + '  →  queda "' + u.nombre + '" (fila ' + u.fila + '), se borran las filas ' + u.borrar.join(', ') +
                (u.datos[9] === 'no' ? '  [no suma puntos]' : ''));
  });
  console.log('NO SUMAN PUNTOS:');
  plan.sinPuntos.forEach(function(c) { console.log('  ' + c.nombre + ' (fila ' + c.fila + ')'); });
  if (plan.crearTiendaNube) console.log('  Tienda Nube (cliente nuevo)');
  if (plan.titulos) console.log('TÍTULOS: se agregan "DNI" (col I) y "SumaPuntos" (col J)');
}
