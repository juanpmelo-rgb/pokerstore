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

// Mismo nombre pero personas distintas: se renombran en vez de unirse.
// Se reconocen por el teléfono, no por la fila.
var CLIENTES_A_RENOMBRAR = [
  { nombre: 'nacho', telefono: '1144477616', nuevo: 'Nacho (4447-7616)' }
];

// Columnas que, si tienen valores distintos, indican que son dos personas
// (0-based: C Telefono, D Email, I DNI)
var COLS_IDENTIDAD = [2, 3, 8];

function revisarUnificacion() {
  var plan = planUnificacion_();
  mostrarUnificacion_(plan);
  if (plan.errores.length) {
    console.log('⛔ Hay problemas: NO ejecutes aplicarUnificacion() hasta revisarlos.');
  } else if (plan.renombrar.length || plan.unir.length || plan.sinPuntos.length || plan.crearTiendaNube || plan.titulos) {
    console.log('Si esto es correcto, ejecutá aplicarUnificacion().');
  }
}

function aplicarUnificacion() {
  var plan = planUnificacion_();
  mostrarUnificacion_(plan);
  if (plan.errores.length) throw new Error('Unificación cancelada: ' + plan.errores.join(' | '));
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CLIENTES);

  asegurarEncabezadosClientes_(sh);
  plan.renombrar.forEach(function(x) {
    sh.getRange(x.fila, 1).setValue(x.nuevo);
  });
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
  var plan = { renombrar: [], unir: [], sinPuntos: [], crearTiendaNube: false, titulos: false, errores: [] };
  plan.titulos = !String(data[0][8] || '').trim() || !String(data[0][9] || '').trim();

  // Primero los homónimos: con el nombre nuevo ya no se agrupan
  for (var r = 1; r < data.length; r++) {
    CLIENTES_A_RENOMBRAR.forEach(function(x) {
      var tel = soloDigitos_(data[r][2]);
      if (claveNombre_(data[r][0]) === x.nombre && tel && tel.slice(-x.telefono.length) === x.telefono) {
        plan.renombrar.push({ fila: r + 1, antes: data[r][0], nuevo: x.nuevo });
        data[r][0] = x.nuevo;
      }
    });
  }

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

    // Si dos filas tienen teléfono, email o DNI distintos, son dos personas: no se unen
    var conflicto = filas.length > 1 && COLS_IDENTIDAD.some(function(c) {
      var vals = {};
      filas.forEach(function(f) {
        var v = c === 3 ? String(data[f][c] || '').trim().toLowerCase() : soloDigitos_(data[f][c]);
        if (v) vals[v] = true;
      });
      return Object.keys(vals).length > 1;
    });
    if (conflicto) {
      plan.errores.push('"' + data[filas[0]][0] + '" está ' + filas.length + ' veces (filas ' +
        filas.map(function(f) { return f + 1; }).join(', ') + ') con teléfono, email o DNI distintos. ' +
        'Si son personas distintas, cambiale el nombre a una; si es la misma, borrá el dato que sobra.');
    } else if (filas.length > 1) {
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
  if (!plan.renombrar.length && !plan.unir.length && !plan.sinPuntos.length && !plan.crearTiendaNube && !plan.titulos && !plan.errores.length) {
    console.log('Nada para hacer: no hay clientes repetidos y las cuentas sin puntos ya están marcadas.');
    return;
  }
  console.log('HOMÓNIMOS a renombrar: ' + plan.renombrar.length);
  plan.renombrar.forEach(function(x) { console.log('  fila ' + x.fila + ': "' + x.antes + '" → "' + x.nuevo + '"'); });
  console.log('CLIENTES REPETIDOS a unir: ' + plan.unir.length);
  plan.unir.forEach(function(u) {
    console.log('  ' + u.nombres.join(' + ') + '  →  queda "' + u.nombre + '" (fila ' + u.fila + '), se borran las filas ' + u.borrar.join(', ') +
                (u.datos[9] === 'no' ? '  [no suma puntos]' : ''));
  });
  console.log('NO SUMAN PUNTOS:');
  plan.sinPuntos.forEach(function(c) { console.log('  ' + c.nombre + ' (fila ' + c.fila + ')'); });
  if (plan.crearTiendaNube) console.log('  Tienda Nube (cliente nuevo)');
  if (plan.titulos) console.log('TÍTULOS: se agregan "DNI" (col I) y "SumaPuntos" (col J)');
  plan.errores.forEach(function(e) { console.log('⛔ ' + e); });
}
