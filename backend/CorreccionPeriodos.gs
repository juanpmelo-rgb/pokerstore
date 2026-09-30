// ============================================================
//  CORRECCIÓN ÚNICA DE PERÍODOS (septiembre 2026)
//
//  Los cierres del 1/7, 2/8 y 2/9 se etiquetaron con el mes del día del
//  cierre en vez del mes cerrado, y Sheets guardó el período como fecha.
//  Este script:
//    · pasa esos períodos al mes anterior, como texto (julio → "2026/06", etc.)
//      en Historico (col G) y Arqueos (col B)
//    · borra el arqueo duplicado (0 movimientos, mismo día y período que otro)
//
//  Uso, desde el editor de Apps Script:
//    1. Ejecutar revisarCorreccion() → solo muestra en el registro qué cambiaría.
//    2. Si está bien, ejecutar aplicarCorreccion().
//  Correrlo dos veces no hace nada: después de aplicarlo no quedan períodos
//  guardados como fecha. Una vez aplicado, este archivo se puede borrar.
// ============================================================

function revisarCorreccion() {
  var plan = planCorreccion_();
  mostrarPlan_(plan);
  if (plan.errores.length) {
    console.log('⛔ Hay problemas: NO ejecutes aplicarCorreccion() hasta revisarlos.');
  } else if (plan.cambiosHisto.length || plan.cambiosArq.length || plan.borrarArq.length) {
    console.log('Si esto es correcto, ejecutá aplicarCorreccion().');
  }
}

function aplicarCorreccion() {
  var plan = planCorreccion_();
  mostrarPlan_(plan);
  if (plan.errores.length) throw new Error('Corrección cancelada: ' + plan.errores.join(' | '));

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var histo = ss.getSheetByName(SHEET_HISTORICO);
  var arq = ss.getSheetByName(SHEET_ARQUEOS);

  plan.cambiosHisto.forEach(function(c) {
    histo.getRange(c.fila, 7).setNumberFormat('@').setValue(c.despues);
  });
  plan.cambiosArq.forEach(function(c) {
    arq.getRange(c.fila, 2).setNumberFormat('@').setValue(c.despues);
  });
  // De abajo hacia arriba, para que no se corran los números de fila
  plan.borrarArq.slice().sort(function(a, b) { return b.fila - a.fila; }).forEach(function(c) {
    arq.deleteRow(c.fila);
  });

  console.log('✓ Corrección aplicada.');
}

function planCorreccion_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var histo = ss.getSheetByName(SHEET_HISTORICO).getDataRange().getValues();
  var arq = ss.getSheetByName(SHEET_ARQUEOS).getDataRange().getValues();
  var plan = { cambiosHisto: [], cambiosArq: [], borrarArq: [], grupos: {}, errores: [] };

  // Historico: períodos guardados como fecha → mes anterior
  for (var i = 1; i < histo.length; i++) {
    var p = histo[i][6];
    if (!esFecha_(p)) continue;
    var antes = aaaamm_(p);
    var despues = mesAnterior_(p);
    plan.cambiosHisto.push({ fila: i + 1, antes: antes, despues: despues });
    var g = plan.grupos[antes] || (plan.grupos[antes] = { despues: despues, meses: {} });
    var mesReal = esFecha_(histo[i][0]) ? aaaamm_(histo[i][0]) : '(fecha ilegible)';
    g.meses[mesReal] = (g.meses[mesReal] || 0) + 1;
  }

  // Control: en cada grupo, el mes real más frecuente tiene que ser el nuevo período
  Object.keys(plan.grupos).forEach(function(antes) {
    var g = plan.grupos[antes];
    var top = Object.keys(g.meses).sort(function(a, b) { return g.meses[b] - g.meses[a]; })[0];
    if (top !== g.despues) {
      plan.errores.push('El grupo ' + antes + ' tiene mayoría de movimientos de ' + top + ', no de ' + g.despues);
    }
  });

  // Arqueos: duplicados (0 movimientos, mismo día y período que otro arqueo con movimientos)
  for (var j = 1; j < arq.length; j++) {
    if (Number(arq[j][6]) !== 0) continue;
    for (var k = 1; k < arq.length; k++) {
      if (k !== j && Number(arq[k][6]) > 0 &&
          mismoValor_(arq[k][0], arq[j][0]) && mismoValor_(arq[k][1], arq[j][1])) {
        plan.borrarArq.push({ fila: j + 1, fecha: arq[j][0], periodo: arq[j][1] });
        break;
      }
    }
  }
  var filasABorrar = plan.borrarArq.map(function(c) { return c.fila; });

  // Arqueos: períodos guardados como fecha → mes anterior (salvo los que se borran)
  for (var m = 1; m < arq.length; m++) {
    if (!esFecha_(arq[m][1]) || filasABorrar.indexOf(m + 1) >= 0) continue;
    plan.cambiosArq.push({ fila: m + 1, antes: aaaamm_(arq[m][1]), despues: mesAnterior_(arq[m][1]) });
  }

  return plan;
}

function mostrarPlan_(plan) {
  if (!plan.cambiosHisto.length && !plan.cambiosArq.length && !plan.borrarArq.length) {
    console.log('Nada para corregir: no quedan períodos guardados como fecha ni arqueos duplicados.');
    return;
  }
  console.log('HISTORICO — ' + plan.cambiosHisto.length + ' filas a re-etiquetar:');
  Object.keys(plan.grupos).sort().forEach(function(antes) {
    var g = plan.grupos[antes];
    var n = Object.keys(g.meses).reduce(function(s, k) { return s + g.meses[k]; }, 0);
    console.log('  ' + antes + ' → ' + g.despues + '  (' + n + ' filas; meses reales: ' + JSON.stringify(g.meses) + ')');
  });
  console.log('ARQUEOS — períodos a corregir:');
  plan.cambiosArq.forEach(function(c) { console.log('  fila ' + c.fila + ': ' + c.antes + ' → ' + c.despues); });
  console.log('ARQUEOS — duplicados a borrar:');
  plan.borrarArq.forEach(function(c) { console.log('  fila ' + c.fila + ' (0 movimientos, período ' + aaaamm_(c.periodo) + ')'); });
  plan.errores.forEach(function(e) { console.log('⛔ ' + e); });
}

function esFecha_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}
function aaaamm_(d) {
  if (!esFecha_(d)) return String(d);
  var m = d.getMonth() + 1;
  return d.getFullYear() + '/' + (m < 10 ? '0' + m : m);
}
function mesAnterior_(d) {
  return aaaamm_(new Date(d.getFullYear(), d.getMonth() - 1, 1));
}
function mismoValor_(a, b) {
  if (esFecha_(a) && esFecha_(b)) return a.getTime() === b.getTime();
  return String(a) === String(b);
}
