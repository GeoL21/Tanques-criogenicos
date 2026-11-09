/*****************************************************************
 * COSTOS DE LAS RECARGAS (Messer cobra por metro cúbico)
 *
 * - Los m³ de cada recarga salen de la tabla de conversión del tanque:
 *   m³(pulgadas finales) − m³(pulgadas iniciales).
 * - El precio por m³ va por año en la hoja "Precios y presupuesto"
 *   (sin IVA, incluye transporte). Ahí también va el presupuesto mensual.
 * - El administrador escribe desde el tablero, con su huella, el número
 *   de factura y los m³ que Messer facturó; así se ve la diferencia.
 *****************************************************************/

const HOJA_PRE = 'Precios y presupuesto';
const ENC_PRE = ['Año', 'Precio por m³', 'Presupuesto mensual Barranquilla', 'Presupuesto mensual Baranoa'];
const PRECIO_INICIAL = { anio: 2026, precio: 2787 };

/**
 * Tablas de conversión de los tanques (pulgadas de H₂O → m³).
 * La posición es la pulgada: TABLA_M3.Baranoa[44] = 886 m³. La posición 0 vale 0.
 */
const TABLA_M3 = {
  Baranoa: [0, 3, 8, 12, 27, 38, 49, 65, 81, 103, 124, 144, 163, 183, 202, 221, 264, 286, 307, 329, 350,
    372, 393, 415, 436, 458, 479, 501, 522, 543, 564, 586, 607, 629, 650, 672, 693, 736, 758, 779, 801,
    822, 843, 864, 886, 907, 929, 950, 972, 993, 1015, 1036, 1058, 1079, 1101, 1122, 1143, 1164, 1207, 1229, 1250,
    1272, 1293, 1315, 1336, 1358, 1379, 1387, 1396, 1404, 1423, 1443, 1462, 1481, 1503, 1525, 1541, 1556, 1578, 1586, 1594,
    1598, 1602, 1605],
  Barranquilla: [0, 7, 7, 28, 49, 70, 92, 113, 134, 155, 186, 204, 235, 266, 287, 337, 374, 411, 466, 504, 525,
    564, 608, 643, 682, 712, 761, 817, 839, 879, 921, 957, 996, 1025, 1075, 1130, 1154, 1193, 1234, 1271, 1311,
    1338, 1389, 1443, 1468, 1507, 1547, 1586, 1625, 1651, 1704, 1756, 1782, 1822, 1860, 1900, 1940, 1964, 2018, 2069, 2097,
    2136, 2173, 2215, 2254, 2277, 2332, 2382, 2411, 2450, 2486, 2529, 2568, 2590, 2647, 2695, 2725, 2765, 2799, 2843, 2883,
    2903, 2961, 3008, 3040, 3079, 3112, 3158, 3217, 3236, 3276, 3321, 3354, 3393, 3425, 3472, 3530, 3551, 3590, 3634, 3669,
    3708, 3738, 3786, 3843, 3865, 3904, 3947, 3983, 4022, 4051, 4101, 4156, 4179, 4219, 4260, 4260, 4261, 4262, 4308, 4345,
    4382, 4419, 4441, 4485, 4517, 4541, 4578, 4614, 4634, 4666, 4679, 4701, 4709, 4721, 4724, 4728]
};

/** m³ que hay en el tanque con ese nivel (si el nivel tiene decimales, se reparte entre las dos pulgadas). */
function m3DeNivel(sede, pulg) {
  const t = TABLA_M3[sede];
  if (!t || pulg === '' || pulg === null || isNaN(pulg)) return null;
  const p = Math.max(0, Math.min(t.length - 1, Number(pulg)));
  const a = Math.floor(p), b = Math.min(t.length - 1, a + 1);
  return t[a] + (t[b] - t[a]) * (p - a);
}

/** m³ cargados en una recarga según la tabla. */
function m3Recarga(sede, inicial, final) {
  const a = m3DeNivel(sede, inicial), b = m3DeNivel(sede, final);
  if (a === null || b === null || b <= a) return null;
  return Math.round(b - a);
}

/** Hoja de precios: la crea con el precio de 2026 si no existe. */
function hojaPrecios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HOJA_PRE);
  if (!h) {
    h = asegurarHoja(ss, HOJA_PRE, ENC_PRE);
    h.getRange(2, 1, 1, 2).setValues([[PRECIO_INICIAL.anio, PRECIO_INICIAL.precio]]);
    try {
      h.getRange(1, 1, 1, ENC_PRE.length).setFontWeight('bold').setBackground('#6A1B3A').setFontColor('#FFFFFF');
      h.getRange('B:D').setNumberFormat('$ #,##0');
      h.getRange('A:A').setNumberFormat('0');
      h.setColumnWidths(1, ENC_PRE.length, 230);
    } catch (e) {}
  }
  return h;
}

/** Precios y presupuestos por año: [{anio, precio, pres: {Barranquilla, Baranoa}}]. */
function leerPrecios() {
  const h = hojaPrecios();
  if (h.getLastRow() < 2) return [];
  const v = h.getRange(1, 1, h.getLastRow(), ENC_PRE.length).getValues();
  const ix = indices(v[0]);
  return v.slice(1).filter(function(f) { return Number(f[ix['Año']]) > 2000; }).map(function(f, i) {
    const n = function(x) { return x === '' || x === null || isNaN(x) ? null : Number(x); };
    return { fila: i + 2, anio: Number(f[ix['Año']]), precio: n(f[ix['Precio por m³']]),
      pres: { Barranquilla: n(f[ix['Presupuesto mensual Barranquilla']]), Baranoa: n(f[ix['Presupuesto mensual Baranoa']]) } };
  }).sort(function(a, b) { return a.anio - b.anio; });
}

/** Precio del año de la recarga (si ese año no tiene precio, el del último año anterior que lo tenga). */
function precioDe(precios, tMs) {
  const anio = Number(Utilities.formatDate(new Date(tMs), Session.getScriptTimeZone(), 'yyyy'));
  let p = null;
  precios.forEach(function(x) { if (x.anio <= anio && x.precio) p = x.precio; });
  return p;
}

/** Lee las columnas de costos de la tabla Recargas (por número de fila). */
function costosPorFila() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_REC);
  const out = {};
  if (!h || h.getLastRow() < 2) return out;
  const v = h.getRange(1, 1, h.getLastRow(), h.getLastColumn()).getValues();
  const ix = indices(v[0]);
  const val = function(f, n) { return ix[n] === undefined ? '' : f[ix[n]]; };
  for (let r = 1; r < v.length; r++) {
    out[r + 1] = { m3Tabla: val(v[r], 'm³ por tabla'), precio: val(v[r], 'Precio por m³'), costoEst: val(v[r], 'Costo estimado'),
      factura: String(val(v[r], 'N.º factura') || ''), m3Fact: val(v[r], 'm³ facturados'), costoFact: val(v[r], 'Costo facturado') };
  }
  return out;
}

/**
 * Llena o corrige los costos de todas las recargas (m³ por tabla, precio, costo estimado y facturado).
 * Se ejecuta sola cada 5 minutos con procesarTelegram; solo escribe lo que cambió.
 */
function sincronizarCostos() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_REC);
  if (!h || h.getLastRow() < 2) return 0;
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(function(t) { return String(t).trim(); });
  if (enc.indexOf('Costo estimado') < 0) return 0; // aún no están las columnas
  const precios = leerPrecios(), costos = costosPorFila();
  let cambios = 0;
  leerRecargas().forEach(function(r) {
    const c = costos[r.fila] || {};
    const m3 = m3Recarga(r.sede, r.inicial, r.final);
    const precio = precioDe(precios, r.t);
    const nuevo = {
      'm³ por tabla': m3 === null ? '' : m3,
      'Precio por m³': precio === null ? '' : precio,
      'Costo estimado': m3 !== null && precio ? Math.round(m3 * precio) : '',
      'Costo facturado': c.m3Fact !== '' && c.m3Fact !== null && !isNaN(c.m3Fact) && precio ? Math.round(Number(c.m3Fact) * precio) : ''
    };
    const actual = { 'm³ por tabla': c.m3Tabla, 'Precio por m³': c.precio, 'Costo estimado': c.costoEst, 'Costo facturado': c.costoFact };
    Object.keys(nuevo).forEach(function(n) {
      if (String(nuevo[n]) !== String(actual[n] === undefined ? '' : actual[n])) {
        ponerEnRecarga(r.fila, n, nuevo[n]); cambios++;
      }
    });
  });
  return cambios;
}

/** Recargas con costos para el tablero (último año y medio). */
function costosTablero() {
  const precios = leerPrecios(), costos = costosPorFila();
  const desde = Date.now() - 550 * 86400000;
  const numO = function(x) { return x === '' || x === null || isNaN(x) ? null : Number(x); };
  return {
    precios: precios.map(function(p) { return { anio: p.anio, precio: p.precio, pres: p.pres }; }),
    recargas: leerRecargas().filter(function(r) { return r.t >= desde; }).map(function(r) {
      const c = costos[r.fila] || {};
      const m3 = m3Recarga(r.sede, r.inicial, r.final), precio = precioDe(precios, r.t), m3Fact = numO(c.m3Fact);
      return { id: r.id, t: r.t, sede: r.sede, inicial: numO(r.inicial), final: numO(r.final), remision: r.remision,
        m3: m3, precio: precio, costoEst: m3 !== null && precio ? Math.round(m3 * precio) : null,
        factura: c.factura || '', m3Fact: m3Fact, costoFact: m3Fact !== null && precio ? Math.round(m3Fact * precio) : null };
    })
  };
}

/** El administrador escribe la factura de una recarga (con su huella). */
function guardarFactura(d) {
  const id = String(d.id || '');
  const v = verificar_(d);
  if (v.para !== 'factura:' + id) throw new Error('La huella se pidió para otra recarga.');
  const factura = String(d.factura || '').trim().slice(0, 40);
  const m3 = d.m3 === '' || d.m3 === null || d.m3 === undefined ? '' : Number(String(d.m3).replace(/\./g, '').replace(',', '.'));
  if (m3 !== '' && (isNaN(m3) || m3 <= 0 || m3 > 10000)) throw new Error('Los m³ facturados no son válidos.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (PropertiesService.getScriptProperties().getProperty('COLUMNAS_V') !== VERSION_COLUMNAS) agregarColumnasRecargas();
    const r = leerRecargas().filter(function(x) { return x.id === id; })[0];
    if (!r) throw new Error('No encontré esa recarga.');
    ponerEnRecarga(r.fila, 'N.º factura', factura);
    ponerEnRecarga(r.fila, 'm³ facturados', m3);
    sincronizarCostos();
    return { ok: true, nombre: v.nombre };
  } finally {
    lock.releaseLock();
  }
}

/** El administrador cambia el precio por m³ o el presupuesto de un año (con su huella). */
function guardarPrecios(d) {
  const v = verificar_(d);
  if (v.para !== 'precios') throw new Error('La huella se pidió para otra cosa.');
  const anio = Number(d.anio);
  if (!(anio >= 2020 && anio <= 2100)) throw new Error('El año no es válido.');
  const n = function(x, nombre) {
    if (x === '' || x === null || x === undefined) return '';
    const y = Number(String(x).replace(/[$\s.]/g, '').replace(',', '.'));
    if (isNaN(y) || y < 0) throw new Error(nombre + ' no es válido.');
    return y;
  };
  const precio = n(d.precio, 'El precio'), pBq = n(d.presBarranquilla, 'El presupuesto de Barranquilla'), pBn = n(d.presBaranoa, 'El presupuesto de Baranoa');
  if (precio === '' || precio < 100 || precio > 100000) throw new Error('Escribe el precio por m³ (por ejemplo 2787).');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const h = hojaPrecios(), x = leerPrecios().filter(function(p) { return p.anio === anio; })[0];
    const fila = [anio, precio, pBq, pBn];
    if (x) h.getRange(x.fila, 1, 1, 4).setValues([fila]); else h.appendRow(fila);
    sincronizarCostos();
    return { ok: true, nombre: v.nombre };
  } finally {
    lock.releaseLock();
  }
}

/** Para probar desde el editor: muestra los m³ y el costo de las recargas. */
function probarCostos() {
  Logger.log('Baranoa 44 → 76: ' + m3Recarga('Baranoa', 44, 76) + ' m³ (debe dar 655)');
  Logger.log('Barranquilla 88 → 136: ' + m3Recarga('Barranquilla', 88, 136) + ' m³ (debe dar 1511)');
  Logger.log('Precios: ' + JSON.stringify(leerPrecios()));
  Logger.log('Celdas actualizadas: ' + sincronizarCostos());
  costosTablero().recargas.slice(-10).forEach(function(r) {
    Logger.log(r.sede + ' ' + new Date(r.t).toLocaleDateString() + ': ' + r.inicial + ' → ' + r.final + ' = ' + r.m3 + ' m³ · $' + r.costoEst);
  });
}
