/*****************************************************************
 * TANQUES CRIOGÉNICOS — CÁLCULOS
 * Funciones puras (sin servicios de Google): clasifican cada lectura,
 * calculan consumo, promedios y arman el reporte del bot.
 * Se pueden probar fuera de Apps Script.
 *****************************************************************/

/** Reglas de clasificación (en pulgadas de agua). */
const REGLAS = {
  recargaMin: 3,      // subir 3 pulg o más = recarga
  mantuvoMax: 1,      // cambiar 1 pulg o menos = se mantuvo
  margenCapacidad: 8, // se acepta hasta capacidad + 8 (la aguja puede pasar un poco)
  diasPromedio: 7,
  diasPronto: 7,      // faltan 7 días o menos para el nivel de recarga = "Recarga pronto"
  minutosRepetida: 30 // otra foto del mismo tanque antes de 30 min reemplaza a la anterior
};

const MESES_CORTOS = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];

/** Número con coma decimal y sin decimales innecesarios: 1.5 -> "1,5", 20 -> "20". */
function num(n, decimales) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const d = decimales === undefined ? 1 : decimales;
  const f = Math.pow(10, d);
  const r = Math.round(n * f) / f;
  return String(r).replace('.', ',');
}

/** Con signo: -20 -> "−20", 78 -> "+78". */
function conSigno(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const r = Math.round(n * 10) / 10;
  if (r > 0) return '+' + num(r);
  if (r < 0) return '−' + num(-r);
  return '0';
}

/** "4 nov, 7:54 a. m." (la fecha ya viene en hora de Colombia como partes). */
function fechaCorta(p) {
  return p.dia + ' ' + MESES_CORTOS[p.mes - 1] + ', ' + horaAmPm(p.hora, p.min);
}

function horaAmPm(h, m) {
  const sufijo = h < 12 ? 'a. m.' : 'p. m.';
  const h12 = h % 12 === 0 ? 12 : h % 12;
  return h12 + ':' + (m < 10 ? '0' : '') + m + ' ' + sufijo;
}

/**
 * Revisa si una lectura de la IA es confiable.
 * ia = { es_manometro, lectura, confianza, motivo }
 * Devuelve { valida, motivo }.
 */
function validarLectura(ia, sede) {
  if (!ia || ia.es_manometro === false) return { valida: false, motivo: 'La foto no parece mostrar el manómetro de nivel.' };
  const v = Number(ia.lectura);
  if (ia.lectura === null || ia.lectura === undefined || ia.lectura === '' || isNaN(v)) {
    return { valida: false, motivo: ia.motivo || 'No se pudo ver la aguja.' };
  }
  if (v < 0 || v > sede.escala) return { valida: false, motivo: 'La lectura (' + num(v) + ') está fuera de la escala del manómetro.' };
  if (v > sede.capacidad + REGLAS.margenCapacidad) {
    return { valida: false, motivo: 'La lectura (' + num(v) + ') pasa la capacidad del tanque (' + sede.capacidad + ').' };
  }
  if (String(ia.confianza || '').toLowerCase() === 'baja') {
    return { valida: false, motivo: ia.motivo || 'La aguja no se ve con claridad.' };
  }
  return { valida: true, motivo: '' };
}

/** Porcentaje del tanque lleno, sin pasar de 100. */
function porcentaje(nivel, capacidad) {
  if (nivel === null || isNaN(nivel) || !capacidad) return null;
  return Math.min(100, Math.round(nivel / capacidad * 100));
}

/** Clasifica el cambio entre dos lecturas válidas. */
function movimiento(cambio) {
  if (cambio === null || isNaN(cambio)) return 'Primera lectura';
  if (cambio >= REGLAS.recargaMin) return 'Recarga';
  if (Math.abs(cambio) <= REGLAS.mantuvoMax) return 'Se mantuvo';
  if (cambio < 0) return 'Consumo';
  return 'Se mantuvo'; // subidas pequeñas (1 a 3 pulg) se tratan como ruido de lectura
}

/**
 * Promedio de consumo diario (pulg/día, número positivo) en los últimos N días.
 * lecturas: válidas de UNA sede, ordenadas de la más vieja a la más nueva,
 * cada una { t (ms), nivel }. Se ignoran los tramos con recarga.
 */
function promedioConsumo(lecturas, ahoraMs, dias) {
  const desde = ahoraMs - dias * 86400000;
  let bajada = 0, horas = 0;
  for (let i = 1; i < lecturas.length; i++) {
    const a = lecturas[i - 1], b = lecturas[i];
    if (b.t < desde || b.t > ahoraMs) continue;
    const inicio = Math.max(a.t, desde);
    const h = (b.t - inicio) / 3600000;
    if (h <= 0) continue;
    const cambio = b.nivel - a.nivel;
    if (cambio >= REGLAS.recargaMin) continue; // tramo con recarga: no cuenta
    const fraccion = (b.t - inicio) / (b.t - a.t);
    bajada += Math.max(0, -cambio) * fraccion;
    horas += h;
  }
  if (horas < 12) return null; // muy poca información
  return bajada / horas * 24;
}

/**
 * Calcula todo lo de una lectura nueva.
 * nueva: { t, nivel } (nivel null si no fue válida)
 * previas: lecturas válidas de la sede (orden cronológico) ANTES de la nueva.
 * recargas: [{ t, cambio, desde, hasta }] de la sede, en orden.
 */
function analizar(nueva, previas, recargas, sede) {
  const ant = previas.length ? previas[previas.length - 1] : null;
  // La aguja puede pasar la capacidad real del tanque: se toma la capacidad.
  const aguja = nueva.nivel;
  if (nueva.nivel !== null && nueva.nivel > sede.capacidad) nueva = { t: nueva.t, nivel: sede.capacidad };
  const r = {
    nivel: nueva.nivel,
    aguja: aguja,
    pct: porcentaje(nueva.nivel, sede.capacidad),
    anterior: ant,
    cambio: null, horas: null, porDia: null,
    movimiento: 'Primera lectura',
    promedio: null, diasHastaRecarga: null,
    bajoNivel: nueva.nivel !== null && nueva.nivel <= sede.aviso,
    ultimaRecarga: recargas.length ? recargas[recargas.length - 1] : null
  };
  if (nueva.nivel === null) { r.movimiento = 'Sin lectura'; return r; }
  if (ant) {
    r.cambio = Math.round((nueva.nivel - ant.nivel) * 10) / 10;
    r.horas = (nueva.t - ant.t) / 3600000;
    r.movimiento = movimiento(r.cambio);
    if (r.movimiento !== 'Recarga' && r.horas > 0) r.porDia = r.cambio / r.horas * 24;
  }
  const serie = previas.concat([{ t: nueva.t, nivel: nueva.nivel }]);
  r.promedio = promedioConsumo(serie, nueva.t, REGLAS.diasPromedio);
  if (r.movimiento === 'Recarga') {
    r.ultimaRecarga = { t: nueva.t, cambio: r.cambio, desde: ant.nivel, hasta: nueva.nivel };
    // El promedio de antes de la recarga es el que sirve para estimar.
    r.promedio = promedioConsumo(previas, ant.t, REGLAS.diasPromedio);
  }
  if (r.promedio && r.promedio > 0.2 && nueva.nivel > sede.aviso) {
    r.diasHastaRecarga = (nueva.nivel - sede.aviso) / r.promedio;
  }
  return r;
}

/**
 * Arma el reporte (HTML de Telegram) que el bot publica debajo de la foto.
 * partes(ms) -> { dia, mes, hora, min } en hora de Colombia.
 */
function reporte(sede, a, tMs, partes, motivoDudosa, lecturaAprox) {
  const L = [];
  L.push('<b>Tanque ' + sede.nombre + '</b> · ' + fechaCorta(partes(tMs)));
  if (motivoDudosa) {
    L.push('No pude leer bien el nivel en esta foto.');
    L.push('<i>' + escaparHtml(motivoDudosa) + '</i>');
    if (lecturaAprox !== null && lecturaAprox !== undefined && !isNaN(lecturaAprox)) {
      L.push('Lectura aproximada: ' + num(lecturaAprox) + ' pulg (no se usa en los cálculos).');
    }
    L.push('Quedó como <b>dudosa</b> en la hoja para revisarla.');
    return L.join('\n');
  }
  L.push('Nivel: <b>' + num(a.nivel) + ' pulg</b> (' + a.pct + ' % de ' + sede.capacidad + ')' +
    (a.aguja !== null && a.aguja > a.nivel ? ' · la aguja marcaba ' + num(a.aguja) : ''));
  if (a.movimiento === 'Recarga') {
    L.push('<b>Recarga detectada: ' + conSigno(a.cambio) + ' pulg</b> (' + num(a.anterior.nivel) + ' → ' + num(a.nivel) + ')');
    if (a.promedio !== null) L.push('Promedio 7 días antes de recargar: −' + num(a.promedio) + ' pulg/día');
  } else if (a.anterior) {
    const desde = fechaCorta(partes(a.anterior.t));
    const lapso = a.horas < 36 ? num(a.horas, 0) + ' h' : num(a.horas / 24) + ' días';
    if (a.movimiento === 'Consumo') {
      L.push('Consumo desde ' + desde + ': <b>' + conSigno(a.cambio) + ' pulg</b> en ' + lapso);
    } else {
      L.push('Se mantuvo desde ' + desde + ' (' + conSigno(a.cambio) + ' pulg en ' + lapso + ')');
    }
    if (a.promedio !== null) L.push('Promedio 7 días: −' + num(a.promedio) + ' pulg/día');
    if (a.ultimaRecarga) {
      L.push('Última recarga: ' + partes(a.ultimaRecarga.t).dia + ' ' + MESES_CORTOS[partes(a.ultimaRecarga.t).mes - 1] +
        ' (' + conSigno(a.ultimaRecarga.cambio) + ' pulg)');
    }
  } else {
    L.push('Primera lectura registrada para este tanque.');
  }
  if (a.bajoNivel) {
    L.push('⚠️ <b>En el nivel de recarga o por debajo (' + sede.aviso + ' pulg).</b> Necesita recarga.');
  } else if (a.diasHastaRecarga !== null && a.diasHastaRecarga <= REGLAS.diasPronto) {
    L.push('🟠 <b>Recarga pronto:</b> al ritmo promedio llega a ' + sede.aviso + ' pulg en ~' + num(a.diasHastaRecarga) + ' días.');
  } else if (a.diasHastaRecarga !== null) {
    L.push('Al ritmo promedio llega a ' + sede.aviso + ' pulg en ~' + num(a.diasHastaRecarga) + ' días.');
  }
  return L.join('\n');
}

function escaparHtml(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Para pruebas fuera de Apps Script.
if (typeof module !== 'undefined') {
  module.exports = { REGLAS, num, conSigno, fechaCorta, validarLectura, porcentaje, movimiento, promedioConsumo, analizar, reporte };
}
