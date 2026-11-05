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
  diasPronto: 7,
  horasRitmo: 6,      // ventana para el "ritmo reciente" (consumo de las últimas horas)
  ritmoAlto: 1.5,     // si el ritmo reciente supera 1,5 veces el promedio, se usa para estimar      // faltan 7 días o menos para el nivel de recarga = "Recarga pronto"
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
 * Ritmo reciente (pulg/día): pendiente de la recta que mejor pasa por las lecturas de las
 * últimas horas, sin cruzar una recarga. Con lecturas cada hora resiste el error de ±1 pulg
 * de cada foto. null si hay menos de 3 lecturas o menos de 2 horas de datos.
 */
function ritmoReciente(lecturas, ahoraMs, horas) {
  const desde = ahoraMs - (horas || REGLAS.horasRitmo) * 3600000;
  let pts = lecturas.filter(function(l) { return l.t >= desde && l.t <= ahoraMs; });
  for (let i = pts.length - 1; i > 0; i--) {
    if (pts[i].nivel - pts[i - 1].nivel >= REGLAS.recargaMin) { pts = pts.slice(i); break; }
  }
  if (pts.length < 3 || (pts[pts.length - 1].t - pts[0].t) < 2 * 3600000) return null;
  const n = pts.length;
  const mx = pts.reduce(function(a, p) { return a + p.t / 3600000; }, 0) / n;
  const my = pts.reduce(function(a, p) { return a + p.nivel; }, 0) / n;
  let sxy = 0, sxx = 0;
  pts.forEach(function(p) { const dx = p.t / 3600000 - mx; sxy += dx * (p.nivel - my); sxx += dx * dx; });
  if (!sxx) return null;
  return Math.max(0, -(sxy / sxx) * 24);
}

/** Decide con qué ritmo se estima: el promedio de 7 días, o el reciente si es mucho más alto. */
function ritmoParaEstimar(promedio, reciente) {
  const alto = reciente !== null && reciente > 0.5 && (promedio === null || reciente > promedio * REGLAS.ritmoAlto);
  return { ritmo: alto ? reciente : promedio, alto: alto };
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
  r.reciente = r.movimiento === 'Recarga' ? null : ritmoReciente(serie, nueva.t);
  const rt = ritmoParaEstimar(r.promedio, r.reciente);
  r.ritmoAlto = rt.alto;
  // Con consumo alto y fotos seguidas, una bajada pequeña también es consumo (no "se mantuvo")
  if (r.ritmoAlto && r.movimiento === 'Se mantuvo' && r.cambio < 0) r.movimiento = 'Consumo';
  if (rt.ritmo && rt.ritmo > 0.2 && nueva.nivel > sede.aviso) {
    r.diasHastaRecarga = (nueva.nivel - sede.aviso) / rt.ritmo;
  }
  return r;
}

/**
 * Arma el reporte (HTML de Telegram) que el bot publica debajo de la foto.
 * partes(ms) -> { dia, mes, hora, min } en hora de Colombia.
 */
function reporte(sede, a, tMs, partes, motivoDudosa, lecturaAprox, revisar) {
  const L = [];
  L.push('<b>Tanque ' + sede.nombre + '</b> · ' + fechaCorta(partes(tMs)));
  // El cálculo por ángulo y la estimación de la IA no coinciden: se pide confirmar el valor calculado
  if (revisar) {
    const v = num(Math.min(revisar.calc, sede.capacidad), 0);
    L.push('⚠️ <b>Revisa esta lectura:</b> calculé <b>' + num(revisar.calc) + ' pulg</b>, pero la IA estimaba ' + num(revisar.estim) + '.');
    L.push('Mira el manómetro: si marca ' + v + ', responde a este mensaje con <code>/corregir ' + v + '</code>; si no, con el valor que ves.');
    L.push('Mientras tanto queda como <b>dudosa</b> y no cuenta en los cálculos.');
    return L.join('\n');
  }
  if (motivoDudosa) {
    L.push('⚠️ No pude leer bien el nivel en esta foto.');
    L.push('<i>' + escaparHtml(motivoDudosa) + '</i>');
    const hay = lecturaAprox !== null && lecturaAprox !== undefined && !isNaN(lecturaAprox);
    if (hay) L.push('Lectura aproximada: ' + num(lecturaAprox) + ' pulg (no se usa en los cálculos).');
    L.push('Quedó como <b>dudosa</b>. Si ves el valor en el manómetro, responde a este mensaje con ' +
      (hay ? '<code>/corregir ' + num(Math.min(lecturaAprox, sede.capacidad), 0) + '</code> (o el valor que veas).'
        : '<code>/corregir</code> y el número, por ejemplo <code>/corregir 41</code>.'));
    return L.join('\n');
  }
  L.push('Nivel: <b>' + num(a.nivel) + ' pulg</b> (' + a.pct + ' % de ' + sede.capacidad + ')' +
    (a.aguja !== null && a.aguja > a.nivel ? ' · la aguja marcaba ' + num(a.aguja) : ''));
  if (a.movimiento === 'Recarga') {
    L.push('<b>Recarga detectada: ' + conSigno(a.cambio) + ' pulg</b> (' + num(a.anterior.nivel) + ' → ' + num(a.nivel) + ')');
    if (a.promedio !== null) L.push('Promedio 7 días antes de recargar: ' + promTxt(a.promedio));
  } else if (a.anterior) {
    const desde = fechaCorta(partes(a.anterior.t));
    const lapso = a.horas < 36 ? num(a.horas, 0) + ' h' : num(a.horas / 24) + ' días';
    if (a.movimiento === 'Consumo') {
      L.push('Consumo desde ' + desde + ': <b>' + conSigno(a.cambio) + ' pulg</b> en ' + lapso);
    } else {
      L.push('Se mantuvo desde ' + desde + ' (' + conSigno(a.cambio) + ' pulg en ' + lapso + ')');
    }
    if (a.ritmoAlto) {
      L.push('🔺 <b>Consumo alto:</b> en las últimas horas va a ' + promTxt(a.reciente) +
        (a.reciente >= 12 ? ' (~' + num(a.reciente / 24) + ' pulg por hora)' : '') +
        (a.promedio !== null ? '; el promedio de 7 días es ' + promTxt(a.promedio) : '') + '.');
    } else if (a.promedio !== null) L.push('Promedio 7 días: ' + promTxt(a.promedio));
    if (a.ultimaRecarga) {
      L.push('Última recarga: ' + partes(a.ultimaRecarga.t).dia + ' ' + MESES_CORTOS[partes(a.ultimaRecarga.t).mes - 1] +
        ' (' + conSigno(a.ultimaRecarga.cambio) + ' pulg)');
    }
  } else {
    L.push('Primera lectura registrada para este tanque.');
  }
  if (a.bajoNivel) {
    L.push('⚠️ <b>En el nivel de recarga o por debajo (' + sede.aviso + ' pulg).</b> Necesita recarga.');
  } else if (a.diasHastaRecarga !== null) {
    const ritmo = a.ritmoAlto ? 'al ritmo actual' : 'al ritmo promedio';
    const cuando = cuandoLlega(a.diasHastaRecarga, tMs, partes);
    if (a.diasHastaRecarga < 1) {
      L.push('🔴 <b>Recarga urgente:</b> ' + ritmo + ' llega a ' + sede.aviso + ' pulg en ' + cuando + '. Avisar a Messer o a la administración.');
    } else if (a.diasHastaRecarga <= REGLAS.diasPronto) {
      L.push('🟠 <b>Recarga pronto:</b> ' + ritmo + ' llega a ' + sede.aviso + ' pulg en ' + cuando + '.');
    } else {
      L.push(ritmo.charAt(0).toUpperCase() + ritmo.slice(1) + ' llega a ' + sede.aviso + ' pulg en ' + cuando + '.');
    }
  }
  return L.join('\n');
}

/** "~8 h (hacia las 6:40 p. m.)" si falta menos de 2 días; "~5,2 días" si falta más. */
function cuandoLlega(dias, tMs, partes) {
  if (dias >= 2) return '~' + num(dias) + ' días';
  const h = dias * 24;
  const p = partes(tMs + h * 3600000);
  return '~' + num(h, 0) + ' h (hacia ' + (p.dia !== partes(tMs).dia ? 'el ' + p.dia + ' ' + MESES_CORTOS[p.mes - 1] + ' a ' : '') +
    'las ' + horaAmPm(p.hora, p.min) + ')';
}

/** Consumo promedio: "−2,8 pulg/día", o "sin consumo" si es casi cero. */
function promTxt(p) {
  return p < 0.05 ? 'sin consumo' : '−' + num(p) + ' pulg/día';
}

/***********************
 * LECTURA POR GEOMETRÍA
 * La IA solo ubica puntos en la foto (eje, punta de la aguja y marcas con número);
 * el ángulo y las pulgadas se calculan aquí, sin adivinar.
 ***********************/

/** Ancho y alto de una foto JPEG leyendo su encabezado (bytes 0–255 o con signo). */
function tamanoJpeg(bytes) {
  const b = function(i) { return bytes[i] & 255; };
  if (b(0) !== 0xFF || b(1) !== 0xD8) return null;
  let i = 2;
  while (i + 9 < bytes.length) {
    if (b(i) !== 0xFF) { i++; continue; }
    const m = b(i + 1);
    if (m === 0xD8 || m === 0x01 || (m >= 0xD0 && m <= 0xD7)) { i += 2; continue; }
    const largo = (b(i + 2) << 8) + b(i + 3);
    if ((m >= 0xC0 && m <= 0xC3) || (m >= 0xC5 && m <= 0xC7) || (m >= 0xC9 && m <= 0xCB) || (m >= 0xCD && m <= 0xCF)) {
      return { alto: (b(i + 5) << 8) + b(i + 6), ancho: (b(i + 7) << 8) + b(i + 8) };
    }
    i += 2 + largo;
  }
  return null;
}

/** Punto de la IA ([y, x] de 0 a 1000, o {x, y}) → píxeles. */
function aPixeles(p, ancho, alto) {
  if (!p) return null;
  let y, x;
  if (Array.isArray(p) && p.length >= 2) { y = Number(p[0]); x = Number(p[1]); }
  else if (typeof p === 'object') { y = Number(p.y); x = Number(p.x); }
  if (isNaN(x) || isNaN(y)) return null;
  return { x: x / 1000 * ancho, y: y / 1000 * alto };
}

/** Ángulo en grados, creciendo en el sentido de las agujas del reloj (como sube la escala). */
function anguloDesde(c, p) {
  return Math.atan2(p.y - c.y, p.x - c.x) * 180 / Math.PI;
}
function giro(desde, hasta) { // de 0 a 360, en sentido horario
  return ((hasta - desde) % 360 + 360) % 360;
}

/**
 * Calcula la lectura a partir de los puntos que devolvió la IA.
 * ia = { centro:[y,x], punta:[y,x], marcas:[{valor, punto:[y,x]}, ...] }
 * Devuelve { lectura, entre:[a,b], marcasUsadas } o null si no alcanza la información.
 */
function lecturaPorGeometria(ia, ancho, alto, escala) {
  if (!ia || !ancho || !alto) return null;
  const c = aPixeles(ia.centro, ancho, alto), t = aPixeles(ia.punta, ancho, alto);
  if (!c || !t) return null;
  const radioAguja = Math.hypot(t.x - c.x, t.y - c.y);
  if (radioAguja < Math.min(ancho, alto) * 0.03) return null; // punta pegada al eje: puntos mal ubicados
  let marcas = (ia.marcas || []).map(function(m) {
    const p = aPixeles(m.punto || m.p || m.pos, ancho, alto);
    const v = Number(m.valor);
    return p && !isNaN(v) && v >= 0 && v <= escala ? { v: v, a: anguloDesde(c, p) } : null;
  }).filter(Boolean).sort(function(a, b) { return a.v - b.v; });
  // Sin valores repetidos
  marcas = marcas.filter(function(m, i) { return i === 0 || m.v !== marcas[i - 1].v; });
  if (marcas.length < 2) return null;
  // Ángulos medidos desde la primera marca: deben crecer junto con el valor
  const a0 = marcas[0].a;
  marcas.forEach(function(m) { m.g = giro(a0, m.a); });
  marcas[0].g = 0;
  const buenas = [marcas[0]];
  for (let i = 1; i < marcas.length; i++) {
    if (marcas[i].g > buenas[buenas.length - 1].g + 1) buenas.push(marcas[i]);
  }
  if (buenas.length < 2) return null;
  // Grados por unidad de escala parecidos en todos los tramos (una marca mal ubicada se descarta)
  const pend = [];
  for (let i = 1; i < buenas.length; i++) pend.push((buenas[i].g - buenas[i - 1].g) / (buenas[i].v - buenas[i - 1].v));
  const med = pend.slice().sort(function(a, b) { return a - b; })[Math.floor(pend.length / 2)];
  const finales = [buenas[0]];
  for (let i = 1; i < buenas.length; i++) {
    const pd = (buenas[i].g - finales[finales.length - 1].g) / (buenas[i].v - finales[finales.length - 1].v);
    if (pd > med * 0.6 && pd < med * 1.6) finales.push(buenas[i]);
  }
  if (finales.length < 2) return null;
  // La aguja: mismo origen; si queda un poco antes de la primera marca, es negativa (cerca de 0)
  let gn = giro(a0, anguloDesde(c, t));
  if (gn > 330) gn -= 360;
  let i = 1;
  while (i < finales.length - 1 && gn > finales[i].g) i++;
  const A = finales[i - 1], B = finales[i];
  let lectura = A.v + (gn - A.g) * (B.v - A.v) / (B.g - A.g);
  lectura = Math.max(0, Math.min(escala, lectura));
  return { lectura: Math.round(lectura * 10) / 10, entre: [A.v, B.v], marcasUsadas: finales.length };
}

function escaparHtml(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Para pruebas fuera de Apps Script.
if (typeof module !== 'undefined') {
  module.exports = { REGLAS, num, conSigno, fechaCorta, validarLectura, porcentaje, movimiento, promedioConsumo, analizar, reporte, tamanoJpeg, lecturaPorGeometria, ritmoReciente };
}
