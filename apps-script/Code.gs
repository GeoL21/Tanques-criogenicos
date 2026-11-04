/*****************************************************************
 * TANQUES CRIOGÉNICOS — CLÍNICA REINA CATALINA
 * Apps Script vinculado a la hoja "Tanques criogénicos" (cuenta cilindrocrc).
 *
 * Cómo funciona:
 *  1. El técnico sube la foto del manómetro de nivel al tema de su sede
 *     en el grupo de Telegram "Gases Medicinales".
 *  2. Cada 5 minutos este script revisa el bot, descarga la foto,
 *     la guarda en Drive y le pide a Gemini (IA de Google) que lea la aguja.
 *  3. Guarda la lectura en la tabla "Lecturas", calcula consumo, recarga
 *     y promedio, actualiza "Resumen" y responde debajo de la foto.
 *  4. La PWA (tablero) consulta los datos por POST con un código de acceso.
 *
 * Propiedades del script (Configuración del proyecto → Propiedades del script):
 *   TELEGRAM_TOKEN  clave del bot que dio @BotFather
 *   GEMINI_KEY      clave de la API de Gemini (aistudio.google.com)
 *   CODIGO_PWA      código de acceso para el tablero
 *   GEMINI_MODELO   (opcional) modelo de Gemini; por defecto gemini-2.5-flash
 *
 * Servicios avanzados que hay que activar: Google Sheets API.
 *****************************************************************/

const HOJA_LEC = 'Lecturas';
const HOJA_RES = 'Resumen';
const HOJA_CFG = 'Config';
const CARPETA_FOTOS = 'Fotos tanques criogénicos';
const MINUTOS_REVISION = 5;
const MODELO_POR_DEFECTO = 'gemini-2.5-flash';
const IA_NO_DISPONIBLE = 'La IA de Google no estuvo disponible';
const MAX_REINTENTOS = 12; // cada 5 minutos: hasta 1 hora intentando

const ENC_LEC = ['ID', 'Fecha y hora', 'Sede', 'Nivel (pulg)', '% lleno', 'Cambio (pulg)', 'Movimiento',
  'Consumo por día (pulg)', 'Horas desde la anterior', 'Lectura', 'Técnico', 'Foto', 'Nota de la lectura',
  'ID del mensaje', 'ID de la foto'];
const ENC_CFG = ['Sede', 'Capacidad (pulg)', 'Escala del manómetro (pulg)', 'Nivel de recarga (pulg)', 'Tema de Telegram (ID)'];
const ENC_RES = ['Sede', 'Nivel actual (pulg)', '% lleno', 'Última lectura', 'Promedio 7 días (pulg/día)',
  'Promedio 30 días (pulg/día)', 'Última recarga', 'Pulgadas recargadas', 'Días hasta el nivel de recarga', 'Estado'];

const MOVIMIENTOS = ['Consumo', 'Se mantuvo', 'Recarga', 'Primera lectura', 'Sin lectura'];
const ESTADOS_LECTURA = ['Automática', 'Dudosa', 'Corregida'];

/** Cómo es cada manómetro, para que la IA lo lea bien. */
const DESCRIPCION_MANOMETRO = {
  barranquilla: 'Manómetro diferencial MID-WEST INSTRUMENT de carátula NEGRA con números blancos de 0 a 200 ' +
    '"INCHES OF WATER" (Δp). Números cada 20 y marcas pequeñas cada 2. La aguja principal es BLANCA y larga. ' +
    'Ignora el dial pequeño de la derecha (increase switch / set point), las pestañas amarillas del borde y los reflejos del vidrio.',
  baranoa: 'Indicador de contenido de carátula BEIGE con números negros de 0 a 100 "Inch H2O". ' +
    'Números cada 20 y marcas cada 2 aproximadamente. La aguja es NEGRA y larga; ignora los tornillos y el contrapeso.'
};


/***********************
 * MENÚ
 ***********************/

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Tanques')
    .addItem('Revisar Telegram ahora', 'procesarTelegram')
    .addItem('Recalcular todo', 'recalcularTodo')
    .addItem('Probar la IA con la última foto', 'probarIA')
    .addItem('Ver modelo de Gemini', 'probarModelo')
    .addSeparator()
    .addItem('Configurar (primera vez)', 'configurar')
    .addToUi();
}


/***********************
 * CONFIGURACIÓN INICIAL (ejecutar una vez)
 ***********************/

function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetLocale('es_CO');
  ss.setSpreadsheetTimeZone('America/Bogota');
  const props = PropertiesService.getScriptProperties();

  const faltan = ['TELEGRAM_TOKEN', 'GEMINI_KEY', 'CODIGO_PWA'].filter(function(k) { return !props.getProperty(k); });

  // Pestañas
  const lec = asegurarHoja(ss, HOJA_LEC, ENC_LEC);
  const cfg = asegurarHoja(ss, HOJA_CFG, ENC_CFG);
  const res = asegurarHoja(ss, HOJA_RES, ENC_RES);

  // Datos de las sedes
  if (!String(cfg.getRange(2, 1).getValue()).trim()) {
    cfg.getRange(2, 1, 2, 4).setValues([
      ['Barranquilla', 136, 200, 40],
      ['Baranoa', 80, 100, 24]
    ]);
  }

  // Tablas (con tipos de columna)
  crearTablas(ss, lec, cfg);

  // Formato del resumen (sin tabla)
  try { res.getRange(1, 1, 1, ENC_RES.length).setFontWeight('bold').setBackground('#6A1B3A').setFontColor('#FFFFFF'); } catch (e) {}

  // Carpeta de fotos
  if (!props.getProperty('CARPETA_ID')) {
    const it = DriveApp.getFoldersByName(CARPETA_FOTOS);
    const carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_FOTOS);
    props.setProperty('CARPETA_ID', carpeta.getId());
  }

  // El bot trabaja revisando (getUpdates): no puede tener webhook.
  if (props.getProperty('TELEGRAM_TOKEN')) telegram('deleteWebhook', { drop_pending_updates: false });

  // Revisión automática
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'procesarTelegram') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('procesarTelegram').timeBased().everyMinutes(MINUTOS_REVISION).create();

  const msg = faltan.length
    ? 'Listo, pero faltan estas propiedades del script: ' + faltan.join(', ') + '. Agrégalas y vuelve a ejecutar Configurar.'
    : 'Listo. El bot revisa Telegram cada ' + MINUTOS_REVISION + ' minutos. Sube una foto de prueba al tema de cada sede.';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function asegurarHoja(ss, nombre, encabezados) {
  let h = ss.getSheetByName(nombre);
  if (!h) h = ss.insertSheet(nombre);
  // Encabezados solo si la fila 1 está vacía (no se usa getLastRow, que falla si la hoja quedó a medias)
  const fila1 = h.getRange(1, 1, 1, encabezados.length).getValues()[0];
  if (fila1.every(function(v) { return v === '' || v === null; })) {
    h.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  }
  try { h.setFrozenRows(1); } catch (e) {}
  return h;
}

/** Convierte Lecturas y Config en tablas de Sheets con tipos de columna. */
function crearTablas(ss, lec, cfg) {
  const id = ss.getId();
  let existentes = [];
  try {
    const info = Sheets.Spreadsheets.get(id, { fields: 'sheets(properties(sheetId,title),tables(tableId,name))' });
    info.sheets.forEach(function(s) { (s.tables || []).forEach(function(t) { existentes.push(t.name); }); });
  } catch (e) {
    Logger.log('Activa el servicio avanzado "Google Sheets API": ' + e.message);
    formatoSinTabla(lec); formatoSinTabla(cfg);
    return;
  }
  const lista = function(vals) {
    return { condition: { type: 'ONE_OF_LIST', values: vals.map(function(v) { return { userEnteredValue: v }; }) } };
  };
  const tipos = {
    'ID': 'TEXT', 'Fecha y hora': 'DATE_TIME', 'Sede': 'DROPDOWN', 'Nivel (pulg)': 'DOUBLE', '% lleno': 'DOUBLE',
    'Cambio (pulg)': 'DOUBLE', 'Movimiento': 'DROPDOWN', 'Consumo por día (pulg)': 'DOUBLE', 'Horas desde la anterior': 'DOUBLE',
    'Lectura': 'DROPDOWN', 'Técnico': 'TEXT', 'Foto': 'FILES_CHIP', 'Nota de la lectura': 'TEXT',
    'ID del mensaje': 'TEXT', 'ID de la foto': 'TEXT'
  };
  const validaciones = { 'Sede': lista(['Barranquilla', 'Baranoa']), 'Movimiento': lista(MOVIMIENTOS), 'Lectura': lista(ESTADOS_LECTURA) };
  const pedidos = [];
  if (existentes.indexOf('Lecturas') < 0) {
    pedidos.push({ addTable: { table: {
      name: 'Lecturas',
      range: { sheetId: lec.getSheetId(), startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: ENC_LEC.length },
      columnProperties: ENC_LEC.map(function(n, i) {
        const c = { columnIndex: i, columnName: n, columnType: tipos[n] };
        if (validaciones[n]) c.dataValidationRule = validaciones[n];
        return c;
      })
    } } });
  }
  if (existentes.indexOf('Config') < 0) {
    pedidos.push({ addTable: { table: {
      name: 'Config',
      range: { sheetId: cfg.getSheetId(), startRowIndex: 0, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: ENC_CFG.length },
      columnProperties: [
        { columnIndex: 0, columnName: ENC_CFG[0], columnType: 'TEXT' },
        { columnIndex: 1, columnName: ENC_CFG[1], columnType: 'DOUBLE' },
        { columnIndex: 2, columnName: ENC_CFG[2], columnType: 'DOUBLE' },
        { columnIndex: 3, columnName: ENC_CFG[3], columnType: 'DOUBLE' },
        { columnIndex: 4, columnName: ENC_CFG[4], columnType: 'TEXT' }
      ]
    } } });
  }
  if (!pedidos.length) return;
  try {
    Sheets.Spreadsheets.batchUpdate({ requests: pedidos }, id);
  } catch (e) {
    Logger.log('No se pudieron crear las tablas, se usa formato simple: ' + e.message);
    formatoSinTabla(lec); formatoSinTabla(cfg);
  }
  try { lec.hideColumns(14, 2); } catch (e) {}
}

/** Da formato si se puede; en columnas de tabla con tipo definido, Sheets no lo permite y no hace falta. */
function formato(rango, patron) {
  try { rango.setNumberFormat(patron); SpreadsheetApp.flush(); } catch (e) { /* la tabla ya define el tipo de la columna */ }
}

function formatoSinTabla(h) {
  h.getRange(1, 1, 1, h.getLastColumn()).setFontWeight('bold').setBackground('#6A1B3A').setFontColor('#FFFFFF');
  if (h.getName() === HOJA_LEC) {
    formato(h.getRange('B:B'), 'dd/mm/yyyy h:mm am/pm');
    formato(h.getRange('A:A'), '@');
    formato(h.getRange('N:O'), '@');
  } else {
    formato(h.getRange('E:E'), '@');
  }
}


/***********************
 * LECTURA DE CONFIG Y DATOS
 ***********************/

function leerSedes() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CFG);
  const v = h.getDataRange().getValues();
  const idx = indices(v[0]);
  return v.slice(1).filter(function(f) { return String(f[idx['Sede']]).trim(); }).map(function(f, i) {
    return {
      fila: i + 2,
      nombre: String(f[idx['Sede']]).trim(),
      capacidad: Number(f[idx['Capacidad (pulg)']]),
      escala: Number(f[idx['Escala del manómetro (pulg)']]),
      aviso: Number(f[idx['Nivel de recarga (pulg)']]),
      tema: String(f[idx['Tema de Telegram (ID)']] || '').trim()
    };
  });
}

function indices(encabezados) {
  const m = {};
  encabezados.forEach(function(t, i) { m[String(t).trim()] = i; });
  return m;
}

function normalizar(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
}

/** Todas las lecturas como objetos, en orden de fecha. */
function leerLecturas() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_LEC);
  if (h.getLastRow() < 2) return [];
  const v = h.getRange(1, 1, h.getLastRow(), h.getLastColumn()).getValues();
  const ix = indices(v[0]);
  const out = [];
  for (let r = 1; r < v.length; r++) {
    const f = v[r];
    const fecha = f[ix['Fecha y hora']];
    if (!(fecha instanceof Date)) continue;
    const nivel = f[ix['Nivel (pulg)']];
    out.push({
      fila: r + 1,
      id: String(f[ix['ID']] || ''),
      t: fecha.getTime(),
      sede: String(f[ix['Sede']] || '').trim(),
      nivel: nivel === '' || nivel === null ? null : Number(nivel),
      pct: f[ix['% lleno']] === '' ? null : Number(f[ix['% lleno']]),
      cambio: f[ix['Cambio (pulg)']] === '' ? null : Number(f[ix['Cambio (pulg)']]),
      movimiento: String(f[ix['Movimiento']] || ''),
      porDia: f[ix['Consumo por día (pulg)']] === '' ? null : Number(f[ix['Consumo por día (pulg)']]),
      estado: String(f[ix['Lectura']] || ''),
      tecnico: String(f[ix['Técnico']] || ''),
      nota: String(f[ix['Nota de la lectura']] || ''),
      msgId: String(f[ix['ID del mensaje']] || ''),
      fotoId: String(f[ix['ID de la foto']] || '')
    });
  }
  out.sort(function(a, b) { return a.t - b.t; });
  return out;
}

function esValida(l) {
  return (l.estado === 'Automática' || l.estado === 'Corregida') && l.nivel !== null && !isNaN(l.nivel);
}

function validasDe(lecturas, sede, antesDe) {
  return lecturas.filter(function(l) {
    return normalizar(l.sede) === normalizar(sede) && esValida(l) && (antesDe === undefined || l.t < antesDe);
  }).map(function(l) { return { t: l.t, nivel: l.nivel }; });
}

/** Recargas a partir de la serie de lecturas válidas (no depende de la columna Movimiento). */
function recargasDe(serie) {
  const out = [];
  for (let i = 1; i < serie.length; i++) {
    const c = serie[i].nivel - serie[i - 1].nivel;
    if (c >= REGLAS.recargaMin) out.push({ t: serie[i].t, cambio: Math.round(c * 10) / 10, desde: serie[i - 1].nivel, hasta: serie[i].nivel });
  }
  return out;
}

function partesFecha(ms) {
  const tz = Session.getScriptTimeZone();
  const d = new Date(ms);
  return {
    dia: Number(Utilities.formatDate(d, tz, 'd')),
    mes: Number(Utilities.formatDate(d, tz, 'M')),
    hora: Number(Utilities.formatDate(d, tz, 'H')),
    min: Number(Utilities.formatDate(d, tz, 'm'))
  };
}


/***********************
 * TELEGRAM
 ***********************/

function telegram(metodo, datos) {
  const token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + metodo, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(datos || {}), muteHttpExceptions: true
  });
  const j = JSON.parse(r.getContentText());
  if (!j.ok) Logger.log('Telegram ' + metodo + ': ' + r.getContentText());
  return j;
}

function responderEnTema(msg, html) {
  const d = {
    chat_id: msg.chat.id, text: html, parse_mode: 'HTML', disable_web_page_preview: true,
    reply_parameters: { message_id: msg.message_id, allow_sending_without_reply: true }
  };
  if (msg.is_topic_message && msg.message_thread_id) d.message_thread_id = msg.message_thread_id;
  return telegram('sendMessage', d);
}

/** Se ejecuta cada 5 minutos: trae lo nuevo del grupo y lo procesa. */
function procesarTelegram() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const props = PropertiesService.getScriptProperties();
    let offset = Number(props.getProperty('OFFSET') || 0);
    const r = telegram('getUpdates', { offset: offset, timeout: 0, allowed_updates: ['message'] });
    try { reintentarPendientes(); } catch (e) { Logger.log('Reintentos: ' + e.message); }
    if (!r.ok || !r.result.length) return;
    r.result.forEach(function(u) {
      offset = u.update_id + 1;
      props.setProperty('OFFSET', String(offset)); // se guarda antes, para no repetir si algo falla
      if (!u.message) return;
      try {
        manejarMensaje(u.message);
      } catch (e) {
        Logger.log('Error con el mensaje ' + u.message.message_id + ': ' + e.stack);
        try { responderEnTema(u.message, 'No pude procesar esta foto por un error: <i>' + escaparHtml(e.message) + '</i>'); } catch (e2) {}
      }
    });
  } finally {
    lock.releaseLock();
  }
}

function manejarMensaje(msg) {
  if (msg.chat.type !== 'supergroup' && msg.chat.type !== 'group') return;
  const props = PropertiesService.getScriptProperties();
  const grupo = props.getProperty('GRUPO_ID');
  if (!grupo) props.setProperty('GRUPO_ID', String(msg.chat.id)); // el primer grupo donde está el bot queda fijo
  else if (grupo !== String(msg.chat.id)) return;

  const tema = msg.is_topic_message ? String(msg.message_thread_id) : '';
  const texto = String(msg.text || msg.caption || '').trim();
  const sedes = leerSedes();
  let sede = sedeDeTema(sedes, tema, msg);

  // Comandos
  const cmd = texto.match(/^\/(sede|estado)(@\w+)?\s*(.*)$/i);
  if (cmd && cmd[1].toLowerCase() === 'sede') {
    const s = sedes.filter(function(x) { return normalizar(x.nombre) === normalizar(cmd[3]); })[0];
    if (!s || !tema) return responderEnTema(msg, 'Escribe /sede Barranquilla o /sede Baranoa dentro del tema de esa sede.');
    guardarTema(s, tema);
    return responderEnTema(msg, 'Listo: este tema queda para el tanque de <b>' + s.nombre + '</b>.');
  }
  if (cmd && cmd[1].toLowerCase() === 'estado') {
    const lista = sede ? [sede] : sedes;
    return responderEnTema(msg, lista.map(function(s) { return textoEstado(s); }).join('\n\n'));
  }

  // Fotos
  let fileId = null;
  if (msg.photo && msg.photo.length) fileId = msg.photo[msg.photo.length - 1].file_id;
  else if (msg.document && /^image\//.test(msg.document.mime_type || '')) fileId = msg.document.file_id;
  if (!fileId) return;

  if (!sede) return; // tema que no es de un tanque: el bot no dice nada

  const lecturas = leerLecturas();
  const msgId = msg.chat.id + ':' + msg.message_id;
  if (lecturas.some(function(l) { return l.msgId === msgId; })) return; // ya procesada

  // Foto a Drive
  const tMs = msg.date * 1000;
  const tz = Session.getScriptTimeZone();
  const info = telegram('getFile', { file_id: fileId });
  const token = props.getProperty('TELEGRAM_TOKEN');
  const blob = UrlFetchApp.fetch('https://api.telegram.org/file/bot' + token + '/' + info.result.file_path).getBlob();
  const nombre = 'Tanque_' + sede.nombre + '_' + Utilities.formatDate(new Date(tMs), tz, 'yyyy-MM-dd_HHmm') + '.jpg';
  if (!/^image\//.test(blob.getContentType() || '')) blob.setContentType('image/jpeg'); // Telegram a veces no marca el tipo
  blob.setName(nombre);
  const archivo = DriveApp.getFolderById(props.getProperty('CARPETA_ID')).createFile(blob);

  // Lectura con IA
  const ia = leerManometro(blob, sede);
  const val = validarLectura(ia, sede);
  const aguja = val.valida ? Math.round(Number(ia.lectura) * 10) / 10 : null;
  const nivel = aguja === null ? null : Math.min(aguja, sede.capacidad);

  // Cálculos
  const previas = validasDe(lecturas, sede.nombre, tMs);
  const a = analizar({ t: tMs, nivel: nivel }, previas, recargasDe(previas), sede);

  const aprox = !val.valida && ia && ia.lectura !== null && ia.lectura !== undefined && ia.lectura !== '' && !isNaN(Number(ia.lectura)) ? Number(ia.lectura) : null;
  const tecnico = [msg.from && msg.from.first_name, msg.from && msg.from.last_name].filter(Boolean).join(' ');
  agregarFila({
    'ID': 'T-' + Utilities.formatDate(new Date(tMs), tz, 'yyyyMMdd-HHmmss') + '-' + sede.nombre.slice(0, 3).toUpperCase(),
    'Fecha y hora': new Date(tMs),
    'Sede': sede.nombre,
    'Nivel (pulg)': val.valida ? nivel : (aprox === null ? '' : aprox),
    '% lleno': val.valida ? a.pct : '',
    'Cambio (pulg)': a.cambio === null ? '' : a.cambio,
    'Movimiento': val.valida ? a.movimiento : 'Sin lectura',
    'Consumo por día (pulg)': a.porDia === null ? '' : Math.round(a.porDia * 10) / 10,
    'Horas desde la anterior': a.horas === null ? '' : Math.round(a.horas * 10) / 10,
    'Lectura': val.valida ? 'Automática' : 'Dudosa',
    'Técnico': tecnico,
    'Foto': archivo.getUrl(),
    'Nota de la lectura': !val.valida ? val.motivo : (aguja > sede.capacidad ? 'La aguja marcaba ' + num(aguja) + '; se toma la capacidad del tanque (' + sede.capacidad + ').' : (ia.motivo || '')),
    'ID del mensaje': msgId,
    'ID de la foto': archivo.getId()
  });

  if (ia && ia.errorIA) {
    responderEnTema(msg, '<b>Tanque ' + sede.nombre + '</b> · ' + fechaCorta(partesFecha(tMs)) + '\n' +
      'La IA de Google está ocupada en este momento. Reintentaré sola en unos minutos y responderé aquí con el reporte.');
  } else {
    responderEnTema(msg, reporte(sede, a, tMs, partesFecha, val.valida ? null : val.motivo, aprox));
  }
  actualizarResumen();
}

/** Vuelve a leer con la IA las fotos que quedaron pendientes porque Google estaba ocupado. */
function reintentarPendientes() {
  const props = PropertiesService.getScriptProperties();
  const pend = leerLecturas().filter(function(l) {
    // También las de antes de este cambio que dijeron "código 503/500/429" (Google ocupado)
    return l.estado === 'Dudosa' && l.fotoId && (l.nota.indexOf(IA_NO_DISPONIBLE) === 0 || /código (503|500|429)/.test(l.nota));
  }).slice(0, 3);
  if (!pend.length) return;
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_LEC);
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(function(t) { return String(t).trim(); });
  const poner = function(fila, n, v) { const c = enc.indexOf(n); if (c >= 0) h.getRange(fila, c + 1).setValue(v === null || v === undefined ? '' : v); };
  const sedes = leerSedes();
  pend.forEach(function(l) {
    const sede = sedes.filter(function(x) { return normalizar(x.nombre) === normalizar(l.sede); })[0];
    if (!sede) return;
    const k = 'REINTENTO_' + l.id;
    const n = Number(props.getProperty(k) || 0) + 1;
    props.setProperty(k, String(n));
    const ia = leerManometro(DriveApp.getFileById(l.fotoId).getBlob(), sede);
    if (ia.errorIA && n < MAX_REINTENTOS) return; // sigue ocupada: se intenta en la próxima revisión
    props.deleteProperty(k);
    const val = ia.errorIA ? { valida: false, motivo: 'La IA de Google no respondió después de varios intentos. Revisa la foto y escribe el nivel a mano.' } : validarLectura(ia, sede);
    const aguja = val.valida ? Math.round(Number(ia.lectura) * 10) / 10 : null;
    const nivel = aguja === null ? null : Math.min(aguja, sede.capacidad);
    const previas = validasDe(leerLecturas(), sede.nombre, l.t);
    const a = analizar({ t: l.t, nivel: nivel }, previas, recargasDe(previas), sede);
    const aprox = !val.valida && !ia.errorIA && ia.lectura !== null && ia.lectura !== undefined && ia.lectura !== '' && !isNaN(Number(ia.lectura)) ? Number(ia.lectura) : null;
    poner(l.fila, 'Nivel (pulg)', val.valida ? nivel : aprox);
    poner(l.fila, '% lleno', val.valida ? a.pct : null);
    poner(l.fila, 'Cambio (pulg)', a.cambio);
    poner(l.fila, 'Movimiento', val.valida ? a.movimiento : 'Sin lectura');
    poner(l.fila, 'Consumo por día (pulg)', a.porDia === null ? null : Math.round(a.porDia * 10) / 10);
    poner(l.fila, 'Horas desde la anterior', a.horas === null ? null : Math.round(a.horas * 10) / 10);
    poner(l.fila, 'Lectura', val.valida ? 'Automática' : 'Dudosa');
    poner(l.fila, 'Nota de la lectura', val.valida ? (aguja > sede.capacidad ? 'La aguja marcaba ' + num(aguja) + '; se toma la capacidad del tanque (' + sede.capacidad + ').' : (ia.motivo || '')) : val.motivo);
    // Responder debajo de la foto original
    const ids = l.msgId.split(':');
    if (ids.length === 2 && sede.tema) {
      responderEnTema({ chat: { id: ids[0] }, message_id: Number(ids[1]), is_topic_message: true, message_thread_id: Number(sede.tema) },
        reporte(sede, a, l.t, partesFecha, val.valida ? null : val.motivo, aprox));
    }
  });
  actualizarResumen();
}

function sedeDeTema(sedes, tema, msg) {
  if (tema) {
    const s = sedes.filter(function(x) { return x.tema === tema; })[0];
    if (s) return s;
    // Primera vez: se reconoce por el nombre del tema, que debe decir "tanque" y la sede
    // ("Fotos del tanque Baranoa"). Así no se confunde con "Reporte de equipos Barranquilla".
    const creado = msg.reply_to_message && msg.reply_to_message.forum_topic_created;
    if (creado) {
      const n = normalizar(creado.name);
      const s2 = n.indexOf('tanque') >= 0 && sedes.filter(function(x) { return n.indexOf(normalizar(x.nombre)) >= 0; })[0];
      if (s2 && !s2.tema) { guardarTema(s2, tema); return s2; }
    }
  }
  return null;
}

function guardarTema(sede, tema) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CFG);
  const ix = indices(h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0]);
  h.getRange(sede.fila, ix['Tema de Telegram (ID)'] + 1).setValue(String(tema));
  sede.tema = String(tema);
}


/***********************
 * GEMINI (lectura de la aguja)
 ***********************/

/** Busca en la lista de modelos de Gemini el "flash" más reciente que pueda leer imágenes. */
function buscarModelo() {
  const props = PropertiesService.getScriptProperties();
  const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', {
    muteHttpExceptions: true, headers: { 'x-goog-api-key': props.getProperty('GEMINI_KEY') }
  });
  if (r.getResponseCode() !== 200) throw new Error('No se pudo consultar la lista de modelos (código ' + r.getResponseCode() + ').');
  const lista = (JSON.parse(r.getContentText()).models || []).filter(function(m) {
    const n = m.name.toLowerCase();
    return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0 && n.indexOf('flash') >= 0 &&
      !/lite|image|tts|audio|live|preview|exp|thinking|embedding/.test(n);
  }).map(function(m) { return m.name.replace('models/', ''); });
  if (!lista.length) throw new Error('No hay un modelo Gemini flash disponible para esta clave.');
  const ver = function(n) { const m = n.match(/gemini-(\d+(?:\.\d+)?)/); return m ? Number(m[1]) : 0; };
  lista.sort(function(a, b) { return ver(b) - ver(a) || a.length - b.length; });
  props.setProperty('GEMINI_MODELO_AUTO', lista[0]);
  return lista[0];
}

function leerManometro(blob, sede) {
  const props = PropertiesService.getScriptProperties();
  let modelo = props.getProperty('GEMINI_MODELO') || props.getProperty('GEMINI_MODELO_AUTO') || MODELO_POR_DEFECTO;
  const desc = DESCRIPCION_MANOMETRO[normalizar(sede.nombre)] ||
    ('Manómetro de nivel con escala de 0 a ' + sede.escala + ' pulgadas de agua.');
  const prompt =
    'Eres un técnico que lee manómetros de nivel de tanques criogénicos de oxígeno medicinal.\n' +
    'Tanque de la sede ' + sede.nombre + '. ' + desc + '\n' +
    'Lee el valor que señala la PUNTA de la aguja principal sobre la escala, con un decimal si se puede. ' +
    'Interpola entre las marcas. Si la foto está de lado, corrige mentalmente la perspectiva.\n' +
    'Responde SOLO este JSON: {"es_manometro": true/false, "lectura": número o null, ' +
    '"confianza": "alta" | "media" | "baja", "motivo": "frase corta en español"}.\n' +
    'Usa confianza "baja" si la aguja no se ve clara, hay reflejo sobre ella o no estás seguro de la escala.';
  const cuerpo = {
    contents: [{ parts: [
      { text: prompt },
      { inline_data: { mime_type: blob.getContentType() || 'image/jpeg', data: Utilities.base64Encode(blob.getBytes()) } }
    ] }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json' }
  };
  try {
    const pedir = function(m) {
      return UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent', {
        method: 'post', contentType: 'application/json', payload: JSON.stringify(cuerpo), muteHttpExceptions: true,
        headers: { 'x-goog-api-key': props.getProperty('GEMINI_KEY') }
      });
    };
    let r = pedir(modelo);
    if (r.getResponseCode() === 404) { // el modelo ya no existe: se busca el vigente y se reintenta
      modelo = buscarModelo();
      r = pedir(modelo);
    }
    // Google ocupado (429/500/503): esperar y reintentar
    for (let i = 0; i < 3 && [429, 500, 503].indexOf(r.getResponseCode()) >= 0; i++) {
      Utilities.sleep(4000 * (i + 1));
      r = pedir(modelo);
    }
    if (r.getResponseCode() !== 200) {
      Logger.log('Gemini ' + modelo + ' respondió ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 500));
      return { es_manometro: true, lectura: null, confianza: 'baja', errorIA: true,
        motivo: IA_NO_DISPONIBLE + ' (código ' + r.getResponseCode() + ').' };
    }
    const j = JSON.parse(r.getContentText());
    const txt = j.candidates[0].content.parts.map(function(p) { return p.text || ''; }).join('');
    const ia = JSON.parse(txt.replace(/^```(json)?|```$/g, '').trim());
    return ia;
  } catch (e) {
    Logger.log('Error con Gemini: ' + e.message);
    return { es_manometro: true, lectura: null, confianza: 'baja', errorIA: true, motivo: IA_NO_DISPONIBLE + ': ' + e.message };
  }
}

/** Menú: muestra qué modelo de Gemini se está usando (lo busca si hace falta). */
function probarModelo() {
  let m;
  try { m = buscarModelo(); } catch (e) { m = 'Error: ' + e.message; }
  SpreadsheetApp.getUi().alert('Modelo de Gemini: ' + m);
}

/** Menú: prueba la IA con la foto más reciente de la carpeta (no guarda nada). */
function probarIA() {
  const props = PropertiesService.getScriptProperties();
  const it = DriveApp.getFolderById(props.getProperty('CARPETA_ID')).getFiles();
  let ultimo = null;
  while (it.hasNext()) { const f = it.next(); if (!ultimo || f.getDateCreated() > ultimo.getDateCreated()) ultimo = f; }
  if (!ultimo) return SpreadsheetApp.getUi().alert('Aún no hay fotos en la carpeta.');
  const sede = leerSedes().filter(function(s) { return ultimo.getName().indexOf(s.nombre) >= 0; })[0] || leerSedes()[0];
  const ia = leerManometro(ultimo.getBlob(), sede);
  SpreadsheetApp.getUi().alert(ultimo.getName() + '\n\n' + JSON.stringify(ia, null, 2));
}


/***********************
 * ESCRITURA EN LA HOJA
 ***********************/

/** Agrega una fila dentro de la tabla, por nombre de encabezado, con la foto como chip de Drive. */
function agregarFila(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(HOJA_LEC);
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(function(t) { return String(t).trim(); });
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    // 1) Por la API, dentro de la tabla y con el chip en la misma escritura
    try {
      const tablaId = idTabla(ss, 'Lecturas');
      const celdas = enc.map(function(n) { return celdaApi(n, datos[n]); });
      const pedido = { appendCells: { sheetId: h.getSheetId(), rows: [{ values: celdas }], fields: 'userEnteredValue,chipRuns' } };
      if (tablaId) pedido.appendCells.tableId = tablaId;
      Sheets.Spreadsheets.batchUpdate({ requests: [pedido] }, ss.getId());
      return;
    } catch (e) {
      Logger.log('appendCells falló, se usa el método simple: ' + e.message);
    }
    // 2) Método simple: fila normal + chip aparte (o enlace con texto corto)
    const fila = enc.map(function(n) { return n === 'Foto' ? '' : (datos[n] === undefined ? '' : datos[n]); });
    h.appendRow(fila);
    const r = h.getLastRow();
    const c = enc.indexOf('Foto') + 1;
    if (c > 0 && datos['Foto']) ponerChip(ss, h, r, c, datos['Foto']);
  } finally {
    lock.releaseLock();
  }
}

function celdaApi(nombre, v) {
  if (nombre === 'Foto') {
    return v ? { userEnteredValue: { stringValue: '@' }, chipRuns: [{ startIndex: 0, chip: { richLinkProperties: { uri: v } } }] } : {};
  }
  if (v === '' || v === null || v === undefined) return {};
  if (v instanceof Date) {
    return { userEnteredValue: { numberValue: serialFecha(v) } };
  }
  if (typeof v === 'number') return { userEnteredValue: { numberValue: v } };
  return { userEnteredValue: { stringValue: String(v) } };
}

/** Fecha de JavaScript -> número de serie de Sheets en hora de Colombia. */
function serialFecha(d) {
  const z = Utilities.formatDate(d, Session.getScriptTimeZone(), 'Z'); // "-0500"
  const off = (z[0] === '-' ? -1 : 1) * (Number(z.substr(1, 2)) * 60 + Number(z.substr(3, 2)));
  return (d.getTime() + off * 60000) / 86400000 + 25569;
}

function idTabla(ss, nombre) {
  const cache = PropertiesService.getScriptProperties();
  const k = 'TABLA_' + nombre;
  if (cache.getProperty(k)) return cache.getProperty(k);
  const info = Sheets.Spreadsheets.get(ss.getId(), { fields: 'sheets(tables(tableId,name))' });
  let id = null;
  info.sheets.forEach(function(s) { (s.tables || []).forEach(function(t) { if (t.name === nombre) id = t.tableId; }); });
  if (id) cache.setProperty(k, id);
  return id;
}

function ponerChip(ss, h, fila, col, url) {
  try {
    Sheets.Spreadsheets.batchUpdate({ requests: [{ updateCells: {
      rows: [{ values: [{ userEnteredValue: { stringValue: '@' }, chipRuns: [{ startIndex: 0, chip: { richLinkProperties: { uri: url } } }] }] }],
      fields: 'userEnteredValue,chipRuns',
      start: { sheetId: h.getSheetId(), rowIndex: fila - 1, columnIndex: col - 1 }
    } }] }, ss.getId());
  } catch (e) {
    h.getRange(fila, col).setRichTextValue(SpreadsheetApp.newRichTextValue().setText('Ver foto').setLinkUrl(url).build());
  }
}

/** Menú: recalcula %, cambio, movimiento y consumo de todas las filas (útil tras corregir lecturas a mano). */
function recalcularTodo() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_LEC);
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(function(t) { return String(t).trim(); });
  const lecturas = leerLecturas();
  leerSedes().forEach(function(sede) {
    const previas = [];
    lecturas.filter(function(l) { return normalizar(l.sede) === normalizar(sede.nombre); }).forEach(function(l) {
      if (!esValida(l)) return;
      const a = analizar({ t: l.t, nivel: l.nivel }, previas.slice(), recargasDe(previas), sede);
      const poner = function(n, v) { const c = enc.indexOf(n); if (c >= 0) h.getRange(l.fila, c + 1).setValue(v === null ? '' : v); };
      poner('% lleno', a.pct);
      poner('Cambio (pulg)', a.cambio);
      poner('Movimiento', a.movimiento);
      poner('Consumo por día (pulg)', a.porDia === null ? null : Math.round(a.porDia * 10) / 10);
      poner('Horas desde la anterior', a.horas === null ? null : Math.round(a.horas * 10) / 10);
      previas.push({ t: l.t, nivel: l.nivel });
    });
  });
  actualizarResumen();
}


/***********************
 * RESUMEN
 ***********************/

function resumenSede(sede, lecturas) {
  const serie = validasDe(lecturas, sede.nombre);
  const ahora = Date.now();
  const ult = serie.length ? serie[serie.length - 1] : null;
  const rec = recargasDe(serie);
  const ultRec = rec.length ? rec[rec.length - 1] : null;
  const p7 = promedioConsumo(serie, ahora, 7);
  const p30 = promedioConsumo(serie, ahora, 30);
  const dias = ult && p7 && p7 > 0.2 && ult.nivel > sede.aviso ? (ult.nivel - sede.aviso) / p7 : null;
  const estado = !ult ? 'Sin lecturas' : (ult.nivel <= sede.aviso ? 'Necesita recarga' : (dias !== null && dias < 2 ? 'Recarga pronto' : 'Normal'));
  return {
    sede: sede.nombre, capacidad: sede.capacidad, escala: sede.escala, aviso: sede.aviso,
    nivel: ult ? ult.nivel : null, pct: ult ? porcentaje(ult.nivel, sede.capacidad) : null, ultimaT: ult ? ult.t : null,
    prom7: p7, prom30: p30, recargaT: ultRec ? ultRec.t : null, recargaCambio: ultRec ? ultRec.cambio : null,
    dias: dias, estado: estado
  };
}

function actualizarResumen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(HOJA_RES);
  const lecturas = leerLecturas();
  const filas = leerSedes().map(function(s) {
    const r = resumenSede(s, lecturas);
    const red = function(x) { return x === null ? '' : Math.round(x * 10) / 10; };
    return [r.sede, red(r.nivel), r.pct === null ? '' : r.pct, r.ultimaT ? new Date(r.ultimaT) : '',
      red(r.prom7), red(r.prom30), r.recargaT ? new Date(r.recargaT) : '', red(r.recargaCambio), red(r.dias), r.estado];
  });
  if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, ENC_RES.length).clearContent();
  if (filas.length) {
    h.getRange(2, 1, filas.length, ENC_RES.length).setValues(filas);
    formato(h.getRange(2, 4, filas.length, 1), 'dd/mm/yyyy h:mm am/pm');
    formato(h.getRange(2, 7, filas.length, 1), 'dd/mm/yyyy');
  }
}

function textoEstado(sede) {
  const r = resumenSede(sede, leerLecturas());
  if (r.nivel === null) return '<b>Tanque ' + sede.nombre + '</b>: aún sin lecturas.';
  const L = ['<b>Tanque ' + sede.nombre + '</b> · ' + fechaCorta(partesFecha(r.ultimaT)),
    'Nivel: <b>' + num(r.nivel) + ' pulg</b> (' + r.pct + ' % de ' + sede.capacidad + ')'];
  if (r.prom7 !== null) L.push('Promedio 7 días: −' + num(r.prom7) + ' pulg/día');
  if (r.recargaT) L.push('Última recarga: ' + partesFecha(r.recargaT).dia + ' ' + MESES_CORTOS[partesFecha(r.recargaT).mes - 1] + ' (' + conSigno(r.recargaCambio) + ' pulg)');
  if (r.estado === 'Necesita recarga') L.push('⚠️ <b>Necesita recarga</b> (nivel de recarga: ' + sede.aviso + ' pulg).');
  else if (r.dias !== null) L.push('Llega a ' + sede.aviso + ' pulg en ~' + num(r.dias) + ' días.');
  return L.join('\n');
}


/***********************
 * API PARA LA PWA (tablero)
 ***********************/

function doGet() {
  return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">Servidor de tanques criogénicos activo. Usa el tablero instalado.</p>');
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'Solicitud no válida.' }); }
  const codigo = PropertiesService.getScriptProperties().getProperty('CODIGO_PWA');
  if (!codigo || String(d.codigo || '').trim() !== codigo) return json({ ok: false, codigo: 'ACCESO', error: 'Código de acceso incorrecto.' });
  try {
    if (d.accion === 'validar') return json({ ok: true });
    if (d.accion === 'datos') return json(datosTablero(Number(d.dias) || 60));
    return json({ ok: false, error: 'Acción desconocida.' });
  } catch (err) {
    return json({ ok: false, error: 'Error en el servidor: ' + err.message });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function datosTablero(dias) {
  const lecturas = leerLecturas();
  const desde = Date.now() - dias * 86400000;
  return {
    ok: true,
    generado: Date.now(),
    sedes: leerSedes().map(function(s) { return resumenSede(s, lecturas); }),
    lecturas: lecturas.filter(function(l) { return l.t >= desde; }).map(function(l) {
      return { t: l.t, sede: l.sede, nivel: l.nivel, pct: l.pct, cambio: l.cambio, mov: l.movimiento, porDia: l.porDia,
        estado: l.estado, tecnico: l.tecnico, nota: l.nota,
        foto: l.fotoId ? 'https://drive.google.com/file/d/' + l.fotoId + '/view' : '' };
    })
  };
}
