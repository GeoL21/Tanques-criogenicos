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
const HOJA_REC = 'Recargas';
const CARPETA_FOTOS = 'Fotos tanques criogénicos';
const MINUTOS_REVISION = 5;
const MODELO_POR_DEFECTO = 'gemini-2.5-flash';
const IA_NO_DISPONIBLE = 'La IA de Google no estuvo disponible';
const HORAS_REINTENTO = 24; // las fotos pendientes se reintentan hasta 24 h

const ENC_LEC = ['ID', 'Fecha y hora', 'Sede', 'Nivel (pulg)', '% lleno', 'Cambio (pulg)', 'Movimiento',
  'Consumo por día (pulg)', 'Horas desde la anterior', 'Lectura', 'Técnico', 'Foto', 'Nota de la lectura',
  'ID del mensaje', 'ID de la foto'];
const ENC_CFG = ['Sede', 'Capacidad (pulg)', 'Escala del manómetro (pulg)', 'Nivel de recarga (pulg)', 'Tema de Telegram (ID)'];
const ENC_RES = ['Sede', 'Nivel actual (pulg)', '% lleno', 'Última lectura', 'Promedio 7 días (pulg/día)',
  'Promedio 30 días (pulg/día)', 'Última recarga', 'Pulgadas recargadas', 'Días hasta el nivel de recarga', 'Estado'];

const ENC_REC = ['ID', 'Fecha y hora', 'Sede', 'Inicial (pulg)', 'Final (pulg)', 'Pulgadas recargadas', 'N.º remisión',
  'Estado', 'Registró', 'Foto del nivel', 'IDs de mensajes del bot', 'ID de la lectura',
  'Fecha de la remisión', 'Presión inicial (bar)', 'Presión final (bar)', 'Foto de la remisión',
  'm³ por tabla', 'Precio por m³', 'Costo estimado', 'N.º factura', 'm³ facturados', 'Costo facturado'];
/** Columnas que se agregaron después de crear la tabla Recargas (se añaden solas una vez). */
const COLS_REC_NUEVAS = { 'Fecha de la remisión': 'DATE', 'Presión inicial (bar)': 'DOUBLE', 'Presión final (bar)': 'DOUBLE', 'Foto de la remisión': 'FILES_CHIP',
  // Costos (ver Costos.gs)
  'm³ por tabla': 'DOUBLE', 'Precio por m³': 'CURRENCY', 'Costo estimado': 'CURRENCY',
  'N.º factura': 'TEXT', 'm³ facturados': 'DOUBLE', 'Costo facturado': 'CURRENCY' };
const VERSION_COLUMNAS = '2';
const ESTADOS_RECARGA = ['Falta remisión', 'Completa'];
const PEDIR_REMISION = '📝 <b>Responde a este mensaje con el número de remisión de Messer.</b>';
const HORA_RECORDATORIO = 7; // 7:00 a. m.
const COLUMNAS_CHIP = ['Foto', 'Foto del nivel', 'Foto de la remisión'];

const MOVIMIENTOS = ['Consumo', 'Se mantuvo', 'Recarga', 'Primera lectura', 'Sin lectura'];
const ESTADOS_LECTURA = ['Automática', 'Dudosa', 'Corregida', 'Reemplazada', 'Descartada'];
const VERSION_LISTAS = '3'; // sube cuando cambian las opciones de los desplegables

/** Cómo es cada manómetro, para que la IA lo lea bien. */
/**
 * "Lupa": segunda mirada a una foto ampliada alrededor de la punta de la aguja.
 * Solo en los tanques que aparecen aquí; paso = pulgadas que vale cada raya pequeña.
 * Necesita el servicio avanzado "Google Slides API" (se usa para recortar la foto).
 */
const LUPA = { barranquilla: { paso: 2, cadaNumero: 20 } };
const LUPA_MAX_DIF = 8; // si la lupa y el cálculo se separan más que esto, se deja el cálculo

const DESCRIPCION_MANOMETRO = {
  barranquilla: 'Manómetro diferencial MID-WEST INSTRUMENT de carátula NEGRA con números blancos de 0 a 200 ' +
    '"INCHES OF WATER" (Δp). Números cada 20 y marcas pequeñas cada 2. La aguja principal es BLANCA y larga. ' +
    'Ignora el dial pequeño de la derecha (increase switch / set point), las pestañas amarillas del borde y los reflejos del vidrio.',
  baranoa: 'Indicador de contenido de carátula BEIGE con números negros de 0 a 100 "Inch H2O". ' +
    'La escala recorre el anillo exterior en sentido horario: 0 abajo a la izquierda, 20 a la izquierda, 40 arriba a la izquierda, ' +
    '50 arriba al centro (marca larga sin número), 60 arriba a la derecha, 80 a la derecha y 100 abajo a la derecha. ' +
    'Marcas largas cada 10 y cortas cada 2. Los NÚMEROS impresos están MÁS ADENTRO que las marcas y la punta de la aguja ' +
    'no llega al anillo de marcas, así que NO tomes el número que tapa o toca la punta. ' +
    'La aguja es NEGRA, sale del eje central (sector negro con forma de abanico); ignora los tornillos y los puntos grises.'
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
  const rec = asegurarHoja(ss, HOJA_REC, ENC_REC);

  // Datos de las sedes
  if (!String(cfg.getRange(2, 1).getValue()).trim()) {
    cfg.getRange(2, 1, 2, 4).setValues([
      ['Barranquilla', 136, 200, 40],
      ['Baranoa', 80, 100, 24]
    ]);
  }

  // Tablas (con tipos de columna)
  crearTablas(ss, lec, cfg, rec);
  hojaPrecios();
  fechasEnEspanol(ss);

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
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'recordarRemisiones') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('recordarRemisiones').timeBased().atHour(HORA_RECORDATORIO).everyDays(1).inTimezone('America/Bogota').create();
  if (props.getProperty('TELEGRAM_TOKEN')) botId();

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
function crearTablas(ss, lec, cfg, rec) {
  const id = ss.getId();
  let existentes = [];
  try {
    const info = Sheets.Spreadsheets.get(id, { fields: 'sheets(properties(sheetId,title),tables(tableId,name))' });
    info.sheets.forEach(function(s) { (s.tables || []).forEach(function(t) { existentes.push(t.name); }); });
  } catch (e) {
    Logger.log('Activa el servicio avanzado "Google Sheets API": ' + e.message);
    formatoSinTabla(lec); formatoSinTabla(cfg); formatoSinTabla(rec);
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
  if (existentes.indexOf('Recargas') < 0) {
    const tiposRec = { 'ID': 'TEXT', 'Fecha y hora': 'DATE_TIME', 'Sede': 'DROPDOWN', 'Inicial (pulg)': 'DOUBLE', 'Final (pulg)': 'DOUBLE',
      'Pulgadas recargadas': 'DOUBLE', 'N.º remisión': 'TEXT', 'Estado': 'DROPDOWN', 'Registró': 'TEXT', 'Foto del nivel': 'FILES_CHIP',
      'IDs de mensajes del bot': 'TEXT', 'ID de la lectura': 'TEXT',
      'Fecha de la remisión': 'DATE', 'Presión inicial (bar)': 'DOUBLE', 'Presión final (bar)': 'DOUBLE', 'Foto de la remisión': 'FILES_CHIP',
      'm³ por tabla': 'DOUBLE', 'Precio por m³': 'CURRENCY', 'Costo estimado': 'CURRENCY',
      'N.º factura': 'TEXT', 'm³ facturados': 'DOUBLE', 'Costo facturado': 'CURRENCY' };
    const valRec = { 'Sede': lista(['Barranquilla', 'Baranoa']), 'Estado': lista(ESTADOS_RECARGA) };
    pedidos.push({ addTable: { table: {
      name: 'Recargas',
      range: { sheetId: rec.getSheetId(), startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: ENC_REC.length },
      columnProperties: ENC_REC.map(function(n, i) {
        const c = { columnIndex: i, columnName: n, columnType: tiposRec[n] };
        if (valRec[n]) c.dataValidationRule = valRec[n];
        return c;
      })
    } } });
  }
  if (!pedidos.length) return;
  try {
    Sheets.Spreadsheets.batchUpdate({ requests: pedidos }, id);
  } catch (e) {
    Logger.log('No se pudieron crear las tablas, se usa formato simple: ' + e.message);
    formatoSinTabla(lec); formatoSinTabla(cfg); formatoSinTabla(rec);
  }
  try { lec.hideColumns(14, 2); } catch (e) {}
  try { rec.hideColumns(11, 2); } catch (e) {}
}

/** Muestra las fechas de las tablas como día/mes/año y hora con a. m./p. m. */
function fechasEnEspanol(ss) {
  try {
    const info = Sheets.Spreadsheets.get(ss.getId(), { fields: 'sheets(tables(tableId,name,columnProperties))' });
    const pedidos = [];
    info.sheets.forEach(function(h) {
      (h.tables || []).forEach(function(t) {
        let cambia = false;
        const cols = (t.columnProperties || []).map(function(c) {
          if (c.columnType === 'DATE_TIME') { c.columnFormat = { pattern: 'dd/mm/yyyy h:mm am/pm' }; cambia = true; }
          return c;
        });
        if (cambia) pedidos.push({ updateTable: { table: { tableId: t.tableId, columnProperties: cols }, fields: 'columnProperties' } });
      });
    });
    if (pedidos.length) Sheets.Spreadsheets.batchUpdate({ requests: pedidos }, ss.getId());
  } catch (e) {
    Logger.log('No se pudo cambiar el formato de fecha de las tablas: ' + e.message);
  }
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
  } else if (h.getName() === HOJA_CFG) {
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
    if (props.getProperty('LISTAS_V') !== VERSION_LISTAS) actualizarListas();
    if (props.getProperty('COLUMNAS_V') !== VERSION_COLUMNAS) agregarColumnasRecargas();
    let offset = Number(props.getProperty('OFFSET') || 0);
    const r = telegram('getUpdates', { offset: offset, timeout: 0, allowed_updates: ['message'] });
    try { reintentarPendientes(); } catch (e) { Logger.log('Reintentos: ' + e.message); }
    try { sincronizarCostos(); } catch (e) { Logger.log('Costos: ' + e.message); }
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
  const cmd = texto.match(/^\/(sede|estado|corregir)(@\w+)?\s*(.*)$/i);
  if (cmd && cmd[1].toLowerCase() === 'sede') {
    const s = sedes.filter(function(x) { return normalizar(x.nombre) === normalizar(cmd[3]); })[0];
    if (!s || !tema) return responderEnTema(msg, 'Escribe /sede Barranquilla o /sede Baranoa dentro del tema de esa sede.');
    guardarTema(s, tema);
    return responderEnTema(msg, 'Listo: este tema queda para el tanque de <b>' + s.nombre + '</b>.');
  }
  if (cmd && cmd[1].toLowerCase() === 'corregir') return manejarCorregir(msg, sede, cmd[3]);
  if (cmd && cmd[1].toLowerCase() === 'estado') {
    const lista = sede ? [sede] : sedes;
    return responderEnTema(msg, lista.map(function(s) { return textoEstado(s); }).join('\n\n'));
  }

  // Respuesta a un mensaje del bot (número de remisión)
  if (msg.text && msg.reply_to_message && esDelBot(msg.reply_to_message)) return manejarRespuestaRemision(msg);

  // Fotos
  let fileId = null;
  if (msg.photo && msg.photo.length) fileId = msg.photo[msg.photo.length - 1].file_id;
  else if (msg.document && /^image\//.test(msg.document.mime_type || '')) fileId = msg.document.file_id;
  if (!fileId) return;

  if (!sede) return; // tema que no es de un tanque: el bot no dice nada

  registrarTecnico(msg.from);
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
  if (ia && ia.es_remision === true && ia.es_manometro !== true) {
    archivo.setName('Remision_' + sede.nombre + '_' + Utilities.formatDate(new Date(tMs), tz, 'yyyy-MM-dd_HHmm') + '.jpg');
    return manejarFotoRemision(msg, sede, blob, archivo, tMs);
  }
  const val = validarLectura(ia, sede);
  const aguja = val.valida ? Math.round(Number(ia.lectura) * 10) / 10 : null;
  const nivel = aguja === null ? null : Math.min(aguja, sede.capacidad);

  // ¿Otra foto del mismo tanque hace menos de 30 min? La nueva reemplaza a la anterior.
  const repetida = lecturas.filter(function(l) {
    return normalizar(l.sede) === normalizar(sede.nombre) && l.estado !== 'Reemplazada' && l.estado !== 'Descartada' &&
      l.t <= tMs && tMs - l.t < REGLAS.minutosRepetida * 60000;
  }).pop() || null;
  const base = repetida ? lecturas.filter(function(l) { return l !== repetida; }) : lecturas;

  // Cálculos
  const previas = validasDe(base, sede.nombre, tMs);
  const a = analizar({ t: tMs, nivel: nivel }, previas, recargasDe(previas), sede);

  const aprox = !val.valida && ia && ia.lectura !== null && ia.lectura !== undefined && ia.lectura !== '' && !isNaN(Number(ia.lectura)) ? Number(ia.lectura) : null;
  const tecnico = [msg.from && msg.from.first_name, msg.from && msg.from.last_name].filter(Boolean).join(' ');
  const idLectura = 'T-' + Utilities.formatDate(new Date(tMs), tz, 'yyyyMMdd-HHmmss') + '-' + sede.nombre.slice(0, 3).toUpperCase();
  agregarFila({
    'ID': idLectura,
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

  const aviso = repetida ? '♻️ <i>Esta foto reemplaza la de las ' + horaAmPm(partesFecha(repetida.t).hora, partesFecha(repetida.t).min) + '</i>\n' : '';
  if (repetida) marcarReemplazada(repetida, tMs, val.valida && a.movimiento === 'Recarga');

  if (ia && ia.errorIA) {
    const envO = responderEnTema(msg, aviso + '<b>Tanque ' + sede.nombre + '</b> · ' + fechaCorta(partesFecha(tMs)) + '\n' +
      'La IA de Google está ocupada en este momento. Reintentaré sola en unos minutos y responderé aquí con el reporte.');
    guardarReporte(envO, idLectura);
  } else {
    const esRecarga = val.valida && a.movimiento === 'Recarga';
    const env = responderEnTema(msg, aviso + reporte(sede, a, tMs, partesFecha, val.valida ? null : val.motivo, aprox, revisarDe(ia, val)) + (esRecarga ? '\n\n' + PEDIR_REMISION : ''));
    guardarReporte(env, idLectura);
    if (esRecarga) registrarRecarga(sede, a, tMs, idLectura, archivo.getUrl(), env && env.ok ? env.result.message_id : '', repetida ? repetida.id : '');
  }
  actualizarResumen();
}

/** Vuelve a leer con la IA las fotos que quedaron pendientes porque Google estaba ocupado. */
function reintentarPendientes() {
  const props = PropertiesService.getScriptProperties();
  const pend = leerLecturas().filter(function(l) {
    // También las de antes de este cambio que dijeron "código 503/500/429" (Google ocupado)
    return l.estado === 'Dudosa' && l.fotoId && (l.nota.indexOf(IA_NO_DISPONIBLE) === 0 || /código (503|500|429)/.test(l.nota));
  }).slice(0, 2);
  if (!pend.length || iaEnPausa()) return;
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_LEC);
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(function(t) { return String(t).trim(); });
  const poner = function(fila, n, v) { const c = enc.indexOf(n); if (c >= 0) h.getRange(fila, c + 1).setValue(v === null || v === undefined ? '' : v); };
  const sedes = leerSedes();
  pend.forEach(function(l) {
    const sede = sedes.filter(function(x) { return normalizar(x.nombre) === normalizar(l.sede); })[0];
    if (!sede) return;
    if (iaEnPausa()) return;
    const ia = leerManometro(DriveApp.getFileById(l.fotoId).getBlob(), sede);
    const vencida = Date.now() - l.t > HORAS_REINTENTO * 3600000;
    if (ia.errorIA && !vencida) return; // sigue sin IA: se intenta en la próxima revisión
    const val = ia.errorIA ? { valida: false, motivo: 'La IA de Google no respondió en 24 horas. Revisa la foto y escribe el nivel a mano.' } : validarLectura(ia, sede);
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
      const esRecarga = val.valida && a.movimiento === 'Recarga';
      const env = responderEnTema({ chat: { id: ids[0] }, message_id: Number(ids[1]), is_topic_message: true, message_thread_id: Number(sede.tema) },
        reporte(sede, a, l.t, partesFecha, val.valida ? null : val.motivo, aprox, revisarDe(ia, val)) + (esRecarga ? '\n\n' + PEDIR_REMISION : ''));
      guardarReporte(env, l.id);
      if (esRecarga) registrarRecarga(sede, a, l.t, l.id, 'https://drive.google.com/file/d/' + l.fotoId + '/view', env && env.ok ? env.result.message_id : '');
    }
  });
  actualizarResumen();
}


/** Escribe un valor en una celda de una hoja, por nombre de encabezado. */
function ponerEn(nombreHoja, fila, nombre, valor) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(function(t) { return String(t).trim(); });
  const c = enc.indexOf(nombre);
  if (c >= 0) h.getRange(fila, c + 1).setValue(valor === null || valor === undefined ? '' : valor);
}

/** Recuerda qué mensaje del bot es el reporte de qué lectura (para /corregir respondiendo al reporte). */
function guardarReporte(env, idLectura) {
  if (!env || !env.ok || !idLectura) return;
  const props = PropertiesService.getScriptProperties();
  const m = JSON.parse(props.getProperty('REPORTES') || '{}');
  m[String(env.result.message_id)] = idLectura;
  const k = Object.keys(m);
  if (k.length > 300) k.slice(0, k.length - 300).forEach(function(x) { delete m[x]; });
  props.setProperty('REPORTES', JSON.stringify(m));
}

/**
 * /corregir 41 — un técnico corrige el nivel de una lectura desde Telegram.
 * Respondiendo al reporte del bot o a la foto: corrige esa lectura.
 * Sin responder, dentro del tema de la sede: corrige la última foto de ese tanque.
 */
function manejarCorregir(msg, sedeTema, argumento) {
  if (!esTecnico(msg.from)) return responderEnTema(msg, 'Solo los técnicos que suben las fotos del tanque pueden corregir lecturas.');
  const lecturas = leerLecturas();
  let l = null, porUltima = false;
  const r = msg.reply_to_message;
  if (r) {
    const idPorReporte = JSON.parse(PropertiesService.getScriptProperties().getProperty('REPORTES') || '{}')[String(r.message_id)];
    const idFoto = msg.chat.id + ':' + r.message_id;
    l = lecturas.filter(function(x) { return (idPorReporte && x.id === idPorReporte) || x.msgId === idFoto; })[0] || null;
  }
  if (!l) {
    if (!sedeTema) return responderEnTema(msg, 'Escribe <b>/corregir</b> dentro del tema del tanque, o respondiendo al reporte del bot.');
    l = lecturas.filter(function(x) { return normalizar(x.sede) === normalizar(sedeTema.nombre) && x.estado !== 'Reemplazada' && x.estado !== 'Descartada'; }).pop() || null;
    porUltima = true;
    if (!l) return responderEnTema(msg, 'Aún no hay lecturas del tanque ' + sedeTema.nombre + ' para corregir.');
  }
  const sede = leerSedes().filter(function(x) { return normalizar(x.nombre) === normalizar(l.sede); })[0];
  if (!sede) return;
  const m = String(argumento || '').match(/\d+(?:[.,]\d+)?/);
  if (!m) return responderEnTema(msg, 'Escribe el nivel en pulgadas, por ejemplo <b>/corregir 41</b>.');
  const v = Number(m[0].replace(',', '.'));
  if (v < 0 || v > sede.escala) return responderEnTema(msg, num(v) + ' no es posible: el manómetro de ' + sede.nombre + ' va de 0 a ' + sede.escala + ' pulg.');
  const nivel = Math.min(v, sede.capacidad);
  const quien = [msg.from.first_name, msg.from.last_name].filter(Boolean).join(' ');
  const antes = esValida(l) ? num(l.nivel) : 'sin lectura';

  ponerEn(HOJA_LEC, l.fila, 'Nivel (pulg)', nivel);
  ponerEn(HOJA_LEC, l.fila, 'Lectura', 'Corregida');
  ponerEn(HOJA_LEC, l.fila, 'Nota de la lectura', 'Corregida en Telegram por ' + quien + ' (antes ' + antes + ')' +
    (v > sede.capacidad ? '; la aguja marcaba ' + num(v) + ', se toma la capacidad' : '') + '.');
  recalcularTodo();

  // Recalcular esta lectura para el reporte y ajustar la pestaña Recargas
  const previas = validasDe(leerLecturas(), sede.nombre, l.t);
  const a = analizar({ t: l.t, nivel: nivel }, previas, recargasDe(previas), sede);
  const esRecarga = a.movimiento === 'Recarga';
  const recPrevia = leerRecargas().filter(function(x) { return x.lectura === l.id; })[0];
  if (!esRecarga && recPrevia && recPrevia.estado === 'Falta remisión') {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_REC).deleteRow(recPrevia.fila);
  } else if (esRecarga && recPrevia) {
    ponerEnRecarga(recPrevia.fila, 'Final (pulg)', a.nivel);
    ponerEnRecarga(recPrevia.fila, 'Pulgadas recargadas', a.cambio);
  }
  const pedir = esRecarga && (!recPrevia || recPrevia.estado === 'Falta remisión');
  const env = responderEnTema(msg, '✏️ <b>Lectura corregida</b> por ' + escaparHtml(quien) + ': <b>' + antes + ' → ' + num(nivel) + ' pulg</b>' +
    (porUltima ? '\n<i>Era la última foto de este tema (' + fechaCorta(partesFecha(l.t)) + ').</i>' : '') + '\n\n' +
    reporte(sede, a, l.t, partesFecha, null, null) + (pedir ? '\n\n' + PEDIR_REMISION : ''));
  guardarReporte(env, l.id);
  if (esRecarga && !recPrevia) {
    registrarRecarga(sede, a, l.t, l.id, l.fotoId ? 'https://drive.google.com/file/d/' + l.fotoId + '/view' : '', env && env.ok ? env.result.message_id : '');
  } else if (pedir && recPrevia && env && env.ok) {
    agregarIdBot(recPrevia, env.result.message_id);
  }
}

/** La lectura anterior queda como "Reemplazada" (no cuenta en los cálculos). */
function marcarReemplazada(l, tNueva, nuevaEsRecarga) {
  const p = partesFecha(tNueva);
  ponerEn(HOJA_LEC, l.fila, 'Lectura', 'Reemplazada');
  ponerEn(HOJA_LEC, l.fila, 'Nota de la lectura', 'Reemplazada por la foto de las ' + horaAmPm(p.hora, p.min) + (l.nota ? ' · ' + l.nota : ''));
  // Si la anterior había creado una recarga sin remisión y la nueva ya no es recarga, se quita esa fila.
  if (!nuevaEsRecarga) {
    const r = leerRecargas().filter(function(x) { return x.lectura === l.id && x.estado === 'Falta remisión'; })[0];
    if (r) SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_REC).deleteRow(r.fila);
  }
}

/** Pone en los desplegables de las tablas las opciones nuevas (se hace sola una vez). */
function actualizarListas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const info = Sheets.Spreadsheets.get(ss.getId(), { fields: 'sheets(tables(tableId,name,columnProperties))' });
    const pedidos = [];
    info.sheets.forEach(function(h) {
      (h.tables || []).forEach(function(t) {
        if (t.name !== 'Lecturas') return;
        const cols = (t.columnProperties || []).map(function(c) {
          if (c.columnName === 'Lectura') {
            c.dataValidationRule = { condition: { type: 'ONE_OF_LIST', values: ESTADOS_LECTURA.map(function(v) { return { userEnteredValue: v }; }) } };
          }
          return c;
        });
        pedidos.push({ updateTable: { table: { tableId: t.tableId, columnProperties: cols }, fields: 'columnProperties' } });
      });
    });
    if (pedidos.length) Sheets.Spreadsheets.batchUpdate({ requests: pedidos }, ss.getId());
    PropertiesService.getScriptProperties().setProperty('LISTAS_V', VERSION_LISTAS);
  } catch (e) {
    Logger.log('No se pudieron actualizar los desplegables: ' + e.message);
  }
}

/** Agrega a la tabla Recargas las columnas nuevas (datos leídos de la remisión), una sola vez. */
function agregarColumnasRecargas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(HOJA_REC);
  if (!h) return;
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(function(t) { return String(t).trim(); });
  const faltan = Object.keys(COLS_REC_NUEVAS).filter(function(n) { return enc.indexOf(n) < 0; });
  if (faltan.length) {
    let hecho = false;
    try {
      const info = Sheets.Spreadsheets.get(ss.getId(), { fields: 'sheets(properties(sheetId),tables(tableId,name,range,columnProperties))' });
      info.sheets.forEach(function(sh) {
        (sh.tables || []).forEach(function(t) {
          if (t.name !== 'Recargas' || hecho) return;
          const fin = t.range.endColumnIndex;
          const cols = (t.columnProperties || []).concat(faltan.map(function(n, i) {
            return { columnIndex: fin - t.range.startColumnIndex + i, columnName: n, columnType: COLS_REC_NUEVAS[n] };
          }));
          const rango = JSON.parse(JSON.stringify(t.range));
          rango.endColumnIndex = fin + faltan.length;
          Sheets.Spreadsheets.batchUpdate({ requests: [{ updateTable: {
            table: { tableId: t.tableId, range: rango, columnProperties: cols }, fields: 'range,columnProperties' } }] }, ss.getId());
          hecho = true;
        });
      });
    } catch (e) {
      Logger.log('No se pudo ampliar la tabla Recargas: ' + e.message);
    }
    if (!hecho) { // sin tabla: encabezados al final
      const c = h.getLastColumn() + 1;
      h.getRange(1, c, 1, faltan.length).setValues([faltan]);
    }
  }
  PropertiesService.getScriptProperties().setProperty('COLUMNAS_V', VERSION_COLUMNAS);
}

/***********************
 * RECARGAS Y REMISIONES
 ***********************/

/** ID del bot (para reconocer cuando alguien responde a un mensaje suyo). */
function botId() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('BOT_ID');
  if (!id) {
    const r = telegram('getMe', {});
    if (r.ok) { id = String(r.result.id); props.setProperty('BOT_ID', id); }
  }
  return id;
}

function esDelBot(m) {
  return !!(m && m.from && m.from.is_bot && String(m.from.id) === botId());
}

/** Técnico = quien ha subido alguna foto del tanque. Se guarda su ID de Telegram. */
function registrarTecnico(from) {
  if (!from || from.is_bot) return;
  const props = PropertiesService.getScriptProperties();
  const t = JSON.parse(props.getProperty('TECNICOS') || '{}');
  const nombre = [from.first_name, from.last_name].filter(Boolean).join(' ');
  if (t[from.id] !== nombre) { t[from.id] = nombre; props.setProperty('TECNICOS', JSON.stringify(t)); }
}

function esTecnico(from) {
  if (!from) return false;
  const t = JSON.parse(PropertiesService.getScriptProperties().getProperty('TECNICOS') || '{}');
  if (t[from.id] !== undefined) return true;
  // Quienes subieron fotos antes de que existiera esta lista: se reconocen por el nombre en Lecturas
  const nombre = normalizar([from.first_name, from.last_name].filter(Boolean).join(' '));
  if (nombre && leerLecturas().some(function(l) { return normalizar(l.tecnico) === nombre; })) { registrarTecnico(from); return true; }
  return false;
}

/** Lee la tabla Recargas como objetos. */
function leerRecargas() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_REC);
  if (!h || h.getLastRow() < 2) return [];
  const v = h.getRange(1, 1, h.getLastRow(), h.getLastColumn()).getValues();
  const ix = indices(v[0]);
  const out = [];
  for (let r = 1; r < v.length; r++) {
    const f = v[r];
    if (!(f[ix['Fecha y hora']] instanceof Date)) continue;
    out.push({
      fila: r + 1, id: String(f[ix['ID']] || ''), t: f[ix['Fecha y hora']].getTime(), sede: String(f[ix['Sede']] || ''),
      inicial: f[ix['Inicial (pulg)']], final: f[ix['Final (pulg)']], cambio: f[ix['Pulgadas recargadas']],
      remision: String(f[ix['N.º remisión']] || ''), estado: String(f[ix['Estado']] || ''), registro: String(f[ix['Registró']] || ''),
      idsBot: String(f[ix['IDs de mensajes del bot']] || '').split(',').map(function(x) { return x.trim(); }).filter(Boolean),
      lectura: String(f[ix['ID de la lectura']] || '')
    });
  }
  return out;
}

function ponerEnRecarga(fila, nombre, valor) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_REC);
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(function(t) { return String(t).trim(); });
  const c = enc.indexOf(nombre);
  if (c >= 0) h.getRange(fila, c + 1).setValue(valor);
}

function agregarIdBot(rec, id) {
  if (!id) return;
  const ids = rec.idsBot.concat([String(id)]);
  ponerEnRecarga(rec.fila, 'IDs de mensajes del bot', ids.join(','));
}

/** Crea la fila de la recarga (una sola vez por lectura) con estado "Falta remisión". */
function registrarRecarga(sede, a, tMs, idLectura, fotoUrl, idMensajeBot, idReemplazada) {
  const todas = leerRecargas();
  if (todas.some(function(r) { return r.lectura === idLectura; })) return;
  // La foto reemplaza a otra que ya había creado la recarga: se actualiza esa misma fila.
  const previa = idReemplazada ? todas.filter(function(r) { return r.lectura === idReemplazada; })[0] : null;
  if (previa) {
    ponerEnRecarga(previa.fila, 'Final (pulg)', a.nivel);
    ponerEnRecarga(previa.fila, 'Pulgadas recargadas', a.cambio);
    ponerEnRecarga(previa.fila, 'ID de la lectura', idLectura);
    agregarIdBot(previa, idMensajeBot);
    return;
  }
  agregarFilaEn(HOJA_REC, 'Recargas', {
    'ID': 'R-' + idLectura.replace(/^T-/, ''),
    'Fecha y hora': new Date(tMs),
    'Sede': sede.nombre,
    'Inicial (pulg)': a.anterior ? a.anterior.nivel : '',
    'Final (pulg)': a.nivel,
    'Pulgadas recargadas': a.cambio,
    'N.º remisión': '',
    'Estado': 'Falta remisión',
    'Registró': '',
    'Foto del nivel': fotoUrl,
    'IDs de mensajes del bot': idMensajeBot ? String(idMensajeBot) : '',
    'ID de la lectura': idLectura
  });
}

/**
 * El técnico respondió a un mensaje del bot de una recarga: se guarda el número de remisión.
 * También acepta correcciones de lo leído en la remisión: "remisión 1001702 antes 11 después 74".
 */
function manejarRespuestaRemision(msg) {
  const idResp = String(msg.reply_to_message.message_id);
  const rec = leerRecargas().filter(function(r) { return r.idsBot.indexOf(idResp) >= 0; })[0];
  if (!rec) return; // respondió a otro mensaje del bot: no hay nada que hacer
  if (!esTecnico(msg.from)) {
    return responderEnTema(msg, 'Solo los técnicos que suben las fotos del tanque pueden registrar el número de remisión.');
  }
  const txt = String(msg.text);
  const nAntes = txt.match(/antes\D{0,3}(\d+(?:[.,]\d+)?)/i), nDesp = txt.match(/despu[eé]s\D{0,3}(\d+(?:[.,]\d+)?)/i);
  const sinNiveles = txt.replace(/(antes|despu[eé]s)\D{0,3}\d+(?:[.,]\d+)?/gi, ' ');
  const m = sinNiveles.match(/\d{3,}/);
  if (!m && !nAntes && !nDesp) return responderEnTema(msg, 'Escribe solo el número de remisión, por ejemplo <b>995292</b>.');
  const nombre = [msg.from.first_name, msg.from.last_name].filter(Boolean).join(' ');
  const numero = m ? m[0] : rec.remision;
  if (numero) {
    ponerEnRecarga(rec.fila, 'N.º remisión', numero);
    ponerEnRecarga(rec.fila, 'Estado', 'Completa');
    ponerEnRecarga(rec.fila, 'Registró', nombre);
  }
  const aNum = function(x) { return x ? Number(x[1].replace(',', '.')) : null; };
  let ini = aNum(nAntes), fin = aNum(nDesp);
  if (ini !== null) ponerEnRecarga(rec.fila, 'Inicial (pulg)', ini); else ini = rec.inicial === '' ? null : Number(rec.inicial);
  if (fin !== null) ponerEnRecarga(rec.fila, 'Final (pulg)', fin); else fin = rec.final === '' ? null : Number(rec.final);
  const ajusteNivel = (nAntes || nDesp) && ini !== null && fin !== null;
  if (ajusteNivel) ponerEnRecarga(rec.fila, 'Pulgadas recargadas', Math.round((fin - ini) * 10) / 10);
  const p = partesFecha(rec.t);
  const env = responderEnTema(msg, (numero ? '✅ Remisión <b>' + numero + '</b> registrada' : '✅ Recarga actualizada') +
    ' para la recarga del ' + p.dia + ' ' + MESES_CORTOS[p.mes - 1] + ' en ' + rec.sede + '.' +
    (ajusteNivel ? '\nNivel: <b>' + num(ini) + ' → ' + num(fin) + ' pulg (' + signoTxt(fin - ini) + ')</b>.' : '') +
    '\n<i>Si te equivocaste, responde de nuevo al mensaje de la recarga con el número correcto.</i>');
  if (env && env.ok) agregarIdBot(rec, env.result.message_id);
}

function signoTxt(n) { const r = Math.round(n * 10) / 10; return (r > 0 ? '+' : r < 0 ? '−' : '') + num(Math.abs(r)); }

/** Lee con la IA los datos de una remisión de Messer (número, fecha, niveles y presiones). */
function leerRemision(blob, sede) {
  const prompt =
    'La foto es una REMISIÓN de Messer (entrega de oxígeno líquido) para el tanque criogénico de la sede ' + sede.nombre + '.\n' +
    'Lee estos datos; muchos están escritos A MANO, léelos con cuidado:\n' +
    '- "numero": el número impreso de la remisión, arriba a la derecha junto al código de barras (ej.: 1001702).\n' +
    '- "fecha": la fecha de entrega (casillas Día / Mes / Año), en formato AAAA-MM-DD.\n' +
    '- "indicador_antes" e "indicador_despues": en el recuadro "Entrega por indicador", Indicador Antes e Indicador Después.\n' +
    '- "nivel_inicial" y "nivel_final": en la tabla de abajo, columnas NIVEL INICIAL y NIVEL FINAL.\n' +
    '- "presion_inicial", "presion_final" y "unidad_presion" (por ejemplo "bar" o "psi"): columnas PRESIÓN INICIAL, PRESIÓN FINAL y UNIDAD.\n' +
    'Usa null en lo que no se vea o no esté escrito. No inventes números.\n' +
    'Responde SOLO este JSON: {"es_remision": true/false, "numero": "texto o null", "fecha": "AAAA-MM-DD o null", ' +
    '"indicador_antes": número o null, "indicador_despues": número o null, "nivel_inicial": número o null, "nivel_final": número o null, ' +
    '"presion_inicial": número o null, "presion_final": número o null, "unidad_presion": "texto o null", ' +
    '"confianza": "alta" | "media" | "baja", "motivo": "frase corta en español"}';
  return llamarGemini(prompt, blob);
}

/** Convierte lo que leyó la IA de la remisión en valores listos para guardar. */
function datosRemision(r, sede) {
  const n = function(v) { if (v === null || v === undefined || v === '') return null; const x = Number(String(v).replace(',', '.')); return isNaN(x) ? null : x; };
  const enEscala = function(v) { return v !== null && v >= 0 && v <= sede.escala ? v : null; };
  const antes = enEscala(n(r.indicador_antes) !== null ? n(r.indicador_antes) : n(r.nivel_inicial));
  const despues = enEscala(n(r.indicador_despues) !== null ? n(r.indicador_despues) : n(r.nivel_final));
  const psi = /psi/i.test(String(r.unidad_presion || ''));
  const bar = function(v) { v = n(v); return v === null ? null : Math.round((psi ? v / 14.504 : v) * 10) / 10; };
  const numero = String(r.numero || '').replace(/\D/g, '');
  let fecha = null;
  const f = String(r.fecha || '').match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (f) fecha = new Date(Number(f[1]), Number(f[2]) - 1, Number(f[3]), 12);
  return { numero: numero.length >= 4 ? numero : '', fecha: fecha, antes: antes,
    despues: despues !== null && antes !== null && despues <= antes ? null : despues,
    presIni: bar(r.presion_inicial), presFin: bar(r.presion_final), confianza: r.confianza || 'media' };
}

/** Llegó la foto de una remisión: se leen sus datos y se guardan en la recarga reciente de esa sede. */
function manejarFotoRemision(msg, sede, blob, archivo, tMs) {
  const cab = '🧾 <b>Remisión de Messer · ' + sede.nombre + '</b>\n';
  const r = leerRemision(blob, sede);
  if (r.errorIA) {
    return responderEnTema(msg, cab + 'La IA de Google está ocupada y no pude leer la remisión. Vuelve a enviar la foto en unos minutos, ' +
      'o responde al mensaje de la recarga con el número de remisión.');
  }
  const d = datosRemision(r, sede);
  // Recarga a la que pertenece: la más reciente de la sede (hasta 2 días antes) sin remisión o con este mismo número
  const recs = leerRecargas().filter(function(x) {
    return normalizar(x.sede) === normalizar(sede.nombre) && x.t <= tMs + 2 * 3600000 && tMs - x.t <= 48 * 3600000 &&
      (x.estado === 'Falta remisión' || (d.numero && x.remision === d.numero));
  }).sort(function(a, b) { return b.t - a.t; });
  const rec = recs[0];
  // La fecha va a mano: si la IA lee un año que no cuadra (2024 en vez de 2026), se usa el año de la recarga.
  if (d.fecha && rec && Math.abs(d.fecha.getTime() - rec.t) > 60 * 86400000) {
    const ajustada = new Date(new Date(rec.t).getFullYear(), d.fecha.getMonth(), d.fecha.getDate(), 12);
    d.fecha = Math.abs(ajustada.getTime() - rec.t) <= 60 * 86400000 ? ajustada : null;
  }
  const leido = [];
  if (d.numero) leido.push('N.º remisión: <b>' + d.numero + '</b>' + (d.fecha ? ' · ' + d.fecha.getDate() + ' ' + MESES_CORTOS[d.fecha.getMonth()] + ' ' + d.fecha.getFullYear() : ''));
  if (d.antes !== null || d.despues !== null) leido.push('Nivel: <b>' + (d.antes !== null ? num(d.antes) : '?') + ' → ' + (d.despues !== null ? num(d.despues) : '?') + ' pulg</b>' +
    (d.antes !== null && d.despues !== null ? ' (' + signoTxt(d.despues - d.antes) + ')' : ''));
  if (d.presIni !== null || d.presFin !== null) leido.push('Presión: ' + (d.presIni !== null ? num(d.presIni) : '?') + ' → ' + (d.presFin !== null ? num(d.presFin) : '?') + ' bar');

  if (!rec) {
    return responderEnTema(msg, cab + (leido.length ? leido.join('\n') + '\n' : '') +
      '⚠️ No encontré una recarga reciente de ' + sede.nombre + ' esperando remisión.\n' +
      '<b>Sube primero la foto del manómetro ya lleno</b> y luego vuelve a enviar la remisión.');
  }
  if (!d.numero) {
    const envN = responderEnTema(msg, cab + (leido.length ? leido.join('\n') + '\n' : '') +
      '⚠️ No pude leer el número de la remisión. Responde a este mensaje con el número, por ejemplo <b>1001702</b>.');
    if (envN && envN.ok) agregarIdBot(rec, envN.result.message_id);
    return envN;
  }

  const nombre = [msg.from && msg.from.first_name, msg.from && msg.from.last_name].filter(Boolean).join(' ');
  ponerEnRecarga(rec.fila, 'N.º remisión', d.numero);
  ponerEnRecarga(rec.fila, 'Estado', 'Completa');
  ponerEnRecarga(rec.fila, 'Registró', nombre);
  if (d.fecha) ponerEnRecarga(rec.fila, 'Fecha de la remisión', d.fecha);
  if (d.presIni !== null) ponerEnRecarga(rec.fila, 'Presión inicial (bar)', d.presIni);
  if (d.presFin !== null) ponerEnRecarga(rec.fila, 'Presión final (bar)', d.presFin);
  ponerChipRecarga(rec.fila, 'Foto de la remisión', archivo.getUrl());

  // El nivel de llegada de la remisión es más fiel que la última foto (el camión puede llegar horas después)
  const finalMan = rec.final === '' || rec.final === null ? null : Number(rec.final);
  const iniFoto = rec.inicial === '' || rec.inicial === null ? null : Number(rec.inicial);
  const notas = [];
  let ini = iniFoto, espera = '';
  if (d.antes !== null && finalMan !== null && d.antes < finalMan) {
    ini = d.antes;
    ponerEnRecarga(rec.fila, 'Inicial (pulg)', ini);
    if (iniFoto !== null && iniFoto - ini >= 0.5) {
      // Lo normal: el camión llega horas después de la última foto y el tanque siguió gastando
      const previa = ultimaLecturaAntes(sede, rec.t);
      const pf = previa ? partesFecha(previa.t) : null;
      espera = '\n🕒 <b>Mientras llegaba el camión, el tanque siguió gastando:</b>' +
        '\n• Última foto' + (pf ? ' (' + pf.dia + ' ' + MESES_CORTOS[pf.mes - 1] + ', ' + horaAmPm(pf.hora, pf.min) + ')' : '') + ': <b>' + num(iniFoto) + ' pulg</b>' +
        '\n• Cuando llegó el camión (remisión): <b>' + num(ini) + ' pulg</b>' +
        '\n• Gastó <b>' + num(iniFoto - ini) + ' pulg</b> esperando. Es normal; para calcular la recarga uso el nivel de la remisión.';
    } else if (iniFoto !== null && Math.abs(iniFoto - ini) >= 0.5) {
      notas.push('La última foto marcaba ' + num(iniFoto) + ' pulg y la remisión dice ' + num(ini) + '; uso el de la remisión.');
    }
  }
  if (ini !== null && finalMan !== null) ponerEnRecarga(rec.fila, 'Pulgadas recargadas', Math.round((finalMan - ini) * 10) / 10);
  if (d.despues !== null && finalMan !== null && Math.abs(d.despues - finalMan) > 3) {
    notas.push('La remisión dice ' + num(d.despues) + ' pulg después de llenar y el manómetro marcó ' + num(finalMan) + '; dejé el del manómetro.');
  }
  // Costo aproximado (tabla de conversión × precio del año)
  let costo = '';
  try {
    const m3 = ini !== null && finalMan !== null ? m3Recarga(sede.nombre, ini, finalMan) : null;
    const precio = precioDe(leerPrecios(), rec.t);
    if (m3 !== null) costo = '\n💰 Aprox. <b>' + miles(m3) + ' m³' + (precio ? ' · $' + miles(Math.round(m3 * precio / 1000) * 1000) : '') + '</b> (estimado; falta la factura)';
  } catch (e) { Logger.log('Costo en remisión: ' + e.message); }
  // El nivel ya va explicado abajo: aquí solo número, fecha y presión
  const env = responderEnTema(msg, cab + leido.filter(function(x) { return x.indexOf('Nivel:') !== 0; }).join('\n') + '\n\n' +
    '✅ <b>Recarga completa' + (ini !== null && finalMan !== null ? ': ' + num(ini) + ' → ' + num(finalMan) + ' pulg (' + signoTxt(finalMan - ini) + ')' : '') + '</b>' +
    (espera ? '\n' + espera : '') +
    (notas.length ? '\n<i>' + notas.join(' ') + '</i>' : '') +
    (costo ? '\n' + costo : '') +
    (d.confianza === 'baja' ? '\n⚠️ <i>La letra a mano no se veía muy clara; revisa los números.</i>' : '') +
    '\n\n<i>¿Leí algo mal? Responde a este mensaje así:</i> <code>remisión ' + d.numero + ' antes ' +
    num(ini !== null ? ini : 11) + ' después ' + num(finalMan !== null ? finalMan : 74) + '</code>');
  if (env && env.ok) agregarIdBot(rec, env.result.message_id);
  return env;
}

/** Última lectura buena de la sede antes de un momento (la foto anterior a la recarga). */
function ultimaLecturaAntes(sede, tMs) {
  const L = leerLecturas().filter(function(l) {
    return normalizar(l.sede) === normalizar(sede.nombre) && l.t < tMs - 60000 && l.nivel !== null && l.nivel !== '' &&
      (l.estado === 'Automática' || l.estado === 'Corregida');
  });
  return L.length ? L[L.length - 1] : null;
}

/** Pone un chip de Drive en una celda de la pestaña Recargas. */
function ponerChipRecarga(fila, nombre, url) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), h = ss.getSheetByName(HOJA_REC);
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(function(t) { return String(t).trim(); });
  const c = enc.indexOf(nombre);
  if (c >= 0) ponerChip(ss, h, fila, c + 1, url);
}

/** Se ejecuta a las 7:00 a. m.: avisa si falta la foto de algún tanque y recuerda las remisiones pendientes. */
function recordarRemisiones() {
  const grupo = PropertiesService.getScriptProperties().getProperty('GRUPO_ID');
  if (!grupo) return;
  const sedes = leerSedes();
  // Tanques sin foto en más de 24 horas
  const lecturas = leerLecturas();
  sedes.forEach(function(sede) {
    if (!sede.tema) return;
    const ult = lecturas.filter(function(l) { return normalizar(l.sede) === normalizar(sede.nombre); }).pop();
    if (!ult || Date.now() - ult.t < 24 * 3600000) return;
    telegram('sendMessage', {
      chat_id: grupo, message_thread_id: Number(sede.tema), parse_mode: 'HTML',
      text: '📷 <b>Falta la foto del tanque ' + sede.nombre + '.</b> La última llegó el ' + fechaCorta(partesFecha(ult.t)) +
        '. Sube una foto del manómetro de nivel en este tema.'
    });
  });
  leerRecargas().filter(function(r) { return r.estado === 'Falta remisión'; }).forEach(function(r) {
    const sede = sedes.filter(function(x) { return normalizar(x.nombre) === normalizar(r.sede); })[0];
    if (!sede || !sede.tema) return;
    const p = partesFecha(r.t);
    const env = telegram('sendMessage', {
      chat_id: grupo, message_thread_id: Number(sede.tema), parse_mode: 'HTML',
      text: '⏰ <b>Falta la remisión</b> de la recarga del <b>' + p.dia + ' ' + MESES_CORTOS[p.mes - 1] + '</b> (' +
        num(r.inicial) + ' → ' + num(r.final) + ' pulg) en ' + r.sede + '.\n' + PEDIR_REMISION
    });
    if (env && env.ok) agregarIdBot(r, env.result.message_id);
  });
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
/** Modelos "flash" disponibles para esta clave, del más nuevo al más viejo; al final los "flash-lite". */
function listaModelos() {
  const props = PropertiesService.getScriptProperties();
  const guardada = props.getProperty('GEMINI_MODELOS');
  if (guardada) return JSON.parse(guardada);
  const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', {
    muteHttpExceptions: true, headers: { 'x-goog-api-key': props.getProperty('GEMINI_KEY') }
  });
  if (r.getResponseCode() !== 200) throw new Error('No se pudo consultar la lista de modelos (código ' + r.getResponseCode() + ').');
  const nombres = (JSON.parse(r.getContentText()).models || []).filter(function(m) {
    const n = m.name.toLowerCase();
    return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0 && n.indexOf('flash') >= 0 &&
      !/image|tts|audio|live|preview|exp|thinking|embedding/.test(n);
  }).map(function(m) { return m.name.replace('models/', ''); });
  if (!nombres.length) throw new Error('No hay un modelo Gemini flash disponible para esta clave.');
  const ver = function(n) { const m = n.match(/gemini-(\d+(?:\.\d+)?)/); return m ? Number(m[1]) : 0; };
  const orden = function(a, b) { return ver(b) - ver(a) || a.length - b.length; };
  const lista = nombres.filter(function(n) { return n.indexOf('lite') < 0; }).sort(orden)
    .concat(nombres.filter(function(n) { return n.indexOf('lite') >= 0; }).sort(orden)).slice(0, 6);
  props.setProperty('GEMINI_MODELOS', JSON.stringify(lista));
  return lista;
}

function buscarModelo() {
  PropertiesService.getScriptProperties().deleteProperty('GEMINI_MODELOS');
  const m = listaModelos()[0];
  PropertiesService.getScriptProperties().setProperty('GEMINI_MODELO_AUTO', m);
  return m;
}

/** La IA quedó en pausa por límite de uso: hasta cuándo (ms). */
function iaEnPausa() {
  return Number(PropertiesService.getScriptProperties().getProperty('IA_PAUSA_HASTA') || 0) > Date.now();
}

function leerManometro(blob, sede) {
  const desc = DESCRIPCION_MANOMETRO[normalizar(sede.nombre)] ||
    ('Manómetro de nivel con escala de 0 a ' + sede.escala + ' pulgadas de agua.');
  const prompt =
    'Eres un técnico que lee manómetros de nivel de tanques criogénicos de oxígeno medicinal.\n' +
    'Tanque de la sede ' + sede.nombre + '. ' + desc + '\n' +
    'NO calcules la lectura de memoria: necesito que UBIQUES PUNTOS en la foto, y yo hago la cuenta.\n' +
    'Coordenadas: [y, x] normalizadas de 0 a 1000 sobre la foto completa (0,0 = esquina superior izquierda).\n' +
    '1) "centro": el eje de la aguja (el tornillo o círculo donde gira).\n' +
    '2) "punta": el extremo más lejano de la aguja principal (no el contrapeso del otro lado).\n' +
    '3) "marcas": TODAS las marcas largas de la escala que tengan número impreso y se vean. Para cada una, su "valor" ' +
    'y el "punto" exactamente sobre la RAYA de la escala (en el anillo de rayas), no sobre el número.\n' +
    '4) "lectura": tu estimación de lo que marca (solo para comparar).\n' +
    'Si la foto NO es un manómetro sino un documento (remisión o factura de Messer), responde solo {"es_manometro": false, "es_remision": true}.\n' +
    'Responde SOLO este JSON: {"es_manometro": true/false, "centro": [y, x], "punta": [y, x], ' +
    '"marcas": [{"valor": número, "punto": [y, x]}, ...], "lectura": número o null, ' +
    '"confianza": "alta" | "media" | "baja", "motivo": "frase corta en español"}.\n' +
    'Usa confianza "baja" si la aguja no se ve clara, hay reflejo sobre ella o no estás seguro de la escala.';
  const ia = llamarGemini(prompt, blob);
  if (ia.errorIA) return { es_manometro: true, lectura: null, confianza: 'baja', errorIA: true, motivo: ia.motivo };
  if (ia.es_remision === true && ia.es_manometro !== true) return ia;
  return aplicarGeometria(ia, blob, sede);
}

/**
 * Envía la foto y la instrucción a Gemini y devuelve el JSON que contesta.
 * Maneja modelo retirado (404), Google ocupado (500/503) y límite de uso (429).
 * Si falla, devuelve { errorIA: true, motivo }.
 */
function llamarGemini(prompt, blob) {
  const props = PropertiesService.getScriptProperties();
  let modelo = props.getProperty('GEMINI_MODELO') || props.getProperty('GEMINI_MODELO_AUTO');
  if (!modelo) { try { modelo = listaModelos()[0]; } catch (e) { modelo = MODELO_POR_DEFECTO; } }
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
    if (iaEnPausa()) {
      return { errorIA: true, motivo: IA_NO_DISPONIBLE + ' (límite de uso, en pausa).' };
    }
    let r = pedir(modelo);
    if (r.getResponseCode() === 404) { // el modelo ya no existe: se busca el vigente y se reintenta
      modelo = buscarModelo();
      r = pedir(modelo);
    }
    // Google ocupado (500/503): un solo reintento corto
    if ([500, 503].indexOf(r.getResponseCode()) >= 0) {
      Utilities.sleep(5000);
      r = pedir(modelo);
    }
    // Límite de uso (429): probar los otros modelos, que tienen su propio cupo gratis
    if (r.getResponseCode() === 429) {
      const otros = listaModelos().filter(function(m) { return m !== modelo; });
      for (let i = 0; i < otros.length && r.getResponseCode() === 429; i++) {
        const r2 = pedir(otros[i]);
        if (r2.getResponseCode() === 200) { modelo = otros[i]; props.setProperty('GEMINI_MODELO_AUTO', modelo); }
        if (r2.getResponseCode() !== 404) r = r2;
      }
      if (r.getResponseCode() === 429) {
        // Todos agotados: pausa de 1 hora para no seguir gastando intentos
        props.setProperty('IA_PAUSA_HASTA', String(Date.now() + 60 * 60000));
      }
    }
    if (r.getResponseCode() !== 200) {
      Logger.log('Gemini ' + modelo + ' respondió ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 500));
      return { errorIA: true, motivo: IA_NO_DISPONIBLE + ' (código ' + r.getResponseCode() + ').' };
    }
    const j = JSON.parse(r.getContentText());
    const txt = j.candidates[0].content.parts.map(function(p) { return p.text || ''; }).join('');
    const ia = JSON.parse(txt.replace(/^```(json)?|```$/g, '').trim());
    return ia;
  } catch (e) {
    Logger.log('Error con Gemini: ' + e.message);
    return { errorIA: true, motivo: IA_NO_DISPONIBLE + ': ' + e.message };
  }
}

/** Si la lectura salió dudosa porque el cálculo por ángulo y la estimación de la IA no coinciden. */
function revisarDe(ia, val) {
  if (val.valida || !ia || ia.metodo !== 'geometria' || ia.lectura_ia === null || ia.lectura_ia === undefined || isNaN(ia.lectura_ia)) return null;
  return { calc: Number(ia.lectura), estim: Number(ia.lectura_ia) };
}

/**
 * Con los puntos que ubicó la IA se calcula la lectura por ángulo (Calculos.gs).
 * Si no alcanzan los puntos, queda la estimación de la IA tal cual.
 */
function aplicarGeometria(ia, blob, sede) {
  if (!ia || ia.es_manometro === false) return ia;
  const t = tamanoJpeg(blob.getBytes());
  const g = t ? lecturaPorGeometria(ia, t.ancho, t.alto, sede.escala) : null;
  const estim = ia.lectura === null || ia.lectura === undefined || ia.lectura === '' ? null : Number(ia.lectura);
  if (!g) {
    ia.motivo = (ia.motivo ? ia.motivo + ' · ' : '') + 'Estimación de la IA (no ubicó suficientes marcas).';
    return ia;
  }
  ia.lectura_ia = estim;
  ia.lectura = g.lectura;
  ia.metodo = 'geometria';
  let nota = 'Calculada por ángulo entre las marcas ' + g.entre[0] + ' y ' + g.entre[1] + ' (' + g.marcasUsadas + ' marcas).';
  if (estim !== null && !isNaN(estim)) {
    const dif = Math.abs(estim - g.lectura);
    if (dif >= 3) nota += ' La IA estimaba ' + num(estim) + '.';
    // Si el cálculo y la estimación se separan mucho, algo se ubicó mal: que lo revise una persona
    if (dif > sede.escala * 0.12) { ia.confianza = 'baja'; nota += ' Diferencia grande: revisar la foto.'; }
  }
  ia.motivo = nota;
  if (LUPA[normalizar(sede.nombre)] && t) aplicarLupa(ia, blob, sede, t);
  return ia;
}

/** Segunda mirada con la foto ampliada en la punta de la aguja: cuenta rayas desde el número más cercano. */
function aplicarLupa(ia, blob, sede, tam) {
  const cfg = LUPA[normalizar(sede.nombre)];
  try {
    const c = aPixeles(ia.centro, tam.ancho, tam.alto), p = aPixeles(ia.punta, tam.ancho, tam.alto);
    if (!c || !p) return;
    const r = Math.hypot(p.x - c.x, p.y - c.y);
    if (r < Math.min(tam.ancho, tam.alto) * 0.1) return;
    // Cuadro centrado un poco hacia adentro de la punta, de lado ~0,9 del largo de la aguja: entra la punta y el número más cercano
    const lado = r * 0.9, k = 0.92;
    const cx = c.x + (p.x - c.x) * k, cy = c.y + (p.y - c.y) * k;
    const recorte = recortarFoto(blob, tam.ancho, tam.alto, cx - lado / 2, cy - lado / 2, lado);
    if (!recorte) return;
    const prompt =
      'Esta imagen es un ACERCAMIENTO de un manómetro de nivel: se ve la punta de la aguja BLANCA sobre la escala.\n' +
      'La escala tiene números cada ' + cfg.cadaNumero + ' y una raya pequeña cada ' + cfg.paso + ' pulgadas; los números crecen en sentido horario.\n' +
      '1) "marca": el número impreso de la escala más cercano a la punta de la aguja (lee el número, por ejemplo 120).\n' +
      '2) "rayas": cuántas rayas pequeñas hay desde la raya de ese número hasta la aguja. Positivo si la aguja está del lado de los números ' +
      'mayores, negativo si está del lado de los menores. Usa decimales si la aguja queda entre dos rayas (por ejemplo 0.5).\n' +
      'Responde SOLO este JSON: {"marca": número, "rayas": número, "confianza": "alta" | "media" | "baja"}';
    const l = llamarGemini(prompt, recorte);
    if (l.errorIA || l.marca === null || l.marca === undefined || isNaN(Number(l.marca)) || isNaN(Number(l.rayas))) return;
    const lupa = Math.round((Number(l.marca) + Number(l.rayas) * cfg.paso) * 10) / 10;
    if (lupa < 0 || lupa > sede.escala) return;
    ia.lectura_lupa = lupa;
    const dif = Math.abs(lupa - Number(ia.lectura));
    if (dif <= LUPA_MAX_DIF) {
      if (dif >= 0.5) ia.motivo += ' Con lupa: ' + num(lupa) + ' (el cálculo daba ' + num(ia.lectura) + ').';
      ia.lectura = lupa;
      ia.metodo = 'lupa';
    } else {
      ia.motivo += ' La lupa leía ' + num(lupa) + '; se deja el cálculo.';
    }
  } catch (e) {
    Logger.log('Lupa: ' + e.message);
  }
}

/**
 * Recorta un cuadro de la foto (en píxeles) y lo devuelve ampliado como PNG.
 * Apps Script no recorta imágenes: se pone la foto en una diapositiva cuadrada de una presentación
 * de trabajo (la parte de afuera queda fuera de la hoja) y se pide su miniatura grande.
 */
function recortarFoto(blob, ancho, alto, x, y, lado) {
  if (typeof Slides === 'undefined') { Logger.log('Lupa: falta activar el servicio "Google Slides API".'); return null; }
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('LUPA_PRES'), pres = null;
  if (id) { try { pres = SlidesApp.openById(id); } catch (e) { pres = null; } }
  if (!pres) {
    const nueva = Slides.Presentations.create({ title: 'Lupa del bot de tanques (no borrar)',
      pageSize: { width: { magnitude: 400, unit: 'PT' }, height: { magnitude: 400, unit: 'PT' } } });
    id = nueva.presentationId;
    props.setProperty('LUPA_PRES', id);
    try { DriveApp.getFileById(id).moveTo(DriveApp.getFolderById(props.getProperty('CARPETA_ID'))); } catch (e) {}
    pres = SlidesApp.openById(id);
  }
  const hoja = pres.getSlides()[0];
  hoja.getPageElements().forEach(function(e) { e.remove(); });
  const esc = 400 / lado;
  const img = hoja.insertImage(blob);
  img.setLeft(-x * esc).setTop(-y * esc).setWidth(ancho * esc).setHeight(alto * esc);
  pres.saveAndClose();
  const th = Slides.Presentations.Pages.getThumbnail(id, hoja.getObjectId(),
    { 'thumbnailProperties.thumbnailSize': 'LARGE', 'thumbnailProperties.mimeType': 'PNG' });
  return UrlFetchApp.fetch(th.contentUrl).getBlob().setContentType('image/png');
}

/**
 * Ejecútala una vez desde el editor: pide el permiso nuevo (Presentaciones) y prueba la lupa
 * con la foto más reciente de Barranquilla. El resultado sale en el registro de ejecución. No guarda nada.
 */
function probarLupa() {
  const props = PropertiesService.getScriptProperties();
  const sede = leerSedes().filter(function(s) { return LUPA[normalizar(s.nombre)]; })[0];
  if (!sede) return Logger.log('Ninguna sede tiene lupa.');
  const it = DriveApp.getFolderById(props.getProperty('CARPETA_ID')).getFiles();
  let ultimo = null;
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf('Tanque_' + sede.nombre) === 0 && (!ultimo || f.getDateCreated() > ultimo.getDateCreated())) ultimo = f;
  }
  if (!ultimo) return Logger.log('No hay fotos de ' + sede.nombre + '.');
  const ia = leerManometro(ultimo.getBlob(), sede);
  Logger.log(ultimo.getName() + '\nLectura final: ' + ia.lectura + ' (método: ' + ia.metodo + ')\nLupa: ' + ia.lectura_lupa +
    '\nEstimación IA: ' + ia.lectura_ia + '\nNota: ' + ia.motivo);
}

/** Menú: muestra qué modelo de Gemini se está usando (lo busca si hace falta). */
function probarModelo() {
  const props = PropertiesService.getScriptProperties();
  let t;
  try { t = 'Modelos disponibles (en orden de uso):\n' + listaModelos().join('\n'); } catch (e) { t = 'Error: ' + e.message; }
  t += '\n\nEn uso ahora: ' + (props.getProperty('GEMINI_MODELO') || props.getProperty('GEMINI_MODELO_AUTO') || MODELO_POR_DEFECTO);
  if (iaEnPausa()) t += '\n\nIA en pausa por límite de uso hasta las ' + horaAmPm(partesFecha(Number(props.getProperty('IA_PAUSA_HASTA'))).hora, partesFecha(Number(props.getProperty('IA_PAUSA_HASTA'))).min) + '.';
  SpreadsheetApp.getUi().alert(t);
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

/** Agrega una lectura a la tabla Lecturas. */
function agregarFila(datos) { agregarFilaEn(HOJA_LEC, 'Lecturas', datos); }

/** Agrega una fila dentro de una tabla, por nombre de encabezado, con las fotos como chip de Drive. */
function agregarFilaEn(hoja, tabla, datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(hoja);
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(function(t) { return String(t).trim(); });
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    // 1) Por la API, dentro de la tabla y con el chip en la misma escritura
    try {
      const tablaId = idTabla(ss, tabla);
      const celdas = enc.map(function(n) { return celdaApi(n, datos[n]); });
      const pedido = { appendCells: { sheetId: h.getSheetId(), rows: [{ values: celdas }], fields: 'userEnteredValue,chipRuns' } };
      if (tablaId) pedido.appendCells.tableId = tablaId;
      Sheets.Spreadsheets.batchUpdate({ requests: [pedido] }, ss.getId());
      return;
    } catch (e) {
      Logger.log('appendCells falló, se usa el método simple: ' + e.message);
    }
    // 2) Método simple: fila normal + chip aparte (o enlace con texto corto)
    const fila = enc.map(function(n) { return COLUMNAS_CHIP.indexOf(n) >= 0 ? '' : (datos[n] === undefined ? '' : datos[n]); });
    h.appendRow(fila);
    const r = h.getLastRow();
    COLUMNAS_CHIP.forEach(function(n) {
      const c = enc.indexOf(n) + 1;
      if (c > 0 && datos[n]) ponerChip(ss, h, r, c, datos[n]);
    });
  } finally {
    lock.releaseLock();
  }
}

function celdaApi(nombre, v) {
  if (COLUMNAS_CHIP.indexOf(nombre) >= 0) {
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

/** Menú: recalcula %, cambio, movimiento y consumo de todas las filas (útil tras corregir lecturas a mano).
 *  Lee y escribe cada columna de una sola vez, para que sea rápido aunque haya miles de lecturas. */
function recalcularTodo() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_LEC);
  const n = h.getLastRow() - 1;
  if (n < 1) return actualizarResumen();
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(function(t) { return String(t).trim(); });
  const nombres = ['% lleno', 'Cambio (pulg)', 'Movimiento', 'Consumo por día (pulg)', 'Horas desde la anterior'];
  const col = {};
  nombres.forEach(function(nm) { const c = enc.indexOf(nm); if (c >= 0) col[nm] = { c: c + 1, v: h.getRange(2, c + 1, n, 1).getValues() }; });
  const poner = function(fila, nm, v) { if (col[nm]) col[nm].v[fila - 2][0] = (v === null || v === undefined) ? '' : v; };
  const lecturas = leerLecturas();
  leerSedes().forEach(function(sede) {
    const previas = [];
    lecturas.filter(function(l) { return normalizar(l.sede) === normalizar(sede.nombre); }).forEach(function(l) {
      if (!esValida(l)) return;
      const a = analizar({ t: l.t, nivel: l.nivel }, previas.slice(), recargasDe(previas), sede);
      poner(l.fila, '% lleno', a.pct);
      poner(l.fila, 'Cambio (pulg)', a.cambio);
      poner(l.fila, 'Movimiento', a.movimiento);
      poner(l.fila, 'Consumo por día (pulg)', a.porDia === null ? null : Math.round(a.porDia * 10) / 10);
      poner(l.fila, 'Horas desde la anterior', a.horas === null ? null : Math.round(a.horas * 10) / 10);
      previas.push({ t: l.t, nivel: l.nivel });
    });
  });
  Object.keys(col).forEach(function(nm) { h.getRange(2, col[nm].c, n, 1).setValues(col[nm].v); });
  actualizarResumen();
}

/** Al corregir a mano el nivel en Lecturas: la fila pasa a "Corregida" y todo se recalcula solo. */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const h = e.range.getSheet();
    if (h.getName() !== HOJA_LEC || e.range.getLastRow() < 2) return;
    const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(function(t) { return String(t).trim(); });
    const cNivel = enc.indexOf('Nivel (pulg)') + 1, cLec = enc.indexOf('Lectura') + 1;
    const toca = function(c) { return c >= e.range.getColumn() && c <= e.range.getLastColumn(); };
    if (!toca(cNivel) && !toca(cLec)) return;
    if (toca(cNivel) && cLec > 0) {
      const f0 = Math.max(2, e.range.getRow()), nf = e.range.getLastRow() - f0 + 1;
      const niv = h.getRange(f0, cNivel, nf, 1).getValues(), lec = h.getRange(f0, cLec, nf, 1).getValues();
      let cambio = false;
      for (let i = 0; i < nf; i++) {
        if (niv[i][0] !== '' && niv[i][0] !== null && lec[i][0] !== 'Corregida' && lec[i][0] !== 'Reemplazada') { lec[i][0] = 'Corregida'; cambio = true; }
      }
      if (cambio) h.getRange(f0, cLec, nf, 1).setValues(lec);
    }
    recalcularTodo();
    e.source.toast('Cálculos actualizados con la corrección.', 'Tanques', 4);
  } catch (err) {
    Logger.log('onEdit: ' + err.message);
  }
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
  const p48 = ult ? promedioConsumo(serie, ult.t, REGLAS.diasEstimar) : null;
  const reciente = ult ? ritmoReciente(serie, ult.t) : null;
  const rt = ritmoParaEstimar(p48 !== null ? p48 : p7, reciente);
  const dias = ult && rt.ritmo && rt.ritmo > 0.2 && ult.nivel > sede.aviso ? (ult.nivel - sede.aviso) / rt.ritmo : null;
  const lleno = ult && porcentaje(ult.nivel, sede.capacidad) > REGLAS.pctVerde;
  const estado = !ult ? 'Sin lecturas' : (ult.nivel <= sede.aviso ? 'Necesita recarga' : lleno ? 'Normal' :
    (dias !== null && dias < 1 ? 'Recarga urgente' : (dias !== null && dias <= REGLAS.diasPronto ? 'Recarga pronto' : 'Normal')));
  return {
    sede: sede.nombre, capacidad: sede.capacidad, escala: sede.escala, aviso: sede.aviso,
    nivel: ult ? ult.nivel : null, pct: ult ? porcentaje(ult.nivel, sede.capacidad) : null, ultimaT: ult ? ult.t : null,
    prom7: p7, prom30: p30, prom48: p48, recargaT: ultRec ? ultRec.t : null, recargaCambio: ultRec ? ultRec.cambio : null,
    dias: dias, estado: estado, reciente: reciente, ritmoAlto: rt.alto
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
  if (r.prom7 !== null) L.push('Promedio 7 días: ' + promTxt(r.prom7));
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
    if (d.accion === 'foto') return json(fotoTablero(String(d.id || ''), d.tam === 'mini'));
    // Administrador con huella
    const adm = { huellaEstado: function() { return { ok: true, nombre: huellaRegistrada_(d.credId) }; },
      huellaReto: function() { return reto_(d); }, huellaRegistrar: function() { return registrarHuellaAdmin_(d); },
      descartar: function() { return descartarLectura(d); }, restaurar: function() { return restaurarLectura(String(d.token || '')); },
      factura: function() { return guardarFactura(d); }, precios: function() { return guardarPrecios(d); } };
    if (adm[d.accion]) { try { return json(adm[d.accion]()); } catch (x) { return json({ ok: false, error: String(x.message || x) }); } }
    return json({ ok: false, error: 'Acción desconocida.' });
  } catch (err) {
    return json({ ok: false, error: 'Error en el servidor: ' + err.message });
  }
}

/**
 * Foto del manómetro para verla dentro del tablero, sin pedir cuenta de Google.
 * Solo entrega fotos que están en el registro de lecturas: el código del tablero
 * no da acceso a ningún otro archivo del Drive. mini = vista previa pequeña.
 */
function fotoTablero(id, mini) {
  if (!/^[\w-]{20,}$/.test(id)) return { ok: false, error: 'Foto no válida.' };
  const cache = CacheService.getScriptCache();
  let ids = cache.get('FOTOS_IDS');
  if (!ids || ids.indexOf('|' + id + '|') < 0) {
    ids = '|' + leerLecturas().map(function(l) { return l.fotoId; }).filter(String).join('|') + '|';
    try { cache.put('FOTOS_IDS', ids, 600); } catch (e) {} // si son demasiadas, se lee la hoja cada vez
  }
  if (ids.indexOf('|' + id + '|') < 0) return { ok: false, error: 'Foto no encontrada.' };
  const archivo = DriveApp.getFileById(id);
  const b = mini ? archivo.getThumbnail() : archivo.getBlob();
  if (!b) return { ok: true, vacia: true }; // Drive aún no genera la vista previa
  return { ok: true, tipo: b.getContentType() || 'image/jpeg', datos: Utilities.base64Encode(b.getBytes()) };
}

/**
 * Descarta una lectura dudosa desde el tablero. Solo el administrador, con su huella:
 * el reto se pidió para "descartar:<ID>" y la firma se comprueba aquí.
 * La fila no se borra: queda como "Descartada" (no se ve en el tablero) y la foto sigue en Drive.
 */
function descartarLectura(d) {
  const id = String(d.id || '');
  const v = verificar_(d); // lanza error si la huella no es válida o el reto venció
  if (v.para !== 'descartar:' + id) throw new Error('La huella se pidió para otra lectura.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (PropertiesService.getScriptProperties().getProperty('LISTAS_V') !== VERSION_LISTAS) actualizarListas();
    const l = leerLecturas().filter(function(x) { return x.id === id; })[0];
    if (!l) throw new Error('No encontré esa lectura.');
    if (l.estado !== 'Dudosa') throw new Error('Solo se pueden borrar lecturas dudosas.');
    const notaAntes = l.nota;
    const p = partesFecha(Date.now());
    ponerEn(HOJA_LEC, l.fila, 'Lectura', 'Descartada');
    ponerEn(HOJA_LEC, l.fila, 'Nota de la lectura', 'Descartada desde el tablero por ' + v.nombre + ' el ' + p.dia + ' ' +
      MESES_CORTOS[p.mes - 1] + ', ' + horaAmPm(p.hora, p.min) + (notaAntes ? ' · ' + notaAntes : ''));
    const token = Utilities.getUuid();
    CacheService.getScriptCache().put('DESHACER_' + token, JSON.stringify({ id: id, nota: notaAntes }), 300);
    return { ok: true, token: token, nombre: v.nombre };
  } finally {
    lock.releaseLock();
  }
}

/** "Deshacer" justo después de descartar (sirve 5 minutos, sin volver a pedir la huella). */
function restaurarLectura(token) {
  const c = CacheService.getScriptCache(), g = c.get('DESHACER_' + token);
  if (!g) throw new Error('Ya no se puede deshacer.');
  c.remove('DESHACER_' + token);
  const x = JSON.parse(g);
  const l = leerLecturas().filter(function(y) { return y.id === x.id; })[0];
  if (!l) throw new Error('No encontré esa lectura.');
  ponerEn(HOJA_LEC, l.fila, 'Lectura', 'Dudosa');
  ponerEn(HOJA_LEC, l.fila, 'Nota de la lectura', x.nota || '');
  return { ok: true };
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
    costos: costosTablero(),
    recargas: leerRecargas().filter(function(r) { return r.t >= desde; }).map(function(r) {
      return { t: r.t, sede: r.sede, inicial: r.inicial === '' ? null : Number(r.inicial), final: r.final === '' ? null : Number(r.final),
        cambio: r.cambio === '' ? null : Number(r.cambio), remision: r.remision, estado: r.estado, registro: r.registro };
    }),
    lecturas: lecturas.filter(function(l) { return l.t >= desde && l.estado !== 'Descartada'; }).map(function(l) {
      return { id: l.id, t: l.t, sede: l.sede, nivel: l.nivel, pct: l.pct, cambio: l.cambio, mov: l.movimiento, porDia: l.porDia,
        estado: l.estado, tecnico: l.tecnico, nota: l.nota,
        foto: l.fotoId ? 'https://drive.google.com/file/d/' + l.fotoId + '/view' : '' };
    })
  };
}
