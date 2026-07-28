// ═══════════════════════════════════════════════════════════
//  HIPOPOTAMO PINTURAS — Actualización automática Google Sheets
//  Versión corregida (28/07/2026)
//
//  QUÉ CAMBIA RESPECTO A LA VERSIÓN ANTERIOR:
//   (1) Valida el PERIODO de los datos ("Fecha desde / Fecha hasta" dentro
//       del Excel) contra el mes comercial actual. Si no coincide, NO vuelca
//       nada y avisa. Antes se volcaba cualquier adjunto y la cabecera
//       mostraba la fecha del CORREO, no la de los datos → parecía
//       actualizado mostrando el mes anterior.
//   (2) Elige entre TODOS los adjuntos candidatos (el ERP envía 2 correos
//       por día) el que casa con el mes comercial; empate → generación más
//       reciente. Antes cogía el primero que encontraba por posición.
//   (3) Escribe HPIN_LAST_RUN_DATE al terminar (antes el bloque estaba vacío
//       → yaEjecutadoHoy() nunca funcionaba y se ejecutaba ~20 veces/día,
//       convirtiendo un xlsx en Drive cada minuto → 403 rate limit).
//   (4) Cachea el análisis por mensaje: cada adjunto se convierte UNA vez.
//   (5) Ventana ampliada a 19:55–21:30 para tolerar retrasos del ERP.
//   (6) Cabecera con el periodo real de los datos y la hora de generación.
//   (7) Eliminada la doble declaración de actualizarDiaMes().
// ═══════════════════════════════════════════════════════════

// ── CONFIGURACIÓN ──────────────────────────────────────────
const CONFIG = {
  remitente: 'reportes@hipopotamo.com',
  asunto: 'Ventas por caja',
  nombreAdjunto: 'Ventas por caja',
  pestana: 'DATOS',
  buscarUltimasHoras: 30,
  maxCandidatos: 8,
  spreadsheetId: '1-J_e03Rue9yYg-jGBUfboFgDkhxTUCdLQAjADCLo4tA',

  // Hojas donde se escriben los días del mes comercial
  hojasDiaMes: ['Dia mes', 'Ventas MES'],
  celdaDiasTranscurridos: 'Q222',
  celdaDiasTotales: 'Q223',
  celdaResumen: 'P225',

  // Ventana de ejecución (Madrid). El correo llega ~20:00.
  ventanaInicio: 19 * 60 + 55,
  ventanaFin: 21 * 60 + 30,

  // Si true, cuando ningún adjunto casa con el mes comercial actual se aborta
  // sin escribir. Si false, vuelca el más reciente igualmente (no recomendado).
  exigirPeriodoCorrecto: true
};

const HPIN_TZ = 'Europe/Madrid';
const K_LAST_EXCEL_DATE = 'HPIN_LAST_EXCEL_DATE';   // yyyy-MM-dd de generación del Excel cargado
const K_LAST_RUN_DATE   = 'HPIN_LAST_RUN_DATE';     // yyyy-MM-dd de la última carga correcta
const K_LAST_PERIODO    = 'HPIN_LAST_PERIODO';      // 'yyyy-MM-dd..yyyy-MM-dd' del Excel cargado
const K_LAST_ALERT_DATE = 'HPIN_LAST_ALERT_DATE';   // para no spamear avisos
const K_CACHE_MSG       = 'HPIN_CACHE_MSGS';        // {msgId: {gen, desde, hasta}}

// Festivos Zaragoza (revisar cada año)
const FESTIVOS_ZARAGOZA = new Set([
  '2026-01-01', '2026-01-06', '2026-04-02', '2026-04-03', '2026-04-23',
  '2026-05-01', '2026-08-15', '2026-10-12', '2026-11-02', '2026-12-08', '2026-12-25'
]);

// ═══════════════════════════════════════════════════════════
//  FLUJO PRINCIPAL
// ═══════════════════════════════════════════════════════════
function actualizarDesdeGmail() {
  try {
    Logger.log('Inicio @ ' + ahoraMadridStr() + ' (' + HPIN_TZ + ')');

    if (!enVentanaHoraria()) { Logger.log('Fuera de ventana. No ejecuta.'); return; }
    if (yaEjecutadoHoy())    { Logger.log('Ya cargado hoy correctamente. No repite.'); return; }

    procesarCarga(false);
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
    Logger.log('STACK: ' + (e.stack || 'sin stack'));
    avisarUnaVezAlDia('⚠ Hipopotamo PIN: error en actualización',
      e.message + '\n\n' + (e.stack || ''));
  }
}

/**
 * Ejecución manual desde el menú: ignora ventana horaria y bloqueo diario.
 */
function forzarActualizacion() {
  Logger.log('=== FORZADO MANUAL ===');
  procesarCarga(true);
}

/**
 * Núcleo de la carga. `forzado` solo salta ventana/bloqueo, NUNCA la
 * validación de periodo: volcar el mes equivocado es peor que no volcar.
 */
function procesarCarga(forzado) {
  const hoyMadrid = parseYmd(Utilities.formatDate(new Date(), HPIN_TZ, 'yyyy-MM-dd'));
  const periodo = getPeriodoComercial(hoyMadrid);
  Logger.log('Mes comercial vigente: ' + ymd(periodo.inicio) + ' → ' + ymd(periodo.fin));

  const candidatos = analizarCandidatos();
  if (!candidatos.length) {
    avisarUnaVezAlDia('⚠ Hipopotamo PIN: no se encontró el Excel',
      'No hay correos recientes de ' + CONFIG.remitente + ' con adjunto "' +
      CONFIG.nombreAdjunto + '".');
    return;
  }

  const elegido = elegirPorPeriodo(candidatos, periodo);

  if (!elegido) {
    const detalle = candidatos.map(function (c) {
      return '· correo ' + fmt(c.fechaEmail, 'dd/MM/yyyy HH:mm:ss') +
             ' | generado ' + (c.gen ? fmt(c.gen, 'dd/MM/yyyy HH:mm:ss') : '?') +
             ' | periodo ' + (c.desde ? ymd(c.desde) : '?') + ' → ' + (c.hasta ? ymd(c.hasta) : '?');
    }).join('\n');

    Logger.log('Ningún adjunto corresponde al mes comercial vigente.\n' + detalle);

    if (CONFIG.exigirPeriodoCorrecto) {
      avisarUnaVezAlDia('⚠ Hipopotamo PIN: el ERP sigue enviando el mes anterior',
        'Mes comercial vigente: ' + ymd(periodo.inicio) + ' → ' + ymd(periodo.fin) + '\n\n' +
        'Adjuntos recibidos:\n' + detalle + '\n\n' +
        'No se ha volcado nada para no mostrar datos del periodo anterior.\n' +
        'Revisa el rango de fechas del informe "Ventas por caja" en el ERP.');
      return;
    }
    Logger.log('exigirPeriodoCorrecto=false → se vuelca el más reciente igualmente.');
  }

  const elegidoFinal = elegido || candidatos[0];

  // Dedup: si ya está cargado ese mismo periodo + generación, no reescribe.
  const firma = firmaCarga(elegidoFinal);
  const props = PropertiesService.getScriptProperties();
  if (!forzado && props.getProperty(K_LAST_PERIODO) === firma) {
    Logger.log('Ese Excel ya está cargado (' + firma + '). No repite.');
    props.setProperty(K_LAST_RUN_DATE, Utilities.formatDate(new Date(), HPIN_TZ, 'yyyy-MM-dd'));
    return;
  }

  const datos = leerExcel(elegidoFinal.blob);
  if (!datos || !datos.length) throw new Error('Excel leído pero sin filas útiles.');

  escribirEnSheets(datos, elegidoFinal);

  props.setProperty(K_LAST_PERIODO, firma);
  props.setProperty(K_LAST_RUN_DATE, Utilities.formatDate(new Date(), HPIN_TZ, 'yyyy-MM-dd'));

  Logger.log('✓ Carga completada. Filas=' + datos.length + ' | ' + firma);
}

// ═══════════════════════════════════════════════════════════
//  BÚSQUEDA Y ANÁLISIS DE CANDIDATOS
// ═══════════════════════════════════════════════════════════
/**
 * Devuelve todos los adjuntos candidatos con su periodo y fecha de generación
 * leídos DEL PROPIO EXCEL. El análisis se cachea por id de mensaje para no
 * convertir el mismo adjunto en cada ejecución.
 */
function analizarCandidatos() {
  const desde = new Date(new Date().getTime() - CONFIG.buscarUltimasHoras * 3600 * 1000);
  const query = 'from:' + CONFIG.remitente + ' subject:"' + CONFIG.asunto + '" after:' +
                Utilities.formatDate(desde, 'GMT', 'yyyy/MM/dd') + ' has:attachment';
  Logger.log('Query Gmail: ' + query);

  const threads = GmailApp.search(query, 0, 20);
  const cache = leerCache();
  const out = [];

  for (let t = 0; t < threads.length && out.length < CONFIG.maxCandidatos; t++) {
    const messages = threads[t].getMessages();
    for (let m = messages.length - 1; m >= 0 && out.length < CONFIG.maxCandidatos; m--) {
      const msg = messages[m];
      if (msg.getDate().getTime() < desde.getTime()) continue;

      const adjs = msg.getAttachments();
      for (let a = 0; a < adjs.length; a++) {
        const adj = adjs[a];
        const n = (adj.getName() || '').toLowerCase();
        if (!(n.endsWith('.xlsx') || n.endsWith('.xlsm'))) continue;
        if (n.indexOf(CONFIG.nombreAdjunto.toLowerCase()) === -1) continue;

        const key = msg.getId() + '#' + a;
        let meta = cache[key];

        if (!meta) {
          const cabecera = leerCabeceraExcel(adj.copyBlob());
          meta = {
            gen:   cabecera.gen   ? cabecera.gen.toISOString()   : null,
            desde: cabecera.desde ? ymd(cabecera.desde)          : null,
            hasta: cabecera.hasta ? ymd(cabecera.hasta)          : null
          };
          cache[key] = meta;
          Logger.log('Analizado ' + key + ' → generado ' + meta.gen +
                     ' | periodo ' + meta.desde + ' → ' + meta.hasta);
        }

        out.push({
          blob: adj.copyBlob(),
          nombre: adj.getName(),
          bytes: adj.getSize(),
          fechaEmail: msg.getDate(),
          gen:   meta.gen   ? new Date(meta.gen) : null,
          desde: meta.desde ? parseYmd(meta.desde) : null,
          hasta: meta.hasta ? parseYmd(meta.hasta) : null
        });
      }
    }
  }

  guardarCache(cache);
  Logger.log('Candidatos analizados: ' + out.length);
  return out;
}

/**
 * De los candidatos, el que corresponde al mes comercial vigente.
 * Criterio principal: "Fecha desde" == inicio del periodo.
 * Respaldo: "Fecha hasta" == fin del periodo.
 * Empate: el de generación más reciente.
 */
function elegirPorPeriodo(candidatos, periodo) {
  const ini = ymd(periodo.inicio);
  const fin = ymd(periodo.fin);

  const casan = candidatos.filter(function (c) {
    if (c.desde) return ymd(c.desde) === ini;
    if (c.hasta) return ymd(c.hasta) === fin;
    return false;
  });

  if (!casan.length) return null;

  casan.sort(function (a, b) {
    const ga = a.gen ? a.gen.getTime() : a.fechaEmail.getTime();
    const gb = b.gen ? b.gen.getTime() : b.fechaEmail.getTime();
    return gb - ga;
  });

  Logger.log('Adjunto elegido: generado ' +
    (casan[0].gen ? fmt(casan[0].gen, 'dd/MM/yyyy HH:mm:ss') : '?') +
    ' | periodo ' + ymd(casan[0].desde || periodo.inicio));
  return casan[0];
}

function firmaCarga(c) {
  return (c.desde ? ymd(c.desde) : '?') + '..' + (c.hasta ? ymd(c.hasta) : '?') +
         '@' + (c.gen ? c.gen.toISOString() : c.fechaEmail.toISOString());
}

// ═══════════════════════════════════════════════════════════
//  LECTURA DEL EXCEL
// ═══════════════════════════════════════════════════════════
/**
 * Convierte el xlsx a Sheets y devuelve el array 2D, recortando filas
 * totalmente vacías al final (la conversión sobre-cuenta getLastRow()).
 */
function leerExcel(blob) {
  const ssId = convertirAHojaTemporal(blob);
  try {
    const hoja = SpreadsheetApp.openById(ssId).getSheets()[0];
    const filas = hoja.getLastRow();
    const cols = hoja.getLastColumn();
    if (!filas || !cols) return [];

    const datos = hoja.getRange(1, 1, filas, cols).getValues();
    let quitadas = 0;
    while (datos.length && datos[datos.length - 1].every(esCeldaVacia)) { datos.pop(); quitadas++; }
    if (quitadas) Logger.log('Recortadas ' + quitadas + ' fila(s) vacía(s) final(es).');

    Logger.log('Leídas ' + filas + ' filas brutas; útiles ' + datos.length + ' (col=' + cols + ').');
    return datos;
  } finally {
    borrarTemporal(ssId);
  }
}

/**
 * Lee solo la cabecera del informe: fecha/hora de generación y periodo.
 * Formato real del ERP:
 *   fila 2 → "Fecha: 27/07/26 Hora: 20:00:01"
 *   fila 5 → "Fecha desde: 26/06/26   Fecha hasta: 25/07/26"
 */
function leerCabeceraExcel(blob) {
  const ssId = convertirAHojaTemporal(blob);
  try {
    const hoja = SpreadsheetApp.openById(ssId).getSheets()[0];
    const filas = Math.min(hoja.getLastRow(), 12);
    const cols = Math.min(hoja.getLastColumn(), 12);
    if (!filas || !cols) return {};

    const vals = hoja.getRange(1, 1, filas, cols).getValues();
    const texto = vals.map(function (r) {
      return r.map(function (c) {
        return (c instanceof Date) ? Utilities.formatDate(c, HPIN_TZ, 'dd/MM/yyyy HH:mm:ss') : String(c);
      }).join(' ');
    }).join('\n');

    return {
      gen:   parseFechaHora(texto),
      desde: parseFechaEtiquetada(texto, 'desde'),
      hasta: parseFechaEtiquetada(texto, 'hasta')
    };
  } finally {
    borrarTemporal(ssId);
  }
}

function convertirAHojaTemporal(blob) {
  blob.setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

  const temp = conReintentos(function () {
    const f = DriveApp.createFile(blob);
    f.setName('_temp_hipopotamo_erp.xlsx');
    return f;
  }, 4, 1200);

  const resp = conReintentos(function () {
    return UrlFetchApp.fetch('https://www.googleapis.com/drive/v2/files/' + temp.getId() + '/copy', {
      method: 'post',
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        title: '_temp_hipopotamo_erp',
        mimeType: 'application/vnd.google-apps.spreadsheet'
      }),
      muteHttpExceptions: true
    });
  }, 4, 1200);

  try { temp.setTrashed(true); } catch (e) { Logger.log('Aviso xlsx temporal: ' + e.message); }

  if (resp.getResponseCode() !== 200) {
    throw new Error('No se pudo convertir el Excel: ' + resp.getContentText());
  }
  return JSON.parse(resp.getContentText()).id;
}

function borrarTemporal(ssId) {
  try {
    conReintentos(function () { DriveApp.getFileById(ssId).setTrashed(true); }, 3, 800);
  } catch (e) {
    Logger.log('Aviso: no se pudo eliminar el temporal convertido: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════
//  ESCRITURA EN EL SHEET
// ═══════════════════════════════════════════════════════════
function escribirEnSheets(datos, elegido) {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  Logger.log('Destino: ' + ss.getName() + ' | ' + ss.getUrl());

  let hoja = ss.getSheetByName(CONFIG.pestana) || ss.insertSheet(CONFIG.pestana);
  hoja.clearContents();

  const refGen = elegido.gen || elegido.fechaEmail;
  const periodoTxt = (elegido.desde ? fmt(elegido.desde, 'dd/MM/yy') : '?') + ' – ' +
                     (elegido.hasta ? fmt(elegido.hasta, 'dd/MM/yy') : '?');

  const cabecera = [[
    'Última actualización:', fmt(new Date(), 'dd/MM/yyyy HH:mm'),
    'Generación informe ERP:', fmt(refGen, 'dd/MM/yyyy HH:mm'),
    'Periodo datos:', periodoTxt,
    'Filas:', datos.length
  ]];

  hoja.getRange(1, 1, 1, cabecera[0].length).setValues(cabecera)
    .setBackground('#1a56e8').setFontColor('#ffffff').setFontWeight('bold');

  hoja.getRange(2, 1, datos.length, datos[0].length).setValues(datos);
  for (let c = 1; c <= datos[0].length; c++) hoja.autoResizeColumn(c);

  // Referencia de 'Dia mes' = fecha de GENERACIÓN del informe, no la del correo.
  const refYmd = Utilities.formatDate(refGen, HPIN_TZ, 'yyyy-MM-dd');
  PropertiesService.getScriptProperties().setProperty(K_LAST_EXCEL_DATE, refYmd);

  actualizarDiaMesConReferencia(parseYmd(refYmd));
  SpreadsheetApp.flush();
  Logger.log('✓ DATOS + Dia mes escritos. Ref=' + refYmd + ' | periodo ' + periodoTxt);
}

function actualizarDiaMes() {
  actualizarDiaMesConReferencia(getFechaReferencia());
}

function actualizarDiaMesConReferencia(refDate) {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const periodo = getPeriodoComercial(refDate);
  const hasta = refDate < periodo.fin ? refDate : periodo.fin;

  const transcurridos = contarLaborablesLV(periodo.inicio, hasta, FESTIVOS_ZARAGOZA);
  const totales = contarLaborablesLV(periodo.inicio, periodo.fin, FESTIVOS_ZARAGOZA);
  const resumen = transcurridos + ' de ' + totales + ' días laborables';

  CONFIG.hojasDiaMes.forEach(function (nombre) {
    const sh = ss.getSheetByName(nombre);
    if (!sh) { Logger.log('Aviso: no existe la hoja "' + nombre + '". Se omite.'); return; }
    sh.getRange(CONFIG.celdaDiasTranscurridos).setValue(transcurridos);
    sh.getRange(CONFIG.celdaDiasTotales).setValue(totales);
    sh.getRange(CONFIG.celdaResumen).setValue(resumen);
    Logger.log('✓ "' + nombre + '" → ' + resumen);
  });

  SpreadsheetApp.flush();
}

// ═══════════════════════════════════════════════════════════
//  MES COMERCIAL Y FECHAS
// ═══════════════════════════════════════════════════════════
/** Mes comercial: del 26 al 25. */
function getPeriodoComercial(refDate) {
  const d = toUtcNoon(refDate);
  const y = d.getUTCFullYear(), m = d.getUTCMonth(), day = d.getUTCDate();
  return (day >= 26)
    ? { inicio: new Date(Date.UTC(y, m, 26, 12)),     fin: new Date(Date.UTC(y, m + 1, 25, 12)) }
    : { inicio: new Date(Date.UTC(y, m - 1, 26, 12)), fin: new Date(Date.UTC(y, m, 25, 12)) };
}

function contarLaborablesLV(desde, hasta, festivos) {
  const ini = toUtcNoon(desde), fin = toUtcNoon(hasta);
  if (ini > fin) return 0;

  let n = 0;
  const cur = new Date(ini.getTime());
  while (cur <= fin) {
    const dow = cur.getUTCDay();
    if (dow >= 1 && dow <= 5 && !festivos.has(ymd(cur))) n++;
    cur.setUTCDate(cur.getUTCDate() + 1);
  }
  return n;
}

function getFechaReferencia() {
  const raw = PropertiesService.getScriptProperties().getProperty(K_LAST_EXCEL_DATE);
  const p = raw ? parseYmd(raw) : null;
  return p || parseYmd(Utilities.formatDate(new Date(), HPIN_TZ, 'yyyy-MM-dd'));
}

/**
 * "Fecha: 27/07/26 Hora: 20:00:01" → Date.
 * La hora del informe es hora de Madrid, no UTC: hay que descontar el offset
 * o el instante se desplaza +1/+2 h (y con informes nocturnos saltaría de día,
 * corriendo un día de más en 'Dia mes').
 */
function parseFechaHora(texto) {
  const m = /Fecha:\s*(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s*Hora:\s*(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/i.exec(texto);
  if (!m) return null;
  let y = Number(m[3]); if (y < 100) y += 2000;
  return madridADate(y, Number(m[2]), Number(m[1]),
    Number(m[4] || 12), Number(m[5] || 0), Number(m[6] || 0));
}

/** Construye el instante correcto a partir de una fecha/hora expresada en Madrid. */
function madridADate(y, mo, d, hh, mi, ss) {
  const tentativa = new Date(Date.UTC(y, mo - 1, d, hh, mi, ss));
  const z = Utilities.formatDate(tentativa, HPIN_TZ, 'Z'); // '+0200' / '+0100'
  const signo = z.charAt(0) === '-' ? -1 : 1;
  const offMin = signo * (Number(z.substr(1, 2)) * 60 + Number(z.substr(3, 2)));
  return new Date(tentativa.getTime() - offMin * 60000);
}

/** "Fecha desde: 26/06/26" / "Fecha hasta: 25/07/26" → Date (mediodía UTC) */
function parseFechaEtiquetada(texto, etiqueta) {
  const re = new RegExp('Fecha\\s+' + etiqueta + '[:\\s]+(\\d{1,2})[\\/\\-](\\d{1,2})[\\/\\-](\\d{2,4})', 'i');
  const m = re.exec(texto);
  if (!m) return null;
  let y = Number(m[3]); if (y < 100) y += 2000;
  return new Date(Date.UTC(y, Number(m[2]) - 1, Number(m[1]), 12));
}

function parseYmd(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s || '')) return null;
  const p = s.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2], 12));
}

function toUtcNoon(d) { return parseYmd(ymd(d)); }
function ymd(d) { return Utilities.formatDate(d, HPIN_TZ, 'yyyy-MM-dd'); }
function fmt(d, pat) { return Utilities.formatDate(d, HPIN_TZ, pat); }
function ahoraMadridStr() { return fmt(new Date(), 'yyyy-MM-dd HH:mm:ss'); }
function esCeldaVacia(c) {
  return c === '' || c === null || c === undefined || (typeof c === 'string' && c.trim() === '');
}

// ═══════════════════════════════════════════════════════════
//  VENTANA, BLOQUEO DIARIO, CACHÉ Y AVISOS
// ═══════════════════════════════════════════════════════════
function enVentanaHoraria() {
  const h = Number(fmt(new Date(), 'H')), m = Number(fmt(new Date(), 'm'));
  const t = h * 60 + m;
  return t >= CONFIG.ventanaInicio && t <= CONFIG.ventanaFin;
}

function yaEjecutadoHoy() {
  return PropertiesService.getScriptProperties().getProperty(K_LAST_RUN_DATE) ===
         Utilities.formatDate(new Date(), HPIN_TZ, 'yyyy-MM-dd');
}

function resetEjecucionHoy() {
  const p = PropertiesService.getScriptProperties();
  p.deleteProperty(K_LAST_RUN_DATE);
  p.deleteProperty(K_LAST_PERIODO);
  Logger.log('✓ Bloqueo diario reseteado.');
}

function leerCache() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(K_CACHE_MSG) || '{}'); }
  catch (e) { return {}; }
}

function guardarCache(cache) {
  const keys = Object.keys(cache);
  if (keys.length > 20) keys.slice(0, keys.length - 20).forEach(function (k) { delete cache[k]; });
  PropertiesService.getScriptProperties().setProperty(K_CACHE_MSG, JSON.stringify(cache));
}

function avisarUnaVezAlDia(asunto, cuerpo) {
  const p = PropertiesService.getScriptProperties();
  const hoy = Utilities.formatDate(new Date(), HPIN_TZ, 'yyyy-MM-dd');
  if (p.getProperty(K_LAST_ALERT_DATE) === hoy) { Logger.log('Aviso ya enviado hoy.'); return; }
  try {
    GmailApp.sendEmail(Session.getActiveUser().getEmail(), asunto, cuerpo);
    p.setProperty(K_LAST_ALERT_DATE, hoy);
  } catch (e) {
    Logger.log('No se pudo enviar el aviso: ' + e.message);
  }
}

function conReintentos(fn, maxIntentos, esperaBase) {
  let ultimo = null;
  for (let i = 1; i <= maxIntentos; i++) {
    try { return fn(); }
    catch (e) {
      ultimo = e;
      Logger.log('Intento ' + i + '/' + maxIntentos + ' falló: ' + (e.message || e));
      if (i < maxIntentos) Utilities.sleep(esperaBase * i);
    }
  }
  throw ultimo;
}

// ═══════════════════════════════════════════════════════════
//  DIAGNÓSTICO Y TRIGGERS
// ═══════════════════════════════════════════════════════════
/**
 * EJECUTA ESTO PRIMERO. Lista todos los adjuntos recientes con su periodo
 * real, sin escribir nada. Responde a: ¿el ERP está enviando el mes nuevo?
 */
function diagnosticarAdjuntos() {
  PropertiesService.getScriptProperties().deleteProperty(K_CACHE_MSG);

  const hoy = parseYmd(Utilities.formatDate(new Date(), HPIN_TZ, 'yyyy-MM-dd'));
  const periodo = getPeriodoComercial(hoy);
  Logger.log('=== DIAGNÓSTICO ADJUNTOS ===');
  Logger.log('Mes comercial vigente: ' + ymd(periodo.inicio) + ' → ' + ymd(periodo.fin));

  const cands = analizarCandidatos();
  cands.forEach(function (c, i) {
    Logger.log('[' + i + '] correo ' + fmt(c.fechaEmail, 'dd/MM/yyyy HH:mm:ss') +
      ' | ' + c.bytes + ' bytes' +
      ' | generado ' + (c.gen ? fmt(c.gen, 'dd/MM/yyyy HH:mm:ss') : '?') +
      ' | periodo ' + (c.desde ? ymd(c.desde) : '?') + ' → ' + (c.hasta ? ymd(c.hasta) : '?') +
      ((c.desde && ymd(c.desde) === ymd(periodo.inicio)) ? '  ← CASA' : ''));
  });

  const elegido = elegirPorPeriodo(cands, periodo);
  Logger.log(elegido
    ? 'RESULTADO: hay adjunto del mes comercial vigente.'
    : 'RESULTADO: NINGÚN adjunto corresponde al mes vigente → el ERP sigue enviando el periodo anterior.');
}

function verEstadoSistema() {
  const p = PropertiesService.getScriptProperties();
  const triggers = ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'actualizarDesdeGmail'; });

  Logger.log('=== ESTADO ===');
  Logger.log('Ahora Madrid: ' + ahoraMadridStr());
  Logger.log('En ventana: ' + enVentanaHoraria());
  Logger.log(K_LAST_RUN_DATE + ': ' + (p.getProperty(K_LAST_RUN_DATE) || '(vacío)'));
  Logger.log(K_LAST_EXCEL_DATE + ': ' + (p.getProperty(K_LAST_EXCEL_DATE) || '(vacío)'));
  Logger.log(K_LAST_PERIODO + ': ' + (p.getProperty(K_LAST_PERIODO) || '(vacío)'));
  Logger.log('Triggers actualizarDesdeGmail: ' + triggers.length);
}

function crearActivador5Min() {
  eliminarTriggersActualizar();
  ScriptApp.newTrigger('actualizarDesdeGmail').timeBased().everyMinutes(5).create();
  Logger.log('✓ Trigger cada 5 min (filtro interno 19:55–21:30 Madrid + bloqueo diario).');
}

function eliminarTriggersActualizar() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'actualizarDesdeGmail') { ScriptApp.deleteTrigger(t); n++; }
  });
  Logger.log('✓ Triggers eliminados: ' + n);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🦛 Hipopotamo')
    .addItem('Actualizar ahora (forzar)', 'forzarActualizacion')
    .addItem('Diagnosticar adjuntos', 'diagnosticarAdjuntos')
    .addSeparator()
    .addItem('Ver estado sistema', 'verEstadoSistema')
    .addItem('Reset bloqueo diario', 'resetEjecucionHoy')
    .addSeparator()
    .addItem('Crear trigger 5 min', 'crearActivador5Min')
    .addItem('Eliminar triggers', 'eliminarTriggersActualizar')
    .addToUi();
}
