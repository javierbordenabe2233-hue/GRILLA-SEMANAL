/**
 * Apps Script de la app "Grillas Paso San Sebastián".
 * Se pega en la planilla GRILLAS: Extensiones > Apps Script.
 *
 * - Las grillas guardadas quedan en una pestaña nueva llamada APP_DATOS
 *   (una fila por semana). No toca las otras pestañas.
 * - La clave de carga es la MISMA que ya usás en la app actual.
 *   Se puede cambiar desde la app (Cargar > Cambiar clave).
 * - Los nombres para "Completar con los nombres de la planilla" están abajo,
 *   en PLANILLAS: así no quedan escritos en GitHub.
 */

var HOJA = 'APP_DATOS';

// Clave actual (la que ya creaste en la app). Solo se usa la primera vez.
var SALT_INICIAL = '1a244b54d74ed4a8c382e98eaf75e58a';
var HASH_INICIAL = '0ecae8728639407245847f693ba8af29e78063c1ba89427ab8fd22bd96ef32a7';

// Nombres de referencia por tipo de grilla (podés editarlos acá).
var PLANILLAS = {
  '6':  { insp: ['ARIAS', 'VALLEJOS', 'MEDINA', 'DIRASSAR', 'CORONEL', 'JEREZ T.'], sup: ['TERCERO', 'FENOGLIO', 'LEDESMA', 'JEREZ'] },
  '7':  { insp: ['ARIAS', 'VALLEJOS', 'MEDINA', 'DIRASSAR', 'JEREZ T.', 'CORONEL', 'SOLIS'], sup: ['LEDESMA', 'JEREZ', 'TERCERO', 'FENOGLIO'] },
  '7F': { insp: ['CHELI', 'ARGUELLO', 'ALDALLA', 'SANTILLAN', 'MEDINA', 'PAGNONI', 'LOBOS'], sup: ['BARRIA', 'ARROYO', 'LEDESMA', 'MANSILLA'] },
  '7V': { insp: ['ARGUELLO', 'ALVAREZ M.', 'JEREZ T.', 'CORONEL', 'ARIAS', 'CHELI', 'PAGNONI'], sup: ['JEREZ', 'FULGENZI', 'FENOGLIO', 'MANSILLA'] },
  '8':  { insp: ['ARGUELLO', 'LOBOS', 'PAGNONI', 'JEREZ T.', 'MEDINA', 'CORONEL', 'ARIAS', 'DIRASSAR'], sup: ['MANSILLA', 'LEDESMA', 'ARROYO', 'BARRIA'] }
};

// Grilla que ya tenías cargada; se copia sola la primera vez.
var SEMILLA = [
  { id: 'gmvyn4cpe', tipo: '7', desde: '2026-11-06', hasta: '2026-11-13',
    insp: ['ARGUELLO', 'PAGNONI', 'DIRASSAR', 'CORONEL', 'SOLIS', 'MEDINA', 'VALLEJOS'],
    sup: ['JEREZ', 'ORSI', 'BARRIA', 'ARROYO'], ajustes: {}, guardado: '2026-11-02T02:29:06.770Z' }
];

var COLUMNAS = ['ID', 'DESDE', 'HASTA', 'TIPO', 'GUARDADO', 'DATOS (no editar a mano)'];

/* ---------- lectura pública: la usa la app para Buscar y Grilla ---------- */
function doGet() {
  prepararClave_();
  return salida_({
    semanas: leerSemanas_(),
    clave: { salt: PropertiesService.getScriptProperties().getProperty('SALT') },
    planillas: PLANILLAS
  });
}

/* ---------- escritura: siempre con la clave ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var b = JSON.parse(e.postData.contents);
    prepararClave_();
    var props = PropertiesService.getScriptProperties();
    if (!b || b.h !== props.getProperty('HASH')) return salida_({ ok: false, error: 'clave' });

    if (b.accion === 'verificar') return salida_({ ok: true });

    if (b.accion === 'guardar') {
      var semanas = validarSemanas_(b.semanas);
      escribirSemanas_(semanas);
      return salida_({ ok: true, semanas: leerSemanas_() });
    }

    if (b.accion === 'clave') {
      if (!/^[0-9a-f]{32}$/.test(b.salt) || !/^[0-9a-f]{64}$/.test(b.nuevo)) return salida_({ ok: false, error: 'datos' });
      props.setProperties({ SALT: b.salt, HASH: b.nuevo });
      return salida_({ ok: true });
    }
    return salida_({ ok: false, error: 'accion' });
  } catch (err) {
    return salida_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- auxiliares ---------- */
function prepararClave_() {
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('HASH')) p.setProperties({ SALT: SALT_INICIAL, HASH: HASH_INICIAL });
}

function hoja_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.getRange(1, 1, 1, COLUMNAS.length).setValues([COLUMNAS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:F').setNumberFormat('@');
    if (SEMILLA.length) escribirEn_(sh, SEMILLA);
  }
  return sh;
}

function leerSemanas_() {
  var sh = hoja_(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 6, n, 1).getValues()
    .map(function (r) { try { return JSON.parse(r[0]); } catch (e) { return null; } })
    .filter(function (w) { return w && w.id; })
    .sort(function (a, b) { return String(b.desde).localeCompare(String(a.desde)); });
}

function escribirSemanas_(semanas) { escribirEn_(hoja_(), semanas); }

function escribirEn_(sh, semanas) {
  var n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, COLUMNAS.length).clearContent();
  if (!semanas.length) return;
  var filas = semanas.map(function (w) {
    return [w.id, w.desde, w.hasta, w.tipo, w.guardado || '', JSON.stringify(w)];
  });
  sh.getRange(2, 1, filas.length, COLUMNAS.length).setNumberFormat('@').setValues(filas);
}

function validarSemanas_(s) {
  if (!Array.isArray(s) || s.length > 1000) throw new Error('datos');
  var fecha = /^\d{4}-\d{2}-\d{2}$/;
  return s.map(function (w) {
    if (!w || typeof w.id !== 'string' || !fecha.test(w.desde) || !fecha.test(w.hasta) ||
        !PLANILLAS.hasOwnProperty(String(w.tipo)) || !Array.isArray(w.insp) || !Array.isArray(w.sup)) throw new Error('datos');
    return { id: w.id, tipo: String(w.tipo), desde: w.desde, hasta: w.hasta,
      insp: w.insp.map(String), sup: w.sup.map(String),
      ajustes: (w.ajustes && typeof w.ajustes === 'object') ? w.ajustes : {}, guardado: String(w.guardado || '') };
  });
}

function salida_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* Ejecutalo una vez desde el editor (botón ▶ Ejecutar) para dar permisos y crear la pestaña APP_DATOS. */
function inicializar() {
  prepararClave_();
  Logger.log('Listo: ' + leerSemanas_().length + ' grilla(s) en ' + HOJA);
}
