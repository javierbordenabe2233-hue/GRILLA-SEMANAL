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
 * - Las grillas mensuales (PDF) quedan guardadas en tu Google Drive, en la
 *   carpeta "Grillas mensuales - App", y la lista en la pestaña APP_MENSUALES.
 *   Ver y descargar: cualquiera con la app. Subir y borrar: con la clave.
 */

var HOJA = 'APP_DATOS';
var HOJA_MEN = 'APP_MENSUALES';
var CARPETA = 'Grillas mensuales - App';
var MAX_MB = 15;

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
  { id: 'gmvzcg186', tipo: '7', desde: '2026-11-06', hasta: '2026-11-13',
    insp: ['ARGUELLO', 'PAGNONI', 'DIRASSAR', 'CORONEL', 'SOLIS', 'MEDINA', 'VALLEJOS'],
    sup: ['JEREZ', 'ORSI', 'BARRIA', 'ARROYO'], ajustes: {}, cambios: [{ d: 0, x: 'insp:6', y: 'insp:3' }],
    guardado: '2026-11-02T14:18:02.166Z' }
];

var COLUMNAS = ['ID', 'DESDE', 'HASTA', 'TIPO', 'GUARDADO', 'DATOS (no editar a mano)'];

/* ---------- lectura pública: la usa la app para Buscar y Grilla ---------- */
function doGet(e) {
  var pdf = e && e.parameter && e.parameter.pdf;
  if (pdf) return salida_(bajarMensual_(String(pdf)));
  prepararClave_();
  return salida_({
    semanas: leerSemanas_(),
    clave: { salt: PropertiesService.getScriptProperties().getProperty('SALT') },
    planillas: PLANILLAS,
    mensuales: leerMensuales_()
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

    if (b.accion === 'subirMensual') return salida_(subirMensual_(b));
    if (b.accion === 'borrarMensual') return salida_(borrarMensual_(String(b.id || '')));

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
      ajustes: (w.ajustes && typeof w.ajustes === 'object') ? w.ajustes : {},
      cambios: Array.isArray(w.cambios) ? w.cambios.filter(function (c) {
        return c && (c.d === 0 || c.d === 7) && typeof c.x === 'string' && typeof c.y === 'string';
      }).map(function (c) { return { d: c.d, x: c.x, y: c.y }; }) : [],
      guardado: String(w.guardado || '') };
  });
}

/* ---------- grillas mensuales (PDF en Google Drive) ---------- */
var COL_MEN = ['ID', 'MES', 'TÍTULO', 'ARCHIVO', 'TAMAÑO', 'SUBIDO', 'ID DRIVE'];

function hojaMen_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA_MEN);
  if (!sh) {
    sh = ss.insertSheet(HOJA_MEN);
    sh.getRange(1, 1, 1, COL_MEN.length).setValues([COL_MEN]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:G').setNumberFormat('@');
  }
  return sh;
}

function carpeta_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('CARPETA_ID');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var c = DriveApp.createFolder(CARPETA);
  props.setProperty('CARPETA_ID', c.getId());
  return c;
}

function filasMen_() {
  var sh = hojaMen_(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, COL_MEN.length).getValues().map(function (r, i) {
    return { fila: i + 2, id: String(r[0]), mes: String(r[1]), titulo: String(r[2]), nombre: String(r[3]),
      size: Number(r[4]) || 0, subido: String(r[5]), drive: String(r[6]) };
  }).filter(function (g) { return g.id && /^\d{4}-\d{2}$/.test(g.mes); });
}

function leerMensuales_() {
  return filasMen_().map(function (g) {
    return { id: g.id, mes: g.mes, titulo: g.titulo, nombre: g.nombre, size: g.size, subido: g.subido };
  });
}

function subirMensual_(b) {
  if (!/^\d{4}-\d{2}$/.test(String(b.mes)) || typeof b.b64 !== 'string' || !b.b64) return { ok: false, error: 'datos' };
  var bytes = Utilities.base64Decode(b.b64);
  if (bytes.length > MAX_MB * 1048576) return { ok: false, error: 'grande' };
  var nombre = String(b.nombre || ('Grilla mensual ' + b.mes + '.pdf')).slice(0, 200);
  var archivo = carpeta_().createFile(Utilities.newBlob(bytes, 'application/pdf', nombre));
  var id = 'm' + new Date().getTime().toString(36);
  hojaMen_().appendRow([id, String(b.mes), String(b.titulo || '').slice(0, 200), nombre, bytes.length,
    new Date().toISOString(), archivo.getId()]);
  return { ok: true, mensuales: leerMensuales_() };
}

function borrarMensual_(id) {
  var g = filasMen_().filter(function (x) { return x.id === id; })[0];
  if (!g) return { ok: true, mensuales: leerMensuales_() };
  try { DriveApp.getFileById(g.drive).setTrashed(true); } catch (e) {}
  hojaMen_().deleteRow(g.fila);
  return { ok: true, mensuales: leerMensuales_() };
}

function bajarMensual_(id) {
  var g = filasMen_().filter(function (x) { return x.id === id; })[0];
  if (!g) return { ok: false, error: 'no-existe' };
  try {
    return { ok: true, b64: Utilities.base64Encode(DriveApp.getFileById(g.drive).getBlob().getBytes()) };
  } catch (e) {
    return { ok: false, error: 'no-existe' };
  }
}

function salida_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* Ejecutalo una vez desde el editor (botón ▶ Ejecutar) para dar permisos y crear la pestaña APP_DATOS. */
function inicializar() {
  prepararClave_();
  hojaMen_();
  carpeta_();
  Logger.log('Listo: ' + leerSemanas_().length + ' grilla(s) en ' + HOJA + ', ' + leerMensuales_().length + ' grilla(s) mensuales.');
}
