/** @OnlyCurrentDoc */
/**
 * PresupuestIA · conexión con Google Drive (versión 1.0)
 *
 * Este script va DENTRO de la hoja de Google "Presupuestia – Datos"
 * (Extensiones → Apps Script). La guía paso a paso está en drive/LEEME.md.
 *
 * Qué hace:
 *  - Guarda cada presupuesto como una fila (pestaña Presupuestos) y sus partidas (pestaña Partidas).
 *  - Guarda el PDF en la carpeta "PDFs", junto a esta hoja.
 *  - Guarda el preciario (pestaña Preciario) y los datos de la empresa (pestaña Empresa).
 *  - Solo responde a quien conoce la CLAVE secreta (se genera al ejecutar "configurar").
 */

var VERSION = '1.1';

var COLS = {
  'Presupuestos': ['ID', 'Número', 'Fecha', 'Obra', 'Cliente', 'Dirección', 'Tipo', 'Estado', 'Margen %', 'Base imponible', 'IVA 21%', 'Total', 'Autor', 'PDF', 'Modificado', 'Notas'],
  'Partidas': ['ID presupuesto', 'Capítulo', 'Subcapítulo', 'Código', 'Descripción', 'Cantidad', 'Unidad', 'Precio unitario', 'Importe'],
  'Preciario': ['Descripción', 'Unidad', 'Precio', 'Capítulo', 'Zona', 'Autor', 'Modificado'],
  'Empresa': ['Clave', 'Valor']
};
// Columnas (empezando en 1) que son texto plano: así Sheets no convierte fechas ni fórmulas
var TEXTO = {
  'Presupuestos': [1, 2, 3, 4, 5, 6, 7, 8, 13, 14, 15, 16],
  'Partidas': [1, 2, 3, 4, 5, 7],
  'Preciario': [1, 2, 4, 5, 6, 7],
  'Empresa': [1, 2]
};
var ESTADOS = ['Borrador', 'Enviado', 'Aceptado', 'Rechazado'];
var CLAVES_EMPRESA = ['nombre', 'nif', 'dir', 'tel', 'email', 'web', 'validez', 'cond'];

// ───────────────────────── Preparación (se ejecuta una vez a mano) ─────────────────────────

function configurar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(COLS).forEach(function (nombre) { hoja(ss, nombre); });
  var props = PropertiesService.getScriptProperties();
  var clave = props.getProperty('CLAVE');
  if (!clave) {
    clave = claveNueva();
    props.setProperty('CLAVE', clave);
  }
  mostrarConexion(ss, clave);
  Logger.log('Listo. Abre la pestaña "Conexión" de la hoja: ahí está tu clave.');
}

// Si la clave se filtra, ejecuta esta función: crea una nueva y la anterior deja de valer.
function nuevaClave() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var clave = claveNueva();
  PropertiesService.getScriptProperties().setProperty('CLAVE', clave);
  mostrarConexion(ss, clave);
  Logger.log('Clave nueva creada. Cópiala de la pestaña "Conexión" y pégala en la web (Ajustes → Drive).');
}

function claveNueva() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

function mostrarConexion(ss, clave) {
  var h = ss.getSheetByName('Conexión') || ss.insertSheet('Conexión');
  var anterior = '';
  try { anterior = String(h.getRange(2, 2).getValue() || ''); } catch (e) {}
  h.clear();
  var filas = [
    ['Clave de conexión', clave],
    ['Dirección de la web app', anterior.indexOf('https://script.google.com/') === 0 ? anterior : 'Pega aquí la dirección que te da Google al implementar (empieza por https://script.google.com/ y acaba en /exec)'],
    ['', ''],
    ['Para qué sirve', 'En la web, Ajustes → Drive compartido: pega la dirección y la clave. Cada persona lo hace en su móvil u ordenador.'],
    ['Importante', 'Quien tenga la dirección y la clave puede leer y escribir en esta hoja. No las compartas fuera de vosotros.']
  ];
  h.getRange(1, 1, filas.length, 2).setNumberFormat('@');
  h.getRange(1, 1, filas.length, 2).setValues(filas);
  try { h.getRange(1, 1, filas.length, 1).setFontWeight('bold'); h.setColumnWidth(1, 190); h.setColumnWidth(2, 640); } catch (e) {}
}

// ───────────────────────── Puerta de entrada ─────────────────────────

function doGet() {
  return salida({ ok: true, mensaje: 'PresupuestIA: la conexión está activa. Se usa desde la web.' });
}

function doPost(e) {
  var res;
  try {
    var texto = (e && e.postData && e.postData.contents) || '';
    if (texto.length > 30000000) throw new Error('Petición demasiado grande');
    var req = JSON.parse(texto);
    if (!claveValida(req.clave)) return salida({ ok: false, error: 'La clave no es correcta' });
    var lock = LockService.getScriptLock();
    try { lock.waitLock(25000); } catch (x) { throw new Error('Drive está ocupado, inténtalo de nuevo en unos segundos'); }
    try { res = despachar(req); } finally { lock.releaseLock(); }
  } catch (err) {
    res = { ok: false, error: String((err && err.message) || err) };
  }
  return salida(res);
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function claveValida(c) {
  var k = PropertiesService.getScriptProperties().getProperty('CLAVE');
  if (!k || typeof c !== 'string' || c.length !== k.length) return false;
  var d = 0;
  for (var i = 0; i < k.length; i++) d |= k.charCodeAt(i) ^ c.charCodeAt(i);
  return d === 0;
}

function despachar(req) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  switch (req.accion) {
    case 'ping': return { ok: true, version: VERSION, hoja: ss.getName() };
    case 'listar': return listar(ss);
    case 'abrir': return abrir(ss, req);
    case 'guardar': return guardar(ss, req);
    case 'subirPdf': return subirPdf(ss, req);
    case 'estado': return cambiarEstado(ss, req);
    case 'preciario.listar': return preciarioListar(ss);
    case 'preciario.guardar': return preciarioGuardar(ss, req);
    case 'empresa.leer': return empresaLeer(ss);
    case 'empresa.guardar': return empresaGuardar(ss, req);
    case 'ia.leer': return iaLeer();
    case 'ia.guardar': return iaGuardar(req);
    default: throw new Error('Acción desconocida');
  }
}

// ───────────────────────── Utilidades ─────────────────────────

function hoja(ss, nombre) {
  var h = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  var cols = COLS[nombre];
  if (h.getLastRow() < 1) {
    h.getRange(1, 1, 1, cols.length).setValues([cols]);
    try { h.getRange(1, 1, 1, cols.length).setFontWeight('bold'); h.setFrozenRows(1); } catch (e) {}
    TEXTO[nombre].forEach(function (c) {
      var filas = Math.max(1, h.getMaxRows() - 1);
      h.getRange(2, c, filas, 1).setNumberFormat('@');
    });
  }
  return h;
}

// Texto seguro para una celda: sin caracteres de control y sin poder convertirse en fórmula
function t(v, max) {
  var s = (v === null || v === undefined ? '' : String(v)).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').slice(0, max || 500);
  if (/^[=+\-@]/.test(s)) s = ' ' + s;
  return s;
}
function sinGuardia(s) { return String(s === null || s === undefined ? '' : s).replace(/^ (?=[=+\-@])/, ''); }
function n(v) { var x = Number(v); return isFinite(x) ? x : 0; }
function norm(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim(); }
function ahora(ss) { return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm'); }

function formatos(nombre, filas) {
  var tx = TEXTO[nombre], cols = COLS[nombre].length, fila = [];
  for (var c = 1; c <= cols; c++) fila.push(tx.indexOf(c) >= 0 ? '@' : '#,##0.00');
  var out = [];
  for (var i = 0; i < filas; i++) out.push(fila);
  return out;
}

function leer(h, nombre) {
  var filas = h.getLastRow() - 1;
  if (filas < 1) return [];
  return h.getRange(2, 1, filas, COLS[nombre].length).getValues();
}

// Escribe todas las filas de la hoja (y borra lo que sobre por debajo)
function escribirTodo(h, nombre, filas, antes) {
  var cols = COLS[nombre].length;
  var falta = filas.length + 1 - h.getMaxRows();
  if (falta > 0) h.insertRowsAfter(h.getMaxRows(), falta + 100);
  if (antes > 0) h.getRange(2, 1, antes, cols).clearContent();
  if (filas.length > 0) {
    var r = h.getRange(2, 1, filas.length, cols);
    r.setNumberFormats(formatos(nombre, filas.length));
    r.setValues(filas);
  }
}

function filaPresupuesto(f) {
  return {
    id: sinGuardia(f[0]), numero: sinGuardia(f[1]), fecha: sinGuardia(f[2]), obra: sinGuardia(f[3]), cliente: sinGuardia(f[4]),
    direccion: sinGuardia(f[5]), tipo: sinGuardia(f[6]), estado: sinGuardia(f[7]) || 'Borrador', margen: n(f[8]), base: n(f[9]),
    iva: n(f[10]), total: n(f[11]), autor: sinGuardia(f[12]), pdf: sinGuardia(f[13]), modificado: sinGuardia(f[14]), notas: sinGuardia(f[15])
  };
}

function idValido(id) {
  id = String(id || '');
  if (!/^[A-Za-z0-9_-]{8,60}$/.test(id)) throw new Error('Identificador de presupuesto no válido');
  return id;
}

function buscarFila(filas, id) {
  for (var i = 0; i < filas.length; i++) if (String(filas[i][0]) === id) return i;
  return -1;
}

// ───────────────────────── Presupuestos ─────────────────────────

function listar(ss) {
  var h = hoja(ss, 'Presupuestos');
  var lista = leer(h, 'Presupuestos').filter(function (f) { return f[0] !== ''; }).map(filaPresupuesto);
  lista.sort(function (a, b) { return a.modificado < b.modificado ? 1 : (a.modificado > b.modificado ? -1 : 0); });
  return { ok: true, presupuestos: lista.slice(0, 300) };
}

function abrir(ss, req) {
  var id = idValido(req.id);
  var filas = leer(hoja(ss, 'Presupuestos'), 'Presupuestos');
  var i = buscarFila(filas, id);
  if (i < 0) throw new Error('No se encuentra ese presupuesto en Drive');
  var partidas = leer(hoja(ss, 'Partidas'), 'Partidas').filter(function (f) { return String(f[0]) === id; }).map(function (f) {
    return { capitulo: sinGuardia(f[1]), subcapitulo: sinGuardia(f[2]), codigo: sinGuardia(f[3]), descripcion: sinGuardia(f[4]), cantidad: n(f[5]), unidad: sinGuardia(f[6]), precio: n(f[7]) };
  });
  return { ok: true, presupuesto: filaPresupuesto(filas[i]), partidas: partidas };
}

function guardar(ss, req) {
  var p = req.presupuesto || {};
  var partidas = req.partidas;
  var id = idValido(p.id);
  if (!Array.isArray(partidas)) throw new Error('Faltan las partidas');
  if (partidas.length > 3000) throw new Error('Demasiadas partidas (máximo 3000)');

  var hp = hoja(ss, 'Presupuestos');
  var filas = leer(hp, 'Presupuestos');
  var i = buscarFila(filas, id);
  var estado = i >= 0 ? sinGuardia(filas[i][7]) : '';
  var pdf = i >= 0 ? sinGuardia(filas[i][13]) : '';
  var nueva = [
    id, t(p.numero, 40), t(p.fecha, 20), t(p.obra, 200), t(p.cliente, 200), t(p.direccion, 250), t(p.tipo, 100),
    ESTADOS.indexOf(estado) >= 0 ? estado : 'Borrador',
    n(p.margen), n(p.base), n(p.iva), n(p.total), t(p.autor, 60), pdf, ahora(ss), t(p.notas, 2000)
  ];
  if (i >= 0) filas[i] = nueva; else filas.push(nueva);
  escribirTodo(hp, 'Presupuestos', filas, filas.length - (i >= 0 ? 0 : 1));

  var hq = hoja(ss, 'Partidas');
  var todas = leer(hq, 'Partidas');
  var antes = todas.length;
  var resto = todas.filter(function (f) { return String(f[0]) !== id; });
  partidas.forEach(function (q) {
    resto.push([id, t(q.capitulo, 200), t(q.subcapitulo, 200), t(q.codigo, 30), t(q.descripcion, 1500), n(q.cantidad), t(q.unidad, 20), n(q.precio), '']);
  });
  escribirTodo(hq, 'Partidas', resto, antes);
  if (resto.length > 0) hq.getRange(2, 9, resto.length, 1).setFormulaR1C1('=RC[-3]*RC[-1]');
  return { ok: true, id: id, partidas: partidas.length };
}

function carpetaPdfs(ss) {
  var archivo = DriveApp.getFileById(ss.getId());
  var padres = archivo.getParents();
  var padre = padres.hasNext() ? padres.next() : DriveApp.getRootFolder();
  var it = padre.getFoldersByName('PDFs');
  return it.hasNext() ? it.next() : padre.createFolder('PDFs');
}

function subirPdf(ss, req) {
  var id = idValido(req.id);
  var nombre = String(req.nombre || 'presupuesto.pdf').replace(/[^A-Za-z0-9._-]/g, '_').slice(0, 80);
  if (!/\.pdf$/i.test(nombre)) nombre += '.pdf';
  var b64 = String(req.base64 || '');
  if (!b64 || b64.length > 20000000) throw new Error('PDF vacío o demasiado grande');
  var bytes = Utilities.base64Decode(b64);
  if (!(bytes.length > 4 && bytes[0] === 37 && bytes[1] === 80 && bytes[2] === 68 && bytes[3] === 70)) throw new Error('El archivo no es un PDF');

  var hp = hoja(ss, 'Presupuestos');
  var filas = leer(hp, 'Presupuestos');
  var i = buscarFila(filas, id);
  if (i < 0) throw new Error('Guarda primero el presupuesto en Drive');

  var archivo = carpetaPdfs(ss).createFile(Utilities.newBlob(bytes, 'application/pdf', nombre));
  var viejo = sinGuardia(filas[i][13]);
  var m = viejo.match(/\/d\/([-\w]{20,})/);
  if (m) { try { DriveApp.getFileById(m[1]).setTrashed(true); } catch (e) {} }
  var url = archivo.getUrl();
  hp.getRange(i + 2, 14, 1, 2).setValues([[url, ahora(ss)]]);
  return { ok: true, url: url };
}

function cambiarEstado(ss, req) {
  var id = idValido(req.id);
  var estado = String(req.estado || '');
  if (ESTADOS.indexOf(estado) < 0) throw new Error('Estado no válido');
  var hp = hoja(ss, 'Presupuestos');
  var filas = leer(hp, 'Presupuestos');
  var i = buscarFila(filas, id);
  if (i < 0) throw new Error('No se encuentra ese presupuesto en Drive');
  hp.getRange(i + 2, 8, 1, 1).setValues([[estado]]);
  hp.getRange(i + 2, 15, 1, 1).setValues([[ahora(ss)]]);
  return { ok: true };
}

// ───────────────────────── Preciario ─────────────────────────

function preciarioListar(ss) {
  var filas = leer(hoja(ss, 'Preciario'), 'Preciario').filter(function (f) { return f[0] !== ''; });
  return {
    ok: true,
    items: filas.map(function (f) {
      return { descripcion: sinGuardia(f[0]), unidad: sinGuardia(f[1]), precio: n(f[2]), capitulo: sinGuardia(f[3]), zona: sinGuardia(f[4]), autor: sinGuardia(f[5]), modificado: sinGuardia(f[6]) };
    })
  };
}

function preciarioGuardar(ss, req) {
  var items = req.items;
  if (!Array.isArray(items)) throw new Error('Faltan las partidas del preciario');
  if (items.length > 500) throw new Error('Demasiadas partidas a la vez (máximo 500)');
  var h = hoja(ss, 'Preciario');
  var filas = leer(h, 'Preciario');
  var antes = filas.length;
  var indice = {};
  filas.forEach(function (f, i) { indice[norm(f[0]) + '|' + norm(f[1]) + '|' + norm(f[4])] = i; });
  var nuevas = 0, actualizadas = 0, ts = ahora(ss);
  items.forEach(function (q) {
    var desc = t(q.descripcion, 300);
    var precio = n(q.precio);
    if (!desc.trim() || precio <= 0) return;
    var fila = [desc, t(q.unidad, 20), precio, t(q.capitulo, 100), t(q.zona, 60), t(q.autor, 60), ts];
    var k = norm(desc) + '|' + norm(fila[1]) + '|' + norm(fila[4]);
    if (indice[k] !== undefined) { filas[indice[k]] = fila; actualizadas++; }
    else { indice[k] = filas.length; filas.push(fila); nuevas++; }
  });
  escribirTodo(h, 'Preciario', filas, antes);
  return { ok: true, anadidas: nuevas, actualizadas: actualizadas };
}

// ───────────────────────── Datos de la empresa ─────────────────────────

function empresaLeer(ss) {
  var filas = leer(hoja(ss, 'Empresa'), 'Empresa');
  var datos = {};
  filas.forEach(function (f) {
    var k = sinGuardia(f[0]);
    if (CLAVES_EMPRESA.indexOf(k) >= 0) datos[k] = sinGuardia(f[1]);
  });
  return { ok: true, datos: datos };
}

function empresaGuardar(ss, req) {
  var d = req.datos || {};
  var h = hoja(ss, 'Empresa');
  var antes = leer(h, 'Empresa').length;
  var filas = CLAVES_EMPRESA.map(function (k) { return [k, t(d[k], k === 'cond' ? 800 : 150)]; });
  escribirTodo(h, 'Empresa', filas, antes);
  return { ok: true };
}

// ───────────────────────── Clave de Gemini compartida (opcional) ─────────────────────────
// Se guarda en las propiedades del script, NO en la hoja: no se ve al abrir la hoja,
// solo la devuelve el script a quien tenga la clave de acceso.

function iaLeer() {
  var p = PropertiesService.getScriptProperties();
  return { ok: true, key: p.getProperty('GEMINI_KEY') || '', modelo: p.getProperty('GEMINI_MODELO') || '' };
}

function iaGuardar(req) {
  var p = PropertiesService.getScriptProperties();
  var key = String(req.key || '').trim();
  var modelo = String(req.modelo || '').trim();
  if (!key) {
    p.deleteProperty('GEMINI_KEY');
    p.deleteProperty('GEMINI_MODELO');
    return { ok: true, compartida: false };
  }
  if (!/^[\x21-\x7E]{20,200}$/.test(key)) throw new Error('La clave de Gemini no parece válida');
  if (modelo && !/^[a-z0-9._-]{3,60}$/.test(modelo)) throw new Error('Nombre de modelo no válido');
  p.setProperty('GEMINI_KEY', key);
  if (modelo) p.setProperty('GEMINI_MODELO', modelo); else p.deleteProperty('GEMINI_MODELO');
  return { ok: true, compartida: true };
}
