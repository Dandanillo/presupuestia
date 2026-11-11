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

var VERSION = '2.3';

var COLS = {
  'Presupuestos': ['ID', 'Número', 'Fecha', 'Obra', 'Cliente', 'Dirección', 'Tipo', 'Estado', 'Margen %', 'Base imponible', 'IVA', 'Total', 'Autor', 'PDF', 'Modificado', 'Notas', 'IVA %'],
  'Partidas': ['ID presupuesto', 'Capítulo', 'Subcapítulo', 'Código', 'Descripción', 'Cantidad', 'Unidad', 'Precio unitario', 'Importe'],
  'Preciario': ['Descripción', 'Unidad', 'Precio', 'Capítulo', 'Zona', 'Autor', 'Modificado'],
  'Empresa': ['Clave', 'Valor'],
  'Usuarios': ['Usuario', 'Nombre', 'Correo', 'Teléfono', 'Dirección', 'Rol', 'Creado', 'Último acceso'],
  'Accesos': ['Fecha y hora', 'Usuario', 'Resultado', 'Dispositivo (lo declara el navegador)', 'Detalle'],
  'Plantillas': ['Plantilla', 'Capítulo', 'Subcapítulo', 'Código', 'Descripción', 'Cantidad', 'Unidad', 'Precio unitario', 'Autor', 'Modificado']
};
// Columnas (empezando en 1) que son texto plano: así Sheets no convierte fechas ni fórmulas
var TEXTO = {
  'Presupuestos': [1, 2, 3, 4, 5, 6, 7, 8, 13, 14, 15, 16],
  'Partidas': [1, 2, 3, 4, 5, 7],
  'Preciario': [1, 2, 4, 5, 6, 7],
  'Empresa': [1, 2],
  'Usuarios': [1, 2, 3, 4, 5, 6, 7, 8],
  'Accesos': [1, 2, 3, 4, 5],
  'Plantillas': [1, 2, 3, 4, 5, 7, 9, 10]
};
var PUBLICAS = ['auth.estado', 'auth.login'];
var MAX_FALLOS = 5;        // intentos fallidos seguidos…
var BLOQUEO_SEG = 900;     // …y se bloquea ese usuario 15 minutos
var ESTADOS = ['Borrador', 'Enviado', 'Aceptado', 'Rechazado'];
var CLAVES_EMPRESA = ['nombre', 'nif', 'dir', 'tel', 'email', 'web', 'validez', 'cond', 'logo', 'logoW', 'logoH'];
var MAX_LOGO = 45000;   // una celda de Google Sheets admite 50.000 caracteres

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
    var publica = PUBLICAS.indexOf(req.accion) >= 0;   // entrar con usuario y contraseña no necesita la clave de Drive
    if (!publica && !claveValida(req.clave)) return salida({ ok: false, error: 'La clave no es correcta' });
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
    case 'numero.siguiente': return numeroSiguiente(ss, req);
    case 'guardar': return guardar(ss, req);
    case 'subirPdf': return subirPdf(ss, req);
    case 'estado': return cambiarEstado(ss, req);
    case 'preciario.listar': return preciarioListar(ss);
    case 'preciario.guardar': return preciarioGuardar(ss, req);
    case 'empresa.leer': return empresaLeer(ss);
    case 'empresa.guardar': return empresaGuardar(ss, req);
    case 'auth.estado': return authEstado();
    case 'auth.login': return authLogin(ss, req);
    case 'auth.setup': return authSetup(ss, req);
    case 'usuarios.listar': return usuariosListar(ss, req);
    case 'usuarios.crear': return usuariosCrear(ss, req);
    case 'usuarios.borrar': return usuariosBorrar(ss, req);
    case 'usuarios.cambiarPass': return usuariosCambiarPass(ss, req);
    case 'plantillas.listar': return plantillasListar(ss);
    case 'plantillas.guardar': return plantillasGuardar(ss, req);
    case 'plantillas.borrar': return plantillasBorrar(ss, req);
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
  } else if (String(h.getRange(1, cols.length, 1, 1).getValues()[0][0]) !== cols[cols.length - 1]) {
    // hoja de una versión anterior: se ponen al día las cabeceras (por ejemplo, la columna nueva "IVA %")
    h.getRange(1, 1, 1, cols.length).setValues([cols]);
    try { h.getRange(1, 1, 1, cols.length).setFontWeight('bold'); } catch (e) {}
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
    iva: n(f[10]), total: n(f[11]), autor: sinGuardia(f[12]), pdf: sinGuardia(f[13]), modificado: sinGuardia(f[14]), notas: sinGuardia(f[15]),
    ivaPct: ivaValido(f[16], 21)   // las filas antiguas no tienen este dato: eran todas al 21 %
  };
}

function ivaValido(v, porDefecto) {
  var x = (v === '' || v === null || v === undefined) ? NaN : Number(v);
  return (x === 21 || x === 10 || x === 0) ? x : porDefecto;
}

// Número de presupuesto correlativo por año (P-2026-001, P-2026-002…), el mismo para todos los usuarios
function numeroSiguiente(ss, req) {
  var anio = parseInt(req.anio, 10);
  if (!(anio >= 2000 && anio <= 2100)) throw new Error('Año no válido');
  var props = PropertiesService.getScriptProperties();
  var clave = 'NUM_' + anio;
  var actual = parseInt(props.getProperty(clave), 10) || 0;
  // por si ya hay números puestos a mano o de antes, se parte del mayor que exista en la hoja
  var re = new RegExp('^P-' + anio + '-(\\d{1,6})$');
  leer(hoja(ss, 'Presupuestos'), 'Presupuestos').forEach(function (f) {
    var m = re.exec(sinGuardia(f[1]));
    if (m) actual = Math.max(actual, parseInt(m[1], 10));
  });
  actual += 1;
  props.setProperty(clave, String(actual));
  var num = String(actual);
  while (num.length < 3) num = '0' + num;
  return { ok: true, numero: 'P-' + anio + '-' + num };
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
    n(p.margen), n(p.base), n(p.iva), n(p.total), t(p.autor, 60), pdf, ahora(ss), t(p.notas, 2000), ivaValido(p.ivaPct, 21)
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
  var logo = String(d.logo || '');
  if (logo && (logo.length > MAX_LOGO || !/^data:image\/(png|jpeg);base64,[A-Za-z0-9+\/=]+$/.test(logo))) {
    throw new Error('El logo no es válido o es demasiado grande');
  }
  var filas = CLAVES_EMPRESA.map(function (k) {
    if (k === 'logo') return [k, logo];
    if (k === 'logoW' || k === 'logoH') return [k, logo ? String(Math.min(2000, Math.max(1, parseInt(d[k], 10) || 1))) : ''];
    return [k, t(d[k], k === 'cond' ? 800 : 150)];
  });
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

// ───────────────────────── Cuentas de usuario ─────────────────────────
// La contraseña NUNCA llega al script: la web la convierte antes en una huella lenta (PBKDF2, 100.000 vueltas)
// y aquí solo se guarda otra huella de esa huella, con sal. Las huellas viven en las propiedades del script
// (no en la hoja): quien abra la hoja no las ve. La pestaña Usuarios solo muestra datos de contacto.

function usuariosLeer() {
  var s = PropertiesService.getScriptProperties().getProperty('USUARIOS');
  try { return s ? JSON.parse(s) : {}; } catch (e) { return {}; }
}
function usuariosGuardar(o) { PropertiesService.getScriptProperties().setProperty('USUARIOS', JSON.stringify(o)); }

function usuarioNorm(u) {
  u = String(u || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,20}$/.test(u)) throw new Error('Usuario: de 3 a 20 letras, números, punto, guion o guion bajo');
  return u;
}
function huellaValida(ph) {
  ph = String(ph || '');
  if (!/^[0-9a-f]{64}$/.test(ph)) throw new Error('Contraseña no válida');
  return ph;
}
function hex(bytes) {
  return bytes.map(function (b) { return ('0' + ((b < 0 ? b + 256 : b)).toString(16)).slice(-2); }).join('');
}
function huellaServidor(salt, ph) {
  return hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + ph, Utilities.Charset.UTF_8));
}
function iguales(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  var d = 0;
  for (var i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}

function registrarAcceso(ss, usuario, resultado, req, detalle) {
  try {
    var h = hoja(ss, 'Accesos');
    var ua = t(req && req.dispositivo, 160);
    var fila = [[t(ahora(ss), 20), t(usuario, 40), t(resultado, 20), ua, t(detalle, 200)]];
    var r = h.getLastRow() + 1;
    if (r > h.getMaxRows()) h.insertRowsAfter(h.getMaxRows(), 100);
    var rg = h.getRange(r, 1, 1, 5);
    rg.setNumberFormats([['@', '@', '@', '@', '@']]);
    rg.setValues(fila);
    if (h.getLastRow() > 5000) h.deleteRows(2, 500);   // se guardan los últimos accesos
  } catch (e) { /* el registro nunca debe impedir entrar */ }
}

function espejoUsuarios(ss) {
  var us = usuariosLeer();
  var h = hoja(ss, 'Usuarios');
  var antes = Math.max(0, h.getLastRow() - 1);
  var filas = Object.keys(us).sort().map(function (k) {
    var u = us[k];
    // el teléfono ya está validado (solo cifras y + ( ) . - espacio), así que no necesita la guardia de fórmulas y conserva su +
    return [t(k, 20), t(u.nombre, 80), t(u.correo, 100), String(u.telefono || '').slice(0, 30), t(u.direccion, 150), u.rol === 'admin' ? 'Admin' : 'Usuario', t(u.creado, 20), t(u.ultimo, 20)];
  });
  escribirTodo(h, 'Usuarios', filas, antes);
}

function perfil(u) {
  return { nombre: u.nombre || '', correo: u.correo || '', telefono: u.telefono || '', direccion: u.direccion || '' };
}
function limpiarPerfil(p) {
  p = p || {};
  var correo = String(p.correo || '').trim().slice(0, 100);
  if (correo && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) throw new Error('El correo no parece válido');
  var tel = String(p.telefono || '').trim().slice(0, 30);
  if (tel && !/^[0-9 +().-]{6,30}$/.test(tel)) throw new Error('El teléfono no parece válido');
  return { nombre: String(p.nombre || '').trim().slice(0, 80), correo: correo, telefono: tel, direccion: String(p.direccion || '').trim().slice(0, 150) };
}

function authEstado() {
  return { ok: true, configurado: Object.keys(usuariosLeer()).length > 0 };
}

// Comprueba usuario + huella con bloqueo tras varios fallos. Devuelve la ficha del usuario o lanza error.
function verificar(ss, usuario, ph, req, accionLog) {
  var u = usuarioNorm(usuario);
  ph = huellaValida(ph);
  var cache = CacheService.getScriptCache();
  var k = 'fallos:' + u;
  var fallos = Number(cache.get(k) || 0);
  if (fallos >= MAX_FALLOS) {
    registrarAcceso(ss, u, 'Bloqueado', req, 'Demasiados intentos seguidos');
    throw new Error('Demasiados intentos. Espera 15 minutos y vuelve a probar');
  }
  var us = usuariosLeer();
  var f = us[u];
  var ok = !!f && iguales(huellaServidor(f.salt, ph), f.hash);
  if (!ok) {
    cache.put(k, String(fallos + 1), BLOQUEO_SEG);
    registrarAcceso(ss, u, 'Fallo', req, accionLog || 'Usuario o contraseña incorrectos');
    throw new Error('Usuario o contraseña incorrectos');
  }
  cache.remove(k);
  return { usuario: u, ficha: f, todos: us };
}

function authLogin(ss, req) {
  var v = verificar(ss, req.usuario, req.ph, req, '');
  v.ficha.ultimo = ahora(ss);
  v.todos[v.usuario] = v.ficha;
  usuariosGuardar(v.todos);
  try { espejoUsuarios(ss); } catch (e) {}
  registrarAcceso(ss, v.usuario, 'Correcto', req, '');
  return { ok: true, usuario: v.usuario, rol: v.ficha.rol === 'admin' ? 'admin' : 'user', perfil: perfil(v.ficha),
    clave: PropertiesService.getScriptProperties().getProperty('CLAVE') };
}

// La primera cuenta (administrador) solo se puede crear con la clave de Drive y mientras no haya ninguna
function authSetup(ss, req) {
  var us = usuariosLeer();
  if (Object.keys(us).length > 0) throw new Error('Ya hay cuentas creadas. Pide al administrador que cree la tuya');
  var u = usuarioNorm(req.usuario);
  var ph = huellaValida(req.ph);
  var salt = Utilities.getUuid().replace(/-/g, '');
  var ficha = limpiarPerfil(req.perfil);
  ficha.salt = salt; ficha.hash = huellaServidor(salt, ph); ficha.rol = 'admin'; ficha.creado = ahora(ss);
  us[u] = ficha;
  usuariosGuardar(us);
  espejoUsuarios(ss);
  registrarAcceso(ss, u, 'Alta', req, 'Primer administrador');
  return { ok: true };
}

function exigirAdmin(ss, req) {
  var a = req.admin || {};
  var v = verificar(ss, a.usuario, a.ph, req, 'Comprobación de administrador');
  if (v.ficha.rol !== 'admin') throw new Error('Solo un administrador puede hacer esto');
  return v;
}

function usuariosListar(ss, req) {
  var v = exigirAdmin(ss, req);
  return { ok: true, usuarios: Object.keys(v.todos).sort().map(function (k) {
    var f = v.todos[k];
    return { usuario: k, rol: f.rol === 'admin' ? 'admin' : 'user', nombre: f.nombre || '', correo: f.correo || '', telefono: f.telefono || '', direccion: f.direccion || '', ultimo: f.ultimo || '' };
  }) };
}

function usuariosCrear(ss, req) {
  var v = exigirAdmin(ss, req);
  var q = req.nuevo || {};
  var u = usuarioNorm(q.usuario);
  if (v.todos[u]) throw new Error('Ese usuario ya existe');
  var ph = huellaValida(q.ph);
  var ficha = limpiarPerfil(q);
  var salt = Utilities.getUuid().replace(/-/g, '');
  ficha.salt = salt; ficha.hash = huellaServidor(salt, ph); ficha.rol = q.rol === 'admin' ? 'admin' : 'user'; ficha.creado = ahora(ss);
  v.todos[u] = ficha;
  usuariosGuardar(v.todos);
  espejoUsuarios(ss);
  registrarAcceso(ss, v.usuario, 'Alta', req, 'Creó la cuenta de ' + u);
  return { ok: true };
}

function usuariosBorrar(ss, req) {
  var v = exigirAdmin(ss, req);
  var u = usuarioNorm(req.usuario);
  if (u === v.usuario) throw new Error('No puedes borrar tu propia cuenta');
  if (!v.todos[u]) throw new Error('Ese usuario no existe');
  delete v.todos[u];
  usuariosGuardar(v.todos);
  espejoUsuarios(ss);
  registrarAcceso(ss, v.usuario, 'Baja', req, 'Borró la cuenta de ' + u);
  return { ok: true };
}

function usuariosCambiarPass(ss, req) {
  var v = verificar(ss, req.usuario, req.ph, req, 'Cambio de contraseña');
  var nuevo = huellaValida(req.nuevoPh);
  var salt = Utilities.getUuid().replace(/-/g, '');
  v.ficha.salt = salt; v.ficha.hash = huellaServidor(salt, nuevo);
  v.todos[v.usuario] = v.ficha;
  usuariosGuardar(v.todos);
  registrarAcceso(ss, v.usuario, 'Cambio', req, 'Cambió su contraseña');
  return { ok: true };
}

// ───────────────────────── Plantillas de partidas ─────────────────────────
// Una plantilla es un grupo de partidas con nombre (por ejemplo "Baño completo").
// Cada fila de la pestaña es una partida; las filas con el mismo nombre forman una plantilla.

function plantillasListar(ss) {
  var filas = leer(hoja(ss, 'Plantillas'), 'Plantillas').filter(function (f) { return f[0] !== ''; });
  var orden = [], mapa = {};
  filas.forEach(function (f) {
    var nombre = sinGuardia(f[0]);
    var k = norm(nombre);
    if (!mapa[k]) { mapa[k] = { nombre: nombre, autor: sinGuardia(f[8]), modificado: sinGuardia(f[9]), partidas: [] }; orden.push(k); }
    mapa[k].partidas.push({ capitulo: sinGuardia(f[1]), subcapitulo: sinGuardia(f[2]), codigo: sinGuardia(f[3]), descripcion: sinGuardia(f[4]), cantidad: n(f[5]), unidad: sinGuardia(f[6]), precio: n(f[7]) });
  });
  return { ok: true, plantillas: orden.map(function (k) { return mapa[k]; }) };
}

function plantillasGuardar(ss, req) {
  var nombre = t(req.nombre, 80).trim();
  if (!nombre) throw new Error('La plantilla necesita un nombre');
  var partidas = req.partidas;
  if (!Array.isArray(partidas) || !partidas.length) throw new Error('La plantilla no tiene partidas');
  if (partidas.length > 300) throw new Error('Demasiadas partidas en una plantilla (máximo 300)');
  var h = hoja(ss, 'Plantillas');
  var filas = leer(h, 'Plantillas');
  var antes = filas.length;
  var k = norm(nombre);
  var resto = filas.filter(function (f) { return norm(sinGuardia(f[0])) !== k; });
  if (resto.length + partidas.length > 5000) throw new Error('Hay demasiadas plantillas guardadas. Borra alguna');
  var ts = ahora(ss), autor = t(req.autor, 60), nuevas = 0;
  partidas.forEach(function (q) {
    var desc = t(q.descripcion, 400);
    if (!desc.trim()) return;
    resto.push([nombre, t(q.capitulo, 100), t(q.subcapitulo, 100), t(q.codigo, 20), desc, n(q.cantidad), t(q.unidad, 20), n(q.precio), autor, ts]);
    nuevas++;
  });
  if (!nuevas) throw new Error('La plantilla no tiene partidas con descripción');
  escribirTodo(h, 'Plantillas', resto, antes);
  return { ok: true, partidas: nuevas };
}

function plantillasBorrar(ss, req) {
  var k = norm(t(req.nombre, 80));
  if (!k) throw new Error('Falta el nombre de la plantilla');
  var h = hoja(ss, 'Plantillas');
  var filas = leer(h, 'Plantillas');
  var resto = filas.filter(function (f) { return norm(sinGuardia(f[0])) !== k; });
  if (resto.length === filas.length) throw new Error('Esa plantilla no existe');
  escribirTodo(h, 'Plantillas', resto, filas.length);
  return { ok: true };
}
