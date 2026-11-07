/**
 * PRESUPUESTIA · Script de Google (hoja de cálculo + Drive)
 *
 * Cómo se usa: se pega ENTERO en Extensiones → Apps Script de la hoja "Presupuestia"
 * y se publica como aplicación web (instrucciones en drive/INSTRUCCIONES.md).
 *
 * LO ÚNICO QUE TIENES QUE CAMBIAR: la clave (TOKEN) de la línea siguiente.
 * Inventa una clave larga (mínimo 16 caracteres, mejor 30) con letras y números.
 * La misma clave se pega después en la web (Ajustes → Conexión con Drive).
 */
const TOKEN = 'CAMBIA-ESTA-CLAVE-POR-UNA-LARGA';

// Carpeta de tu Drive donde se guardan los PDF y los datos de cada presupuesto
const NOMBRE_CARPETA = 'Presupuestia';

const VERSION = 1;

const COLUMNAS = {
  Presupuestos: ['ID', 'Número', 'Fecha', 'Obra', 'Cliente', 'Dirección', 'Tipo', 'Autor',
    'Base imponible', 'Margen %', 'IVA', 'Total', 'PDF', 'PDF (id)', 'Datos (id)', 'Guardado'],
  Partidas: ['Presupuesto', 'ID', 'Obra', 'Capítulo', 'Subcapítulo', 'Código', 'Descripción',
    'Cantidad', 'Unidad', 'Precio ud. (base)', 'Importe (con margen)'],
  Preciario: ['ID', 'Descripción', 'Unidad', 'Precio', 'Capítulo', 'Añadido por', 'Fecha'],
  Empresa: ['Clave', 'Valor']
};

const CLAVES_EMPRESA = ['nombre', 'nif', 'dir', 'tel', 'email', 'web', 'validez', 'cond'];

// ─────────────────────────── Entrada de la web ───────────────────────────

// Si abres la dirección del script en el navegador, solo verás este aviso (no enseña datos)
function doGet() {
  return ContentService.createTextOutput('Presupuestia: el script está activo ✓');
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return salida_({ ok: false, error: 'Petición no válida' });
  }
  if (TOKEN.indexOf('CAMBIA') === 0 || TOKEN.length < 16) {
    return salida_({ ok: false, error: 'Falta cambiar la clave (TOKEN) en el script' });
  }
  if (body.token !== TOKEN) {
    return salida_({ ok: false, error: 'Clave incorrecta' });
  }
  const accion = ACCIONES[body.action];
  if (!accion) {
    return salida_({ ok: false, error: 'Acción desconocida' });
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const resultado = accion(body);
    return salida_(Object.assign({ ok: true }, resultado));
  } catch (err) {
    return salida_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) { /* nada */ }
  }
}

function salida_(objeto) {
  return ContentService.createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────── Acciones ────────────────────────────────

const ACCIONES = {
  // Comprueba la conexión y devuelve los enlaces al Excel y a la carpeta
  ping: function () {
    asegurarHojas_();
    return {
      version: VERSION,
      hojaUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
      carpetaUrl: carpeta_().getUrl()
    };
  },

  // Guarda (o actualiza) un presupuesto: PDF + datos en Drive y filas en el Excel
  guardar: function (b) {
    const e = b.entry || {};
    const d = e.data;
    if (!d || !Array.isArray(d.capitulos)) throw new Error('El presupuesto está vacío');
    if (!b.pdfBase64) throw new Error('Falta el PDF');
    const pres = hoja_('Presupuestos');
    const part = hoja_('Partidas');
    const carpeta = carpeta_();

    let id = String(e.remoteId || '');
    let fila = id ? buscarFila_(pres, 1, id) : 0;
    const nombrePdf = nombreArchivo_(b.pdfName || ('Presupuesto_' + (e.numero || 'sin_numero') + '.pdf'));

    if (fila) {
      // Es una nueva versión de uno ya guardado: los archivos viejos van a la papelera
      const viejo = pres.getRange(fila, 1, 1, COLUMNAS.Presupuestos.length).getValues()[0];
      papelera_(viejo[13]);
      papelera_(viejo[14]);
    } else {
      id = 'p' + Utilities.getUuid().replace(/-/g, '').slice(0, 7);
    }

    const pdf = carpeta.createFile(
      Utilities.newBlob(Utilities.base64Decode(b.pdfBase64), 'application/pdf', nombrePdf));
    const json = carpeta.createFile(
      nombrePdf.replace(/\.pdf$/i, '') + '_' + id + '.json', JSON.stringify(d), 'application/json');

    const linea = [id, txt_(e.numero), txt_(e.fecha), txt_(e.obra), txt_(e.cliente), txt_(e.direccion),
      txt_(e.tipo), txt_(e.autor), num_(e.base), num_(e.margin), num_(e.iva), num_(e.total),
      pdf.getUrl(), pdf.getId(), json.getId(), new Date()];
    if (fila) {
      pres.getRange(fila, 1, 1, linea.length).setValues([linea]);
    } else {
      pres.getRange(pres.getLastRow() + 1, 1, 1, linea.length).setValues([linea]);
    }

    // Una fila por partida en la hoja "Partidas" (el "Excel" para analizar)
    quitarFilas_(part, 2, id);
    const margen = num_(e.margin);
    const filas = [];
    (d.capitulos || []).forEach(function (cap, ci) {
      (cap.subcapitulos || []).forEach(function (sub) {
        (sub.partidas || []).forEach(function (p) {
          if (p._del) return;
          const cant = num_(p.cantidad);
          const precio = num_(p.precio_unitario);
          const importe = redondear_(cant * redondear_(precio * (1 + margen / 100)));
          filas.push([txt_(e.numero), id, txt_(e.obra), txt_((ci + 1) + '. ' + (cap.nombre || '')),
            txt_(sub.nombre), txt_(p.codigo), txt_(p.descripcion), cant, txt_(p.unidad), precio, importe]);
        });
      });
    });
    if (filas.length) {
      part.getRange(part.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);
    }
    return { id: id, pdfUrl: pdf.getUrl() };
  },

  // Lista de presupuestos guardados (el más reciente primero)
  listar: function () {
    const sh = hoja_('Presupuestos');
    const n = sh.getLastRow();
    if (n < 2) return { items: [] };
    const v = sh.getRange(2, 1, n - 1, COLUMNAS.Presupuestos.length).getValues();
    return { items: v.map(filaPresupuesto_).reverse() };
  },

  // Devuelve los datos completos de uno para poder abrirlo y editarlo
  obtener: function (b) {
    const sh = hoja_('Presupuestos');
    const fila = buscarFila_(sh, 1, String(b.id || ''));
    if (!fila) throw new Error('No se encuentra ese presupuesto');
    const r = sh.getRange(fila, 1, 1, COLUMNAS.Presupuestos.length).getValues()[0];
    const data = JSON.parse(DriveApp.getFileById(r[14]).getBlob().getDataAsString());
    return { data: data, entry: filaPresupuesto_(r) };
  },

  eliminar: function (b) {
    const sh = hoja_('Presupuestos');
    const id = String(b.id || '');
    const fila = buscarFila_(sh, 1, id);
    if (!fila) throw new Error('No se encuentra ese presupuesto');
    const r = sh.getRange(fila, 1, 1, COLUMNAS.Presupuestos.length).getValues()[0];
    papelera_(r[13]);
    papelera_(r[14]);
    quitarFilas_(hoja_('Partidas'), 2, id);
    sh.deleteRow(fila);
    return {};
  },

  // ── Preciario compartido ──
  preciarioListar: function () {
    const sh = hoja_('Preciario');
    const n = sh.getLastRow();
    if (n < 2) return { items: [] };
    const v = sh.getRange(2, 1, n - 1, COLUMNAS.Preciario.length).getValues();
    return {
      items: v.map(function (r) {
        return { id: String(r[0]), desc: String(r[1]), unit: String(r[2]), price: num_(r[3]),
          cap: String(r[4]), autor: String(r[5]), fecha: fechaTexto_(r[6]) };
      })
    };
  },

  // Añade partidas al preciario; si ya existe la misma (descripción + unidad), actualiza su precio
  preciarioAnadir: function (b) {
    const items = Array.isArray(b.items) ? b.items : [];
    if (!items.length) throw new Error('No hay partidas que añadir');
    const sh = hoja_('Preciario');
    const n = sh.getLastRow();
    const existentes = n > 1 ? sh.getRange(2, 1, n - 1, COLUMNAS.Preciario.length).getValues() : [];
    const indice = {};
    existentes.forEach(function (r, i) { indice[clave_(r[1], r[2])] = i + 2; });
    const hoy = hoyTexto_();
    let nuevas = 0;
    let actualizadas = 0;
    items.forEach(function (it) {
      const desc = String(it.desc || '').trim();
      if (!desc) return;
      const k = clave_(desc, it.unit);
      const fila = indice[k];
      if (fila) {
        sh.getRange(fila, 4).setValue(num_(it.price));
        sh.getRange(fila, 7).setValue(txt_(hoy));
        actualizadas++;
      } else {
        const nueva = ['p' + Utilities.getUuid().replace(/-/g, '').slice(0, 7), txt_(desc), txt_(it.unit),
          num_(it.price), txt_(it.cap), txt_(it.autor), txt_(hoy)];
        const destino = sh.getLastRow() + 1;
        sh.getRange(destino, 1, 1, nueva.length).setValues([nueva]);
        indice[k] = destino;
        nuevas++;
      }
    });
    return { nuevas: nuevas, actualizadas: actualizadas };
  },

  preciarioBorrar: function (b) {
    const sh = hoja_('Preciario');
    const fila = buscarFila_(sh, 1, String(b.id || ''));
    if (!fila) throw new Error('No se encuentra esa partida');
    sh.deleteRow(fila);
    return {};
  },

  // ── Datos de la empresa (compartidos) ──
  empresaLeer: function () {
    const sh = hoja_('Empresa');
    const n = sh.getLastRow();
    const config = {};
    if (n > 1) {
      sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (r) {
        if (CLAVES_EMPRESA.indexOf(String(r[0])) >= 0) config[String(r[0])] = String(r[1]);
      });
    }
    return { config: config };
  },

  empresaGuardar: function (b) {
    const sh = hoja_('Empresa');
    const cfg = b.config || {};
    CLAVES_EMPRESA.forEach(function (k) {
      if (cfg[k] === undefined) return;
      const fila = buscarFila_(sh, 1, k);
      const valor = txt_(String(cfg[k]).slice(0, 800));
      if (fila) {
        sh.getRange(fila, 2).setValue(valor);
      } else {
        sh.getRange(sh.getLastRow() + 1, 1, 1, 2).setValues([[k, valor]]);
      }
    });
    return {};
  }
};

// ─────────────────────────────── Utilidades ──────────────────────────────

function hoja_(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    const cols = COLUMNAS[nombre];
    sh.getRange(1, 1, 1, cols.length).setValues([cols])
      .setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function asegurarHojas_() {
  Object.keys(COLUMNAS).forEach(hoja_);
}

function carpeta_() {
  const it = DriveApp.getFoldersByName(NOMBRE_CARPETA);
  return it.hasNext() ? it.next() : DriveApp.createFolder(NOMBRE_CARPETA);
}

function papelera_(idArchivo) {
  if (!idArchivo) return;
  try { DriveApp.getFileById(String(idArchivo)).setTrashed(true); } catch (err) { /* ya no existe */ }
}

// Número de fila (1 = cabecera) cuya columna "col" vale "valor"; 0 si no está
function buscarFila_(sh, col, valor) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const v = sh.getRange(2, col, n - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]) === valor) return i + 2;
  }
  return 0;
}

// Borra todas las filas cuya columna "col" vale "valor" (recorriendo de abajo arriba)
function quitarFilas_(sh, col, valor) {
  const n = sh.getLastRow();
  if (n < 2) return;
  const v = sh.getRange(2, col, n - 1, 1).getValues();
  let fin = -1;
  for (let i = v.length - 1; i >= 0; i--) {
    if (String(v[i][0]) === valor) {
      if (fin < 0) fin = i;
    } else if (fin >= 0) {
      sh.deleteRows(i + 3, fin - i);
      fin = -1;
    }
  }
  if (fin >= 0) sh.deleteRows(2, fin + 1);
}

function filaPresupuesto_(r) {
  return {
    id: String(r[0]), numero: String(r[1]), fecha: fechaTexto_(r[2]), obra: String(r[3]),
    cliente: String(r[4]), direccion: String(r[5]), tipo: String(r[6]), autor: String(r[7]),
    base: num_(r[8]), margin: num_(r[9]), iva: num_(r[10]), total: num_(r[11]), pdfUrl: String(r[12])
  };
}

// Texto seguro para la hoja: evita que algo escrito como "=FÓRMULA(...)" se ejecute y que
// códigos como 1.1.1 o fechas se conviertan en fechas. El apóstrofo inicial fuerza "texto".
function txt_(v) {
  const s = v === null || v === undefined ? '' : String(v);
  if (s === '') return s;
  if (/^[=+\-@\t\r]/.test(s) || /^[\d.,\/:\-\s]+$/.test(s)) return "'" + s;
  return s;
}

function num_(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function redondear_(x) {
  return Math.round((x + Number.EPSILON) * 100) / 100;
}

function clave_(desc, unidad) {
  return String(desc || '').toLowerCase().replace(/\s+/g, ' ').trim() + '|' +
    String(unidad || '').toLowerCase().trim();
}

function nombreArchivo_(nombre) {
  const limpio = String(nombre).replace(/[\\\/:*?"<>|\r\n]+/g, '_').slice(0, 120);
  return /\.pdf$/i.test(limpio) ? limpio : limpio + '.pdf';
}

function hoyTexto_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function fechaTexto_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return String(v === null || v === undefined ? '' : v);
}
