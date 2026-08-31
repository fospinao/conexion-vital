/**
 * BACKEND · Conexión Vital · Google Sheets + Apps Script
 * =======================================================
 * Todo vive en UN solo Google Sheet con estas pestañas (se crean solas):
 *
 *  - parques        → id | nombre | activo          (la editas tú: agrega filas y aparecen en el registro de la app)
 *  - participantes  → base de datos de la gente (una fila por cédula)
 *  - estaciones     → log de cada estación completada (para medir embudo/deserción)
 *  - codigos        → base de códigos de WeShop (pegas el Excel aquí cada mes)
 *  - intenciones    → intenciones elegidas en la estación Conexión
 *
 * CÓMO INSTALARLO (una sola vez):
 *  1. Crea un Google Sheet nuevo (ej: "Conexión Vital - Datos").
 *  2. Extensiones → Apps Script → borra todo y pega este código completo. Guarda.
 *  3. Ejecuta una vez la función "instalar" (botón ▶) y acepta los permisos.
 *     Esto crea las pestañas con sus encabezados y parques de ejemplo.
 *  4. Implementar → Nueva implementación → tipo "Aplicación web":
 *       - Ejecutar como: Yo
 *       - Quién tiene acceso: Cualquier persona
 *  5. Copia la URL del Web App (termina en /exec) y pégala en
 *     CONFIG.SHEET_API dentro del index.html de la app.
 *
 * CÓMO CARGAR LOS CÓDIGOS DE WESHOP (cada mes):
 *  En la pestaña "codigos", pega en las columnas:
 *    codigo  |  mes (formato 2026-09)  |  las demás columnas se llenan solas
 *  Puedes copiar/pegar directo desde el Excel de WeShop.
 *
 * REGLAS QUE APLICA ESTE BACKEND:
 *  - Máximo 2 códigos por cédula por mes calendario.
 *  - Solo asigna códigos del mes en curso (columna "mes").
 *  - Cada código queda con vigencia de 2 meses desde su asignación.
 */

var ZONA = 'America/Bogota';
var MAX_CODIGOS_MES = 2;
var MESES_VIGENCIA = 2;

/* ============================= INSTALACIÓN ============================= */
function instalar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  hoja(ss, 'parques', ['id','nombre','activo']);
  var hp = ss.getSheetByName('parques');
  if (hp.getLastRow() < 2) {
    hp.getRange(2,1,4,3).setValues([
      ['comfama','COMFAMA','si'],
      ['epm','EPM','si'],
      ['colsubsidio','COLSUBSIDIO','si'],
      ['arvi','PARQUE ARVÍ','si']
    ]);
  }
  hoja(ss, 'participantes', ['documento','tipo_doc','nombre','correo','genero','nacimiento',
                             'parque_registro','parque_juego','fecha_registro','vuelta',
                             'progreso','estaciones_ok','circuitos_completos','ultimo_codigo','ultima_actividad']);
  hoja(ss, 'estaciones', ['fecha','documento','parque_qr','parque_juego','estacion','vuelta']);
  hoja(ss, 'codigos', ['codigo','mes','documento','parque_juego','fecha_asignado','valido_hasta']);
  hoja(ss, 'intenciones', ['fecha','documento','intencion','vuelta']);
}

function hoja(ss, nombre, encabezados) {
  var h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.appendRow(encabezados);
    h.setFrozenRows(1);
  }
  return h;
}

/* ================================ GET ================================ */
function doGet(e) {
  var accion = (e.parameter.accion || '').toLowerCase();
  if (accion === 'parques')  return json(listaParques());
  if (accion === 'progreso') return json(progreso(e.parameter.doc, e.parameter.parque));
  return json({ ok: true, servicio: 'conexion-vital' });
}

function listaParques() {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('parques');
  if (!h) return { parques: [] };
  var filas = h.getDataRange().getValues().slice(1);
  var parques = filas
    .filter(function(f){ return f[0] && String(f[2]).toLowerCase().indexOf('s') === 0; })
    .map(function(f){ return { id: String(f[0]).trim(), nombre: String(f[1]).trim() }; });
  return { parques: parques };
}

function progreso(doc, parque) {
  doc = String(doc || '').replace(/\D/g, '');
  if (!doc) return { existe: false };
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('participantes');
  if (!h) return { existe: false };
  var filas = h.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === doc) {
      var prog = {};
      try { prog = JSON.parse(filas[i][10] || '{}'); } catch (err) {}
      return {
        existe: true,
        usuario: {
          doc: doc, tipodoc: filas[i][1], nombre: filas[i][2], correo: filas[i][3],
          genero: filas[i][4], nac: filas[i][5], parque: filas[i][6],
          parqueJuego: filas[i][7], registro: filas[i][8]
        },
        vuelta: Number(filas[i][9]) || 1,
        progreso: prog,
        codigo: filas[i][13] || null
      };
    }
  }
  return { existe: false };
}

/* ================================ POST ================================ */
function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); }
  catch (err) { return json({ error: 'json' }); }
  var tipo = d.tipo;
  if (tipo === 'registro')  return json(registrar(d));
  if (tipo === 'estacion')  return json(reportarEstacion(d));
  if (tipo === 'codigo')    return json(asignarCodigo(d));
  if (tipo === 'intencion') return json(guardarIntencion(d));
  if (tipo === 'reinicio')  return json(registrar(d)); // el reinicio actualiza vuelta y progreso
  return json({ error: 'tipo desconocido' });
}

/* crea o actualiza la fila del participante */
function registrar(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var h = hoja(ss, 'participantes', []);
  var doc = String(d.doc || '').replace(/\D/g, '');
  if (!doc) return { error: 'doc' };
  var filas = h.getDataRange().getValues();
  var fila = -1;
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === doc) { fila = i + 1; break; }
  }
  var prog = d.progreso || {};
  var progTxt = JSON.stringify(prog);
  var okCount = Object.keys(prog).filter(function(k){ return prog[k]; }).length;
  var ahora = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd HH:mm:ss');

  if (fila === -1) {
    h.appendRow([doc, d.tipodoc || '', d.nombre || '', d.correo || '', d.genero || '',
                 d.nac || '', d.parque || '', d.parqueJuego || '', ahora,
                 d.vuelta || 1, progTxt, okCount, 0, '', ahora]);
  } else {
    if (d.nombre) h.getRange(fila, 3).setValue(d.nombre);
    if (d.correo) h.getRange(fila, 4).setValue(d.correo);
    if (d.parqueJuego) h.getRange(fila, 8).setValue(d.parqueJuego);
    h.getRange(fila, 10).setValue(d.vuelta || 1);
    h.getRange(fila, 11).setValue(progTxt);
    h.getRange(fila, 12).setValue(okCount);
    h.getRange(fila, 15).setValue(ahora);
  }
  return { ok: true };
}

function reportarEstacion(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var h = hoja(ss, 'estaciones', []);
  var ahora = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd HH:mm:ss');
  h.appendRow([ahora, "'" + String(d.doc || ''), d.parque || '', d.parque_juego || '',
               d.estacion || '', d.vuelta || 1]);
  return { ok: true };
}

function guardarIntencion(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var h = hoja(ss, 'intenciones', []);
  var ahora = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd HH:mm:ss');
  h.appendRow([ahora, "'" + String(d.doc || ''), d.intencion || '', d.vuelta || 1]);
  return { ok: true };
}

/* asigna un código del mes en curso, con tope de 2 por cédula al mes.
   Usa LockService para que dos personas no reciban el mismo código. */
function asignarCodigo(d) {
  var doc = String(d.doc || '').replace(/\D/g, '');
  if (!doc) return { error: 'doc' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var h = hoja(ss, 'codigos', []);
    var filas = h.getDataRange().getValues();
    var mesActual = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM');

    /* 1. ¿cuántos códigos tiene esta cédula este mes? */
    var delMes = 0;
    for (var i = 1; i < filas.length; i++) {
      var asignadoA = String(filas[i][2] || '').replace(/\D/g, '');
      var fechaAsig = String(filas[i][4] || '');
      if (asignadoA === doc && fechaAsig.indexOf(mesActual) === 0) delMes++;
    }
    if (delMes >= MAX_CODIGOS_MES) return { error: 'limite_mes' };

    /* 2. primer código libre del mes en curso */
    for (var j = 1; j < filas.length; j++) {
      var mes = normalizarMes(filas[j][1]);
      var libre = !filas[j][2];
      if (libre && mes === mesActual && filas[j][0]) {
        var vence = new Date();
        vence.setMonth(vence.getMonth() + MESES_VIGENCIA);
        var validoHasta = Utilities.formatDate(vence, ZONA, 'yyyy-MM-dd');
        var ahora = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd HH:mm:ss');
        h.getRange(j + 1, 3).setValue("'" + doc);
        h.getRange(j + 1, 4).setValue(d.parque_juego || d.parque || '');
        h.getRange(j + 1, 5).setValue(ahora);
        h.getRange(j + 1, 6).setValue(validoHasta);
        /* marca el circuito completo en participantes */
        marcarCircuito(ss, doc, String(filas[j][0]));
        return { codigo: String(filas[j][0]), valido_hasta: validoHasta };
      }
    }
    return { error: 'agotados' };
  } finally {
    lock.releaseLock();
  }
}

/* acepta "2026-09", fechas de celda, "sep-2026", etc. */
function normalizarMes(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ZONA, 'yyyy-MM');
  var s = String(v || '').trim();
  var m = s.match(/^(\d{4})[-\/](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  return s;
}

function marcarCircuito(ss, doc, codigo) {
  var h = ss.getSheetByName('participantes');
  if (!h) return;
  var filas = h.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === doc) {
      var completos = Number(filas[i][12]) || 0;
      h.getRange(i + 1, 13).setValue(completos + 1);
      h.getRange(i + 1, 14).setValue(codigo);
      return;
    }
  }
}

/* ================================ UTIL ================================ */
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
