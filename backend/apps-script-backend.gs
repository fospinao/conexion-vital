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
                             'progreso','estaciones_ok','circuitos_completos','ultimo_codigo','ultima_actividad',
                             'e1_movimiento','e2_conexion','e3_bienestar','e4_interaccion','pendientes']);
  hoja(ss, 'estaciones', ['fecha','documento','parque_qr','parque_juego','estacion','vuelta']);
  hoja(ss, 'codigos', ['codigo','mes','documento','parque_juego','fecha_asignado','valido_hasta','tipo']);
  hoja(ss, 'intenciones', ['fecha','documento','intencion','vuelta']);
  hoja(ss, 'redenciones', ['fecha','documento','codigo','parque_juego','valido_hasta','tipo','vuelta']);
  hoja(ss, 'usuarios', ['usuario','hash','sal','rol','activo','creado']);
  hoja(ss, 'sesiones', ['token','usuario','rol','expira']);
  var hu = ss.getSheetByName('usuarios');
  if (hu.getLastRow() < 2) {
    var sal = Utilities.getUuid();
    hu.appendRow(['felipe', hashClave('Activar2026*CV', sal), sal, 'superadmin', 'si',
                  Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd HH:mm:ss')]);
  }
}

function hoja(ss, nombre, encabezados) {
  var h = ss.getSheetByName(nombre);
  if (!h) h = ss.insertSheet(nombre);
  if (encabezados && encabezados.length) {
    h.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    h.setFrozenRows(1);
  }
  return h;
}

/* ================================ GET ================================ */
function doGet(e) {
  var accion = (e.parameter.accion || '').toLowerCase();
  if (accion === 'parques')  return json(listaParques());
  if (accion === 'progreso') return json(progreso(e.parameter.doc, e.parameter.parque));
  if (accion === 'metricas') return json(metricas());
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

/* métricas agregadas para el panel público (sin datos personales) */
function metricas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hp = ss.getSheetByName('participantes');
  var filas = hp && hp.getLastRow() > 1 ? hp.getDataRange().getValues().slice(1) : [];
  var est = { 1: 0, 2: 0, 3: 0, 4: 0 }, pendientes = 0, circuitos = 0, vueltas = 0;
  var porParque = {};
  filas.forEach(function(f) {
    if (!f[0]) return;
    var prog = {};
    try { prog = JSON.parse(f[10] || '{}'); } catch (err) {}
    var falta = false;
    [1, 2, 3, 4].forEach(function(n) { if (prog[n]) est[n]++; else falta = true; });
    if (falta) pendientes++;
    circuitos += Number(f[12]) || 0;
    vueltas += Number(f[9]) || 1;
    var p = String(f[7] || f[6] || '').trim().toLowerCase() || 'otro';
    porParque[p] = (porParque[p] || 0) + 1;
  });

  var hc = ss.getSheetByName('codigos');
  var cf = hc && hc.getLastRow() > 1 ? hc.getDataRange().getValues().slice(1) : [];
  var mesActual = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM');
  var disponibles = 0, modoMes = 'ninguno';
  cf.forEach(function(f) {
    if (!f[0] || normalizarMes(f[1]) !== mesActual) return;
    if (String(f[6] || '').toLowerCase().indexOf('comp') === 0) { modoMes = 'compartido'; return; }
    if (modoMes === 'ninguno') modoMes = 'unico';
    if (!f[2]) disponibles++;
  });
  if (modoMes === 'compartido') disponibles = -1;  /* -1 = ilimitado */
  var hr0 = ss.getSheetByName('redenciones');
  var entregados = hr0 && hr0.getLastRow() > 1 ? hr0.getLastRow() - 1 : 0;

  var parques = [];
  var hpq = ss.getSheetByName('parques');
  if (hpq && hpq.getLastRow() > 1) {
    hpq.getDataRange().getValues().slice(1).forEach(function(f) {
      if (!f[0]) return;
      var id = String(f[0]).trim().toLowerCase();
      parques.push({ id: id, nombre: String(f[1]).trim(), participantes: porParque[id] || 0 });
      delete porParque[id];
    });
  }
  Object.keys(porParque).forEach(function(id) {
    parques.push({ id: id, nombre: id.toUpperCase(), participantes: porParque[id] });
  });

  return {
    participantes: filas.filter(function(f){ return f[0]; }).length,
    circuitos: circuitos,
    entregados: entregados,
    disponibles: disponibles,
    estaciones: est,
    pendientes: pendientes,
    parques: parques,
    modo_mes: modoMes,
    actualizado: Utilities.formatDate(new Date(), ZONA, "yyyy-MM-dd HH:mm")
  };
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
  if (tipo === 'login')     return json(login(d));
  if (tipo === 'admin')     return json(adminRouter(d));
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

  /* columnas legibles por estación: sí / no + en cuáles quedó pendiente */
  var NOMBRES = { 1:'Movimiento', 2:'Conexión', 3:'Bienestar', 4:'Interacción' };
  var porEstacion = [1,2,3,4].map(function(n){ return prog[n] ? 'sí' : 'no'; });
  var pendientes = [1,2,3,4].filter(function(n){ return !prog[n]; })
                            .map(function(n){ return NOMBRES[n]; }).join(', ');

  if (fila === -1) {
    h.appendRow([doc, d.tipodoc || '', d.nombre || '', d.correo || '', d.genero || '',
                 d.nac || '', d.parque || '', d.parqueJuego || '', ahora,
                 d.vuelta || 1, progTxt, okCount, 0, '', ahora]
                .concat(porEstacion).concat([pendientes]));
  } else {
    if (d.nombre) h.getRange(fila, 3).setValue(d.nombre);
    if (d.correo) h.getRange(fila, 4).setValue(d.correo);
    if (d.parqueJuego) h.getRange(fila, 8).setValue(d.parqueJuego);
    h.getRange(fila, 10).setValue(d.vuelta || 1);
    h.getRange(fila, 11).setValue(progTxt);
    h.getRange(fila, 12).setValue(okCount);
    h.getRange(fila, 15).setValue(ahora);
    h.getRange(fila, 16, 1, 5).setValues([porEstacion.concat([pendientes])]);
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
    var hr = hoja(ss, 'redenciones', []);
    var mesActual = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM');

    /* 1. ¿cuántos códigos recibió esta cédula este mes? (log de redenciones) */
    var delMes = 0;
    if (hr.getLastRow() > 1) {
      hr.getDataRange().getValues().slice(1).forEach(function(f) {
        if (String(f[1] || '').replace(/\D/g, '') === doc && mesDe(f[0]) === mesActual) delMes++;
      });
    }
    if (delMes >= MAX_CODIGOS_MES) return { error: 'limite_mes' };

    var vence = new Date();
    vence.setMonth(vence.getMonth() + MESES_VIGENCIA);
    var validoHasta = Utilities.formatDate(vence, ZONA, 'yyyy-MM-dd');
    var ahora = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd HH:mm:ss');
    var parque = d.parque_juego || d.parque || '';
    var filas = h.getDataRange().getValues();

    /* 2. ¿hay código compartido para el mes? lo reciben todos */
    for (var i = 1; i < filas.length; i++) {
      if (filas[i][0] && normalizarMes(filas[i][1]) === mesActual &&
          String(filas[i][6] || '').toLowerCase().indexOf('comp') === 0) {
        var codC = String(filas[i][0]);
        hr.appendRow([ahora, "'" + doc, codC, parque, validoHasta, 'compartido', d.vuelta || 1]);
        marcarCircuito(ss, doc, codC);
        return { codigo: codC, valido_hasta: validoHasta, compartido: true };
      }
    }

    /* 3. primer código único libre del mes en curso */
    for (var j = 1; j < filas.length; j++) {
      if (filas[j][0] && !filas[j][2] && normalizarMes(filas[j][1]) === mesActual) {
        var codU = String(filas[j][0]);
        h.getRange(j + 1, 3).setValue("'" + doc);
        h.getRange(j + 1, 4).setValue(parque);
        h.getRange(j + 1, 5).setValue(ahora);
        h.getRange(j + 1, 6).setValue(validoHasta);
        hr.appendRow([ahora, "'" + doc, codU, parque, validoHasta, 'unico', d.vuelta || 1]);
        marcarCircuito(ss, doc, codU);
        return { codigo: codU, valido_hasta: validoHasta };
      }
    }
    return { error: 'agotados' };
  } finally {
    lock.releaseLock();
  }
}

/* mes (yyyy-MM) de una fecha guardada: Sheets a veces la devuelve como Date */
function mesDe(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ZONA, 'yyyy-MM');
  return String(v || '').slice(0, 7);
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

/* ====================== LOOK & FEEL (ejecutar 1 vez) ======================
   Formatea todas las pestañas con la marca y crea la pestaña "Dashboard"
   con métricas en vivo. Se puede volver a ejecutar sin dañar nada. */
function embellecer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  /* separador de argumentos según el idioma del archivo (es → ";") */
  var SEP = ((ss.getSpreadsheetLocale() || '').indexOf('en') === 0) ? ',' : ';';
  var TINTA = '#123A5C', CIELO = '#57AEE4', VERDE = '#6FBF73',
      LILA = '#9B7FD4', SOL = '#F58220', ROJO = '#E45B5B';
  var S_VERDE = '#DFF3E0', S_CIELO = '#E3F2FC', S_LILA = '#EFE9FA',
      S_SOL = '#FDE9D7', HIELO = '#EAF6FE', ROJO_S = '#FDECEC';

  decorarTab(ss, 'parques',       VERDE, S_VERDE);
  decorarTab(ss, 'participantes', TINTA, HIELO);
  decorarTab(ss, 'estaciones',    CIELO, S_CIELO);
  decorarTab(ss, 'codigos',       SOL,   S_SOL);
  decorarTab(ss, 'intenciones',   LILA,  S_LILA);
  decorarTab(ss, 'redenciones',   SOL,   S_SOL);

  /* semáforos sí/no y pendientes en participantes */
  var hp = ss.getSheetByName('participantes');
  if (hp) {
    var rEst = hp.getRange('P2:S1000'), rPen = hp.getRange('T2:T1000');
    var reglas = [
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('sí')
        .setBackground(S_VERDE).setFontColor('#2F7D33').setRanges([rEst]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('no')
        .setBackground(ROJO_S).setFontColor(ROJO).setRanges([rEst]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenCellNotEmpty()
        .setBackground(S_SOL).setFontColor('#9A5B14').setRanges([rPen]).build()
    ];
    hp.setConditionalFormatRules(reglas);
    hp.getRange('K2:K1000').setFontSize(8).setFontColor('#8A99A8');
  }

  /* ------------------------------ Dashboard ------------------------------ */
  var d = ss.getSheetByName('Dashboard');
  if (!d) d = ss.insertSheet('Dashboard', 0);
  d.clear();
  d.getBandings().forEach(function(b){ b.remove(); });
  d.setHiddenGridlines(true);
  d.setTabColor(TINTA);
  d.getRange('A1:J30').setBackground('#FFFFFF').setFontFamily('Montserrat').setFontColor(TINTA);

  d.setColumnWidth(1, 18);
  d.setColumnWidths(2, 1, 150); d.setColumnWidths(3, 4, 72);
  d.setColumnWidth(7, 40); d.setColumnWidth(8, 150); d.setColumnWidth(9, 100);
  d.setRowHeights(1, 1, 14); d.setRowHeights(2, 1, 8);
  d.setRowHeights(3, 1, 34); d.setRowHeights(4, 1, 22);
  d.setRowHeights(5, 1, 12); d.setRowHeights(6, 1, 22);
  d.setRowHeights(7, 2, 24); d.setRowHeights(9, 1, 16);

  /* franja de marca */
  d.getRange('B2:I4').setBackground(TINTA);
  d.getRange('B3:I3').merge().setValue('CONEXIÓN VITAL · PROGRAMA ACTIVAR')
    .setFontColor('#FFFFFF').setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center');
  d.getRange('B4:I4').merge().setValue('Fundación Postobón · datos en vivo del circuito de bienestar')
    .setFontColor(CIELO).setFontSize(10).setHorizontalAlignment('center');

  /* tarjetas de indicadores */
  tarjeta(d, 'B6:C6', 'B7:C8', 'PARTICIPANTES', '=COUNTA(participantes!A2:A)', HIELO, TINTA);
  tarjeta(d, 'D6:E6', 'D7:E8', 'CIRCUITOS COMPLETOS', '=SUM(participantes!M2:M)', S_VERDE, '#2F7D33');
  tarjeta(d, 'F6:G6', 'F7:G8', 'CÓDIGOS ENTREGADOS', '=COUNTIF(codigos!C2:C' + SEP + '"?*")', S_SOL, '#9A5B14');
  tarjeta(d, 'H6:I6', 'H7:I8', 'DISPONIBLES ESTE MES',
    '=SUMPRODUCT((LEFT(TO_TEXT(codigos!B2:B)' + SEP + '7)=TEXT(TODAY()' + SEP + '"yyyy-MM"))*(codigos!C2:C="")*(codigos!A2:A<>""))',
    S_LILA, '#5B3E9E');

  /* embudo por estación */
  d.getRange('B10:F10').merge().setValue('EMBUDO POR ESTACIÓN (completaron)')
    .setFontWeight('bold').setFontSize(11);
  var estaciones = [
    ['1 · Movimiento',  'P', VERDE],
    ['2 · Conexión',    'Q', LILA],
    ['3 · Bienestar',   'R', CIELO],
    ['4 · Interacción', 'S', SOL]
  ];
  for (var i = 0; i < 4; i++) {
    var fila = 11 + i;
    d.getRange('B' + fila).setValue(estaciones[i][0]).setFontSize(10);
    d.getRange('C' + fila).setFormula('=COUNTIF(participantes!' + estaciones[i][1] + '2:' + estaciones[i][1] + SEP + '"sí")')
      .setFontWeight('bold').setHorizontalAlignment('center');
    d.getRange('D' + fila + ':F' + fila).merge()
      .setFormula('=REPT("█"' + SEP + 'ROUND(24*C' + fila + '/MAX(1' + SEP + '$B$7)))')
      .setFontColor(estaciones[i][2]).setFontSize(12).setHorizontalAlignment('left');
  }
  d.getRange('B16').setValue('Con estaciones pendientes (posible deserción):').setFontSize(10);
  d.getRange('C16').setFormula('=COUNTIF(participantes!T2:T' + SEP + '"?*")')
    .setFontWeight('bold').setFontColor(ROJO).setHorizontalAlignment('center');

  /* participantes por parque */
  d.getRange('H10:I10').merge().setValue('POR PARQUE / CLUB').setFontWeight('bold').setFontSize(11);
  d.getRange('H11').setFormula('=ARRAYFORMULA(IF(parques!A2:A10=""' + SEP + '""' + SEP + 'parques!B2:B10))');
  d.getRange('I11').setFormula('=ARRAYFORMULA(IF(parques!A2:A10=""' + SEP + '""' + SEP + 'COUNTIF(participantes!H2:H' + SEP + 'parques!A2:A10)))');
  d.getRange('H11:H19').setFontSize(10);
  d.getRange('I11:I19').setFontWeight('bold').setHorizontalAlignment('center');

  d.getRange('B19:I19').merge().setValue('Los datos se actualizan solos con cada participación en los parques.')
    .setFontSize(9).setFontColor('#8A99A8');

  ss.setActiveSheet(d);
  ss.moveActiveSheet(1);
}

function decorarTab(ss, nombre, color, suave) {
  var h = ss.getSheetByName(nombre);
  if (!h) return;
  var cols = Math.max(1, h.getLastColumn());
  h.setTabColor(color);
  h.setFrozenRows(1);
  h.getRange(1, 1, 1, cols).setBackground(color).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10).setFontFamily('Montserrat');
  h.setRowHeight(1, 28);
  h.getBandings().forEach(function(b){ b.remove(); });
  h.getRange(2, 1, 999, cols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
    .setFirstRowColor('#FFFFFF').setSecondRowColor(suave);
  for (var c = 1; c <= cols; c++) h.autoResizeColumn(c);
  if (h.getColumnWidth(1) < 90) h.setColumnWidth(1, 90);
}

function tarjeta(d, rLabel, rNum, texto, formula, fondo, colorNum) {
  d.getRange(rLabel).merge().setValue(texto).setBackground(fondo)
    .setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center')
    .setFontColor(colorNum);
  d.getRange(rNum).merge().setFormula(formula).setBackground(fondo)
    .setFontSize(24).setFontWeight('bold').setHorizontalAlignment('center')
    .setVerticalAlignment('middle').setFontColor(colorNum);
}

/* ====================== ADMINISTRACIÓN (login + roles) ======================
   Usuarios en la pestaña "usuarios" (contraseña con hash SHA-256 + sal).
   Roles: superadmin (todo), admin (cifras + descargas + cargar códigos),
   lector (solo cifras). Las sesiones duran 12 horas. */

var SESION_HORAS = 12;

function hashClave(clave, sal) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + '|' + clave, Utilities.Charset.UTF_8);
  return raw.map(function(b){ var v = (b + 256) % 256; return ('0' + v.toString(16)).slice(-2); }).join('');
}

function login(d) {
  var usuario = String(d.usuario || '').trim().toLowerCase();
  var clave = String(d.clave || '');
  if (!usuario || !clave) return { error: 'credenciales' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hu = ss.getSheetByName('usuarios');
  if (!hu) return { error: 'credenciales' };
  var filas = hu.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).trim().toLowerCase() === usuario &&
        String(filas[i][4]).toLowerCase().indexOf('s') === 0) {
      if (hashClave(clave, String(filas[i][2])) === String(filas[i][1])) {
        var hs = hoja(ss, 'sesiones', []);
        limpiarSesiones(hs);
        var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
        hs.appendRow([token, usuario, String(filas[i][3]), Date.now() + SESION_HORAS * 3600 * 1000]);
        return { token: token, usuario: usuario, rol: String(filas[i][3]) };
      }
      break;
    }
  }
  Utilities.sleep(600);
  return { error: 'credenciales' };
}

function limpiarSesiones(hs) {
  var filas = hs.getDataRange().getValues();
  for (var i = filas.length - 1; i >= 1; i--) {
    if (Number(filas[i][3]) < Date.now()) hs.deleteRow(i + 1);
  }
}

function sesionDe(token) {
  if (!token) return null;
  var hs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sesiones');
  if (!hs) return null;
  var filas = hs.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(token) && Number(filas[i][3]) > Date.now()) {
      return { usuario: String(filas[i][1]), rol: String(filas[i][2]), fila: i + 1 };
    }
  }
  return null;
}

function adminRouter(d) {
  var s = sesionDe(d.token);
  if (!s) return { error: 'sesion' };
  var accion = String(d.accion || '');
  var esAdmin = (s.rol === 'admin' || s.rol === 'superadmin');
  var esSuper = (s.rol === 'superadmin');

  if (accion === 'datos')   return adminDatos(s);
  if (accion === 'logout')  { SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sesiones').deleteRow(s.fila); return { ok: true }; }
  if (accion === 'clave')   return cambiarClave(s, d);
  if (accion === 'descargar')          return esAdmin ? descargarTabla(d) : { error: 'permiso' };
  if (accion === 'cargar_codigos')     return esAdmin ? cargarCodigos(d) : { error: 'permiso' };
  if (accion === 'usuarios')           return esSuper ? usuariosLista() : { error: 'permiso' };
  if (accion === 'usuario_guardar')    return esSuper ? usuarioGuardar(s, d) : { error: 'permiso' };
  if (accion === 'borrar_participante')return esSuper ? borrarParticipante(d) : { error: 'permiso' };
  return { error: 'accion' };
}

function adminDatos(s) {
  var m = metricas();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hc = ss.getSheetByName('codigos');
  var hr = ss.getSheetByName('redenciones');
  var porCodigo = {};
  if (hr && hr.getLastRow() > 1) {
    hr.getDataRange().getValues().slice(1).forEach(function(f) {
      var c = String(f[2] || '');
      if (c) porCodigo[c] = (porCodigo[c] || 0) + 1;
    });
  }
  var meses = {};
  if (hc && hc.getLastRow() > 1) {
    hc.getDataRange().getValues().slice(1).forEach(function(f) {
      if (!f[0]) return;
      var mes = normalizarMes(f[1]) || 'sin mes';
      if (!meses[mes]) meses[mes] = { mes: mes, tipo: 'unico', total: 0, usados: 0, libres: 0, codigo: '' };
      if (String(f[6] || '').toLowerCase().indexOf('comp') === 0) {
        meses[mes].tipo = 'compartido';
        meses[mes].codigo = String(f[0]);
        meses[mes].total = 1;
        meses[mes].usados = porCodigo[String(f[0])] || 0;
        meses[mes].libres = null;
      } else {
        meses[mes].total++;
        if (f[2]) meses[mes].usados++; else meses[mes].libres++;
      }
    });
  }
  var lista = Object.keys(meses).sort().reverse().map(function(k){ return meses[k]; });
  return { usuario: s.usuario, rol: s.rol, metricas: m, codigos_mes: lista };
}

function cambiarClave(s, d) {
  var actual = String(d.actual || ''), nueva = String(d.nueva || '');
  if (nueva.length < 8) return { error: 'clave_corta' };
  var hu = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('usuarios');
  var filas = hu.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).trim().toLowerCase() === s.usuario) {
      if (hashClave(actual, String(filas[i][2])) !== String(filas[i][1])) return { error: 'credenciales' };
      var sal = Utilities.getUuid();
      hu.getRange(i + 1, 2).setValue(hashClave(nueva, sal));
      hu.getRange(i + 1, 3).setValue(sal);
      return { ok: true };
    }
  }
  return { error: 'credenciales' };
}

function descargarTabla(d) {
  var permitidas = ['participantes', 'estaciones', 'codigos', 'intenciones', 'parques', 'redenciones'];
  var tabla = String(d.tabla || '');
  if (permitidas.indexOf(tabla) === -1) return { error: 'tabla' };
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabla);
  if (!h) return { error: 'tabla' };
  var filas = h.getDataRange().getValues();
  var csv = filas.map(function(fila) {
    return fila.map(function(v) {
      if (v instanceof Date) v = Utilities.formatDate(v, ZONA, 'yyyy-MM-dd HH:mm:ss');
      v = String(v == null ? '' : v);
      return /[",\n;]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
    }).join(',');
  }).join('\r\n');
  return { nombre: tabla + '.csv', csv: csv };
}

function cargarCodigos(d) {
  var mes = String(d.mes || '').trim();
  if (!/^\d{4}-\d{2}$/.test(mes)) return { error: 'mes' };
  var compartido = (String(d.modo || '') === 'compartido');
  var lista = (d.codigos || []).map(function(c){ return String(c).trim(); }).filter(function(c){ return c; });
  if (!lista.length) return { error: 'vacio' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var h = hoja(SpreadsheetApp.getActiveSpreadsheet(), 'codigos', []);
    var filas = h.getDataRange().getValues();
    var existentes = {}, filaComp = -1, hayUnicos = false;
    for (var i = 1; i < filas.length; i++) {
      if (!filas[i][0]) continue;
      existentes[String(filas[i][0]).trim()] = true;
      if (normalizarMes(filas[i][1]) !== mes) continue;
      if (String(filas[i][6] || '').toLowerCase().indexOf('comp') === 0) filaComp = i + 1;
      else hayUnicos = true;
    }
    /* un mes es de códigos únicos o de código compartido, nunca mezclado */
    if (compartido) {
      if (hayUnicos) return { error: 'mixto' };
      var cod = lista[0];
      if (filaComp > 0) {
        h.getRange(filaComp, 1).setValue(cod);
        return { ok: true, compartido: true, codigo: cod, reemplazado: true };
      }
      h.appendRow([cod, mes, '', '', '', '', 'compartido']);
      return { ok: true, compartido: true, codigo: cod };
    }
    if (filaComp > 0) return { error: 'mixto' };
    var nuevas = [], duplicados = 0;
    lista.forEach(function(c) {
      if (existentes[c]) { duplicados++; return; }
      existentes[c] = true;
      nuevas.push([c, mes, '', '', '', '', 'unico']);
    });
    if (nuevas.length) h.getRange(h.getLastRow() + 1, 1, nuevas.length, 7).setValues(nuevas);
    return { ok: true, agregados: nuevas.length, duplicados: duplicados };
  } finally {
    lock.releaseLock();
  }
}

function usuariosLista() {
  var hu = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('usuarios');
  if (!hu || hu.getLastRow() < 2) return { usuarios: [] };
  return { usuarios: hu.getDataRange().getValues().slice(1).filter(function(f){ return f[0]; }).map(function(f) {
    return { usuario: String(f[0]), rol: String(f[3]), activo: String(f[4]).toLowerCase().indexOf('s') === 0,
             creado: f[5] instanceof Date ? Utilities.formatDate(f[5], ZONA, 'yyyy-MM-dd') : String(f[5]).slice(0, 10) };
  }) };
}

function usuarioGuardar(s, d) {
  var usuario = String(d.usuario || '').trim().toLowerCase();
  var rol = String(d.rol || '');
  var activo = d.activo === false ? 'no' : 'si';
  var clave = String(d.clave || '');
  if (!/^[a-z0-9._-]{3,30}$/.test(usuario)) return { error: 'usuario' };
  if (['superadmin', 'admin', 'lector'].indexOf(rol) === -1) return { error: 'rol' };
  if (usuario === s.usuario && (rol !== 'superadmin' || activo === 'no')) return { error: 'propio' };
  var hu = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('usuarios');
  var filas = hu.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).trim().toLowerCase() === usuario) {
      hu.getRange(i + 1, 4).setValue(rol);
      hu.getRange(i + 1, 5).setValue(activo);
      if (clave) {
        if (clave.length < 8) return { error: 'clave_corta' };
        var sal = Utilities.getUuid();
        hu.getRange(i + 1, 2).setValue(hashClave(clave, sal));
        hu.getRange(i + 1, 3).setValue(sal);
      }
      return { ok: true, actualizado: true };
    }
  }
  if (clave.length < 8) return { error: 'clave_corta' };
  var sal2 = Utilities.getUuid();
  hu.appendRow([usuario, hashClave(clave, sal2), sal2, rol, activo,
                Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd HH:mm:ss')]);
  return { ok: true, creado: true };
}

function borrarParticipante(d) {
  var doc = String(d.doc || '').replace(/\D/g, '');
  if (!doc) return { error: 'doc' };
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('participantes');
  if (!h) return { error: 'no_existe' };
  var filas = h.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === doc) { h.deleteRow(i + 1); return { ok: true }; }
  }
  return { error: 'no_existe' };
}
