/* ============================================================================
   KAPITAL CRM · BACKEND
   Lima Yedi SAPI de CV SOFOM ENR · Kapital Partners
   Google Apps Script + Google Sheets

   INSTALACIÓN (ver la Guía de instalación):
     1. Pega este archivo completo en Extensiones > Apps Script.
     2. Guarda.
     3. En el menú de funciones elige "configurarSistema" y presiona Ejecutar.
     4. Implementar > Nueva implementación > Aplicación web
        - Ejecutar como: Yo
        - Quién tiene acceso: Cualquier persona
     5. Copia la URL /exec y pégala en CONFIG.API_URL del archivo HTML.
   ========================================================================== */

var HOJAS = {
  EJECUTIVOS:  ['id','nombre','usuario','hash','clave_visible','rol','email','telefono','activo','fecha_alta','fecha_baja'],
  CONTACTOS:   ['id','empresa','nombre_comercial','rfc','sector','ciudad','estado','telefono','sitio_web',
                'ventas_estimadas','producto_interes','fuente','estatus','temperatura','ejecutivo',
                'personas_json','proxima_accion','fecha_compromiso','hora_compromiso','nota_proxima','ultima_comentario',
                'etapa','cita_efectuada',
                'notas','activo','fecha_alta','creado_por','ultima_actividad','actualizado_en'],
  ACTIVIDADES: ['id','contacto_id','empresa','ejecutivo','fecha','hora','semana','anio','intento',
                'tipo','resultado','persona','puesto','telefono','correo','duracion','producto','monto',
                'comentarios','proxima_accion','fecha_compromiso','hora_compromiso','nota_proxima','estatus_nuevo','etapa','temperatura','creado_en'],
  TRASPASOS:   ['id','fecha','de_ejecutivo','a_ejecutivo','n','motivo','nota','autorizo','ids','creado_en'],
  CONFIG:      ['clave','valor']
};

/* Equipo inicial: nombre · usuario · rol · correo · CONTRASEÑA INICIAL
   Cada ejecutivo trae su propia contraseña. La de Gerencia es genérica.
   Se pueden cambiar desde el sistema (Administración > Contraseña). */
var EQUIPO_INICIAL = [
  ['Rossana Soto Gaytán',          'rossana',      'GERENCIA',    'rsoto@kpg.mx',           'Kapital2026'],
  ['Dirección Comercial',           'direccion',    'DIRECCION',   'direccion@kpg.mx',    'Direccion4907'],
  ['Alfredo Azuara Miranda',       'aazuara',      'EJECUTIVO',   'aazuaram@kpg.mx',      'Azuara7412'],
  ['Álvaro Rascón',                'arascon',      'EJECUTIVO',   'alvaro.rascon@kpg.mx',      'Rascon3856'],
  ['Catherin Diaz Monge',          'cdiaz',        'EJECUTIVO',   'cdiazm@kpg.mx',        'Diaz6209'],
  ['Gustavo Cruz Aragonez',        'gcruz',        'EJECUTIVO',   'gustavo.cruz@kpg.mx',        'CruzA5137'],
  ['Mario Cosio Villagrana',       'mcosio',       'EJECUTIVO',   'mario.cosio@kpg.mx',       'Cosio8465'],
  ['Martín Humberto Lopez Polo',   'mlopez',       'EJECUTIVO',   'mlopez@kpg.mx',       'Lopez2793'],
  ['Maurilio Cruz Amaro',          'mcruz',        'EJECUTIVO',   'maurilio.cruz@kpg.mx',        'CruzM4128'],
  ['Pavel Aarón González López',   'pgonzalez',    'EJECUTIVO',   'pgonzalez@kpg.mx',    'Gonzalez9351'],
  ['Sofia Rodriguez Melendez',     'srodriguez',   'EJECUTIVO',   'sofia.rodriguez@kpg.mx',   'Rodriguez6074'],
  ['Cesar Ponce Ponce',            'cponce',       'EJECUTIVO',   'cesar.ponce@kpg.mx',       'Ponce5286']
];
var PASS_GERENCIA = 'Kapital2026';

/* ==========================================================================
   1. CONFIGURACIÓN INICIAL  ·  ejecutar UNA sola vez desde el editor
   ========================================================================== */
function configurarSistema(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HOJAS).forEach(function(nombre){
    var sh = ss.getSheetByName(nombre);
    if(!sh){
      sh = ss.insertSheet(nombre);
      sh.getRange(1,1,sh.getMaxRows(),sh.getMaxColumns()).setNumberFormat('@');
    }
    var cols = HOJAS[nombre];
    migrarColumnas(sh, cols);
    sh.getRange(1,1,1,cols.length).setValues([cols])
      .setFontWeight('bold').setBackground('#0D2433').setFontColor('#FFFFFF')
      .setFontFamily('Montserrat').setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.getRange(1,1,sh.getMaxRows(),cols.length).setNumberFormat('@');
  });

  // Semilla del equipo comercial
  var shE = ss.getSheetByName('EJECUTIVOS');
  if(shE.getLastRow() < 2){
    var filas = EQUIPO_INICIAL.map(function(e,i){
      return ['EJ-'+pad2(i+1), e[0], e[1], hashPass(e[4]), e[4], e[2], e[3], '', 'SI', hoyISO(), ''];
    });
    shE.getRange(2,1,filas.length,HOJAS.EJECUTIVOS.length).setValues(filas);
  }

  // Configuración por defecto
  var shC = ss.getSheetByName('CONFIG');
  if(shC.getLastRow() < 2){
    shC.getRange(2,1,1,2).setValues([['metas', JSON.stringify({diaria:4,semanal:20,dias:5})]]);
  }

  var hoja0 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja1');
  if(hoja0 && ss.getSheets().length > 5){ try{ ss.deleteSheet(hoja0); }catch(err){} }

  SpreadsheetApp.getUi().alert(
    'KAPITAL CRM configurado\n\n' +
    'Se crearon las hojas y se dieron de alta ' + EQUIPO_INICIAL.length + ' usuarios.\n\n' +
    'Cada ejecutivo tiene su propia contraseña inicial (ver la hoja de credenciales).\n' +
    'Gerencia (rossana): ' + PASS_GERENCIA + '  ·  Dirección Comercial (direccion): Direccion4907\n\n' +
    'Siguientes pasos:\n' +
    '1. Implementar > Nueva implementación > Aplicación web ' +
    '(Ejecutar como: Yo · Acceso: Cualquier persona) y pegar la URL /exec en el archivo HTML.\n' +
    '2. Menú KAPITAL CRM > Definir la liga que abre el equipo.\n' +
    '3. Menú KAPITAL CRM > Activar correos automáticos.'
  );
}

/* Si la hoja ya tiene datos y el sistema estrenó columnas nuevas, las inserta
   en su lugar en vez de sobrescribir encabezados y desalinear la información. */
function migrarColumnas(sh, cols){
  if(sh.getLastRow() < 1) return;
  var ancho = Math.max(1, sh.getLastColumn());
  var actuales = sh.getRange(1,1,1,ancho).getValues()[0].map(function(v){ return String(v||'').trim(); });
  if(!actuales[0]) return;                       // hoja recién creada, sin encabezados
  for(var i=0;i<cols.length;i++){
    if(actuales.indexOf(cols[i]) >= 0) continue; // la columna ya existe
    var pos = i + 1;
    if(sh.getMaxColumns() < pos) sh.insertColumnsAfter(sh.getMaxColumns(), pos - sh.getMaxColumns());
    sh.insertColumnBefore(pos);
    sh.getRange(1,pos).setValue(cols[i]);
    actuales.splice(i, 0, cols[i]);
  }
}

function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('KAPITAL CRM')
    .addItem('Configurar sistema (primera vez)', 'configurarSistema')
    .addItem('Restablecer contraseña de un usuario', 'restablecerContrasena')
    .addItem('Ver URL del servidor', 'verUrl')
    .addItem('Definir la liga que abre el equipo', 'configurarLigaApp')
    .addSeparator()
    .addItem('Activar correos automáticos', 'instalarCorreos')
    .addItem('Desactivar correos automáticos', 'quitarCorreos')
    .addItem('Enviar correos de prueba ahora', 'enviarCorreosAhora')
    .addSeparator()
    .addItem('Guardar el reporte semanal en PDF (Drive)', 'guardarReporteSemanalPDF')
    .addSeparator()
    .addItem('Respaldar la base ahora', 'respaldarBase')
    .addItem('Activar respaldo automático semanal', 'instalarRespaldo')
    .addItem('Desactivar respaldo automático', 'quitarRespaldo')
    .addToUi();
}
function verUrl(){
  SpreadsheetApp.getUi().alert('URL del servidor:\n\n' + ScriptApp.getService().getUrl() +
    '\n\nPégala en CONFIG.API_URL dentro del archivo KAPITAL_CRM.html');
}
/* Liga que se pone en el botón de los correos automáticos: es la que abre el
   equipo (GitHub Pages o donde hayas publicado el HTML), no la del servidor. */
function configurarLigaApp(){
  var ui = SpreadsheetApp.getUi();
  var actual = PropertiesService.getScriptProperties().getProperty('KP_URL_APP') || '';
  var r = ui.prompt('Liga que abre el equipo',
    'Pega la dirección donde publicaste el KAPITAL_CRM.html ' +
    '(por ejemplo https://tuusuario.github.io/kapital-crm/).\n\n' +
    'Es la liga del botón "Abrir el Kapital CRM" de los correos automáticos.' +
    (actual ? '\n\nActual: ' + actual : ''), ui.ButtonSet.OK_CANCEL);
  if(r.getSelectedButton() !== ui.Button.OK) return;
  var u = r.getResponseText().trim();
  if(!u){ ui.alert('No se guardó nada.'); return; }
  PropertiesService.getScriptProperties().setProperty('KP_URL_APP', u);
  ui.alert('Liga guardada:\n\n' + u);
}
function restablecerContrasena(){
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt('Restablecer contraseña', 'Escribe el USUARIO (ej. aazuara):', ui.ButtonSet.OK_CANCEL);
  if(r1.getSelectedButton() !== ui.Button.OK) return;
  var r2 = ui.prompt('Nueva contraseña', 'Mínimo 6 caracteres:', ui.ButtonSet.OK_CANCEL);
  if(r2.getSelectedButton() !== ui.Button.OK) return;
  var usuario = r1.getResponseText().trim().toLowerCase();
  var pass = r2.getResponseText();
  if(pass.length < 6){ ui.alert('La contraseña debe tener al menos 6 caracteres.'); return; }
  var sh = hoja('EJECUTIVOS');
  var datos = sh.getDataRange().getValues();
  var col = HOJAS.EJECUTIVOS.indexOf('usuario');
  var colHash = HOJAS.EJECUTIVOS.indexOf('hash');
  var colClara = HOJAS.EJECUTIVOS.indexOf('clave_visible');
  for(var i=1;i<datos.length;i++){
    if(String(datos[i][col]).trim().toLowerCase() === usuario){
      sh.getRange(i+1, colHash+1).setValue(hashPass(pass));
      if(colClara >= 0) sh.getRange(i+1, colClara+1).setValue(pass);
      ui.alert('Contraseña actualizada para ' + datos[i][1] + '.');
      return;
    }
  }
  ui.alert('No se encontró el usuario "' + usuario + '".');
}

/* ==========================================================================
   1-bis. CORREOS AUTOMÁTICOS
   · 5:00 pm (L-V): a cada ejecutivo, su cierre del día y lo que trae agendado
     para el siguiente día de trabajo. El viernes, además, el corte de la semana.
   · 3:00 pm (viernes): a Gerencia y Dirección, el cierre semanal con el PDF.
   Los correos salen desde la cuenta dueña de este archivo.
   Se activan una sola vez desde el menú KAPITAL CRM > Activar correos automáticos.
   ========================================================================== */
var HORA_CIERRE_DIA = 17;   // 5 pm, lunes a viernes
var HORA_CIERRE     = 15;   // 3 pm del viernes

function instalarCorreos(){
  quitarCorreos(true);
  ScriptApp.newTrigger('correoCierreDiaEjecutivos').timeBased()
    .everyDays(1).atHour(HORA_CIERRE_DIA).nearMinute(0).create();
  ScriptApp.newTrigger('correoCierreSemanal').timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(HORA_CIERRE).nearMinute(0).create();
  SpreadsheetApp.getUi().alert(
    'Correos automáticos activados\n\n' +
    '· Lunes a viernes ' + HORA_CIERRE_DIA + ':00 hrs — cada ejecutivo recibe su cierre del día y las llamadas ' +
      'que trae agendadas para el día siguiente (el viernes, las del lunes).\n' +
    '· Viernes ' + HORA_CIERRE + ':00 hrs — Gerencia recibe el cierre de la semana.\n\n' +
    'Sólo llegan a los usuarios que tengan correo capturado en el sistema ' +
    '(Usuarios y metas > Editar).\n\n' +
    'Los correos salen desde la cuenta con la que abriste este archivo.'
  );
}
function quitarCorreos(silencioso){
  ScriptApp.getProjectTriggers().forEach(function(t){
    var f = t.getHandlerFunction();
    if(f === 'correoPendientesEjecutivos' || f === 'correoCierreDiaEjecutivos' ||
       f === 'correoCierreSemanal' || f === 'correoResumenGerencia') ScriptApp.deleteTrigger(t);
  });
  if(!silencioso) SpreadsheetApp.getUi().alert('Correos automáticos desactivados.');
}
function enviarCorreosAhora(){
  var r2 = correoCierreDiaEjecutivos(true);
  var r3 = correoCierreSemanal(true);
  SpreadsheetApp.getUi().alert('Prueba de correos\n\n' + r2 + '\n' + r3);
}
/* Genera el PDF de la semana en curso y lo deja en Drive, por si se necesita
   fuera del correo del viernes. */
function guardarReporteSemanalPDF(){
  var hoy = hoyISO();
  var lunes = lunesDeEstaSemana(), domingo = domingoDeISO(hoy);
  var todos = leer('EJECUTIVOS');
  var ejecutivos = todos.filter(function(e){
    return e.rol === 'EJECUTIVO' && String(e.activo).toUpperCase() === 'SI';
  });
  var blob = pdfSemanal(lunes, domingo, ejecutivos, leer('ACTIVIDADES'), leer('CONTACTOS'), metasGuardadas());
  var archivo = DriveApp.createFile(blob);
  SpreadsheetApp.getUi().alert(
    'Reporte semanal generado\n\n' + archivo.getName() +
    '\n\nQuedó en la raíz de tu Google Drive:\n' + archivo.getUrl()
  );
}

/* ==========================================================================
   1-ter. RESPALDO DE LA BASE
   Copia completa del archivo a una carpeta de Drive, con la fecha en el
   nombre. Se conservan los últimos 12 respaldos y los más viejos se retiran
   solos, para que la carpeta no crezca sin control.
   ========================================================================== */
var CARPETA_RESPALDO = 'KAPITAL CRM · Respaldos';
var RESPALDOS_A_CONSERVAR = 12;

function carpetaRespaldo(){
  var it = DriveApp.getFoldersByName(CARPETA_RESPALDO);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_RESPALDO);
}

/* Devuelve el archivo creado. Sirve igual desde el menú y desde el disparador. */
function respaldarBaseSilencioso(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var carpeta = carpetaRespaldo();
  var nombre = 'KAPITAL CRM · respaldo ' + hoyISO();
  var copia = DriveApp.getFileById(ss.getId()).makeCopy(nombre, carpeta);

  /* Se retiran los respaldos más viejos */
  var todos = [];
  var it = carpeta.getFiles();
  while(it.hasNext()){
    var f = it.next();
    if(f.getName().indexOf('KAPITAL CRM · respaldo') === 0) todos.push(f);
  }
  todos.sort(function(a,b){ return b.getDateCreated() - a.getDateCreated(); });
  for(var i = RESPALDOS_A_CONSERVAR; i < todos.length; i++){
    try{ todos[i].setTrashed(true); }catch(err){}
  }
  return { archivo: copia, total: Math.min(todos.length, RESPALDOS_A_CONSERVAR), carpeta: carpeta };
}

function respaldarBase(){
  var r = respaldarBaseSilencioso();
  SpreadsheetApp.getUi().alert(
    'Respaldo generado\n\n' + r.archivo.getName() +
    '\n\nCarpeta en tu Drive: ' + CARPETA_RESPALDO +
    '\nRespaldos guardados: ' + r.total + ' (se conservan los últimos ' + RESPALDOS_A_CONSERVAR + ')' +
    '\n\n' + r.carpeta.getUrl()
  );
}

function respaldoSemanal(){ try{ respaldarBaseSilencioso(); }catch(err){} }

function instalarRespaldo(){
  quitarRespaldo(true);
  ScriptApp.newTrigger('respaldoSemanal').timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(3).nearMinute(0).create();
  SpreadsheetApp.getUi().alert(
    'Respaldo automático activado\n\n' +
    'Cada sábado a las 3:00 am se guarda una copia completa del archivo en la carpeta "' +
    CARPETA_RESPALDO + '" de tu Google Drive.\n\n' +
    'Se conservan los últimos ' + RESPALDOS_A_CONSERVAR + ' respaldos.'
  );
}
function quitarRespaldo(silencioso){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction() === 'respaldoSemanal') ScriptApp.deleteTrigger(t);
  });
  if(!silencioso) SpreadsheetApp.getUi().alert('Respaldo automático desactivado.');
}

function esDiaHabil(){
  var d = new Date().getDay();
  return d >= 1 && d <= 5;
}
function metasGuardadas(){
  var m = { diaria:4, semanal:20, dias:5 };
  leer('CONFIG').forEach(function(r){
    if(r.clave === 'metas'){ try{ m = JSON.parse(r.valor); }catch(err){} }
  });
  return m;
}
function lunesDeEstaSemana(){
  var d = new Date();
  d.setDate(d.getDate() - ((d.getDay()+6)%7));
  return isoDeFecha(d);
}
function isoDeFecha(d){ return d.getFullYear()+'-'+pad2(d.getMonth()+1)+'-'+pad2(d.getDate()); }
/* El siguiente día de trabajo: de lunes a jueves es mañana; el viernes,
   el sábado y el domingo, es el lunes. */
function siguienteDiaHabil(iso){
  var d = fechaDeISO(iso);
  do { d.setDate(d.getDate()+1); } while(d.getDay() === 0 || d.getDay() === 6);
  return isoDeFecha(d);
}
function fmtLargo(iso){
  var p = String(iso).split('-');
  var M = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  var d = new Date(+p[0], +p[1]-1, +p[2]);
  var D = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
  return D[d.getDay()] + ' ' + d.getDate() + ' de ' + M[d.getMonth()];
}
/* Tipos de llamada, incluidos los que se usaron en versiones anteriores */
var TIPOS_LLAMADA = ['Primera llamada','Llamada de seguimiento','Llamada'];
function esLlamadaTipo(t){
  return TIPOS_LLAMADA.indexOf(String(t)) >= 0;
}
/* Estatus con el que quedó el prospecto en esa llamada. Los registros viejos
   guardaban un "resultado": se traduce al vuelo. */
var RESULTADO_A_ESTATUS = {
  'Dar seguimiento':'En seguimiento', 'Volver a llamar':'En seguimiento',
  'Se envió información':'Información enviada', 'Solicitó información':'En seguimiento',
  'Solicitó propuesta':'Propuesta solicitada', 'Cita concretada':'Cita concretada',
  'No interesado':'Descartado', 'Ya trabaja con otra institución':'Descartado',
  'No califica':'Descartado', 'Datos incorrectos':'Descartado'
};
var ESTATUS_CIERRE = ['Cita concretada','Visitado'];
/* Estatus en los que la empresa ya no tiene pendientes de llamada */
var ESTATUS_ABIERTO = ['Cita concretada','Visitado','Descartado'];
function estatusDeActividad(a){
  var e = a.estatus_nuevo || RESULTADO_A_ESTATUS[a.resultado] || a.resultado || '';
  return RESULTADO_A_ESTATUS[e] || e;
}
function esCitaGS(a){ return ESTATUS_CIERRE.indexOf(estatusDeActividad(a)) >= 0; }

/* Logotipo Kapital Partners en blanco (PNG), tomado del manual de marca.
   Va como imagen incrustada (cid) porque Gmail no renderiza SVG ni data:URI. */
var LOGO_KP_B64 =
  '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';
function logoBlob(){
  return Utilities.newBlob(Utilities.base64Decode(LOGO_KP_B64), 'image/png', 'kplogo.png');
}

/* --- Plantilla de correo con la identidad KP (tablas + estilos en línea) --- */
function marcoCorreo(titulo, saludo, cuerpo, pie){
  return '' +
  '<div style="margin:0;padding:0;background:#F5EFEF;font-family:Montserrat,Arial,Helvetica,sans-serif;">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#F5EFEF;padding:24px 12px;">' +
  '<tr><td align="center">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:640px;background:#FFFFFF;border-radius:14px;overflow:hidden;border:1px solid #E4DFDF;">' +
    '<tr><td style="background:#0D2433;padding:22px 26px;">' +
      '<img src="cid:kplogo" width="200" height="29" alt="Kapital Partners" ' +
        'style="display:block;border:0;outline:none;width:200px;height:29px;">' +
      '<div style="color:#00DA9B;font-size:10px;font-weight:bold;letter-spacing:2.4px;margin-top:11px;">' +
        titulo + '</div>' +
    '</td></tr>' +
    '<tr><td style="padding:26px;">' +
      '<div style="font-size:17px;font-weight:bold;color:#0D2433;margin-bottom:14px;">' + saludo + '</div>' +
      cuerpo +
    '</td></tr>' +
    '<tr><td style="background:#FBFAFA;padding:16px 26px;border-top:1px solid #E4DFDF;">' +
      '<div style="font-size:11px;color:#8098A8;line-height:1.7;">' + (pie||'') +
      'Kapital Partners · Lima Yedi SAPI de CV SOFOM ENR<br>' +
      'Correo automático del Kapital CRM. No es necesario responderlo.</div>' +
    '</td></tr>' +
  '</table></td></tr></table></div>';
}
function cajaMeta(hechas, meta, texto){
  var ok = meta && hechas >= meta;
  var bg = ok ? '#E6FBF4' : '#FDF1EE';
  var col = ok ? '#04795A' : '#A8331A';
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + bg +
    ';border-radius:10px;margin-bottom:18px;"><tr><td style="padding:14px 18px;">' +
    '<span style="font-size:11px;font-weight:bold;letter-spacing:1.5px;color:#4E6675;">META DE HOY</span><br>' +
    '<span style="font-size:26px;font-weight:bold;color:' + col + ';">' + hechas + '</span>' +
    '<span style="font-size:14px;color:#4E6675;font-weight:bold;"> / ' + meta + ' llamadas</span><br>' +
    '<span style="font-size:12.5px;color:#4E6675;">' + texto + '</span>' +
    '</td></tr></table>';
}

/* --- 5:00 pm · cierre del día para cada ejecutivo -------------------------
   Le dice cómo cerró: cuántas llamadas puso contra su meta, cómo se repartieron
   y qué trae agendado para el siguiente día de trabajo. El viernes ese siguiente
   día es el lunes, y además se le suma el corte de la semana completa.
--------------------------------------------------------------------------- */
function correoCierreDiaEjecutivos(forzar){
  if(!forzar && !esDiaHabil()) return 'Cierre del día: fin de semana, no se envió nada.';
  var hoy = hoyISO();
  var manana = siguienteDiaHabil(hoy);
  var esViernes = fechaDeISO(hoy).getDay() === 5;
  var lunes = lunesDeEstaSemana(), domingo = domingoDeISO(hoy);
  var metas = metasGuardadas();
  var contactos = leer('CONTACTOS');
  var actividades = leer('ACTIVIDADES');
  var enviados = 0, sinCorreo = [];

  leer('EJECUTIVOS').forEach(function(e){
    if(String(e.activo).toUpperCase() !== 'SI') return;
    if(e.rol !== 'EJECUTIVO') return;
    if(!e.email){ sinCorreo.push(e.nombre); return; }

    var mias = actividades.filter(function(a){ return a.ejecutivo === e.nombre; });
    var hoyActs = mias.filter(function(a){ return a.fecha === hoy && esLlamadaTipo(a.tipo); });
    var hechas  = hoyActs.length;
    var primeras = hoyActs.filter(function(a){ return String(a.tipo) === 'Primera llamada'; }).length;
    var seguim   = hechas - primeras;
    var citas    = hoyActs.filter(esCitaGS).length;
    var empresas = {};
    hoyActs.forEach(function(a){ if(a.contacto_id) empresas[a.contacto_id] = 1; });

    var cuerpo = cajaMeta(hechas, metas.diaria,
      hechas >= metas.diaria
        ? (hechas > metas.diaria ? 'Cerraste ' + (hechas - metas.diaria) + ' llamada(s) por arriba de la meta.'
                                 : 'Meta del día cumplida.')
        : 'Cerraste ' + Math.max(0, metas.diaria - hechas) + ' llamada(s) por debajo de la meta del día.');

    /* Cómo se repartieron las llamadas de hoy */
    cuerpo += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" ' +
      'style="border:1px solid #E4DFDF;border-radius:9px;margin-bottom:18px;"><tr>' +
      celdaKpi(primeras, 'Primeras llamadas') +
      celdaKpi(seguim, 'De seguimiento') +
      celdaKpi(Object.keys(empresas).length, 'Empresas trabajadas') +
      celdaKpi(citas, 'Citas concretadas') +
      '</tr></table>';

    /* El viernes, además, el corte de la semana */
    if(esViernes){
      var sem = mias.filter(function(a){
        return a.fecha >= lunes && a.fecha <= domingo && esLlamadaTipo(a.tipo);
      }).length;
      var pctSem = metas.semanal ? Math.round(sem * 100 / metas.semanal) : 0;
      cuerpo += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" ' +
        'style="background:' + (sem >= metas.semanal ? '#E6FBF4' : '#FDF1EE') + ';border-radius:10px;' +
        'margin-bottom:18px;"><tr><td style="padding:14px 18px;">' +
        '<span style="font-size:11px;font-weight:bold;letter-spacing:1.5px;color:#4E6675;">' +
          'CIERRE DE LA SEMANA</span><br>' +
        '<span style="font-size:26px;font-weight:bold;color:' +
          (sem >= metas.semanal ? '#04795A' : '#A8331A') + ';">' + sem + '</span>' +
        '<span style="font-size:14px;color:#4E6675;font-weight:bold;"> / ' + metas.semanal +
          ' llamadas &nbsp;·&nbsp; ' + pctSem + '%</span><br>' +
        '<span style="font-size:12.5px;color:#4E6675;">' +
          (sem >= metas.semanal
            ? 'Semana cerrada en meta. Bien hecho.'
            : 'Te faltaron ' + Math.max(0, metas.semanal - sem) + ' llamadas para cerrar la semana en meta.') +
        '</span></td></tr></table>';
    }

    /* Lo que trae agendado para el siguiente día de trabajo */
    var mios = contactos.filter(function(c){ return c.ejecutivo === e.nombre; });
    var prox = mios.filter(function(c){
      return c.fecha_compromiso === manana &&
             c.proxima_accion && c.proxima_accion !== 'Sin acción pendiente' &&
             ESTATUS_ABIERTO.indexOf(c.estatus) < 0;
    }).sort(function(a,b){
      return String(a.hora_compromiso||'99:99').localeCompare(String(b.hora_compromiso||'99:99'));
    });

    cuerpo += '<div style="font-size:12px;font-weight:bold;letter-spacing:1.5px;color:#108896;' +
      'margin:0 0 10px;">PARA ' + fmtLargo(manana).toUpperCase() +
      (prox.length ? ' · ' + prox.length + ' LLAMADA(S)' : '') + '</div>';

    if(prox.length){
      prox.forEach(function(c){
        var personas = [];
        try{ personas = c.personas_json ? JSON.parse(c.personas_json) : []; }catch(err){}
        var p = personas[0] || {};
        cuerpo += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" ' +
          'style="border:1px solid #E4DFDF;border-left:4px solid #00DA9B;border-radius:9px;' +
          'margin-bottom:9px;"><tr><td style="padding:13px 16px;">' +
          '<div style="font-size:14px;font-weight:bold;color:#0D2433;">' + c.proxima_accion + ' · ' +
            '<span style="color:#108896;">' + c.empresa + '</span></div>' +
          '<div style="font-size:12px;color:#8098A8;margin-top:3px;">' +
            (p.nombre ? p.nombre + (p.puesto ? ' · ' + p.puesto : '') + ' &nbsp;·&nbsp; ' : '') +
            (p.telefono || c.telefono || 'sin teléfono') +
            (c.hora_compromiso ? ' &nbsp;·&nbsp; <b style="color:#0D2433;">' + c.hora_compromiso + ' hrs</b>' : '') +
          '</div>' +
          (c.nota_proxima ? '<div style="font-size:12.5px;color:#4E6675;margin-top:7px;background:#FBFAFA;' +
            'padding:8px 11px;border-radius:6px;"><b style="color:#0D2433;">Quedamos en:</b> ' +
            c.nota_proxima + '</div>' : '') +
          '</td></tr></table>';
      });
    } else {
      var frias = mios.filter(function(c){ return c.estatus === 'Sin contactar'; }).slice(0, 6);
      cuerpo += '<div style="font-size:13px;color:#4E6675;line-height:1.7;margin-bottom:14px;">' +
        'No traes llamadas agendadas para ese día. Déjalas programadas hoy y mañana sólo registras.' +
        (frias.length ? ' Estas empresas de tu cartera todavía no reciben su primera llamada:' : '') +
        '</div>';
      if(frias.length){
        cuerpo += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" ' +
          'style="border:1px solid #E4DFDF;border-radius:9px;">';
        frias.forEach(function(c, i){
          var pp = [];
          try{ pp = c.personas_json ? JSON.parse(c.personas_json) : []; }catch(err){}
          var q = pp[0] || {};
          cuerpo += '<tr><td style="padding:10px 15px;' + (i ? 'border-top:1px solid #E4DFDF;' : '') + '">' +
            '<b style="color:#0D2433;font-size:13px;">' + c.empresa + '</b>' +
            '<span style="color:#8098A8;font-size:12px;"> · ' + (q.nombre || 'sin contacto') +
            ' · ' + (q.telefono || c.telefono || 'sin teléfono') + '</span></td></tr>';
        });
        cuerpo += '</table>';
      }
    }

    cuerpo += '<div style="margin-top:20px;"><a href="' + urlApp() + '" ' +
      'style="display:inline-block;background:#00DA9B;color:#0D2433;text-decoration:none;' +
      'padding:12px 24px;border-radius:8px;font-size:13px;font-weight:bold;">Programar en el Kapital CRM</a></div>';

    try{
      GmailApp.sendEmail(e.email,
        'Cierre del día · ' + hechas + ' llamada(s) · ' + prox.length + ' para ' +
          (esViernes ? 'el lunes' : 'mañana'),
        'Abre el Kapital CRM para ver tu cierre del día.',
        { htmlBody: marcoCorreo(esViernes ? 'CIERRE DEL DÍA Y DE LA SEMANA' : 'CIERRE DEL DÍA',
            'Buenas tardes, ' + String(e.nombre).split(' ')[0] + '.', cuerpo),
          inlineImages: { kplogo: logoBlob() }, name: 'Kapital Partners' });
      enviados++;
    }catch(err){}
  });

  return 'Cierre del día: ' + enviados + ' correo(s) enviado(s).' +
    (sinCorreo.length ? ' Sin correo capturado: ' + sinCorreo.join(', ') + '.' : '');
}


/* ==========================================================================
   REPORTE SEMANAL EN PDF
   Se arma en HTML sencillo y Apps Script lo convierte a PDF para adjuntarlo
   al correo del viernes. Trae el resumen del equipo y una hoja por ejecutivo.
   ========================================================================== */
function pdfSemanal(lunes, domingo, ejecutivos, actividades, contactos, metas){
  function enSem(f){ return f && f >= lunes && f <= domingo; }
  var actSem = actividades.filter(function(a){ return enSem(a.fecha); });
  var metaEquipo = metas.semanal * ejecutivos.length;

  var filas = ejecutivos.map(function(e){
    var mias   = actSem.filter(function(a){ return a.ejecutivo === e.nombre; });
    var llam   = mias.filter(function(a){ return esLlamadaTipo(a.tipo); });
    var citas  = mias.filter(esCitaGS).length;
    var nuevas = contactos.filter(function(c){ return c.ejecutivo===e.nombre && enSem(c.fecha_alta); }).length;
    var dias = {}; llam.forEach(function(a){ dias[a.fecha] = 1; });
    var primeras = llam.filter(function(a){ return a.tipo !== 'Llamada de seguimiento'; }).length;
    return { nombre:e.nombre, acts:mias, llamadas:llam.length, primeras:primeras,
             seguimiento: llam.length - primeras, citas:citas, nuevas:nuevas,
             dias:Object.keys(dias).length };
  }).sort(function(a,b){ return b.llamadas - a.llamadas; });

  var totalLl = filas.reduce(function(s2,f){ return s2+f.llamadas; }, 0);
  var totalCi = filas.reduce(function(s2,f){ return s2+f.citas; }, 0);
  var totalNu = filas.reduce(function(s2,f){ return s2+f.nuevas; }, 0);
  var enMeta  = filas.filter(function(f){ return metas.semanal && f.llamadas >= metas.semanal; }).length;

  /* territorio */
  var ficha = {}; contactos.forEach(function(c){ ficha[c.id] = c; });
  var porEstado = {};
  actSem.forEach(function(a){
    if(!esLlamadaTipo(a.tipo)) return;
    var c = ficha[a.contacto_id];
    var k = (c && c.estado) ? String(c.estado) : 'Sin estado';
    if(!porEstado[k]) porEstado[k] = { n:0, ejec:{} };
    porEstado[k].n++;
    porEstado[k].ejec[a.ejecutivo] = (porEstado[k].ejec[a.ejecutivo]||0)+1;
  });
  var terr = Object.keys(porEstado).map(function(k){
    return { k:k, n:porEstado[k].n, ejec:Object.keys(porEstado[k].ejec) };
  }).sort(function(a,b){ return b.n - a.n; });

  var css =
    '<style>'+
    'body{font-family:Arial,Helvetica,sans-serif;color:#0D2433;font-size:10px;margin:0;}'+
    'h1{font-size:17px;margin:0 0 2px;}'+
    'h2{font-size:14px;margin:0 0 3px;}'+
    '.sub{font-size:10px;color:#4E6675;margin:0 0 14px;}'+
    '.sec{font-size:9px;letter-spacing:1.3px;text-transform:uppercase;color:#108896;'+
      'font-weight:bold;margin:16px 0 7px;}'+
    'table{width:100%;border-collapse:collapse;}'+
    'th{background:#0D2433;color:#fff;font-size:8px;letter-spacing:.5px;text-transform:uppercase;'+
      'padding:6px 7px;text-align:left;}'+
    'td{padding:5px 7px;border-bottom:1px solid #EDEAEA;font-size:9px;}'+
    '.num{text-align:right;}.b{font-weight:bold;}'+
    '.ok{color:#04795A;font-weight:bold;}.mal{color:#A8331A;font-weight:bold;}'+
    '.kpis td{border:1px solid #E4DFDF;background:#FBFAFA;text-align:center;padding:9px 6px;}'+
    '.kpis .v{font-size:18px;font-weight:bold;display:block;}'+
    '.kpis .e{font-size:7.5px;letter-spacing:.8px;text-transform:uppercase;color:#8098A8;}'+
    '.hoja{page-break-before:always;}'+
    '.pie{margin-top:16px;padding-top:7px;border-top:1px solid #E4DFDF;font-size:8px;color:#8098A8;}'+
    '.cab{border-bottom:2.5px solid #00DA9B;padding-bottom:8px;margin-bottom:14px;}'+
    '.marca{float:right;font-size:11px;font-weight:bold;letter-spacing:2px;color:#0D2433;}'+
    '</style>';

  var h = '<html><head><meta charset="utf-8">'+css+'</head><body>'+
    '<div class="cab"><span class="marca">KAPITAL PARTNERS</span>'+
    '<h1>Reporte semanal de actividad comercial</h1>'+
    '<p class="sub">Semana del ' + fmtLargo(lunes) + ' al ' + fmtLargo(domingo) +
    ' &nbsp;·&nbsp; Lima Yedi SAPI de CV SOFOM ENR</p></div>'+

    '<table class="kpis"><tr>'+
      '<td><span class="v">' + totalLl + '</span><span class="e">Llamadas del equipo</span></td>'+
      '<td><span class="v">' + (metaEquipo ? Math.round(totalLl*100/metaEquipo) : 0) + '%</span>'+
        '<span class="e">Avance de la meta conjunta</span></td>'+
      '<td><span class="v">' + enMeta + ' / ' + ejecutivos.length + '</span>'+
        '<span class="e">En meta de ' + metas.semanal + '</span></td>'+
      '<td><span class="v">' + totalCi + '</span><span class="e">Citas concretadas</span></td>'+
      '<td><span class="v">' + totalNu + '</span><span class="e">Empresas nuevas</span></td>'+
    '</tr></table>'+

    '<div class="sec">Cumplimiento por ejecutivo</div>'+
    '<table><tr><th>#</th><th>Ejecutivo</th><th class="num">Llamadas</th>'+
      '<th class="num">Meta ' + metas.semanal + '</th><th class="num">Avance</th>'+
      '<th class="num">1&ordf; vez</th><th class="num">Seguimiento</th>'+
      '<th class="num">Días</th><th class="num">Citas</th><th class="num">Nuevas</th></tr>';
  filas.forEach(function(f, i){
    var pct = metas.semanal ? Math.round(f.llamadas*100/metas.semanal) : 0;
    h += '<tr><td>' + (i+1) + '</td><td class="b">' + f.nombre + '</td>'+
      '<td class="num b">' + f.llamadas + '</td><td class="num">' + metas.semanal + '</td>'+
      '<td class="num ' + (pct>=100?'ok':'mal') + '">' + pct + '%</td>'+
      '<td class="num">' + f.primeras + '</td><td class="num">' + f.seguimiento + '</td>'+
      '<td class="num">' + f.dias + ' / ' + metas.dias + '</td>'+
      '<td class="num">' + f.citas + '</td><td class="num">' + f.nuevas + '</td></tr>';
  });
  h += '</table>';

  if(terr.length){
    h += '<div class="sec">Dónde prospectó el equipo</div>'+
      '<table><tr><th>Estado</th><th class="num">Llamadas</th><th>Ejecutivos</th></tr>';
    terr.slice(0, 15).forEach(function(t){
      h += '<tr><td class="b">' + t.k + '</td><td class="num b">' + t.n + '</td>'+
        '<td>' + t.ejec.join(' · ') + '</td></tr>';
    });
    h += '</table>';
  }

  /* Una hoja por ejecutivo */
  filas.forEach(function(f){
    var pct = metas.semanal ? Math.round(f.llamadas*100/metas.semanal) : 0;
    h += '<div class="hoja"><div class="cab"><span class="marca">KAPITAL PARTNERS</span>'+
      '<h2>' + f.nombre + '</h2>'+
      '<p class="sub">Semana del ' + fmtLargo(lunes) + ' al ' + fmtLargo(domingo) + '</p></div>'+
      '<table class="kpis"><tr>'+
        '<td><span class="v ' + (pct>=100?'ok':'mal') + '">' + pct + '%</span><span class="e">Avance de meta</span></td>'+
        '<td><span class="v">' + f.llamadas + ' / ' + metas.semanal + '</span><span class="e">Llamadas</span></td>'+
        '<td><span class="v">' + f.primeras + '</span><span class="e">Primeras llamadas</span></td>'+
        '<td><span class="v">' + f.seguimiento + '</span><span class="e">De seguimiento</span></td>'+
        '<td><span class="v">' + f.dias + '</span><span class="e">Días con actividad</span></td>'+
        '<td><span class="v">' + f.citas + '</span><span class="e">Citas</span></td>'+
      '</tr></table>'+
      '<div class="sec">Detalle de las llamadas</div>';
    if(f.acts.length){
      h += '<table><tr><th>Fecha</th><th>Hora</th><th>Empresa</th><th class="num">Nº</th>'+
        '<th>Persona</th><th>Estatus</th><th>Qué se habló</th><th>Siguiente paso</th></tr>';
      f.acts.sort(function(a,b){
        return String(a.fecha+(a.hora||'')).localeCompare(String(b.fecha+(b.hora||'')));
      }).forEach(function(a){
        h += '<tr><td>' + a.fecha + '</td><td>' + (a.hora||'') + '</td>'+
          '<td class="b">' + a.empresa + '</td><td class="num">' + (a.intento||'') + '</td>'+
          '<td>' + (a.persona||'') + '</td><td>' + estatusDeActividad(a) + '</td>'+
          '<td>' + (a.comentarios||'') + '</td>'+
          '<td>' + (a.proxima_accion||'') + (a.fecha_compromiso ? ' · ' + a.fecha_compromiso : '') + '</td></tr>';
      });
      h += '</table>';
    } else {
      h += '<p class="sub">Sin llamadas registradas en la semana.</p>';
    }
    h += '</div>';
  });

  h += '<div class="pie">Kapital Partners · Nuestra experiencia, tu crecimiento — '+
    'Reporte automático del Kapital CRM.</div></body></html>';

  return Utilities.newBlob(h, 'text/html', 'reporte.html')
    .getAs('application/pdf')
    .setName('KP_Reporte_semanal_' + lunes + '.pdf');
}

/* --- Viernes 3:00 pm · cierre de la semana para Gerencia --- */
function correoCierreSemanal(forzar){
  var hoy = hoyISO();
  var lunes = lunesDeEstaSemana();
  var metas = metasGuardadas();
  var actividades = leer('ACTIVIDADES');
  var contactos = leer('CONTACTOS');
  var todos = leer('EJECUTIVOS');
  var ejecutivos = todos.filter(function(e){
    return e.rol === 'EJECUTIVO' && String(e.activo).toUpperCase() === 'SI';
  });
  var gerentes = todos.filter(function(e){
    return ROLES_MANDO.indexOf(e.rol) >= 0 && String(e.activo).toUpperCase() === 'SI' && e.email;
  });
  if(!gerentes.length) return 'Gerencia: no hay usuarios de Gerencia con correo capturado.';

  function enSemana(f){ return f && f >= lunes && f <= hoy; }
  var actSem = actividades.filter(function(a){ return enSemana(a.fecha); });

  var filas = ejecutivos.map(function(e){
    var mias = actSem.filter(function(a){ return a.ejecutivo === e.nombre; });
    var llam = mias.filter(function(a){ return esLlamadaTipo(a.tipo); }).length;
    var citas = mias.filter(esCitaGS).length;
    var nuevas = contactos.filter(function(c){
      return c.ejecutivo === e.nombre && enSemana(c.fecha_alta);
    }).length;
    var dias = {};
    mias.forEach(function(a){ if(esLlamadaTipo(a.tipo)) dias[a.fecha] = 1; });
    var vencidos = contactos.filter(function(c){
      return c.ejecutivo === e.nombre && c.fecha_compromiso && c.fecha_compromiso <= hoy &&
             c.proxima_accion && c.proxima_accion !== 'Sin acción pendiente' &&
             ESTATUS_ABIERTO.indexOf(c.estatus) < 0;
    }).length;
    return { nombre:e.nombre, llamadas:llam, citas:citas, nuevas:nuevas,
             dias:Object.keys(dias).length, vencidos:vencidos };
  }).sort(function(a,b){ return b.llamadas - a.llamadas; });

  var totalLlam  = filas.reduce(function(s2,f){ return s2 + f.llamadas; }, 0);
  var totalCitas = filas.reduce(function(s2,f){ return s2 + f.citas; }, 0);
  var totalNuev  = filas.reduce(function(s2,f){ return s2 + f.nuevas; }, 0);
  var cumplen    = filas.filter(function(f){ return f.llamadas >= metas.semanal; }).length;
  var totalVenc  = filas.reduce(function(s2,f){ return s2 + f.vencidos; }, 0);
  var metaEquipo = metas.semanal * ejecutivos.length;
  /* Embudos que se cerraron esta semana: la cita ya se efectuó */
  var cerradosSem = contactos.filter(function(c){ return enSemana(c.cita_efectuada); }).length;
  var porEfectuar = contactos.filter(function(c){
    return c.estatus === 'Cita concretada' && !c.cita_efectuada;
  }).length;

  /* Territorio de la semana: estados donde sí se llamó */
  var porEstado = {}, ficha = {};
  contactos.forEach(function(c){ ficha[c.id] = c; });
  actSem.forEach(function(a){
    if(!esLlamadaTipo(a.tipo)) return;
    var c = ficha[a.contacto_id];
    var est = (c && c.estado) ? String(c.estado) : 'Sin estado';
    porEstado[est] = (porEstado[est] || 0) + 1;
  });
  var territorio = Object.keys(porEstado).map(function(k){
    return { estado:k, n:porEstado[k] };
  }).sort(function(a,b){ return b.n - a.n; }).slice(0, 6);

  var cuerpo =
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px;"><tr>' +
      celdaKpi(totalLlam, 'LLAMADAS') +
      celdaKpi(cumplen + ' / ' + ejecutivos.length, 'EN META DE ' + metas.semanal) +
      celdaKpi(totalCitas, 'CITAS') +
      celdaKpi(totalNuev, 'EMPRESAS NUEVAS') +
    '</tr></table>' +
    '<div style="font-size:12.5px;color:#4E6675;line-height:1.7;margin-bottom:18px;">' +
      'El equipo cerró la semana con <b style="color:#0D2433;">' + totalLlam + ' llamadas</b> ' +
      'de una meta conjunta de ' + metaEquipo + ' (' +
      (metaEquipo ? Math.round(totalLlam * 100 / metaEquipo) : 0) + '%). ' +
      '<b style="color:#04795A;">&#10003; ' + cerradosSem + '</b> empresa(s) pasaron al CRM de citas esta semana' +
      (porEfectuar ? ' y quedan <b style="color:#A8331A;">' + porEfectuar + '</b> cita(s) concretadas sin confirmar que se efectuaran' : '') +
      '.</div>' +
    '<div style="font-size:12px;font-weight:bold;letter-spacing:1.5px;color:#108896;margin-bottom:10px;">' +
      'DETALLE POR EJECUTIVO</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" ' +
      'style="border:1px solid #E4DFDF;border-radius:9px;border-collapse:separate;overflow:hidden;">' +
      '<tr style="background:#FAF7F7;">' +
        '<td style="padding:9px 12px;font-size:10px;font-weight:bold;letter-spacing:1px;color:#4E6675;">EJECUTIVO</td>' +
        '<td align="right" style="padding:9px 8px;font-size:10px;font-weight:bold;letter-spacing:1px;color:#4E6675;">LLAM.</td>' +
        '<td align="right" style="padding:9px 8px;font-size:10px;font-weight:bold;letter-spacing:1px;color:#4E6675;">DÍAS</td>' +
        '<td align="right" style="padding:9px 8px;font-size:10px;font-weight:bold;letter-spacing:1px;color:#4E6675;">CITAS</td>' +
        '<td align="right" style="padding:9px 8px;font-size:10px;font-weight:bold;letter-spacing:1px;color:#4E6675;">NUEVAS</td>' +
        '<td align="right" style="padding:9px 12px;font-size:10px;font-weight:bold;letter-spacing:1px;color:#4E6675;">META ' + metas.semanal + '</td>' +
      '</tr>';
  filas.forEach(function(f){
    var ok = f.llamadas >= metas.semanal;
    cuerpo += '<tr style="border-top:1px solid #E4DFDF;">' +
      '<td style="padding:10px 12px;font-size:13px;font-weight:bold;color:#0D2433;">' + f.nombre + '</td>' +
      '<td align="right" style="padding:10px 8px;font-size:14px;font-weight:bold;color:' + (ok ? '#04795A' : '#A8331A') + ';">' + f.llamadas + '</td>' +
      '<td align="right" style="padding:10px 8px;font-size:12.5px;color:#4E6675;">' + f.dias + ' / ' + metas.dias + '</td>' +
      '<td align="right" style="padding:10px 8px;font-size:12.5px;color:#4E6675;">' + f.citas + '</td>' +
      '<td align="right" style="padding:10px 8px;font-size:12.5px;color:#4E6675;">' + f.nuevas + '</td>' +
      '<td align="right" style="padding:10px 12px;font-size:12px;font-weight:bold;color:' + (ok ? '#04795A' : '#A8331A') + ';">' +
        (ok ? (f.llamadas > metas.semanal ? '&#10003; +' + (f.llamadas - metas.semanal) : '&#10003; Cumplió')
            : 'Faltaron ' + (metas.semanal - f.llamadas)) + '</td></tr>';
  });
  cuerpo += '</table>';

  if(territorio.length){
    var maxT = territorio[0].n || 1;
    cuerpo += '<div style="font-size:12px;font-weight:bold;letter-spacing:1.5px;color:#108896;margin:22px 0 10px;">' +
      'DÓNDE SE PROSPECTÓ</div>' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">';
    territorio.forEach(function(t){
      var pct = Math.max(6, Math.round(t.n * 100 / maxT));
      cuerpo += '<tr><td width="38%" style="padding:5px 10px 5px 0;font-size:12.5px;font-weight:bold;color:#0D2433;">' + t.estado + '</td>' +
        '<td style="padding:5px 0;">' +
          '<table role="presentation" width="' + pct + '%" cellpadding="0" cellspacing="0"><tr>' +
          '<td style="background:#00DA9B;height:9px;border-radius:5px;font-size:0;line-height:0;">&nbsp;</td>' +
          '</tr></table></td>' +
        '<td width="46" align="right" style="padding:5px 0;font-size:12.5px;font-weight:bold;color:#4E6675;">' + t.n + '</td></tr>';
    });
    cuerpo += '</table>';
  }

  if(totalVenc){
    cuerpo += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" ' +
      'style="border:1px solid #F6C9BC;border-left:4px solid #E84D2E;border-radius:9px;margin-top:22px;">' +
      '<tr><td style="padding:13px 16px;">' +
      '<div style="font-size:13px;font-weight:bold;color:#A8331A;">Cabos sueltos: ' + totalVenc +
        ' compromiso(s) de llamada sin atender</div>' +
      '<div style="font-size:12.5px;color:#4E6675;margin-top:5px;line-height:1.7;">' +
        filas.filter(function(f){ return f.vencidos; }).map(function(f){
          return f.nombre + ' (' + f.vencidos + ')';
        }).join(' &nbsp;·&nbsp; ') +
      '</div></td></tr></table>';
  }

  cuerpo += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" ' +
    'style="border:1px solid #E4DFDF;border-left:4px solid #108896;border-radius:9px;margin-top:22px;">' +
    '<tr><td style="padding:13px 16px;">' +
    '<div style="font-size:13px;font-weight:bold;color:#0D2433;">&#128206; Reporte semanal en PDF adjunto</div>' +
    '<div style="font-size:12.5px;color:#4E6675;margin-top:4px;line-height:1.6;">' +
    'Trae el resumen del equipo, dónde se prospectó y una hoja por ejecutivo con el detalle de todas ' +
    'sus llamadas de la semana.</div></td></tr></table>' +
    '<div style="margin-top:20px;"><a href="' + urlApp() + '" ' +
    'style="display:inline-block;background:#0D2433;color:#FFFFFF;text-decoration:none;' +
    'padding:12px 24px;border-radius:8px;font-size:13px;font-weight:bold;">Abrir el Kapital CRM</a></div>';

  var adjuntos = [];
  try{ adjuntos.push(pdfSemanal(lunes, domingo, ejecutivos, actividades, contactos, metas)); }catch(err){}

  var enviados = 0;
  gerentes.forEach(function(g){
    try{
      GmailApp.sendEmail(g.email,
        'Cierre de semana · ' + totalLlam + ' llamadas · ' + cumplen + ' de ' + ejecutivos.length + ' en meta',
        'Abre el Kapital CRM para ver el detalle de la semana.',
        { htmlBody: marcoCorreo('CIERRE DE SEMANA',
            'Semana del ' + fmtLargo(lunes) + ' al ' + fmtLargo(hoy) + '.', cuerpo),
          inlineImages: { kplogo: logoBlob() }, name: 'Kapital Partners' });
      enviados++;
    }catch(err){}
  });
  return 'Gerencia: ' + enviados + ' correo(s) enviado(s).';
}
function celdaKpi(valor, etiqueta){
  return '<td width="25%" style="padding:0 4px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" ' +
    'style="background:#FBFAFA;border:1px solid #E4DFDF;border-radius:9px;"><tr><td align="center" style="padding:12px 6px;">' +
    '<div style="font-size:22px;font-weight:bold;color:#0D2433;">' + valor + '</div>' +
    '<div style="font-size:9px;font-weight:bold;letter-spacing:1px;color:#8098A8;">' + etiqueta + '</div>' +
    '</td></tr></table></td>';
}
function urlApp(){
  var props = PropertiesService.getScriptProperties();
  var u = props.getProperty('KP_URL_APP');
  if(u) return u;
  try{ return ScriptApp.getService().getUrl(); }catch(err){ return '#'; }
}


/* ==========================================================================
   3 bis. CORREO DE META CUMPLIDA
   Se dispara en el momento en que el ejecutivo captura la llamada con la que
   alcanza su meta. Sólo una vez por día y una vez por semana.
   ========================================================================== */
function fechaDeISO(iso){
  var p = String(iso).split('-');
  return new Date(+p[0], +p[1]-1, +p[2]);
}
function lunesDeISO(iso){
  var d = fechaDeISO(iso);
  d.setDate(d.getDate() - ((d.getDay()+6)%7));
  return isoDeFecha(d);
}
function domingoDeISO(iso){
  var d = fechaDeISO(lunesDeISO(iso));
  d.setDate(d.getDate() + 6);
  return isoDeFecha(d);
}
function leerConfigClave(clave){
  var v = '';
  leer('CONFIG').forEach(function(r){ if(r.clave === clave) v = r.valor; });
  return v;
}
function guardarConfigClave(clave, valor){
  var sh = hoja('CONFIG');
  var ultima = sh.getLastRow();
  if(ultima >= 2){
    var claves = sh.getRange(2,1,ultima-1,1).getValues();
    for(var i=0;i<claves.length;i++){
      if(String(claves[i][0]) === clave){ sh.getRange(i+2,2).setValue(valor); return; }
    }
  }
  sh.appendRow([clave, valor]);
}
/* Bitácora de avisos ya enviados, para no repetir. Se poda sola a 120 días. */
function avisosEnviados(){
  var o = {};
  try{ o = JSON.parse(leerConfigClave('avisos_meta') || '{}'); }catch(err){ o = {}; }
  return o;
}
function marcarAviso(clave, fechaRef){
  var o = avisosEnviados();
  o[clave] = fechaRef;
  var corte = new Date(); corte.setDate(corte.getDate() - 120);
  var corteISO = isoDeFecha(corte);
  var limpio = {};
  Object.keys(o).forEach(function(k){ if(String(o[k]) >= corteISO) limpio[k] = o[k]; });
  guardarConfigClave('avisos_meta', JSON.stringify(limpio));
}
function buscarEjecutivoPorNombre(nombre){
  var lista = leer('EJECUTIVOS');
  for(var i=0;i<lista.length;i++) if(lista[i].nombre === nombre) return lista[i];
  return null;
}

function revisarMetaCumplida(a){
  if(!esLlamadaTipo(a.tipo)) return null;
  var e = buscarEjecutivoPorNombre(a.ejecutivo);
  if(!e || !e.email || String(e.activo).toUpperCase() !== 'SI') return null;

  var metas = metasGuardadas();
  var acts = leer('ACTIVIDADES').filter(function(x){
    return x.ejecutivo === a.ejecutivo && esLlamadaTipo(x.tipo);
  });
  var lunes = lunesDeISO(a.fecha), domingo = domingoDeISO(a.fecha);
  var enDia    = acts.filter(function(x){ return x.fecha === a.fecha; }).length;
  var enSemana = acts.filter(function(x){ return x.fecha >= lunes && x.fecha <= domingo; }).length;

  var avisos = avisosEnviados();
  var out = null;

  var kDia = 'D|' + a.ejecutivo + '|' + a.fecha;
  if(metas.diaria && enDia >= metas.diaria && !avisos[kDia]){
    correoMetaCumplida(e, 'dia', enDia, metas.diaria, enSemana, metas.semanal, a.fecha, lunes, domingo);
    marcarAviso(kDia, a.fecha);
    out = { tipo:'dia', hechas:enDia, meta:metas.diaria };
  }
  var kSem = 'S|' + a.ejecutivo + '|' + lunes;
  if(metas.semanal && enSemana >= metas.semanal && !avisos[kSem]){
    correoMetaCumplida(e, 'semana', enSemana, metas.semanal, enSemana, metas.semanal, a.fecha, lunes, domingo);
    marcarAviso(kSem, a.fecha);
    out = { tipo:'semana', hechas:enSemana, meta:metas.semanal };
  }
  return out;
}

function correoMetaCumplida(e, tipo, hechas, meta, enSemana, metaSem, fecha, lunes, domingo){
  var esDia = tipo === 'dia';
  var extra = hechas - meta;
  var titulo = esDia ? 'META DEL DÍA CUMPLIDA' : 'META DE LA SEMANA CUMPLIDA';
  var asunto = esDia
    ? '&#10003; Meta del día cumplida · ' + hechas + ' llamadas'
    : '&#10003; Meta de la semana cumplida · ' + hechas + ' llamadas';

  var cuerpo =
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" ' +
      'style="background:#E6FBF4;border:1px solid #9BEBD3;border-radius:12px;margin-bottom:20px;">' +
      '<tr><td align="center" style="padding:26px 20px;">' +
        '<div style="font-size:34px;line-height:1;color:#04795A;font-weight:bold;">&#10003;</div>' +
        '<div style="font-size:42px;font-weight:bold;color:#0D2433;line-height:1.1;margin-top:8px;">' + hechas + '</div>' +
        '<div style="font-size:13px;font-weight:bold;letter-spacing:1.5px;color:#04795A;">' +
          'DE ' + meta + ' LLAMADAS ' + (esDia ? 'DEL DÍA' : 'DE LA SEMANA') + '</div>' +
        (extra > 0
          ? '<div style="font-size:12.5px;color:#4E6675;margin-top:9px;">Te pasaste por <b style="color:#04795A;">' +
            extra + '</b> llamada' + (extra > 1 ? 's' : '') + '. Todas cuentan.</div>'
          : '') +
      '</td></tr></table>';

  if(esDia){
    var faltanSem = Math.max(0, metaSem - enSemana);
    cuerpo += '<div style="font-size:13px;color:#4E6675;line-height:1.75;margin-bottom:18px;">' +
      'Cerraste el día con la meta hecha. ' +
      (faltanSem
        ? 'En la semana llevas <b style="color:#0D2433;">' + enSemana + ' de ' + metaSem + '</b>; ' +
          'te faltan <b style="color:#0D2433;">' + faltanSem + '</b> para cerrarla completa.'
        : 'Y la semana también la tienes cerrada: <b style="color:#0D2433;">' + enSemana + ' de ' + metaSem + '</b>.') +
      '</div>';
  } else {
    cuerpo += '<div style="font-size:13px;color:#4E6675;line-height:1.75;margin-bottom:18px;">' +
      'Cerraste la semana del <b style="color:#0D2433;">' + fmtLargo(lunes) + '</b> al <b style="color:#0D2433;">' +
      fmtLargo(domingo) + '</b> con la meta cumplida. Bien hecho.' +
      '</div>';
  }

  /* Lo que sigue: sus próximas llamadas agendadas */
  var pend = leer('CONTACTOS').filter(function(c){
    return c.ejecutivo === e.nombre && c.fecha_compromiso &&
           c.proxima_accion && c.proxima_accion !== 'Sin acción pendiente' &&
           ESTATUS_ABIERTO.indexOf(c.estatus) < 0;
  }).sort(function(a2,b2){ return String(a2.fecha_compromiso).localeCompare(String(b2.fecha_compromiso)); }).slice(0, 4);

  if(pend.length){
    cuerpo += '<div style="font-size:12px;font-weight:bold;letter-spacing:1.5px;color:#108896;margin-bottom:10px;">' +
      'LO QUE TIENES AGENDADO</div>' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #E4DFDF;border-radius:9px;">';
    pend.forEach(function(c, i){
      cuerpo += '<tr><td style="padding:10px 15px;' + (i ? 'border-top:1px solid #E4DFDF;' : '') + '">' +
        '<b style="color:#0D2433;font-size:13px;">' + c.empresa + '</b>' +
        '<span style="color:#8098A8;font-size:12px;"> · ' + (c.proxima_accion || '') +
        ' · ' + fmtLargo(c.fecha_compromiso) + '</span></td></tr>';
    });
    cuerpo += '</table>';
  }

  cuerpo += '<div style="margin-top:20px;"><a href="' + urlApp() + '" ' +
    'style="display:inline-block;background:#00DA9B;color:#0D2433;text-decoration:none;' +
    'padding:12px 24px;border-radius:8px;font-size:13px;font-weight:bold;">Abrir el Kapital CRM</a></div>';

  try{
    GmailApp.sendEmail(e.email,
      (esDia ? '✓ Meta del día cumplida · ' : '✓ Meta de la semana cumplida · ') + hechas + ' llamadas',
      'Abre el Kapital CRM para ver tu avance.',
      { htmlBody: marcoCorreo(titulo,
          (esDia ? 'Bien, ' : 'Excelente semana, ') + String(e.nombre).split(' ')[0] + '.', cuerpo),
        inlineImages: { kplogo: logoBlob() }, name: 'Kapital Partners' });
  }catch(err){}
}

/* ==========================================================================
   2. PUNTO DE ENTRADA HTTP
   ========================================================================== */
function doPost(e){
  try{
    var req = JSON.parse(e.postData.contents);
    return json(procesar(req.action, req.data || {}, req.token));
  }catch(err){
    return json({ ok:false, error: String(err && err.message ? err.message : err) });
  }
}
function doGet(e){
  if(e && e.parameter && e.parameter.ping) return json({ ok:true, data:{ estado:'activo', version:'1.0' } });
  return HtmlService.createHtmlOutput(
    '<div style="font-family:Montserrat,Arial,sans-serif;padding:40px;background:#0D2433;color:#fff;height:100%">' +
    '<h2 style="margin:0 0 8px">KAPITAL CRM · Servidor activo</h2>' +
    '<p style="opacity:.7;font-size:13px">Lima Yedi SAPI de CV SOFOM ENR</p>' +
    '<p style="font-size:13px">Esta dirección es el backend del sistema. Copia esta URL completa y pégala en ' +
    '<b>CONFIG.API_URL</b> dentro del archivo KAPITAL_CRM.html</p></div>'
  );
}
function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function procesar(action, data, token){
  if(action === 'login') return { ok:true, data: login(data) };

  var user = validarToken(token);
  if(!user) return { ok:false, error:'Tu sesión expiró. Vuelve a iniciar sesión.' };

  switch(action){
    case 'bootstrap':      return { ok:true, data: bootstrap(user) };
    case 'saveContacto':   return conLock(function(){ return { ok:true, data: saveContacto(data, user) }; });
    case 'saveActividad':  return conLock(function(){ return { ok:true, data: saveActividad(data, user) }; });
    case 'deleteContacto': return conLock(function(){ return { ok:true, data: borrarFila('CONTACTOS', data.id, user) }; });
    case 'deleteActividad':return conLock(function(){ return { ok:true, data: borrarFila('ACTIVIDADES', data.id, user) }; });
    case 'saveEjecutivo':  return conLock(function(){ return { ok:true, data: saveEjecutivo(data, user) }; });
    case 'resetPassword':  return conLock(function(){ return { ok:true, data: resetPassword(data, user) }; });
    case 'deleteEjecutivo':return conLock(function(){ return { ok:true, data: deleteEjecutivo(data, user) }; });
    case 'traspaso':       return conLock(function(){ return { ok:true, data: traspaso(data, user) }; });
    case 'saveConfig':     return conLock(function(){ return { ok:true, data: saveConfig(data, user) }; });
    case 'importBase':     return conLock(function(){ return { ok:true, data: importBase(data, user) }; });
    default: return { ok:false, error:'Acción no reconocida: ' + action };
  }
}
function conLock(fn){
  var lock = LockService.getScriptLock();
  try{
    lock.waitLock(25000);
    return fn();
  }catch(err){
    return { ok:false, error:'El sistema está ocupado. Intenta de nuevo en unos segundos.' };
  }finally{
    try{ lock.releaseLock(); }catch(err2){}
  }
}

/* ==========================================================================
   3. SEGURIDAD
   ========================================================================== */
function secreto(){
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('KP_SECRET');
  if(!s){
    s = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('KP_SECRET', s);
  }
  return s;
}
function hashPass(pass){
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'KPG::' + secreto() + '::' + pass, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(raw);
}
function firmar(texto){
  return Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(texto, secreto(), Utilities.Charset.UTF_8)
  );
}
function crearToken(idEjecutivo){
  var exp = Date.now() + 1000 * 60 * 60 * 24 * 30;   // 30 días
  var payload = idEjecutivo + '|' + exp;
  return Utilities.base64EncodeWebSafe(payload + '|' + firmar(payload));
}
function validarToken(token){
  if(!token) return null;
  try{
    var texto = Utilities.newBlob(Utilities.base64DecodeWebSafe(token)).getDataAsString();
    var p = texto.split('|');
    if(p.length !== 3) return null;
    var payload = p[0] + '|' + p[1];
    if(firmar(payload) !== p[2]) return null;
    if(Number(p[1]) < Date.now()) return null;
    var e = buscarEjecutivoPorId(p[0]);
    if(!e || String(e.activo).toUpperCase() !== 'SI') return null;
    return { id:e.id, nombre:e.nombre, usuario:e.usuario, rol:e.rol };
  }catch(err){ return null; }
}
function login(data){
  var usuario = String(data.usuario || '').trim().toLowerCase();
  var pass = String(data.password || '');
  var lista = leer('EJECUTIVOS');
  for(var i=0;i<lista.length;i++){
    if(String(lista[i].usuario).trim().toLowerCase() === usuario){
      if(String(lista[i].activo).toUpperCase() !== 'SI'){
        throw new Error('Este usuario está dado de baja. Contacta a Gerencia.');
      }
      if(String(lista[i].hash) !== hashPass(pass)){
        throw new Error('Usuario o contraseña incorrectos.');
      }
      return {
        token: crearToken(lista[i].id),
        user: { id:lista[i].id, nombre:lista[i].nombre, usuario:lista[i].usuario, rol:lista[i].rol }
      };
    }
  }
  throw new Error('Usuario o contraseña incorrectos.');
}
/* Gerencia y Dirección Comercial tienen el mismo alcance: ven toda la operación. */
var ROLES_MANDO = ['GERENCIA','DIRECCION'];
function esGerente(user){ return !!(user && ROLES_MANDO.indexOf(user.rol) >= 0); }

/* ==========================================================================
   4. ACCESO A HOJAS
   ========================================================================== */
function hoja(nombre){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if(!sh) throw new Error('Falta la hoja "' + nombre + '". Ejecuta KAPITAL CRM > Configurar sistema.');
  return sh;
}
function leer(nombre){
  var sh = hoja(nombre);
  var ultima = sh.getLastRow();
  if(ultima < 2) return [];
  var cols = HOJAS[nombre];
  var datos = sh.getRange(2, 1, ultima - 1, cols.length).getValues();
  var out = [];
  for(var i=0;i<datos.length;i++){
    if(!datos[i][0]) continue;
    var o = {};
    for(var j=0;j<cols.length;j++) o[cols[j]] = normalizar(datos[i][j]);
    out.push(o);
  }
  return out;
}
function normalizar(v){
  if(v instanceof Date){
    return v.getFullYear() + '-' + pad2(v.getMonth()+1) + '-' + pad2(v.getDate());
  }
  return v === null || v === undefined ? '' : v;
}
function filaDe(nombre, obj){
  return HOJAS[nombre].map(function(c){
    var v = obj[c];
    return (v === null || v === undefined) ? '' : v;
  });
}
function guardarFila(nombre, obj){
  var sh = hoja(nombre);
  var cols = HOJAS[nombre];
  var ultima = sh.getLastRow();
  var fila = filaDe(nombre, obj);
  if(ultima >= 2){
    var ids = sh.getRange(2,1,ultima-1,1).getValues();
    for(var i=0;i<ids.length;i++){
      if(String(ids[i][0]) === String(obj.id)){
        sh.getRange(i+2, 1, 1, cols.length).setValues([fila]);
        return { actualizado:true };
      }
    }
  }
  sh.appendRow(fila);
  return { creado:true };
}
function borrarFila(nombre, id, user){
  var sh = hoja(nombre);
  var ultima = sh.getLastRow();
  if(ultima < 2) return {};
  var ids = sh.getRange(2,1,ultima-1,1).getValues();
  for(var i=0;i<ids.length;i++){
    if(String(ids[i][0]) === String(id)){
      sh.deleteRow(i+2);
      return { borrado:true };
    }
  }
  return { borrado:false };
}
function buscarEjecutivoPorId(id){
  var lista = leer('EJECUTIVOS');
  for(var i=0;i<lista.length;i++) if(String(lista[i].id) === String(id)) return lista[i];
  return null;
}
function pad2(n){ return n < 10 ? '0'+n : ''+n; }
function hoyISO(){
  var d = new Date();
  return d.getFullYear() + '-' + pad2(d.getMonth()+1) + '-' + pad2(d.getDate());
}

/* ==========================================================================
   5. LECTURA GENERAL
   ========================================================================== */
function bootstrap(user){
  /* La contraseña legible sólo viaja hacia Gerencia y Dirección Comercial.
     A un ejecutivo el servidor nunca se la manda, ni la suya ni la de nadie. */
  var mando = esGerente(user);
  var ejecutivos = leer('EJECUTIVOS').map(function(e){
    return { id:e.id, nombre:e.nombre, usuario:e.usuario, rol:e.rol, email:e.email,
             telefono:e.telefono, activo: String(e.activo).toUpperCase()==='SI',
             fecha_alta:e.fecha_alta, fecha_baja:e.fecha_baja,
             clave_visible: mando ? String(e.clave_visible || '') : '' };
  });
  var contactos = leer('CONTACTOS').map(function(c){
    var personas = [];
    try{ personas = c.personas_json ? JSON.parse(c.personas_json) : []; }catch(err){ personas = []; }
    return {
      id:c.id, empresa:c.empresa, nombre_comercial:c.nombre_comercial, rfc:c.rfc, sector:c.sector,
      ciudad:c.ciudad, estado:c.estado, telefono:String(c.telefono||''), sitio_web:c.sitio_web,
      ventas_estimadas:c.ventas_estimadas, producto_interes:c.producto_interes, fuente:c.fuente,
      estatus:c.estatus, temperatura:c.temperatura, ejecutivo:c.ejecutivo, personas:personas,
      proxima_accion:c.proxima_accion, fecha_compromiso:c.fecha_compromiso,
      hora_compromiso:c.hora_compromiso || '',
      nota_proxima:c.nota_proxima, ultima_comentario:c.ultima_comentario,
      etapa:c.etapa || '', cita_efectuada:c.cita_efectuada || '', notas:c.notas,
      activo: String(c.activo).toUpperCase() !== 'NO', fecha_alta:c.fecha_alta,
      creado_por:c.creado_por, ultima_actividad:c.ultima_actividad
    };
  });
  var actividades = leer('ACTIVIDADES').map(function(a){
    a.telefono = String(a.telefono || '');
    a.monto = Number(a.monto) || 0;
    a.duracion = Number(a.duracion) || 0;
    a.intento = Number(a.intento) || 0;
    return a;
  });
  var traspasos = leer('TRASPASOS');
  var config = { metas:{ diaria:4, semanal:20, dias:5 } };
  leer('CONFIG').forEach(function(r){
    if(r.clave === 'metas'){ try{ config.metas = JSON.parse(r.valor); }catch(err){} }
  });

  // Un ejecutivo sólo recibe su propia cartera y su propia bitácora
  if(!esGerente(user)){
    contactos = contactos.filter(function(c){ return c.ejecutivo === user.nombre; });
    actividades = actividades.filter(function(a){ return a.ejecutivo === user.nombre; });
    traspasos = [];
  }
  return { ejecutivos:ejecutivos, contactos:contactos, actividades:actividades,
           traspasos:traspasos, config:config };
}

/* ==========================================================================
   6. ESCRITURAS
   ========================================================================== */
function contactoAFila(c, user){
  return {
    id: c.id, empresa: c.empresa, nombre_comercial: c.nombre_comercial || '', rfc: c.rfc || '',
    sector: c.sector || '', ciudad: c.ciudad || '', estado: c.estado || '',
    telefono: c.telefono || '', sitio_web: c.sitio_web || '',
    ventas_estimadas: c.ventas_estimadas || '', producto_interes: c.producto_interes || '',
    fuente: c.fuente || '', estatus: c.estatus || '', temperatura: c.temperatura || '',
    ejecutivo: c.ejecutivo || '', personas_json: JSON.stringify(c.personas || []),
    proxima_accion: c.proxima_accion || '', fecha_compromiso: c.fecha_compromiso || '',
    hora_compromiso: c.hora_compromiso || '',
    nota_proxima: c.nota_proxima || '', ultima_comentario: c.ultima_comentario || '',
    etapa: c.etapa || '', cita_efectuada: c.cita_efectuada || '',
    notas: c.notas || '', activo: (c.activo === false ? 'NO' : 'SI'),
    fecha_alta: c.fecha_alta || hoyISO(), creado_por: c.creado_por || user.nombre,
    ultima_actividad: c.ultima_actividad || '',
    actualizado_en: new Date().toISOString()
  };
}
function permitido(c, user){
  if(esGerente(user)) return true;
  return !c.ejecutivo || c.ejecutivo === user.nombre;
}
function saveContacto(data, user){
  var c = data.contacto;
  if(!c || !c.empresa) throw new Error('Falta el nombre de la empresa.');
  if(!permitido(c, user)) throw new Error('No puedes modificar la cartera de otro ejecutivo.');
  guardarFila('CONTACTOS', contactoAFila(c, user));
  return { id:c.id };
}
function saveActividad(data, user){
  var a = data.actividad;
  if(!a || !a.fecha) throw new Error('Falta la fecha de la actividad.');
  if(!esGerente(user) && a.ejecutivo !== user.nombre){
    throw new Error('Sólo puedes registrar actividad a tu propio nombre.');
  }
  a.telefono = String(a.telefono || '');
  guardarFila('ACTIVIDADES', a);
  if(data.contacto) guardarFila('CONTACTOS', contactoAFila(data.contacto, user));
  var felicitacion = null;
  try{ felicitacion = revisarMetaCumplida(a); }catch(err){}   // nunca frenar el guardado
  return { id:a.id, meta: felicitacion };
}
function saveEjecutivo(data, user){
  if(!esGerente(user)) throw new Error('Sólo Gerencia puede administrar ejecutivos.');
  var e = data.ejecutivo;
  if(!e || !e.nombre || !e.usuario) throw new Error('Faltan datos del ejecutivo.');

  var lista = leer('EJECUTIVOS');
  var previo = null;
  for(var i=0;i<lista.length;i++){
    if(String(lista[i].id) === String(e.id)) previo = lista[i];
    else if(String(lista[i].usuario).toLowerCase() === String(e.usuario).toLowerCase()){
      throw new Error('El usuario "' + e.usuario + '" ya está en uso por ' + lista[i].nombre + '.');
    }
  }
  var hash = previo ? previo.hash : '';
  var clara = previo ? String(previo.clave_visible || '') : '';
  if(e.pass){ hash = hashPass(e.pass); clara = String(e.pass); }
  if(!hash) throw new Error('Debes asignar una contraseña.');

  guardarFila('EJECUTIVOS', {
    id: e.id, nombre: e.nombre, usuario: String(e.usuario).toLowerCase(), hash: hash,
    clave_visible: clara,
    rol: e.rol || 'EJECUTIVO', email: e.email || '', telefono: e.telefono || '',
    activo: e.activo ? 'SI' : 'NO',
    fecha_alta: e.fecha_alta || hoyISO(),
    fecha_baja: e.activo ? '' : (e.fecha_baja || hoyISO())
  });

  // Si cambió el nombre, se propaga a la cartera y a la bitácora
  var anterior = data.nombre_anterior;
  if(anterior && anterior !== e.nombre){
    renombrarEnColumna('CONTACTOS', 'ejecutivo', anterior, e.nombre);
    renombrarEnColumna('ACTIVIDADES', 'ejecutivo', anterior, e.nombre);
  }
  return { id:e.id };
}
function resetPassword(data, user){
  if(!esGerente(user)) throw new Error('Sólo Gerencia puede restablecer contraseñas.');
  var pass = String(data.password || '');
  if(pass.length < 6) throw new Error('La contraseña debe tener al menos 6 caracteres.');
  var lista = leer('EJECUTIVOS');
  for(var i=0;i<lista.length;i++){
    if(String(lista[i].id) === String(data.id)){
      var sh = hoja('EJECUTIVOS');
      var col = HOJAS.EJECUTIVOS.indexOf('hash') + 1;
      sh.getRange(i + 2, col).setValue(hashPass(pass));
      var colClara = HOJAS.EJECUTIVOS.indexOf('clave_visible') + 1;
      if(colClara > 0) sh.getRange(i + 2, colClara).setValue(pass);
      return { ok:true, nombre: lista[i].nombre };
    }
  }
  throw new Error('No se encontró el usuario.');
}
function deleteEjecutivo(data, user){
  if(!esGerente(user)) throw new Error('Sólo Gerencia puede eliminar usuarios.');
  if(String(data.id) === String(user.id)) throw new Error('No puedes eliminar tu propio usuario.');

  var e = buscarEjecutivoPorId(data.id);
  if(!e) throw new Error('No se encontró el usuario.');

  // La cartera no se pierde nunca: primero se traspasa.
  var conCartera = leer('CONTACTOS').filter(function(c){ return c.ejecutivo === e.nombre; }).length;
  if(conCartera){
    throw new Error('Este ejecutivo todavía tiene ' + conCartera +
      ' empresa(s) asignadas. Traspasa su cartera antes de eliminarlo.');
  }
  // Nunca dejar el sistema sin Gerencia.
  if(ROLES_MANDO.indexOf(e.rol) >= 0){
    var gerentes = leer('EJECUTIVOS').filter(function(x){ return ROLES_MANDO.indexOf(x.rol) >= 0; }).length;
    if(gerentes <= 1) throw new Error('No puedes eliminar al único usuario de Gerencia.');
  }
  borrarFila('EJECUTIVOS', data.id, user);
  return { eliminado: e.nombre };
}
function renombrarEnColumna(nombreHoja, columna, de, a){
  var sh = hoja(nombreHoja);
  var ultima = sh.getLastRow();
  if(ultima < 2) return;
  var col = HOJAS[nombreHoja].indexOf(columna) + 1;
  var rango = sh.getRange(2, col, ultima - 1, 1);
  var vals = rango.getValues();
  var cambio = false;
  for(var i=0;i<vals.length;i++){
    if(String(vals[i][0]) === de){ vals[i][0] = a; cambio = true; }
  }
  if(cambio) rango.setValues(vals);
}
function traspaso(data, user){
  if(!esGerente(user)) throw new Error('Sólo Gerencia puede traspasar cartera.');
  var ids = data.ids || [];
  if(!ids.length) throw new Error('No se recibieron empresas para traspasar.');

  var sh = hoja('CONTACTOS');
  var ultima = sh.getLastRow();
  if(ultima < 2) throw new Error('La base de contactos está vacía.');
  var colEj = HOJAS.CONTACTOS.indexOf('ejecutivo') + 1;
  var idsHoja = sh.getRange(2,1,ultima-1,1).getValues();
  var rangoEj = sh.getRange(2,colEj,ultima-1,1);
  var vals = rangoEj.getValues();
  var n = 0;
  for(var i=0;i<idsHoja.length;i++){
    if(ids.indexOf(String(idsHoja[i][0])) >= 0){ vals[i][0] = data.a_ejecutivo; n++; }
  }
  rangoEj.setValues(vals);

  var r = data.registro || {};
  guardarFila('TRASPASOS', {
    id: r.id, fecha: r.fecha || hoyISO(), de_ejecutivo: data.de_ejecutivo,
    a_ejecutivo: data.a_ejecutivo, n: n, motivo: r.motivo || '', nota: r.nota || '',
    autorizo: user.nombre, ids: ids.join('|'), creado_en: new Date().toISOString()
  });
  return { traspasadas: n };
}
function saveConfig(data, user){
  if(!esGerente(user)) throw new Error('Sólo Gerencia puede cambiar la configuración.');
  var sh = hoja('CONFIG');
  var ultima = sh.getLastRow();
  var valor = JSON.stringify((data.config && data.config.metas) || {});
  if(ultima >= 2){
    var claves = sh.getRange(2,1,ultima-1,1).getValues();
    for(var i=0;i<claves.length;i++){
      if(String(claves[i][0]) === 'metas'){
        sh.getRange(i+2,2).setValue(valor);
        return { ok:true };
      }
    }
  }
  sh.appendRow(['metas', valor]);
  return { ok:true };
}
function importBase(data, user){
  var lista = data.contactos || [];
  var sh = hoja('CONTACTOS');
  var existentes = {};
  leer('CONTACTOS').forEach(function(c){ existentes[normNombre(c.empresa)] = true; });
  var filas = [];
  lista.forEach(function(c){
    if(existentes[normNombre(c.empresa)]) return;
    if(!permitido(c, user)) return;
    filas.push(filaDe('CONTACTOS', contactoAFila(c, user)));
    existentes[normNombre(c.empresa)] = true;
  });
  if(filas.length){
    sh.getRange(sh.getLastRow()+1, 1, filas.length, HOJAS.CONTACTOS.length).setValues(filas);
  }
  return { n: filas.length };
}
function normNombre(s){
  return String(s||'').toLowerCase().replace(/\s+/g,' ').trim();
}
