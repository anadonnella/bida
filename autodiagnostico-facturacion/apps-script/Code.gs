/**
 * Autodiagnóstico de preparación - Facturación digital en Venezuela (BIDA)
 * Recibe las respuestas de la página (GitHub Pages) y las guarda en un Google Sheet
 * dentro de la carpeta "AUTODIAGNOSTICO TRIBUTARIO FACTURACION DIGITAL".
 *
 * Uso: ejecute setup() UNA vez (crea la carpeta y la hoja) y luego implemente como aplicación web.
 */
var FOLDER_NAME = 'AUTODIAGNOSTICO TRIBUTARIO FACTURACION DIGITAL';
var SHEET_FILE  = 'Respuestas Autodiagnóstico Facturación Digital';
var TAB         = 'Respuestas';
var NOTIFY_EMAIL = 'consultoria@bidaglobal.com'; // aviso cuando alguien pide ser contactado
var NQ = 17;

function headers_() {
  var h = ['Fecha', 'ID', 'Sector', 'Tamaño', 'Medio actual'];
  for (var i = 1; i <= NQ; i++) h.push('P' + (i < 10 ? '0' : '') + i);
  return h.concat(['% Medio de emisión', '% Contingencia', '% Datos maestros', '% Sistema y gobierno',
    '% Total', 'Semáforo', 'Correo', 'Nombre', 'Teléfono', 'Desea contacto', 'Fecha del correo']);
}

function setup() {
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
  var files = folder.getFilesByName(SHEET_FILE);
  var ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SHEET_FILE);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
  }
  var sh = ss.getSheetByName(TAB) || ss.insertSheet(TAB);
  var def = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  var h = headers_();
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#10283F').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  Logger.log('Carpeta: ' + folder.getUrl());
  Logger.log('Hoja: ' + ss.getUrl());
}

function sheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Ejecute setup() primero.');
  return SpreadsheetApp.openById(id).getSheetByName(TAB);
}

function safe_(v) {            // evita que el texto se interprete como fórmula
  v = (v === null || v === undefined) ? '' : String(v).slice(0, 300);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function doGet() {
  return ContentService.createTextOutput('Autodiagnóstico BIDA: servicio activo');
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var p = JSON.parse(e.postData.contents);
    if (p.hp) return ok_();                       // trampa para bots
    var sh = sheet_();
    var col = headers_();
    var row = findRow_(sh, p.id);
    if (p.action === 'submit') {
      if (row) return ok_();                      // ya registrado
      var r = [p.ts ? new Date(p.ts) : new Date(), safe_(p.id), safe_(p.sector), safe_(p.size), safe_(p.medio)];
      for (var i = 0; i < NQ; i++) r.push(safe_((p.answers || [])[i]));
      for (var b = 0; b < 4; b++) r.push((p.blocks || [])[b] === undefined ? '' : Number(p.blocks[b]) / 100);
      r.push(p.total === '' ? '' : Number(p.total) / 100, safe_(p.light), '', '', '', '', '');
      sh.appendRow(r);
      var n = sh.getLastRow();
      sh.getRange(n, 6 + NQ, 1, 5).setNumberFormat('0%');
    } else if (p.action === 'lead') {
      if (!row) {                                 // llegó el correo antes que las respuestas
        var base = new Array(col.length).fill('');
        base[0] = new Date(); base[1] = safe_(p.id);
        sh.appendRow(base); row = sh.getLastRow();
      }
      var c0 = col.indexOf('Correo') + 1;
      sh.getRange(row, c0, 1, 5).setValues([[safe_(p.email), safe_(p.nombre), safe_(p.tel), safe_(p.contact), new Date()]]);
      if (p.contact === 'Sí' && NOTIFY_EMAIL) {
        MailApp.sendEmail(NOTIFY_EMAIL, 'Autodiagnóstico facturación digital: solicitud de contacto',
          'Una empresa pidió ser contactada.\n\nCorreo: ' + p.email + '\nNombre: ' + (p.nombre || '(no indicó)') +
          '\nTeléfono: ' + (p.tel || '(no indicó)') + '\nResultado: ' + p.total + '% (' + p.light + ')' +
          '\nSector: ' + (p.sector || '-') + ' | Tamaño: ' + (p.size || '-') + '\nID: ' + p.id);
      }
    }
    return ok_();
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}

function findRow_(sh, id) {
  if (!id) return 0;
  var f = sh.getRange(2, 2, Math.max(sh.getLastRow() - 1, 1), 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}

function ok_() { return ContentService.createTextOutput('ok'); }
