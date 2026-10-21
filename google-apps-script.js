// ═══════════════════════════════════════════════════════════
// GOOGLE APPS SCRIPT — Ruleta Salentino x Ritvales 2026
// ═══════════════════════════════════════════════════════════
//
// INSTRUCCIONES:
// 1. Abrir el Google Sheet existente (o crear uno nuevo)
// 2. En la primera fila poner estos headers:
//    Timestamp | Nombre | Email | WhatsApp | Codigo | Premio | Hora
// 3. Ir a Extensions > Apps Script
// 4. Borrar todo el contenido y pegar ESTE archivo completo
// 5. Click en Deploy > Manage deployments
// 6. Editar la deployment existente (o crear nueva si es la primera vez)
// 7. Version: "New version"
// 8. Execute as: "Me", Who has access: "Anyone"
// 9. Click Deploy — la URL NO cambia si editas la deployment existente
//
// ═══════════════════════════════════════════════════════════

function doPost(e) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    // Registro de jugada
    sheet.appendRow([
      new Date().toISOString(),
      data.name || '',
      data.email || '',
      data.phone || '',
      data.code || '',
      data.prize || '',
      data.time || ''
    ]);

    lock.releaseLock();

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var params = e.parameter || {};
    var action = params.action || '';

    if (action === 'counts') {
      // Devuelve premios entregados en una fecha + todos los códigos usados
      var dateFilter = params.date || new Date().toISOString().slice(0, 10);
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var rows = sheet.getDataRange().getValues();

      var prizes = {};
      var usedCodes = [];

      for (var i = 1; i < rows.length; i++) {
        var timestamp = String(rows[i][0]);
        var codigo = String(rows[i][4]).trim().toUpperCase();
        var premio = String(rows[i][5]).trim();

        // Recopilar TODOS los códigos usados (cualquier fecha)
        if (codigo && codigo !== '') {
          usedCodes.push(codigo);
        }

        // Contar premios solo de la fecha solicitada
        var rowDate = timestamp.slice(0, 10);
        if (rowDate === dateFilter && premio && premio !== '') {
          prizes[premio] = (prizes[premio] || 0) + 1;
        }
      }

      // Eliminar duplicados de usedCodes
      var uniqueCodes = [];
      var seen = {};
      for (var j = 0; j < usedCodes.length; j++) {
        if (!seen[usedCodes[j]]) {
          seen[usedCodes[j]] = true;
          uniqueCodes.push(usedCodes[j]);
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        date: dateFilter,
        prizes: prizes,
        usedCodes: uniqueCodes
      })).setMimeType(ContentService.MimeType.JSON);

    } else if (action === 'validate') {
      // Verificar si un código ya fue usado
      var code = (params.code || '').toUpperCase();
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var rows = sheet.getDataRange().getValues();
      var used = false;

      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][4]).trim().toUpperCase() === code) {
          used = true;
          break;
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        code: code,
        used: used
      })).setMimeType(ContentService.MimeType.JSON);

    } else {
      // Default: status
      return ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        message: 'Ruleta Salentino x Ritvales 2026 API activa'
      })).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
