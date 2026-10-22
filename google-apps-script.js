// ═══════════════════════════════════════════════════════════
// GOOGLE APPS SCRIPT — Ruleta Salentino x Ritvales 2026
// Fase 1: Server-side security + numero column
// ═══════════════════════════════════════════════════════════
//
// SETUP:
// 1. Google Sheet con TRES pestañas:
//    - "log"     → Timestamp | Nombre | Email | WhatsApp | Codigo | Premio | Hora
//    - "codigos" → numero | codigo | usado | premio | nombre | email | whatsapp | ganado_en | canjeado_en | cedula
//    - "Resumen" → Dashboard con fórmulas (solo lectura)
// 2. Script Properties → agregar: ADMIN_PIN = 2610
// 3. Ejecutar initCodes() una vez desde el editor de Apps Script
// 4. Deploy como Web App: Execute as Me, Anyone can access
//
// ═══════════════════════════════════════════════════════════

// ─── CONFIG ───
var CUPO_CAFE     = 30;
var CUPO_FRIA     = 15;
var CUPO_COLLAR   = 10;
var CUPO_DESC20   = 8;
var CUPO_ENTRADA  = 3;

var PRIZE_CONFIG = [
  { id: 'cafe',        weight: 30, daily_limit: CUPO_CAFE },
  { id: 'fria',        weight: 15, daily_limit: CUPO_FRIA },
  { id: 'collar',      weight: 10, daily_limit: CUPO_COLLAR },
  { id: 'desc20',      weight: 8,  daily_limit: CUPO_DESC20 },
  { id: 'entrada',     weight: 3,  daily_limit: CUPO_ENTRADA },
  { id: 'consolacion', weight: 0,  daily_limit: 9999 },
];

// ─── HELPERS ───
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function colTimestamp() {
  var now = new Date();
  var col = new Date(now.getTime() - 5 * 3600000 + now.getTimezoneOffset() * 60000);
  var pad = function(n) { return n < 10 ? '0' + n : '' + n; };
  return col.getFullYear() + '-' + pad(col.getMonth()+1) + '-' + pad(col.getDate()) + ' ' + pad(col.getHours()) + ':' + pad(col.getMinutes()) + ':' + pad(col.getSeconds());
}

function colToday() {
  return colTimestamp().slice(0, 10);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getAdminPin() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || '2610';
}

// ─── CODE NORMALIZATION ───
function normalizeCode(raw) {
  var c = (raw || '').trim().toUpperCase().replace(/\s/g, '');
  // Add RTV- prefix if missing
  if (c.length === 4 && /^[A-Z0-9]{4}$/.test(c)) c = 'RTV-' + c;
  if (c.length === 7 && c.indexOf('-') === -1) c = 'RTV-' + c;
  // O↔0 and I↔1 equivalence
  c = c.replace(/O/g, '0').replace(/I/g, '1');
  return c;
}

function normalizeStored(code) {
  return String(code).trim().toUpperCase().replace(/O/g, '0').replace(/I/g, '1');
}

// ─── RATE LIMITING ───
function checkRateLimit(fingerprint) {
  if (!fingerprint) return false;
  var cache = CacheService.getScriptCache();
  var key = 'rl_' + fingerprint;
  var val = cache.get(key);
  if (val) return true; // rate limited
  cache.put(key, '1', 10); // 10 second cooldown
  return false;
}

// ─── PRIZE SELECTION (server-side) ───
function countPrizesToday(logSheet, today) {
  var counts = {};
  if (!logSheet) return counts;
  var rows = logSheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var ts = String(rows[i][0]);
    var prize = String(rows[i][5]).trim();
    if (ts.slice(0, 10) === today && prize) {
      counts[prize] = (counts[prize] || 0) + 1;
    }
  }
  return counts;
}

function pickPrize(logSheet, today) {
  var counts = countPrizesToday(logSheet, today);

  // Filter available prizes (not consolacion, not exhausted)
  var available = PRIZE_CONFIG.filter(function(p) {
    if (p.id === 'consolacion') return false;
    var used = counts[p.id] || 0;
    return used < p.daily_limit;
  });

  if (available.length === 0) {
    return 'consolacion';
  }

  // Weighted random selection
  var total = available.reduce(function(s, p) { return s + p.weight; }, 0);
  var r = Math.random() * total;
  for (var i = 0; i < available.length; i++) {
    r -= available[i].weight;
    if (r <= 0) return available[i].id;
  }
  return available[available.length - 1].id;
}

// ─── POST HANDLER ───
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || '';

    if (action === 'spin') return handleSpin(data);
    if (action === 'redeem') return handleRedeem(data);
    if (action === 'claim') return handleClaim(data);

    return jsonResponse({ status: 'error', error: 'UNKNOWN_ACTION' });
  } catch (error) {
    return jsonResponse({ status: 'error', error: 'SERVER_ERROR', message: error.toString() });
  }
}

// ─── ACTION: SPIN ───
function handleSpin(data) {
  var code = normalizeCode(data.code || '');
  var name = (data.name || '').trim();
  var email = (data.email || '').trim();
  var phone = (data.phone || '').replace(/\D/g, '');
  var fingerprint = data.fingerprint || '';

  // Rate limit check
  if (checkRateLimit(fingerprint)) {
    return jsonResponse({ status: 'error', error: 'RATE_LIMIT' });
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var codigosSheet = getSheet('codigos');
    var logSheet = getSheet('log');

    if (!codigosSheet || !logSheet) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'SERVER_ERROR', message: 'Sheets not found' });
    }

    // Find code in codigos sheet
    // Columns: A=numero, B=codigo, C=usado, D=premio, E=nombre, F=email, G=whatsapp, H=ganado_en, I=canjeado_en, J=cedula
    var codigosData = codigosSheet.getDataRange().getValues();
    var codeRow = -1;
    for (var i = 1; i < codigosData.length; i++) {
      if (normalizeStored(codigosData[i][1]) === code) {
        codeRow = i + 1; // 1-indexed for Sheet
        break;
      }
    }

    if (codeRow === -1) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'NOT_FOUND' });
    }

    // Check if already used
    var usedFlag = String(codigosData[codeRow - 1][2]).trim().toUpperCase();
    if (usedFlag === 'SI' || usedFlag === 'SÍ' || usedFlag === 'TRUE' || usedFlag === '1') {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'USED' });
    }

    // Select prize server-side
    var today = colToday();
    var prize = pickPrize(logSheet, today);
    var ts = colTimestamp();

    // Mark code as used in codigos sheet
    codigosSheet.getRange(codeRow, 3).setValue('SI');      // C=usado
    codigosSheet.getRange(codeRow, 4).setValue(prize);     // D=premio
    codigosSheet.getRange(codeRow, 5).setValue(name);      // E=nombre
    codigosSheet.getRange(codeRow, 6).setValue(email);     // F=email
    codigosSheet.getRange(codeRow, 7).setValue(phone);     // G=whatsapp
    codigosSheet.getRange(codeRow, 8).setValue(ts);        // H=ganado_en

    // Log the play
    logSheet.appendRow([ts, name, email, phone, code, prize, ts]);

    lock.releaseLock();

    return jsonResponse({ status: 'ok', prize: prize, ganado_en: ts });

  } catch (error) {
    try { lock.releaseLock(); } catch(e2) {}
    return jsonResponse({ status: 'error', error: 'SERVER_ERROR', message: error.toString() });
  }
}

// ─── ACTION: REDEEM (carrito/physical prizes — staff marks as delivered) ───
function handleRedeem(data) {
  var code = normalizeCode(data.code || '');
  var pin = (data.pin || '').trim();

  if (pin !== getAdminPin()) {
    return jsonResponse({ status: 'error', error: 'BAD_PIN' });
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var codigosSheet = getSheet('codigos');
    var codigosData = codigosSheet.getDataRange().getValues();
    var codeRow = -1;
    for (var i = 1; i < codigosData.length; i++) {
      if (normalizeStored(codigosData[i][1]) === code) {
        codeRow = i + 1;
        break;
      }
    }

    if (codeRow === -1) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'NOT_FOUND' });
    }

    // Check if already redeemed (canjeado_en = col I, index 8)
    var canjeadoEn = String(codigosData[codeRow - 1][8] || '').trim();
    if (canjeadoEn) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'ALREADY_REDEEMED', canjeado_en: canjeadoEn });
    }

    // Check expiry for carrito prizes
    var premio = String(codigosData[codeRow - 1][3] || '').trim();
    var ganado = String(codigosData[codeRow - 1][7] || '').trim();
    var carritoPrizes = ['cafe', 'fria', 'collar', 'consolacion'];
    if (carritoPrizes.indexOf(premio) >= 0 && ganado.slice(0, 10) !== colToday()) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'EXPIRED', ganado_en: ganado });
    }

    // Mark as redeemed
    var ts = colTimestamp();
    codigosSheet.getRange(codeRow, 9).setValue(ts); // I=canjeado_en

    lock.releaseLock();
    return jsonResponse({ status: 'ok', redeemed: true, canjeado_en: ts });

  } catch (error) {
    try { lock.releaseLock(); } catch(e2) {}
    return jsonResponse({ status: 'error', error: 'SERVER_ERROR', message: error.toString() });
  }
}

// ─── ACTION: CLAIM (tuboleta prizes — user submits cedula) ───
function handleClaim(data) {
  var code = normalizeCode(data.code || '');
  var cedula = (data.cedula || '').trim();
  var claimName = (data.name || '').trim();
  var claimEmail = (data.email || '').trim();

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var codigosSheet = getSheet('codigos');
    var codigosData = codigosSheet.getDataRange().getValues();
    var codeRow = -1;
    for (var i = 1; i < codigosData.length; i++) {
      if (normalizeStored(codigosData[i][1]) === code) {
        codeRow = i + 1;
        break;
      }
    }

    if (codeRow === -1) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'NOT_FOUND' });
    }

    // Check if already claimed (cedula = col J, index 9)
    var existingCedula = String(codigosData[codeRow - 1][9] || '').trim();
    if (existingCedula) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'ALREADY_REDEEMED' });
    }

    // Save claim data
    codigosSheet.getRange(codeRow, 10).setValue(cedula);        // J=cedula
    codigosSheet.getRange(codeRow, 9).setValue(colTimestamp());  // I=canjeado_en
    if (claimName) codigosSheet.getRange(codeRow, 5).setValue(claimName);   // E=nombre
    if (claimEmail) codigosSheet.getRange(codeRow, 6).setValue(claimEmail); // F=email

    lock.releaseLock();
    return jsonResponse({ status: 'ok', claimed: true });

  } catch (error) {
    try { lock.releaseLock(); } catch(e2) {}
    return jsonResponse({ status: 'error', error: 'SERVER_ERROR', message: error.toString() });
  }
}

// ─── GET HANDLER ───
function doGet(e) {
  try {
    var params = e.parameter || {};
    var action = params.action || '';

    if (action === 'status') return handleStatus(params);
    if (action === 'lookup') return handleLookup(params);

    // Default: API status
    return jsonResponse({
      status: 'ok',
      message: 'Ruleta Salentino x Ritvales 2026 API activa'
    });

  } catch (error) {
    return jsonResponse({ status: 'error', error: 'SERVER_ERROR', message: error.toString() });
  }
}

// ─── ACTION: STATUS ───
function handleStatus(params) {
  var today = params.date || colToday();
  var logSheet = getSheet('log');
  var counts = logSheet ? countPrizesToday(logSheet, today) : {};

  // Also get used codes for client cache sync
  // Columns: A=numero, B=codigo(index 1), C=usado(index 2)
  var codigosSheet = getSheet('codigos');
  var usedCodes = [];
  if (codigosSheet) {
    var data = codigosSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var usedFlag = String(data[i][2]).trim().toUpperCase();
      if (usedFlag === 'SI' || usedFlag === 'SÍ' || usedFlag === 'TRUE' || usedFlag === '1') {
        usedCodes.push(normalizeStored(data[i][1]));
      }
    }
  }

  return jsonResponse({
    status: 'ok',
    date: today,
    prizes: counts,
    usedCodes: usedCodes
  });
}

// ─── ACTION: LOOKUP ───
function handleLookup(params) {
  var code = normalizeCode(params.code || '');
  var codigosSheet = getSheet('codigos');
  if (!codigosSheet) return jsonResponse({ status: 'error', error: 'SERVER_ERROR' });
  var data = codigosSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeStored(data[i][1]) === code) {
      var usado = String(data[i][2]).trim().toUpperCase();
      if (usado === 'SI' || usado === 'SÍ' || usado === 'TRUE' || usado === '1') {
        return jsonResponse({
          status: 'ok',
          used: true,
          prize: String(data[i][3]).trim(),
          name: String(data[i][4]).trim(),
          ganado_en: String(data[i][7]).trim(),
          canjeado_en: String(data[i][8]).trim()
        });
      }
      return jsonResponse({ status: 'ok', used: false });
    }
  }
  return jsonResponse({ status: 'error', error: 'NOT_FOUND' });
}

// ─── UTILITY: Initialize codes sheet ───
// Run this once manually from Apps Script editor to populate the codigos sheet
function initCodes() {
  var codes = ["RTV-04OG","RTV-0E09","RTV-18YW","RTV-1LXA","RTV-1VIW","RTV-26LO","RTV-27GJ","RTV-2O2N","RTV-2OQL","RTV-2R6M","RTV-2SR5","RTV-2T01","RTV-2VXW","RTV-2XDN","RTV-34I8","RTV-3GI8","RTV-3HYM","RTV-3J1F","RTV-3SHA","RTV-42U6","RTV-4L9O","RTV-4OHB","RTV-4PRV","RTV-4Y8M","RTV-50NW","RTV-576O","RTV-5DWQ","RTV-5ET2","RTV-5GJ8","RTV-5SND","RTV-5XO2","RTV-62KN","RTV-63L7","RTV-648Y","RTV-6XMM","RTV-6YOS","RTV-70AU","RTV-7BT4","RTV-7LHQ","RTV-7RSA","RTV-7YLJ","RTV-834Z","RTV-868X","RTV-8BDX","RTV-8QVV","RTV-8RZE","RTV-8SJW","RTV-8U2F","RTV-90E3","RTV-95Y5","RTV-9729","RTV-98I3","RTV-99GS","RTV-9B8X","RTV-9OAH","RTV-9Q8Y","RTV-9UEK","RTV-AAU7","RTV-ADAE","RTV-ADVZ","RTV-AH4U","RTV-AS54","RTV-AVS8","RTV-BGFX","RTV-BHE4","RTV-BQFI","RTV-BRA3","RTV-BV2F","RTV-BVC7","RTV-C2IW","RTV-CZWF","RTV-D0MJ","RTV-D0OQ","RTV-D3Q3","RTV-D5BM","RTV-D60E","RTV-DG6H","RTV-DIVU","RTV-DJ8G","RTV-DSSL","RTV-DY2S","RTV-DZUU","RTV-E1L6","RTV-EDDT","RTV-ENHC","RTV-ES1V","RTV-ESS4","RTV-EWEE","RTV-F0XE","RTV-FLU1","RTV-FLZB","RTV-FUUT","RTV-G099","RTV-G0AV","RTV-G3PP","RTV-GODD","RTV-GTF5","RTV-GWF6","RTV-GXG8","RTV-HF9P","RTV-HKJU","RTV-HQVG","RTV-IEBZ","RTV-IV85","RTV-J1HJ","RTV-J2S6","RTV-J4L4","RTV-J5RU","RTV-J80C","RTV-JAHQ","RTV-JFAJ","RTV-JMXO","RTV-JWCY","RTV-KAWB","RTV-KLJC","RTV-KMJR","RTV-KQ5Z","RTV-KTM9","RTV-LKLE","RTV-LMG9","RTV-LSL6","RTV-LUZY","RTV-M0PA","RTV-M1AS","RTV-MHMW","RTV-MXXB","RTV-NNHF","RTV-NTQ0","RTV-O61O","RTV-OCZX","RTV-OPOC","RTV-P1VV","RTV-P4Z8","RTV-PHET","RTV-PKJB","RTV-POVR","RTV-Q1GZ","RTV-Q884","RTV-QFWI","RTV-QOAZ","RTV-QQ79","RTV-QRA6","RTV-R7JP","RTV-R7JR","RTV-RA8N","RTV-RCJR","RTV-RE22","RTV-RE2U","RTV-RI5W","RTV-RJL7","RTV-RL1I","RTV-RP36","RTV-RTBH","RTV-S33U","RTV-S3YQ","RTV-S5XA","RTV-S8F4","RTV-SCMU","RTV-SDY7","RTV-SJP0","RTV-SPMV","RTV-SSS2","RTV-STDJ","RTV-SXFS","RTV-TADL","RTV-TB89","RTV-THML","RTV-TMQG","RTV-TMUZ","RTV-TR2U","RTV-TUYL","RTV-U8N8","RTV-UAZN","RTV-UDJP","RTV-UMNI","RTV-UNXU","RTV-V4TX","RTV-VC2D","RTV-VE86","RTV-VE9R","RTV-VHQH","RTV-VJ54","RTV-WP8V","RTV-WSJ4","RTV-WTFT","RTV-X7YO","RTV-XCXB","RTV-XM6E","RTV-XVQS","RTV-Y10G","RTV-Y1KR","RTV-Y4S7","RTV-Y7O8","RTV-YCOH","RTV-YMY7","RTV-YQM3","RTV-YU7M","RTV-YYHJ","RTV-ZK82","RTV-ZLQL"];

  var sheet = getSheet('codigos');
  if (!sheet) {
    SpreadsheetApp.getActiveSpreadsheet().insertSheet('codigos');
    sheet = getSheet('codigos');
  }

  // Set headers (10 columns: numero first)
  sheet.getRange(1, 1, 1, 10).setValues([['numero', 'codigo', 'usado', 'premio', 'nombre', 'email', 'whatsapp', 'ganado_en', 'canjeado_en', 'cedula']]);

  // Write codes with numero 001-200
  var pad3 = function(n) { return n < 10 ? '00' + n : n < 100 ? '0' + n : '' + n; };
  var rows = codes.map(function(c, idx) { return [pad3(idx + 1), c, '', '', '', '', '', '', '', '']; });
  sheet.getRange(2, 1, rows.length, 10).setValues(rows);

  // Also ensure log sheet exists
  var logSheet = getSheet('log');
  if (!logSheet) {
    SpreadsheetApp.getActiveSpreadsheet().insertSheet('log');
    logSheet = getSheet('log');
  }
  logSheet.getRange(1, 1, 1, 7).setValues([['Timestamp', 'Nombre', 'Email', 'WhatsApp', 'Codigo', 'Premio', 'Hora']]);

  // Create Resumen dashboard
  initResumen();
}

// ─── UTILITY: Initialize Resumen dashboard ───
// Creates a "Resumen" sheet with formulas that auto-update from codigos
function initResumen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Resumen');
  if (!sheet) {
    sheet = ss.insertSheet('Resumen');
  } else {
    sheet.clear();
  }

  // ── Section 1: Códigos jugados ──
  sheet.getRange('A1').setValue('RESUMEN RULETA SALENTINO x RITVALES 2026').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A3').setValue('CÓDIGOS JUGADOS').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A4').setValue('Jugados');
  sheet.getRange('B4').setFormula('=COUNTIF(codigos!C2:C201,"SI")');
  sheet.getRange('A5').setValue('Disponibles');
  sheet.getRange('B5').setFormula('=200-COUNTIF(codigos!C2:C201,"SI")');

  // ── Section 2: Premios ganados por tipo vs cupo ──
  sheet.getRange('A7').setValue('PREMIOS POR TIPO').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A8').setValue('Premio');
  sheet.getRange('B8').setValue('Ganados');
  sheet.getRange('C8').setValue('Cupo');
  sheet.getRange('D8').setValue('Restantes');
  sheet.getRange('A8:D8').setFontWeight('bold');

  var prizes = [
    { name: 'Aeropress',          id: 'cafe',        cupo: CUPO_CAFE },
    { name: 'Bebida fría',        id: 'fria',        cupo: CUPO_FRIA },
    { name: 'Collar Ritvales',    id: 'collar',      cupo: CUPO_COLLAR },
    { name: '20% boletería',      id: 'desc20',      cupo: CUPO_DESC20 },
    { name: 'Entrada general',    id: 'entrada',     cupo: CUPO_ENTRADA },
    { name: '15% café',           id: 'consolacion', cupo: 9999 },
  ];

  for (var i = 0; i < prizes.length; i++) {
    var row = 9 + i;
    sheet.getRange('A' + row).setValue(prizes[i].name);
    sheet.getRange('B' + row).setFormula('=COUNTIF(codigos!D2:D201,"' + prizes[i].id + '")');
    sheet.getRange('C' + row).setValue(prizes[i].cupo);
    if (prizes[i].id !== 'consolacion') {
      sheet.getRange('D' + row).setFormula('=C' + row + '-B' + row);
    } else {
      sheet.getRange('D' + row).setValue('∞');
    }
  }

  // ── Section 3: Premios del carrito pendientes hoy ──
  var secRow = 9 + prizes.length + 1;
  sheet.getRange('A' + secRow).setValue('PREMIOS DEL CARRITO PENDIENTES HOY').setFontWeight('bold').setFontSize(11);
  secRow++;
  sheet.getRange('A' + secRow).setValue('Premio');
  sheet.getRange('B' + secRow).setValue('Pendientes');
  sheet.getRange('A' + secRow + ':B' + secRow).setFontWeight('bold');

  var carritoPrizes = ['cafe', 'fria', 'collar', 'consolacion'];
  for (var j = 0; j < carritoPrizes.length; j++) {
    secRow++;
    var pName = prizes.filter(function(p) { return p.id === carritoPrizes[j]; })[0].name;
    sheet.getRange('A' + secRow).setValue(pName);
    // Ganado (premio = X) AND usado = SI AND canjeado_en está vacío
    sheet.getRange('B' + secRow).setFormula('=COUNTIFS(codigos!D2:D201,"' + carritoPrizes[j] + '",codigos!I2:I201,"",codigos!C2:C201,"SI")');
  }

  // ── Section 4: Últimas 10 jugadas ──
  secRow += 2;
  sheet.getRange('A' + secRow).setValue('ÚLTIMAS 10 JUGADAS').setFontWeight('bold').setFontSize(11);
  secRow++;
  sheet.getRange('A' + secRow).setValue('Número');
  sheet.getRange('B' + secRow).setValue('Código');
  sheet.getRange('C' + secRow).setValue('Premio');
  sheet.getRange('D' + secRow).setValue('Ganado');
  sheet.getRange('E' + secRow).setValue('Canjeado');
  sheet.getRange('A' + secRow + ':E' + secRow).setFontWeight('bold');
  secRow++;

  // QUERY: last 10 played codes sorted by ganado_en descending
  var queryFormula = '=IFERROR(QUERY(codigos!A2:J201,"SELECT A,B,D,H,I WHERE C=\'SI\' ORDER BY H DESC LIMIT 10",0),"Sin jugadas aún")';
  sheet.getRange('A' + secRow).setFormula(queryFormula);

  // ── Formatting ──
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 160);

  // Protect sheet as read-only
  var protection = sheet.protect().setDescription('Resumen - Solo lectura');
  protection.setWarningOnly(true);
}
