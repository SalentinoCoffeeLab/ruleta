// ═══════════════════════════════════════════════════════════
// GOOGLE APPS SCRIPT — Ruleta Salentino x Ritvales 2026
// Fase 2: Endpoints completos (spin/status/redeem/claim/board/ping/event)
// ═══════════════════════════════════════════════════════════
//
// SETUP:
// 1. Google Sheet con TRES pestañas:
//    - "log"     → Timestamp | Nombre | Email | WhatsApp | Codigo | Premio | Hora
//    - "codigos" → numero | codigo | usado | premio | nombre | email | whatsapp | ganado_en | canjeado_en | cedula | datos_enviados
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

// Expiry rules: mostrador (cafe/fria/collar) = same day, tuboleta (desc20/entrada) = 30 days, consolacion = 30 days
var EXPIRY_SAME_DAY = ['cafe', 'fria', 'collar'];
var EXPIRY_30_DAYS  = ['desc20', 'entrada', 'consolacion'];

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

function isTestMode() {
  return PropertiesService.getScriptProperties().getProperty('TEST_MODE') === 'true';
}

function isUsed(flag) {
  var f = String(flag).trim().toUpperCase();
  return f === 'SI' || f === 'SÍ' || f === 'TRUE' || f === '1';
}

function isExpired(premio, ganadoEn) {
  var ganDate = String(ganadoEn || '').slice(0, 10);
  if (!ganDate) return false;
  var today = colToday();
  if (EXPIRY_SAME_DAY.indexOf(premio) >= 0) {
    return ganDate !== today;
  }
  if (EXPIRY_30_DAYS.indexOf(premio) >= 0) {
    var gan = new Date(ganDate + 'T12:00:00Z');
    var now = new Date(today + 'T12:00:00Z');
    var diffDays = Math.floor((now - gan) / 86400000);
    return diffDays > 30;
  }
  return false;
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

function getRemainingQuotas(counts) {
  var remaining = {};
  for (var i = 0; i < PRIZE_CONFIG.length; i++) {
    var p = PRIZE_CONFIG[i];
    var used = counts[p.id] || 0;
    remaining[p.id] = Math.max(0, p.daily_limit - used);
  }
  return remaining;
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

// ─── FIND CODE ROW ───
// Returns { row: 1-indexed sheet row, index: 0-indexed data index } or null
function findCodeRow(codigosData, code) {
  for (var i = 1; i < codigosData.length; i++) {
    if (normalizeStored(codigosData[i][1]) === code) {
      return { row: i + 1, index: i };
    }
  }
  return null;
}

// ═══════════════════════════════════════════════════════════
// POST HANDLER
// ═══════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || '';

    if (action === 'spin') return handleSpin(data);
    if (action === 'redeem') return handleRedeem(data);
    if (action === 'claim') return handleClaim(data);
    if (action === 'event') return handleEvent(data);

    return jsonResponse({ status: 'error', error: 'UNKNOWN_ACTION' });
  } catch (error) {
    return jsonResponse({ status: 'error', error: 'SERVER_ERROR', message: error.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
// ACTION: SPIN
// ═══════════════════════════════════════════════════════════
function handleSpin(data) {
  var code = normalizeCode(data.code || '');
  var name = (data.name || '').trim();
  var email = (data.email || '').trim();
  var phone = (data.phone || '').replace(/\D/g, '');
  var fingerprint = data.fingerprint || '';
  var forzar = (data.forzar || '').trim();
  var testMode = isTestMode();

  // Rate limit check (skip for test codes)
  var isTestCode = testMode && code.indexOf('TEST-') === 0;
  if (!isTestCode && checkRateLimit(fingerprint)) {
    return jsonResponse({ status: 'error', error: 'RATE_LIMIT' });
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    // Test codes: use codigos_prueba sheet, reusable
    var codigosSheet;
    if (isTestCode) {
      codigosSheet = getSheet('codigos_prueba');
    } else {
      codigosSheet = getSheet('codigos');
    }
    var logSheet = getSheet('log');

    if (!codigosSheet || !logSheet) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'SERVER_ERROR', message: 'Sheets not found' });
    }

    var codigosData = codigosSheet.getDataRange().getValues();
    var found = findCodeRow(codigosData, code);

    if (!found) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'NOT_FOUND' });
    }

    // Check if already used (skip for test codes — they are reusable)
    if (!isTestCode && isUsed(codigosData[found.index][2])) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'USED' });
    }

    // Select prize server-side
    var today = colToday();
    var prize;
    // Force prize if TEST_MODE and forzar param provided
    if (testMode && forzar) {
      var valid = PRIZE_CONFIG.some(function(p) { return p.id === forzar; });
      prize = valid ? forzar : pickPrize(logSheet, today);
    } else {
      prize = pickPrize(logSheet, today);
    }
    var ts = colTimestamp();
    var counts = countPrizesToday(logSheet, today);
    counts[prize] = (counts[prize] || 0) + 1;
    var remaining = getRemainingQuotas(counts);
    var numero = String(codigosData[found.index][0]).trim();

    if (isTestCode) {
      // Test codes: reset for reuse, log but don't mark permanently
      codigosSheet.getRange(found.row, 4).setValue(prize);
      codigosSheet.getRange(found.row, 8).setValue(ts);
    } else {
      // Mark code as used in codigos sheet
      codigosSheet.getRange(found.row, 3).setValue('SI');      // C=usado
      codigosSheet.getRange(found.row, 4).setValue(prize);     // D=premio
      codigosSheet.getRange(found.row, 5).setValue(name);      // E=nombre
      codigosSheet.getRange(found.row, 6).setValue(email);     // F=email
      codigosSheet.getRange(found.row, 7).setValue(phone);     // G=whatsapp
      codigosSheet.getRange(found.row, 8).setValue(ts);        // H=ganado_en
    }

    // Log the play
    logSheet.appendRow([ts, name, email, phone, code, prize, ts]);

    lock.releaseLock();

    return jsonResponse({
      status: 'ok',
      prize: prize,
      ganado_en: ts,
      n: numero,
      cupo: remaining[prize] !== undefined ? remaining[prize] : null
    });

  } catch (error) {
    try { lock.releaseLock(); } catch(e2) {}
    return jsonResponse({ status: 'error', error: 'SERVER_ERROR', message: error.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
// ACTION: REDEEM (physical prizes — staff marks as delivered)
// ═══════════════════════════════════════════════════════════
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
    var found = findCodeRow(codigosData, code);

    if (!found) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'NOT_FOUND' });
    }

    // Check if already redeemed (canjeado_en = col I, index 8)
    var canjeadoEn = String(codigosData[found.index][8] || '').trim();
    if (canjeadoEn) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'ALREADY_REDEEMED', canjeado_en: canjeadoEn });
    }

    // Check expiry based on prize type
    var premio = String(codigosData[found.index][3] || '').trim();
    var ganado = String(codigosData[found.index][7] || '').trim();
    if (isExpired(premio, ganado)) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'EXPIRED', ganado_en: ganado });
    }

    // Mark as redeemed
    var ts = colTimestamp();
    codigosSheet.getRange(found.row, 9).setValue(ts); // I=canjeado_en

    lock.releaseLock();
    return jsonResponse({ status: 'ok', redeemed: true, canjeado_en: ts });

  } catch (error) {
    try { lock.releaseLock(); } catch(e2) {}
    return jsonResponse({ status: 'error', error: 'SERVER_ERROR', message: error.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
// ACTION: CLAIM (tuboleta prizes — user submits cedula)
// ═══════════════════════════════════════════════════════════
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
    var found = findCodeRow(codigosData, code);

    if (!found) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'NOT_FOUND' });
    }

    // Check if already claimed (datos_enviados = col K, index 10)
    var existingDatos = String(codigosData[found.index][10] || '').trim();
    if (existingDatos) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', error: 'ALREADY_CLAIMED' });
    }

    // Save claim data
    var ts = colTimestamp();
    codigosSheet.getRange(found.row, 10).setValue(cedula);    // J=cedula
    codigosSheet.getRange(found.row, 11).setValue(ts);        // K=datos_enviados
    if (claimName) codigosSheet.getRange(found.row, 5).setValue(claimName);   // E=nombre
    if (claimEmail) codigosSheet.getRange(found.row, 6).setValue(claimEmail); // F=email

    lock.releaseLock();
    return jsonResponse({ status: 'ok', claimed: true, datos_enviados: ts });

  } catch (error) {
    try { lock.releaseLock(); } catch(e2) {}
    return jsonResponse({ status: 'error', error: 'SERVER_ERROR', message: error.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
// ACTION: EVENT (analytics beacon)
// ═══════════════════════════════════════════════════════════
function handleEvent(data) {
  var logSheet = getSheet('log');
  if (!logSheet) return jsonResponse({ status: 'ok' });

  var ts = colTimestamp();
  var eventName = String(data.event || '').slice(0, 50);
  var code = String(data.code || '').slice(0, 10);
  var extra = String(data.extra || '').slice(0, 200);

  // Log events as a special row: timestamp, event type, code, extra, '', 'EVENT', timestamp
  logSheet.appendRow([ts, eventName, code, extra, '', 'EVENT', ts]);

  return jsonResponse({ status: 'ok' });
}

// ═══════════════════════════════════════════════════════════
// GET HANDLER
// ═══════════════════════════════════════════════════════════
function doGet(e) {
  try {
    var params = e.parameter || {};
    var action = params.action || '';

    if (action === 'status') return handleStatus(params);
    if (action === 'lookup') return handleLookup(params);
    if (action === 'board')  return handleBoard(params);
    if (action === 'ping')   return handlePing();

    // Default: API status
    return jsonResponse({
      status: 'ok',
      message: 'Ruleta Salentino x Ritvales 2026 API activa'
    });

  } catch (error) {
    return jsonResponse({ status: 'error', error: 'SERVER_ERROR', message: error.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
// ACTION: STATUS
// ═══════════════════════════════════════════════════════════
function handleStatus(params) {
  var today = params.date || colToday();
  var logSheet = getSheet('log');
  var counts = logSheet ? countPrizesToday(logSheet, today) : {};
  var remaining = getRemainingQuotas(counts);

  // Also get used codes for client cache sync
  var codigosSheet = getSheet('codigos');
  var usedCodes = [];
  if (codigosSheet) {
    var data = codigosSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (isUsed(data[i][2])) {
        usedCodes.push(normalizeStored(data[i][1]));
      }
    }
  }

  return jsonResponse({
    status: 'ok',
    date: today,
    prizes: counts,
    remaining: remaining,
    usedCodes: usedCodes
  });
}

// ═══════════════════════════════════════════════════════════
// ACTION: LOOKUP
// ═══════════════════════════════════════════════════════════
function handleLookup(params) {
  var code = normalizeCode(params.code || '');
  var codigosSheet = getSheet('codigos');
  if (!codigosSheet) return jsonResponse({ status: 'error', error: 'SERVER_ERROR' });
  var data = codigosSheet.getDataRange().getValues();
  var found = findCodeRow(data, code);

  if (!found) {
    return jsonResponse({ status: 'error', error: 'NOT_FOUND' });
  }

  var row = data[found.index];
  if (isUsed(row[2])) {
    var premio = String(row[3]).trim();
    var ganadoEn = String(row[7]).trim();
    return jsonResponse({
      status: 'ok',
      used: true,
      prize: premio,
      name: String(row[4]).trim(),
      n: String(row[0]).trim(),
      ganado_en: ganadoEn,
      canjeado_en: String(row[8]).trim(),
      datos_enviados: String(row[10] || '').trim(),
      expired: isExpired(premio, ganadoEn)
    });
  }
  return jsonResponse({ status: 'ok', used: false });
}

// ═══════════════════════════════════════════════════════════
// ACTION: BOARD (tablero — last 8 plays + remaining quotas)
// ═══════════════════════════════════════════════════════════
function handleBoard(params) {
  var today = params.date || colToday();
  var codigosSheet = getSheet('codigos');
  var logSheet = getSheet('log');

  if (!codigosSheet) return jsonResponse({ status: 'error', error: 'SERVER_ERROR' });

  var counts = logSheet ? countPrizesToday(logSheet, today) : {};
  var remaining = getRemainingQuotas(counts);

  // Get last 8 played codes sorted by ganado_en descending
  var data = codigosSheet.getDataRange().getValues();
  var played = [];
  for (var i = 1; i < data.length; i++) {
    if (isUsed(data[i][2])) {
      played.push({
        n: String(data[i][0]).trim(),
        premio: String(data[i][3]).trim(),
        nombre: String(data[i][4]).trim(),
        ganado_en: String(data[i][7]).trim()
      });
    }
  }

  // Sort by ganado_en descending
  played.sort(function(a, b) {
    return a.ganado_en > b.ganado_en ? -1 : a.ganado_en < b.ganado_en ? 1 : 0;
  });

  // Take last 8
  var last8 = played.slice(0, 8);

  return jsonResponse({
    status: 'ok',
    date: today,
    remaining: remaining,
    total_played: played.length,
    last: last8
  });
}

// ═══════════════════════════════════════════════════════════
// ACTION: PING (health check)
// ═══════════════════════════════════════════════════════════
function handlePing() {
  return jsonResponse({
    status: 'ok',
    ts: colTimestamp(),
    date: colToday()
  });
}

// ═══════════════════════════════════════════════════════════
// UTILITY: Initialize codes sheet
// Run this once manually from Apps Script editor
// ═══════════════════════════════════════════════════════════
function initCodes() {
  var codes = ["RTV-04OG","RTV-0E09","RTV-18YW","RTV-1LXA","RTV-1VIW","RTV-26LO","RTV-27GJ","RTV-2O2N","RTV-2OQL","RTV-2R6M","RTV-2SR5","RTV-2T01","RTV-2VXW","RTV-2XDN","RTV-34I8","RTV-3GI8","RTV-3HYM","RTV-3J1F","RTV-3SHA","RTV-42U6","RTV-4L9O","RTV-4OHB","RTV-4PRV","RTV-4Y8M","RTV-50NW","RTV-576O","RTV-5DWQ","RTV-5ET2","RTV-5GJ8","RTV-5SND","RTV-5XO2","RTV-62KN","RTV-63L7","RTV-648Y","RTV-6XMM","RTV-6YOS","RTV-70AU","RTV-7BT4","RTV-7LHQ","RTV-7RSA","RTV-7YLJ","RTV-834Z","RTV-868X","RTV-8BDX","RTV-8QVV","RTV-8RZE","RTV-8SJW","RTV-8U2F","RTV-90E3","RTV-95Y5","RTV-9729","RTV-98I3","RTV-99GS","RTV-9B8X","RTV-9OAH","RTV-9Q8Y","RTV-9UEK","RTV-AAU7","RTV-ADAE","RTV-ADVZ","RTV-AH4U","RTV-AS54","RTV-AVS8","RTV-BGFX","RTV-BHE4","RTV-BQFI","RTV-BRA3","RTV-BV2F","RTV-BVC7","RTV-C2IW","RTV-CZWF","RTV-D0MJ","RTV-D0OQ","RTV-D3Q3","RTV-D5BM","RTV-D60E","RTV-DG6H","RTV-DIVU","RTV-DJ8G","RTV-DSSL","RTV-DY2S","RTV-DZUU","RTV-E1L6","RTV-EDDT","RTV-ENHC","RTV-ES1V","RTV-ESS4","RTV-EWEE","RTV-F0XE","RTV-FLU1","RTV-FLZB","RTV-FUUT","RTV-G099","RTV-G0AV","RTV-G3PP","RTV-GODD","RTV-GTF5","RTV-GWF6","RTV-GXG8","RTV-HF9P","RTV-HKJU","RTV-HQVG","RTV-IEBZ","RTV-IV85","RTV-J1HJ","RTV-J2S6","RTV-J4L4","RTV-J5RU","RTV-J80C","RTV-JAHQ","RTV-JFAJ","RTV-JMXO","RTV-JWCY","RTV-KAWB","RTV-KLJC","RTV-KMJR","RTV-KQ5Z","RTV-KTM9","RTV-LKLE","RTV-LMG9","RTV-LSL6","RTV-LUZY","RTV-M0PA","RTV-M1AS","RTV-MHMW","RTV-MXXB","RTV-NNHF","RTV-NTQ0","RTV-O61O","RTV-OCZX","RTV-OPOC","RTV-P1VV","RTV-P4Z8","RTV-PHET","RTV-PKJB","RTV-POVR","RTV-Q1GZ","RTV-Q884","RTV-QFWI","RTV-QOAZ","RTV-QQ79","RTV-QRA6","RTV-R7JP","RTV-R7JR","RTV-RA8N","RTV-RCJR","RTV-RE22","RTV-RE2U","RTV-RI5W","RTV-RJL7","RTV-RL1I","RTV-RP36","RTV-RTBH","RTV-S33U","RTV-S3YQ","RTV-S5XA","RTV-S8F4","RTV-SCMU","RTV-SDY7","RTV-SJP0","RTV-SPMV","RTV-SSS2","RTV-STDJ","RTV-SXFS","RTV-TADL","RTV-TB89","RTV-THML","RTV-TMQG","RTV-TMUZ","RTV-TR2U","RTV-TUYL","RTV-U8N8","RTV-UAZN","RTV-UDJP","RTV-UMNI","RTV-UNXU","RTV-V4TX","RTV-VC2D","RTV-VE86","RTV-VE9R","RTV-VHQH","RTV-VJ54","RTV-WP8V","RTV-WSJ4","RTV-WTFT","RTV-X7YO","RTV-XCXB","RTV-XM6E","RTV-XVQS","RTV-Y10G","RTV-Y1KR","RTV-Y4S7","RTV-Y7O8","RTV-YCOH","RTV-YMY7","RTV-YQM3","RTV-YU7M","RTV-YYHJ","RTV-ZK82","RTV-ZLQL"];

  var sheet = getSheet('codigos');
  if (!sheet) {
    SpreadsheetApp.getActiveSpreadsheet().insertSheet('codigos');
    sheet = getSheet('codigos');
  }

  // Set headers (11 columns: numero + datos_enviados)
  sheet.getRange(1, 1, 1, 11).setValues([['numero', 'codigo', 'usado', 'premio', 'nombre', 'email', 'whatsapp', 'ganado_en', 'canjeado_en', 'cedula', 'datos_enviados']]);

  // Write codes with numero 001-200
  var pad3 = function(n) { return n < 10 ? '00' + n : n < 100 ? '0' + n : '' + n; };
  var rows = codes.map(function(c, idx) { return [pad3(idx + 1), c, '', '', '', '', '', '', '', '', '']; });
  sheet.getRange(2, 1, rows.length, 11).setValues(rows);

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

// ═══════════════════════════════════════════════════════════
// UTILITY: Initialize Resumen dashboard
// ═══════════════════════════════════════════════════════════
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

  // ── Section 3: Premios del mostrador pendientes hoy ──
  var secRow = 9 + prizes.length + 1;
  sheet.getRange('A' + secRow).setValue('PREMIOS MOSTRADOR PENDIENTES HOY').setFontWeight('bold').setFontSize(11);
  secRow++;
  sheet.getRange('A' + secRow).setValue('Premio');
  sheet.getRange('B' + secRow).setValue('Pendientes');
  sheet.getRange('A' + secRow + ':B' + secRow).setFontWeight('bold');

  var mostradorPrizes = ['cafe', 'fria', 'collar', 'consolacion'];
  for (var j = 0; j < mostradorPrizes.length; j++) {
    secRow++;
    var pName = prizes.filter(function(p) { return p.id === mostradorPrizes[j]; })[0].name;
    sheet.getRange('A' + secRow).setValue(pName);
    sheet.getRange('B' + secRow).setFormula('=COUNTIFS(codigos!D2:D201,"' + mostradorPrizes[j] + '",codigos!I2:I201,"",codigos!C2:C201,"SI")');
  }

  // ── Section 4: Datos tuboleta pendientes ──
  secRow += 2;
  sheet.getRange('A' + secRow).setValue('DATOS TUBOLETA PENDIENTES').setFontWeight('bold').setFontSize(11);
  secRow++;
  sheet.getRange('A' + secRow).setValue('Premio');
  sheet.getRange('B' + secRow).setValue('Sin datos');
  sheet.getRange('A' + secRow + ':B' + secRow).setFontWeight('bold');

  var tuboletaPrizes = ['desc20', 'entrada'];
  for (var k = 0; k < tuboletaPrizes.length; k++) {
    secRow++;
    var tName = prizes.filter(function(p) { return p.id === tuboletaPrizes[k]; })[0].name;
    sheet.getRange('A' + secRow).setValue(tName);
    // Premio = X AND usado = SI AND datos_enviados vacío
    sheet.getRange('B' + secRow).setFormula('=COUNTIFS(codigos!D2:D201,"' + tuboletaPrizes[k] + '",codigos!K2:K201,"",codigos!C2:C201,"SI")');
  }

  // ── Section 5: Últimas 10 jugadas ──
  secRow += 2;
  sheet.getRange('A' + secRow).setValue('ÚLTIMAS 10 JUGADAS').setFontWeight('bold').setFontSize(11);
  secRow++;
  sheet.getRange('A' + secRow).setValue('Número');
  sheet.getRange('B' + secRow).setValue('Código');
  sheet.getRange('C' + secRow).setValue('Premio');
  sheet.getRange('D' + secRow).setValue('Ganado');
  sheet.getRange('E' + secRow).setValue('Canjeado');
  sheet.getRange('F' + secRow).setValue('Datos env.');
  sheet.getRange('A' + secRow + ':F' + secRow).setFontWeight('bold');
  secRow++;

  // QUERY: last 10 played codes sorted by ganado_en descending
  var queryFormula = '=IFERROR(QUERY(codigos!A2:K201,"SELECT A,B,D,H,I,K WHERE C=\'SI\' ORDER BY H DESC LIMIT 10",0),"Sin jugadas aún")';
  sheet.getRange('A' + secRow).setFormula(queryFormula);

  // ── Formatting ──
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 160);
  sheet.setColumnWidth(6, 160);

  // Protect sheet as read-only
  var protection = sheet.protect().setDescription('Resumen - Solo lectura');
  protection.setWarningOnly(true);
}

// ═══════════════════════════════════════════════════════════
// UTILITY: Initialize test codes sheet
// Run manually. Set Script Properties: TEST_MODE = true
// ═══════════════════════════════════════════════════════════
function initTestCodes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('codigos_prueba');
  if (!sheet) {
    sheet = ss.insertSheet('codigos_prueba');
  } else {
    sheet.clear();
  }

  // Same columns as codigos
  sheet.getRange(1, 1, 1, 11).setValues([['numero', 'codigo', 'usado', 'premio', 'nombre', 'email', 'whatsapp', 'ganado_en', 'canjeado_en', 'cedula', 'datos_enviados']]);

  var pad3 = function(n) { return n < 10 ? '00' + n : n < 100 ? '0' + n : '' + n; };
  var rows = [];
  for (var i = 1; i <= 10; i++) {
    rows.push([pad3(900 + i), 'TEST-' + pad3(i + 0), '', '', '', '', '', '', '', '', '']);
  }
  // TEST-001 through TEST-010
  sheet.getRange(2, 1, rows.length, 11).setValues(rows);
}
