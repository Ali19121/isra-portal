// ═══════════════════════════════════════════════════════════
// ISRA SCHOOLS — APPS SCRIPT BACKEND (FINAL)
// Properly handles CORS for GitHub Pages
// All 4 operations: questions, settings, results, history
// ═══════════════════════════════════════════════════════════

// ── CORS HEADERS ─────────────────────────────────────────
function setCORS(output) {
  return output
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function jsonResponse(obj) {
  const out = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return out;
}

// ── MAIN GET HANDLER ─────────────────────────────────────
function doGet(e) {
  const p = e.parameter || {};

  // Write via GET (payload param) for CORS compatibility
  if (p.payload) {
    try {
      const data = JSON.parse(decodeURIComponent(p.payload));
      return jsonResponse(handleAction(data));
    } catch(err) {
      return jsonResponse({error: 'Bad payload: ' + err.message});
    }
  }

  // Read operations
  try {
    const action = p.action || '';
    if (action === 'ping')
      return jsonResponse({status:'ok', time: new Date().toISOString()});
    if (action === 'getSettings')
      return jsonResponse(getSettings());
    if (action === 'getQuestions')
      return jsonResponse(getQuestions(p.cat || ''));
    if (action === 'getHistory') {
      if (!checkSecret(p.secret))
        return jsonResponse({error:'Unauthorized'});
      return jsonResponse(getHistory());
    }
    return jsonResponse({error: 'Unknown action: ' + action});
  } catch(err) {
    return jsonResponse({error: err.message});
  }
}

// ── MAIN POST HANDLER ────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    return jsonResponse(handleAction(data));
  } catch(err) {
    return jsonResponse({error: err.message});
  }
}

// ── ACTION ROUTER ─────────────────────────────────────────
function handleAction(data) {
  const action = data.action || '';
  try {
    if (action === 'saveResult')
      return saveResult(data);
    if (action === 'saveQuestions') {
      if (!checkSecret(data.secret)) return {error:'Unauthorized'};
      return saveQuestions(data.cat, data.questions);
    }
    if (action === 'saveSettings') {
      if (!checkSecret(data.secret)) return {error:'Unauthorized'};
      return saveSettings(data.settings);
    }
    if (action === 'clearHistory') {
      if (!checkSecret(data.secret)) return {error:'Unauthorized'};
      return clearHistory();
    }
    return {error: 'Unknown action: ' + action};
  } catch(err) {
    return {error: err.message};
  }
}

// ── SHEET HELPERS ─────────────────────────────────────────
function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(name, headers) {
  let sheet = ss().getSheetByName(name);
  if (!sheet) {
    sheet = ss().insertSheet(name);
    if (headers && headers.length) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#0D2B5E')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// ── SETTINGS ─────────────────────────────────────────────
function getAllSettingsMap() {
  const sheet = getOrCreateSheet('Settings', ['Key','Value']);
  const rows = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) map[String(rows[i][0])] = String(rows[i][1] || '');
  }
  return map;
}

function checkSecret(secret) {
  const map = getAllSettingsMap();
  return secret && map['admin_secret'] && secret === map['admin_secret'];
}

function getSettings() {
  const s = getAllSettingsMap();
  const safe = Object.assign({}, s);
  delete safe.admin_secret; // never expose secret
  return safe;
}

function saveSettings(settings) {
  if (!settings) return {error: 'No settings provided'};
  const sheet = getOrCreateSheet('Settings', ['Key','Value']);
  const rows = sheet.getDataRange().getValues();
  // Build row index map
  const idx = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) idx[String(rows[i][0])] = i + 1;
  }
  Object.entries(settings).forEach(([k, v]) => {
    if (idx[k]) {
      sheet.getRange(idx[k], 2).setValue(v);
    } else {
      sheet.appendRow([k, v]);
    }
  });
  return {saved: true};
}

// ── QUESTIONS ─────────────────────────────────────────────
function getQuestions(cat) {
  const sheet = getOrCreateSheet('Questions',
    ['Category','Type','Question','Option A','Option B','Option C','Option D','Correct','Marks']);
  const rows = sheet.getDataRange().getValues();
  const questions = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] || !r[2]) continue;
    if (cat && String(r[0]) !== String(cat)) continue;
    if (String(r[1]) === 'mcq') {
      questions.push({
        type: 'mcq',
        text: String(r[2]),
        options: [String(r[3]||''), String(r[4]||''), String(r[5]||''), String(r[6]||'')],
        correct: parseInt(r[7]) || 0,
        marks: 1
      });
    } else {
      questions.push({
        type: 'desc',
        text: String(r[2]),
        marks: parseInt(r[7]) || 5
      });
    }
  }
  return {questions: questions, count: questions.length};
}

function saveQuestions(cat, questions) {
  if (!cat) return {error: 'Category required'};
  if (!Array.isArray(questions)) return {error: 'Questions must be an array'};

  const sheet = getOrCreateSheet('Questions',
    ['Category','Type','Question','Option A','Option B','Option C','Option D','Correct','Marks']);

  // Delete existing rows for this category (bottom up)
  const rows = sheet.getDataRange().getValues();
  const toDelete = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(cat)) toDelete.push(i + 1);
  }
  toDelete.forEach(r => sheet.deleteRow(r));

  // Add new questions
  questions.forEach(q => {
    if (q.type === 'mcq') {
      sheet.appendRow([
        cat, 'mcq', q.text || '',
        (q.options||[])[0]||'', (q.options||[])[1]||'',
        (q.options||[])[2]||'', (q.options||[])[3]||'',
        q.correct || 0, 1
      ]);
    } else {
      sheet.appendRow([cat, 'desc', q.text||'', '','','','', q.marks||5, q.marks||5]);
    }
  });

  return {saved: true, count: questions.length, cat: cat};
}

// ── RESULTS ───────────────────────────────────────────────
function saveResult(data) {
  const sheet = getOrCreateSheet('Results', [
    'Date','Time','Name','Father Name','CNIC','Test Category',
    'Score %','Correct','Wrong','Skipped','Total MCQ',
    'Time Taken','Reference ID','AI Analysis','Cheat Count'
  ]);
  sheet.appendRow([
    data.date||'', data.time||'',
    data.name||'', data.father||'', data.cnic||'',
    data.test||'', (data.score||0)+'%',
    data.correct||0, data.wrong||0, data.skipped||0,
    data.totalMcq||0, data.timeTaken||'',
    data.refId||'', data.aiAnalysis||'',
    data.cheatCount||0
  ]);
  return {saved: true};
}

// ── HISTORY ───────────────────────────────────────────────
function getHistory() {
  const sheet = getOrCreateSheet('Results', [
    'Date','Time','Name','Father Name','CNIC','Test Category',
    'Score %','Correct','Wrong','Skipped','Total MCQ',
    'Time Taken','Reference ID','AI Analysis','Cheat Count'
  ]);
  if (sheet.getLastRow() < 2) return {records:[], count:0};
  const rows = sheet.getDataRange().getValues();
  const records = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[2]) continue;
    records.push({
      date:      String(r[0]||''),
      time:      String(r[1]||''),
      name:      String(r[2]||''),
      father:    String(r[3]||''),
      cnic:      String(r[4]||''),
      test:      String(r[5]||''),
      score:     String(r[6]||'0%').replace('%',''),
      correct:   r[7]||0,
      wrong:     r[8]||0,
      skipped:   r[9]||0,
      totalMcq:  r[10]||0,
      timeTaken: String(r[11]||''),
      refId:     String(r[12]||''),
      cheatCount:r[14]||0
    });
  }
  return {records: records, count: records.length};
}

function clearHistory() {
  const sheet = getOrCreateSheet('Results');
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
  return {cleared: true};
}

// ── ONE-TIME SETUP ────────────────────────────────────────
function setupSheets() {
  getOrCreateSheet('Settings',   ['Key','Value']);
  getOrCreateSheet('Questions',  ['Category','Type','Question','Option A','Option B','Option C','Option D','Correct','Marks']);
  getOrCreateSheet('Results',    ['Date','Time','Name','Father Name','CNIC','Test Category','Score %','Correct','Wrong','Skipped','Total MCQ','Time Taken','Reference ID','AI Analysis','Cheat Count']);

  // Set defaults if not already set
  const map = getAllSettingsMap();
  const defaults = {
    duration:               '60',
    pw_sindhi_hyd:          'isra123',
    pw_math_hyd:            'isra123',
    pw_english_hyd:         'isra123',
    pw_math_mirpur:         'isra123',
    pw_sst_mirpur:          'isra123',
    pw_montessori_mirpur:   'isra123',
  };
  if (!map['admin_secret']) {
    defaults['admin_secret'] = 'ISRA-' + Math.random().toString(36).substring(2,10).toUpperCase();
  }
  Object.entries(defaults).forEach(([k,v]) => {
    if (!map[k]) {
      const sheet = ss().getSheetByName('Settings');
      sheet.appendRow([k, v]);
    }
  });

  const secret = getAllSettingsMap()['admin_secret'];
  SpreadsheetApp.getUi().alert(
    '✅ ISRA Portal Setup Complete!\n\n' +
    'Admin Secret Key:\n' + secret + '\n\n' +
    '→ Copy this key\n' +
    '→ Open portal → Admin Panel → Settings\n' +
    '→ Paste in "Admin Secret Key" field\n' +
    '→ Click Save URL & Key\n\n' +
    '3 sheets created: Settings, Questions, Results'
  );
}
