// ═══════════════════════════════════════════════════════════
// ISRA SCHOOLS — APPS SCRIPT BACKEND v2
// Handles GET requests with optional 'payload' param for writes
// This avoids CORS preflight issues from GitHub Pages
// ═══════════════════════════════════════════════════════════

const SHEET_NAMES = {
  RESULTS:   'Results',
  QUESTIONS: 'Questions',
  SETTINGS:  'Settings',
};

// ── MAIN ENTRY POINTS ─────────────────────────────────────

function doGet(e) {
  // All requests come as GET — either pure read or write via 'payload' param
  const params = e.parameter || {};
  
  // If 'payload' exists → it's a write operation encoded as GET
  if (params.payload) {
    try {
      const data = JSON.parse(decodeURIComponent(params.payload));
      return handleWrite(data);
    } catch(err) {
      return jsonOk({error: 'Invalid payload: ' + err.toString()});
    }
  }
  
  // Pure read operations
  const action = params.action || '';
  try {
    if (action === 'ping')        return jsonOk({status:'ok', time:new Date().toISOString()});
    if (action === 'getSettings') return jsonOk(getSettings());
    if (action === 'getQuestions')return jsonOk(getQuestions(params.cat));
    if (action === 'getHistory')  {
      if (!verifySecret(params.secret)) return jsonOk({error:'Unauthorized'});
      return jsonOk(getHistory(params.cat));
    }
    return jsonOk({error:'Unknown action: ' + action});
  } catch(err) {
    return jsonOk({error: err.toString()});
  }
}

function doPost(e) {
  // Keep doPost working too (in case browser sends it)
  try {
    const data = JSON.parse(e.postData.contents);
    return handleWrite(data);
  } catch(err) {
    return jsonOk({error: err.toString()});
  }
}

function handleWrite(data) {
  const action = data.action || '';
  try {
    if (action === 'saveResult')    return jsonOk(saveResult(data));
    if (action === 'saveQuestions') {
      if (!verifySecret(data.secret)) return jsonOk({error:'Unauthorized — wrong secret key'});
      return jsonOk(saveQuestions(data.cat, data.questions));
    }
    if (action === 'saveSettings') {
      if (!verifySecret(data.secret)) return jsonOk({error:'Unauthorized — wrong secret key'});
      return jsonOk(saveSettings(data.settings));
    }
    if (action === 'clearHistory') {
      if (!verifySecret(data.secret)) return jsonOk({error:'Unauthorized — wrong secret key'});
      return jsonOk(clearHistory());
    }
    return jsonOk({error:'Unknown write action: ' + action});
  } catch(err) {
    return jsonOk({error: err.toString()});
  }
}

// ── HELPERS ───────────────────────────────────────────────

function jsonOk(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = {
      Results:   ['Date','Time','Name','Father Name','CNIC','Test','Score%',
                  'Correct','Wrong','Skipped','Total MCQ','Time Taken',
                  'Ref ID','AI Analysis','Cheats'],
      Questions: ['Category','Type','Question','Option A','Option B','Option C','Option D','Correct Index','Marks'],
      Settings:  ['Key','Value'],
    };
    if (headers[name]) {
      sheet.appendRow(headers[name]);
      sheet.getRange(1,1,1,headers[name].length)
        .setBackground('#0D2B5E').setFontColor('#FFFFFF').setFontWeight('bold');
    }
    // Seed default settings
    if (name === 'Settings') {
      const secret = 'ISRA-' + Math.random().toString(36).substring(2,10).toUpperCase();
      const defaults = [
        ['duration','60'],
        ['admin_secret', secret],
        ['pw_sindhi_hyd','isra123'],
        ['pw_math_hyd','isra123'],
        ['pw_english_hyd','isra123'],
        ['pw_math_mirpur','isra123'],
        ['pw_sst_mirpur','isra123'],
        ['pw_montessori_mirpur','isra123'],
      ];
      defaults.forEach(row => sheet.appendRow(row));
    }
  }
  return sheet;
}

function getAllSettingsObj() {
  const sheet = getSheet('Settings');
  const rows = sheet.getDataRange().getValues();
  const obj = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) obj[String(rows[i][0])] = String(rows[i][1]);
  }
  return obj;
}

function verifySecret(secret) {
  const settings = getAllSettingsObj();
  return secret && secret === settings['admin_secret'];
}

// ── SETTINGS ──────────────────────────────────────────────

function getSettings() {
  const s = getAllSettingsObj();
  // Don't expose admin_secret to public
  const safe = Object.assign({}, s);
  delete safe.admin_secret;
  return safe;
}

function saveSettings(settings) {
  const sheet = getSheet('Settings');
  const rows = sheet.getDataRange().getValues();
  const rowMap = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) rowMap[String(rows[i][0])] = i + 1; // 1-indexed row number
  }
  Object.entries(settings).forEach(([key, value]) => {
    if (rowMap[key] !== undefined) {
      sheet.getRange(rowMap[key], 2).setValue(value);
    } else {
      sheet.appendRow([key, value]);
    }
  });
  return {saved: true};
}

// ── QUESTIONS ─────────────────────────────────────────────

function getQuestions(cat) {
  const sheet = getSheet('Questions');
  const rows = sheet.getDataRange().getValues();
  const questions = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0] || !row[2]) continue;
    if (cat && String(row[0]) !== String(cat)) continue;
    if (String(row[1]) === 'mcq') {
      questions.push({
        type: 'mcq',
        text: String(row[2]),
        options: [String(row[3]||''), String(row[4]||''), String(row[5]||''), String(row[6]||'')],
        correct: parseInt(row[7]) || 0,
        marks: 1
      });
    } else {
      questions.push({
        type: 'desc',
        text: String(row[2]),
        marks: parseInt(row[7]) || 5
      });
    }
  }
  return {questions, count: questions.length};
}

function saveQuestions(cat, questions) {
  if (!cat || !questions) return {error: 'Missing cat or questions'};
  const sheet = getSheet('Questions');
  const rows = sheet.getDataRange().getValues();
  
  // Delete existing rows for this category (from bottom up)
  const toDelete = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(cat)) toDelete.push(i + 1);
  }
  toDelete.forEach(r => sheet.deleteRow(r));
  
  // Append new questions
  questions.forEach(q => {
    if (q.type === 'mcq') {
      sheet.appendRow([
        cat, 'mcq', q.text,
        (q.options||[])[0]||'', (q.options||[])[1]||'',
        (q.options||[])[2]||'', (q.options||[])[3]||'',
        q.correct||0, 1
      ]);
    } else {
      sheet.appendRow([cat, 'desc', q.text, '','','','', q.marks||5, q.marks||5]);
    }
  });
  return {saved: true, count: questions.length, cat};
}

// ── RESULTS ───────────────────────────────────────────────

function saveResult(data) {
  const sheet = getSheet('Results');
  sheet.appendRow([
    data.date||'', data.time||'', data.name||'', data.father||'',
    data.cnic||'', data.test||'',
    (data.score||'0')+'%', data.correct||0, data.wrong||0,
    data.skipped||0, data.totalMcq||0, data.timeTaken||'',
    data.refId||'', data.aiAnalysis||'', data.cheatCount||0
  ]);
  return {saved: true};
}

function getHistory(cat) {
  const sheet = getSheet('Results');
  if (sheet.getLastRow() < 2) return {records:[], count:0};
  const rows = sheet.getDataRange().getValues();
  const records = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[2]) continue; // no name
    if (cat && cat !== 'all' && row[5] && !String(row[5]).includes(cat)) continue;
    records.push({
      date:String(row[0]), time:String(row[1]),
      name:String(row[2]), father:String(row[3]),
      cnic:String(row[4]), test:String(row[5]),
      score:String(row[6]), correct:row[7], wrong:row[8],
      skipped:row[9], totalMcq:row[10],
      timeTaken:String(row[11]), refId:String(row[12]),
      cheatCount:row[14]||0
    });
  }
  return {records, count: records.length};
}

function clearHistory() {
  const sheet = getSheet('Results');
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
  return {cleared: true};
}

// ── ONE-TIME SETUP ─────────────────────────────────────────

function setupSheets() {
  getSheet('Results');
  getSheet('Questions');
  getSheet('Settings');
  
  const secret = getAllSettingsObj()['admin_secret'];
  SpreadsheetApp.getUi().alert(
    'ISRA Portal Setup Complete!\n\n' +
    'Your Admin Secret Key:\n\n' +
    secret + '\n\n' +
    'Copy this key and paste it in:\n' +
    'Portal → Admin Panel → Settings → Admin Secret Key\n\n' +
    '3 sheets created: Results, Questions, Settings'
  );
}
