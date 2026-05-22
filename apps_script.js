// ═══════════════════════════════════════════════════════════ // ISRA SCHOOLS — APPS SCRIPT BACKEND v3 (FIXED) // ═══════════════════════════════════════════════════════════
const SHEET_NAMES = { RESULTS: ‘Results’, QUESTIONS: ‘Questions’, SETTINGS: ‘Settings’, };
function doGet(e) { const params = e.parameter || {}; const output = ContentService.createTextOutput();
if (params.payload) { try { const data = JSON.parse(decodeURIComponent(params.payload)); const result = handleWrite(data); output.setContent(JSON.stringify(result)); } catch(err) { output.setContent(JSON.stringify({error: ‘Invalid payload:’ + err.toString()})); } } else { const action = params.action || ’‘; try { let result; if (action === ’ping’) result = {status:‘ok’, time:new Date().toISOString()}; else if (action === ‘getSettings’) result = getSettings(); else if (action === ‘getQuestions’) result = getQuestions(params.cat); else if (action === ‘getHistory’) { if (!verifySecret(params.secret)) result = {error:‘Unauthorized’}; else result = getHistory(params.cat); } else result = {error:‘Unknown action:’ + action}; output.setContent(JSON.stringify(result)); } catch(err) { output.setContent(JSON.stringify({error: err.toString()})); } }
output.setMimeType(ContentService.MimeType.JSON); return output; }
function doPost(e) { try { const data = JSON.parse(e.postData.contents); const result = handleWrite(data); return jsonOk(result); } catch(err) { return jsonOk({error: err.toString()}); } }
function handleWrite(data) { const action = data.action || ’‘; try { if (action === ’saveResult’) return saveResult(data); if (action === ‘saveQuestions’) { if (!verifySecret(data.secret)) return {error:‘Unauthorized’}; return saveQuestions(data.cat, data.questions); } if (action === ‘saveSettings’) { if (!verifySecret(data.secret)) return {error:‘Unauthorized’}; return saveSettings(data.settings); } if (action === ‘clearHistory’) { if (!verifySecret(data.secret)) return {error:‘Unauthorized’}; return clearHistory(); } if (action === ‘batchSaveQuestions’) { if (!verifySecret(data.secret)) return {error:‘Unauthorized’}; return batchSaveQuestions(data.cat, data.questions, data.batchIndex, data.totalBatches); } return {error:‘Unknown action:’ + action}; } catch(err) { return {error: err.toString()}; } }
function jsonOk(obj) { const output = ContentService.createTextOutput(JSON.stringify(obj)); output.setMimeType(ContentService.MimeType.JSON); return output; }
function getSheet(name) { const ss = SpreadsheetApp.getActiveSpreadsheet(); let sheet = ss.getSheetByName(name); if (!sheet) { sheet = ss.insertSheet(name); const headers = { Results: [‘Date’,‘Time’,‘Name’,‘Father Name’,‘CNIC’,‘Test’,‘Score%’, ‘Correct’,‘Wrong’,‘Skipped’,‘Total MCQ’,‘Time Taken’, ‘Ref ID’,‘AI Analysis’,‘Cheats’], Questions: [‘Category’,‘Type’,‘Question’,‘Option A’,‘Option B’,‘Option C’,‘Option D’,‘Correct Index’,‘Marks’], Settings: [‘Key’,‘Value’], }; if (headers[name]) { sheet.appendRow(headers[name]); sheet.getRange(1,1,1,headers[name].length) .setBackground(‘#0D2B5E’).setFontColor(‘#FFFFFF’).setFontWeight(‘bold’); } if (name === ‘Settings’) { const secret = ‘ISRA-’ + Math.random().toString(36).substring(2,10).toUpperCase(); const defaults = [ [‘duration’,‘60’], [‘admin_secret’, secret], [‘pw_sindhi_hyd’,‘isra123’], [‘pw_math_hyd’,‘isra123’], [‘pw_english_hyd’,‘isra123’], [‘pw_math_mirpur’,‘isra123’], [‘pw_sst_mirpur’,‘isra123’], [‘pw_montessori_mirpur’,‘isra123’], ]; defaults.forEach(row => sheet.appendRow(row)); } } return sheet; }
function getAllSettingsObj() { const sheet = getSheet(‘Settings’); const rows = sheet.getDataRange().getValues(); const obj = {}; for (let i = 1; i < rows.length; i++) { if (rows[i][0]) obj[String(rows[i][0])] = String(rows[i][1]); } return obj; }
function verifySecret(secret) { const settings = getAllSettingsObj(); return secret && secret === settings[‘admin_secret’]; }
function getSettings() { const s = getAllSettingsObj(); const safe = Object.assign({}, s); delete safe.admin_secret; return safe; }
function saveSettings(settings) { const sheet = getSheet(‘Settings’); const rows = sheet.getDataRange().getValues(); const rowMap = {}; for (let i = 1; i < rows.length; i++) { if (rows[i][0]) rowMap[String(rows[i][0])] = i + 1; } Object.entries(settings).forEach(([key, value]) => { if (rowMap[key] !== undefined) { sheet.getRange(rowMap[key], 2).setValue(value); } else { sheet.appendRow([key, value]); } }); return {saved: true}; }
function getQuestions(cat) { const sheet = getSheet(‘Questions’); const rows = sheet.getDataRange().getValues(); const questions = []; for (let i = 1; i < rows.length; i++) { const row = rows[i]; if (!row[0] || !row[2]) continue; if (cat && String(row[0]) !== String(cat)) continue; if (String(row[1]) === ‘mcq’) { questions.push({ type: ‘mcq’, text: String(row[2]), options: [String(row[3]||’‘), String(row[4]||’‘), String(row[5]||’‘), String(row[6]||’’)], correct: parseInt(row[7]) || 0, marks: 1 }); } else { questions.push({ type: ‘desc’, text: String(row[2]), marks: parseInt(row[7]) || 5 }); } } return {questions, count: questions.length}; }
function saveQuestions(cat, questions) { if (!cat || !questions) return {error: ‘Missing cat or questions’}; const sheet = getSheet(‘Questions’); const rows = sheet.getDataRange().getValues(); const toDelete = []; for (let i = rows.length - 1; i >= 1; i–) { if (String(rows[i][0]) === String(cat)) toDelete.push(i + 1); } toDelete.forEach(r => sheet.deleteRow(r)); questions.forEach(q => { if (q.type === ‘mcq’) { sheet.appendRow([cat, ‘mcq’, q.text, (q.options||[])[0]||’‘, (q.options||[])[1]||’‘, (q.options||[])[2]||’‘, (q.options||[])[3]||’’, q.correct||0, 1]); } else { sheet.appendRow([cat, ‘desc’, q.text, ’‘,’‘,’‘,’’, q.marks||5, q.marks||5]); } }); return {saved: true, count: questions.length, cat}; }
// NEW: Batch save for large imports function batchSaveQuestions(cat, questions, batchIndex, totalBatches) { if (!cat || !questions) return {error: ‘Missing cat or questions’}; const sheet = getSheet(‘Questions’);
// Only delete existing on first batch if (batchIndex === 0) { const rows = sheet.getDataRange().getValues(); const toDelete = []; for (let i = rows.length - 1; i >= 1; i–) { if (String(rows[i][0]) === String(cat)) toDelete.push(i + 1); } toDelete.forEach(r => sheet.deleteRow(r)); }
questions.forEach(q => { if (q.type === ‘mcq’) { sheet.appendRow([cat, ‘mcq’, q.text, (q.options||[])[0]||’‘, (q.options||[])[1]||’‘, (q.options||[])[2]||’‘, (q.options||[])[3]||’’, q.correct||0, 1]); } else { sheet.appendRow([cat, ‘desc’, q.text, ’‘,’‘,’‘,’’, q.marks||5, q.marks||5]); } });
return {saved: true, count: questions.length, cat, batch: batchIndex + 1, total: totalBatches}; }
function saveResult(data) { const sheet = getSheet(‘Results’); sheet.appendRow([ data.date||’‘, data.time||’‘, data.name||’‘, data.father||’‘, data.cnic||’‘, data.test||’‘, (data.score||’0’)+‘%’, data.correct||0, data.wrong||0, data.skipped||0, data.totalMcq||0, data.timeTaken||’‘, data.refId||’‘, data.aiAnalysis||’’, data.cheatCount||0 ]); return {saved: true}; }
function getHistory(cat) { const sheet = getSheet(‘Results’); if (sheet.getLastRow() < 2) return {records:[], count:0}; const rows = sheet.getDataRange().getValues(); const records = []; for (let i = 1; i < rows.length; i++) { const row = rows[i]; if (!row[2]) continue; if (cat && cat !== ‘all’ && row[5] && !String(row[5]).includes(cat)) continue; records.push({ date:String(row[0]), time:String(row[1]), name:String(row[2]), father:String(row[3]), cnic:String(row[4]), test:String(row[5]), score:String(row[6]), correct:row[7], wrong:row[8], skipped:row[9], totalMcq:row[10], timeTaken:String(row[11]), refId:String(row[12]), cheatCount:row[14]||0 }); } return {records, count: records.length}; }
function clearHistory() { const sheet = getSheet(‘Results’); if (sheet.getLastRow() > 1) { sheet.deleteRows(2, sheet.getLastRow() - 1); } return {cleared: true}; }
function setupSheets() { getSheet(‘Results’); getSheet(‘Questions’); getSheet(‘Settings’); const secret = getAllSettingsObj()[‘admin_secret’]; SpreadsheetApp.getUi().alert( ‘ISRA Portal Setup Complete!Secret Key:’ + secret + ‘this to: Portal → Admin → Settings’ ); }
═══════════════════════════════════════════════════════════════════════════════ FRONTEND CHANGES (index.html) ═══════════════════════════════════════════════════════════════════════════════
CHANGE #1: Hard-code URL (Line ~95) ──────────────────────────────────── REPLACE: const gurl = ()=>localStorage.getItem(UK)||’’;
WITH: // ═══════════════════════════════════════════════════ // HARD-CODED URL — Works on ALL computers automatically // ═══════════════════════════════════════════════════ const APPS_SCRIPT_URL = ‘https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec’; const gurl = () => APPS_SCRIPT_URL;
CHANGE #2: Add batch save function (After saveQuestionsRemote) ─────────────────────────────────────────────────────────────── ADD: const BATCH_SIZE = 50;
async function saveQuestionsBatch(cat, questions) { if (!gurl()) { toast(‘URL not configured’, ‘er’); return false; } if (!getASK()) { toast(‘Secret key not configured’, ‘er’); return false; }
const totalBatches = Math.ceil(questions.length / BATCH_SIZE);
let totalSaved = 0;
showLoad(`Uploading 0/${questions.length} questions...`);

for (let i = 0; i < totalBatches; i++) {
  const batch = questions.slice(i * BATCH_SIZE, (i + 1) * BATCH_SIZE);
  let retries = 3, success = false;

  while (retries > 0 && !success) {
    try {
      const res = await apiPost({
        action: 'batchSaveQuestions', secret: getASK(), cat,
        questions: batch, batchIndex: i, totalBatches
      });
      if (res && res.saved) {
        totalSaved += batch.length;
        document.getElementById('loadMsg').textContent = 
          `Uploading ${totalSaved}/${questions.length}...`;
        success = true;
      } else throw new Error(res?.error || 'Unknown');
    } catch (e) {
      retries--;
      if (retries === 0) { hideLoad(); toast(`Batch ${i+1} failed: ${e.message}`, 'er'); return false; }
      await new Promise(r => setTimeout(r, 2000));
    }
  }
}
hideLoad();
toast(`${totalSaved} questions saved ✓`, 'ok');
_qCache[cat] = questions;
broadcastUpdate('QUESTIONS_UPDATED', { cat });
return true;
}
CHANGE #3: Update confirmImport to use batch ───────────────────────────────────────────── REPLACE confirmImport() with:
async function confirmImport() { const cat = document.getElementById(‘imp_cat’).value; if (!parsedImport || (!parsedImport.questions && !parsedImport.map)) { toast(‘No data to import’, ‘er’); return; } if (!gurl()) { toast(‘URL not set’, ‘er’); return; } if (!getASK()) { toast(‘Secret key not set’, ‘er’); return; }
if (parsedImport.type === 'answers') {
  const map = parsedImport.map;
  let updated = 0;
  if (!_qCache[cat]) _qCache[cat] = [];
  _qCache[cat].forEach((q, i) => {
    if (q.type === 'mcq') {
      const key = map[i + 1];
      if (key !== undefined) {
        const idx = key.charCodeAt(0) - 65;
        if (idx >= 0 && idx <= 3) { q.correct = idx; updated++; }
      }
    }
  });
  const ok = await saveQuestionsBatch(cat, _qCache[cat]);
  if (ok) toast('Updated ' + updated + ' answers ✓', 'ok');
} else {
  const newQuestions = parsedImport.questions;
  if (!_qCache[cat]) _qCache[cat] = [];
  _qCache[cat] = _qCache[cat].concat(newQuestions);
  const ok = await saveQuestionsBatch(cat, _qCache[cat]);
  if (ok) toast('Imported ' + newQuestions.length + ' questions ✓', 'ok');
}

document.getElementById('wordImportPanel').style.display = 'none';
document.getElementById('wordPreview').style.display = 'none';
document.getElementById('wordFileName').textContent = 'No file selected';
parsedImport = [];
await renderCatFilters();
await renderQList();
renderDash();
}
CHANGE #4: Add auto-polling (In DOMContentLoaded) ────────────────────────────────────────────────── ADD at end of DOMContentLoaded:
// Auto-poll for near real-time sync let pollTimer = null; function startPolling() { if (pollTimer) clearInterval(pollTimer); pollTimer = setInterval(async () => { if (document.getElementById(‘pgAdmin’).style.display !== ‘block’) return; try { const res = await apiGet({action: ‘ping’}); if (res && res.status === ‘ok’) { // Refresh current tab const activeTab = document.querySelector(‘.tab.on’); if (activeTab) { const tabId = activeTab.id.replace(‘tab_’, ’‘); if (tabId === ’qbank’) { await renderCatFilters(); await renderQList(); } else if (tabId === ‘dash’) await renderDash(); else if (tabId === ‘hist’) await renderHist(); } } } catch (e) { console.log(‘Poll error:’, e); } }, 15000); } startPolling();
CHANGE #5: Improve sendSheet with retry ──────────────────────────────────────── REPLACE sendSheet() with:
async function sendSheet(data) { const url = gurl(); const bar = document.getElementById(‘r_savebar’);
if (!url) {
  if (bar) { bar.className = 'savebar sb-er'; bar.textContent = '⚠ URL not configured'; }
  saveResultLocally(data);
  return;
}

if (bar) { bar.className = 'savebar sb-w'; bar.textContent = '⏳ Saving to Sheet...'; }

let retries = 3;
while (retries > 0) {
  try {
    const res = await apiPost({
      action: 'saveResult', date: data.date, time: data.time,
      name: data.name, father: data.father, cnic: data.cnic,
      test: data.test, score: data.score, correct: data.correct,
      wrong: data.wrong, skipped: data.skipped, totalMcq: data.totalMcq,
      timeTaken: data.timeTaken, refId: data.refId || '',
      aiAnalysis: data.aiAnalysis || '', cheatCount: data.cheatCount || 0
    });
    if (bar) {
      bar.className = 'savebar sb-ok';
      bar.textContent = '✅ Saved to Google Sheet!';
    }
    return;
  } catch (e) {
    retries--;
    if (retries === 0) {
      if (bar) { bar.className = 'savebar sb-er'; bar.textContent = '❌ Saved locally (sync later)'; }
      saveResultLocally(data);
    } else await new Promise(r => setTimeout(r, 2000));
  }
}
}
function saveResultLocally(data) { const key = ‘isra_pending_results’; const existing = JSON.parse(localStorage.getItem(key) || ‘[]’); existing.push({…data, pending: true, savedAt: new Date().toISOString()}); localStorage.setItem(key, JSON.stringify(existing)); }
