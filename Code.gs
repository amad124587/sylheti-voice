/**
 * সিলেটি কণ্ঠ সংগ্রহ — Google Apps Script web app (backend)
 * রেকর্ডিং পেজ থাকে GitHub Pages-এ (মাইকের জন্য); এই স্ক্রিপ্ট শুধু বাক্য দেয় আর রেকর্ডিং Drive-এ রাখে:
 *   My Drive/my_voice/collection/speakers/<code>/sylheti|bangla_english/<file>
 * প্রতিটা রেকর্ডিং "voice_collection_log" শিটে লেখা হয় (একই ফোল্ডারে)।
 *
 * ফাইলের নাম = পেজে দেখানো বাক্য নম্বর:  বাক্য ১৫০১ → syl_1501,  বাক্য ১২ → me_012
 *
 * Deploy: Deploy → Manage deployments → ✏️ → Version: New version → Deploy (লিংক একই থাকে)।
 * পুরনো নামের ফাইল ঠিক করতে: উপরের ফাংশন তালিকা থেকে fixOldNames বেছে একবার Run চাপুন।
 */
const CFG = {
  CORPUS_ID: '1TycLVwUxg7hU0O68ykDny_dvZ7Hv04zR',      // My Drive/my_voice/corpus.csv (9,320 Sylheti sentences)
  LIST_ID:   '1esSLT0nmKUhvs75H1QuaoOVbOc4WVqdJKTE6S8z_XPg', // my_voice_recording_list (Bangla/English sentences)
  ROOT_ID:   '1WubrJi63PqYy3n4yrSr2Z1VSL3ESK9Tv',      // My Drive/my_voice
  MAX_BYTES: 15 * 1024 * 1024,
  MAX_SEC: 30
};
const EXTS = ['webm', 'm4a', 'ogg', 'mp3', 'wav', '3gp', 'amr', 'audio'];

function doGet() {
  return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">রেকর্ডিং পেজ: <a href="https://amad124587.github.io/sylheti-voice/">amad124587.github.io/sylheti-voice</a></p>')
    .setTitle('সিলেটি কণ্ঠ সংগ্রহ');
}

// ---------- sentences ----------
function getSentences(set) {
  if (set === 'A') {
    const v = SpreadsheetApp.openById(CFG.LIST_ID).getSheets()[0].getDataRange().getValues();
    const h = v[0].map(String);
    const iId = h.indexOf('id'), iS = h.indexOf('sentence'), iL = h.indexOf('language');
    return v.slice(1).filter(r => r[iS]).map(r => [Number(r[iId]), String(r[iS]).trim(),
      r[iL] === 'English' ? 'ইংরেজি বাক্য' : 'বাংলা বাক্য']);
  }
  let csv = DriveApp.getFileById(CFG.CORPUS_ID).getBlob().getDataAsString('UTF-8');
  if (csv.charCodeAt(0) === 0xFEFF) csv = csv.slice(1);
  const rows = Utilities.parseCsv(csv);
  const h = rows[0];
  const iR = h.indexOf('excel_row'), iS = h.indexOf('sylheti'), iB = h.indexOf('bangla');
  return rows.slice(1).filter(r => r[iS]).map(r => [Number(r[iR]), String(r[iS]).trim(), String(r[iB] || '').trim()]);
}

// ---------- storage helpers ----------
function folder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function collection_() { return folder_(DriveApp.getFolderById(CFG.ROOT_ID), 'collection'); }
function speakerDir_(code, set) {                       // cached folder id → no folder search on every save
  const props = PropertiesService.getScriptProperties(), key = 'dir_' + code + '_' + set;
  const id = props.getProperty(key);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const dir = folder_(folder_(folder_(collection_(), 'speakers'), code), set === 'B' ? 'sylheti' : 'bangla_english');
  props.setProperty(key, dir.getId());
  return dir;
}
function book_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('LOG_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const ss = SpreadsheetApp.create('voice_collection_log');
  DriveApp.getFileById(ss.getId()).moveTo(collection_());
  const sp = ss.getSheets()[0]; sp.setName('speakers');
  sp.appendRow(['code', 'gender', 'age', 'district', 'consent_text', 'consent_at', 'first_seen', 'last_seen']);
  const rc = ss.insertSheet('recordings');
  rc.appendRow(['time', 'code', 'set', 'index', 'excel_row', 'file_name', 'file_id', 'duration_sec', 'mime', 'size_bytes', 'sentence_no']);
  sp.setFrozenRows(1); rc.setFrozenRows(1);
  props.setProperty('LOG_ID', ss.getId());
  return ss;
}
function cleanCode_(c) {
  c = String(c || '').trim().toLowerCase().replace(/[^a-z0-9_-]/g, '');
  if (c.length < 3 || c.length > 24) throw new Error('code অন্তত ৩ অক্ষরের হোক (ইংরেজি অক্ষর বা সংখ্যা)।');
  return c;
}
const pad_ = (n, w) => String(n).padStart(w, '0');
// file name from the sentence number shown on the page (index is 0-based → +1)
function baseName_(set, index) { return set === 'B' ? 'syl_' + pad_(index + 1, 4) : 'me_' + pad_(index + 1, 3); }

// ---------- speakers ----------
function register(p) {
  const code = cleanCode_(p && p.code);
  if (!p.consent) throw new Error('এগোতে সম্মতির ঘরে টিক দিন।');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = book_().getSheetByName('speakers');
    const v = sh.getDataRange().getValues();
    const now = new Date();
    const clip = s => String(s || '').slice(0, 60);
    let r = -1;
    for (let i = 1; i < v.length; i++) if (v[i][0] === code) { r = i + 1; break; }
    if (r > 0) {
      sh.getRange(r, 2, 1, 3).setValues([[clip(p.gender), clip(p.age), clip(p.district)]]);
      sh.getRange(r, 8).setValue(now);
    } else {
      sh.appendRow([code, clip(p.gender), clip(p.age), clip(p.district), 'v1', now, now, now]);
    }
  } finally { lock.releaseLock(); }
  return progress(code);
}

function progress(code) {
  code = cleanCode_(code);
  const sh = book_().getSheetByName('recordings');
  const n = sh.getLastRow();
  const done = { A: {}, B: {} };
  if (n > 1) {
    const v = sh.getRange(2, 2, n - 1, 3).getValues();   // code, set, index
    v.forEach(r => { if (r[0] === code && done[r[1]]) done[r[1]][r[2]] = 1; });
  }
  return { code: code, A: Object.keys(done.A).map(Number), B: Object.keys(done.B).map(Number) };
}

// ---------- recordings ----------
function extFor_(mime) {
  mime = String(mime || '').toLowerCase();
  if (mime.indexOf('webm') >= 0) return 'webm';
  if (/mp4|m4a|aac|quicktime/.test(mime)) return 'm4a';
  if (mime.indexOf('ogg') >= 0) return 'ogg';
  if (/mpeg|mp3/.test(mime)) return 'mp3';
  if (mime.indexOf('wav') >= 0) return 'wav';
  if (mime.indexOf('3gp') >= 0) return '3gp';
  if (mime.indexOf('amr') >= 0) return 'amr';
  return 'audio';
}
function trashSameName_(dir, base, keepId) {           // one Drive query instead of one per extension
  const q = EXTS.map(e => "title = '" + base + '.' + e + "'").join(' or ');
  const it = dir.searchFiles('trashed = false and (' + q + ')');
  while (it.hasNext()) { const f = it.next(); if (f.getId() !== keepId) f.setTrashed(true); }
}

function saveRecording(o) {
  const code = cleanCode_(o.code);
  const set = o.set === 'A' ? 'A' : (o.set === 'B' ? 'B' : null);
  const index = Math.floor(Number(o.index)), row = Math.floor(Number(o.row));
  if (!set || !(index >= 0) || !(row > 0)) throw new Error('ভুল অনুরোধ।');
  const dur = o.duration == null ? '' : Math.round(Number(o.duration) * 100) / 100;
  if (dur !== '' && (dur < 0.5 || dur > CFG.MAX_SEC)) throw new Error('রেকর্ডিং ০.৫ থেকে ৩০ সেকেন্ডের মধ্যে হতে হবে।');
  const bytes = Utilities.base64Decode(String(o.data || ''));
  if (!bytes.length || bytes.length > CFG.MAX_BYTES) throw new Error('ফাইলটা খালি বা খুব বড়।');
  const mime = String(o.mime || 'application/octet-stream').split(';')[0];
  const base = baseName_(set, index);
  const name = base + '.' + extFor_(o.mime);

  const dir = speakerDir_(code, set);
  const file = dir.createFile(Utilities.newBlob(bytes, mime, name));
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    trashSameName_(dir, base, file.getId());             // re-recording replaces the old take
    book_().getSheetByName('recordings').appendRow([new Date(), code, set, index, row, name, file.getId(), dur, mime, bytes.length, index + 1]);
  } finally { lock.releaseLock(); }
  return { ok: true, name: name };
}

// ---------- একবার চালানোর জন্য: পুরনো নাম (excel row) → নতুন নাম (বাক্য নম্বর) ----------
function fixOldNames() {
  const sh = book_().getSheetByName('recordings');
  const n = sh.getLastRow();
  if (n < 2) { Logger.log('কোনো রেকর্ডিং নেই।'); return; }
  const width = Math.max(sh.getLastColumn(), 11);
  if (sh.getRange(1, 11).getValue() !== 'sentence_no') sh.getRange(1, 11).setValue('sentence_no');
  const v = sh.getRange(2, 1, n - 1, width).getValues();
  const rows = v.map((r, i) => ({ i: i, code: r[1], set: r[2], index: Number(r[3]), name: String(r[5]), id: String(r[6]) }))
                .sort((a, b) => a.index - b.index);              // ascending so a freed name is reused safely
  let renamed = 0, skipped = 0;
  rows.forEach(r => {
    const want = baseName_(r.set, r.index) + '.' + r.name.split('.').pop();
    v[r.i][10] = r.index + 1;
    if (r.name === want) return;
    let f; try { f = DriveApp.getFileById(r.id); } catch (e) { return; }
    if (f.isTrashed()) return;
    const dir = f.getParents().hasNext() ? f.getParents().next() : null;
    if (dir) {
      const clash = dir.getFilesByName(want);
      while (clash.hasNext()) { const c = clash.next(); if (c.getId() !== r.id && !c.isTrashed()) { skipped++; Logger.log('বাদ: ' + want + ' আগে থেকেই আছে (' + r.code + ')'); return; } }
    }
    f.setName(want); v[r.i][5] = want; renamed++;
  });
  sh.getRange(2, 1, n - 1, width).setValues(v);
  Logger.log('নাম ঠিক করা হলো: ' + renamed + ' টা, বাদ: ' + skipped + ' টা');
}

// ---------- API for the recorder page on GitHub Pages ----------
function doPost(e) {
  const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, error: 'bad request' }); }
  try {
    let r;
    if (req.action === 'sentences') {
      const all = getSentences(req.set);
      if (req.limit) { const o = Number(req.offset) || 0; r = { total: all.length, items: all.slice(o, o + Number(req.limit)) }; }
      else r = all;
    }
    else if (req.action === 'register') r = register(req);
    else if (req.action === 'progress') r = progress(req.code);
    else if (req.action === 'save') r = saveRecording(req);
    else throw new Error('unknown action');
    return out({ ok: true, result: r });
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  }
}
