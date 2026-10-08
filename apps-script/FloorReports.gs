/**
 * FLOOR REPORTS -> GOOGLE SHEETS (monthly store + mismatch alert)
 *
 * What it does, once a day (or whenever you press Run):
 *  1. Looks in your Drive folder for the files exported by the Floor Report Reader:
 *       ACTUAL FLOOR REPORT -03 OCT 26 (U-1).xlsx
 *       FINAL PRODUCTION REPORT - 03 OCT (UNIT-1).xlsx
 *  2. Copies each new / changed ACTUAL file into one Google Sheet per month,
 *     "FLOOR REPORT OCT 2026", with tabs  U1 DATA, U2 DATA, U1 CATEGORY, U2 CATEGORY.
 *     A day that is loaded again replaces the old rows of that day.
 *  3. Checks every file and sends you ONE email listing anything that does not match:
 *       - the CATEGORY tab says WRONG (color + DPD + white/wash + sample is not the production)
 *       - the rows on Sheet1 do not add up to the CATEGORY production figure
 *       - the same batch number appears on more than one row
 *       - the day is already stored in Google with DIFFERENT figures (the local file changed)
 *       - the FINAL PRODUCTION REPORT file of that day has a different total than the ACTUAL file
 *
 * SET UP (about 5 minutes, once)
 *  1. In Google Drive create a folder, for example "Floor Reports", and save/sync your daily files there.
 *  2. Open https://script.google.com -> New project -> paste this whole file.
 *  3. Left side: Services (+) -> add "Drive API" (version v3).
 *  4. Edit CONFIG below (folder name, your email).
 *  5. Press Run on  syncFloorReports  once and allow the permissions.
 *  6. Press Run on  installDailyTrigger  to make it run by itself every evening.
 *
 * Note: this script was written without being able to run it against your Google account.
 * Try it first on two or three files and look at the monthly sheet and the email.
 */
const CONFIG = {
  FOLDER_NAME: 'Floor Reports',     // the Drive folder that holds the exported files
  EMAIL: '',                        // where to send the alert; empty = your own Google account
  KEY: 'change-this-secret-word',   // the page sends this word with its data; type the same word in the page. Anyone without it is refused
  TRIGGER_HOUR: 22,                 // daily run time (hour, 0-23, script time zone)
  RANGE: {1: [1, 23], 2: [101, 106]},
  TOLERANCE_KG: 0.5
};
const MON = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
const CAT_HEAD = ['DATE','PRODUCTION','COLOR','DPD','WHITE / WASH','REPROCESS','SAMPLE','CHECK','FILE','UPDATED'];
const DATA_HEAD = ['DATE','MACHINE','CAP','BUYER','ORDER','BATCH NO','COLOR','FABRICATION','QTY','LOAD TIME','UNLOAD TIME','DURATION','RFT','CATEGORY','SHADE'];

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncFloorReports').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncFloorReports').timeBased().everyDays(1).atHour(CONFIG.TRIGGER_HOUR).create();
}

function syncFloorReports() {
  const folder = getFolder_(), props = PropertiesService.getScriptProperties();
  const todo = [], it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const m = /^ACTUAL FLOOR REPORT\s*-\s*(\d{1,2})\s+([A-Za-z]{3})\s+(\d{2})\s*\(U-([12])\)\.xlsx$/i.exec(f.getName());
    if (!m) continue;
    const stamp = String(f.getLastUpdated().getTime());
    if (props.getProperty('f_' + f.getId()) === stamp) continue;          // already stored, unchanged
    todo.push({file: f, stamp, day: +m[1], mon: MON.indexOf(m[2].toUpperCase()), year: 2000 + (+m[3]), unit: +m[4]});
  }
  todo.sort((a, b) => (a.year - b.year) || (a.mon - b.mon) || (a.day - b.day) || (a.unit - b.unit));
  const problems = [], done = [];
  todo.forEach(t => {
    const label = pad_(t.day) + ' ' + MON[t.mon] + ' ' + t.year + ' UNIT ' + t.unit;
    try {
      const out = processFile_(folder, t);
      done.push(label);
      out.forEach(p => problems.push(label + ': ' + p));
      props.setProperty('f_' + t.file.getId(), t.stamp);
    } catch (e) {
      problems.push(label + ': could not be read (' + e.message + ')');
    }
  });
  if (!todo.length) return;
  const to = CONFIG.EMAIL || Session.getEffectiveUser().getEmail();
  const body = (problems.length
      ? 'These need your attention:\n\n- ' + problems.join('\n- ')
      : 'Everything matches.') + '\n\nStored in Google: ' + done.join(', ');
  MailApp.sendEmail(to, (problems.length ? 'Floor report MISMATCH (' + problems.length + ')' : 'Floor reports stored, all match'), body);
}

/* ---- one ACTUAL file ---- */
function processFile_(folder, t) {
  const probs = [], u = t.unit, rng = CONFIG.RANGE[u];
  const dateKey = pad_(t.day) + '.' + pad_(t.mon + 1) + '.' + t.year;
  const conv = convert_(folder, t.file);
  let rows, cat;
  try {
    const ss = SpreadsheetApp.openById(conv);
    const s1 = ss.getSheetByName('Sheet1'), sc = ss.getSheetByName('CATEGORY');
    if (!s1 || !sc) throw new Error('Sheet1 or CATEGORY tab not found');
    const last = s1.getLastRow();
    rows = last >= 4 ? s1.getRange(4, 1, last - 3, 15).getValues().filter(r => r.some(c => c !== '' && c !== null)) : [];
    cat = readCategoryDay_(sc, t);
  } finally { DriveApp.getFileById(conv).setTrashed(true); }

  if (!cat) { probs.push('this day is not on the CATEGORY tab'); }
  else {
    if (String(cat.check).toUpperCase() !== 'OK') probs.push('CATEGORY check says WRONG (color + DPD + white/wash + sample is ' + fmt_(cat.color + cat.dpd + cat.ww + cat.sample) + ' but production is ' + fmt_(cat.prod) + ')');
    const sum = sheetTotal_(rows, rng, u);
    if (Math.abs(sum - cat.prod) > CONFIG.TOLERANCE_KG) probs.push('rows on Sheet1 add up to ' + fmt_(sum) + ' but CATEGORY shows ' + fmt_(cat.prod));
  }
  const dup = duplicateBatches_(rows);
  if (dup.length) probs.push('same batch number on more than one row: ' + dup.join(', '));

  // monthly store
  const store = monthBook_(folder, t.year, t.mon);
  const dTab = tab_(store, 'U' + u + ' DATA', DATA_HEAD), cTab = tab_(store, 'U' + u + ' CATEGORY', CAT_HEAD);
  if (cat) {
    const old = findRow_(cTab, dateKey);
    if (old) {
      const o = cTab.getRange(old, 1, 1, 7).getValues()[0];
      const nw = [cat.prod, cat.color, cat.dpd, cat.ww, cat.reproc, cat.sample], names = ['production','color','DPD','white/wash','reprocess','sample'], diff = [];
      nw.forEach((v, i) => { if (Math.abs(Number(o[i + 1]) - v) > CONFIG.TOLERANCE_KG) diff.push(names[i] + ' ' + fmt_(o[i + 1]) + ' in Google, ' + fmt_(v) + ' in the file'); });
      if (diff.length) probs.push('DIFFERENT from what is already stored in Google: ' + diff.join('; '));
    }
    replaceRows_(cTab, dateKey, [[dateKey, cat.prod, cat.color, cat.dpd, cat.ww, cat.reproc, cat.sample, cat.check, t.file.getName(), new Date()]]);
  }
  replaceRows_(dTab, dateKey, rows.map(r => [dateKey].concat(r.slice(1))));
  dTab.getRange('J2:K').setNumberFormat('h:mm AM/PM');

  // compare with the FINAL PRODUCTION REPORT of the same day
  const fname = 'FINAL PRODUCTION REPORT - ' + pad_(t.day) + ' ' + MON[t.mon] + ' (UNIT-' + u + ').xlsx';
  const fi = folder.getFilesByName(fname);
  if (fi.hasNext() && cat) {
    const fid = convert_(folder, fi.next());
    try {
      const fs = SpreadsheetApp.openById(fid).getSheetByName('Sheet1');
      const fl = fs.getLastRow(), fr = fl >= 3 ? fs.getRange(3, 1, fl - 2, 9).getValues() : [];
      const ftot = fr.reduce((s, r) => s + (typeof r[8] === 'number' ? r[8] : 0), 0);
      if (Math.abs(ftot - cat.prod) > CONFIG.TOLERANCE_KG) probs.push('FINAL PRODUCTION REPORT total ' + fmt_(ftot) + ' is different from the ACTUAL file production ' + fmt_(cat.prod));
    } finally { DriveApp.getFileById(fid).setTrashed(true); }
  } else if (!fi.hasNext()) {
    probs.push('no "' + fname + '" in the folder to compare with');
  }
  return probs;
}

/* ---- helpers ---- */
function readCategoryDay_(sc, t) {
  const v = sc.getRange(6, 1, 31, 12).getValues();
  for (let i = 0; i < v.length; i++) {
    const d = v[i][0];
    if (d instanceof Date && d.getDate() === t.day && d.getMonth() === t.mon && d.getFullYear() === t.year) {
      const n = x => Number(x) || 0;
      return {prod: n(v[i][1]), color: n(v[i][2]), dpd: n(v[i][3]), ww: n(v[i][5]), reproc: n(v[i][7]), sample: n(v[i][9]), check: v[i][11]};
    }
  }
  return null;
}
function sheetTotal_(rows, rng, u) {
  return rows.reduce((s, r) => {
    const mc = r[1], q = r[8];
    if (typeof q !== 'number') return s;                                   // bracketed / blank quantity is not counted
    if (typeof mc === 'number' && mc >= rng[0] && mc <= rng[1]) return s + q;
    if (/^SAMPLE\s*-?\s*([12])$/i.test(String(mc).trim()) && +String(mc).replace(/\D/g, '') === u) return s + q;
    return s;
  }, 0);
}
function duplicateBatches_(rows) {
  const seen = {}, dup = [];
  rows.forEach(r => {
    if (/SAMPLE/i.test(String(r[1]))) return;
    const k = String(r[5]).replace(/[()\s]/g, '');
    if (!k) return;
    seen[k] = (seen[k] || 0) + 1;
    if (seen[k] === 2) dup.push(k);
  });
  return dup;
}
function convert_(folder, file) {
  const res = Drive.Files.create({name: 'tmp ' + file.getName(), mimeType: MimeType.GOOGLE_SHEETS, parents: [folder.getId()]}, file.getBlob());
  return res.id;
}
function getFolder_() {
  const it = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  if (!it.hasNext()) throw new Error('Drive folder "' + CONFIG.FOLDER_NAME + '" not found');
  return it.next();
}
function monthBook_(folder, year, mon) {
  const name = 'FLOOR REPORT ' + MON[mon] + ' ' + year, it = folder.getFilesByName(name);
  if (it.hasNext()) return SpreadsheetApp.openById(it.next().getId());
  const ss = SpreadsheetApp.create(name);
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  return ss;
}
function tab_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0].setName(name) : ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#DDEBF7');
    sh.setFrozenRows(1);
  }
  return sh;
}
function findRow_(sh, dateKey) {
  const v = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getDisplayValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === dateKey) return i + 1;
  return 0;
}
function replaceRows_(sh, dateKey, rows) {
  const last = sh.getLastRow();
  if (last > 1) {
    const v = sh.getRange(1, 1, last, 1).getDisplayValues();
    for (let i = v.length - 1; i >= 1; i--) if (v[i][0] === dateKey) sh.deleteRow(i + 1);
  }
  if (!rows.length) return;
  sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@');   // keep the date as text dd.mm.yyyy
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}
const pad_ = n => (n < 10 ? '0' : '') + n;
const fmt_ = x => Number(x).toLocaleString('en-US');


/* ============================================================
 * WEB APP: the Floor Report Reader page (GitHub version) sends its data here.
 * Deploy ->  New deployment -> type "Web app" -> Execute as: Me, Who has access: Anyone -> Deploy.
 * Copy the /exec link into the page ("Google link") together with the KEY above.
 * After you change this code, use Deploy -> Manage deployments -> pencil -> Version: New version.
 * ============================================================ */
function doGet() { return json_({ok: true, note: 'Floor report receiver is running. Send data from the page.'}); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const p = JSON.parse(e.postData.contents);
    if (!CONFIG.KEY || p.key !== CONFIG.KEY) return json_({error: 'wrong secret word (KEY)'});
    const folder = getFolder_(), problems = [], changed = [];
    (p.items || []).forEach(it => {
      const u = +it.unit, y = +it.iso.slice(0, 4), m = +it.iso.slice(5, 7) - 1, d = +it.iso.slice(8);
      const dateKey = pad_(d) + '.' + pad_(m + 1) + '.' + y, label = 'UNIT ' + u + ' ' + dateKey;
      const store = monthBook_(folder, y, m);
      const dTab = tab_(store, 'U' + u + ' DATA', DATA_HEAD), cTab = tab_(store, 'U' + u + ' CATEGORY', CAT_HEAD);
      const s = it.sum, old = findRow_(cTab, dateKey);
      if (old) {
        const o = cTab.getRange(old, 1, 1, 7).getValues()[0], names = ['production', 'color', 'DPD', 'white/wash', 'reprocess', 'sample'];
        [s.prod, s.color, s.dpd, s.ww, s.reproc, s.sample].forEach((v, i) => {
          if (Math.abs(Number(o[i + 1]) - v) > CONFIG.TOLERANCE_KG) changed.push(label + ': ' + names[i] + ' was ' + fmt_(o[i + 1]) + ' in Google, now ' + fmt_(v));
        });
      }
      if (!s.ok) problems.push(label + ': CATEGORY check is WRONG (color + DPD + white/wash + sample is not the production)');
      const dup = duplicateBatches_(it.rows.map(r => [null, r[0], null, null, null, r[4]]));
      if (dup.length) problems.push(label + ': same batch number on more than one row: ' + dup.join(', '));
      replaceRows_(cTab, dateKey, [[dateKey, s.prod, s.color, s.dpd, s.ww, s.reproc, s.sample, s.ok ? 'OK' : 'WRONG', 'web page', new Date()]]);
      replaceRows_(dTab, dateKey, it.rows.map(r => [dateKey].concat(r)));
    });
    if (changed.length && CONFIG.EMAIL) MailApp.sendEmail(CONFIG.EMAIL, 'Floor report: stored figures changed', changed.join('\n'));
    return json_({ok: true, saved: (p.items || []).length, problems, changed});
  } catch (err) {
    return json_({error: String(err && err.message || err)});
  } finally { try { lock.releaseLock(); } catch (x) {} }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
