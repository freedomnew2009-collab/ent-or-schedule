// ══ Google Apps Script (Code.gs) – ENT OR Schedule v5 ══════════
// อัพเดทจาก v4 (Fixed): เพิ่ม saveLeaves / saveSwaps / saveConfig + ส่ง config กลับใน getAll
// v5 navy: เพิ่มคอลัมน์ Pre-Op (preopDate, preopStatus, lab, cxr, ekg, npo, preopNote, history)
// v5.2: เพิ่มตารางแพทย์ออกตรวจ OPD (แท็บ ENT_OPD เก็บแบบ key/value ทีละวัน)
// v5.3: บันทึกประวัติการแก้ไข (แท็บ ENT_Log) — Email ไหนแก้อะไร เมื่อไหร่ ทุกส่วนของแอพ
// Sheet tabs: ENT_Schedule, ENT_Doctors, ENT_Ops, ENT_Leaves, ENT_Swaps, ENT_Config, ENT_OPD, ENT_Log
const SHEET_ID = '1KWH-9JobfctIp-prqGSlZIJ4xOuP4Yb6QVA0spTE8sk';

// ★ v5: เสิร์ฟหน้าเว็บล่าสุดจาก GitHub (branch main) โดยตรง
// → อัพเดทโค้ดที่ GitHub ที่เดียว ผู้ใช้ /exec ทุกคนได้เวอร์ชันใหม่อัตโนมัติ
//   ไม่ต้อง copy index.html มาวางใน Apps Script และไม่ต้อง Deploy ใหม่อีก
// หมายเหตุ: ถ้าเปลี่ยน repo เป็น private ภายหลัง ลิงก์ raw นี้จะใช้ไม่ได้ ต้องกลับไปใช้ไฟล์ index ที่ฝังไว้
const LIVE_HTML_URL = 'https://raw.githubusercontent.com/freedomnew2009-collab/ent-or-schedule/main/index.html';
// เลขนี้เปลี่ยนทุกครั้งที่แก้ Code.gs — ใช้เช็คว่า deployment รับโค้ดใหม่แล้วหรือยัง
const APP_VERSION = 'v5.3';

const APT_COLS   = ['id','hn','name','date','ts','te','op','doctorName','di','di2','doctor2Name','anesthesia','tel1','tel2','status','note','preopDate','preopStatus','lab','cxr','ekg','npo','preopNote','history'];
const LEAVE_COLS = ['id','di','start','end','reason','status'];
const SWAP_COLS  = ['id','di','date','type','note'];
const DOC_COLS   = ['di','name','color','sched','orDays'];
const OP_COLS    = ['name'];
const CFG_COLS   = ['key','value'];   // ★ v5: เก็บค่าตั้งระบบ เช่น วันผ่าตัดของแผนก (orDays)
// ★ v5.2: ตาราง OPD — key เช่น doctors, settings, template, d2026-10-01 (1 แถวต่อ 1 วัน) / value เป็น JSON
const OPD_COLS   = ['key','value'];
// ★ v5.3: ประวัติการแก้ไข — 1 แถวต่อ 1 รายการที่เปลี่ยน (เพิ่ม/แก้/ลบ)
const LOG_COLS   = ['time','email','section','action','item','detail'];
const LOG_MAX_PER_SAVE = 200;   // กันบันทึกครั้งเดียวเขียน log ยาวเกินไป (เช่น อัพโหลดทั้งหมด)

function doGet(e) {
  const p = e.parameter || {};
  // เปิด URL?action=ping ใน browser เพื่อเช็คว่า deployment เป็นเวอร์ชันล่าสุดหรือยัง
  // ping = ตรวจสุขภาพระบบ + บอกว่า deployment ที่ใช้อยู่มีอะไรบ้าง
  // lock:true คือมีระบบกันเขียนชนกันแล้ว (ถ้าไม่ขึ้นบรรทัดนี้เลย = deployment ยังเป็นโค้ดเก่า)
  if (p.action === 'ping') return ok({
    status: 'online',
    version: APP_VERSION,
    serving: 'github-main',
    lock: (typeof handlePost === 'function'),
    opd: (typeof upsertKV === 'function'),
    log: (typeof appendLog === 'function')
  });

  // ประวัติการแก้ไขล่าสุด (ใหม่สุดก่อน) — n = จำนวนแถว (ไม่เกิน 500)
  if (p.action === 'getLog') {
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      return ok({ status: 'ok', log: readLog(ss, Math.min(Number(p.n) || 100, 500)) });
    } catch (err) {
      return ok({ status: 'error', message: err.message });
    }
  }

  if (p.action === 'getAll') {
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      // ★ ถ้าส่ง only=apts มา ดึงแค่ appointments sheet เดียว เร็วกว่ามาก
      if (p.only === 'apts') {
        return ok({
          appointments: readSheet(ss, p.sheet || 'ENT_Schedule', APT_COLS),
        });
      }
      // ★ v5: อ่านค่าตั้งระบบ (ENT_Config) กลับไปด้วย — ถ้ายังไม่มีแท็บนี้จะได้ config ว่าง
      const config = {};
      readSheet(ss, 'ENT_Config', CFG_COLS).forEach(r => { config[r.key] = r.value; });
      const opd = {};
      readSheet(ss, 'ENT_OPD', OPD_COLS).forEach(r => { if (r.key) opd[String(r.key)] = String(r.value); });
      return ok({
        appointments: readSheet(ss, p.sheet || 'ENT_Schedule', APT_COLS),
        leaves:       readSheet(ss, 'ENT_Leaves',  LEAVE_COLS),
        swaps:        readSheet(ss, 'ENT_Swaps',   SWAP_COLS),
        doctors:      readSheet(ss, 'ENT_Doctors', DOC_COLS),
        operations:   readSheet(ss, 'ENT_Ops',     OP_COLS),
        config:       config,
        opd:          opd,
      });
    } catch (err) {
      return ok({ status: 'error', message: err.message });
    }
  }

  // ★ หน้าเว็บ: ดึงเวอร์ชันล่าสุดจาก GitHub main มาเสิร์ฟเสมอ
  try {
    const res = UrlFetchApp.fetch(LIVE_HTML_URL, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      return HtmlService.createHtmlOutput(res.getContentText())
        .setTitle('ENT OR Schedule')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (err) { /* GitHub ล่ม/เข้าไม่ได้ → ใช้ตัวสำรองด้านล่าง */ }

  // ตัวสำรอง: ไฟล์ index ที่ฝังในโปรเจกต์ (ใช้เมื่อดึงจาก GitHub ไม่สำเร็จเท่านั้น)
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ENT OR Schedule')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ★ v5: ล็อกก่อนเขียนทุกครั้ง — ถ้าเจ้าหน้าที่ 2 คนกดบันทึกพร้อมกัน
// การเขียนจะเรียงคิวทีละคน ไม่อ่าน-เขียนซ้อนกันจนข้อมูลหาย
function doPost(e){
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(28000);
  } catch (err) {
    return ok({ status: 'busy', message: 'ระบบกำลังบันทึกรายการอื่นอยู่ กรุณาลองใหม่' });
  }
  try {
    return handlePost(e);
  } finally {
    lock.releaseLock();
  }
}

function handlePost(e){
  const body = JSON.parse(e.postData.contents || '{}');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  // ★ v5.3: ทุกการบันทึกจดว่าใครแก้ — หน้าเว็บส่ง email ที่ผู้ใช้กรอกไว้มาใน body.editor
  //   เทียบข้อมูลเก่า/ใหม่ก่อนเขียน แล้วจดเฉพาะรายการที่เปลี่ยนจริงลงแท็บ ENT_Log
  const log = [];
  const res = savePost(ss, body, log);
  try { appendLog(ss, body.editor, log); } catch (err) { console.log('log error: ' + err.message); }
  return ok(res);
}

function savePost(ss, body, log){
  if(body.action === 'upsert'){
    const old = upsert(ss, body.sheet, APT_COLS, body.row);
    safeDiff(log, old ? [old] : [], [body.row], aptSpec(), true);
    return {status: 'upserted', id: body.row.id};
  }

  if(body.action === 'pushAll' || body.action === 'saveSettings'){
    const names = docNames(ss, body.doctors);
    if(body.appointments) writeLogged(ss, body.sheet || 'ENT_Schedule', APT_COLS, body.appointments, aptSpec(), log);
    if(body.leaves)       writeLogged(ss, 'ENT_Leaves', LEAVE_COLS, body.leaves, leaveSpec(names), log);
    if(body.swaps)        writeLogged(ss, 'ENT_Swaps',  SWAP_COLS,  body.swaps,  swapSpec(names),  log);
    if(body.doctors)      writeLogged(ss, 'ENT_Doctors', DOC_COLS,  body.doctors, docSpec(), log);
    if(body.operations)   writeLogged(ss, 'ENT_Ops',     OP_COLS,   body.operations, opSpec(), log);
    return {status: 'success'};
  }

  // ★ v5: บันทึกวันลา / Override / ค่าตั้งระบบ แยกส่วน (หน้าเว็บซิงค์อัตโนมัติทุกครั้งที่แก้ไข)
  if (body.action === 'saveLeaves') {
    writeLogged(ss, 'ENT_Leaves', LEAVE_COLS, body.leaves || [], leaveSpec(docNames(ss)), log);
    return { status: 'saved' };
  }
  if (body.action === 'saveSwaps') {
    writeLogged(ss, 'ENT_Swaps', SWAP_COLS, body.swaps || [], swapSpec(docNames(ss)), log);
    return { status: 'saved' };
  }
  if (body.action === 'saveConfig') {
    const rows = Object.keys(body.config || {}).map(k => ({ key: k, value: body.config[k] }));
    const old = readRaw(ss, 'ENT_Config');
    write(ss, 'ENT_Config', CFG_COLS, rows);
    try { log.push.apply(log, configDiff(old, rows)); } catch (err) { console.log('diff error: ' + err.message); }
    return { status: 'saved' };
  }

  // ★ v5.2: ตาราง OPD — บันทึกเฉพาะ key ที่แก้ (ทีละวัน) ไม่เขียนทับทั้งแท็บ
  // สองคนแก้คนละวันพร้อมกันจึงไม่ทับกัน
  if (body.action === 'saveOpd') {
    const old = {};
    readRaw(ss, 'ENT_OPD').forEach(r => { if (r.key) old[String(r.key)] = String(r.value); });
    upsertKV(ss, 'ENT_OPD', body.items || []);
    try { log.push.apply(log, opdDiff(old, body.items || [])); } catch (err) { console.log('diff error: ' + err.message); }
    return { status: 'saved' };
  }

  return {status: 'unknown_action'};
}

// ── ฐานข้อมูล: อ่าน ─────────────────────────────────────────────────────
function readSheet(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getDataRange().getValues();
  const hdrs = rows[0].map(h => String(h).trim());

  return rows.slice(1).map(r => {
    const o = {};
    hdrs.forEach((h, i) => {
      if (!cols.includes(h)) return;
      let val = r[i];

      // 1. จัดการวันที่
      if (val instanceof Date) {
        if (isNaN(val.getTime())) val = '';
        else {
          const yr = val.getFullYear();
          const mo = String(val.getMonth() + 1).padStart(2, '0');
          const dy = String(val.getDate()).padStart(2, '0');
          const hh = String(val.getHours()).padStart(2, '0');
          const mm = String(val.getMinutes()).padStart(2, '0');

          // ถ้าเป็นคอลัมน์เวลา (ts, te) หรือปี 1899 (Google Sheets เก็บเวลาไว้ในปีนี้) ให้ดึงมาแค่เวลา
          if (h === 'ts' || h === 'te' || yr === 1899) {
            val = `${hh}:${mm}`;
          } else {
            // ถ้าเป็นคอลัมน์วันที่ปกติ
            val = `${yr}-${mo}-${dy}`;
          }
        }
      }

      // 2. ตรวจสอบและแปลง JSON (แก้ไขจุดที่ทำให้เกิด Error)
      if (h === 'sched' || h === 'orDays') {
        if (typeof val === 'string' && val.trim() !== '') {
          if (val.startsWith('{') || val.startsWith('[')) {
            try {
              val = JSON.parse(val);
            } catch(e) {
              console.log("JSON Parse error at " + h + ": " + e.message);
              // หาก Error ให้ส่งค่า Default สำรองเพื่อไม่ให้ Web App พัง
              val = (h === 'sched') ? {} : [];
            }
          }
        } else if (val === '' || val === null || val === undefined) {
          // 💡 จุดสำคัญ: ถ้าใน Sheet เป็นช่องว่างเปล่าๆ ให้แปลงเป็น Object/Array เริ่มต้นให้ Web App ทันที
          val = (h === 'sched') ? {} : [];
        }
      }

      // 3. จัดการคอลัมน์ di (ID แพทย์)
      if (h === 'di') val = (val === '' || val === null) ? 0 : Number(val);

      o[h] = val;
    });
    return o;
  }).filter(r => Object.values(r).some(v => v !== ''));
}

// ── ฐานข้อมูล: เขียน ─────────────────────────────────────────────────────
function write(ss, name, cols, rows) {
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);

  if (sh.getLastRow() === 0) {
    sh.appendRow(cols);
  } else {
    // ✅ บังคับเขียน header ให้ถูกต้องเสมอ
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);

    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      sh.getRange(2, 1, lastRow - 1, cols.length).clearContent();
    }
  }

  // เขียนข้อมูลใหม่ลงไป
  if (rows && rows.length > 0) {
    const data = rows.map(r => cols.map(c => {
      let v = r[c];
      if (v === undefined || v === null) return '';
      if (typeof v === 'object') return JSON.stringify(v);
      return v;
    }));
    sh.getRange(2, 1, data.length, cols.length).setValues(data);
  }
}

function upsert(ss, name, cols, row) {
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(cols);
  const data = sh.getDataRange().getValues();

  const formattedRow = cols.map(c => {
    let v = row[c];
    if (v === undefined || v === null) return '';
    if (typeof v === 'object') return JSON.stringify(v);
    return v;
  });

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(row.id)) {
      sh.getRange(i + 1, 1, 1, cols.length).setValues([formattedRow]);
      // คืนค่าแถวเดิมไว้เทียบว่าแก้อะไรไป (ใช้จดประวัติ)
      const old = {};
      data[0].forEach((h, j) => { old[String(h).trim()] = data[i][j]; });
      return old;
    }
  }
  sh.appendRow(formattedRow);
  return null;
}

// เขียนทับเฉพาะแถวที่ key ตรงกัน ถ้ายังไม่มีให้ต่อท้าย
function upsertKV(ss, name, items) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(OPD_COLS);
  // คอลัมน์ key ต้องเป็นข้อความเสมอ กัน Sheets แปลง key ที่หน้าตาคล้ายวันที่
  sh.getRange('A:B').setNumberFormat('@');
  const last = sh.getLastRow();
  const keys = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
  const rowOf = {};
  keys.forEach((k, i) => { rowOf[k] = i + 2; });
  const fresh = [];
  items.forEach(it => {
    const key = String(it.key || '');
    if (!key) return;
    const val = typeof it.value === 'string' ? it.value : JSON.stringify(it.value);
    if (rowOf[key] > 0) sh.getRange(rowOf[key], 2).setValue(val);
    else if (rowOf[key] < 0) fresh[-rowOf[key] - 1][1] = val;
    else { fresh.push([key, val]); rowOf[key] = -fresh.length; }
  });
  if (fresh.length) sh.getRange(sh.getLastRow() + 1, 1, fresh.length, 2).setValues(fresh);
}

function debugPost_doctors() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName('ENT_Doctors');

  Logger.log("=== ENT_Doctors Debug ===");
  Logger.log("Headers: " + JSON.stringify(sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]));

  const data = readSheet(ss, 'ENT_Doctors', DOC_COLS);
  data.forEach((d, i) => {
    Logger.log(`Row ${i+1}: di=${d.di}, name=${d.name}, sched=${JSON.stringify(d.sched)}, orDays=${JSON.stringify(d.orDays)}`);
  });
}

// ══ ประวัติการแก้ไข (ENT_Log) ═══════════════════════════════════════
// แต่ละรายการ = [section, action, item, detail] เช่น ['นัดผ่าตัด','แก้','HN 123 นาย ก','สถานะ: รอยืนยัน → ยืนยันแล้ว']

function appendLog(ss, editor, entries) {
  if (!entries || !entries.length) return;
  const who = String(editor || '').trim().slice(0, 120) || '(ไม่ระบุ)';
  let list = entries;
  if (list.length > LOG_MAX_PER_SAVE) {
    const more = list.length - LOG_MAX_PER_SAVE;
    list = list.slice(0, LOG_MAX_PER_SAVE).concat([['ระบบ', 'สรุป', '', 'และอีก ' + more + ' รายการ']]);
  }
  const sh = ss.getSheetByName('ENT_Log') || ss.insertSheet('ENT_Log');
  if (sh.getLastRow() === 0) { sh.appendRow(LOG_COLS); sh.setFrozenRows(1); }
  const now = new Date();
  const rows = list.map(en => [now, who].concat(en.map(safeCell)));
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, LOG_COLS.length).setValues(rows);
}

// ข้อความที่ขึ้นต้นด้วย = + - @ Sheets จะมองเป็นสูตร — เติม ' นำหน้าให้เป็นข้อความธรรมดา
function safeCell(v) {
  const t = String(v == null ? '' : v);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function readLog(ss, n) {
  const sh = ss.getSheetByName('ENT_Log');
  if (!sh || sh.getLastRow() < 2) return [];
  const last = sh.getLastRow(), start = Math.max(2, last - n + 1);
  const tz = Session.getScriptTimeZone();
  return sh.getRange(start, 1, last - start + 1, LOG_COLS.length).getValues().reverse().map(r => {
    const o = {};
    LOG_COLS.forEach((c, i) => {
      o[c] = (r[i] instanceof Date) ? Utilities.formatDate(r[i], tz, 'yyyy-MM-dd HH:mm:ss') : String(r[i]);
    });
    return o;
  });
}

// อ่านแท็บแบบดิบ (ไม่แปลงค่า) เป็น object ตามหัวตาราง
function readRaw(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getDataRange().getValues();
  const hdrs = rows[0].map(h => String(h).trim());
  return rows.slice(1).map(r => {
    const o = {};
    hdrs.forEach((h, i) => { if (h) o[h] = r[i]; });
    return o;
  }).filter(o => Object.keys(o).some(k => o[k] !== ''));
}

// เขียนทั้งแท็บ + จดเฉพาะรายการที่เปลี่ยน
function writeLogged(ss, name, cols, rows, spec, log) {
  const old = readRaw(ss, name);
  write(ss, name, cols, rows);
  safeDiff(log, old, rows, spec, false);
}

// การจดประวัติต้องไม่ทำให้การบันทึกข้อมูลล้ม
function safeDiff(log, oldRows, newRows, spec, onlyUpsert) {
  try { log.push.apply(log, diffRows(oldRows, newRows, spec, onlyUpsert)); }
  catch (err) { console.log('diff error: ' + err.message); }
}

// ค่าในเซลล์ → ข้อความสำหรับเทียบ/แสดง
function nv(v, h) {
  if (v === undefined || v === null) return '';
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    const tz = Session.getScriptTimeZone();
    if (h === 'ts' || h === 'te' || v.getFullYear() === 1899) return Utilities.formatDate(v, tz, 'HH:mm');
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  if (typeof v === 'object') return JSON.stringify(v);
  return String(v).trim();
}
// ถือว่าเท่ากัน ถ้าข้อความตรงกัน หรือเป็นตัวเลขค่าเดียวกัน (Sheets ตัด 0 หน้าเบอร์โทร / แปลง "5" เป็น 5)
function same(a, b) {
  if (a === b) return true;
  if (a === '' || b === '') return false;
  return !isNaN(a) && !isNaN(b) && Number(a) === Number(b);
}
function cut(t, n) { t = String(t); n = n || 60; return t.length > n ? t.slice(0, n - 1) + '…' : t; }

// spec = {section, key(row), label(row), fields:{col:'ชื่อที่แสดง'}, show(col,val), brief:[cols ที่สรุปตอนเพิ่ม/ลบ]}
function diffRows(oldRows, newRows, spec, onlyUpsert) {
  const out = [];
  const oldBy = {};
  oldRows.forEach(r => { const k = nv(spec.key(r)); if (k !== '') oldBy[k] = r; });
  const seen = {};
  const show = (c, v) => cut(spec.show ? spec.show(c, nv(v, c)) : nv(v, c));
  const brief = r => (spec.brief || []).map(c => nv(r[c], c) !== '' ? spec.fields[c] + ' ' + show(c, r[c]) : '').filter(Boolean).join(' · ');
  newRows.forEach(r => {
    const k = nv(spec.key(r));
    if (k === '') return;
    seen[k] = 1;
    const o = oldBy[k];
    if (!o) { out.push([spec.section, 'เพิ่ม', spec.label(r), brief(r)]); return; }
    const ch = [];
    Object.keys(spec.fields).forEach(c => {
      const a = nv(o[c], c), b = nv(r[c], c);
      if (!same(a, b)) ch.push(spec.fields[c] + ': ' + (a === '' ? '–' : show(c, o[c])) + ' → ' + (b === '' ? '–' : show(c, r[c])));
    });
    if (ch.length) out.push([spec.section, 'แก้', spec.label(r), ch.join(' ; ')]);
  });
  if (!onlyUpsert) {
    Object.keys(oldBy).forEach(k => {
      if (!seen[k]) out.push([spec.section, 'ลบ', spec.label(oldBy[k]), brief(oldBy[k])]);
    });
  }
  return out;
}

const APT_STATUS = { confirm: 'ยืนยันแล้ว', wait: 'รอยืนยัน', pending: 'รอยืนยัน', cancel: 'ยกเลิก' };
const PREOP_STATUS = { done: 'เรียบร้อย', wait: 'รอตรวจ', none: 'ยังไม่นัด' };
const DAY_TH = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
const DAY_SH = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];

function aptSpec() {
  return {
    section: 'นัดผ่าตัด',
    key: r => r.id,
    label: r => ['HN ' + nv(r.hn), nv(r.name)].filter(Boolean).join(' '),
    fields: { hn: 'HN', name: 'ชื่อ', date: 'วันที่', ts: 'เริ่ม', te: 'ถึง', op: 'หัตถการ', doctorName: 'แพทย์',
      doctor2Name: 'แพทย์ร่วม', anesthesia: 'ระงับความรู้สึก', tel1: 'โทร 1', tel2: 'โทร 2', status: 'สถานะ', note: 'หมายเหตุ',
      preopDate: 'วัน Pre-Op', preopStatus: 'สถานะ Pre-Op', lab: 'Lab', cxr: 'CXR', ekg: 'EKG', npo: 'NPO', preopNote: 'หมายเหตุ Pre-Op' },
    brief: ['date', 'op', 'doctorName', 'status'],
    show: (c, v) => c === 'status' ? (APT_STATUS[v] || v) : c === 'preopStatus' ? (PREOP_STATUS[v] || v)
      : (c === 'lab' || c === 'cxr' || c === 'ekg') ? (v ? '✓' : '–') : v,
  };
}
// ชื่อแพทย์ตาม di — ใช้รายชื่อที่ส่งมาพร้อมกัน ถ้าไม่มีอ่านจากแท็บ ENT_Doctors
function docNames(ss, rows) {
  const m = {};
  (rows || readRaw(ss, 'ENT_Doctors')).forEach((d, i) => { m[nv(d.di !== undefined && d.di !== '' ? d.di : i)] = nv(d.name); });
  return m;
}
function leaveSpec(names) {
  const dn = di => names[nv(di)] || ('แพทย์ #' + nv(di));
  return {
    section: 'วันลาแพทย์',
    key: r => r.id,
    label: r => dn(r.di) + ' ' + nv(r.start, 'start') + (nv(r.end, 'end') && nv(r.end, 'end') !== nv(r.start, 'start') ? ' ถึง ' + nv(r.end, 'end') : ''),
    fields: { di: 'แพทย์', start: 'ตั้งแต่', end: 'ถึง', reason: 'เหตุผล', status: 'สถานะ' },
    brief: ['reason', 'status'],
    show: (c, v) => c === 'di' ? dn(v) : v,
  };
}
function swapSpec(names) {
  const dn = di => names[nv(di)] || ('แพทย์ #' + nv(di));
  return {
    section: 'สลับวันผ่าตัด',
    key: r => r.id,
    label: r => dn(r.di) + ' ' + nv(r.date, 'date'),
    fields: { di: 'แพทย์', date: 'วันที่', type: 'ประเภท', note: 'หมายเหตุ' },
    brief: ['type', 'note'],
    show: (c, v) => c === 'di' ? dn(v) : v,
  };
}
function docSpec() {
  return {
    section: 'แพทย์',
    key: r => r.di,
    label: r => nv(r.name),
    fields: { name: 'ชื่อ', color: 'สี', sched: 'ตารางผ่าตัด', orDays: 'วันผ่าตัด' },
    brief: [],
    show: (c, v) => c === 'orDays' ? dayList(v) : v,
  };
}
function opSpec() {
  return { section: 'ประเภทการผ่าตัด', key: r => r.name, label: r => nv(r.name), fields: {}, brief: [] };
}
function dayList(v) {
  try { const a = typeof v === 'string' ? JSON.parse(v) : v; if (Array.isArray(a)) return a.map(d => DAY_SH[d] || d).join(','); } catch (e) {}
  return v;
}
function pj(v) {
  if (v && typeof v === 'object') return v;
  try { return JSON.parse(v); } catch (e) { return null; }
}

// ค่าตั้งระบบ: วันผ่าตัดของแผนก / วันหยุดราชการ
function configDiff(oldRows, newRows) {
  const out = [];
  const lbl = { orDays: 'วันผ่าตัดของแผนก', holidays: 'วันหยุดราชการ' };
  const om = {};
  oldRows.forEach(r => { om[nv(r.key)] = nv(r.value); });
  newRows.forEach(r => {
    const k = nv(r.key), a = om[k] || '', b = nv(r.value);
    if (same(a, b)) return;
    if (k === 'holidays') {
      const A = pj(a) || {}, B = pj(b) || {};
      Object.keys(B).forEach(d => {
        if (!(d in A)) out.push(['วันหยุด', 'เพิ่ม', d, cut(B[d])]);
        else if (String(A[d]) !== String(B[d])) out.push(['วันหยุด', 'แก้', d, cut(A[d]) + ' → ' + cut(B[d])]);
      });
      Object.keys(A).forEach(d => { if (!(d in B)) out.push(['วันหยุด', 'ลบ', d, cut(A[d])]); });
    } else {
      const show = k === 'orDays' ? dayList : (x => cut(x));
      out.push(['ค่าตั้งระบบ', 'แก้', lbl[k] || k, (a ? show(a) : '–') + ' → ' + (b ? show(b) : '–')]);
    }
  });
  return out;
}

// ตาราง OPD: แปลง key/value (JSON) เป็นข้อความที่อ่านรู้เรื่อง
function opdDiff(old, items) {
  const out = [];
  const cur = k => { const it = items.find(x => String(x.key) === k); return it ? (typeof it.value === 'string' ? it.value : JSON.stringify(it.value)) : old[k]; };
  const docs = pj(cur('doctors')) || [], set = pj(cur('settings')) || {}, orv = pj(cur('or')) || {};
  const rooms = (set.rooms && set.rooms.length) ? set.rooms : ['ห้อง 1', 'ห้อง 2'];
  const orRows = (orv.rows && orv.rows.length) ? orv.rows : [];
  const dn = id => { const d = docs.find(x => String(x.id) === String(id)); return d ? d.name : '?'; };
  const ses = s => s === 'am' ? 'เช้า' : s === 'pm' ? 'บ่าย' : s;
  const slotName = k => { const p = String(k).split('-'); return (rooms[+p[0]] || ('ห้อง ' + (+p[0] + 1))) + ' ' + ses(p[1]); };
  const tplWho = list => (list || []).map(e => dn(e.d) + (e.t ? ' (' + e.t + ')' : '') + (e.w && e.w.length ? ' [สัปดาห์ ' + e.w.join(',') + ']' : '')).join(' / ') || '–';
  const orWho = list => (list || []).map(e => ((e.d || []).map(dn).join('+') || 'งด') + (e.w && e.w.length ? ' [สัปดาห์ ' + e.w.join(',') + ']' : '')).join(' / ') || '–';
  const mapDiff = (A, B, name, who) => {
    const ch = [];
    const keys = Object.keys(A || {}).concat(Object.keys(B || {})).filter((k, i, a) => a.indexOf(k) === i);
    keys.forEach(k => {
      const a = (A || {})[k], b = (B || {})[k];
      if (JSON.stringify(a) === JSON.stringify(b)) return;
      ch.push(name(k) + ': ' + (a == null ? '–' : who(a)) + ' → ' + (b == null ? '–' : who(b)));
    });
    return ch;
  };
  items.forEach(it => {
    const k = String(it.key || '');
    const b = typeof it.value === 'string' ? it.value : JSON.stringify(it.value);
    const a = old[k];
    if (!k || a === b) return;
    const A = pj(a), B = pj(b);
    let item = k, ch = [];
    if (k === 'doctors') {
      item = 'รายชื่อแพทย์ OPD';
      const ids = (A || []).map(d => String(d.id)), nids = (B || []).map(d => String(d.id));
      (B || []).forEach(d => {
        const i = ids.indexOf(String(d.id));
        if (i < 0) ch.push('เพิ่ม ' + d.name);
        else if (JSON.stringify(A[i]) !== JSON.stringify(d)) ch.push('แก้ ' + (A[i].name !== d.name ? A[i].name + ' → ' : '') + d.name);
      });
      (A || []).forEach(d => { if (nids.indexOf(String(d.id)) < 0) ch.push('ลบ ' + d.name); });
    } else if (k === 'settings') {
      item = 'ค่าตั้ง OPD';
      ch = mapDiff(A, B, x => x, v => cut(Array.isArray(v) ? v.join(', ') : v));
    } else if (k === 'template') {
      item = 'แม่แบบรายสัปดาห์';
      ch = mapDiff(A, B, x => { const p = x.split('|'); return (DAY_TH[+p[0]] || p[0]) + ' ' + slotName(p[1]); }, tplWho);
    } else if (k === 'or') {
      item = 'ตาราง OR (แม่แบบ)';
      ch = mapDiff((A || {}).c, (B || {}).c, x => { const p = x.split('|'); return (DAY_TH[+p[0]] || p[0]) + ' ' + (orRows[+p[1]] || ('OR ' + (+p[1] + 1))); }, orWho);
      if (JSON.stringify((A || {}).rows) !== JSON.stringify((B || {}).rows)) ch.push('แถว OR: ' + cut(((B || {}).rows || []).join(', ')));
    } else if (k.charAt(0) === 'd') {
      item = 'OPD ' + k.slice(1);
      const x = A || {}, y = B || {};
      ch = mapDiff(x.s, y.s, slotName, v => dn(v.d) + (v.t ? ' (' + v.t + ')' : ''));
      if ((x.c || '') !== (y.c || '')) ch.push('วันหยุด: ' + (x.c || '–') + ' → ' + (y.c || '–'));
      if (!!x.o !== !!y.o) ch.push(y.o ? 'เปิดตรวจวันเสาร์/อาทิตย์' : 'ยกเลิกเปิดตรวจวันเสาร์/อาทิตย์');
      if (JSON.stringify(x.x || {}) !== JSON.stringify(y.x || {})) ch.push('ปรับช่องจากแม่แบบ');
      ch = ch.concat(mapDiff(x.or, y.or, r => orRows[+r] || ('OR ' + (+r + 1)), v => (v || []).map(dn).join('+') || 'งด'));
    }
    if (!ch.length) ch = ['แก้ไข'];
    out.push(['ตาราง OPD', a == null ? 'เพิ่ม' : 'แก้', item, cut(ch.join(' ; '), 500)]);
  });
  return out;
}

function ok(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function debugGetAll() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheetName = 'ENT_Schedule';
  const sh = ss.getSheetByName(sheetName);

  if (!sh) {
    Logger.log("❌ หาแผ่นงานชื่อ " + sheetName + " ไม่เจอ!");
  } else {
    Logger.log("✅ เจอแผ่นงาน! จำนวนแถวทั้งหมดคือ: " + sh.getLastRow());
    const data = readSheet(ss, sheetName, APT_COLS);
    Logger.log("📊 ข้อมูลที่ดึงได้: " + JSON.stringify(data));
  }
}

function checkMyHeaders() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName('ENT_Schedule');
  const actualHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  Logger.log("--- ตรวจสอบหัวตาราง ---");
  Logger.log("สิ่งที่อยู่ใน Sheet: " + JSON.stringify(actualHeaders));
  Logger.log("สิ่งที่ Code ต้องการ: " + JSON.stringify(APT_COLS));

  actualHeaders.forEach((h, i) => {
    if (APT_COLS.includes(String(h).trim())) {
      Logger.log("✅ คอลัมน์ที่ " + (i+1) + " [" + h + "] -> ถูกต้อง");
    } else {
      Logger.log("❌ คอลัมน์ที่ " + (i+1) + " [" + h + "] -> ไม่ตรง! (ต้องแก้ใน Sheet ให้เป็น " + (APT_COLS[i] || 'ค่าที่กำหนด') + ")");
    }
  });
}
