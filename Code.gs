// ══ Google Apps Script (Code.gs) – ENT OR Schedule v5 ══════════
// อัพเดทจาก v4 (Fixed): เพิ่ม saveLeaves / saveSwaps / saveConfig + ส่ง config กลับใน getAll
// v5 navy: เพิ่มคอลัมน์ Pre-Op (preopDate, preopStatus, lab, cxr, ekg, npo, preopNote, history)
// Sheet tabs: ENT_Schedule, ENT_Doctors, ENT_Ops, ENT_Leaves, ENT_Swaps, ENT_Config
const SHEET_ID = '1KWH-9JobfctIp-prqGSlZIJ4xOuP4Yb6QVA0spTE8sk';

// ★ v5: เสิร์ฟหน้าเว็บล่าสุดจาก GitHub (branch main) โดยตรง
// → อัพเดทโค้ดที่ GitHub ที่เดียว ผู้ใช้ /exec ทุกคนได้เวอร์ชันใหม่อัตโนมัติ
//   ไม่ต้อง copy index.html มาวางใน Apps Script และไม่ต้อง Deploy ใหม่อีก
// หมายเหตุ: ถ้าเปลี่ยน repo เป็น private ภายหลัง ลิงก์ raw นี้จะใช้ไม่ได้ ต้องกลับไปใช้ไฟล์ index ที่ฝังไว้
const LIVE_HTML_URL = 'https://raw.githubusercontent.com/freedomnew2009-collab/ent-or-schedule/main/index.html';
const APP_VERSION = 'v5';

const APT_COLS   = ['id','hn','name','date','ts','te','op','doctorName','di','di2','doctor2Name','anesthesia','tel1','tel2','status','note','preopDate','preopStatus','lab','cxr','ekg','npo','preopNote','history'];
const LEAVE_COLS = ['id','di','start','end','reason','status'];
const SWAP_COLS  = ['id','di','date','type','note'];
const DOC_COLS   = ['di','name','color','sched','orDays'];
const OP_COLS    = ['name'];
const CFG_COLS   = ['key','value'];   // ★ v5: เก็บค่าตั้งระบบ เช่น วันผ่าตัดของแผนก (orDays)

function doGet(e) {
  const p = e.parameter || {};
  // เปิด URL?action=ping ใน browser เพื่อเช็คว่า deployment เป็นเวอร์ชันล่าสุดหรือยัง
  if (p.action === 'ping') return ok({ status: 'online', version: APP_VERSION, serving: 'github-main' });

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
      return ok({
        appointments: readSheet(ss, p.sheet || 'ENT_Schedule', APT_COLS),
        leaves:       readSheet(ss, 'ENT_Leaves',  LEAVE_COLS),
        swaps:        readSheet(ss, 'ENT_Swaps',   SWAP_COLS),
        doctors:      readSheet(ss, 'ENT_Doctors', DOC_COLS),
        operations:   readSheet(ss, 'ENT_Ops',     OP_COLS),
        config:       config,
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

  if(body.action === 'upsert'){
    upsert(ss, body.sheet, APT_COLS, body.row);
    return ok({status: 'upserted', id: body.row.id});
  }

  if(body.action === 'pushAll' || body.action === 'saveSettings'){
    if(body.appointments) write(ss, body.sheet || 'ENT_Schedule', APT_COLS, body.appointments);
    if(body.leaves)       write(ss, 'ENT_Leaves', LEAVE_COLS, body.leaves);
    if(body.swaps)        write(ss, 'ENT_Swaps',  SWAP_COLS,  body.swaps);
    if(body.doctors)      write(ss, 'ENT_Doctors', DOC_COLS,  body.doctors);
    if(body.operations)   write(ss, 'ENT_Ops',     OP_COLS,   body.operations);
    return ok({status: 'success'});
  }

  // ★ v5: บันทึกวันลา / Override / ค่าตั้งระบบ แยกส่วน (หน้าเว็บซิงค์อัตโนมัติทุกครั้งที่แก้ไข)
  if (body.action === 'saveLeaves') {
    write(ss, 'ENT_Leaves', LEAVE_COLS, body.leaves || []);
    return ok({ status: 'saved' });
  }
  if (body.action === 'saveSwaps') {
    write(ss, 'ENT_Swaps', SWAP_COLS, body.swaps || []);
    return ok({ status: 'saved' });
  }
  if (body.action === 'saveConfig') {
    const rows = Object.keys(body.config || {}).map(k => ({ key: k, value: body.config[k] }));
    write(ss, 'ENT_Config', CFG_COLS, rows);
    return ok({ status: 'saved' });
  }

  return ok({status: 'unknown_action'});
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
      return;
    }
  }
  sh.appendRow(formattedRow);
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
