/**********************************************************************
 * Code.gs — โหมด STANDALONE ล้วน
 * - อ้างสเปรดชีตด้วย CONFIG.SPREADSHEET_ID เสมอ
 * - ค้นหาผ่าน Web App (doGet) เท่านั้น (ไม่มีเมนู/Sidebar แบบ bound)
 * - ตั้ง API key และสร้างดัชนี โดยรันฟังก์ชันจาก editor
 **********************************************************************/

/* ---------- ตัวช่วยอ้างชีต ---------- */
function getSS_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

/** log อย่างเดียว (standalone ไม่มี UI) */
function notify_(msg) { Logger.log(msg); }

/* ---------- ตั้งค่า / บำรุงรักษา (รันจาก editor) ---------- */

/** ตั้ง OpenRouter API key: แก้เป็น saveApiKey("sk-or-...") แล้ว Run ครั้งเดียว (เสร็จแล้วลบคีย์ออกจากโค้ด) */
function saveApiKey(key) {
  key = (key || '').trim();
  if (!key) throw new Error('เรียกแบบ saveApiKey("sk-or-v1-...") พร้อมคีย์จริง');
  PropertiesService.getScriptProperties().setProperty('OPENROUTER_API_KEY', key);
  Logger.log('บันทึก OpenRouter API Key แล้ว ✅');
}

/* ---------- อัปเดตอัตโนมัติเมื่อมีไฟล์ใหม่เข้าโฟลเดอร์ ----------
 * Apps Script ไม่มี trigger "ไฟล์เข้าโฟลเดอร์" แบบเรียลไทม์ จึงใช้ trigger ตามเวลา (ทุก 15 นาที)
 * เรียก buildIndex ซึ่งอ่านเฉพาะไฟล์ใหม่/แก้ไข (cache ตาม modifiedTime) แล้วประกอบ Index ใหม่
 * รัน installAutoUpdateTrigger() หนึ่งครั้งจาก editor เพื่อเริ่ม (และ removeAutoUpdateTrigger() เพื่อหยุด)
 */
const AUTO_UPDATE_FN = 'autoUpdate';
const AUTO_UPDATE_EVERY_MIN = 15; // ค่าที่ Apps Script รองรับ: 1, 5, 10, 15, 30

/** ติดตั้ง trigger อัตโนมัติทุก 15 นาที (รันครั้งเดียว) */
function installAutoUpdateTrigger() {
  removeAutoUpdateTrigger_();
  ScriptApp.newTrigger(AUTO_UPDATE_FN).timeBased().everyMinutes(AUTO_UPDATE_EVERY_MIN).create();
  Logger.log('ติดตั้ง trigger อัตโนมัติทุก ' + AUTO_UPDATE_EVERY_MIN + ' นาทีแล้ว ✅');
}

/** ลบ trigger อัตโนมัติทั้งหมด */
function removeAutoUpdateTrigger() {
  const n = removeAutoUpdateTrigger_();
  Logger.log('ลบ trigger อัตโนมัติแล้ว (' + n + ' ตัว)');
}
function removeAutoUpdateTrigger_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === AUTO_UPDATE_FN) { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}

/** ตัวที่ trigger เรียก — กันรอบทับกันด้วย LockService (buildIndex อาจใช้เวลาถึง ~5 นาที) */
function autoUpdate() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) { Logger.log('autoUpdate: มีงานทำอยู่ ข้ามรอบนี้'); return; }
  try { buildIndex(); }
  catch (e) { Logger.log('autoUpdate error: ' + e); }
  finally { lock.releaseLock(); }
}

/** ล้าง cache (รอบถัดไปจะอ่าน PDF ใหม่ทั้งหมด) */
function clearCache() {
  const sh = getSS_().getSheetByName(CONFIG.CACHE_SHEET);
  if (sh) {
    sh.clearContents();
    // เขียนหัวตารางกลับทันที มิฉะนั้นแถวข้อมูลแรกรอบหน้าจะถูกอ่านเป็นหัวตาราง (bug อ่าน PDF ซ้ำ/แถวซ้ำ)
    const headers = ['fileId', 'modifiedTime', 'fileName', 'url', 'folder', 'json'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  Logger.log('ล้าง Cache แล้ว');
}

/* ---------- Web App ---------- */
function doGet() {
  // ใช้ค่าปริยาย (DEFAULT) ของ HtmlService — ยัง embed ใน Google Sites ได้ แต่กันฝังในเว็บภายนอกใด ๆ (clickjacking)
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .setTitle('ระบบสืบค้นบัญชีคุม')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- Search backend (เรียกจาก HTML ด้วย google.script.run) ---------- */

/** ค้นหา: params = {q, category, number} */
function apiSearch(params) {
  params = params || {};
  let res = readIndex_();
  if (params.number) {
    const n = Number(params.number);
    // ค้นได้ทั้งเลขตามลำดับบัญชีคุม และเลขเดิมในเอกสาร PDF
    res = res.filter(r => r.number === n || r.docNumber === n);
  }
  if (params.category) {
    res = res.filter(r => r.category === params.category);
  }
  if (params.q) {
    const q = String(params.q).toLowerCase().trim();
    res = res.filter(r =>
      [r.name, r.owner_unit, r.project, r.category, r.subtags, r.summary, r.search, r.code, String(r.number), String(r.docNumber)]
        .join(' ').toLowerCase().indexOf(q) >= 0);
  }
  const LIMIT = 300;
  // คืน total ด้วย เพื่อให้ฝั่งเว็บรู้ว่าผลถูกตัด (ไม่โชว์จำนวนที่ตัดแล้วว่าเป็นทั้งหมด)
  return { items: res.slice(0, LIMIT), total: res.length };
}

/** รายการหมวดหมู่ + จำนวน (สำหรับปุ่มกรอง) */
function apiCategories() {
  const data = readIndex_();
  const m = {};
  data.forEach(r => { const c = r.category || 'อื่น ๆ'; m[c] = (m[c] || 0) + 1; });
  return Object.keys(m).sort().map(k => ({ category: k, count: m[k] }));
}

/** สรุปจำนวนรวม (โชว์บนหัวหน้าเว็บ) — อ่านดัชนีครั้งเดียว */
function apiStats() {
  const data = readIndex_();
  const seen = {};
  data.forEach(r => { seen[r.category || 'อื่น ๆ'] = 1; });
  return { total: data.length, categories: Object.keys(seen).length };
}

/** วินิจฉัยระบบ — ช่วยหาสาเหตุ "ไม่มีข้อมูล" */
function apiDiag() {
  const ss = getSS_();
  const idx = ss.getSheetByName(CONFIG.INDEX_SHEET);
  const cache = ss.getSheetByName(CONFIG.CACHE_SHEET);
  const ctrl = ss.getSheetByName(CONFIG.CONTROL_SHEET);
  const ledger = ss.getSheetByName(CONFIG.LEDGER_SHEET);
  return {
    spreadsheetName: ss.getName(),
    hasApiKey: !!PropertiesService.getScriptProperties().getProperty('OPENROUTER_API_KEY'),
    indexExists: !!idx,
    indexRows: idx ? Math.max(0, idx.getLastRow() - 1) : 0,
    cacheExists: !!cache,
    cacheRows: cache ? Math.max(0, cache.getLastRow() - 1) : 0,
    controlExists: !!ctrl,
    ledgerExists: !!ledger,
    ledgerRows: ledger ? Math.max(0, ledger.getLastRow() - 1) : 0
  };
}

/** อ่านชีต Index เป็น array ของ object */
function readIndex_() {
  const sh = getSS_().getSheetByName(CONFIG.INDEX_SHEET);
  if (!sh) return [];
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if (!r[0] && !r[2]) continue; // ข้ามแถวว่าง (ไม่มีทั้งเลขและชื่อ)
    out.push({
      code: r[0],                                                 // "ส-001" ตามบัญชีคุม
      number: parseInt(String(r[0]).split('/')[0].replace(/[^0-9]/g, ''), 10) || null, // เลขจริงสำหรับค้น/เรียง (ตัด /ปีงบ ออกก่อน)
      docNumber: Number(r[1]) || null,                            // เลขที่ OCR อ่านได้
      name: r[2], unit: r[3], category: r[4], subtags: r[5],
      owner_unit: r[6], file: r[7], url: r[8], projectRow: r[9], project: r[10],
      status: r[11], budget: r[12], summary: r[13], search: r[14]
    });
  }
  return out;
}

/* ---------- Helpers ทั่วไป ---------- */

function getOrCreateSheet_(name, headers) {
  const ss = getSS_();
  let sh = ss.getSheetByName(name);
  // เผื่อชีตถูกตั้งชื่อโดยมีช่องว่างหน้า/หลังติดมา (เช่น "Index_อุปกรณ์ ") — หาแบบ trim ก่อนสร้างใหม่ กันชีตซ้ำ
  if (!sh) sh = ss.getSheets().find(s => s.getName().trim() === String(name).trim());
  if (!sh) sh = ss.insertSheet(name);
  // เขียนหัวตารางเสมอเมื่อชีตยังว่าง — สำคัญหลัง clearCache()/clearContents ที่ลบหัวตารางทิ้ง
  // มิฉะนั้น loadCache_ จะเข้าใจว่าแถวข้อมูลแรกคือหัวตาราง (อ่านข้าม) ทำให้อ่าน PDF ซ้ำและเกิดแถวซ้ำทุกครั้ง
  if (headers && sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sh;
}

/** normalize ข้อความไทยสำหรับจับคู่ (ตัดช่องว่าง/จุด/วรรคตอน) */
function norm_(s) {
  return String(s || '').toLowerCase().replace(/[\s\.\,\(\)\-_/]/g, '');
}
