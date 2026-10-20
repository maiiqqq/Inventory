/**********************************************************************
 * Index.gs — สร้างดัชนี: ไล่ PDF -> OpenRouter (cache) -> เชื่อมบัญชีคุม -> เขียน Index
 * รันซ้ำได้ (idempotent). ถ้าใกล้ชน 6 นาที จะหยุด ให้รัน buildIndex อีกครั้งเพื่อทำต่อ
 **********************************************************************/

function buildIndex() {
  getApiKey_(); // เช็คก่อนว่าตั้ง key แล้ว

  const files = listPdfFiles_();
  const { sh: cacheSh, map: cache } = loadCache_();
  const t0 = Date.now();
  let processed = 0, skipped = 0, remaining = 0, failed = 0;

  for (const f of files) {
    if (cache[f.id] && cache[f.id].modifiedTime === f.modifiedTime) { skipped++; continue; }
    if (Date.now() - t0 > CONFIG.MAX_RUNTIME_MS) { remaining++; continue; }
    try {
      const items = openrouterExtractFromPdf_(f);
      saveCacheRow_(cacheSh, cache, f, items);
      processed++;
    } catch (e) {
      failed++;
      Logger.log('ERR ' + f.name + ': ' + e);
      notify_('ผิดพลาดไฟล์ ' + f.name + ': ' + e.message);
    }
  }

  rebuildIndexSheet_(cache);

  let msg = `รอบนี้: อ่านใหม่ ${processed} • ข้าม(เดิม) ${skipped} • พลาด ${failed}`;
  if (remaining > 0) msg += ` • เหลืออีก ${remaining} ไฟล์ → รัน buildIndex อีกครั้งเพื่อทำต่อ`;
  else msg += ' • ✅ เสร็จสมบูรณ์';
  notify_(msg);
  return msg;
}

/* ---------- ไล่ไฟล์ PDF ทุกโฟลเดอร์ย่อย + ถอดช่วงหมายเลขจากชื่อไฟล์ ---------- */
function listPdfFiles_() {
  const root = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const out = [];
  (function walk(folder) {
    const files = folder.getFilesByType(MimeType.PDF);
    while (files.hasNext()) {
      const f = files.next();
      const m = f.getName().match(/^(\d+)(?:\s*-\s*(\d+))?\.pdf$/i);
      out.push({
        id: f.getId(),
        name: f.getName(),
        url: f.getUrl(),
        modifiedTime: f.getLastUpdated().toISOString(),
        start: m ? Number(m[1]) : null,
        end: m ? (m[2] ? Number(m[2]) : Number(m[1])) : null,
        folder: folder.getName()
      });
    }
    const subs = folder.getFolders();
    while (subs.hasNext()) walk(subs.next());
  })(root);
  return out;
}

/* ---------- Cache (ชีต _Cache) ---------- */
function loadCache_() {
  const sh = getOrCreateSheet_(CONFIG.CACHE_SHEET, ['fileId', 'modifiedTime', 'fileName', 'url', 'folder', 'json']);
  const data = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    let items = [], ok = true;
    try { items = JSON.parse(r[5] || '[]'); } catch (e) { ok = false; }
    // ถ้า JSON เสีย (พัง/ถูกตัดเพราะเกิน 50k ต่อเซลล์/แก้มือ) ให้ modifiedTime = '' เพื่อบังคับอ่านไฟล์ใหม่รอบหน้า
    // (แทนที่แถวเดิมในที่เดิม ไม่เกิดแถวซ้ำ) แทนที่จะปล่อยให้ไฟล์นั้นหายจากดัชนีถาวร
    map[r[0]] = { rowIndex: i + 1, modifiedTime: ok ? r[1] : '', fileName: r[2], url: r[3], folder: r[4], items: items };
  }
  return { sh, map };
}

function saveCacheRow_(sh, map, f, items) {
  const json = JSON.stringify(items);
  if (map[f.id]) {
    sh.getRange(map[f.id].rowIndex, 1, 1, 6).setValues([[f.id, f.modifiedTime, f.name, f.url, f.folder, json]]);
    map[f.id] = { rowIndex: map[f.id].rowIndex, modifiedTime: f.modifiedTime, fileName: f.name, url: f.url, folder: f.folder, items };
  } else {
    sh.appendRow([f.id, f.modifiedTime, f.name, f.url, f.folder, json]);
    map[f.id] = { rowIndex: sh.getLastRow(), modifiedTime: f.modifiedTime, fileName: f.name, url: f.url, folder: f.folder, items };
  }
}

/* ---------- ประกอบชีต Index จาก cache + เชื่อมบัญชีคุม + สถานะโครงการ ----------
 * "หมายเลขอุปกรณ์" = เลขทางการจากชีตบัญชีคุม (คอลัมน์ หมายเลข → แสดง "ส-001")
 *   จับคู่โดย "ชื่อสิ่งอุปกรณ์" ที่ OCR อ่านได้ กับชื่อในบัญชีคุม
 *   จำกัดผู้สมัครด้วย "ช่วงเลขจากชื่อไฟล์ PDF" (เช่น 100-103.pdf = เลข 100–103) → จับคู่แม่นขึ้นมาก
 * เลขที่ OCR อ่านได้เก็บไว้คอลัมน์ "เลขที่ OCR อ่าน" เพื่อเทียบ/ตรวจ (ถ้าไม่ตรงบัญชีคุมจะเห็น)
 * สถานะ/วงเงิน มาจากชีตสถานะโครงการ (จับคู่ชื่อโครงการ)
 */
function rebuildIndexSheet_(cache) {
  const control = readControlSheet_(); // สถานะโครงการ (สถานะ/วงเงิน)
  const ledger = readLedger_();        // บัญชีคุม (เลขอุปกรณ์ทางการ)

  const items = [];
  Object.keys(cache).forEach(fid => {
    const c = cache[fid];
    const list = c.items || [];
    if (!list.length) return;
    const range = parseRange_(c.fileName); // ช่วงเลขจากชื่อไฟล์ (อาจเป็น [null,null])
    list.forEach(it => {
      const led = ledger.matchName(it.name, range);            // แถวบัญชีคุมที่ตรงชื่อที่สุด
      const proj = (led && led.project) || it.project || '';
      const unit = it.owner_unit || (led && led.unit) || '';
      const link = control.matchProject(proj, unit);           // สถานะ/วงเงิน จากสถานะโครงการ
      items.push({
        number: led ? led.number : (Number(it.number) || null), // ไว้เรียง/ค้น
        code:   led ? (led.code + (led.year ? '/' + led.year : '')) : (it.number != null ? String(it.number) : ''), // แสดง "ส-001/69"
        ocrNumber: it.number,
        name: it.name, unitCount: it.unit, category: it.category,
        subtags: (it.subtags || []).join(', '), owner_unit: unit,
        fileName: c.fileName, url: c.url,
        projRow: link.row || '', project: proj || link.project || '', status: link.status || '', budget: link.budget || '',
        summary: it.summary, rawText: it.rawText || ''
      });
    });
  });

  // เรียงตามเลขบัญชีคุม (รายการที่จับคู่ไม่ได้ไปอยู่ท้ายสุด)
  items.sort((a, b) => ((a.number || 1e9) - (b.number || 1e9)) || String(a.name).localeCompare(String(b.name), 'th'));

  const rows = items.map(it => ([
    it.code, it.ocrNumber, it.name, it.unitCount, it.category, it.subtags,
    it.owner_unit, it.fileName, it.url,
    it.projRow, it.project, it.status, it.budget,
    it.summary, it.rawText, new Date()
  ]));

  const headers = ['เลขอุปกรณ์', 'เลขที่ OCR อ่าน', 'ชื่ออุปกรณ์', 'หน่วยนับ', 'หมวดหมู่', 'หมวดย่อย/tag',
    'หน่วยเจ้าของ', 'ไฟล์ PDF', 'ลิงก์ PDF', 'แถวโครงการ', 'ชื่อโครงการ',
    'สถานะโครงการ', 'วงเงิน', 'สรุปคุณลักษณะ', 'ข้อความค้นหา', 'อัปเดตเมื่อ'];
  const sh = getOrCreateSheet_(CONFIG.INDEX_SHEET, headers);
  sh.clearContents();
  sh.showColumns(1, sh.getMaxColumns()); // เลิกซ่อนของเดิมก่อน (เลย์เอาต์คอลัมน์เปลี่ยน)
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.hideColumns(15); // ซ่อนคอลัมน์ "ข้อความค้นหา" (คอลัมน์ที่ 15)
}

/* ถอดช่วงหมายเลขจากชื่อไฟล์ เช่น "100-103.pdf" → [100,103], "7.pdf" → [7,7] */
function parseRange_(fileName) {
  const m = String(fileName || '').match(/^(\d+)(?:\s*-\s*(\d+))?\.pdf$/i);
  if (!m) return [null, null];
  const s = Number(m[1]);
  return [s, m[2] ? Number(m[2]) : s];
}

/* ---------- อ่านบัญชีคุม (ทะเบียนเลขอุปกรณ์ทางการ) + ฟังก์ชันจับคู่ตามชื่อ ---------- */
function readLedger_() {
  const sh = getSS_().getSheetByName(CONFIG.LEDGER_SHEET);
  if (!sh) throw new Error('ไม่พบชีต ' + CONFIG.LEDGER_SHEET);
  const values = sh.getDataRange().getValues();
  const display = sh.getDataRange().getDisplayValues(); // เอา "ส-001" ตามรูปแบบตัวเลขที่ตั้งไว้

  // หาแถวหัวตาราง (มีทั้ง "สิ่งอุปกรณ์" และ "หมายเลข")
  let hr = -1;
  for (let i = 0; i < values.length; i++) {
    const row = values[i].map(x => String(x).trim());
    if (row.indexOf('สิ่งอุปกรณ์') >= 0 && row.indexOf('หมายเลข') >= 0) { hr = i; break; }
  }
  if (hr < 0) throw new Error('ไม่พบหัวตาราง (สิ่งอุปกรณ์/หมายเลข) ในชีต ' + CONFIG.LEDGER_SHEET);

  const H = values[hr].map(x => String(x).trim());
  const col = name => { const e = H.findIndex(h => h === name); return e >= 0 ? e : H.findIndex(h => h.indexOf(name) >= 0); };
  const cName = col('สิ่งอุปกรณ์'), cNum = col('หมายเลข'), cProj = col('โครงการ'), cUnit = col('หน่วย'), cYear = col('ปีงบประมาณ');

  const rows = [];
  for (let i = hr + 1; i < values.length; i++) {
    const name = String(values[i][cName] || '').trim();
    const numRaw = values[i][cNum];
    if (!name && (numRaw === '' || numRaw == null)) continue;
    rows.push({
      number: Number(numRaw) || null,
      code: cNum >= 0 ? String(display[i][cNum] || '').trim() : String(Number(numRaw) || ''), // "ส-001"
      name: name,
      nkey: norm_(name),
      project: cProj >= 0 ? String(values[i][cProj] || '').trim() : '',
      unit: cUnit >= 0 ? String(values[i][cUnit] || '').trim() : '',
      year: cYear >= 0 ? String(values[i][cYear] || '').trim() : ''
    });
  }

  return {
    rows,
    /** จับคู่ชื่อ → แถวบัญชีคุม; ถ้ามี range=[a,b] จำกัดผู้สมัครไว้ในช่วงเลขนั้นก่อน (แม่นกว่า) */
    matchName(name, range) {
      const nk = norm_(name);
      if (!nk) return null;
      let pool = rows;
      if (range && range[0] != null) {
        const inRange = rows.filter(e => e.number >= range[0] && e.number <= range[1]);
        if (inRange.length) pool = inRange;
      }
      // 1) ชื่อตรงเป๊ะ (หลัง normalize)
      const exact = pool.filter(e => e.nkey === nk);
      if (exact.length) return exact[0];
      // 2) ความคล้ายสูงสุด (ไทยไม่เว้นวรรค → ใช้ substring + bigram similarity)
      let best = null, bestScore = 0;
      for (const e of pool) {
        const s = simNorm_(name, e.name);
        if (s > bestScore) { bestScore = s; best = e; }
      }
      // pool ถูกจำกัดด้วยช่วงเลขแล้วจึงยอมรับได้ตั้งแต่ 0.45; ถ้าเทียบทั้งบัญชีต้อง 0.6 ขึ้นไป
      const threshold = (pool !== rows) ? 0.45 : 0.6;
      return (best && bestScore >= threshold) ? best : null;
    }
  };
}

/* ความคล้ายชื่อ (0–1): เท่ากัน=1, เป็น substring กัน=0.9, ที่เหลือใช้ Dice บน bigram อักขระ */
function simNorm_(a, b) {
  a = norm_(a); b = norm_(b);
  if (!a || !b) return 0;
  if (a === b) return 1;
  if (a.indexOf(b) >= 0 || b.indexOf(a) >= 0) return 0.9;
  const bg = s => { const m = {}; for (let i = 0; i < s.length - 1; i++) { const g = s.substr(i, 2); m[g] = (m[g] || 0) + 1; } return m; };
  const A = bg(a), B = bg(b);
  let inter = 0, na = 0, nb = 0;
  for (const g in A) { na += A[g]; if (B[g]) inter += Math.min(A[g], B[g]); }
  for (const g in B) nb += B[g];
  return (na + nb) ? (2 * inter) / (na + nb) : 0;
}

/* ---------- อ่านชีตบัญชีคุม + ฟังก์ชันจับคู่โครงการ ---------- */
function readControlSheet_() {
  const sh = getSS_().getSheetByName(CONFIG.CONTROL_SHEET);
  if (!sh) throw new Error('ไม่พบชีต ' + CONFIG.CONTROL_SHEET);
  const values = sh.getDataRange().getValues();

  // หาแถวหัวตาราง (แถวที่มีทั้ง "หน่วย" และ "โครงการ")
  let hr = -1;
  for (let i = 0; i < values.length; i++) {
    const row = values[i].map(x => String(x).trim());
    if (row.indexOf('หน่วย') >= 0 && row.indexOf('โครงการ') >= 0) { hr = i; break; }
  }
  if (hr < 0) throw new Error('ไม่พบหัวตารางในชีต ' + CONFIG.CONTROL_SHEET);

  const H = values[hr].map(x => String(x).trim());
  // จับคู่แบบตรงเป๊ะก่อน แล้วค่อย fallback เป็น substring — กันเลือกคอลัมน์ผิด
  // (เช่น 'หน่วยนับ' มาก่อน 'หน่วย' หรือ 'เลขที่โครงการ' มาก่อน 'ชื่อโครงการ')
  const col = name => {
    const exact = H.findIndex(h => h === name);
    return exact >= 0 ? exact : H.findIndex(h => h.indexOf(name) >= 0);
  };
  const cUnit = col('หน่วย'), cProj = col('โครงการ'), cStatus = col('สถานะ'), cBudget = col('วงเงิน');

  const recs = [];
  for (let i = hr + 1; i < values.length; i++) {
    const r = values[i];
    const proj = String(r[cProj] || '').trim();
    if (!proj) continue;
    recs.push({
      order: recs.length + 1, // ลำดับตามแถวในชีตบัญชีคุม (บนลงล่าง) — ใช้ไล่หมายเลขอุปกรณ์
      rowNum: i + 1,
      unit: String(r[cUnit] || '').trim(),
      project: proj,
      status: cStatus >= 0 ? String(r[cStatus] || '').trim() : '',
      budget: cBudget >= 0 ? r[cBudget] : ''
    });
  }

  // จับคู่ชื่อโครงการ (+ โบนัสหน่วยตรง) กับแถวในบัญชีคุม
  const matchStr = (projectStr, unitStr) => {
    const q = String(projectStr || '').split(/\s+/).filter(x => x.length > 1);
    const qu = norm_(unitStr || '');
    let best = null, bestScore = 0;
    for (const rec of recs) {
      const projTokens = new Set(String(rec.project).split(/\s+/).filter(x => x.length > 1));
      let hit = 0;
      q.forEach(t => { if (projTokens.has(t)) hit++; });
      let s = q.length ? hit / q.length : 0;
      if (qu && norm_(rec.unit).indexOf(qu) >= 0) s += 0.3; // โบนัสถ้าหน่วยตรง
      if (s > bestScore) { bestScore = s; best = rec; }
    }
    if (best && bestScore >= 0.25) {
      return { order: best.order, row: best.rowNum, project: best.project, status: best.status, budget: best.budget, score: bestScore };
    }
    return {};
  };

  return {
    recs,
    matchProject: matchStr,                              // จับคู่ระดับไฟล์ (ส่งชื่อโครงการ+หน่วย)
    match(it) { return matchStr(it.project || it.name, it.owner_unit); } // เผื่อโค้ดเดิมเรียกใช้
  };
}
