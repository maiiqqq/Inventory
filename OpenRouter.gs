/**********************************************************************
 * OpenRouter.gs — ส่ง PDF ให้โมเดลผ่าน OpenRouter (OpenAI-compatible)
 * คืน JSON รายการอุปกรณ์ + หมวดหมู่
 * ใช้ content type "file" (base64 data URL) — โมเดลต้องอ่าน PDF ได้ เช่น Gemini/Claude/GPT-4o
 **********************************************************************/

function openrouterExtractFromPdf_(f) {
  const blob = DriveApp.getFileById(f.id).getBlob();
  const b64 = Utilities.base64Encode(blob.getBytes());

  const payload = {
    model: CONFIG.OPENROUTER_MODEL,
    temperature: 0,
    response_format: { type: 'json_object' },
    plugins: [{ id: 'file-parser', pdf: { engine: CONFIG.PDF_ENGINE } }],
    messages: [{
      role: 'user',
      content: [
        { type: 'text', text: buildExtractPrompt_(f) },
        { type: 'file', file: { filename: f.name, file_data: 'data:application/pdf;base64,' + b64 } }
      ]
    }]
  };

  const res = UrlFetchApp.fetch(CONFIG.OPENROUTER_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + getApiKey_(),
      'HTTP-Referer': 'https://script.google.com',
      'X-Title': 'BanchiKhum Search'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code !== 200) throw new Error('OpenRouter HTTP ' + code + ': ' + res.getContentText().slice(0, 500));

  const j = JSON.parse(res.getContentText());
  if (j.error) throw new Error('OpenRouter error: ' + JSON.stringify(j.error).slice(0, 400));
  const text = j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content;
  if (!text) throw new Error('ไม่มีผลลัพธ์จากโมเดล: ' + res.getContentText().slice(0, 400));

  let parsed;
  try { parsed = JSON.parse(stripFences_(text)); }
  catch (e) { throw new Error('แปลง JSON ไม่ได้: ' + String(text).slice(0, 400)); }

  const items = parsed.items || (Array.isArray(parsed) ? parsed : []);
  return items.map(x => {
    const it = {
      number: Number(x.number) || null,
      name: String(x.name || '').trim(),
      unit: String(x.unit || '').trim(),
      owner_unit: String(x.owner_unit || '').trim(),
      project: String(x.project || '').trim(),
      subtags: Array.isArray(x.subtags) ? x.subtags : [],
      summary: String(x.summary || '').trim()
    };
    it.category = normalizeCategory_(x.category, it); // AI กำหนด, กัน "อื่น ๆ"
    it.rawText = [it.name, it.summary, it.subtags.join(' ')].join(' ');
    return it;
  });
}

/** ตัด ```json ... ``` ถ้าโมเดลใส่มา */
function stripFences_(t) {
  return String(t).replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/i, '').trim();
}

/** ใช้หมวดที่ AI วิเคราะห์เป็นหลัก — ห้าม "อื่น ๆ"/ว่าง (ถ้าเจอ ใช้ safety net เดาจากคำสำคัญ) */
function normalizeCategory_(c, item) {
  c = String(c || '').trim();
  const bad = /^(อื่น\s*ๆ|อื่นๆ|other|others|misc|ทั่วไป|n\/?a|ไม่ระบุ|ไม่ทราบ|-)$/i;
  if (c && !bad.test(c)) return c;
  return guessCategory_(item);
}

/** safety net: เดาหมวดจากคำสำคัญในชื่อ/สรุป/tag (ไม่คืน "อื่น ๆ") */
function guessCategory_(item) {
  const t = [item && item.name, item && item.summary, (item && item.subtags || []).join(' ')]
    .join(' ').toLowerCase();
  const rules = [
    [/(drone|โดรน|อากาศยานไร้คนขับ|ต่อต้านอากาศยาน|uav|จับสัญญาณ.*อากาศยาน)/, 'โดรน/ต่อต้านอากาศยานไร้คนขับ'],
    [/(server|เซิร์ฟเวอร์|แม่ข่าย|คอมพิวเตอร์|คลาวด์|cloud|ประมวลผล|data\s*center)/, 'คอมพิวเตอร์/เซิร์ฟเวอร์/คลาวด์'],
    [/(cctv|กล้อง|วงจรปิด|บันทึกภาพ)/, 'กล้อง/CCTV/บันทึกภาพ'],
    [/(วิทยุ|ดาวเทียม|satellite|ku-band|manpack|สื่อสาร|เชื่อมโยงสัญญาณ|iridium|เครือข่าย)/, 'สื่อสาร/วิทยุ/ดาวเทียม'],
    [/(บัตร|rfid|เข้าออก|เครื่องอ่านบัตร|access\s*control|ประตู)/, 'ควบคุมการเข้าออก/บัตร/RFID'],
    [/(ไซเบอร์|cyber|security|รักษาความมั่นคงปลอดภัย)/, 'ความมั่นคงปลอดภัยไซเบอร์'],
    [/(ปรับปรุง|ห้องประชุม|อาคาร|สถานที่|ดับเพลิง|wiring|piping)/, 'ปรับปรุงอาคาร/สถานที่'],
    [/(รถ|ยานพาหนะ|vehicle|ปฏิบัติการเคลื่อนที่)/, 'ยานพาหนะ/รถปฏิบัติการ'],
    [/(ชิ้นส่วน|อะไหล่|ซ่อมบำรุง|เครื่องมือ)/, 'เครื่องมือซ่อมบำรุง/ชิ้นส่วนอะไหล่']
  ];
  for (let i = 0; i < rules.length; i++) if (rules[i][0].test(t)) return rules[i][1];
  // สุดท้ายจริง ๆ: ใช้ชื่ออุปกรณ์แบบย่อเป็นหมวด (ยังคงเฉพาะเจาะจง ไม่ใช่ "อื่น ๆ")
  const nm = String(item && item.name || '').trim();
  return nm ? nm.split(/\s+/).slice(0, 4).join(' ') : 'ไม่ระบุชื่ออุปกรณ์';
}

function buildExtractPrompt_(f) {
  const range = (f.start != null)
    ? `ไฟล์นี้ชื่อ "${f.name}" ครอบคลุมหมายเลขสิ่งอุปกรณ์ (ลำดับ .../ปีงบ) ตั้งแต่ ${f.start} ถึง ${f.end}. ให้ใช้เลขในเอกสารเป็นหลัก แต่ควรอยู่ในช่วงนี้`
    : `ไฟล์นี้ชื่อ "${f.name}"`;

  return [
    'คุณคือผู้ช่วยจัดทำบัญชีคุมสิ่งอุปกรณ์ของกองทัพ อ่านเอกสาร PDF (บันทึกข้อความขออนุมัติ + คุณลักษณะเฉพาะสิ่งอุปกรณ์) ที่แนบมา',
    range,
    'ในเอกสารมีสิ่งอุปกรณ์หลายรายการ แต่ละรายการระบุด้วย "หมายเลขสิ่งอุปกรณ์" และ "ลำดับ n / ปีงบ" (เช่น 2/69 = หมายเลข 2), "ชื่อสิ่งอุปกรณ์", "หน่วยนับ", และคุณลักษณะเฉพาะ',
    'จงสกัดข้อมูลทุกรายการออกมาเป็น JSON ตามรูปแบบนี้เท่านั้น:',
    '{"items":[{"number":<เลขจำนวนเต็ม>,"name":"<ชื่อสิ่งอุปกรณ์>","unit":"<หน่วยนับ>","owner_unit":"<หน่วยเจ้าของ เช่น ศตก. / ศทส.ฯ>","project":"<ชื่องาน/โครงการจัดหา จากบันทึกข้อความ>","category":"<หมวดหมู่>","subtags":["<คำสำคัญ>"],"summary":"<สรุปคุณลักษณะเด่น 1-2 ประโยค>"}]}',
    'การกำหนด category: ให้ "วิเคราะห์หน้าที่/ประเภทการใช้งานของอุปกรณ์" แล้วตั้งชื่อหมวดหมู่ภาษาไทยที่กระชับและเฉพาะเจาะจง',
    'ควรใช้หมวดที่แนะนำต่อไปนี้ก่อนเพื่อความสม่ำเสมอ (ใช้ข้อความให้ตรงเป๊ะ): ' + CONFIG.CATEGORIES.map(c => '"' + c + '"').join(', '),
    'หากไม่มีหมวดใดเหมาะสมจริง ๆ ให้สร้างหมวดใหม่ที่เฉพาะเจาะจงตามหน้าที่ของอุปกรณ์ และควรใช้ชื่อหมวดเดียวกันสำหรับอุปกรณ์ประเภทเดียวกัน',
    'ห้ามเด็ดขาด: category ต้องไม่เป็น "อื่น ๆ", "อื่นๆ", "ทั่วไป", "ไม่ระบุ", "other" หรือค่าว่าง — ต้องเป็นหมวดที่มีความหมายเสมอ',
    'ข้อกำหนดอื่น: number เป็นตัวเลขล้วน (ไม่มี /ปีงบ). ถ้าอ่านไม่ชัดให้เดาอย่างสมเหตุผลตามลำดับในเอกสาร. ตอบเป็น JSON ล้วน ไม่มีคำอธิบายอื่น ไม่มี markdown.'
  ].join('\n');
}

/* ---------- ทดสอบ key อย่างเร็ว (รันจาก editor ได้) ---------- */
function testOpenRouterKey() {
  const res = UrlFetchApp.fetch(CONFIG.OPENROUTER_URL, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + getApiKey_() },
    payload: JSON.stringify({
      model: CONFIG.OPENROUTER_MODEL,
      messages: [{ role: 'user', content: 'ตอบว่า OK' }]
    })
  });
  Logger.log(res.getResponseCode() + ' ' + res.getContentText().slice(0, 300));
}
