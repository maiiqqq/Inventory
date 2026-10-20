/**********************************************************************
 * Config.gs — ค่าตั้งต้นของระบบ
 * แก้ FOLDER_ID / ชื่อชีต ได้ที่นี่ ส่วน API key เก็บใน Script Properties
 **********************************************************************/

const CONFIG = {
  // ID ของ Google Sheet (จำเป็นสำหรับสคริปต์แบบ standalone — อ้างชีตด้วย ID)
  SPREADSHEET_ID: '1thRDzg3m5j7yGta0HLzKdGm-Gm19hpKsgG1RjV4t1Oc',

  // โฟลเดอร์ Drive ที่เก็บ PDF (มีโฟลเดอร์ย่อยรายปี เช่น 69, 70 ได้)
  FOLDER_ID: '1KlzsV1mEEvit9JxjvYMt69-bSAJ8X929',

  // ชื่อชีต
  LEDGER_SHEET:  'บัญชีคุม',       // ทะเบียนเลขอุปกรณ์ทางการ (คอลัมน์: สิ่งอุปกรณ์ | หมายเลข | ปีงบประมาณ | โครงการ | หน่วย) — แหล่งความจริงของ "หมายเลขอุปกรณ์"
  CONTROL_SHEET: 'สถานะโครงการ',   // ชีตติดตามสถานะโครงการ 8 ขั้น + วงเงิน (หัวตารางไม่อยู่แถวแรก) — อ่านอย่างเดียว
  INDEX_SHEET:   'Index_อุปกรณ์',  // ชีตดัชนีกลาง (สร้างให้อัตโนมัติ)
  CACHE_SHEET:   '_Cache',         // เก็บผล Gemini รายไฟล์ (กันเรียกซ้ำ)

  // OpenRouter (OpenAI-compatible API)
  OPENROUTER_URL: 'https://openrouter.ai/api/v1/chat/completions',
  // ต้องเป็นโมเดลที่อ่าน PDF สแกน (OCR ภาพ) ได้ — Gemini/Claude/GPT-4o ผ่าน OpenRouter ทำได้
  OPENROUTER_MODEL: 'openrouter/auto-beta',
  // engine อ่าน PDF ของ OpenRouter:
  //   'native'     = ให้โมเดลอ่านไฟล์เอง (Gemini/Claude OCR ในตัว) — เหมาะกับเอกสารสแกน ไม่มีค่าใช้จ่ายเพิ่ม
  //   'mistral-ocr'= OCR คุณภาพสูง (คิดเงิน ~$2/1000 หน้า)
  //   'pdf-text'   = ดึงเฉพาะ text-layer (ฟรี แต่ใช้ไม่ได้กับเอกสารสแกน)
  PDF_ENGINE: 'native',

  // กันชนก่อนชน limit 6 นาทีของ Apps Script (หยุดแล้วให้กดทำต่อ)
  MAX_RUNTIME_MS: 4.5 * 60 * 1000,

  // หมวดหมู่ "แนะนำ" (เพื่อความสม่ำเสมอ) — AI วิเคราะห์เองได้ และสร้างหมวดใหม่ที่เฉพาะเจาะจงได้
  // ห้ามใช้ "อื่น ๆ" เด็ดขาด (ดู prompt ใน OpenRouter.gs)
  CATEGORIES: [
    'โดรน/ต่อต้านอากาศยานไร้คนขับ',
    'คอมพิวเตอร์/เซิร์ฟเวอร์/คลาวด์',
    'กล้อง/CCTV/บันทึกภาพ',
    'สื่อสาร/วิทยุ/ดาวเทียม',
    'ควบคุมการเข้าออก/บัตร/RFID',
    'ความมั่นคงปลอดภัยไซเบอร์',
    'ปรับปรุงอาคาร/สถานที่',
    'ยานพาหนะ/รถปฏิบัติการ',
    'เครื่องมือซ่อมบำรุง/ชิ้นส่วนอะไหล่'
  ]
};

/** ดึง OpenRouter API key จาก Script Properties */
function getApiKey_() {
  const k = PropertiesService.getScriptProperties().getProperty('OPENROUTER_API_KEY');
  if (!k) throw new Error('ยังไม่ได้ตั้งค่า OpenRouter API Key — รันฟังก์ชัน saveApiKey("คีย์ของคุณ") หนึ่งครั้ง หรือใส่ใน Project Settings > Script Properties (ชื่อ OPENROUTER_API_KEY)');
  return k;
}
