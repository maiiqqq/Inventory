# Inventory

เว็บค้นหาบัญชีคุมสิ่งอุปกรณ์ที่ทำงานบน Google Apps Script (V8) โดยอ่าน PDF จาก
Google Drive, สกัดข้อมูลผ่าน OpenRouter และจับคู่เลขอุปกรณ์กับชีต `บัญชีคุม`
ก่อนเขียนผลลงชีต `Index_อุปกรณ์`

## พัฒนาและทดสอบ

ไม่ต้องติดตั้งแพ็กเกจ ใช้ Node.js รันชุดทดสอบใน checkout:

```sh
node --test tests/index.test.cjs
```

ชุดทดสอบใช้ข้อมูลจำลองในหน่วยความจำ ไม่เรียก Google Sheets, Drive หรือ OpenRouter
ไฟล์ `WebApp.html` เป็น UI ของ Apps Script และไม่ใช่เว็บเซิร์ฟเวอร์ที่รันในเครื่อง

## อัปเดตระบบจริง

หากทำงานบน Windows ให้ใช้ **PowerShell** สำหรับ `clasp` และเข้าสู่ระบบ Google
ด้วยบัญชีที่แก้ไขโปรเจกต์ Apps Script ได้:

```powershell
npm install --global @google/clasp
clasp login
```

ก่อนส่งโค้ดขึ้นระบบจริง ให้ `clasp clone` โปรเจกต์ที่ใช้งานอยู่ลงในโฟลเดอร์ใหม่
เพื่อสำรองและตรวจไฟล์ฝั่ง Google ก่อนเสมอ:

```powershell
$scriptId = '1RvINzBsOibTkW9hrRpEPvcbDKQz4Gg5zUi10hfmO4Lkbi_XdXV3uZOEe'
$remoteDir = Join-Path $env:TEMP ('inventory-apps-script-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $remoteDir | Out-Null
Set-Location $remoteDir
clasp clone $scriptId
clasp status
```

ตรวจความต่างกับไฟล์ใน repository ก่อนคัดลอก `Code.gs`, `Config.gs`,
`Index.gs`, `OpenRouter.gs`, `WebApp.html` และ `appsscript.json` ไปยังโฟลเดอร์
ที่ clone แล้วสั่ง `clasp push` จากโฟลเดอร์นั้น การ clone ไว้ก่อนช่วยเก็บไฟล์
อื่นที่มีอยู่ในโปรเจกต์จริง หลัง push ให้ทำตามขั้นตอนด้านล่างเพื่อสร้างดัชนี
และอัปเดต deployment

1. นำไฟล์ `.gs`, `WebApp.html` และ `appsscript.json` ไปยังโปรเจกต์ Apps Script
   ที่ใช้งานอยู่ โปรเจกต์นี้ยังไม่มี `.clasp.json` สำหรับซิงก์อัตโนมัติ
2. ตรวจ `SPREADSHEET_ID` และ `FOLDER_ID` ใน `Config.gs` และตั้ง
   `OPENROUTER_API_KEY` ใน **Script Properties** ของ Apps Script อย่าใส่คีย์ใน Git
3. รัน `buildIndex()` ใน Apps Script เพื่อสร้างดัชนีใหม่จาก cache ปัจจุบัน
   ระบบเรียก OpenRouter เฉพาะ PDF ใหม่หรือ PDF ที่เปลี่ยนแปลง
4. เปิดหน้าเว็บแอปและตรวจผลค้นหา เลขบนการ์ดอ้างอิงจากชีต `บัญชีคุม` เท่านั้น
   หากจับคู่ไม่ได้หรือผลใกล้เคียงกัน ระบบจะแสดง `รอตรวจสอบเลข` พร้อมเลข OCR
   เพื่อให้ตรวจรายการจริงในบัญชีคุม
5. หากใช้ deployment แบบกำหนดเวอร์ชัน ให้อัปเดต deployment หลังแก้ UI

การรัน `buildIndex()` เขียนชีต `Index_อุปกรณ์` และลบ cache ของ PDF ที่ไม่อยู่ใน
โฟลเดอร์แล้ว ควรตรวจปลายทางที่กำหนดใน `Config.gs` ก่อนรันกับข้อมูลจริง
