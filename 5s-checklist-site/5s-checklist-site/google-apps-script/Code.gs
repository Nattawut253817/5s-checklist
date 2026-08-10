/**
 * 5ส Checklist - Backend (Google Apps Script)
 * ทำหน้าที่เป็นฐานข้อมูลกลาง เก็บสถานะการตรวจเช็กแต่ละข้อไว้ใน Google Sheet
 * เพื่อให้ทุกคนที่เปิดหน้าเว็บ checklist เห็นสถานะเดียวกัน (เกือบเรียลไทม์)
 *
 * วิธีใช้งาน:
 * 1. สร้าง Google Sheet ใหม่ (ไฟล์ว่างๆ 1 ไฟล์) ตั้งชื่อตามสะดวก เช่น "5ส Checklist ฐานข้อมูล"
 * 2. ไปที่เมนู Extensions > Apps Script
 * 3. ลบโค้ดตัวอย่างเดิมทั้งหมดในไฟล์ Code.gs แล้ววางโค้ดทั้งหมดในไฟล์นี้ทับ
 * 4. กด Run เลือกฟังก์ชัน setupSheet แล้วกด Run อีกครั้ง (ครั้งแรกจะให้ขออนุญาต ให้กด Allow)
 *    ขั้นตอนนี้จะสร้างชีตชื่อ "Status" พร้อมหัวตารางให้อัตโนมัติ
 * 5. ไปที่ Deploy > New deployment
 *      - เลือกรูปแบบ (Select type) เป็น "Web app"
 *      - Execute as: Me
 *      - Who has access: "Anyone" (ถ้าต้องการให้คนนอกองค์กรเข้าถึงได้ด้วยลิงก์)
 *        หรือ "Anyone within [ชื่อองค์กร]" ถ้าใช้ Google Workspace ขององค์กร (ปลอดภัยกว่า แนะนำถ้ามี)
 *      - กด Deploy แล้วกด Authorize access ตามที่ระบบขอ
 * 6. คัดลอก "Web app URL" ที่ได้ (จะลงท้ายด้วย /exec)
 * 7. เปิดไฟล์ index.html แล้วนำ URL นี้ไปแทนที่ค่า PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE
 *    ในบรรทัด: const API_URL = "...";
 *
 * หมายเหตุ: ถ้าแก้ไขโค้ดไฟล์นี้ภายหลัง ต้องไปที่ Deploy > Manage deployments > กดไอคอนดินสอ (Edit)
 * แล้วเปลี่ยน Version เป็น "New version" แล้วกด Deploy ใหม่อีกครั้ง ไม่งั้นระบบจะยังใช้โค้ดเวอร์ชันเก่าอยู่
 */

const SHEET_NAME = "Status";

// เรียกครั้งเดียวตอนตั้งค่าระบบ เพื่อสร้างชีตและหัวตารางให้พร้อมใช้งาน
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getRange(1, 1).getValue() !== "ItemID") {
    sheet.getRange(1, 1, 1, 4).setValues([["ItemID", "Status", "Note", "UpdatedAt"]]);
    sheet.setFrozenRows(1);
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName(SHEET_NAME);
  }
  return sheet;
}

// ให้ index.html เรียกตอนโหลดหน้า/รีเฟรช เพื่อดึงสถานะล่าสุดของทุกข้อ
function doGet(e) {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const state = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    state[row[0]] = { status: row[1] || "none", note: row[2] || "" };
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, state: state }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ให้ index.html เรียกทุกครั้งที่มีคนกดเปลี่ยนสถานะหรือพิมพ์หมายเหตุ
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const body = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        rowIndex = i + 1; // แปลงเป็นเลขแถวจริงใน sheet (index เริ่มที่ 1)
        break;
      }
    }
    const now = new Date();
    if (rowIndex === -1) {
      sheet.appendRow([body.id, body.status || "none", body.note || "", now]);
    } else {
      sheet.getRange(rowIndex, 2, 1, 3).setValues([[body.status || "none", body.note || "", now]]);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
