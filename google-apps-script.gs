/**
 * Google Apps Script backend สำหรับ "ตารางลาหยุด"
 * ----------------------------------------------------
 * วิธีติดตั้ง:
 * 1. สร้าง Google Sheet ใหม่ (หรือใช้ไฟล์เดิม)
 * 2. เพิ่มแท็บชื่อ "LeaveData" แล้วใส่หัวตารางแถวแรก:
 *    A1: Date   B1: Name   C1: Status   D1: Timestamp
 * 3. เมนู Extensions > Apps Script วางโค้ดนี้ทั้งหมดแทนของเดิม แล้วกด Save
 * 4. กด Deploy > New deployment
 *    - Select type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    - กด Deploy แล้วอนุญาต (Authorize) ตามที่ระบบขอ
 * 5. คัดลอก "Web app URL" ที่ได้ ไปวางแทนที่
 *    APPS_SCRIPT_URL ในไฟล์ leave-sheet_2.html
 */

const SHEET_NAME = "LeaveData";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Date", "Name", "Status", "Timestamp"]);
  }
  return sheet;
}

function formatDate_(d) {
  if (Object.prototype.toString.call(d) === "[object Date]") {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }
  return String(d);
}

// อ่านข้อมูลทั้งหมด -> คืนเป็น JSON { rows: [ {date, name, status, ts}, ... ] }
function doGet(e) {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const [date, name, status, ts] = data[i];
    if (date && name && status) {
      rows.push({
        date: formatDate_(date),
        name: String(name),
        status: String(status),
        ts: ts ? String(ts) : null
      });
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ rows: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}

// รับข้อมูลทั้งหมดจากหน้าเว็บ (rows ปัจจุบันทุกแถวที่มีสถานะ) แล้วเขียนทับทั้งชีต
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const rows = body.rows || [];

    const sheet = getSheet_();
    sheet.clearContents();
    sheet.appendRow(["Date", "Name", "Status", "Timestamp"]);

    if (rows.length > 0) {
      const values = rows.map(r => [r.date, r.name, r.status, r.ts || ""]);
      sheet.getRange(2, 1, values.length, 4).setValues(values);
      // บังคับคอลัมน์ Date ให้เป็นข้อความล้วน กัน Sheets แปลงเป็นวันที่เอง
      sheet.getRange(2, 1, values.length, 1).setNumberFormat("@");
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
