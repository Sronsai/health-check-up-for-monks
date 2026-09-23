// วางโค้ดนี้ใน Google Sheet → ส่วนขยาย → Apps Script
// จากนั้น ทำให้ใช้งานได้ → การทำให้ใช้งานได้รายการใหม่ → เว็บแอป
//   ดำเนินการในฐานะ: ฉัน | ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน
const SHEET_ID = '1iyxRCc532s3NlFGmMoEkWqnautD9RrXQz4nvPUEY9pc';
const TOKEN = 'pakchom2569'; // เปลี่ยน และใส่ค่าเดียวกันในหน้า "ตั้งค่า" ของแอป
const HEADERS = ['Timestamp', 'เลขคิว', 'เลขประจำตัวประชาชน', 'ชื่อ-นามสกุล', 'เพศ', 'วัน/เดือน/ปี เกิด', 'สัญชาติ', 'หมู่เลือด', 'ข้อมูลที่อยู่(ปัจจุบัน)', 'จังหวัด', 'อำเภอ', 'ตำบล', 'หมายเลขโทรศัพท์'];

function sheet_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  return sh;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.token !== TOKEN) return json_({ ok: false, error: 'invalid token' });
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh = sheet_();
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 3, sh.getLastRow() - 1, 1).getDisplayValues().flat() : [];
    if (ids.indexOf(String(d.idCard)) !== -1) return json_({ ok: false, error: 'duplicate' });
    sh.appendRow([new Date(d.timestamp || Date.now()), d.queue, "'" + d.idCard, d.fullName, d.gender, "'" + d.dob, d.nationality, d.blood, d.address, d.province, d.amphoe, d.tambon, "'" + d.phone]);
    return json_({ ok: true });
  } finally { lock.releaseLock(); }
}

function doGet(e) {
  if ((e.parameter.token || '') !== TOKEN) return json_({ ok: false, error: 'invalid token' });
  const sh = sheet_(); const v = sh.getDataRange().getValues();
  const rows = v.slice(1).filter(r => r[2]).map(r => ({
    timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), queue: r[1], idCard: String(r[2]), fullName: r[3], gender: r[4],
    dob: r[5] instanceof Date ? Utilities.formatDate(r[5], 'Asia/Bangkok', 'dd/MM/yyyy') : String(r[5]),
    nationality: r[6], blood: r[7], address: r[8], province: r[9], amphoe: r[10], tambon: r[11], phone: String(r[12])
  }));
  return json_({ ok: true, rows });
}
