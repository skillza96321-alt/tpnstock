/**
 * ระบบหลังบ้านสำหรับ "คลังสินค้า" — ทำให้ Google ชีตทำหน้าที่เป็นฐานข้อมูลกลาง
 *
 * วิธีติดตั้ง (ดูขั้นตอนละเอียดใน README.md):
 * 1. สร้าง Google ชีตใหม่ พร้อมแท็บ (sheet tabs) ชื่อ categories, products, transactions
 *    (แถวที่ 1 ของแต่ละแท็บจะใส่อะไรก็ได้หรือเว้นว่างไว้ก็ได้ ระบบไม่ได้อ่านชื่อหัวตาราง)
 * 2. เปิด Extensions > Apps Script วางโค้ดนี้ทับของเดิมทั้งหมด แล้วกด Save
 * 3. Deploy > New deployment > เลือกประเภท "Web app"
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. คัดลอกลิงก์ Web app URL ไปใส่ในไฟล์ index.html (ตัวแปร GAS_URL)
 * 5. กลับไปหน้าชีต รีเฟรชหน้าเว็บ แล้วกดเมนู "⚙️ คลังสินค้า > จัดรูปแบบชีตให้สวยงาม"
 *    เพื่อใส่หัวตารางภาษาไทยและจัดรูปแบบให้อัตโนมัติ
 *
 * ข้อควรระวัง: ระบบอ่าน/เขียนข้อมูลตาม "ลำดับคอลัมน์" ไม่ได้อ่านจากชื่อหัวตาราง
 * ห้ามสลับลำดับคอลัมน์ หรือแทรก/ลบคอลัมน์ในแท็บทั้ง 3 นี้ ไม่งั้นข้อมูลจะเพี้ยน
 */

// ลำดับคอลัมน์ของแต่ละตาราง (ห้ามสลับลำดับ)
const SHEET_SCHEMAS = {
  categories: ['id', 'name'],
  products: ['id', 'category_id', 'name', 'initial_stock', 'image'],
  transactions: ['id', 'product_id', 'type', 'qty', 'date', 'note']
};

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {};
  Object.keys(SHEET_SCHEMAS).forEach(name => {
    result[name] = sheetToObjects(ss.getSheetByName(name), SHEET_SCHEMAS[name]);
  });
  return jsonOutput(result);
}

function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput({ success: false, error: 'รูปแบบข้อมูลไม่ถูกต้อง' });
  }

  const fields = SHEET_SCHEMAS[payload.sheet];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(payload.sheet);
  if (!sheet || !fields) {
    return jsonOutput({ success: false, error: 'ไม่พบตาราง: ' + payload.sheet });
  }

  try {
    if (payload.action === 'insert') {
      const newId = getNextId(sheet);
      const row = fields.map(f => (f === 'id' ? newId : valueOrBlank(payload.data[f])));
      sheet.appendRow(row);
      return jsonOutput({ success: true, id: newId });
    }

    if (payload.action === 'update') {
      const rowNum = findRowNumberById(sheet, payload.id);
      if (rowNum === -1) return jsonOutput({ success: false, error: 'ไม่พบรายการที่จะแก้ไข' });
      fields.forEach((f, idx) => {
        if (f === 'id') return;
        if (payload.data[f] !== undefined) {
          sheet.getRange(rowNum, idx + 1).setValue(payload.data[f]);
        }
      });
      return jsonOutput({ success: true });
    }

    if (payload.action === 'delete') {
      const rowNum = findRowNumberById(sheet, payload.id);
      if (rowNum === -1) return jsonOutput({ success: false, error: 'ไม่พบรายการที่จะลบ' });
      sheet.deleteRow(rowNum);
      return jsonOutput({ success: true });
    }

    return jsonOutput({ success: false, error: 'ไม่รู้จักคำสั่ง: ' + payload.action });
  } catch (err) {
    return jsonOutput({ success: false, error: err.message });
  }
}

/* ---------- helpers ---------- */

// อ่านข้อมูล โดยข้ามแถวที่ 1 เสมอ (ถือเป็นแถวหัวตาราง/ป้ายกำกับ ไม่ว่าจะเขียนอะไรไว้ก็ตาม)
function sheetToObjects(sheet, fields) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const tz = sheet.getParent().getSpreadsheetTimeZone();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === '' || data[i][0] === null) continue; // ข้ามแถวว่าง
    const obj = {};
    fields.forEach((f, j) => {
      let v = data[i][j];
      if (f === 'date' && v instanceof Date) {
        v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      }
      obj[f] = v;
    });
    rows.push(obj);
  }
  return rows;
}

function getNextId(sheet) {
  const data = sheet.getDataRange().getValues();
  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const v = Number(data[i][0]);
    if (!isNaN(v) && v > maxId) maxId = v;
  }
  return maxId + 1;
}

function findRowNumberById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // แถวจริงในชีต (1-indexed)
  }
  return -1;
}

function valueOrBlank(v) {
  return v === undefined || v === null ? '' : v;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ==========================================================
   จัดรูปแบบชีตให้สวยงาม + ใส่หัวตารางเป็นภาษาไทย
   (ปลอดภัย รันซ้ำได้เรื่อยๆ ไม่กระทบข้อมูลที่มีอยู่)
   ========================================================== */

// เพิ่มเมนู "⚙️ คลังสินค้า" ในแถบเมนูของ Google ชีตอัตโนมัติเมื่อเปิดไฟล์
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ คลังสินค้า')
    .addItem('จัดรูปแบบชีตให้สวยงาม (ภาษาไทย)', 'formatAllSheets')
    .addSeparator()
    .addItem('เปิดใช้งานปิดรอบอัตโนมัติทุกสิ้นเดือน', 'installMonthlyResetTrigger')
    .addItem('ปิดรอบบัญชีเดือนนี้เลยตอนนี้ (มือ)', 'menuForceCloseNow')
    .addToUi();
}

function formatAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  formatSheet(
    ss.getSheetByName('categories'),
    ['รหัส', 'ชื่อหมวดหมู่ / แบรนด์'],
    [70, 260],
    ['ห้ามแก้ไข ระบบสร้างให้อัตโนมัติ (id)', 'ชื่อหมวดหมู่หรือแบรนด์สินค้า (name)']
  );

  formatSheet(
    ss.getSheetByName('products'),
    ['รหัส', 'รหัสหมวดหมู่', 'ชื่อสินค้า', 'จำนวนสต็อก', 'รูปภาพ'],
    [70, 110, 260, 120, 200],
    [
      'ห้ามแก้ไข ระบบสร้างให้อัตโนมัติ (id)',
      'อ้างอิงรหัสจากแท็บ categories (category_id)',
      'ชื่อสินค้า (name)',
      'ยอดสต็อกตั้งต้นก่อนมีรายการรับ-จ่าย ไม่ใช่ยอดคงเหลือปัจจุบัน (initial_stock) — ยอดคงเหลือจริงคำนวณจากยอดนี้บวกลบรายการในแท็บ transactions แล้วแสดงในหน้าแอป',
      'รูปสินค้า ระบบเก็บเป็นรหัส Base64 อัตโนมัติจากแอป ไม่ต้องแก้เอง (image)'
    ]
  );

  formatSheet(
    ss.getSheetByName('transactions'),
    ['รหัส', 'รหัสสินค้า', 'ประเภท', 'จำนวน', 'วันที่', 'หมายเหตุ'],
    [70, 110, 110, 90, 120, 260],
    [
      'ห้ามแก้ไข ระบบสร้างให้อัตโนมัติ (id)',
      'อ้างอิงรหัสจากแท็บ products (product_id)',
      'พิมพ์ได้แค่ in (รับเข้า) หรือ out (จ่ายออก) (type)',
      'จำนวน (qty)',
      'วันที่ทำรายการ (date)',
      'หมายเหตุ (note)'
    ]
  );

  // ทำ dropdown ให้เลือกได้แค่ in/out ในคอลัมน์ "ประเภท" ป้องกันพิมพ์ผิด
  const txSheet = ss.getSheetByName('transactions');
  if (txSheet) {
    const typeCol = SHEET_SCHEMAS.transactions.indexOf('type') + 1;
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['in', 'out'], true)
      .setAllowInvalid(false)
      .build();
    txSheet.getRange(2, typeCol, Math.max(txSheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  }

  SpreadsheetApp.getUi().alert('จัดรูปแบบชีตเป็นภาษาไทยเรียบร้อยแล้ว ✅');
}

function formatSheet(sheet, thaiLabels, colWidths, notes) {
  if (!sheet) return;
  const lastCol = thaiLabels.length;

  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setValues([thaiLabels]);
  headerRange
    .setBackground('#1F4B3F')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 34);

  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  notes.forEach((n, i) => headerRange.getCell(1, i + 1).setNote(n));

  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
    const bodyRange = sheet.getRange(2, 1, maxRows - 1, lastCol);
    try {
      bodyRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREEN, false, false);
    } catch (e) {
      // มีแถบสีอยู่แล้ว ข้ามไป
    }
  }
  sheet.setTabColor('#1F4B3F');
}

/* ==========================================================
   ปิดรอบบัญชีประจำเดือน
   - ยุบสต็อกปัจจุบันของทุกสินค้าให้เป็นยอดตั้งต้นใหม่ (initial_stock)
   - ย้ายรายการรับ-จ่ายทั้งหมดไปเก็บสำรองไว้ที่แท็บ transactions_archive (ไม่ได้ลบทิ้งถาวร)
   - ล้างแท็บ transactions ให้ว่าง พร้อมเริ่มรอบใหม่
   ========================================================== */

// ฟังก์ชันหลัก: ทำงานจริง ปลอดภัยสำหรับรันอัตโนมัติ (ห้ามเรียก UI ในนี้)
function closeMonthlyPeriod() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName('transactions');
  const prodSheet = ss.getSheetByName('products');
  if (!txSheet || !prodSheet) return;

  const txData = txSheet.getDataRange().getValues();
  const dataRows = txData.slice(1).filter(r => r[0] !== '' && r[0] !== null);
  if (dataRows.length === 0) return; // ไม่มีรายการให้ปิดรอบ

  // 1) คำนวณสต็อกปัจจุบันของทุกสินค้า แล้วยุบเข้าคอลัมน์ initial_stock (คอลัมน์ที่ 4)
  const prodData = prodSheet.getDataRange().getValues();
  const stockMap = {};
  for (let i = 1; i < prodData.length; i++) {
    const id = prodData[i][0];
    if (id === '' || id === null) continue;
    stockMap[id] = Number(prodData[i][3]) || 0;
  }
  dataRows.forEach(r => {
    const pid = r[1], type = r[2], qty = Number(r[3]) || 0;
    if (stockMap[pid] === undefined) stockMap[pid] = 0;
    stockMap[pid] += (type === 'in' ? qty : -qty);
  });
  for (let i = 1; i < prodData.length; i++) {
    const id = prodData[i][0];
    if (id === '' || id === null || stockMap[id] === undefined) continue;
    prodSheet.getRange(i + 1, 4).setValue(stockMap[id]);
  }

  // 2) ย้ายรายการทั้งหมดไปเก็บสำรองก่อนล้าง
  const tz = ss.getSpreadsheetTimeZone();
  const period = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  let archiveSheet = ss.getSheetByName('transactions_archive');
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet('transactions_archive');
    const header = SHEET_SCHEMAS.transactions.concat(['archived_period']);
    archiveSheet.appendRow(header);
    archiveSheet.getRange(1, 1, 1, header.length)
      .setBackground('#1F4B3F').setFontColor('#FFFFFF').setFontWeight('bold');
    archiveSheet.setFrozenRows(1);
    archiveSheet.setTabColor('#5B6B63');
  }
  const archiveRows = dataRows.map(r => r.concat([period]));
  archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, archiveRows.length, archiveRows[0].length)
    .setValues(archiveRows);

  // 3) ล้างข้อมูลในแท็บ transactions (เว้นหัวตารางไว้)
  const lastRow = txSheet.getLastRow();
  if (lastRow > 1) {
    txSheet.getRange(2, 1, lastRow - 1, txSheet.getLastColumn()).clearContent();
  }

  // 4) ส่งอีเมลสรุปผลให้เจ้าของชีต (ถ้าส่งไม่ได้ก็ข้ามไปเฉยๆ ไม่กระทบการปิดรอบ)
  try {
    const email = Session.getEffectiveUser().getEmail();
    if (email) {
      MailApp.sendEmail(email, 'ปิดรอบบัญชีคลังสินค้า ' + period,
        'ปิดรอบสำเร็จ: เก็บรายการรับ-จ่าย ' + dataRows.length + ' รายการไว้ในแท็บ transactions_archive แล้ว\n' +
        'ยอดสต็อกปัจจุบันของสินค้าทุกตัวถูกบันทึกเป็นยอดตั้งต้นใหม่ในแท็บ products เรียบร้อย\n' +
        'ไฟล์: ' + ss.getUrl());
    }
  } catch (e) {
    // ส่งอีเมลไม่ได้ ไม่เป็นไร
  }
}

// เรียกจากเมนู: ปิดรอบทันทีด้วยมือ พร้อมถามยืนยันก่อน
function menuForceCloseNow() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    'ปิดรอบบัญชีตอนนี้เลยหรือไม่?',
    'ระบบจะยุบยอดสต็อกปัจจุบันเป็นยอดตั้งต้นใหม่ แล้วย้ายรายการรับ-จ่ายทั้งหมดไปเก็บที่แท็บ transactions_archive ' +
    'จากนั้นล้างแท็บ transactions ให้ว่าง (ข้อมูลไม่ได้หายไปไหน ยังดูย้อนหลังได้ที่ transactions_archive)',
    ui.ButtonSet.YES_NO
  );
  if (res === ui.Button.YES) {
    closeMonthlyPeriod();
    ui.alert('ปิดรอบบัญชีเรียบร้อยแล้ว ✅');
  }
}

// เรียกจากเมนู: ติดตั้งระบบตรวจสอบอัตโนมัติทุกคืน (ทำครั้งเดียว)
function installMonthlyResetTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkAndCloseMonthEnd') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkAndCloseMonthEnd')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();
  SpreadsheetApp.getUi().alert(
    'ตั้งค่าระบบปิดรอบอัตโนมัติเรียบร้อยแล้ว ✅\n' +
    'ระบบจะตรวจสอบทุกคืนช่วงประมาณ 23:00 น. และปิดรอบให้อัตโนมัติเฉพาะคืนสุดท้ายของแต่ละเดือนเท่านั้น (ทำทุกเดือนตลอด 12 เดือน)'
  );
}

// ฟังก์ชันที่ trigger เรียกทุกคืน: เช็กว่าเป็นวันสุดท้ายของเดือนหรือไม่ ถ้าใช่ค่อยปิดรอบ
function checkAndCloseMonthEnd() {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const today = new Date(todayStr + 'T00:00:00');
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  if (tomorrow.getMonth() !== today.getMonth()) {
    closeMonthlyPeriod();
  }
}
