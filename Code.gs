/**
 * ระบบติดตามวันหยุดเสาร์ — อิงข้อมูลจากชีท "SaturdayTracker"
 * คอลัมน์: A = วันที่ (Date) | B = สถานะ ("หยุด" / "ทำงาน") | C = สถานะ sync ปฏิทิน ("DONE" / ว่าง)
 *
 * รอบนับโควตาเริ่มวันที่ CUTOFF_DAY ของทุกเดือน ถึงวันที่ (CUTOFF_DAY - 1) ของเดือนถัดไป
 * ปรับ CUTOFF_DAY / CAP ด้านล่างได้ตามต้องการ
 */
var CUTOFF_DAY = 16; // วันเริ่มรอบใหม่ของทุกเดือน
var CAP = 2;          // สิทธิ์วันหยุดเสาร์ต่อรอบ
var THAI_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : null;

  // ไม่มี action -> เปิดตรงจาก URL ของ Apps Script ได้เหมือนเดิม (เผื่อไว้ใช้ทดสอบ)
  if (!action) {
    return HtmlService.createHtmlOutputFromFile('Index')
        .setTitle('ระบบติดตามวันหยุดเสาร์')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // มี action -> ทำงานเป็น JSON API ให้หน้าเว็บภายนอก (เช่น GitHub Pages) เรียกผ่าน fetch()
  var result;
  try {
    if (action === "getData") {
      result = getSaturdaysData();
    } else if (action === "updateStatus") {
      result = updateStatusAndSync(parseInt(e.parameter.rowIndex, 10), e.parameter.status);
    } else if (action === "generate") {
      result = generateSaturdays(parseInt(e.parameter.months, 10) || 6);
    } else {
      result = { error: "ไม่รู้จัก action: " + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
}

function thaiShort_(date) {
  return date.getDate() + " " + THAI_MONTHS[date.getMonth()] + " " + String(date.getFullYear() + 543).slice(-2);
}

// คำนวณรอบ (cycle) ที่วันที่นี้สังกัดอยู่ ตาม CUTOFF_DAY
function getCycle_(date) {
  var y = date.getFullYear(), m = date.getMonth(), d = date.getDate();
  var start = d >= CUTOFF_DAY ? new Date(y, m, CUTOFF_DAY) : new Date(y, m - 1, CUTOFF_DAY);
  var end = new Date(start.getFullYear(), start.getMonth() + 1, CUTOFF_DAY - 1);
  var key = start.getFullYear() + "-" + ("0" + (start.getMonth() + 1)).slice(-2) + "-" + ("0" + start.getDate()).slice(-2);
  return { start: start, end: end, key: key, label: thaiShort_(start) + " – " + thaiShort_(end) };
}

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SaturdayTracker");
}

/**
 * อ่านข้อมูลทั้งหมด จัดกลุ่มตามรอบ 16→15 พร้อมนับโควตาแยกรายรอบ
 */
function getSaturdaysData() {
  var sheet = getSheet_();
  if (!sheet) return { groups: [], cap: CAP, error: "ไม่พบชีท SaturdayTracker" };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { groups: [], cap: CAP };

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var today = new Date();
  var todayKey = getCycle_(today).key;

  var groupsMap = {};
  var order = [];

  for (var i = 0; i < data.length; i++) {
    var dateVal = data[i][0];
    var status = data[i][1];
    var syncStatus = data[i][2];
    if (!(dateVal instanceof Date)) continue; // ข้ามแถวที่วันที่ไม่ถูกต้อง

    var cyc = getCycle_(dateVal);
    if (!groupsMap[cyc.key]) {
      groupsMap[cyc.key] = {
        key: cyc.key,
        label: cyc.label,
        isCurrent: cyc.key === todayKey,
        offCount: 0,
        rows: []
      };
      order.push(cyc.key);
    }
    var g = groupsMap[cyc.key];
    if (status === "หยุด") g.offCount++;
    g.rows.push({
      rowIndex: i + 2,
      dateLabel: thaiShort_(dateVal),
      weekday: ["อา", "จ", "อ", "พ", "พฤ", "ศ", "ส"][dateVal.getDay()],
      status: status || "",
      syncStatus: syncStatus || ""
    });
  }

  order.sort();
  var groups = order.map(function (k) {
    var g = groupsMap[k];
    g.atCap = g.offCount >= CAP;
    return g;
  });

  return { groups: groups, cap: CAP };
}

/**
 * สร้างแถววันเสาร์ล่วงหน้าให้อัตโนมัติ ไม่จำกัดจำนวนครั้ง — กดซ้ำเพื่อขยายต่อไปเรื่อยๆ ได้
 * เริ่มสแกนจากต้นรอบปัจจุบันเสมอ (กันวันตกหล่น) และขยายไปอีก monthsAhead เดือน
 * นับจากวันที่ล่าสุดที่มีอยู่แล้วในชีท (ถ้ามี) หรือจากวันนี้ (ถ้าชีทว่าง/ยังไปไม่ถึงวันนี้)
 */
function generateSaturdays(monthsAhead) {
  var sheet = getSheet_();
  if (!sheet) return { error: "ไม่พบชีท SaturdayTracker" };
  monthsAhead = monthsAhead || 6;
  var tz = Session.getScriptTimeZone();

  var existing = {};
  var maxDate = null;
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var existingData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    existingData.forEach(function (r) {
      if (r[0] instanceof Date) {
        var k = Utilities.formatDate(r[0], tz, "yyyy-MM-dd");
        existing[k] = true;
        if (!maxDate || r[0] > maxDate) maxDate = r[0];
      }
    });
  }

  var today = new Date();
  var start = getCycle_(today).start; // สแกนตั้งแต่ต้นรอบปัจจุบันเผื่อมีวันตกหล่น
  var base = (maxDate && maxDate > today) ? maxDate : today;
  var end = new Date(base.getFullYear(), base.getMonth() + monthsAhead, base.getDate());

  var newRows = [];
  var cur = new Date(start);
  while (cur <= end) {
    if (cur.getDay() === 6) {
      var key2 = Utilities.formatDate(cur, tz, "yyyy-MM-dd");
      if (!existing[key2]) {
        newRows.push([new Date(cur), "", ""]);
        existing[key2] = true;
      }
    }
    cur.setDate(cur.getDate() + 1);
  }

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 3).setValues(newRows);
    var finalLastRow = sheet.getLastRow();
    sheet.getRange(2, 1, finalLastRow - 1, 3).sort({ column: 1, ascending: true });
  }

  var result = getSaturdaysData();
  result.generated = newRows.length;
  return result;
}

/**
 * เปลี่ยนสถานะแถว + ซิงค์ปฏิทิน
 * - ถ้าจะตั้งเป็น "หยุด" แต่รอบนั้นครบโควตาแล้ว จะไม่บันทึก และคืนค่า blocked:true
 * - ถ้าเปลี่ยนจาก "หยุด" กลับเป็น "ทำงาน" และเคยซิงค์ปฏิทินไว้ (DONE) จะลบ event ในปฏิทินให้ด้วย
 */
function updateStatusAndSync(rowIndex, status) {
  var sheet = getSheet_();
  var dateVal = sheet.getRange(rowIndex, 1).getValue();
  var prevStatus = sheet.getRange(rowIndex, 2).getValue();
  var syncStatus = sheet.getRange(rowIndex, 3).getValue();

  if (status === "หยุด" && prevStatus !== "หยุด" && dateVal instanceof Date) {
    var cyc = getCycle_(dateVal);
    var lastRow = sheet.getLastRow();
    var all = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    var count = 0;
    for (var i = 0; i < all.length; i++) {
      if (all[i][0] instanceof Date && all[i][1] === "หยุด") {
        if (getCycle_(all[i][0]).key === cyc.key) count++;
      }
    }
    if (count >= CAP) {
      var blockedResult = getSaturdaysData();
      blockedResult.blocked = true;
      blockedResult.message = "ใช้สิทธิ์หยุดเสาร์ครบ " + CAP + " วันในรอบ " + cyc.label + " แล้ว";
      return blockedResult;
    }
  }

  sheet.getRange(rowIndex, 2).setValue(status);

  var calendar = CalendarApp.getDefaultCalendar();
  if (status === "หยุด") {
    if (syncStatus !== "DONE" && dateVal instanceof Date) {
      var event = calendar.createAllDayEvent("หยุดเสาร์ (สิทธิ์โควตา)", dateVal);
      event.addPopupReminder(24 * 60); // แจ้งเตือนล่วงหน้า 1 วัน
      sheet.getRange(rowIndex, 3).setValue("DONE");
    }
  } else if (syncStatus === "DONE" && dateVal instanceof Date) {
    // กลับสถานะเป็นทำงาน — ลบ event ที่เคยซิงค์ไว้ ไม่ให้ค้างในปฏิทิน
    var events = calendar.getEventsForDay(dateVal);
    events.forEach(function (ev) {
      if (ev.getTitle().indexOf("หยุดเสาร์") !== -1) ev.deleteEvent();
    });
    sheet.getRange(rowIndex, 3).setValue("");
  }

  var result = getSaturdaysData();
  result.blocked = false;
  return result;
}
