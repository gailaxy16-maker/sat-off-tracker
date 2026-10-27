/**
 * ระบบติดตามวันหยุดเสาร์ — อิงข้อมูลจากชีท "SaturdayTracker"
 * คอลัมน์: A = วันที่ (Date) | B = สถานะ ("หยุด" / "ทำงาน") | C = สถานะ sync ปฏิทิน ("DONE" / ว่าง)
 *
 * รอบนับโควตาเริ่มวันที่ CUTOFF_DAY ของทุกเดือน ถึงวันที่ (CUTOFF_DAY - 1) ของเดือนถัดไป
 * ปรับ CUTOFF_DAY / CAP ด้านล่างได้ตามต้องการ
 *
 * v2 การเปลี่ยนแปลงหลัก:
 * - deleteCycle(): ลบวันเสาร์ทั้งรอบออกจากชีท (ปุ่ม ✕ หัวการ์ด)
 * - generateSaturdays(months, dryRun): นับจำนวนก่อนเติม แล้วค่อยเขียนจริงเมื่อยืนยัน
 * - ทุกการเขียนข้อมูลถูกห่อด้วย LockService กันการทำงานชนกัน
 * - resolveRow_(): หาแถวจาก dateISO ด้วย ถ้า rowIndex เลื่อนไปแล้ว (หลังลบแถวอื่น)
 *   จะไม่ไปแก้แถวผิด
 */
var CUTOFF_DAY = 16; // วันเริ่มรอบใหม่ของทุกเดือน
var CAP = 2;          // สิทธิ์วันหยุดเสาร์ต่อรอบ
var THAI_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
var SAT_EVENT_MARK = "หยุดเสาร์";

/**
 * ══════════════════════════════════════════
 * Router สำหรับ URL (.../exec)
 * ══════════════════════════════════════════
 */
function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : null;

  if (!action) {
    return HtmlService.createHtmlOutputFromFile('Index')
        .setTitle('ระบบติดตามวันหยุดเสาร์')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var p = e.parameter;
  var result;
  try {
    if (action === "getData") {
      result = getSaturdaysData();
    } else if (action === "updateStatus") {
      result = updateStatusAndSync(parseInt(p.rowIndex, 10), p.status, p.dateISO);
    } else if (action === "generate") {
      result = generateSaturdays(parseInt(p.months, 10) || 6, p.dryRun === "1");
    } else if (action === "deleteRow") {
      result = deleteSaturdayRow(parseInt(p.rowIndex, 10), p.dateISO);
    } else if (action === "deleteCycle") {
      result = deleteCycle(p.cycleKey);
    } else if (action === "clearCycle") {
      result = clearCycle(p.cycleKey);
    } else if (action === "getHolidays") {
      result = getHolidays();
    } else if (action === "addHoliday") {
      result = addHoliday(p.date, p.label);
    } else if (action === "deleteHoliday") {
      result = deleteHoliday(parseInt(p.rowIndex, 10), p.dateISO);
    } else {
      result = { error: "ไม่รู้จัก action: " + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ══════════════════════════════════════════
 * Helper ทั่วไป
 * ══════════════════════════════════════════
 */
function thaiShort_(date) {
  return date.getDate() + " " + THAI_MONTHS[date.getMonth()] + " " + String(date.getFullYear() + 543).slice(-2);
}

function isoDate_(date, tz) {
  return Utilities.formatDate(date, tz || Session.getScriptTimeZone(), "yyyy-MM-dd");
}

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

/** ห่อการเขียนข้อมูลด้วย lock กันการทำงานซ้อนกัน (เช่น กดลบหลายแถวติดกัน) */
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/**
 * หาแถวที่ถูกต้องจาก rowIndex + dateISO
 * - ถ้า rowIndex ยังชี้ไปยังวันที่เดิม ใช้แถวนั้น
 * - ถ้าไม่ตรง (แถวเลื่อนจากการลบ) ค้นหาจากวันที่แทน
 * คืน -1 ถ้าไม่พบ
 */
function resolveRow_(sheet, rowIndex, dateISO) {
  var tz = Session.getScriptTimeZone();
  var last = sheet.getLastRow();
  if (last < 2) return -1;

  if (rowIndex >= 2 && rowIndex <= last) {
    var v = sheet.getRange(rowIndex, 1).getValue();
    if (v instanceof Date && (!dateISO || isoDate_(v, tz) === dateISO)) return rowIndex;
  }
  if (!dateISO) return -1;

  var col = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    if (col[i][0] instanceof Date && isoDate_(col[i][0], tz) === dateISO) return i + 2;
  }
  return -1;
}

/** ผลลัพธ์กรณีแถวหายไปแล้ว: ส่งข้อมูลล่าสุดกลับไปให้หน้าเว็บ พร้อมแจ้งเตือน */
function staleResult_() {
  var r = computeSaturdaysData_();
  r.stale = true;
  r.message = "ข้อมูลเปลี่ยนไปแล้ว (อาจถูกลบหรือแก้จากที่อื่น) — โหลดข้อมูลล่าสุดให้แล้ว";
  return r;
}

function removeSatEvents_(calendar, d) {
  calendar.getEventsForDay(d).forEach(function (ev) {
    if (ev.getTitle().indexOf(SAT_EVENT_MARK) !== -1) ev.deleteEvent();
  });
}

/**
 * ══════════════════════════════════════════
 * อ่านข้อมูลวันเสาร์ (แคช 60 วินาที)
 * ══════════════════════════════════════════
 */
function getSaturdaysData() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('sat_data_v1');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* แคชเสีย ข้ามไปคำนวณใหม่ */ }
  }
  var result = computeSaturdaysData_();
  cacheSaturdaysData_(result);
  return result;
}

function cacheSaturdaysData_(result) {
  try {
    var json = JSON.stringify(result);
    if (json.length < 90000) CacheService.getScriptCache().put('sat_data_v1', json, 60);
  } catch (e) { /* ผลลัพธ์ใหญ่เกินหรือ error อื่น — ข้ามการแคชไปเงียบๆ */ }
}

function computeSaturdaysData_() {
  var sheet = getSheet_();
  if (!sheet) return { groups: [], cap: CAP, error: "ไม่พบชีท SaturdayTracker" };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { groups: [], cap: CAP };

  var tz = Session.getScriptTimeZone();
  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var todayKey = getCycle_(new Date()).key;

  var groupsMap = {};
  var order = [];

  for (var i = 0; i < data.length; i++) {
    var dateVal = data[i][0];
    var status = data[i][1];
    var syncStatus = data[i][2];
    if (!(dateVal instanceof Date)) continue;

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
      dateISO: isoDate_(dateVal, tz),
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
 * ══════════════════════════════════════════
 * เติมวันเสาร์ล่วงหน้า
 * dryRun = true → คืนจำนวนที่จะเพิ่ม โดยไม่เขียนลงชีท (ใช้แสดงให้ยืนยันก่อน)
 * ══════════════════════════════════════════
 */
function generateSaturdays(monthsAhead, dryRun) {
  return withLock_(function () {
    var sheet = getSheet_();
    if (!sheet) return { error: "ไม่พบชีท SaturdayTracker" };
    monthsAhead = monthsAhead || 6;
    var tz = Session.getScriptTimeZone();

    var existing = {};
    var maxDate = null;
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function (r) {
        if (r[0] instanceof Date) {
          existing[Utilities.formatDate(r[0], tz, "yyyy-MM-dd")] = true;
          if (!maxDate || r[0] > maxDate) maxDate = r[0];
        }
      });
    }

    var today = new Date();
    var start = getCycle_(today).start;
    var base = (maxDate && maxDate > today) ? maxDate : today;
    var end = new Date(base.getFullYear(), base.getMonth() + monthsAhead, base.getDate());

    var newRows = [];
    var cur = new Date(start);
    while (cur <= end) {
      if (cur.getDay() === 6) {
        var key = Utilities.formatDate(cur, tz, "yyyy-MM-dd");
        if (!existing[key]) {
          newRows.push([new Date(cur), "", ""]);
          existing[key] = true;
        }
      }
      cur.setDate(cur.getDate() + 1);
    }

    if (dryRun) {
      var n = newRows.length;
      return {
        dryRun: true,
        wouldAdd: n,
        from: n ? thaiShort_(newRows[0][0]) : "",
        to: n ? thaiShort_(newRows[n - 1][0]) : ""
      };
    }

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 3).setValues(newRows);
      var finalLastRow = sheet.getLastRow();
      sheet.getRange(2, 1, finalLastRow - 1, 3).sort({ column: 1, ascending: true });
    }

    var result = computeSaturdaysData_();
    cacheSaturdaysData_(result);
    result.generated = newRows.length;
    return result;
  });
}

/**
 * ══════════════════════════════════════════
 * เปลี่ยนสถานะรายวัน + ซิงค์ปฏิทิน
 * status = "หยุด" / "ทำงาน" / "" (ว่าง = ล้างการเลือก)
 * ══════════════════════════════════════════
 */
function updateStatusAndSync(rowIndex, status, dateISO) {
  return withLock_(function () {
    var sheet = getSheet_();
    if (!sheet) return { error: "ไม่พบชีท SaturdayTracker" };

    var row = resolveRow_(sheet, rowIndex, dateISO);
    if (row < 0) return staleResult_();

    var dateVal = sheet.getRange(row, 1).getValue();
    var prevStatus = sheet.getRange(row, 2).getValue();
    var syncStatus = sheet.getRange(row, 3).getValue();

    if (status === "หยุด" && prevStatus !== "หยุด" && dateVal instanceof Date) {
      var cyc = getCycle_(dateVal);
      var all = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      var count = 0;
      for (var i = 0; i < all.length; i++) {
        if (all[i][0] instanceof Date && all[i][1] === "หยุด" && getCycle_(all[i][0]).key === cyc.key) count++;
      }
      if (count >= CAP) {
        var blocked = getSaturdaysData();
        blocked.blocked = true;
        blocked.message = "ใช้สิทธิ์หยุดเสาร์ครบ " + CAP + " วันในรอบ " + cyc.label + " แล้ว";
        return blocked;
      }
    }

    sheet.getRange(row, 2).setValue(status);

    var calendar = CalendarApp.getDefaultCalendar();
    if (status === "หยุด") {
      if (syncStatus !== "DONE" && dateVal instanceof Date) {
        var ev = calendar.createAllDayEvent("หยุดเสาร์ (สิทธิ์โควตา)", dateVal);
        ev.addPopupReminder(24 * 60);
        sheet.getRange(row, 3).setValue("DONE");
      }
    } else if (syncStatus === "DONE" && dateVal instanceof Date) {
      removeSatEvents_(calendar, dateVal);
      sheet.getRange(row, 3).setValue("");
    }

    var result = computeSaturdaysData_();
    cacheSaturdaysData_(result);
    result.blocked = false;
    return result;
  });
}

/**
 * ลบแถววันเสาร์ทีละวัน (ใช้กับปุ่ม 🗑 รายแถว และตอน Undo หมดเวลา)
 * ถ้าเคยซิงค์ปฏิทิน (DONE) จะลบ event ในปฏิทินก่อนเสมอ
 */
function deleteSaturdayRow(rowIndex, dateISO) {
  return withLock_(function () {
    var sheet = getSheet_();
    if (!sheet) return { error: "ไม่พบชีท SaturdayTracker" };

    var row = resolveRow_(sheet, rowIndex, dateISO);
    if (row < 0) return staleResult_();

    var dateVal = sheet.getRange(row, 1).getValue();
    var syncStatus = sheet.getRange(row, 3).getValue();
    if (syncStatus === "DONE" && dateVal instanceof Date) {
      removeSatEvents_(CalendarApp.getDefaultCalendar(), dateVal);
    }

    sheet.deleteRow(row);

    var result = computeSaturdaysData_();
    cacheSaturdaysData_(result);
    result.deleted = true;
    return result;
  });
}

/**
 * ★ ใหม่: ลบวันเสาร์ทั้งรอบออกจากชีท (ปุ่ม ✕ หัวการ์ด)
 * cycleKey = วันที่เริ่มรอบ รูปแบบ yyyy-MM-dd
 * ลบ event ในปฏิทินของวันที่ซิงค์ไว้ด้วย แล้วลบแถวจากล่างขึ้นบน
 * (กันไม่ให้ rowIndex ของแถวที่เหลือเลื่อนระหว่างลบ)
 */
function deleteCycle(cycleKey) {
  return withLock_(function () {
    var sheet = getSheet_();
    if (!sheet) return { error: "ไม่พบชีท SaturdayTracker" };
    if (!cycleKey) return { error: "ไม่ได้ระบุรอบที่ต้องการลบ" };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return computeSaturdaysData_();

    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    var calendar = null;
    var targets = [];

    for (var i = 0; i < data.length; i++) {
      var d = data[i][0];
      if (!(d instanceof Date) || getCycle_(d).key !== cycleKey) continue;
      if (data[i][2] === "DONE") {
        if (!calendar) calendar = CalendarApp.getDefaultCalendar();
        removeSatEvents_(calendar, d);
      }
      targets.push(i + 2);
    }

    for (var j = targets.length - 1; j >= 0; j--) {
      sheet.deleteRow(targets[j]);
    }

    var result = computeSaturdaysData_();
    cacheSaturdaysData_(result);
    result.deleted = targets.length;
    return result;
  });
}

/**
 * ล้างการเลือกทั้งรอบ — วันเสาร์ยังอยู่ แค่กลับเป็น "ยังไม่ระบุ"
 * (ฟังก์ชันนี้ยังคงไว้ ย้ายปุ่มไปอยู่ท้ายการ์ดแทนปุ่ม ✕)
 */
function clearCycle(cycleKey) {
  return withLock_(function () {
    var sheet = getSheet_();
    if (!sheet) return { error: "ไม่พบชีท SaturdayTracker" };
    if (!cycleKey) return { error: "ไม่ได้ระบุรอบที่ต้องการล้าง" };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return computeSaturdaysData_();

    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    var calendar = null;
    var cleared = 0;

    for (var i = 0; i < data.length; i++) {
      var d = data[i][0];
      if (!(d instanceof Date)) continue;
      if (getCycle_(d).key !== cycleKey) continue;

      var hadStatus = data[i][1] === "หยุด" || data[i][1] === "ทำงาน";
      var hadSync = data[i][2] === "DONE";
      if (!hadStatus && !hadSync) continue;

      if (hadSync) {
        if (!calendar) calendar = CalendarApp.getDefaultCalendar();
        removeSatEvents_(calendar, d);
      }
      sheet.getRange(i + 2, 2, 1, 2).setValues([["", ""]]);
      cleared++;
    }

    var result = computeSaturdaysData_();
    cacheSaturdaysData_(result);
    result.cleared = cleared;
    result.blocked = false;
    return result;
  });
}

/**
 * ══════════════════════════════════════════
 * วันหยุดบริษัท (CompanyHolidays)
 * เก็บในชีทแยกชื่อ "CompanyHolidays" (A=วันที่, B=ชื่อวันหยุด, C=สถานะ sync
 * ปฏิทิน "DONE"/ว่าง) ทุกวันหยุดที่เพิ่มจะสร้าง event ลง Google Calendar
 * พร้อมแจ้งเตือนล่วงหน้า 1 วันอัตโนมัติ
 * ══════════════════════════════════════════
 */
var HOLIDAY_SHEET_NAME = "CompanyHolidays";
var HOLIDAY_EVENT_MARK = "วันหยุดบริษัท";

function getHolidaySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOLIDAY_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(HOLIDAY_SHEET_NAME);
    sh.getRange(1, 1, 1, 3).setValues([["Date", "Label", "SyncStatus"]]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getHolidays() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('holidays_v1');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* แคชเสีย ข้ามไปคำนวณใหม่ */ }
  }
  var result = computeHolidays_();
  cacheHolidays_(result);
  return result;
}

function cacheHolidays_(result) {
  try {
    var json = JSON.stringify(result);
    if (json.length < 90000) CacheService.getScriptCache().put('holidays_v1', json, 60);
  } catch (e) { /* ข้ามการแคชไปเงียบๆ */ }
}

function computeHolidays_() {
  var sh = getHolidaySheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { holidays: [] };
  var tz = Session.getScriptTimeZone();
  var data = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  var list = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] instanceof Date) {
      list.push({
        rowIndex: i + 2,
        dateISO: isoDate_(data[i][0], tz),
        label: String(data[i][1] || ""),
        syncStatus: String(data[i][2] || "")
      });
    }
  }
  list.sort(function (a, b) { return a.dateISO < b.dateISO ? -1 : a.dateISO > b.dateISO ? 1 : 0; });
  return { holidays: list };
}

function addHoliday(dateStr, label) {
  return withLock_(function () {
    if (!dateStr || !label) return { error: "กรุณาระบุวันที่และชื่อวันหยุดให้ครบ" };
    label = String(label).trim();
    if (!label) return { error: "กรุณาระบุชื่อวันหยุด" };

    var d = new Date(dateStr + "T00:00:00");
    if (isNaN(d.getTime())) return { error: "รูปแบบวันที่ไม่ถูกต้อง" };

    var tz = Session.getScriptTimeZone();
    var iso = isoDate_(d, tz);
    var existing = getHolidays().holidays;
    for (var i = 0; i < existing.length; i++) {
      if (existing[i].dateISO === iso) return { error: "มีวันหยุดวันที่นี้อยู่แล้ว: " + existing[i].label };
    }

    var calendar = CalendarApp.getDefaultCalendar();
    var event = calendar.createAllDayEvent(HOLIDAY_EVENT_MARK + ": " + label, d);
    event.addPopupReminder(24 * 60);

    var sh = getHolidaySheet_();
    sh.appendRow([d, label, "DONE"]);
    var lastRow = sh.getLastRow();
    if (lastRow > 2) sh.getRange(2, 1, lastRow - 1, 3).sort({ column: 1, ascending: true });

    var result = computeHolidays_();
    cacheHolidays_(result);
    return result;
  });
}

function deleteHoliday(rowIndex, dateISO) {
  return withLock_(function () {
    var sh = getHolidaySheet_();
    var row = resolveRow_(sh, rowIndex, dateISO);
    if (row < 0) return { error: "ไม่พบวันหยุดนี้แล้ว (อาจถูกลบไปแล้ว) กรุณารีเฟรชหน้า" };

    var dateVal = sh.getRange(row, 1).getValue();
    var syncStatus = sh.getRange(row, 3).getValue();

    if (syncStatus === "DONE" && dateVal instanceof Date) {
      var calendar = CalendarApp.getDefaultCalendar();
      calendar.getEventsForDay(dateVal).forEach(function (ev) {
        if (ev.getTitle().indexOf(HOLIDAY_EVENT_MARK) !== -1) ev.deleteEvent();
      });
    }

    sh.deleteRow(row);
    var result = computeHolidays_();
    cacheHolidays_(result);
    return result;
  });
}

/**
 * รันครั้งเดียวจาก Apps Script editor (เลือกฟังก์ชันนี้ → กด ▶ Run) เพื่อเติม
 * วันหยุดบริษัท TJG ปี 2026 ลงชีท CompanyHolidays — รันซ้ำจะข้ามวันที่มีอยู่แล้ว
 * หมายเหตุ: จะสร้าง event ในปฏิทินหลักทีเดียว 21 รายการ เป็นไปตั้งใจ
 */
function seedTJGHolidays2026() {
  var seed = [
    ["2026-01-01", "วันขึ้นปีใหม่"],
    ["2026-01-02", "วันหยุดปีใหม่"],
    ["2026-03-03", "วันมาฆบูชา"],
    ["2026-04-06", "วันจักรี"],
    ["2026-04-13", "วันสงกรานต์"],
    ["2026-04-14", "วันสงกรานต์"],
    ["2026-04-15", "วันสงกรานต์"],
    ["2026-05-01", "วันแรงงานแห่งชาติ"],
    ["2026-05-04", "วันฉัตรมงคล"],
    ["2026-05-31", "วันวิสาขบูชา"],
    ["2026-06-01", "ชดเชยวันวิสาขบูชา"],
    ["2026-06-03", "วันเฉลิมพระชนมพรรษาราชินี"],
    ["2026-07-28", "วันเฉลิมพระชนมพรรษา ร.10"],
    ["2026-07-29", "วันอาสาฬหบูชา"],
    ["2026-07-30", "วันเข้าพรรษา"],
    ["2026-08-12", "วันแม่แห่งชาติ"],
    ["2026-10-13", "วันคล้ายวันสวรรคต ร.9"],
    ["2026-10-23", "วันปิยมหาราช"],
    ["2026-12-05", "วันพ่อแห่งชาติ — ตรงกับวันเสาร์"],
    ["2026-12-30", "วันหยุดส่งท้ายปี"],
    ["2026-12-31", "วันสิ้นปี"]
  ];
  var added = 0, skipped = 0;
  seed.forEach(function (row) {
    var res = addHoliday(row[0], row[1]);
    if (res.error) { skipped++; } else { added++; }
  });
  Logger.log("เพิ่มใหม่ " + added + " วัน, ข้าม (มีอยู่แล้ว) " + skipped + " วัน");
  return "เพิ่มใหม่ " + added + " วัน, ข้าม (มีอยู่แล้ว) " + skipped + " วัน";
}
