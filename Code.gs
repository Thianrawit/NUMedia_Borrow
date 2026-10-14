/**
 * ============================================================
 * NU Media Club - Backend REST API (Single GAS Project)
 * ============================================================
 * Spreadsheet ID : 1i5niBkkHdqmrIFair6DANcdAfwGhpHjyR2szQz_tqKE
 * Drive Folder ID: 1AlycOE3qm-PbYqXcN44-bObCEYxiyJwTszIaqmAseJDvob53bXPfCq7TI4P-13zSzJcilrij
 * ============================================================
 * ไฟล์นี้เป็นไฟล์เดียวที่ต้องนำไปใส่ใน Google Apps Script Editor
 * Frontend จะโฮสต์อยู่ภายนอกบน GitHub Pages และสื่อสารผ่าน API_URL
 * ============================================================
 */

const SS_ID = "1i5niBkkHdqmrIFair6DANcdAfwGhpHjyR2szQz_tqKE";
const FOLDER_ID =
  "1AlycOE3qm-PbYqXcN44-bObCEYxiyJwTszIaqmAseJDvob53bXPfCq7TI4P-13zSzJcilrij";
const SH_BORROW_V1 = "ยืมอุปกรณ์";
const SH_RETURN = "คืนอุปกรณ์";
const SH_CATALOG = "คลังอุปกรณ์";
const SH_BORROW_V2 = "ยืมอุปกรณ์_v2";
const SH_ADMIN = "Admin";

/* ══════════════════════════════════════════════════
 * 1) doGet — Handle GET Requests
 * ══════════════════════════════════════════════════ */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || "";

    if (!action) {
      return createJsonResponse_({
        status: "online",
        service: "NU Media Club Equipment API",
        version: "2.0.0",
        timestamp: new Date().toISOString(),
      });
    }

    let result;
    if (action === "getDashboardData" || action === "getPublicBorrowList") {
      result = handleGetDashboardData();
    } else if (action === "getBorrowCatalog") {
      result = handleGetBorrowCatalog();
    } else if (
      action === "verifyLogin" ||
      action === "verifyLoginAndGetBorrows"
    ) {
      result = verifyLoginAndGetBorrows(
        e.parameter.email,
        e.parameter.studentId,
      );
    } else if (action === "adminGetDashboard") {
      result = {
        success: true,
        catalog: adminGetDashboard(e.parameter.adminEmail || ""),
      };
    } else {
      result = { success: false, message: "Unknown GET action: " + action };
    }

    return createJsonResponse_(result);
  } catch (err) {
    return createJsonResponse_({
      success: false,
      message: "doGet Error: " + err.toString(),
    });
  }
}

/* ══════════════════════════════════════════════════
 * 2) doPost — Handle POST Requests
 * ══════════════════════════════════════════════════ */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (errJson) {
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const action =
      payload.action || (e && e.parameter && e.parameter.action) || "";
    let result;

    if (action === "submitBorrowV2") {
      result = submitBorrowV2(payload);
    } else if (
      action === "verifyLogin" ||
      action === "verifyLoginAndGetBorrows"
    ) {
      result = verifyLoginAndGetBorrows(payload.email, payload.studentId);
    } else if (action === "submitReturn" || action === "submitReturnItem") {
      result = submitReturnItem(payload);
    } else if (action === "adminLogin") {
      result = adminLogin(payload.email, payload.password);
    } else if (action === "adminAddEquipment") {
      result = adminAddEquipment(
        payload.name,
        payload.category,
        payload.itemId,
        payload.imageBase64,
        payload.imageMime,
      );
    } else if (action === "adminDeleteEquipment") {
      result = adminDeleteEquipment(payload.rowIndex);
    } else if (action === "adminEditEquipment") {
      result = adminEditEquipment(payload);
    } else if (action === "adminDeleteMultiple") {
      result = adminDeleteMultipleEquipment(payload.rowIndexes);
    } else if (action === "adminGetDashboard") {
      result = {
        success: true,
        catalog: adminGetDashboard(payload.adminEmail || ""),
      };
    } else {
      result = { success: false, message: "Unknown POST action: " + action };
    }

    return createJsonResponse_(result);
  } catch (err) {
    return createJsonResponse_({
      success: false,
      message: "doPost Error: " + err.toString(),
    });
  }
}

/* ══════════════════════════════════════════════════
 * HELPER — Response, Spreadsheet, Drive
 * ══════════════════════════════════════════════════ */
function createJsonResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

function ss_() {
  return SpreadsheetApp.openById(SS_ID);
}
function sheet_(n) {
  return ss_().getSheetByName(n);
}

function fmtDate_(v) {
  if (!v) return "";
  if (v instanceof Date)
    return Utilities.formatDate(v, "Asia/Bangkok", "dd/MM/yyyy");
  return v.toString();
}

function driveUrl_(raw) {
  if (!raw) return "";
  raw = raw.toString().trim();
  if (raw.indexOf("lh3.googleusercontent.com") !== -1) return raw;
  var m = raw.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (!m) m = raw.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return m ? "https://lh3.googleusercontent.com/d/" + m[1] : raw;
}

function driveUrls_(raw) {
  if (!raw) return [];
  return raw
    .toString()
    .split(",")
    .map(function (u) {
      return driveUrl_(u.trim());
    })
    .filter(Boolean);
}

function uploadImage_(base64, mime, prefix) {
  if (!base64 || !mime) return "";
  try {
    var decoded = Utilities.base64Decode(base64);
    var ext = mime.indexOf("png") !== -1 ? ".png" : ".jpg";
    var blob = Utilities.newBlob(
      decoded,
      mime,
      prefix + "_" + Date.now() + ext,
    );
    var folder;
    try {
      folder = DriveApp.getFolderById(FOLDER_ID);
    } catch (ef) {
      folder = DriveApp.getRootFolder();
    }
    var file = folder.createFile(blob);
    try {
      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW,
      );
    } catch (es) {
      /* org policy */
    }
    return "https://lh3.googleusercontent.com/d/" + file.getId();
  } catch (eu) {
    Logger.log("Upload error: " + eu.message);
    return "";
  }
}

function authorizeDrive() {
  var f = DriveApp.getFolderById(FOLDER_ID);
  Logger.log("Drive OK: " + f.getName());
}

/* ══════════════════════════════════════════════════
 * 3) API Endpoints Implementation
 * ══════════════════════════════════════════════════ */

// --- Dashboard Data ---
function handleGetDashboardData() {
  var data = getDashboardData();
  return {
    success: true,
    data: data.borrows, // compatibility with old index.html
    borrows: data.borrows,
    catalog: data.catalog,
  };
}

function getDashboardData() {
  var result = [];

  // V1 Legacy
  var s1 = sheet_(SH_BORROW_V1);
  if (s1) {
    var d1 = s1.getDataRange().getValues();
    for (var i = 1; i < d1.length; i++) {
      var r = d1[i];
      var st = (r[14] || "").toString().trim();
      if (st === "คืนแล้ว") continue;
      if (!st) st = "กำลังยืม";
      var items = (r[8] || "").toString().trim();
      var name = (r[3] || "").toString().trim();
      if (!items && !name) continue;
      result.push({
        nickname: r[4] || "",
        department: r[6] || "",
        items: items,
        borrowDate: fmtDate_(r[9]),
        dueDate: fmtDate_(r[13]),
        status: st,
        groupImageUrl: driveUrl_(r[10] || ""),
        imageUrls: driveUrls_(r[10] || ""),
        isLegacy: true,
      });
    }
  }

  // V2 New
  var s2 = sheet_(SH_BORROW_V2);
  if (s2) {
    var d2 = s2.getDataRange().getValues();
    for (var j = 1; j < d2.length; j++) {
      var r2 = d2[j];
      var st2 = (r2[16] || "").toString().trim();
      if (st2 === "คืนแล้ว") continue;
      if (!st2) st2 = "กำลังยืม";
      var items2 = (r2[8] || "").toString().trim();
      var name2 = (r2[3] || "").toString().trim();
      if (!items2 && !name2) continue;
      result.push({
        nickname: r2[4] || "",
        department: r2[6] || "",
        items: items2,
        borrowDate: fmtDate_(r2[9]),
        dueDate: fmtDate_(r2[15]),
        status: st2,
        groupImageUrl: driveUrl_(r2[10] || ""),
        imageUrls: driveUrls_(r2[10] || ""),
        isLegacy: false,
      });
    }
  }

  var catalog = getCatalogRaw_();
  return { borrows: result, catalog: catalog };
}

// --- Borrow Catalog ---
function handleGetBorrowCatalog() {
  var catalog = getBorrowCatalog();
  return {
    success: true,
    data: catalog,
    catalog: catalog,
  };
}

function getBorrowCatalog() {
  var catalog = getCatalogRaw_();
  var borrowedSet = {};

  // Check V1
  var s1 = sheet_(SH_BORROW_V1);
  if (s1) {
    var d1 = s1.getDataRange().getValues();
    for (var i = 1; i < d1.length; i++) {
      var st = (d1[i][14] || "").toString().trim();
      if (st === "คืนแล้ว") continue;
      var itm = (d1[i][8] || "").toString().trim();
      if (itm) {
        itm.split(",").forEach(function (x) {
          borrowedSet[x.trim().toLowerCase()] = true;
        });
      }
    }
  }

  // Check V2
  var s2 = sheet_(SH_BORROW_V2);
  if (s2) {
    var d2 = s2.getDataRange().getValues();
    for (var j = 1; j < d2.length; j++) {
      var st2 = (d2[j][16] || "").toString().trim();
      if (st2 === "คืนแล้ว") continue;
      var itm2 = (d2[j][8] || "").toString().trim();
      if (itm2) {
        itm2.split(",").forEach(function (x) {
          borrowedSet[x.trim().toLowerCase()] = true;
        });
      }
    }
  }

  catalog.forEach(function (c) {
    var st = (c.status || "").trim();
    var isMaint =
      st === "ส่งซ่อม/ชำรุด" ||
      st.indexOf("ส่งซ่อม") !== -1 ||
      st.indexOf("ชำรุด") !== -1;
    var isBorrowedByBorrowSheet = !!borrowedSet[(c.name || "").toLowerCase()];
    var isBorrowedByCatalogStatus = st === "ถูกยืม";
    c.isMaintenance = isMaint;
    c.isBorrowed = isBorrowedByBorrowSheet || isBorrowedByCatalogStatus;
  });

  return catalog;
}

function getCatalogRaw_() {
  var sh = sheet_(SH_CATALOG);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!(r[0] || "").toString().trim() && !(r[1] || "").toString().trim())
      continue;
    out.push({
      rowIndex: i + 1,
      id: (r[0] || "").toString().trim(),
      name: (r[1] || "").toString().trim(),
      imageUrl: driveUrl_(r[2] || ""),
      dateAdded: fmtDate_(r[3]),
      status: (r[4] || "พร้อมใช้งาน").toString().trim(),
      category: (r[5] || "").toString().trim(),
    });
  }
  return out;
}

// --- Submit Borrow V2 ---
function submitBorrowV2(payload) {
  try {
    // ตรวจสอบสถานะอุปกรณ์ใน "คลังอุปกรณ์" ว่ามีรายการส่งซ่อม/ชำรุด หรือไม่
    var selectedList = (payload.selectedItems || "")
      .split(",")
      .map(function (s) {
        return s.trim().toLowerCase();
      })
      .filter(Boolean);

    var catSheet = sheet_(SH_CATALOG);
    if (catSheet && selectedList.length > 0) {
      var catData = catSheet.getDataRange().getValues();
      for (var ck = 1; ck < catData.length; ck++) {
        var cName = (catData[ck][1] || "").toString().trim().toLowerCase();
        if (selectedList.indexOf(cName) !== -1) {
          var cStatus = (catData[ck][4] || "").toString().trim();
          if (
            cStatus === "ส่งซ่อม/ชำรุด" ||
            cStatus.indexOf("ส่งซ่อม") !== -1 ||
            cStatus.indexOf("ชำรุด") !== -1
          ) {
            return {
              success: false,
              message:
                'อุปกรณ์ "' +
                (catData[ck][1] || "") +
                '" อยู่ในสถานะ ส่งซ่อม/ชำรุด ไม่สามารถยืมได้',
            };
          }
        }
      }
    }

    var groupUrl = uploadImage_(
      payload.groupImageBase64,
      payload.groupImageMime || "image/jpeg",
      "borrow_group",
    );
    var otherUrl = "";
    if (payload.otherImageBase64) {
      otherUrl = uploadImage_(
        payload.otherImageBase64,
        payload.otherImageMime || "image/jpeg",
        "borrow_other",
      );
    }

    var sh = sheet_(SH_BORROW_V2);
    var row = [
      new Date(), // A: ประทับเวลา
      (payload.email || "").trim(), // B: อีเมล
      payload.agreement || "ยอมรับ", // C: ข้อตกลง
      (payload.fullName || "").trim(), // D: ชื่อ-สกุล
      (payload.nickname || "").trim(), // E: ชื่อเล่น
      (payload.studentId || "").trim(), // F: รหัสนิสิต
      (payload.department || "").trim(), // G: ฝ่าย
      (payload.contact || "").trim(), // H: ช่องทางติดต่อ
      (payload.selectedItems || "").trim(), // I: รายชื่ออุปกรณ์
      payload.borrowDate || "", // J: วันที่ยืม
      groupUrl, // K: รูปรวมหน้างาน
      otherUrl, // L: รูปอุปกรณ์อื่นๆ
      (payload.otherDetail || "").trim(), // M: รายละเอียดอื่นๆ
      (payload.toldWhom || "").trim(), // N: บอกใคร
      (payload.roomOpener || "").trim(), // O: ใครเปิดห้อง
      payload.dueDate || "", // P: วันที่คืน
      "กำลังยืม", // Q: สถานะ
      "", // R: แจ้งเตือนล่าสุด
    ];
    sh.appendRow(row);

    // อัปเดตสถานะใน "คลังอุปกรณ์" เป็น "ถูกยืม"
    try {
      var catSheet = sheet_(SH_CATALOG);
      if (catSheet) {
        var catData = catSheet.getDataRange().getValues();
        var selectedList = (payload.selectedItems || "")
          .split(",")
          .map(function (s) {
            return s.trim().toLowerCase();
          });
        for (var k = 1; k < catData.length; k++) {
          var eqName = (catData[k][1] || "").toString().trim().toLowerCase();
          if (selectedList.indexOf(eqName) !== -1) {
            catSheet.getRange(k + 1, 5).setValue("ถูกยืม");
          }
        }
      }
    } catch (eCat) {
      Logger.log("Update catalog status on borrow error: " + eCat.message);
    }

    return { success: true, message: "บันทึกการยืมอุปกรณ์สำเร็จ 🎉" };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + err.message };
  }
}

// --- Verify Login (V1 + V2) ---
function verifyLoginAndGetBorrows(email, studentId) {
  email = (email || "").toLowerCase().trim();
  studentId = (studentId || "").toString().replace(/\D/g, "");
  if (!email || !studentId) {
    return { success: false, message: "กรุณากรอกอีเมลและรหัสนิสิตให้ครบถ้วน" };
  }

  var user = null;
  var items = [];

  // V1 Legacy
  var s1 = sheet_(SH_BORROW_V1);
  if (s1) {
    var d1 = s1.getDataRange().getValues();
    for (var i = 1; i < d1.length; i++) {
      var r = d1[i];
      var rEmail = (r[1] || "").toString().toLowerCase().trim();
      var rSid = (r[5] || "").toString().replace(/\D/g, "");
      if (rEmail !== email || rSid !== studentId) continue;
      if (!user) {
        user = {
          name: r[3] || "",
          nickname: r[4] || "",
          studentId: rSid,
          email: rEmail,
          department: r[6] || "",
        };
      }
      var st = (r[14] || "").toString().trim();
      if (st === "คืนแล้ว") continue;
      if (!st) st = "กำลังยืม";
      items.push({
        sourceSheet: SH_BORROW_V1,
        rowIndex: i + 1,
        items: (r[8] || "").toString(),
        borrowDate: fmtDate_(r[9]),
        dueDate: fmtDate_(r[13]),
        imageUrl: driveUrl_(r[10] || ""),
        groupImageUrl: driveUrl_(r[10] || ""),
        status: st,
      });
    }
  }

  // V2 New
  var s2 = sheet_(SH_BORROW_V2);
  if (s2) {
    var d2 = s2.getDataRange().getValues();
    for (var j = 1; j < d2.length; j++) {
      var r2 = d2[j];
      var rEmail2 = (r2[1] || "").toString().toLowerCase().trim();
      var rSid2 = (r2[5] || "").toString().replace(/\D/g, "");
      if (rEmail2 !== email || rSid2 !== studentId) continue;
      if (!user) {
        user = {
          name: r2[3] || "",
          nickname: r2[4] || "",
          studentId: rSid2,
          email: rEmail2,
          department: r2[6] || "",
        };
      }
      var st2 = (r2[16] || "").toString().trim();
      if (st2 === "คืนแล้ว") continue;
      if (!st2) st2 = "กำลังยืม";
      items.push({
        sourceSheet: SH_BORROW_V2,
        rowIndex: j + 1,
        items: (r2[8] || "").toString(),
        borrowDate: fmtDate_(r2[9]),
        dueDate: fmtDate_(r2[15]),
        imageUrl: driveUrl_(r2[10] || ""),
        groupImageUrl: driveUrl_(r2[10] || ""),
        status: st2,
      });
    }
  }

  if (!user) {
    return {
      success: true,
      user: null,
      items: [],
      hasHistory: false,
      message: "ไม่มีรายการยืมอุปกรณ์ในตอนนี้",
    };
  }
  return { success: true, user: user, items: items, hasHistory: true };
}

// --- Submit Return ---
function submitReturnItem(payload) {
  try {
    var srcSheet = payload.sourceSheet || SH_BORROW_V1;
    var rowIdx = Number(payload.rowIndex);
    if (!rowIdx || isNaN(rowIdx)) {
      return { success: false, message: "ข้อมูลแถวไม่ถูกต้อง" };
    }

    var sh = sheet_(srcSheet);
    if (!sh) return { success: false, message: "ไม่พบชีต: " + srcSheet };

    // ตรวจสอบรหัสนิสิต Col F (idx 5)
    var sheetSid = (sh.getRange(rowIdx, 6).getValue() || "")
      .toString()
      .replace(/\D/g, "");
    var cleanSid = (payload.studentId || "").toString().replace(/\D/g, "");
    if (cleanSid && sheetSid && cleanSid !== sheetSid) {
      return { success: false, message: "รหัสนิสิตไม่ตรงกับรายการนี้" };
    }

    // อัปโหลดรูปหลักฐานคืน
    var imageUrl = uploadImage_(
      payload.imageBase64,
      payload.imageMimeType || "image/jpeg",
      "return",
    );

    // ดึงข้อมูลแถวเดิม
    var numCols = srcSheet === SH_BORROW_V2 ? 18 : 15;
    var borrowRow = sh.getRange(rowIdx, 1, 1, numCols).getValues()[0];

    var emailVal = borrowRow[1] || "";
    var nameVal = borrowRow[3] || "";
    var nickVal = borrowRow[4] || "";
    var sidVal = borrowRow[5] || "";
    var deptVal = borrowRow[6] || "";
    var contactVal = borrowRow[7] || "";
    var itemsVal = borrowRow[8] || "";
    var borrowDateVal = borrowRow[9] || "";

    // บันทึกประวัติลง "คืนอุปกรณ์"
    var ts = new Date();
    var returnRow = [
      ts, // A: ประทับเวลา
      emailVal, // B: อีเมล
      nameVal, // C: ชื่อ-สกุล
      nickVal, // D: ชื่อเล่น
      sidVal, // E: รหัสนิสิต
      deptVal, // F: ฝ่าย
      contactVal, // G: ช่องทางติดต่อ
      itemsVal, // H: รายชื่ออุปกรณ์
      borrowDateVal, // I: วันที่ยืม
      imageUrl, // J: รูปหลักฐานคืน
      payload.toldWhom || "", // K: บอกใคร
      payload.roomOpener || "", // L: ใครเปิดห้อง
      Utilities.formatDate(ts, "Asia/Bangkok", "d/M/yyyy, HH:mm:ss"), // M: วันที่คืน
    ];
    var returnSheet = sheet_(SH_RETURN);
    if (returnSheet) returnSheet.appendRow(returnRow);

    // อัปเดตสถานะในแท็บยืมเป็น "คืนแล้ว"
    var statusCol = srcSheet === SH_BORROW_V2 ? 17 : 15;
    sh.getRange(rowIdx, statusCol).setValue("คืนแล้ว");

    // อัปเดตสถานะใน "คลังอุปกรณ์" กลับเป็น "พร้อมใช้งาน"
    try {
      var catSheet = sheet_(SH_CATALOG);
      if (catSheet && itemsVal) {
        var catData = catSheet.getDataRange().getValues();
        var returnedList = itemsVal
          .toString()
          .split(",")
          .map(function (s) {
            return s.trim().toLowerCase();
          });
        for (var c = 1; c < catData.length; c++) {
          var eqName = (catData[c][1] || "").toString().trim().toLowerCase();
          if (returnedList.indexOf(eqName) !== -1) {
            var curSt = (catData[c][4] || "").toString().trim();
            if (
              curSt !== "ส่งซ่อม/ชำรุด" &&
              curSt.indexOf("ส่งซ่อม") === -1 &&
              curSt.indexOf("ชำรุด") === -1
            ) {
              catSheet.getRange(c + 1, 5).setValue("พร้อมใช้งาน");
            }
          }
        }
      }
    } catch (eResetCat) {
      Logger.log("Reset catalog status error: " + eResetCat.message);
    }

    return { success: true, message: "คืนอุปกรณ์สำเร็จเรียบร้อยแล้ว 🎉" };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + err.message };
  }
}

// --- Admin ---
function adminLogin(email, password) {
  email = (email || "").toLowerCase().trim();
  password = (password || "").trim();
  if (!email || !password) {
    return { success: false, message: "กรุณากรอกข้อมูลให้ครบ" };
  }

  var sh = sheet_(SH_ADMIN);
  if (!sh) return { success: false, message: "ไม่พบชีต Admin" };

  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var rEmail = (r[0] || "").toString().toLowerCase().trim();
    var rPass = (r[1] || "").toString().trim();
    if (rEmail === email && rPass === password) {
      return {
        success: true,
        admin: {
          email: rEmail,
          name: r[2] || "",
          role: r[3] || "Staff",
        },
      };
    }
  }
  return { success: false, message: "อีเมลหรือรหัสผ่านไม่ถูกต้อง" };
}

function adminGetDashboard(adminEmail) {
  var catalog = getCatalogRaw_();
  var borrowMap = {};

  // V1
  var s1 = sheet_(SH_BORROW_V1);
  if (s1) {
    var d1 = s1.getDataRange().getValues();
    for (var i = 1; i < d1.length; i++) {
      var r = d1[i];
      var st = (r[14] || "").toString().trim();
      if (st === "คืนแล้ว") continue;
      if (!st) st = "กำลังยืม";
      var names = (r[8] || "").toString().split(",");
      names.forEach(function (n) {
        var key = n.trim().toLowerCase();
        if (key) {
          borrowMap[key] = {
            nickname: r[4] || "",
            studentId: (r[5] || "").toString(),
            borrowDate: fmtDate_(r[9]),
            dueDate: fmtDate_(r[13]),
            status: st,
          };
        }
      });
    }
  }

  // V2
  var s2 = sheet_(SH_BORROW_V2);
  if (s2) {
    var d2 = s2.getDataRange().getValues();
    for (var j = 1; j < d2.length; j++) {
      var r2 = d2[j];
      var st2 = (r2[16] || "").toString().trim();
      if (st2 === "คืนแล้ว") continue;
      if (!st2) st2 = "กำลังยืม";
      var names2 = (r2[8] || "").toString().split(",");
      names2.forEach(function (n) {
        var key = n.trim().toLowerCase();
        if (key) {
          borrowMap[key] = {
            nickname: r2[4] || "",
            studentId: (r2[5] || "").toString(),
            borrowDate: fmtDate_(r2[9]),
            dueDate: fmtDate_(r2[15]),
            status: st2,
          };
        }
      });
    }
  }

  catalog.forEach(function (c) {
    var key = (c.name || "").toLowerCase();
    if (borrowMap[key]) {
      c.borrower = borrowMap[key];
    } else {
      c.borrower = null;
    }
  });

  return catalog;
}

function getNextEquipmentId_(catOrPrefix) {
  var prefixMap = {
    ไมค์: "MIC",
    ขาตั้ง: "POD",
    ไฟ: "LAM",
    เลนส์: "LEN",
    กริมบอล: "GIM",
    กล้อง: "CAM",
    อื่นๆ: "NUM",
    MIC: "MIC",
    POD: "POD",
    LAM: "LAM",
    LEN: "LEN",
    GIM: "GIM",
    CAM: "CAM",
    NUM: "NUM",
  };
  var prefix = prefixMap[catOrPrefix] || "NUM";
  var sh = sheet_(SH_CATALOG);
  if (!sh) return prefix + "-001";

  var data = sh.getDataRange().getValues();
  var maxNum = 0;
  var re = new RegExp("^" + prefix + "[-_]?(\\d+)", "i");

  for (var i = 1; i < data.length; i++) {
    var idStr = (data[i][0] || "").toString().trim();
    if (!idStr) continue;
    var match = idStr.match(re);
    if (match) {
      var num = parseInt(match[1], 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }
  var nextNum = maxNum + 1;
  return prefix + "-" + String(nextNum).padStart(3, "0");
}

function adminAddEquipment(name, category, itemId, imageBase64, imageMime) {
  try {
    var imgUrl = uploadImage_(
      imageBase64,
      imageMime || "image/jpeg",
      "catalog",
    );
    var sh = sheet_(SH_CATALOG);
    var autoId =
      (itemId || "").toString().trim() || getNextEquipmentId_(category);

    sh.appendRow([
      autoId, // A: รหัสอุปกรณ์
      (name || "").trim(), // B: ชื่ออุปกรณ์
      imgUrl, // C: ลิงก์รูปภาพ Drive
      Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd"), // D: วันที่เพิ่ม
      "พร้อมใช้งาน", // E: สถานะ
      (category || "").trim(), // F: ประเภท
    ]);

    return {
      success: true,
      message: "เพิ่มอุปกรณ์สำเร็จ: " + name + " (" + autoId + ")",
      id: autoId,
    };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + err.message };
  }
}

function adminDeleteEquipment(rowIndex) {
  try {
    var sh = sheet_(SH_CATALOG);
    sh.deleteRow(rowIndex);
    return { success: true, message: "ลบอุปกรณ์สำเร็จ" };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + err.message };
  }
}

function adminEditEquipment(payload) {
  try {
    var sh = sheet_(SH_CATALOG);
    var rowIndex = parseInt(payload.rowIndex, 10); // 1-based row in sheet
    if (!rowIndex || rowIndex < 2) {
      return { success: false, message: "ไม่พบแถวที่ต้องการแก้ไข" };
    }

    // อัปเดตรหัสอุปกรณ์ (Col A)
    if (payload.newId !== undefined && payload.newId !== null) {
      sh.getRange(rowIndex, 1).setValue(payload.newId.toString().trim());
    }
    // อัปเดตชื่ออุปกรณ์ (Col B)
    if (payload.newName !== undefined && payload.newName !== null) {
      sh.getRange(rowIndex, 2).setValue(payload.newName.toString().trim());
    }
    // อัปเดตสถานะ (Col E)
    if (payload.newStatus !== undefined && payload.newStatus !== null) {
      sh.getRange(rowIndex, 5).setValue(payload.newStatus.toString().trim());
    }
    // อัปเดตหมวดหมู่ (Col F)
    if (payload.newCategory !== undefined && payload.newCategory !== null) {
      sh.getRange(rowIndex, 6).setValue(payload.newCategory.toString().trim());
    }
    // อัปเดตรูปภาพ (Col C)
    if (payload.removeImage === true) {
      sh.getRange(rowIndex, 3).setValue("");
    } else if (payload.imageBase64 && payload.imageBase64.length > 0) {
      var imgUrl = uploadImage_(
        payload.imageBase64,
        payload.imageMime || "image/jpeg",
        "catalog",
      );
      if (imgUrl) {
        sh.getRange(rowIndex, 3).setValue(imgUrl);
      }
    }

    return { success: true, message: "แก้ไขอุปกรณ์สำเร็จ" };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + err.message };
  }
}

function adminDeleteMultipleEquipment(rowIndexes) {
  try {
    if (!rowIndexes || !Array.isArray(rowIndexes) || rowIndexes.length === 0) {
      return { success: false, message: "ไม่มีรายการที่เลือก" };
    }
    var sh = sheet_(SH_CATALOG);
    // Deduplicate and filter valid rows (>= 2)
    var rowSet = {};
    rowIndexes.forEach(function (r) {
      var n = parseInt(r, 10);
      if (!isNaN(n) && n >= 2) rowSet[n] = true;
    });
    var sorted = Object.keys(rowSet)
      .map(Number)
      .sort(function (a, b) {
        return b - a;
      });

    if (sorted.length === 0) {
      return { success: false, message: "ไม่พบแถวที่สามารถลบได้" };
    }

    for (var i = 0; i < sorted.length; i++) {
      sh.deleteRow(sorted[i]);
    }
    return {
      success: true,
      message: "ลบอุปกรณ์สำเร็จ " + sorted.length + " รายการ",
    };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + err.message };
  }
}
