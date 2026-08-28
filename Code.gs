/**
 * ============================================================
 * NU Media Club - ระบบยืม-คืนอุปกรณ์ (Backend REST API)
 * ============================================================
 * Spreadsheet ID : 1i5niBkkHdqmrIFair6DANcdAfwGhpHjyR2szQz_tqKE
 * Drive Folder ID: 1AlycOE3qm-PbYqXcN44-bObCEYxiyJwTszIaqmAseJDvob53bXPfCq7TI4P-13zSzJcilrij
 * ============================================================
 */

const SPREADSHEET_ID  = '1i5niBkkHdqmrIFair6DANcdAfwGhpHjyR2szQz_tqKE';
const DRIVE_FOLDER_ID = '1AlycOE3qm-PbYqXcN44-bObCEYxiyJwTszIaqmAseJDvob53bXPfCq7TI4P-13zSzJcilrij';
const SHEET_BORROW    = 'ยืมอุปกรณ์';
const SHEET_RETURN    = 'คืนอุปกรณ์';

/**
 * ──────────────────────────────────────────────
 * ฟังก์ชันสำหรับกด Run ใน Apps Script Editor 1 ครั้ง
 * เพื่อเปิดสิทธิ์การเข้าถึง Google Drive (DriveApp)
 * ──────────────────────────────────────────────
 */
function authorizeDrive() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  Logger.log("Drive Access OK! Folder Name: " + folder.getName());
}

/* ──────────────────────────────────────────────
 * API Router – doGet (GET Requests)
 * ────────────────────────────────────────────── */
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const action = params.action || '';

    let result;
    if (action === 'getPublicBorrowList') {
      result = { success: true, data: getPublicBorrowList() };
    } else if (action === 'verifyLogin') {
      result = verifyLoginAndGetBorrows(params.email, params.studentId);
    } else {
      result = {
        success: true,
        message: 'NU Media Borrow API is running 🚀',
        endpoints: ['?action=getPublicBorrowList', '?action=verifyLogin&email=...&studentId=...']
      };
    }

    return createJsonResponse_(result);
  } catch (err) {
    return createJsonResponse_({ success: false, message: 'Server Error: ' + err.message });
  }
}

/* ──────────────────────────────────────────────
 * API Router – doPost (POST Requests)
 * ────────────────────────────────────────────── */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const action = payload.action || '';
    let result;

    if (action === 'submitReturn') {
      result = submitReturnItem(payload);
    } else if (action === 'verifyLogin') {
      result = verifyLoginAndGetBorrows(payload.email, payload.studentId);
    } else if (action === 'getPublicBorrowList') {
      result = { success: true, data: getPublicBorrowList() };
    } else {
      result = { success: false, message: 'Invalid action: ' + action };
    }

    return createJsonResponse_(result);
  } catch (err) {
    return createJsonResponse_({ success: false, message: 'Server Error: ' + err.message });
  }
}

/* ──────────────────────────────────────────────
 * 1) getPublicBorrowList
 *    ดึงรายการที่ยังไม่คืน (สถานะ != "คืนแล้ว")
 *    ไม่ส่ง email & studentId เพื่อความปลอดภัย
 * ────────────────────────────────────────────── */
function getPublicBorrowList() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_BORROW);
  const data  = sheet.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const status = (row[14] || '').toString().trim(); // Col O (idx 14)

    if (status === 'คืนแล้ว' || status === '') continue;

    result.push({
      nickname:   row[4]  || '',                          // Col E – ชื่อเล่น
      department: row[6]  || '',                          // Col G – ฝ่าย
      items:      row[8]  || '',                          // Col I – รายชื่ออุปกรณ์
      borrowDate: formatDate_(row[9]),                   // Col J – วันที่ยืม
      dueDate:    formatDate_(row[13]),                  // Col N – วันที่กำหนดคืน
      status:     status,                                 // Col O – สถานะ
      imageUrl:   convertDriveUrl_(row[10] || '')         // Col K – รูปอุปกรณ์ (แปลงให้โหลดได้)
    });
  }

  return result;
}

/* ──────────────────────────────────────────────
 * 2) verifyLoginAndGetBorrows
 *    ตรวจสอบ email + รหัสนิสิต แล้วส่งรายการยืมกลับ
 * ────────────────────────────────────────────── */
function verifyLoginAndGetBorrows(email, studentId) {
  email     = (email || '').toLowerCase().trim();
  studentId = (studentId || '').toString().replace(/\D/g, '');

  if (!email || !studentId) {
    return { success: false, message: 'กรุณากรอกอีเมลและรหัสนิสิตให้ครบถ้วน' };
  }

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_BORROW);
  const data  = sheet.getDataRange().getValues();

  let user  = null;
  const items = [];

  for (let i = 1; i < data.length; i++) {
    const row       = data[i];
    const rowEmail  = (row[1] || '').toString().toLowerCase().trim();   // Col B
    const rowSid    = (row[5] || '').toString().replace(/\D/g, '');      // Col F
    const rowStatus = (row[14] || '').toString().trim();                  // Col O

    if (rowEmail === email && rowSid === studentId) {
      if (!user) {
        user = {
          name:       row[3] || '',  // Col D – ชื่อ-นามสกุล
          nickname:   row[4] || '',  // Col E – ชื่อเล่น
          studentId:  rowSid,
          email:      rowEmail,
          department: row[6] || ''   // Col G – ฝ่าย
        };
      }
      if (rowStatus !== 'คืนแล้ว' && rowStatus !== '') {
        items.push({
          rowIndex:   i + 1,                            // 1-based row number ใน Sheet
          items:      row[8]  || '',                     // Col I
          borrowDate: formatDate_(row[9]),               // Col J
          dueDate:    formatDate_(row[13]),              // Col N
          imageUrl:   convertDriveUrl_(row[10] || ''),   // Col K
          status:     rowStatus                          // Col O
        });
      }
    }
  }

  // กรณี 1: ไม่เคยมีประวัติการยืมในระบบเลย
  if (!user) {
    return { success: true, user: null, items: [], hasHistory: false, message: 'ไม่มีรายการยืมอุปกรณ์ในตอนนี้' };
  }

  // กรณี 2: มีประวัติการยืม (items อาจว่างถ้าคืนหมดแล้ว หรือมีรายการค้างส่ง)
  return { success: true, user: user, items: items, hasHistory: true };
}

/* ──────────────────────────────────────────────
 * 3) submitReturnItem
 *    บันทึกการคืนอุปกรณ์ + อัปโหลดรูปลง Google Drive
 * ────────────────────────────────────────────── */
function submitReturnItem(payload) {
  try {
    const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    const borrowSheet = ss.getSheetByName(SHEET_BORROW);
    const returnSheet = ss.getSheetByName(SHEET_RETURN);
    const rowIdx      = Number(payload.rowIndex); // 1-based

    if (!rowIdx || isNaN(rowIdx)) {
      return { success: false, message: 'Row index ไม่ถูกต้อง' };
    }

    // --- ดึงข้อมูลเดิมจากแท็บยืม ---
    const borrowRow = borrowSheet.getRange(rowIdx, 1, 1, 15).getValues()[0];
    const emailVal      = borrowRow[1]  || ''; // Col B
    const nameVal       = borrowRow[3]  || ''; // Col D
    const nicknameVal   = borrowRow[4]  || ''; // Col E
    const sidVal        = borrowRow[5]  || ''; // Col F
    const deptVal       = borrowRow[6]  || ''; // Col G
    const contactVal    = borrowRow[7]  || ''; // Col H
    const itemsVal      = borrowRow[8]  || ''; // Col I
    const borrowDateVal = borrowRow[9]  || ''; // Col J

    // ตรวจสอบ studentId ให้ตรงกัน
    const cleanSid = (payload.studentId || '').toString().replace(/\D/g, '');
    const sheetSid = (sidVal || '').toString().replace(/\D/g, '');
    if (cleanSid !== sheetSid) {
      return { success: false, message: 'รหัสนิสิตไม่ตรงกับรายการนี้' };
    }

    // --- อัปโหลดรูปภาพ Base64 ลง Drive (พร้อมระบบป้องกัน Error สิทธิ์โฟลเดอร์) ---
    let imageUrl = '';
    if (payload.imageBase64 && payload.imageMimeType) {
      try {
        const decoded  = Utilities.base64Decode(payload.imageBase64);
        const blob     = Utilities.newBlob(decoded, payload.imageMimeType, 'return_' + Date.now() + '.jpg');
        
        let folder;
        try {
          folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        } catch (eFolder) {
          // ถ้าเข้าถึง Folder ID ไม่ได้ (เช่น สิทธิ์ไม่พอ หรือ ID ผิด) ให้เซฟลง Root Drive
          Logger.log('Cannot access folder ID, fallback to root: ' + eFolder.message);
          folder = DriveApp.getRootFolder();
        }

        const file = folder.createFile(blob);

        // พยายามเปิดแชร์ไฟล์ (หากติด Policy ขององค์กร จะไม่ทำให้ระบบล่ม)
        try {
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (eShare) {
          Logger.log('Set sharing skipped: ' + eShare.message);
        }

        imageUrl = 'https://lh3.googleusercontent.com/d/' + file.getId();
      } catch (eUpload) {
        Logger.log('Image upload error: ' + eUpload.message);
        imageUrl = '';
      }
    }

    // --- Append ลงแท็บ "คืนอุปกรณ์" ---
    const timestamp = new Date();
    // จัดรูปแบบวันที่และเวลา เช่น 24/8/2026, 21:21:00
    const returnDateTimeFormatted = Utilities.formatDate(timestamp, 'Asia/Bangkok', 'd/M/yyyy, HH:mm:ss');

    // Mapping ตาม Col A–M (idx 0–12)
    const newRow = [
      timestamp,                // Col A (idx 0) – ประทับเวลา
      emailVal,                 // Col B (idx 1) – อีเมล
      nameVal,                  // Col C (idx 2) – ชื่อ-นามสกุล
      nicknameVal,              // Col D (idx 3) – ชื่อเล่น
      sidVal,                   // Col E (idx 4) – รหัสนิสิต
      deptVal,                  // Col F (idx 5) – ฝ่าย
      contactVal,               // Col G (idx 6) – ช่องทางติดต่อ
      itemsVal,                 // Col H (idx 7) – รายชื่ออุปกรณ์
      borrowDateVal,            // Col I (idx 8) – วันที่ยืม
      imageUrl,                 // Col J (idx 9) – รูปอุปกรณ์ที่คืน
      payload.toldWhom || '',   // Col K (idx 10) – ได้บอกใครบ้าง
      payload.roomOpener || '', // Col L (idx 11) – ใครเป็นคนเปิดห้องให้
      returnDateTimeFormatted   // Col M (idx 12) – วันที่คืน เช่น 24/8/2026, 21:21:00
    ];
    returnSheet.appendRow(newRow);

    // --- อัปเดตสถานะเป็น "คืนแล้ว" ในแท็บยืม Col O (col 15) ---
    borrowSheet.getRange(rowIdx, 15).setValue('คืนแล้ว');

    return { success: true, message: 'คืนอุปกรณ์สำเร็จเรียบร้อยแล้ว 🎉' };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + err.message };
  }
}

/* ──────────────────────────────────────────────
 * Helper – ตอบกลับ JSON สำหรับ REST API
 * ────────────────────────────────────────────── */
function createJsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ──────────────────────────────────────────────
 * Helper – จัดรูปแบบวันที่
 * ────────────────────────────────────────────── */
function formatDate_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Bangkok', 'dd/MM/yyyy');
  }
  return val.toString();
}

/* ──────────────────────────────────────────────
 * Helper – แปลง Drive URL ให้โหลดรูปได้จริง
 * ────────────────────────────────────────────── */
function convertDriveUrl_(url) {
  if (!url) return '';
  url = url.toString().trim();

  if (url.indexOf('lh3.googleusercontent.com') !== -1) return url;

  var fileId = '';
  var match1 = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match1) { fileId = match1[1]; }

  if (!fileId) {
    var match2 = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (match2) { fileId = match2[1]; }
  }

  if (fileId) {
    return 'https://lh3.googleusercontent.com/d/' + fileId;
  }

  return url;
}
