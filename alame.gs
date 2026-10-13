// ========================================================
// ระบบแจ้งเตือนยืม-คืนอุปกรณ์ ชมรม NU Media (alame.gs)
// ตรวจสอบทั้ง "ยืมอุปกรณ์" (V1) และ "ยืมอุปกรณ์_v2" (V2)
// เทียบประวัติการคืนกับ "คืนอุปกรณ์" พร้อมส่งอีเมลแจ้งเตือน
// ========================================================

const SPREADSHEET_ID = "1i5niBkkHdqmrIFair6DANcdAfwGhpHjyR2szQz_tqKE";

const SH_BORROW_V1 = "ยืมอุปกรณ์";
const SH_BORROW_V2 = "ยืมอุปกรณ์_v2";
const SH_RETURN = "คืนอุปกรณ์";
const SH_CATALOG = "คลังอุปกรณ์";

/**
 * ดึง Spreadsheet Object (รองรับทั้งเปิดตรงจาก ID หรือเปิดแบบ Container-bound)
 */
function getSpreadsheet_() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getId()) return active;
  } catch (e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * ฟังก์ชันหลักที่ใช้ตั้ง Trigger เช็คและส่งอีเมลแจ้งเตือนทุกเช้า 8:00 น.
 * สามารถกด Run ทดสอบได้ทันทีจาก Apps Script Editor
 */
function checkAndSendReminders() {
  const ss = getSpreadsheet_();
  const s1 = ss.getSheetByName(SH_BORROW_V1); // "ยืมอุปกรณ์" (ระบบเดิม)
  const s2 = ss.getSheetByName(SH_BORROW_V2); // "ยืมอุปกรณ์_v2" (ระบบใหม่)
  const returnSheet = ss.getSheetByName(SH_RETURN); // "คืนอุปกรณ์"

  if (!returnSheet) {
    Logger.log("❌ ไม่พบแท็บ 'คืนอุปกรณ์'");
    return { success: false, message: "ไม่พบแท็บ 'คืนอุปกรณ์'" };
  }

  const returnData = returnSheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const timeZone = Session.getScriptTimeZone() || "Asia/Bangkok";
  const todayStr = Utilities.formatDate(today, timeZone, "yyyy-MM-dd");

  // เก็บ Index แถวคืนที่ถูกจับคู่แล้ว ป้องกันการคืนซ้ำ (1-to-1 Match)
  const usedReturnIndexes = new Set();
  let totalChecked = 0;
  let totalReturned = 0;
  let totalAlerted = 0;

  // 1. ตรวจสอบแท็บ "ยืมอุปกรณ์" (V1 ดั้งเดิม)
  if (s1) {
    const res1 = processBorrowSheet_(s1, returnData, today, todayStr, usedReturnIndexes, {
      version: "V1",
      emailColIdx: 1,      // Col B (idx 1): อีเมล
      nameColIdx: 3,       // Col D (idx 3): ชื่อ-นามสกุล
      nickColIdx: 4,       // Col E (idx 4): ชื่อเล่น
      sidColIdx: 5,        // Col F (idx 5): รหัสนิสิต
      itemsColIdx: 8,      // Col I (idx 8): รายการอุปกรณ์
      borrowDateColIdx: 9, // Col J (idx 9): วันที่ยืม
      imgColIdx: 10,       // Col K (idx 10): ลิงก์รูปภาพ
      dueDateColIdx: 13,   // Col N (idx 13): กำหนดคืน
      statusCol: 15,       // Col O (Col 15): สถานะ
      lastAlertCol: 16     // Col P (Col 16): แจ้งเตือนล่าสุด
    });
    totalChecked += res1.checked;
    totalReturned += res1.returned;
    totalAlerted += res1.alerted;
  } else {
    Logger.log("⚠️ ไม่พบแท็บ '" + SH_BORROW_V1 + "'");
  }

  // 2. ตรวจสอบแท็บ "ยืมอุปกรณ์_v2" (V2 ระบบปัจจุบัน)
  if (s2) {
    const res2 = processBorrowSheet_(s2, returnData, today, todayStr, usedReturnIndexes, {
      version: "V2",
      emailColIdx: 1,      // Col B (idx 1): อีเมล
      nameColIdx: 3,       // Col D (idx 3): ชื่อ-นามสกุล
      nickColIdx: 4,       // Col E (idx 4): ชื่อเล่น
      sidColIdx: 5,        // Col F (idx 5): รหัสนิสิต
      itemsColIdx: 8,      // Col I (idx 8): รายการอุปกรณ์
      borrowDateColIdx: 9, // Col J (idx 9): วันที่ยืม
      imgColIdx: 10,       // Col K (idx 10): ลิงก์รูปภาพ
      dueDateColIdx: 15,   // Col P (idx 15): กำหนดคืน
      statusCol: 17,       // Col Q (Col 17): สถานะ
      lastAlertCol: 18     // Col R (Col 18): แจ้งเตือนล่าสุด
    });
    totalChecked += res2.checked;
    totalReturned += res2.returned;
    totalAlerted += res2.alerted;
  } else {
    Logger.log("⚠️ ไม่พบแท็บ '" + SH_BORROW_V2 + "'");
  }

  Logger.log(`✅ ตรวจสอบสมบูรณ์: ตรวจสอบ ${totalChecked} รายการ, คืนแล้ว ${totalReturned} รายการ, แจ้งเตือน ${totalAlerted} รายการ`);
  return {
    success: true,
    message: `ตรวจสอบ ${totalChecked} รายการ, คืนแล้ว ${totalReturned} รายการ, ส่งแจ้งเตือน ${totalAlerted} รายการ`,
    totalChecked: totalChecked,
    totalReturned: totalReturned,
    totalAlerted: totalAlerted
  };
}

/**
 * ประมวลผลรายการยืมในแต่ละ Sheet
 */
function processBorrowSheet_(sheet, returnData, today, todayStr, usedReturnIndexes, config) {
  const data = sheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone() || "Asia/Bangkok";
  let checked = 0;
  let returned = 0;
  let alerted = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const email = String(row[config.emailColIdx] || "").trim();
    const borrowerName = String(row[config.nameColIdx] || "").trim();
    const nickname = String(row[config.nickColIdx] || "").trim();
    const studentId = String(row[config.sidColIdx] || "").trim();
    const items = String(row[config.itemsColIdx] || "").trim();
    const borrowDateRaw = row[config.borrowDateColIdx];
    const imageUrl = String(row[config.imgColIdx] || "").trim();
    const dueDateRaw = row[config.dueDateColIdx];
    let status = String(row[config.statusCol - 1] || "").trim();
    const lastAlert = row[config.lastAlertCol - 1];

    if (!items && !borrowerName) continue;
    checked++;

    const cleanBorrowId = cleanStudentId(studentId);
    const borrowDateFormatted = formatDate(borrowDateRaw);

    // 1. ตรวจสอบกับประวัติในแท็บ "คืนอุปกรณ์"
    let isReturned = false;
    for (let r = 1; r < returnData.length; r++) {
      if (usedReturnIndexes.has(r)) continue;

      const retRow = returnData[r];
      const retStudentId = String(retRow[4] || "").trim(); // Col E: รหัสนิสิต
      const retItems = String(retRow[7] || "").trim();     // Col H: รายการอุปกรณ์ที่คืน
      const retBorrowDateRaw = retRow[8];                  // Col I: วันที่ยืม

      const cleanRetId = cleanStudentId(retStudentId);
      const retBorrowDateFormatted = formatDate(retBorrowDateRaw);

      // เงื่อนไข 1: รหัสนิสิตต้องตรงกัน
      const isIdMatch = cleanBorrowId && cleanRetId && cleanBorrowId === cleanRetId;

      // เงื่อนไข 2: วันที่ยืมต้องตรงกัน (ตัดเวลาออก เทียบแค่วัน/เดือน/ปี)
      const isDateMatch = borrowDateFormatted && retBorrowDateFormatted
        ? borrowDateFormatted === retBorrowDateFormatted
        : true;

      // เงื่อนไข 3: เช็ครายการอุปกรณ์ที่มีความคล้ายกัน (Fuzzy Bigram)
      const isItemMatch = compareItemList(items, retItems);

      if (isIdMatch && isDateMatch && isItemMatch) {
        isReturned = true;
        usedReturnIndexes.add(r);
        break;
      }
    }

    // 2. ถ้ามีประวัติการคืนแล้ว
    if (isReturned) {
      if (status !== "คืนแล้ว") {
        sheet.getRange(i + 1, config.statusCol).setValue("คืนแล้ว");
        returned++;

        // ปลดสถานะใน "คลังอุปกรณ์" กลับเป็น "พร้อมใช้งาน" อัตโนมัติ
        try {
          const ss = getSpreadsheet_();
          const catSheet = ss.getSheetByName(SH_CATALOG);
          if (catSheet && items) {
            const catData = catSheet.getDataRange().getValues();
            const returnedList = items.split(/[\n,]/).map(function(s){ return s.trim().toLowerCase(); });
            for (let c = 1; c < catData.length; c++) {
              const eqName = (catData[c][1] || "").toString().trim().toLowerCase();
              if (returnedList.indexOf(eqName) !== -1) {
                catSheet.getRange(c + 1, 5).setValue("พร้อมใช้งาน");
              }
            }
          }
        } catch (eCat) {
          Logger.log("Reset catalog status error: " + eCat.message);
        }
      }
      continue;
    }

    // 3. ถ้ายังไม่มีประวัติคืน -> ตรวจสอบกำหนดคืนและส่งแจ้งเตือน
    if (dueDateRaw) {
      const dueDate = new Date(dueDateRaw);
      dueDate.setHours(0, 0, 0, 0);

      if (today < dueDate) {
        if (status !== "กำลังยืม") {
          sheet.getRange(i + 1, config.statusCol).setValue("กำลังยืม");
        }
      } else {
        const isOverdue = today > dueDate;
        const targetStatus = isOverdue ? "เลยกำหนดคืน" : "ถึงกำหนดคืน";

        if (status !== targetStatus) {
          sheet.getRange(i + 1, config.statusCol).setValue(targetStatus);
        }

        const lastAlertStr = lastAlert
          ? Utilities.formatDate(new Date(lastAlert), timeZone, "yyyy-MM-dd")
          : "";

        // ส่งแจ้งเตือนถ้ายังไม่ได้ส่งวันนี้
        if (lastAlertStr !== todayStr && email) {
          let imageBlob = null;
          if (imageUrl) {
            imageBlob = getDriveImageBlob(imageUrl);
          }

          sendEmailReminder(
            email,
            borrowerName,
            nickname,
            items,
            dueDate,
            isOverdue,
            imageBlob
          );

          sheet.getRange(i + 1, config.lastAlertCol).setValue(new Date());
          alerted++;
        }
      }
    }
  }

  return { checked: checked, returned: returned, alerted: alerted };
}

/**
 * ตั้งค่า Trigger เช็คสถานะและส่งเมลทุกวันช่วง 8:00 - 9:00 น.
 * รันฟังก์ชันนี้ 1 ครั้งใน Apps Script เพื่อเปิดใช้งาน
 */
function setupDaily8AmTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const fn = triggers[i].getHandlerFunction();
    if (fn === "checkAndSendReminders" || fn === "checkAndNotifyDailyBorrowStatus") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("checkAndSendReminders")
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone("Asia/Bangkok")
    .create();

  Logger.log("✅ ตั้ง Trigger เช็คสถานะ 8:00 น. เรียบร้อยแล้ว!");
  return { success: true, message: "ตั้ง Trigger เช็คสถานะ 8:00 น. เรียบร้อยแล้ว" };
}

// ----------------------------------------------------
// ระบบเทียบรายการอุปกรณ์แบบ Fuzzy Multi-line Match
// ----------------------------------------------------
function compareItemList(borrowItemsText, returnItemsText) {
  if (!borrowItemsText || !returnItemsText) return true;

  const getLines = (text) =>
    String(text)
      .split(/\r?\n|,/)
      .map((line) =>
        line
          .replace(/^[-*•\s\d]+/, "")
          .replace(/[\s\-_.,/\\()]/g, "")
          .toLowerCase()
          .trim()
      )
      .filter((line) => line.length > 1);

  const borrowLines = getLines(borrowItemsText);
  const returnLines = getLines(returnItemsText);

  if (borrowLines.length === 0 || returnLines.length === 0) return true;

  let matchCount = 0;
  for (let bLine of borrowLines) {
    const isMatched = returnLines.some((rLine) => {
      if (bLine === rLine || bLine.includes(rLine) || rLine.includes(bLine)) return true;
      return getSimilarityScore(bLine, rLine) >= 0.4;
    });
    if (isMatched) matchCount++;
  }

  return matchCount / borrowLines.length >= 0.5;
}

function getSimilarityScore(str1, str2) {
  if (str1.length < 2 || str2.length < 2) return 0;
  const getBigrams = (str) => {
    const bigrams = new Map();
    for (let i = 0; i < str.length - 1; i++) {
      const bigram = str.substring(i, i + 2);
      bigrams.set(bigram, (bigrams.get(bigram) || 0) + 1);
    }
    return bigrams;
  };

  const bg1 = getBigrams(str1);
  const bg2 = getBigrams(str2);
  let intersection = 0;

  bg1.forEach((count1, bigram) => {
    if (bg2.has(bigram)) {
      intersection += Math.min(count1, bg2.get(bigram));
    }
  });

  return (2.0 * intersection) / (str1.length - 1 + (str2.length - 1));
}

function formatDate(dateVal) {
  if (!dateVal) return "";
  try {
    const d = new Date(dateVal);
    if (isNaN(d.getTime())) return String(dateVal).split(" ")[0].trim();
    return Utilities.formatDate(d, Session.getScriptTimeZone() || "Asia/Bangkok", "yyyy-MM-dd");
  } catch (e) {
    return String(dateVal).trim();
  }
}

function cleanStudentId(id) {
  return String(id || "")
    .replace(/\D/g, "")
    .trim();
}

function getDriveImageBlob(driveUrl) {
  try {
    if (!driveUrl) return null;
    const match = driveUrl.toString().match(/[-\w]{25,}/);
    if (match) {
      const fileId = match[0];
      const file = DriveApp.getFileById(fileId);
      return file.getBlob();
    }
  } catch (e) {
    Logger.log("⚠️ ดึงรูปไม่สำเร็จ: " + e.message);
  }
  return null;
}

// ----------------------------------------------------
// ฟังก์ชันส่ง Email แจ้งเตือน (HTML Template + Inline Image)
// ----------------------------------------------------
function sendEmailReminder(
  email,
  name,
  nickname,
  items,
  dueDate,
  isOverdue,
  imageBlob
) {
  const dueDateFormatted = Utilities.formatDate(dueDate, "GMT+7", "dd/MM/yyyy");
  const displayName = nickname ? `${name} (${nickname})` : name;

  const subject = isOverdue
    ? `[แจ้งเตือนด่วน!] เลยกำหนดคืนอุปกรณ์ชมรม NU Media - ${items}`
    : `[แจ้งเตือน] ครบรอบกำหนดคืนอุปกรณ์ชมรม NU Media วันนี้ - ${items}`;

  let htmlBody = `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: auto; border: 1px solid #e0e0e0; border-radius: 8px; padding: 20px;">
      <h2 style="color: ${isOverdue ? "#d32f2f" : "#1976d2"};">
        ${isOverdue ? "⚠️ แจ้งเตือนเลยกำหนดคืนอุปกรณ์" : "📦 แจ้งเตือนกำหนดคืนอุปกรณ์วันนี้"}
      </h2>
      <p>สวัสดีคุณ <b>${displayName}</b>,</p>
      <p>ระบบตรวจสอบพบว่าคุณมีรายการยืมอุปกรณ์ของ <b>ชมรม NU Media</b> ที่ถึงกำหนดส่งคืน ดังนี้:</p>
      
      <table style="width: 100%; border-collapse: collapse; margin: 15px 0;">
        <tr style="background-color: #f8f9fa;">
          <td style="padding: 10px; border: 1px solid #ddd; width: 35%;"><b>รายการอุปกรณ์:</b></td>
          <td style="padding: 10px; border: 1px solid #ddd;">${items.replace(/\n/g, "<br>")}</td>
        </tr>
        <tr>
          <td style="padding: 10px; border: 1px solid #ddd;"><b>กำหนดคืน:</b></td>
          <td style="padding: 10px; border: 1px solid #ddd; color: ${isOverdue ? "#d32f2f" : "#2e7d32"};"><b>${dueDateFormatted}</b></td>
        </tr>
      </table>
  `;

  let inlineImagesObj = {};
  let attachmentsArr = [];

  if (imageBlob) {
    inlineImagesObj["itemPic"] = imageBlob;
    attachmentsArr.push(imageBlob);
    htmlBody += `
      <p><b>รูปภาพอุปกรณ์/หลักฐานตอนยืม:</b></p>
      <div style="text-align: center; margin: 15px 0;">
        <img src="cid:itemPic" style="max-width: 100%; max-height: 350px; border-radius: 6px; border: 1px solid #ccc;" />
      </div>
    `;
  }

  htmlBody += `
      <p>${
        isOverdue
          ? '<b style="color: #d32f2f;">ขณะนี้เลยกำหนดส่งคืนแล้ว</b> กรุณานำอุปกรณ์มาคืนที่ห้องชมรมโดยด่วนที่สุด และกรอกแบบฟอร์มคืนของเพื่อปิดรายการ'
          : "กรุณานำอุปกรณ์มาส่งคืนที่ห้องชมรมภายในวันนี้ และกรอกแบบฟอร์มคืนของด้วยนะครับ"
      }
      </p>

      <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;">
      <p style="font-size: 12px; color: #888;">อีเมลนี้เป็นระบบแจ้งเตือนอัตโนมัติจากฝ่ายพัสดุ ชมรม NU Media <b>ห้ามตอบกลับข้อความนี้</b></p>
    </div>
  `;

  try {
    GmailApp.sendEmail(email, subject, "", {
      htmlBody: htmlBody,
      name: "ระบบพัสดุ NU Media Club",
      inlineImages: inlineImagesObj,
      attachments: attachmentsArr,
    });
  } catch (eGmail) {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      name: "ระบบพัสดุ NU Media Club",
      inlineImages: inlineImagesObj,
      attachments: attachmentsArr,
    });
  }
}
