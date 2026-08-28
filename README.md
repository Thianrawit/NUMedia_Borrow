# NU Media Club - ระบบยืม-คืนอุปกรณ์ (GitHub Pages + Google Apps Script)

เว็บแอปพลิเคชันระบบแสดงรายการและคืนอุปกรณ์สำหรับชมรม **NU Media Club** รัน Frontend บน **GitHub Pages** และเชื่อมต่อฐานข้อมูล Google Sheets ผ่าน **Google Apps Script REST API** หมดปัญหาการล็อกอินหลายบัญชีและปัญหา iframe 100%

---

## 🚀 ขั้นตอนการติดตั้งและใช้งาน (Setup Guide)

### 1. ฝั่ง Google Apps Script (Backend API)
1. ไปที่ [script.google.com](https://script.google.com) แล้วสร้างโปรเจกต์ใหม่
2. คัดลอกโค้ดจากไฟล์ `Code.gs` ไปวางในโปรเจกต์
3. กดปุ่มสีน้ำเงินมุมขวาบน **Deploy (การทำให้ใช้งานได้)** > **New deployment (การทำให้ใช้งานได้รายการใหม่)**
4. ตั้งค่าดังนี้:
   * **Select type (เลือกประเภท):** Web app (เว็บแอป)
   * **Execute as (ดำเนินการในฐานะ):** `Me (ฉัน)`
   * **Who has access (ผู้มีสิทธิ์เข้าถึง):** `Anyone (ทุกคน)`
5. กด **Deploy** แล้วให้สิทธิ์การเข้าถึง (Authorize access)
6. **คัดลอก Web App URL** ที่ได้ (ลงท้ายด้วย `/exec`)

---

### 2. ฝั่ง GitHub Pages (Frontend)
1. เปิดไฟล์ `config.js` ในโปรเจกต์นี้
2. นำ Web App URL ที่ได้จากข้อ 1 มาใส่ในตัวแปร `API_URL`:
   ```javascript
   const CONFIG = {
     API_URL: "https://script.google.com/macros/s/AKfycbxxxxxxx/exec"
   };
   ```
3. อัปโหลดไฟล์ทั้งหมดขึ้น GitHub Repository: [Thianrawit/NUMedia_Borrow](https://github.com/Thianrawit/NUMedia_Borrow)
   * `index.html`
   * `return.html`
   * `config.js`
   * `icon.png`
4. ไปที่หน้า **Repository บน GitHub** > **Settings** > **Pages**
5. ตรงหัวข้อ **Branch** ให้เลือก `main` (หรือ `master`) แล้วกด **Save**
6. รอประมาณ 1-2 นาที คุณจะได้ URL เว็บไซต์ เช่น:
   `https://thianrawit.github.io/NUMedia_Borrow/`

---

## 📂 โครงสร้างไฟล์
* `index.html` : หน้ารายการอุปกรณ์ที่กำลังถูกยืม (Dashboard)
* `return.html` : หน้าเข้าสู่ระบบและแบบฟอร์มยืนยันการคืนอุปกรณ์
* `config.js` : ไฟล์ตั้งค่า API URL และฟังก์ชันเรียก API
* `Code.gs` : โค้ด Backend API สำหรับนำไปใส่ใน Google Apps Script
* `icon.png` : ไอคอนและ Favicon ของเว็บไซต์

---

## ✨ ฟีเจอร์เด่น
* 🎨 **Clean & Modern Light Theme:** Nav Bar สีส้ม `#f97316` ตัวหนังสือใหญ่คมชัด อ่านง่ายสบายตา
* 📱 **Mobile-First & Responsive:** ใช้งานได้ลื่นไหลทั้งบนมือถือ แท็บเล็ต และคอมพิวเตอร์
* ⚡ **Client-Side Image Compression:** ย่อขนาดรูปภาพหลักฐานการคืนอัตโนมัติบน Canvas ก่อนส่ง ลดการใช้เน็ตและบันทึกลง Google Drive ได้รวดเร็ว
* 🛡️ **Zero Multi-Account Conflict:** รันบน GitHub Pages จึงหมดปัญหาเซสชัน Google ตีกันหรือปัญหาเปิดบน iframe ไม่ได้ 100%
