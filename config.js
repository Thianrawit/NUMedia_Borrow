/**
 * ============================================================
 * NU Media Club - Web App Configuration
 * ============================================================
 * ใส่ URL ที่ได้จากการ Deploy Google Apps Script Web App ที่นี่
 * ตัวอย่าง: "https://script.google.com/macros/s/AKfycbyxxxxxxx/exec"
 */

const CONFIG = {
  // นำ Web App URL ที่ได้จากขั้นตอน Deploy ใน Google Apps Script มาใส่ตรงนี้
  API_URL: "https://script.google.com/macros/s/AKfycbz_placeholder/exec"
};

// API Helper Functions
async function apiGet(action, params = {}) {
  const queryParams = new URLSearchParams({ action, ...params }).toString();
  const url = `${CONFIG.API_URL}?${queryParams}`;
  const response = await fetch(url, { method: "GET", redirect: "follow" });
  return await response.json();
}

async function apiPost(action, data = {}) {
  const response = await fetch(CONFIG.API_URL, {
    method: "POST",
    redirect: "follow",
    body: JSON.stringify({ action, ...data })
  });
  return await response.json();
}
