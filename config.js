// ตั้งค่าแอป — แก้ไฟล์นี้ไฟล์เดียวพอ
window.APP_CONFIG = {
  // Web app URL จาก Apps Script (ลงท้ายด้วย /exec)
  apiUrl: "ใส่-WEB-APP-URL-ที่นี่",

  // ตัวเลือกผลตรวจ: tone = ok (เขียว), warn (เหลือง), bad (แดง)
  statuses: [
    { label: "ตรวจแล้ว", tone: "ok" },
    { label: "เสื่อมสภาพ", tone: "warn" },
    { label: "ชำรุด", tone: "warn" },
    { label: "ไม่พบ", tone: "bad" }
  ]
};
