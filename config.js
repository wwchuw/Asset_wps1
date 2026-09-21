// ตั้งค่าแอป — แก้ไฟล์นี้ไฟล์เดียวพอ
window.APP_CONFIG = {
  // Application (client) ID จาก Azure App registrations
  clientId: "ใส่-CLIENT-ID-ที่นี่",

  // "common" ใช้ได้ทั้งบัญชีส่วนตัวและบัญชีองค์กร
  // ถ้าย้ายไปใช้เฉพาะบัญชีองค์กร เปลี่ยนเป็น "https://login.microsoftonline.com/<tenant-id>"
  authority: "https://login.microsoftonline.com/common",

  // โฟลเดอร์ใน OneDrive (นับจาก root) และชื่อไฟล์ Excel หลัก
  folder: "AssetDB",
  workbook: "AssetDB.xlsx",

  // ตัวเลือกผลตรวจ: tone = ok (เขียว), warn (เหลือง), bad (แดง)
  statuses: [
    { label: "ตรวจแล้ว", tone: "ok" },
    { label: "เสื่อมสภาพ", tone: "warn" },
    { label: "ชำรุด", tone: "warn" },
    { label: "ไม่พบ", tone: "bad" }
  ]
};
