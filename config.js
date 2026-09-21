// ตั้งค่าแอป — แก้ไฟล์นี้ไฟล์เดียวพอ
window.APP_CONFIG = {
  // Web app URL จาก Apps Script (ลงท้ายด้วย /exec)
  apiUrl: "https://script.google.com/macros/s/AKfycbyar-7X1oNZZASL2lDILoEQiaLN_Guud7wI4QMEYqIYaoQJix5PAFT6H4IEPjrcFaUfyw/exec",

  // ตัวเลือกผลตรวจ: tone = ok (เขียว), warn (เหลือง), bad (แดง)
  statuses: [
    { label: "ตรวจแล้ว", tone: "ok" },
    { label: "เสื่อมสภาพ", tone: "warn" },
    { label: "ชำรุด", tone: "warn" },
    { label: "ไม่พบ", tone: "bad" }
  ]
};
