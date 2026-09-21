/**
 * ตรวจทรัพย์สิน — API สำหรับ PWA
 * วางไฟล์นี้ใน Google Sheet เดิม: ส่วนขยาย (Extensions) → Apps Script
 * ตั้ง Script property ชื่อ TEAM_KEY แล้ว Deploy เป็น Web app (ดู README)
 */
const CONFIG = {
  SHEET_NAME: 'Sheet1',                           // ชีทข้อมูลทรัพย์สิน
  IMAGES_FOLDER_ID: '',                           // ID โฟลเดอร์รูปใน Drive (แนะนำให้ใส่)
  IMAGES_FOLDER_NAME: 'Database_Asset_ssn1_Images', // ใช้เมื่อไม่ได้ใส่ ID
  AUDIT_PREFIX: 'ตรวจทรัพสินย์งบ_',               // หัวคอลัมน์ผลตรวจรายปี (ตามของเดิม)
  LOG_SHEET: 'AuditLog',                          // ประวัติการตรวจทุกครั้ง (สร้างให้อัตโนมัติ)
  TIMEZONE: 'Asia/Bangkok'
};

const COL = {
  id: 'รหัสสินทรัพย์', name: 'ชื่อสินทรัพย์',
  typeCode: 'รหัสประเภททรัพย์สิน', typeName: 'ชื่อประเภททรัพย์สิน',
  qty: 'จำนวน', owner: 'ผู้ครอบครอง', room: 'สถานที่', building: 'อาคาร',
  image: 'รูปภาพ', location: 'Location', status: 'สถานะ', note: 'หมายเหตุ'
};
const LOG_HEADERS = ['เวลาบันทึก', 'ปีงบ', 'รหัสสินทรัพย์', 'ผลตรวจ', 'หมายเหตุ', 'ผู้ตรวจ', 'Lat', 'Lng', 'เวลาตรวจ(ISO)'];

function doGet(e) { return respond_(handle_((e && e.parameter) || {})); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { /* ignore */ }
  return respond_(handle_(body));
}

function handle_(req) {
  try {
    checkKey_(req.key);
    let out;
    switch (req.action) {
      case 'ping': out = {}; break;
      case 'load': out = load_(); break;
      case 'audit': out = saveAudits_(req.records || []); break;
      case 'photo': out = savePhoto_(req.id, req.data); break;
      default: throw new Error('ไม่รู้จักคำสั่ง ' + req.action);
    }
    out.ok = true;
    return out;
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}
function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function checkKey_(k) {
  const key = PropertiesService.getScriptProperties().getProperty('TEAM_KEY');
  if (!key) throw new Error('ยังไม่ได้ตั้ง TEAM_KEY ใน Script properties');
  if (String(k || '') !== key) { Utilities.sleep(800); throw new Error('รหัสทีมไม่ถูกต้อง'); }
}

/* ---------- อ่านข้อมูล ---------- */
function mainSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error('ไม่พบชีท ' + CONFIG.SHEET_NAME);
  return sh;
}
const text_ = (v) => v instanceof Date ? v.toISOString() : String(v == null ? '' : v).trim();

function load_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const values = mainSheet_().getDataRange().getValues();
  const head = values.shift().map(String);
  const ix = {};
  Object.keys(COL).forEach((k) => { ix[k] = head.indexOf(COL[k]); });
  if (ix.id < 0) throw new Error('ไม่พบคอลัมน์ ' + COL.id);
  const get = (row, k) => ix[k] >= 0 ? row[ix[k]] : '';

  const auditCols = [];
  head.forEach((h, i) => { if (h.indexOf(CONFIG.AUDIT_PREFIX) === 0) auditCols.push({ year: h.slice(CONFIG.AUDIT_PREFIX.length), i }); });

  const imgs = imageMap_();
  const assets = [], records = {};
  values.forEach((r) => {
    const id = text_(get(r, 'id'));
    if (!id) return;
    const ll = String(get(r, 'location')).split(',').map(parseFloat);
    const imgName = String(get(r, 'image')).split('/').pop();
    assets.push({
      id, name: text_(get(r, 'name')), typeCode: text_(get(r, 'typeCode')), typeName: text_(get(r, 'typeName')),
      qty: text_(get(r, 'qty')), owner: text_(get(r, 'owner')), room: text_(get(r, 'room')),
      building: text_(get(r, 'building')), status: text_(get(r, 'status')), note: text_(get(r, 'note')),
      lat: isFinite(ll[0]) ? ll[0] : null, lng: isFinite(ll[1]) ? ll[1] : null,
      imageId: imgs[imgName] || null
    });
    auditCols.forEach((c) => {
      const v = text_(r[c.i]);
      if (v) (records[c.year] = records[c.year] || {})[id] = { status: v, note: '', by: '', at: '' };
    });
  });

  // ประวัติใน AuditLog มีรายละเอียดมากกว่า ใช้ทับค่าจากคอลัมน์
  const log = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (log && log.getLastRow() > 1) {
    log.getRange(2, 1, log.getLastRow() - 1, LOG_HEADERS.length).getValues().forEach((r) => {
      const year = text_(r[1]), id = text_(r[2]), at = text_(r[8]) || text_(r[0]);
      if (!year || !id) return;
      const y = (records[year] = records[year] || {});
      if (!y[id] || !y[id].at || y[id].at <= at) {
        y[id] = { status: text_(r[3]), note: text_(r[4]), by: text_(r[5]), at };
      }
    });
  }
  return { assets, records, sheetUrl: ss.getUrl() };
}

function imagesFolder_() {
  if (CONFIG.IMAGES_FOLDER_ID) return DriveApp.getFolderById(CONFIG.IMAGES_FOLDER_ID);
  const it = DriveApp.getFoldersByName(CONFIG.IMAGES_FOLDER_NAME);
  if (!it.hasNext()) throw new Error('ไม่พบโฟลเดอร์รูป ' + CONFIG.IMAGES_FOLDER_NAME);
  return it.next();
}
function imageMap_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('imgmap');
  if (hit) return JSON.parse(hit);
  const map = {};
  const files = imagesFolder_().getFiles();
  while (files.hasNext()) { const f = files.next(); map[f.getName()] = f.getId(); }
  try { cache.put('imgmap', JSON.stringify(map), 600); } catch (e) { /* ใหญ่เกิน cache ก็ไม่เป็นไร */ }
  return map;
}

/* ---------- บันทึกผลตรวจ ---------- */
function logSheet_(ss) {
  let log = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(CONFIG.LOG_SHEET);
    log.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold');
    log.setFrozenRows(1);
    log.getRange('B:C').setNumberFormat('@');
    log.getRange('I:I').setNumberFormat('@');
  }
  return log;
}

function saveAudits_(records) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = logSheet_(ss);
    const seen = {};
    if (log.getLastRow() > 1) {
      log.getRange(2, 3, log.getLastRow() - 1, 7).getValues()
        .forEach((r) => { seen[text_(r[0]) + '|' + text_(r[6])] = true; });
    }
    const sh = mainSheet_();
    const values = sh.getDataRange().getValues();
    const head = values[0].map(String);
    const idIx = head.indexOf(COL.id);
    const rowOf = {};
    for (let i = 1; i < values.length; i++) rowOf[text_(values[i][idIx])] = i + 1;

    const rows = [], saved = [];
    records.forEach((rec) => {
      const id = text_(rec.id), year = text_(rec.year), at = text_(rec.at);
      if (!id || !year || !rec.status) return;
      const k = id + '|' + at;
      saved.push(k);
      if (seen[k]) return; // ส่งซ้ำ (เช่น เน็ตหลุดตอนรอคำตอบ)
      seen[k] = true;
      rows.push([new Date(), year, id, String(rec.status), String(rec.note || ''), String(rec.by || ''),
        rec.lat == null ? '' : rec.lat, rec.lng == null ? '' : rec.lng, at]);

      // อัปเดตคอลัมน์ผลตรวจรายปีในชีทหลัก ให้ดูใน Sheet ได้เหมือนเดิม
      const colName = CONFIG.AUDIT_PREFIX + year;
      let col = head.indexOf(colName);
      if (col < 0) {
        col = head.length;
        if (col + 1 > sh.getMaxColumns()) sh.insertColumnAfter(sh.getMaxColumns());
        sh.getRange(1, col + 1).setValue(colName);
        head.push(colName);
      }
      if (rowOf[id]) sh.getRange(rowOf[id], col + 1).setValue(String(rec.status));
    });
    if (rows.length) log.getRange(log.getLastRow() + 1, 1, rows.length, LOG_HEADERS.length).setValues(rows);
    return { saved };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- บันทึกรูป ---------- */
function savePhoto_(id, data) {
  id = text_(id);
  if (!/^[\w.-]{3,40}$/.test(id)) throw new Error('รหัสสินทรัพย์ไม่ถูกต้อง');
  if (!data) throw new Error('ไม่มีข้อมูลรูป');
  const stamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd-HHmmss');
  const name = id + '.รูปภาพ.' + stamp + '.jpg';
  const folder = imagesFolder_();
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(data), 'image/jpeg', name));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* องค์กรอาจห้าม */ }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = mainSheet_();
    const values = sh.getDataRange().getValues();
    const head = values[0].map(String);
    const idIx = head.indexOf(COL.id), imgIx = head.indexOf(COL.image);
    if (imgIx >= 0) {
      for (let i = 1; i < values.length; i++) {
        if (text_(values[i][idIx]) !== id) continue;
        const old = String(values[i][imgIx]);
        const dir = old.indexOf('/') >= 0 ? old.slice(0, old.lastIndexOf('/')) : folder.getName();
        sh.getRange(i + 1, imgIx + 1).setValue(dir + '/' + name); // รูปเก่ายังอยู่ในโฟลเดอร์
        break;
      }
    }
  } finally {
    lock.releaseLock();
  }
  CacheService.getScriptCache().remove('imgmap');
  return { imageId: file.getId() };
}

/** กดรันฟังก์ชันนี้ครั้งแรกใน editor เพื่อให้สิทธิ์ และเช็คว่าอ่านข้อมูลได้ */
function testSetup() {
  const r = load_();
  Logger.log('ทรัพย์สิน %s รายการ, มีรูป %s รายการ, ปีงบที่มีผลตรวจ: %s',
    r.assets.length, r.assets.filter((a) => a.imageId).length, Object.keys(r.records).join(', '));
}
