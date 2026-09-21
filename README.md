# ตรวจทรัพย์สิน (PWA + OneDrive)

เว็บแอปสำหรับตรวจทรัพย์สินบนมือถือ สแกน QR เดิมที่ติดบนทรัพย์สินได้ทันที ข้อมูลทั้งหมดอยู่ใน OneDrive ของคุณ

## โครงสร้างข้อมูลใน OneDrive

```
AssetDB/
  AssetDB.xlsx          ข้อมูลทรัพย์สิน (แก้ใน Excel ได้ตามปกติ แอปอ่านอย่างเดียว)
  images/<รหัส>.jpg     รูปทรัพย์สิน (แอปเขียนทับเมื่อถ่ายรูปใหม่ OneDrive เก็บเวอร์ชันเก่าไว้ให้)
  audit/<ปีงบ>.json     ผลตรวจที่บันทึกจากแอป (แอปสร้างเอง)
  reports/              ไฟล์ Excel สรุปผลตรวจ (สร้างเมื่อกด "ส่งออกผลตรวจ")
```

เหตุที่ผลตรวจเก็บเป็น JSON ไม่เขียนลง Excel โดยตรง: Microsoft Graph ไม่รองรับการแก้ไข Excel
ที่อยู่ใน OneDrive ส่วนตัว (รองรับเฉพาะ OneDrive for Business/SharePoint) แอปจึงอ่าน Excel
แล้วเขียนผลตรวจแยกไฟล์ และส่งออกเป็น Excel ให้เมื่อต้องการ

## ติดตั้ง (ครั้งเดียว)

### 1. อัปโหลดข้อมูล
แตก `AssetDB.zip` แล้วอัปโหลดโฟลเดอร์ `AssetDB` ไปไว้ที่หน้าแรก (root) ของ OneDrive

### 2. เอาเว็บขึ้น GitHub Pages
1. สร้างบัญชี github.com และสร้าง repository ใหม่ เช่น `asset-audit`
2. อัปโหลดไฟล์ทั้งหมดในโฟลเดอร์นี้ (index.html, app.js, style.css, config.js, sw.js, manifest.webmanifest, icons/)
3. Settings → Pages → Source: Deploy from a branch → Branch: main / root → Save
4. รอสักครู่ จะได้ URL เช่น `https://<ชื่อผู้ใช้>.github.io/asset-audit/`

ในเว็บไม่มีข้อมูลทรัพย์สินหรือรหัสผ่านใดๆ ข้อมูลจริงอยู่ใน OneDrive ที่ต้อง login ก่อนเท่านั้น

### 3. ลงทะเบียนแอปกับ Microsoft
1. เข้า https://portal.azure.com ด้วยบัญชี Microsoft ส่วนตัว
   (ถ้าระบบให้สร้าง Azure account/directory ก่อน ให้สมัครแบบฟรี)
2. Microsoft Entra ID → App registrations → New registration
   - Name: `ตรวจทรัพย์สิน`
   - Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**
   - Redirect URI: เลือก **Single-page application (SPA)** ใส่ URL จากข้อ 2 (มี `/` ปิดท้าย)
3. กด Register แล้วคัดลอก **Application (client) ID**
4. API permissions → Add a permission → Microsoft Graph → Delegated → เพิ่ม `Files.ReadWrite`
   (`User.Read` มีให้อยู่แล้ว)
5. ถ้าจะทดสอบในเครื่องด้วย: Authentication → เพิ่ม SPA redirect URI `http://localhost:8080/`

### 4. ใส่ client ID
แก้ `config.js` บรรทัด `clientId` ใส่ค่าจากข้อ 3 แล้วอัปโหลดขึ้น GitHub อีกครั้ง

### 5. ใช้งานบนมือถือ
เปิด URL ด้วย Chrome (Android) หรือ Safari (iPhone) → เข้าสู่ระบบ → เมนูเบราว์เซอร์ →
"เพิ่มลงในหน้าจอหลัก" แอปจะเปิดแบบเต็มจอเหมือนแอปทั่วไป

## การใช้งาน
- สแกน QR: ใช้กับสติ๊กเกอร์เดิมได้ (หรือพิมพ์รหัสเองก็ได้)
- บันทึกผลตรวจ: เลือกผล ใส่หมายเหตุ ถ่ายรูปใหม่ได้ ระบบบันทึกชื่อผู้ตรวจ เวลา และพิกัด GPS (ถ้าอนุญาต)
- บันทึกและสแกนต่อ: เมื่อเปิดจากการสแกน กดแล้วกลับไปที่กล้องทันที
- ไม่มีเน็ต: ผลตรวจเก็บในเครื่องก่อน แล้วส่งขึ้น OneDrive อัตโนมัติเมื่อต่อเน็ต (รูปใหม่ต้องมีเน็ตตอนกดบันทึก)
- เปลี่ยนปีงบ: ตั้งอัตโนมัติตามปีงบประมาณ (เริ่ม 1 ต.ค.) เลือกปีเก่าเพื่อดูย้อนหลังได้
- ส่งออกผลตรวจ: เมนู ⋯ → ได้ไฟล์ Excel และเก็บสำเนาไว้ใน `AssetDB/reports`

## ทดสอบในเครื่องคอมพิวเตอร์
```
cd โฟลเดอร์นี้
python -m http.server 8080
```
เปิด http://localhost:8080/ (กล้องใช้ได้บน localhost หรือ https เท่านั้น)

## แก้ไขภายหลัง
- เพิ่ม/แก้ทรัพย์สิน: แก้ในชีท `Assets` ของ AssetDB.xlsx แล้วกด "โหลดข้อมูลใหม่" ในแอป (อย่าเปลี่ยนชื่อหัวคอลัมน์)
- เปลี่ยนตัวเลือกผลตรวจ: แก้ `statuses` ใน config.js
- แก้โค้ดแล้วมือถือไม่อัปเดต: เปลี่ยน `VERSION` ใน sw.js

## ข้อจำกัดของบัญชีส่วนตัว และตอนย้ายไปบัญชีองค์กร
- แอปใช้ OneDrive ของคนที่ login จึงเหมาะกับการตรวจคนเดียวในช่วงทดลอง
  ถ้าหลายคนต้องใช้ข้อมูลชุดเดียวกัน ควรย้ายไป OneDrive for Business/SharePoint ขององค์กร
  (แอปมีระบบกันบันทึกทับกันเมื่อหลายคนเขียนไฟล์เดียวกันอยู่แล้ว)
- ย้ายไปบัญชีองค์กร: ให้ IT ลงทะเบียนแอปในองค์กร (ขั้นตอนเดียวกับข้อ 3) แล้วเปลี่ยน
  `clientId` และ `authority` ใน config.js ถ้าเก็บข้อมูลไว้ใน SharePoint site แทน OneDrive
  ต้องแก้ฟังก์ชัน `item()` ใน app.js ให้ชี้ไปที่ drive ของ site นั้น
