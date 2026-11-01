# ระบบรายงาน Pet Shop R9

ระบบรายงานนัดหมาย / ยอดขาย / สต็อก / ตารางเวร บน Google Sheets + Apps Script

## โครงสร้าง

| ไฟล์ | ใช้ทำอะไร |
|---|---|
| `apps-script/Code.gs` | โค้ดฝั่งเซิร์ฟเวอร์ — วางในไฟล์ `รหัส.gs` ของ Apps Script |
| `apps-script/Page.html` | หน้าเว็บ — วางในไฟล์ `Page.html` ของ Apps Script |
| `apps-script/appsscript.json` | ตั้งค่าโปรเจกต์ (ใช้กับ clasp เท่านั้น ไม่ต้องวางเอง) |
| `index.html` | ลิงก์สั้นผ่าน GitHub Pages — เปิดแล้วพาไปหน้าเว็บแอป |

## อัปเดตโค้ด (ลิงก์เดิม)

1. คัดลอก `apps-script/Code.gs` ไปวางทับ `รหัส.gs` และ `apps-script/Page.html` ไปวางทับ `Page.html`
2. Apps Script → **การทำให้ใช้งานได้ → จัดการการทำให้ใช้งานได้ → ✏️ → เวอร์ชันใหม่ → ทำให้ใช้งานได้**
3. ห้ามกด "การทำให้ใช้งานได้รายการใหม่" (จะได้ลิงก์ใหม่ QR ตู้จะใช้ไม่ได้)

## ลิงก์สั้น (GitHub Pages)

1. แก้ `index.html` บรรทัด `WEBAPP_URL` ใส่ลิงก์เว็บแอป (ลงท้าย `/exec`)
2. GitHub → Settings → Pages → Source: **Deploy from a branch** → Branch: `main` / `(root)` → Save
3. ได้ลิงก์ https://thonglorpetshop.github.io/petshopr9/
4. (ถ้าต้องการให้ QR ตู้ใช้ลิงก์สั้น) ใส่ลิงก์นี้ในชีต **ตั้งค่า** ช่อง **ลิงก์หน้าเว็บ (QR)** แล้วพิมพ์ QR ใหม่

> ⚠️ ห้ามใส่รหัสผ่าน หรือข้อมูลลูกค้าลงใน repo นี้ — ถ้า repo เป็น Public ทุกคนเห็นได้
