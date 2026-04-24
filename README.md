# SRI BROWN - B2B Subscription Backend

ระบบ Backend สำหรับจัดการโควตากาแฟ (B2B Coffee Subscription) พัฒนาด้วย Google Apps Script

**Repository URL:** [https://github.com/Issarapong-w/Sribrownsubscription](https://github.com/Issarapong-w/Sribrownsubscription)

## คุณสมบัติ (Features)
- **User Management:** ระบบสมัครสมาชิกและล็อกอินสำหรับ Partner
- **Quota Allocation:** จัดการจำนวนเมล็ดกาแฟที่ได้รับตามแพ็กเกจ
- **Withdrawal System:** ระบบเบิกเมล็ดกาแฟพร้อมการตัดยอดโควตาอัตโนมัติ
- **Admin Dashboard:** สำหรับจัดการการชำระเงินและสถานะการจัดส่ง

## การติดตั้ง (Setup)
1. ใช้ [clasp](https://github.com/google/clasp) ในการจัดการโค้ด
2. เชื่อมโยงกับ Google Sheet โดยระบุ `SHEET_ID` ในไฟล์ `รหัส.js`
3. โครงสร้าง Google Sheet ต้องมี Sheet ชื่อ: `Users`, `Subscriptions`, `Withdrawals`, และ `Packages`

## เทคโนโลยีที่ใช้
- Google Apps Script (V8 Runtime)
- Google Sheets API
- Clasp for Version Control

---
*Developed for SRI BROWN Coffee Roaster*