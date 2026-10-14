# Odoo Reference (Community only)

แผนที่ระบบ Odoo Community ใช้เป็นตำราอ้างอิงสำหรับ gap analysis ERP ของ Laotai
ใช้เฉพาะ Odoo Community (LGPL-3) — เอาแนวคิด ไม่ก๊อปโค้ด

- `Odoo20_Module_Map.xlsx` — สรุปทุก module ของ Odoo 20.0: กลุ่ม, การเชื่อมกัน, App หลัก (ภาษาไทย), สิ่งที่ Community ไม่มี
- `odoo20_modules.json` — ข้อมูลดิบที่สแกนจาก source
- `scripts/` — สคริปต์สร้างไฟล์ใหม่ เมื่อ Odoo ออกเวอร์ชันใหม่

## สร้างใหม่ (เช่น Odoo 21.0)
```bash
git clone --depth 1 --branch 21.0 --filter=blob:none --no-checkout https://github.com/odoo/odoo.git odoo
cd odoo && git sparse-checkout init --no-cone
printf '%s\n' '/addons/*/__manifest__.py' '/addons/*/models/*.py' '/addons/*/views/*.xml' '/addons/*/wizard/*.py' > .git/info/sparse-checkout
git checkout 21.0 && cd ..
python3 scripts/scan.py odoo modules.json
python3 scripts/build.py modules.json Odoo21_Module_Map.xlsx scripts
```
คำอธิบายภาษาไทยอยู่ใน `scripts/content.py` (App ใหม่ต้องเพิ่มคำอธิบายเอง)
