# ระบบเสนอโครงร่างงานวิจัย คณะพยาบาลศาสตร์ มหาวิทยาลัยราชภัฏศรีสะเกษ

เว็บไซต์กรอกโครงร่างงานวิจัยของคณะพยาบาลศาสตร์ เผยแพร่ด้วย GitHub Pages และบันทึกข้อมูลลง Google Sheets

เว็บไซต์: https://panpan22052525-max.github.io/nurs-research/

ผู้ดูแลระบบ: อาจารย์ปัณณทัต จันพานิช

## โครงสร้าง
| ไฟล์ | หน้าที่ |
|---|---|
| `index.html` | หน้าเว็บกรอกโครงร่างวิจัย 7 ขั้นตอน |
| `config.js` | ตั้งค่า URL ของ API (Google Apps Script Web App) |
| `logo.jpg` | ตราคณะพยาบาลศาสตร์ |
| `apps-script/Code.gs` | โค้ดหลังบ้านที่บันทึกลง Google Sheets (นำไปวางใน Apps Script) |

## การเชื่อมต่อ Google Sheets
1. เปิด https://script.google.com > โปรเจกต์ใหม่ > วางโค้ดจาก `apps-script/Code.gs`
2. รันฟังก์ชัน `setup` หนึ่งครั้ง แล้วอนุญาตสิทธิ์
3. Deploy > New deployment > Web app — Execute as: **Me**, Who has access: **Anyone**
4. คัดลอก URL ที่ลงท้ายด้วย `/exec` มาใส่ใน `config.js` ที่ `API_URL`

ชีตที่ใช้: `1BZTbYXSFAPv2sZ6Tv41FwMN496FNinTpTv8CR2FRLp0`
แท็บ: โครงร่างวิจัย · คณะผู้วิจัย · งบประมาณ · เอกสารแนบ · รายชื่ออาจารย์
