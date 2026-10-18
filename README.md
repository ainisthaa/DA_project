# Netflix Data Explorer

ระบบจัดการและวิเคราะห์ข้อมูลยอดดู Netflix (*What We Watched*, ม.ค.–มิ.ย. 2026, 17,370 เรื่อง)

**หน้าเว็บ (React) ↔ Backend (Go) ↔ MySQL ↔ Excel → Tableau → หน้าเว็บ**

## เริ่มใช้งาน (Docker)

```bash
docker compose up --build
```

เปิด http://localhost:8080 · ไฟล์ Excel อยู่ที่ `data/netflix_data.xlsx` (แก้ใน Excel ได้ ระบบซิงก์ให้เอง)

วิธีรันแบบไม่ใช้ Docker ดูใน [docs/USER_MANUAL.md](docs/USER_MANUAL.md) หัวข้อ 2

## โครงสร้างโฟลเดอร์

```
netflix-project/
├── data/netflix_data.xlsx        ไฟล์ Excel หลัก (sheet1 ข้อมูลดิบ, sheet2 วิเคราะห์, sheet3 Top10, sheet4 สรุปสำหรับ Tableau)
├── data/tableau_netflix2.twbx    ไฟล์ Tableau (Dashboard_Overview + Dashboard_Details)
├── data/tableau/                 ภาพ Dashboard ที่ Export จาก Tableau → แสดงในหน้าเว็บ (เมนู กราฟ Tableau)
├── old_tableau_versions/         ไฟล์ Tableau รุ่นเก่า (Book1, tableau_netflix) — ไม่ต้องส่ง
├── source/                       ไฟล์ต้นฉบับจาก Netflix
├── scripts/build_excel.py        สคริปต์สร้าง data/netflix_data.xlsx จากไฟล์ต้นฉบับ
├── backend/                      Go: REST API + MySQL + ซิงก์ Excel สองทาง
├── frontend/                     React + Vite: หน้าเว็บ
├── tableau/TABLEAU_GUIDE.md      ขั้นตอนทำกราฟใน Tableau และฝังบนเว็บ
├── docs/USER_MANUAL.md           คู่มือผู้ใช้พร้อมภาพหน้าจอ (docs/images/)
├── Dockerfile, docker-compose.yml
└── WORKLOG.md                    บันทึกการทำงานทุกขั้นตอนอย่างละเอียด
```

## สิ่งที่ต้องส่ง

| รายการ | ไฟล์ |
|---|---|
| Excel: ข้อมูลดิบ (sheet1), ข้อมูลที่คำนวณ (sheet2), อื่นๆ (sheet3–4) | `data/netflix_data.xlsx` |
| ตัว Tableau | `data/tableau_netflix2.twbx` (ขั้นตอน + รายการที่ควรแก้อยู่ใน `tableau/TABLEAU_GUIDE.md`) |
| Program | โฟลเดอร์ `backend/`, `frontend/`, `docker-compose.yml` |
| User manual พร้อมภาพหน้าจอ | **`docs/USER_MANUAL.docx`** (Word) · ต้นฉบับ `docs/USER_MANUAL.md` + `docs/images/` |

## สร้างไฟล์ Excel ใหม่จากต้นฉบับ (ถ้าต้องการ)

```bash
pip install openpyxl
python scripts/build_excel.py source/Netflix_What_We_Watched_2026Jan-Jun_original.xlsx data/netflix_data.xlsx
```
แล้วเปิดไฟล์ใน Excel → Save หนึ่งครั้งเพื่อให้สูตรคำนวณค่าเก็บไว้ (ระบบใน Docker ทำให้อัตโนมัติ)

## ที่มาข้อมูล

Netflix. (2026, July 16). *What We Watched: A Netflix Engagement Report (January–June 2026)*.
https://about.netflix.com/en/news/what-we-watched-the-first-half-of-2026

## สร้างคู่มือ .docx ใหม่ (หลังแก้ `docs/USER_MANUAL.md`)

ใช้ pandoc ผ่าน Docker (ไม่ต้องติดตั้งเพิ่ม) · `docs/reference.docx` คือแม่แบบ (ฟอนต์ Tahoma รองรับภาษาไทย, ตารางมีเส้นขอบ)

```bash
docker run --rm -v "$PWD":/proj -w /proj/docs pandoc/core:latest USER_MANUAL.md -f gfm -t docx \
  --reference-doc=reference.docx --resource-path=/proj/docs:/proj -o USER_MANUAL.docx
```
