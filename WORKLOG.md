# WORKLOG — Netflix Data Project

บันทึกการทำงานทุกขั้นตอน ตั้งแต่เตรียมข้อมูลจนระบบเสร็จ (เขียนต่อท้ายตามลำดับเวลา)

---

## 0. โจทย์ (จากรายละเอียด project)

1. หน้าเว็บสำหรับ input ข้อมูล แล้วข้อมูลไปอยู่ใน Excel (อัปเดตใน Excel ได้เลย)
2. Database ต้องลิงก์กับ Excel (เพิ่ม/ลบข้อมูลได้)
3. นำข้อมูลไปวิเคราะห์ใน Excel ด้วย tool ต่างๆ แล้วนำข้อมูลที่ได้ (เช่นค่าเฉลี่ย) ไปใช้ทำกราฟใน Tableau
4. ใช้ Tableau ทำกราฟ
5. ข้อมูล ผลลัพธ์ กราฟ และค่าที่คำนวณได้ นำไปแสดงบนหน้าเว็บ

สิ่งที่ต้องส่ง: Excel (sheet1 ข้อมูลดิบ, sheet2 ข้อมูลที่คำนวณ, sheet อื่นๆ), ตัว Tableau, program, แคปส่ง user manual

### ตัวเลือกที่ผู้ใช้เลือก (ถามก่อนเริ่ม)

| เรื่อง | เลือก |
|---|---|
| Tech stack | **Go + MySQL + React** |
| การเชื่อม Excel ↔ Database | **ได้ทั้งสองทาง** (แก้บนเว็บ → Excel อัปเดต / แก้ใน Excel → Database อัปเดต) |
| กราฟบนหน้าเว็บ | **ฝัง Tableau อย่างเดียว** |

---

## 1. Dataset

- แหล่งที่มา: Netflix, *What We Watched: A Netflix Engagement Report (January–June 2026)*
  - หน้าประกาศ: https://about.netflix.com/en/news/what-we-watched-the-first-half-of-2026
  - ไฟล์: `Netflix-s_What_We_Watched_Report_2026Jan-Jun__1_.xlsx`
- ไฟล์ต้นฉบับมี 2 sheet: `Shows` (8,202 แถวรวมหัวรายงาน) และ `Movies` (9,186 แถว)
- โครงสร้างต้นฉบับ: 5 แถวแรกเป็นหัวรายงาน, แถวที่ 6 เป็นชื่อคอลัมน์, คอลัมน์ A และ H ว่าง,
  ท้าย sheet มีแถว "Other Shows"/"Other Movies" (ยอดรวมของเรื่องที่ยอดดู < 50k) และ footnote

### การทำความสะอาด (ทำใน Python + openpyxl)

| ขั้นตอน | รายละเอียด |
|---|---|
| ตัดหัวรายงาน | เริ่มอ่านจากแถวที่ 7 (ใต้ชื่อคอลัมน์) |
| เพิ่มคอลัมน์ `Type` | Shows → `TV`, Movies → `Film` |
| รวม 2 sheet | ต่อ Film ท้าย TV ได้ตารางเดียว |
| ตัดแถวยอดรวม | ลบ "Other Shows" และ "Other Movies" (ไม่ใช่เรื่องจริง ทำให้ค่าเฉลี่ยเพี้ยน) |
| ตัด footnote/แถวว่าง | ข้ามแถวที่ไม่มี Title หรือ Hours Viewed |
| Runtime | ต้นฉบับเป็นข้อความ `H:MM` บางแถวมี `*` ต่อท้าย (เช่น `1:02*` = ไลฟ์) หรือเป็น `*` อย่างเดียว → ลบ `*` ออก, ถ้าว่างให้เป็นค่าว่าง |
| Runtime (Hours) | คอลัมน์ใหม่ แปลง `H:MM` เป็นทศนิยม เช่น `1:30` → `1.50` |
| Release Date | แปลงจากข้อความ `YYYY-MM-DD` เป็นวันที่จริงของ Excel |

ผลลัพธ์: **17,370 แถว** (TV 8,193 + Film 9,177)
- Release Date ว่าง 11,324 แถว (Netflix ใส่วันที่เฉพาะคอนเทนต์ของ Netflix)
- Runtime ว่าง 193 แถว (ส่วนใหญ่เป็นไลฟ์/กีฬา)

### หมายเหตุเรื่อง Original / Licensed

ในไฟล์ไม่มีคอลัมน์บอกตรงๆ ใช้เกณฑ์แทน (proxy): **มี Release Date = Netflix Title**, **ไม่มี = Non-Netflix/Licensed**
อ้างอิงจากที่ Netflix ระบุว่ารายงานจะใส่ premiere date ให้ซีรีส์/หนังของ Netflix
ข้อจำกัด: เรื่อง exclusive บางเรื่องอาจมีวันที่ทั้งที่ไม่ได้ผลิตเอง → ในรายงานควรเขียนหมายเหตุเกณฑ์นี้

---

## 2. สภาพแวดล้อมที่ใช้พัฒนา/ทดสอบ (ในเครื่องของ Claude)

- Go 1.24, Node 22, Python 3 + openpyxl, LibreOffice (ใช้คำนวณสูตรเพื่อตรวจ)
- proxy.golang.org ถูกบล็อกในสภาพแวดล้อมนี้ → ดึง Go modules ตรงจาก GitHub (`GOPROXY=direct` + replace ชั่วคราว)
  **ไม่มีผลกับเครื่องผู้ใช้** go.mod ที่ส่งให้เป็นแบบปกติ ไม่มี replace
- ไม่มี Docker daemon → ทดสอบด้วย MariaDB 10.11 (ใช้ SQL ชุดเดียวกับ MySQL 8 ได้) ส่วนที่ส่งให้ใช้ MySQL 8 ผ่าน Docker Compose

---

## 3. ไฟล์ Excel (`data/netflix_data.xlsx`) — สร้างด้วย `scripts/build_excel.py`

### โครงสร้าง sheet

| Sheet | เนื้อหา |
|---|---|
| `sheet1_RawData` | ข้อมูลดิบ 17,370 แถว คอลัมน์: ID, Type, Title, Available Globally, Release Date, Hours Viewed, Runtime, Runtime (Hours), Views |
| `sheet2_Analysis` | ข้อมูลที่คำนวณได้ด้วยสูตร Excel 7 หัวข้อ |
| `sheet3_Top10` | Top 10 Film และ Top 10 TV ด้วยสูตร |
| `sheet4_TableauSummary` | ตารางยาว `Section / Group / Metric / Value` ดึงค่าจาก sheet2–3 สำหรับให้ Tableau ใช้ทำกราฟ (119 แถว) |

### เทคนิคที่ใช้ใน Excel

- **Named Range** `D_ID, D_Type, D_Title, D_Global, D_Rel, D_Hours, D_Runtime, D_RtHours, D_Views`
  ชี้ไปที่ `sheet1_RawData!$X$2:$X$100000` → สูตรอ่านง่าย และรองรับข้อมูลเพิ่มถึง ~100,000 แถวโดยไม่ต้องแก้สูตร
- ฟังก์ชันที่ใช้: `COUNTA, COUNTIF(S), SUMIF(S), AVERAGEIF(S), MEDIAN(IF()), MAX(IF()), STDEV, LARGE, SMALL, INDEX, ROUNDUP, IFERROR, DATE`
- สูตร array (`MEDIAN(IF(...))`, `LARGE(IF(...))`) บันทึกเป็น array formula เพื่อให้ใช้ได้ทั้ง Excel รุ่นเก่าและใหม่
- Top 10 จัดการกรณียอดดูเท่ากันด้วย `SMALL(IF(...,ROW()),COUNTIF(...))`
- ตั้ง `fullCalcOnLoad` ให้ Excel คำนวณใหม่ทุกครั้งที่เปิดไฟล์

### หัวข้อการวิเคราะห์ใน sheet2

1. KPI ภาพรวม: จำนวนเรื่อง, ชั่วโมงดูรวม, ยอดดูรวม, ค่าเฉลี่ย, มัธยฐาน, ส่วนเบี่ยงเบนมาตรฐาน
2. TV vs Film: จำนวน, ชั่วโมงดู, % ชั่วโมงดู, ยอดดูเฉลี่ย/มัธยฐาน/สูงสุด, Runtime เฉลี่ย
3. Netflix Title (มี Release Date) vs Non-Netflix
4. Available Globally Yes vs No
5. ช่วงความยาว: Film (<1.5h, 1.5–2h, >2h) และ TV (<5h, 5–10h, >10h) — แยกเพราะ Runtime ของ TV คือทั้งซีซัน
6. Netflix Title ตามปีที่ออกฉาย (≤2015, 2016–2026)
7. Pareto: Top 1/5/10/20/50% ของเรื่องครองชั่วโมงดูกี่ %

### ผลการตรวจสอบ

คำนวณด้วย LibreOffice แล้วเทียบกับการคำนวณอิสระใน Python → **ตรงกันทุกค่า**
(จำนวน 17,370, ยอดดูเฉลี่ย 1,346,598, มัธยฐาน 300,000, TV/Film, Runtime group, ปี 2025, Pareto Top 1% = 25.2%)

### ผลวิเคราะห์ที่น่าสนใจ (ข้อมูลต้นฉบับ)

- TV มีจำนวนเรื่องน้อยกว่า Film แต่กิน **76%** ของชั่วโมงดูทั้งหมด
- Netflix Title (6,046 เรื่อง) ได้ยอดดูเฉลี่ย **1.93M** เกือบ 2 เท่าของ Non-Netflix (1.04M) แต่ชั่วโมงดูรวมพอๆ กัน (~50/50)
- เรื่องที่ฉายทั่วโลกได้ยอดดูเฉลี่ย 2.12M vs ไม่ทั่วโลก 1.01M
- **Top 1% (174 เรื่อง) ครอง 25% ของชั่วโมงดูทั้งหมด**, Top 10% ครอง 68%
- หนังยาว >2 ชม. มีมัธยฐานยอดดูสูงสุด (500k) เทียบกับ 200k ของหนังสั้นกว่า

---

## 4. Backend (Go) — โฟลเดอร์ `backend/`

| ไฟล์ | หน้าที่ |
|---|---|
| `main.go` | อ่าน config จาก environment, เปิด DB, เริ่มตัวซิงก์, เปิด HTTP server, เสิร์ฟหน้าเว็บ React ที่ build แล้ว |
| `models.go` | struct `Title` + `Normalize()` ตรวจและจัดรูปแบบข้อมูล (ใช้ทั้งจากเว็บและจาก Excel) |
| `db.go` | สร้างตาราง (migrate), CRUD, ค้นหา/กรอง/เรียง/แบ่งหน้า, `ApplyExcel` ซิงก์ทั้งตารางใน transaction, settings, sync_log |
| `excel.go` | อ่าน/เขียน `sheet1_RawData` ด้วย excelize, อ่านค่าที่คำนวณแล้วจาก `sheet4` |
| `sync.go` | ตัวจัดการซิงก์สองทาง (รายละเอียดด้านล่าง) |
| `stats.go` | คำนวณสถิติชุดเดียวกับ sheet2 จาก DB (ใช้ตรวจเทียบ และแสดงบนเว็บ) |
| `handlers.go` | REST API |

### ตารางใน MySQL

- `titles` (id, type, title, available_globally, release_date, hours_viewed, runtime, runtime_hours, views, updated_at) — utf8mb4 รองรับชื่อเรื่องภาษาไทย/เกาหลี/อาหรับ
- `settings` (k, v) — เก็บลิงก์ Tableau, สถานะซิงก์ (`excel_sig`, `db_dirty`)
- `sync_log` — ประวัติการซิงก์ทุกครั้ง แสดงในหน้าเว็บ

### REST API

| Method | Path | ใช้ทำ |
|---|---|---|
| GET | `/api/titles?q=&type=&global=&origin=&sort=&page=&size=` | ค้นหา/กรอง/เรียง/แบ่งหน้า |
| GET/POST/PUT/DELETE | `/api/titles`, `/api/titles/{id}` | ดู/เพิ่ม/แก้/ลบ |
| GET | `/api/stats` | สถิติที่คำนวณจาก DB |
| GET | `/api/excel/summary` | ค่าที่ Excel คำนวณ (sheet4) |
| GET | `/api/excel/download` | ดาวน์โหลดไฟล์ Excel |
| POST | `/api/excel/upload` | อัปโหลด .xlsx ที่แก้แล้ว เพื่อนำเข้า |
| GET | `/api/sync/status` | สถานะซิงก์ + ประวัติ |
| POST | `/api/sync/import[?force=1]`, `/api/sync/export` | ซิงก์ด้วยมือ |
| GET/PUT | `/api/settings` | ลิงก์ Tableau dashboard ที่ฝังบนเว็บ |

### การซิงก์สองทาง (sync.go)

- **เว็บ → DB → Excel**: เพิ่ม/แก้/ลบบนเว็บ → บันทึก DB ทันที → รอ 0.5 วินาที (รวมหลายคำสั่ง) → เขียน `sheet1` ใหม่ทั้งหมดโดยไม่แตะ sheet อื่น
  → ถ้ามี LibreOffice ให้คำนวณสูตร sheet2–4 ใหม่ทันที (ค่าที่ Tableau และเว็บอ่านจะเป็นปัจจุบัน)
- **Excel → DB**: ตรวจไฟล์ทุก 3 วินาที (เวลาแก้ไข + ขนาด) ถ้าเปลี่ยน → อ่าน sheet1 →
  แถวที่มี ID = แก้, แถวไม่มี ID = เพิ่ม (แล้วเขียน ID กลับลง Excel), ID ที่หายไป = ลบ
- อ่านคอลัมน์จาก**ชื่อหัวคอลัมน์** สลับลำดับคอลัมน์ได้; รองรับวันที่แบบ serial ของ Excel และ Runtime ที่ Excel แปลงเป็นเวลา
- แถวที่ข้อมูลผิดจะไม่ถูกนำเข้า และแสดงเลขแถว + สาเหตุในหน้าซิงก์
- **ระบบกันข้อมูลหาย**: ถ้า Excel มีข้อมูลน้อยกว่าครึ่งของ DB จะไม่นำเข้า ต้องกด "นำเข้าแบบบังคับ" เอง
- ถ้าเขียนไฟล์ไม่ได้ (เช่นเปิดไฟล์ค้างใน Excel บน Windows) จะลองใหม่ทุก 10 วินาที
- เขียนลงไฟล์ชั่วคราวแล้ว rename เพื่อไม่ให้ไฟล์เสียถ้าเขียนไม่จบ

### บั๊กที่เจอและแก้ระหว่างทดสอบ

| ปัญหา | สาเหตุ | วิธีแก้ |
|---|---|---|
| Export ล้มเหลว `unsupported workbook file format` | excelize ไม่ยอม SaveAs ไฟล์ที่นามสกุลไม่ใช่ .xlsx (ไฟล์ชั่วคราว `.xlsx.tmp`) | ตั้งชื่อไฟล์ชั่วคราวเป็น `.tmp-netflix_data.xlsx` |
| แถวที่เพิ่มบนเว็บหายหลังรีสตาร์ต (ถ้าปิดโปรแกรมก่อนเขียน Excel) | ตอนเปิดโปรแกรม ระบบถือ Excel เป็นหลักเสมอ | เก็บสถานะ `db_dirty` และลายเซ็นไฟล์ `excel_sig` ใน DB แล้วตัดสินทิศการซิงก์ตอนเปิดโปรแกรม |
| แก้ข้อแรกแล้วยังหาย | ตอนเปิดโปรแกรมยังไม่ได้จำสถานะไฟล์ในหน่วยความจำ ทำให้ Export ดึงจาก Excel ก่อน | เรียก `rememberFile()` ก่อน Export ตอน startup |

### ผลทดสอบ backend (MariaDB 10.11 + LibreOffice)

| กรณี | ผล |
|---|---|
| เปิดครั้งแรก DB ว่าง | นำเข้า 17,370 แถวใน ~1 วินาที ✅ |
| เพิ่มข้อมูลผ่าน API | ได้ ID ใหม่, Excel อัปเดต + คำนวณสูตรใหม่ ~7 วินาที ✅ |
| ส่งข้อมูลผิด (Type="anime", Runtime="abc") | ตอบ 400 พร้อมข้อความภาษาไทย ✅ |
| ปิดโปรแกรมก่อนเขียน Excel แล้วเปิดใหม่ | เขียนลง Excel ตอนเปิด ข้อมูลไม่หาย ✅ |
| แก้ Excel ภายนอก (แก้ค่า, ลบแถว, เพิ่มแถวไม่มี ID, แถวผิด) | ซิงก์เข้า DB อัตโนมัติ, แถวใหม่ได้ ID, แถวผิดถูกรายงาน ✅ |
| เทียบค่าที่ Excel คำนวณ (sheet4) กับ `/api/stats` | ตรงกัน 63/63 ค่า ✅ |
| แก้/ลบผ่าน API | Excel อัปเดตถูกต้อง, ลบ ID ที่ไม่มีได้ 404 ✅ |
| ลบข้อมูลใน Excel เหลือ 3 แถว | ระบบไม่ยอมนำเข้า แสดงคำเตือน ✅ |
| อัปโหลดไฟล์ Excel ผ่าน API | นำเข้าสำเร็จ sheet2–4 ของไฟล์หลักยังอยู่ ✅ |

---

## 5. Frontend (React + Vite) — โฟลเดอร์ `frontend/`

| ไฟล์ | หน้าที่ |
|---|---|
| `src/App.jsx` | โครงหน้า, เมนู 5 หน้า (ใช้ `#hash` เลือกหน้า ไม่ต้องลง router), ป้ายสถานะซิงก์มุมขวาบน (อัปเดตทุก 3 วิ) |
| `src/api.js` | ฟังก์ชันเรียก REST API + จัดรูปแบบตัวเลข (17,370 / 1.35M / 76.0%) |
| `src/pages/Dashboard.jsx` | **ภาพรวม**: KPI, ตารางวิเคราะห์ 6 หัวข้อ, Top 10, แถบตรวจว่าค่าที่ Excel คำนวณตรงกับ DB |
| `src/pages/Charts.jsx` | **กราฟ Tableau**: ฝัง Tableau Public ด้วย iframe (แปลงลิงก์แชร์เป็นลิงก์ฝังให้อัตโนมัติ), หลายแท็บ |
| `src/pages/DataTable.jsx` | **จัดการข้อมูล**: ค้นหา, กรอง 3 แบบ, เรียง 7 แบบ, แบ่งหน้า, ฟอร์มเพิ่ม/แก้ (ตรวจข้อมูลก่อนส่ง), ลบพร้อมยืนยัน |
| `src/pages/SyncPage.jsx` | **ซิงก์ Excel**: สถานะ, ดาวน์โหลด/อัปโหลด, ซิงก์ด้วยมือ, แถวที่นำเข้าไม่ได้, ประวัติ |
| `src/pages/SettingsPage.jsx` | **ตั้งค่า**: เพิ่ม/ลบลิงก์ Tableau Dashboard |
| `src/styles.css` | ธีมมืดโทน Netflix, ฟอนต์ IBM Plex Sans Thai, responsive มือถือ |

ตอนรันจริง backend Go เสิร์ฟไฟล์ที่ build แล้ว (`frontend/dist`) จึงเปิดเว็บและ API ที่พอร์ตเดียว (8080)

### ปรับแก้หลังดูภาพหน้าจอ

| ปัญหาที่เห็น | แก้ |
|---|---|
| มือถือ: เมนูซ้อนเป็นแนวตั้งกินพื้นที่ครึ่งจอ | เมนูเป็นแถวเดียวเลื่อนแนวนอนได้ |
| หน้าซิงก์: "นำเข้าล่าสุด" ว่างหลังรีสตาร์ต | ดึงเวลาล่าสุดจากตาราง sync_log ตอนเปิดโปรแกรม |
| หัวคอลัมน์ "Runtime เฉลี่ย (ชม.)" ล้นตาราง | ย่อเป็น "Runtime (ชม.)" |
| แถบเหลือง "Excel ยังคำนวณไม่เสร็จ" ค้างทั้งที่คำนวณเสร็จแล้ว | ให้หน้าภาพรวมโหลดใหม่เมื่อเวลาคำนวณล่าสุด (`lastRecalc`) เปลี่ยนด้วย |

### ทดสอบหน้าเว็บด้วย Playwright (Chromium headless)

- เพิ่มข้อมูลผ่านฟอร์ม → แถวขึ้นในตาราง และอยู่ใน Excel sheet1 ✅
- แก้ Views ผ่านฟอร์ม → ค่าเปลี่ยนทั้งในตารางและ Excel ✅
- กดบันทึกฟอร์มว่าง → แจ้ง "กรุณากรอกชื่อเรื่อง" ✅
- ลบผ่านปุ่ม → หายจากตารางและ Excel ✅
- ไม่มี JavaScript error (มีแค่โหลดฟอนต์ Google ไม่ได้ในเครื่องทดสอบซึ่งไม่มีอินเทอร์เน็ต)

---

## 6. Docker

- `Dockerfile` 3 ขั้น: build React (Node 22) → build Go (Go 1.24) → image รันจริง Debian + **LibreOffice Calc** (ใช้คำนวณสูตร Excel อัตโนมัติ) + ฟอนต์ไทย
- `docker-compose.yml`: `mysql` (MySQL 8.0, utf8mb4, healthcheck) + `app` (รอ MySQL พร้อมก่อน) · map `./data` เข้า container เพื่อให้แก้ไฟล์ Excel บนเครื่องได้
- ตรวจไวยากรณ์ด้วย `docker compose config` ผ่าน (เครื่องที่ใช้พัฒนาไม่มี Docker daemon จึง build image จริงไม่ได้ — ส่วนประกอบแต่ละตัวทดสอบแยกแล้ว)
- Build backend สำหรับ Windows (`GOOS=windows`) ผ่าน และแก้ path โปรไฟล์ LibreOffice ให้ใช้ได้บน Windows (`file:///C:/...`)

---

## 7. ไฟล์ Excel ที่ส่ง

ไฟล์ `data/netflix_data.xlsx` ผ่าน LibreOffice คำนวณสูตรแล้ว 1 รอบ จึงมี**ค่าที่คำนวณเก็บไว้** เปิดใน Tableau ได้ทันที
ตรวจแล้วว่ายังมี Named Range ครบ 9 ตัว และสูตร array (เช่น `=MEDIAN(IF(D_Type=B17,D_Views))`) อยู่ครบ

---

## 8. เอกสาร

| ไฟล์ | เนื้อหา |
|---|---|
| `README.md` | ภาพรวม, วิธีรัน, โครงสร้างโฟลเดอร์, รายการสิ่งที่ต้องส่ง |
| `docs/USER_MANUAL.md` | คู่มือผู้ใช้ 9 หัวข้อ + ภาพหน้าจอจริง 11 ภาพ + ลำดับขั้นตอนสาธิต + แก้ปัญหา |
| `docs/images/*.png` | ภาพหน้าจอเรียงตามขั้นตอน (01_dashboard … 12_mobile) |
| `tableau/TABLEAU_GUIDE.md` | ขั้นตอนทำกราฟ 10 แบบ, รวม Dashboard, บันทึก .twbx, publish และฝังบนเว็บ, การอัปเดตข้อมูล |

---

## 9. สิ่งที่ผู้ใช้ต้องทำเองต่อ

1. ~~Tableau: ทำกราฟตามคู่มือ~~ ✅ ทำแล้วเป็น `data/Book1.twbx` (ดูข้อ 11)
2. ~~แก้ Tableau~~ ✅ เสร็จใน `tableau_netflix2.twbx` (เหลือคลิกล้าง selection ก่อน Export)
3. Export ภาพ Dashboard ไปที่ `data/tableau/` (และ/หรือ publish Tableau Public แล้ววางลิงก์ในหน้าตั้งค่า)
4. **ภาพหน้าจอ Tableau** ใส่ใน user manual ตรงที่เขียน `[ใส่ภาพหน้าจอ ...]` (3 จุด: 13, 14, 15)
5. (ถ้าอาจารย์ต้องการ) ถ่ายภาพหน้าจอจากเครื่องตัวเองแทนภาพใน `docs/images/`

---

## 11. Tableau — Book1.twbx (18 ต.ค. 2026)

ผู้ใช้ทำ workbook ใน Tableau Public 2026.2 บันทึกเป็น `data/Book1.twbx` → แตกไฟล์ (.twbx = zip) อ่าน XML ของ `Book1.twb` เพื่อตรวจ

### โครงสร้าง

| ส่วน | รายละเอียด |
|---|---|
| Data source | `sheet4_TableauSummary` (extract `Netflix_Summary.hyper`) + calc `Kind`; `sheet1_RawData` (extract อยู่ในไฟล์) + calc `Valid Runtime = [Runtime (Hours)] > 0` |
| Worksheet (14) | 01_KPI (Text), 02_TV_Film_Share (Pie), 03_AVG_MED (Bar คู่), 04_Origin ×2, 05_Global ×2, 06_Runtime_Views, 07_Release_Year (Line), 07_Release_Year_Titles (Bar), 08_Pareto, 09_Top10_Film, 09_Top10_TV (Top 10 by SUM(Views)), 10_Runtime_Scatter (log, Detail = ID) |
| Dashboard (2) | `Dashboard_Overview` 1400×2000 (กราฟ 1–8), `Dashboard_Details` 1400×1200 (กราฟ 9–10) |

ฟิลเตอร์ Section/Metric ทุกกราฟตรงกับคู่มือ และค่าใน thumbnail ตรงกับ Excel (1,346,598 / 300,000 / 17,370 / TV 76.03% / Pareto Top 10% = 67.95%)

### ปัญหาที่พบ (ไม่แก้ XML เอง เพราะไฟล์เปิดค้างใน Tableau และแก้ด้วยมือเสี่ยงไฟล์เสีย → เขียนวิธีแก้ใน TABLEAU_GUIDE)

| ปัญหา | หลักฐานใน XML |
|---|---|
| Top 10 เรียงตามชื่อ ไม่ใช่ยอดดู | `single-value-per-nest-shelf-sort ... measure-to-sort-by='[none:Title:nk]' direction='DESC'` |
| แกนปีเรียง 2018 ก่อน 2017 | `shelf-sort-v2 ... measure-to-sort-by='[sum:Value:qk]' direction='ASC'` (เรียงตามค่า) |
| data source `netflix_data` ไม่มีกราฟใช้ และ extract ชี้ไป `/var/folders/.../tableau-temp/...hyper` (ไม่ได้แพ็กในไฟล์) | อ้างถึงแค่ครั้งเดียวตอนประกาศ datasource |
| ชื่อ worksheet 6 อันมีช่องว่างนำหน้า, หัวกราฟเป็นชื่อ sheet, แกน log ของ scatter เริ่มที่ 1 | ดูจาก thumbnail ใน twb |
| ยังไม่ publish | ไม่มี `repository-location` |

### เพิ่มในโปรแกรม: แสดงกราฟ Tableau บนเว็บโดยไม่ต้อง publish

เดิมหน้า **กราฟ Tableau** แสดงได้เฉพาะลิงก์ Tableau Public (ต้อง publish + มีเน็ต) → เพิ่มทางเลือกแสดงภาพ Dashboard ที่ Export จาก Book1

| ไฟล์ | เปลี่ยน |
|---|---|
| `backend/tableau.go` (ใหม่) | `GET /api/tableau` (ข้อมูล Book1 + รายการภาพใน `data/tableau/`), `GET /api/tableau/images/{file}`, `GET /api/tableau/workbook` (ดาวน์โหลด Book1.twbx) · กันอ่านไฟล์นอกโฟลเดอร์ (`../`), รับเฉพาะ .png/.jpg · ชื่อไฟล์ตั้งค่าได้ด้วย `TABLEAU_WORKBOOK` |
| `backend/handlers.go`, `main.go` | ผูก route ใหม่ |
| `frontend/src/pages/Charts.jsx` | แท็บรวม: ลิงก์ Tableau Public (iframe) + ภาพ (img) · ปุ่มดาวน์โหลด Book1.twbx · ข้อความแนะนำ 2 วิธีเมื่อยังไม่มีอะไร |
| `frontend/src/api.js`, `styles.css` | `api.tableau()`, สไตล์ `.tableau-image` |

ทดสอบ: `go vet` + `go build` + `vite build` ผ่าน · ทดสอบ handler ด้วย httptest (ใช้ Book1.twbx จริง): รายการภาพเรียงตามชื่อไฟล์, ข้าม `.DS_Store`/ไฟล์ไม่ใช่ภาพ, `..%2Fsecret.png` → 404, ดาวน์โหลด twbx ได้ 785,514 ไบต์ ✅
(เครื่องนี้ไม่มี MySQL/Docker daemon จึงยังไม่ได้ทดสอบทั้งระบบ end-to-end ในรอบนี้)

เอกสารที่อัปเดต: `README.md`, `tableau/TABLEAU_GUIDE.md` (สถานะ Book1 + ตารางแก้ + ข้อ 4.1 Export ภาพ), `docs/USER_MANUAL.md` ข้อ 7–9

## 12. Tableau รุ่นใหม่ — tableau_netflix.twbx (18 ต.ค. 2026 23:41)

ผู้ใช้แก้แล้วบันทึกเป็นไฟล์ใหม่ `data/tableau_netflix.twbx` (แทน `Book1.twbx`)

| เรื่อง | ผลตรวจ |
|---|---|
| Top 10 Film / TV | เปลี่ยนเป็น Manual sort — ลำดับตรงกับยอดดูใน Excel ครบ 10 อันดับทั้ง 2 กราฟ ✅ (คำนวณอิสระจาก sheet1 เทียบ เช่น TV: His & Hers 104.0M, Bridgerton S4 100.2M, …, The Night Agent S3 36.1M) |
| 07_Release_Year_Titles | Manual sort ≤2015 → 2026 ✅ |
| 07_Release_Year (Line) | ยังเรียงตาม SUM(Value) → 2018 มาก่อน 2017 ❌ |
| data source `netflix_data` ที่ไม่ใช้ / ชื่อ sheet มีช่องว่าง / หัวกราฟ / แกน log | ยังเหมือนเดิม (ไม่บังคับ) |

ข้อสังเกต: Manual sort ล็อกลำดับชื่อเรื่อง ถ้าข้อมูลเปลี่ยน (เพิ่มเรื่องดังผ่านเว็บ) เรื่องใหม่จะไปอยู่ท้าย → แนะนำ Sort by Field (Views, Sum) แทน

เปลี่ยนในโปรแกรม: ค่าเริ่มต้น `TABLEAU_WORKBOOK` = `tableau_netflix.twbx` (backend/tableau.go), ข้อความใน Charts.jsx, README, USER_MANUAL, TABLEAU_GUIDE (ตาราง "แก้แล้ว / ยังเหลือ")

## 13. Tableau รุ่นสุดท้าย — tableau_netflix2.twbx (18 ต.ค. 2026 23:55)

| เรื่อง | ผลตรวจ (อ่านจาก XML + thumbnail) |
|---|---|
| Top 10 Film / TV | `computed-sort` ตาม `SUM(Views)` DESC ✅ (เรียงตามข้อมูลเอง ไม่ล็อกแบบ Manual แล้ว) |
| 07_Release_Year (Line) | `manual-sort` ≤2015 → 2026 ✅ (thumbnail: 2016 → 2017 → 2018 ถูกลำดับ) |
| 07_Release_Year_Titles, 06_Runtime_Views | ลำดับถูกต้อง ✅ |
| ข้อสังเกต | มี selection ค้าง 3 จุด (2017, Office Romance, 1 แท่งใน Overview) → แท่งอื่นจางตอน Export · ให้คลิกที่ว่างก่อน Export |
| ไม่บังคับ | data source `netflix_data` ที่ไม่ใช้, ช่องว่างหน้าชื่อ sheet, หัวกราฟ, แกน log, ยังไม่ publish |

จัดไฟล์: ย้าย `Book1.twbx`, `tableau_netflix.twbx` → `old_tableau_versions/` · ลบไฟล์ล็อกเก่าของ Excel `data/~$netflix_data.xlsx` (Excel ไม่ได้เปิดอยู่)
เปลี่ยนในโปรแกรม: ค่าเริ่มต้น `TABLEAU_WORKBOOK` = `tableau_netflix2.twbx` + ข้อความใน Charts.jsx, README, USER_MANUAL, TABLEAU_GUIDE

## 14. Publish Tableau Public + นำกราฟขึ้นเว็บ (19 ต.ค. 2026)

| ขั้น | รายละเอียด |
|---|---|
| ปัญหา Export Image | Tableau Public ไม่มีเมนู Dashboard → Export Image (มีเฉพาะ Desktop) → เปลี่ยนเป็นดึงภาพจาก Tableau Public |
| Publish error `AB9D5A55` "Missing connection name in relation / Unable to create extract" | สาเหตุ: data source `netflix_data` ที่ไม่ใช้ ต้องการ extract `/var/folders/.../#TableauTemp_0toumvx….hyper` ซึ่งถูกลบไปแล้ว (ตรวจในโฟลเดอร์ temp ไม่พบ) → แนะนำ Close data source นั้น · ทำสำเนา `tableau_netflix3.twbx` ที่ลบ datasource ออกจาก XML ไว้สำรอง (ผู้ใช้แก้ในโปรแกรมเองแล้ว publish ผ่าน → ย้าย netflix3 ไป `old_tableau_versions/`) |
| ลิงก์ | `https://public.tableau.com/views/tableau_netflix2/Dashboard_Overview`, `.../Dashboard_Details` · ตรวจแล้วไม่มี X-Frame-Options → ฝัง iframe ได้ |
| ภาพ | ดึงจาก `<url>.png?:display_static_image=y` → `data/tableau/01_Dashboard_Overview.png` (1399×2026), `02_Dashboard_Details.png` (1399×1226) — ตรวจค่าในภาพตรงกับ Excel |
| โปรแกรม | `handlers.go`: `defaultTableauViews` = 2 ลิงก์ข้างบน ใช้เมื่อยังไม่เคยบันทึกในหน้าตั้งค่า → เปิดเว็บแล้วเห็นกราฟทันที (แท็บ live 2 + ภาพ 2) |
| ที่เห็นในภาพ | 03_AVG_MED มี selection ค้าง (แท่งอื่นจาง), แกน Y ของ 07_Release_Year_Titles แสดง `0.00B` — ไม่กระทบค่า แก้แล้ว publish ทับได้ |

## 10. ข้อจำกัดที่ทราบ

- Tableau Public ไม่อัปเดตข้อมูลอัตโนมัติ ต้อง Refresh + Save ใหม่เมื่อข้อมูลเปลี่ยน (ข้อจำกัดของ Tableau Public ฟรี)
- Original/Licensed ใช้เกณฑ์แทนจาก Release Date ไม่ใช่ป้ายกำกับทางการจาก Netflix
- แถวใน Excel ที่ข้อมูลผิดจะไม่ถูกนำเข้า และจะหายจากไฟล์เมื่อระบบเขียนไฟล์ใหม่ครั้งถัดไป (ดูรายการได้ในหน้าซิงก์)
- ถ้าแก้ทั้งบนเว็บและใน Excel พร้อมกันภายในไม่กี่วินาที ระบบจะดึง Excel เข้าก่อนแล้วจึงเขียนการแก้จากเว็บลงไป
- Runtime รองรับรูปแบบ `H:MM` เท่านั้น

## สรุปสถานะ

| งาน | สถานะ |
|---|---|
| Dataset > 10,000 แถว (17,370) | ✅ |
| Excel sheet1 ข้อมูลดิบ / sheet2 คำนวณ / sheet3–4 เพิ่มเติม | ✅ |
| หน้าเว็บ input ข้อมูล → Excel อัปเดต | ✅ |
| MySQL ลิงก์กับ Excel สองทาง (เพิ่ม/แก้/ลบ) | ✅ |
| ค่าที่คำนวณใน Excel → Tableau (sheet4) | ✅ ใช้ sheet4 ทำ 11 กราฟ |
| Tableau (`data/tableau_netflix2.twbx`) | ✅ กราฟถูกต้องครบ (เหลือรายการไม่บังคับ ใน TABLEAU_GUIDE) |
| แสดงข้อมูล/ผลคำนวณ/กราฟบนหน้าเว็บ | ✅ ฝัง Tableau Public 2 dashboard + ภาพสำรอง 2 ภาพ (ตั้งไว้เป็นค่าเริ่มต้น) |
| User manual พร้อมภาพหน้าจอ | ✅ (เหลือภาพ Tableau 3 ภาพ) |

## 15. รันทั้งระบบด้วย Docker บนเครื่องผู้ใช้ (19 ต.ค. 2026)

- `docker compose up --build -d` → build ผ่าน แต่ MySQL start ไม่ได้: `Bind for 0.0.0.0:3306 failed: port is already allocated`
  เพราะมี container `lineoa_mysql` ของโปรเจกต์อื่นใช้ 3306 อยู่ (ไม่ได้ไปหยุดของโปรเจกต์อื่น)
- แก้ `docker-compose.yml` ให้ map MySQL ฝั่งเครื่องเป็น **3308** (app ต่อ `mysql:3306` ภายใน network ของ Docker เหมือนเดิม)
- ผล: app ขึ้นใน ~10 วินาที, นำเข้า Excel → DB 17,370 แถว, `/api/settings` คืนลิงก์ Tableau 2 อัน, `/api/tableau` เห็นภาพ 2 ภาพ + `tableau_netflix2.twbx`, หน้าเว็บ 200 ✅

## 16. ทดสอบเพิ่ม/ลบข้อมูลผ่านหน้าเว็บ + แก้บั๊กสูตร Excel ไม่คำนวณใหม่ (19 ต.ค. 2026)

ผู้ใช้ถามว่าเว็บเพิ่ม row ลง Excel ได้หรือยัง → ทดสอบจริงผ่านหน้าเว็บ (Playwright + Chrome) บนระบบ Docker

| ขั้น | ผล |
|---|---|
| เมนู **จัดการข้อมูล → + เพิ่มข้อมูล** กรอก "TEST ROW from web" → บันทึก | ได้ ID 17371, แถวขึ้นท้าย `sheet1_RawData` ใน `data/netflix_data.xlsx` ภายใน ~6 วินาที ✅ |
| ปุ่ม **ลบ** | แถวหายจาก Excel กลับเป็น 17,370 ✅ |
| ค่าใน sheet2–4 หลังเพิ่มแถว | ❌ **บั๊ก**: log บอก "คำนวณสูตร sheet2–4 ใหม่แล้ว" แต่ Total Titles ยังเป็น 17,370 |

**สาเหตุ**: LibreOffice (รุ่นใน Debian bookworm) ตั้งค่าเริ่มต้น *Recalculation on File Load* ของไฟล์ Excel เป็น "ไม่คำนวณใหม่"
→ `soffice --convert-to` จึงเขียนค่าเดิมที่เก็บในไฟล์กลับไป (ตอนพัฒนาครั้งแรกใช้ LibreOffice คนละรุ่น/คนละโปรไฟล์ จึงไม่เจอ)

**แก้** (`backend/sync.go`): `ensureRecalcOnLoad()` เขียน `OOXMLRecalcMode = 0` (คำนวณใหม่เสมอ) ลง `registrymodifications.xcu` ของโปรไฟล์ LibreOffice ก่อนรันทุกครั้ง (สร้างไฟล์ให้ถ้ายังไม่มี)

**ทดสอบหลังแก้** (container ใหม่ โปรไฟล์ว่าง): เพิ่ม 2 แถว → sheet4 Total Titles = 17,372, Average Views = 1,372,346 **ตรงกับ `/api/stats` ของ DB** ✅ → ลบ 2 แถว → กลับเป็น 17,370 / 1,346,598 ✅

แก้เพิ่ม: ปุ่มบนหัว Card (เช่น "เปิดใน Tableau" + "ดาวน์โหลด") เคยแยกกันคนละฝั่ง → รวมไว้ใน `.card-actions` ชิดขวา

## 17. คู่มือผู้ใช้ .docx + ภาพหน้าจอที่ขาด (19 ต.ค. 2026)

- ถ่ายภาพหน้าจอจริงด้วย Playwright (ใช้ Chrome ในเครื่อง): `15_web_charts.png` (แท็บ Tableau แบบโต้ตอบ), `16_web_charts_image.png` (แท็บภาพ)
- `13_tableau_overview.png`, `14_tableau_details.png` = ภาพ Dashboard ที่ render จาก Tableau Public → ไม่มีช่อง `[ใส่ภาพหน้าจอ]` เหลือในคู่มือ
- แผนภาพระบบแบบตัวอักษร (ASCII) เพี้ยนเมื่อมีภาษาไทย → เปลี่ยนเป็นตารางลำดับการไหลของข้อมูล 6 ขั้น
- แปลง `docs/USER_MANUAL.md` → **`docs/USER_MANUAL.docx`** ด้วย pandoc 3.11 (Docker) + แม่แบบ `docs/reference.docx` (Tahoma, ตารางมีเส้นขอบ/หัวตารางตัวหนา, ภาษา bidi = th-TH) · ภาพฝังครบ 15 ภาพ · ตรวจหน้าตาโดยแปลงเป็น PDF ด้วย LibreOffice Writer (20 หน้า)
