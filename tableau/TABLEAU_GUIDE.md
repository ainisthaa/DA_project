# คู่มือทำกราฟใน Tableau

ส่วนนี้ต้องทำในโปรแกรม Tableau บนเครื่องคุณ (Tableau เป็นโปรแกรมแบบลากวาง สร้างไฟล์ให้จากโค้ดไม่ได้)
ใช้เวลาประมาณ 1–2 ชั่วโมง ทำตามทีละข้อได้เลย

ผลลัพธ์ที่จะได้:

1. ไฟล์ **`data/tableau_netflix2.twbx`** สำหรับส่งอาจารย์ (ทำแล้ว ✅)
2. Dashboard บนหน้าเว็บ (เมนู **กราฟ Tableau**) — เป็นภาพที่ Export จาก tableau_netflix2 และ/หรือฝังจาก Tableau Public

---

## สถานะ tableau_netflix2.twbx (ตรวจจากไฟล์ 18 ต.ค. 2026 23:55)

ทำครบตามข้อ 1–3 ของคู่มือนี้แล้ว: 14 worksheet + 2 dashboard (`Dashboard_Overview`, `Dashboard_Details`)
ค่าในกราฟตรงกับ Excel (เช่น ยอดดูเฉลี่ย 1,346,598, มัธยฐาน 300,000, 17,370 เรื่อง, TV 76.03% ของชั่วโมงดู)
รุ่นก่อนหน้า (`Book1.twbx`, `tableau_netflix.twbx`, `tableau_netflix3.twbx`) ย้ายไปไว้ที่ `old_tableau_versions/` แล้ว — ไม่ต้องส่ง

**Publish แล้ว (19 ต.ค. 2026)** — หน้าเว็บใช้ลิงก์นี้เป็นค่าเริ่มต้น และภาพ Dashboard อยู่ใน `data/tableau/` แล้ว
- https://public.tableau.com/views/tableau_netflix2/Dashboard_Overview
- https://public.tableau.com/views/tableau_netflix2/Dashboard_Details

### แก้แล้ว ✅ — ไม่มีจุดผิดที่ต้องแก้แล้ว

| กราฟ | ผล |
|---|---|
| 09_Top10_Film, 09_Top10_TV | Sort by Field: SUM(Views) มาก→น้อย — ลำดับตรงกับ Excel และจะเรียงใหม่เองเมื่อข้อมูลเปลี่ยน |
| data source `netflix_data` ที่เสีย (ทำให้ publish error) | ลบออกแล้ว (21 ต.ค.) — เหลือ 2 data source ที่กราฟใช้ |
| 07_Release_Year (Line), 07_Release_Year_Titles (Bar) | แกนปีเรียง ≤2015, 2016, 2017 … 2026 ถูกต้อง |
| 06_Runtime_Views | เรียง Film <1.5h … TV >10h ถูกต้อง |

### ก่อน Export ภาพ / นำเสนอ

(ในเวอร์ชันที่ publish แล้ว กราฟ 03_AVG_MED ยังมีแท่ง Film/Average Views ถูกเลือกค้าง แท่งอื่นจาง)
ไฟล์บันทึก "การคลิกเลือก" ค้างไว้ 3 จุด (จุดปี 2017 ใน 07_Release_Year, แท่ง Office Romance ใน 09_Top10_Film, แท่งหนึ่งใน Dashboard_Overview)
ทำให้แท่งอื่นจางลงตอนเปิด/Export → **คลิกพื้นที่ว่างในกราฟ (หรือกด Esc) ให้ทุกแท่งสีเต็มก่อน แล้วค่อย Save / Export Image**

### ปรับเพิ่มได้ (ไม่บังคับ ไม่กระทบความถูกต้อง)

| # | เรื่อง | วิธีแก้ใน Tableau |
|---|---|---|
| 2 | ชื่อ worksheet 5 อันมีช่องว่างนำหน้า (07_Release_Year_Titles, 08_Pareto, 09_Top10_Film, 09_Top10_TV, 10_Runtime_Scatter) | ดับเบิลคลิกแท็บด้านล่าง → ลบช่องว่าง |
| 3 | หัวกราฟบน Dashboard เป็นชื่อ sheet เช่น `01_KPI` | ดับเบิลคลิกหัวกราฟ → พิมพ์ชื่อที่อ่านรู้เรื่อง |
| 4 | 10_Runtime_Scatter แกน Y (log) เริ่มที่ 1 ครึ่งล่างว่าง | คลิกขวาแกน Y → **Edit Axis** → Fixed Start = `50,000` |
| 5 | 07_Release_Year_Titles แกน Y แสดง `0.00B` (ใช้รูปแบบตัวเลขของ Hours) ทั้งที่เป็นจำนวนเรื่อง | คลิกขวาแกน Y → **Format** → Numbers: **Number (Custom)** ทศนิยม 0, Display Units: None |

---

## 0. เตรียมโปรแกรม

- **Tableau Public** (ฟรี): https://public.tableau.com/app/discover → Create → Download Tableau Public
- หรือ **Tableau Desktop** ฟรีสำหรับนักศึกษา (สมัครด้วยอีเมลมหาวิทยาลัย): https://www.tableau.com/academic/students
- สมัครบัญชี Tableau Public (ใช้ตอน publish เพื่อเอาไปฝังบนเว็บ)

> ก่อนเปิดใน Tableau ให้แน่ใจว่าค่าใน sheet2–4 คำนวณแล้ว: ถ้ารันระบบด้วย Docker ระบบคำนวณให้เอง
> ถ้าไม่แน่ใจ ให้เปิด `data/netflix_data.xlsx` ใน Excel แล้วกด **Save** หนึ่งครั้ง แล้วปิดไฟล์

---

## 1. เชื่อมข้อมูล (Data Source)

1. เปิด Tableau → หน้า Connect → **To a File → Microsoft Excel** → เลือก `data/netflix_data.xlsx`
2. ลาก **`sheet1_RawData`** ลงพื้นที่ตรงกลาง (ข้อมูลดิบ 17,370 แถว)
3. ตรวจชนิดข้อมูลในตารางตัวอย่าง: `Release Date` = Date, `Hours Viewed` / `Views` / `Runtime (Hours)` = Number
   ถ้า `ID` เป็น Measure ให้คลิกขวา → **Convert to Dimension**
4. เพิ่มแหล่งข้อมูลที่ 2: เมนู **Data → New Data Source → Microsoft Excel** → ไฟล์เดิม → ลาก **`sheet4_TableauSummary`**
   (ค่าที่ Excel คำนวณไว้ใน sheet2 เช่น ค่าเฉลี่ย มัธยฐาน % — ตรงกับโจทย์ "นำค่าที่คำนวณใน Excel ไปทำกราฟ")
5. เปลี่ยนชื่อ data source ด้านซ้ายบนให้จำง่าย: `Raw Data` และ `Excel Summary`

`sheet4_TableauSummary` มี 4 คอลัมน์

| Section | Group | Metric | Value |
|---|---|---|---|
| KPI / Type / Origin / Available Globally / Runtime Group / Release Year / Pareto / Top10 Film / Top10 TV | ชื่อกลุ่ม | ชื่อค่า เช่น Average Views | ตัวเลข |

ทุกกราฟจาก sheet4 ใช้หลักเดียวกัน: **กรอง Section + Metric → Group ไว้แกนหนึ่ง, SUM(Value) ไว้อีกแกน**

---

## 2. สร้างกราฟ (แต่ละข้อ = 1 Worksheet, ตั้งชื่อ tab ตามหัวข้อ)

### กราฟ 1 — KPI ภาพรวม (Excel Summary)
1. ลาก `Section` ไป Filters → เลือก `KPI`
2. ลาก `Metric` ไป Columns, `SUM(Value)` ไป **Text** (บน Marks)
3. Marks เป็น **Text**, ปรับฟอนต์ใหญ่ · ฟอร์แมตตัวเลข: คลิกขวา Value → Format → Numbers → Number (Custom) ทศนิยม 0, ติ๊ก Display Units = Millions/Billions

### กราฟ 2 — สัดส่วนชั่วโมงดู TV กับ Film (Pie)
1. Filters: `Section` = `Type`, `Metric` = `Share of Hours`
2. Marks → **Pie**, ลาก `Group` ไป **Color**, `SUM(Value)` ไป **Angle**
3. ลาก `Group` และ `SUM(Value)` ไป **Label** → Format Value เป็น Percentage
   ข้อสรุป: TV มีเรื่องน้อยกว่าแต่ครองชั่วโมงดู ~76%

### กราฟ 3 — ยอดดูเฉลี่ย vs มัธยฐาน แยก TV/Film (Bar คู่)
1. Filters: `Section` = `Type`, `Metric` = `Average Views` และ `Median Views`
2. Columns: `Group`, `Metric` · Rows: `SUM(Value)` · Color: `Metric`
   ข้อสรุป: ค่าเฉลี่ยสูงกว่ามัธยฐานมาก = ข้อมูลเบ้ขวา มีเรื่องดังไม่กี่เรื่องดึงค่าเฉลี่ยขึ้น

### กราฟ 4 — Netflix Title กับ Non-Netflix (Bar)
1. Filters: `Section` = `Origin`, `Metric` = `Average Views`
2. Columns: `Group` · Rows: `SUM(Value)` · Color: `Group` · Label: `SUM(Value)`
3. ทำซ้ำอีก worksheet ด้วย `Metric` = `Share of Hours` เพื่อเทียบสัดส่วนชั่วโมงดู

### กราฟ 5 — ฉายทั่วโลกกับไม่ทั่วโลก (Bar)
เหมือนกราฟ 4 แต่ `Section` = `Available Globally`

### กราฟ 6 — ความยาวกับยอดดู (Bar)
1. Filters: `Section` = `Runtime Group`, `Metric` = `Average Views`
2. Rows: `Group` · Columns: `SUM(Value)` (bar แนวนอน)
3. เรียงลำดับเอง: คลิกขวา `Group` → Sort → **Manual** → เรียง Film <1.5h, Film 1.5-2h, Film >2h, TV <5h, TV 5-10h, TV >10h
4. สร้าง Calculated Field ชื่อ `Kind` = `IF CONTAINS([Group],"Film") THEN "Film" ELSE "TV" END` แล้วลากไป Color

### กราฟ 7 — Netflix Title ตามปีที่ออกฉาย (Line + Bar)
1. Filters: `Section` = `Release Year`, `Metric` = `Hours Viewed`
2. Columns: `Group` (เรียง Manual ให้ ≤2015 อยู่ซ้ายสุด) · Rows: `SUM(Value)` · Marks: **Line**
3. (เสริม) สร้าง worksheet ใหม่ใช้ `Metric` = `Titles` แบบ Bar แล้วรวมเป็น Dual Axis
   ข้อสรุป: เรื่องปี 2026 มี 481 เรื่องแต่กินชั่วโมงดูมากที่สุด เรื่องเก่ายังมีคนดูต่อเนื่อง

### กราฟ 8 — การกระจุกตัวของยอดดู (Pareto)
1. Filters: `Section` = `Pareto`
2. Columns: `Group` (เรียง Manual: Top 1%, 5%, 10%, 20%, 50%) · Rows: `SUM(Value)` · Format เป็น %
3. Label: `SUM(Value)`
   ข้อสรุป: เรื่องท็อป 1% (174 เรื่อง) ครองชั่วโมงดู ~25%

### กราฟ 9 — Top 10 Film และ Top 10 TV (จากข้อมูลดิบ)
ใช้ data source **Raw Data**
1. Rows: `Title` · Columns: `SUM(Views)` · Filters: `Type` = `Film`
2. คลิกขวา `Title` ใน Filters → **Top** → By field → Top **10** by `Views` Sum
3. เรียงจากมากไปน้อย (ปุ่ม Sort descending บน toolbar) · Label: `SUM(Views)`
4. Duplicate sheet แล้วเปลี่ยน Filter `Type` = `TV`

### กราฟ 10 — Runtime กับ Views (Scatter, จากข้อมูลดิบ)
1. Columns: `Runtime (Hours)` → คลิกขวาเปลี่ยนเป็น **Dimension** (ให้แต่ละเรื่องเป็นจุด)
2. Rows: `SUM(Views)` → คลิกขวาแกน Y → Edit Axis → ติ๊ก **Logarithmic**
3. Detail: `Title` · Color: `Type` · Filters: `Runtime (Hours)` ไม่เอาค่า Null
4. (เสริม) Analytics pane → ลาก **Trend Line** ลงกราฟ

---

## 3. รวมเป็น Dashboard

1. ปุ่ม **New Dashboard** (ไอคอนด้านล่าง) → Size: **Automatic** (เพื่อให้พอดีกับหน้าเว็บ)
2. ลาก worksheet ลงตามลำดับ เช่น
   - แถวบน: กราฟ 1 (KPI)
   - แถวกลาง: กราฟ 2, 3, 4, 5
   - แถวล่าง: กราฟ 6, 7, 8
3. สร้าง Dashboard ที่ 2 สำหรับ Top 10 (กราฟ 9 ทั้งสอง) + Scatter (กราฟ 10)
4. ใส่ Title ด้านบน: `Netflix What We Watched — Jan–Jun 2026`
5. (เสริม) ทำ Filter Action: Dashboard → Actions → Add Action → Filter → คลิกที่ Type ในกราฟ 2 แล้วกรองกราฟอื่น

---

## 4. บันทึกไฟล์สำหรับส่ง

- **File → Save As** → ชนิด **Tableau Packaged Workbook (.twbx)** → `data/tableau_netflix2.twbx`
  (.twbx รวมข้อมูลไว้ในไฟล์ เปิดเครื่องอื่นได้เลย · ระบบเว็บหาไฟล์ชื่อนี้ในโฟลเดอร์ `data/` เพื่อให้ดาวน์โหลดจากหน้ากราฟ
  ถ้าจะใช้ชื่ออื่น ตั้ง environment `TABLEAU_WORKBOOK=ชื่อไฟล์.twbx`)

### 4.1 ภาพ Dashboard สำหรับแสดงบนเว็บ (ใช้ได้ออฟไลน์)

> **Tableau Public: ⌘S / File → Save บันทึกขึ้น Tableau Public (cloud) ไม่ได้เขียนไฟล์ .twbx ในเครื่อง** — ถ้าต้องการไฟล์ล่าสุด ดาวน์โหลดได้จาก `https://public.tableau.com/workbooks/tableau_netflix2.twb` (ได้ .twbx แต่จะไม่มีลิงก์ไปที่ไฟล์ Excel ต้อง Edit Data Source ก่อน Refresh)
>
> **Tableau Public ไม่มีเมนู Export Image** (มีเฉพาะ Tableau Desktop) → ใช้วิธีดึงภาพจาก Tableau Public หลัง publish:
> เปิด `https://public.tableau.com/views/<ชื่อ workbook>/<ชื่อ dashboard>.png?:display_static_image=y` ในเบราว์เซอร์ → บันทึกภาพ
> หรือในหน้า viz กด **Download → Image** แล้วบันทึกลง `data/tableau/` ตามชื่อด้านล่าง

ขั้นตอนสำหรับ Tableau Desktop:

1. สร้างโฟลเดอร์ `data/tableau/` (ถ้ายังไม่มี)
2. คลิกแท็บ `Dashboard_Overview` → เมนู **Dashboard → Export Image…** → บันทึกเป็น `data/tableau/01_Dashboard_Overview.png`
3. คลิกแท็บ `Dashboard_Details` → **Dashboard → Export Image…** → `data/tableau/02_Dashboard_Details.png`
4. เปิดเว็บ เมนู **กราฟ Tableau** → จะมีแท็บ "Dashboard Overview (ภาพ)" / "Dashboard Details (ภาพ)" ทันที (ไม่ต้องรีสตาร์ต)

ชื่อไฟล์: ตัวเลขหน้าชื่อใช้เรียงลำดับแท็บ ส่วนที่เหลือเป็นชื่อแท็บ (`_` แสดงเป็นช่องว่าง) · รองรับ .png .jpg

---

## 5. Publish ขึ้น Tableau Public แล้วฝังบนเว็บ

1. **File → Save to Tableau Public As…** → ล็อกอิน → ตั้งชื่อ `Netflix Dashboard`
   (ถ้าใช้ Tableau Desktop ระบบจะถามให้ Extract ข้อมูล → ตกลง)
2. เบราว์เซอร์จะเปิดหน้า viz → ตรวจว่าตั้งเป็น Public (ไอคอนรูปตา) — ถ้าตั้ง Hidden เว็บจะฝังไม่ได้
3. กดปุ่ม **Share** → คัดลอก **Link** (รูปแบบ `https://public.tableau.com/views/NetflixDashboard/Dashboard1?...`)
4. เปิดเว็บของโปรเจกต์ → เมนู **ตั้งค่า** → **+ เพิ่มลิงก์** → ใส่ชื่อแท็บ เช่น `ภาพรวม` และวางลิงก์ → **บันทึก**
5. ถ้ามีหลาย Dashboard ให้คัดลอกลิงก์ของแต่ละ dashboard (เปลี่ยน tab ด้านบนของ viz แล้ว Share) มาใส่เพิ่ม
6. ไปที่เมนู **กราฟ Tableau** จะเห็นกราฟฝังอยู่ในหน้าเว็บ

---

## 6. เมื่อข้อมูลเปลี่ยน (เพิ่ม/ลบข้อมูลบนเว็บหรือใน Excel)

ไฟล์ Excel อัปเดตเองอัตโนมัติ แต่ Tableau เก็บสำเนาข้อมูล (extract) ไว้ ต้องอัปเดตด้วยมือ:

1. เปิด workbook ใน Tableau
2. เมนู **Data → Refresh All Extracts** (หรือคลิกขวา data source → **Refresh**) ทำทั้ง 2 data source
3. **File → Save** (ทับ `data/tableau_netflix2.twbx`) แล้ว **Export Image** ทับไฟล์เดิมใน `data/tableau/` (ข้อ 4.1)
4. ถ้า publish ไว้: **File → Save to Tableau Public** (ทับชื่อเดิม) — กราฟที่ฝังบนเว็บจะเปลี่ยนตาม (อาจใช้เวลาไม่กี่นาที)

> Tableau Public ฟรีไม่รองรับการต่อ MySQL หรืออัปเดตอัตโนมัติ จึงใช้ไฟล์ Excel เป็นแหล่งข้อมูล
> (ตรงกับลำดับของโจทย์: ข้อมูล → Excel วิเคราะห์ → Tableau ทำกราฟ → เว็บแสดงผล)

---

## ปัญหาที่พบบ่อย

| อาการ | วิธีแก้ |
|---|---|
| ค่าใน sheet4 เป็น 0 หรือว่าง | เปิดไฟล์ใน Excel → Save → แล้ว Refresh ใน Tableau |
| หน้าเว็บแสดงกล่องขาวว่าง | ตรวจว่า viz ตั้งเป็น Public และลิงก์ขึ้นต้นด้วย `https://public.tableau.com/views/` |
| ชื่อเรื่องภาษาเกาหลี/อาหรับเป็นกล่อง | เป็นเรื่องฟอนต์ของเครื่อง ไม่กระทบตัวเลข |
| `Group` เรียงลำดับผิด | คลิกขวา `Group` → Sort → Manual |
