"""
build_excel.py — สร้างไฟล์ Excel หลักของโปรเจกต์จากรายงานต้นฉบับของ Netflix

ใช้งาน:
    python scripts/build_excel.py source/Netflix_What_We_Watched_2026Jan-Jun_original.xlsx data/netflix_data.xlsx

ผลลัพธ์มี 4 sheet
    sheet1_RawData        ข้อมูลดิบ (TV + Film รวมกัน) — backend อ่าน/เขียน sheet นี้
    sheet2_Analysis       ข้อมูลที่คำนวณได้ (สูตร Excel ทั้งหมด)
    sheet3_Top10          Top 10 Film / TV (สูตร Excel)
    sheet4_TableauSummary สรุปค่าจาก sheet2 ในรูปตารางยาว (long format) สำหรับต่อ Tableau

สูตรทั้งหมดอ้างอิง sheet1 ผ่าน Named Range (D_Type, D_Views, ...) ครอบคลุมแถว 2–100000
จึงรองรับการเพิ่ม/ลบข้อมูลจากหน้าเว็บโดยไม่ต้องแก้สูตร
"""
import datetime
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula

MAX_ROW = 100000
RED = "B20710"
DARK = "221F1F"
LIGHT = "F5F5F1"

S1 = "sheet1_RawData"
S2 = "sheet2_Analysis"
S3 = "sheet3_Top10"
S4 = "sheet4_TableauSummary"

HEADERS = ["ID", "Type", "Title", "Available Globally", "Release Date",
           "Hours Viewed", "Runtime", "Runtime (Hours)", "Views"]
# ชื่อ Named Range -> คอลัมน์ใน sheet1
NAMES = {"D_ID": "A", "D_Type": "B", "D_Title": "C", "D_Global": "D", "D_Rel": "E",
         "D_Hours": "F", "D_Runtime": "G", "D_RtHours": "H", "D_Views": "I"}

HFONT = Font(bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor=RED)
SECFONT = Font(bold=True, size=13, color=RED)
SUBFILL = PatternFill("solid", fgColor="EDEDED")
THIN = Border(bottom=Side(style="thin", color="D0D0D0"))


def runtime_hours(rt):
    if not rt:
        return None
    h, m = rt.split(":")
    return round(int(h) + int(m) / 60, 2)


def read_source(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    rows = []
    for sheet, typ in (("Shows", "TV"), ("Movies", "Film")):
        for r in list(wb[sheet].iter_rows(values_only=True))[6:]:
            title, glob, rel, hrs, rt, views = r[1:7]
            if title is None or hrs is None:
                continue  # แถวว่าง / footnote
            if glob == "*":
                continue  # แถวยอดรวม "Other Shows" / "Other Movies"
            d = datetime.datetime.strptime(rel, "%Y-%m-%d").date() if rel else None
            rt = (rt or "").replace("*", "").strip() or None
            rows.append([typ, str(title), glob, d, int(hrs), rt, runtime_hours(rt), int(views)])
    return rows


def write_sheet1(ws, rows):
    ws.append(HEADERS)
    for i, r in enumerate(rows, start=1):
        ws.append([i] + r)
    for c in ws[1]:
        c.font, c.fill = HFONT, HFILL
        c.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=2, max_row=len(rows) + 1):
        row[4].number_format = "yyyy-mm-dd"
        row[5].number_format = "#,##0"
        row[7].number_format = "0.00"
        row[8].number_format = "#,##0"
    for col, w in zip("ABCDEFGHI", [8, 7, 55, 18, 14, 16, 10, 16, 14]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{len(rows) + 1}"


def header_row(ws, row, labels, col=2):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font, c.fill = HFONT, HFILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)


def section(ws, row, text, note=None):
    ws.cell(row=row, column=2, value=text).font = SECFONT
    if note:
        ws.cell(row=row + 1, column=2, value=note).font = Font(italic=True, color="777777", size=9)
        return row + 2
    return row + 1


def arr(ws, ref, formula):
    ws[ref] = ArrayFormula(ref, formula)


def build_sheet2(ws):
    """sheet2: ข้อมูลที่คำนวณได้ — คืน dict ของตำแหน่งเซลล์ที่ sheet4 จะอ้างอิง"""
    refs = {}
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 30
    for col in "CDEFGHIJ":
        ws.column_dimensions[col].width = 17

    ws["B1"] = "Netflix What We Watched (ม.ค.–มิ.ย. 2026) — ข้อมูลที่คำนวณได้"
    ws["B1"].font = Font(bold=True, size=16, color=DARK)
    ws["B2"] = "ทุกค่าคำนวณด้วยสูตร Excel จาก sheet1_RawData อัตโนมัติ (เพิ่ม/ลบข้อมูลแล้วค่าจะอัปเดตเอง)"
    ws["B2"].font = Font(italic=True, color="777777")

    # ---------- 1. KPI ----------
    r = section(ws, 4, "1. ตัวเลขภาพรวม (KPI)")
    header_row(ws, r, ["ตัวชี้วัด", "ค่า"])
    kpis = [
        ("จำนวนเรื่องทั้งหมด", '=COUNTA(D_Title)', "#,##0", "kpi_titles"),
        ("ชั่วโมงดูรวม (Hours Viewed)", '=SUM(D_Hours)', "#,##0", "kpi_hours"),
        ("ยอดดูรวม (Views)", '=SUM(D_Views)', "#,##0", "kpi_views"),
        ("ยอดดูเฉลี่ยต่อเรื่อง", '=AVERAGE(D_Views)', "#,##0", "kpi_avg_views"),
        ("ยอดดูมัธยฐานต่อเรื่อง", '=MEDIAN(D_Views)', "#,##0", "kpi_median_views"),
        ("ชั่วโมงดูเฉลี่ยต่อเรื่อง", '=AVERAGE(D_Hours)', "#,##0", "kpi_avg_hours"),
        ("ส่วนเบี่ยงเบนมาตรฐานของยอดดู", '=STDEV(D_Views)', "#,##0", "kpi_sd_views"),
    ]
    for i, (lab, f, fmt, key) in enumerate(kpis, start=1):
        ws.cell(row=r + i, column=2, value=lab)
        c = ws.cell(row=r + i, column=3, value=f)
        c.number_format = fmt
        refs[key] = f"C{r + i}"
    r += len(kpis) + 2

    # ---------- 2. TV vs Film ----------
    r = section(ws, r, "2. เปรียบเทียบ TV กับ Film")
    header_row(ws, r, ["ประเภท", "จำนวนเรื่อง", "ชั่วโมงดูรวม", "% ของชั่วโมงดู", "ยอดดูรวม",
                       "ยอดดูเฉลี่ย", "ยอดดูมัธยฐาน", "ยอดดูสูงสุด", "Runtime เฉลี่ย (ชม.)"])
    top = r + 1
    for i, t in enumerate(["TV", "Film"]):
        rr = top + i
        ws.cell(row=rr, column=2, value=t)
        ws[f"C{rr}"] = f'=COUNTIF(D_Type,B{rr})'
        ws[f"D{rr}"] = f'=SUMIF(D_Type,B{rr},D_Hours)'
        ws[f"E{rr}"] = f'=D{rr}/SUM(D_Hours)'
        ws[f"F{rr}"] = f'=SUMIF(D_Type,B{rr},D_Views)'
        ws[f"G{rr}"] = f'=AVERAGEIF(D_Type,B{rr},D_Views)'
        arr(ws, f"H{rr}", f'=MEDIAN(IF(D_Type=B{rr},D_Views))')
        arr(ws, f"I{rr}", f'=MAX(IF(D_Type=B{rr},D_Views))')
        ws[f"J{rr}"] = f'=AVERAGEIF(D_Type,B{rr},D_RtHours)'
        refs[f"type_{t}"] = rr
    rr = top + 2
    ws.cell(row=rr, column=2, value="รวม").font = Font(bold=True)
    ws[f"C{rr}"] = f"=SUM(C{top}:C{top + 1})"
    ws[f"D{rr}"] = f"=SUM(D{top}:D{top + 1})"
    ws[f"E{rr}"] = f"=SUM(E{top}:E{top + 1})"
    ws[f"F{rr}"] = f"=SUM(F{top}:F{top + 1})"
    ws[f"G{rr}"] = "=AVERAGE(D_Views)"
    ws[f"H{rr}"] = "=MEDIAN(D_Views)"
    ws[f"I{rr}"] = "=MAX(D_Views)"
    ws[f"J{rr}"] = "=AVERAGE(D_RtHours)"
    fmt_block(ws, top, rr, {"C": "#,##0", "D": "#,##0", "E": "0.0%", "F": "#,##0", "G": "#,##0",
                            "H": "#,##0", "I": "#,##0", "J": "0.00"})
    r = rr + 2

    # ---------- 3. Netflix Title vs Non-Netflix ----------
    r = section(ws, r, "3. Netflix Title (มี Release Date) กับ Non-Netflix (ไม่มี Release Date)",
                "เกณฑ์แทน: Netflix ระบุ premiere date เฉพาะคอนเทนต์ของ Netflix — ไม่ใช่ป้ายกำกับทางการ")
    header_row(ws, r, ["กลุ่ม", "จำนวนเรื่อง", "ชั่วโมงดูรวม", "% ของชั่วโมงดู", "ยอดดูเฉลี่ย", "ยอดดูมัธยฐาน"])
    top = r + 1
    ws.cell(row=top, column=2, value="Netflix Title")
    ws[f"C{top}"] = '=COUNTIFS(D_Title,"<>",D_Rel,"<>")'
    ws[f"D{top}"] = '=SUMIFS(D_Hours,D_Title,"<>",D_Rel,"<>")'
    ws[f"F{top}"] = '=AVERAGEIFS(D_Views,D_Title,"<>",D_Rel,"<>")'
    arr(ws, f"G{top}", '=MEDIAN(IF((D_Title<>"")*(D_Rel<>""),D_Views))')
    ws.cell(row=top + 1, column=2, value="Non-Netflix")
    ws[f"C{top + 1}"] = '=COUNTIFS(D_Title,"<>",D_Rel,"")'
    ws[f"D{top + 1}"] = '=SUMIFS(D_Hours,D_Title,"<>",D_Rel,"")'
    ws[f"F{top + 1}"] = '=AVERAGEIFS(D_Views,D_Title,"<>",D_Rel,"")'
    arr(ws, f"G{top + 1}", '=MEDIAN(IF((D_Title<>"")*(D_Rel=""),D_Views))')
    for rr in (top, top + 1):
        ws[f"E{rr}"] = f"=D{rr}/SUM(D_Hours)"
    refs["origin_netflix"], refs["origin_other"] = top, top + 1
    fmt_block(ws, top, top + 1, {"C": "#,##0", "D": "#,##0", "E": "0.0%", "F": "#,##0", "G": "#,##0"})
    r = top + 3

    # ---------- 4. Available Globally ----------
    r = section(ws, r, "4. ฉายทั่วโลก (Available Globally) กับ ไม่ทั่วโลก")
    header_row(ws, r, ["Available Globally", "จำนวนเรื่อง", "ชั่วโมงดูรวม", "% ของชั่วโมงดู", "ยอดดูเฉลี่ย", "ยอดดูมัธยฐาน"])
    top = r + 1
    for i, g in enumerate(["Yes", "No"]):
        rr = top + i
        ws.cell(row=rr, column=2, value=g)
        ws[f"C{rr}"] = f'=COUNTIF(D_Global,B{rr})'
        ws[f"D{rr}"] = f'=SUMIF(D_Global,B{rr},D_Hours)'
        ws[f"E{rr}"] = f'=D{rr}/SUM(D_Hours)'
        ws[f"F{rr}"] = f'=AVERAGEIF(D_Global,B{rr},D_Views)'
        arr(ws, f"G{rr}", f'=MEDIAN(IF(D_Global=B{rr},D_Views))')
        refs[f"global_{g}"] = rr
    fmt_block(ws, top, top + 1, {"C": "#,##0", "D": "#,##0", "E": "0.0%", "F": "#,##0", "G": "#,##0"})
    r = top + 3

    # ---------- 5. Runtime groups ----------
    r = section(ws, r, "5. ความยาว (Runtime) กับยอดดู",
                "TV ใช้ความยาวรวมทั้งซีซัน, Film ใช้ความยาวหนัง จึงแบ่งช่วงคนละแบบ (ไม่นับแถวที่ไม่มี Runtime)")
    header_row(ws, r, ["ช่วงความยาว", "ประเภท", "ต่ำสุด (ชม.)", "สูงสุด (ชม.)", "จำนวนเรื่อง",
                       "ยอดดูเฉลี่ย", "ยอดดูมัธยฐาน"])
    groups = [("Film <1.5h", "Film", 0, 1.4999), ("Film 1.5-2h", "Film", 1.5, 2),
              ("Film >2h", "Film", 2.0001, 999), ("TV <5h", "TV", 0, 4.9999),
              ("TV 5-10h", "TV", 5, 10), ("TV >10h", "TV", 10.0001, 9999)]
    top = r + 1
    refs["runtime"] = []
    for i, (lab, t, lo, hi) in enumerate(groups):
        rr = top + i
        ws.cell(row=rr, column=2, value=lab)
        ws.cell(row=rr, column=3, value=t)
        ws.cell(row=rr, column=4, value=lo)
        ws.cell(row=rr, column=5, value=hi)
        crit = f'D_Type,C{rr},D_RtHours,">="&D{rr},D_RtHours,"<="&E{rr}'
        ws[f"F{rr}"] = f'=COUNTIFS({crit})'
        ws[f"G{rr}"] = f'=IFERROR(AVERAGEIFS(D_Views,{crit}),0)'
        arr(ws, f"H{rr}", f'=IFERROR(MEDIAN(IF((D_Type=C{rr})*(D_RtHours<>"")*(D_RtHours>=D{rr})*(D_RtHours<=E{rr}),D_Views)),0)')
        refs["runtime"].append(rr)
    fmt_block(ws, top, top + len(groups) - 1, {"D": "0.00", "E": "0.00", "F": "#,##0", "G": "#,##0", "H": "#,##0"})
    r = top + len(groups) + 1

    # ---------- 6. Release year ----------
    r = section(ws, r, "6. Netflix Title ตามปีที่ออกฉาย",
                "ดูว่าเรื่องเก่ายังทำยอดดูได้แค่ไหนในครึ่งปีแรก 2026")
    header_row(ws, r, ["ปีที่ออกฉาย", "ปีเริ่ม", "ปีสิ้นสุด", "จำนวนเรื่อง", "ชั่วโมงดูรวม", "ยอดดูเฉลี่ย"])
    top = r + 1
    years = [("≤2015", 1900, 2015)] + [(str(y), y, y) for y in range(2016, 2027)]
    refs["year"] = []
    for i, (lab, y0, y1) in enumerate(years):
        rr = top + i
        ws.cell(row=rr, column=2, value=lab)
        ws.cell(row=rr, column=3, value=y0)
        ws.cell(row=rr, column=4, value=y1)
        crit = f'D_Rel,">="&DATE(C{rr},1,1),D_Rel,"<="&DATE(D{rr},12,31)'
        ws[f"E{rr}"] = f'=COUNTIFS({crit})'
        ws[f"F{rr}"] = f'=SUMIFS(D_Hours,{crit})'
        ws[f"G{rr}"] = f'=IFERROR(AVERAGEIFS(D_Views,{crit}),0)'
        refs["year"].append(rr)
    fmt_block(ws, top, top + len(years) - 1, {"E": "#,##0", "F": "#,##0", "G": "#,##0"})
    r = top + len(years) + 1

    # ---------- 7. Concentration ----------
    r = section(ws, r, "7. การกระจุกตัวของยอดดู (Pareto)",
                "เรื่องที่ยอดดูสูงสุด X% ครองสัดส่วนชั่วโมงดูรวมกี่ %")
    header_row(ws, r, ["Top % ของเรื่อง", "สัดส่วน", "จำนวนเรื่อง", "% ของชั่วโมงดูรวม"])
    top = r + 1
    refs["pareto"] = []
    for i, p in enumerate([0.01, 0.05, 0.10, 0.20, 0.50]):
        rr = top + i
        ws.cell(row=rr, column=2, value=f"Top {int(p * 100)}%")
        ws.cell(row=rr, column=3, value=p).number_format = "0%"
        ws[f"D{rr}"] = f"=ROUNDUP(COUNT(D_Hours)*C{rr},0)"
        ws[f"E{rr}"] = f'=SUMIF(D_Hours,">="&LARGE(D_Hours,D{rr}))/SUM(D_Hours)'
        refs["pareto"].append(rr)
    fmt_block(ws, top, top + 4, {"D": "#,##0", "E": "0.0%"})
    ws.freeze_panes = "A3"
    return refs


def fmt_block(ws, r0, r1, fmts):
    for rr in range(r0, r1 + 1):
        for col, f in fmts.items():
            ws[f"{col}{rr}"].number_format = f
        for col in "BCDEFGHIJ":
            ws[f"{col}{rr}"].border = THIN


def build_sheet3(ws):
    ws.column_dimensions["A"].width = 2
    widths = {"B": 7, "C": 50, "D": 15, "E": 16, "F": 13, "G": 10, "H": 12}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws["B1"] = "Top 10 เรื่องที่ยอดดู (Views) สูงสุด"
    ws["B1"].font = Font(bold=True, size=16, color=DARK)
    ws["B2"] = "คำนวณด้วยสูตร LARGE + INDEX/SMALL (รองรับยอดดูเท่ากัน) — คอลัมน์ H คือเลขแถวใน sheet1"
    ws["B2"].font = Font(italic=True, color="777777")
    start = 4
    for t in ("Film", "TV"):
        ws.cell(row=start, column=2, value=f"Top 10 {t}").font = SECFONT
        ws.cell(row=start, column=9, value=t).font = Font(color="FFFFFF")  # ค่าที่ใช้ในสูตร (ซ่อนด้วยสีขาว)
        tcell = f"$I${start}"
        header_row(ws, start + 1, ["อันดับ", "Title", "Views", "Hours Viewed", "Release Date", "Global", "แถวใน sheet1"])
        f0 = start + 2
        for k in range(1, 11):
            rr = f0 + k - 1
            ws.cell(row=rr, column=2, value=k)
            arr(ws, f"D{rr}", f"=LARGE(IF(D_Type={tcell},D_Views),B{rr})")
            arr(ws, f"H{rr}", f"=SMALL(IF((D_Type={tcell})*(D_Views=D{rr}),ROW(D_Views)),COUNTIF(D${f0}:D{rr},D{rr}))")
            ws[f"C{rr}"] = f"=INDEX({S1}!$C:$C,H{rr})"
            ws[f"E{rr}"] = f"=INDEX({S1}!$F:$F,H{rr})"
            ws[f"F{rr}"] = f'=IF(INDEX({S1}!$E:$E,H{rr})="","-",INDEX({S1}!$E:$E,H{rr}))'
            ws[f"G{rr}"] = f"=INDEX({S1}!$D:$D,H{rr})"
            ws[f"D{rr}"].number_format = "#,##0"
            ws[f"E{rr}"].number_format = "#,##0"
            ws[f"F{rr}"].number_format = "yyyy-mm-dd"
            for col in "BCDEFGH":
                ws[f"{col}{rr}"].border = THIN
        start = f0 + 12
    return {"Film": 6, "TV": 6 + 14}


def build_sheet4(ws, refs, top_start):
    """ตารางยาวสำหรับ Tableau: Section | Group | Metric | Value"""
    header = ["Section", "Group", "Metric", "Value"]
    ws.append(header)
    for c in ws[1]:
        c.font, c.fill = HFONT, HFILL
    rows = []
    k = refs
    for key, label in [("kpi_titles", "Total Titles"), ("kpi_hours", "Total Hours Viewed"),
                       ("kpi_views", "Total Views"), ("kpi_avg_views", "Average Views"),
                       ("kpi_median_views", "Median Views"), ("kpi_avg_hours", "Average Hours Viewed")]:
        rows.append(("KPI", "All", label, f"={S2}!{k[key]}"))
    for t in ("TV", "Film"):
        rr = k[f"type_{t}"]
        for col, m in [("C", "Titles"), ("D", "Hours Viewed"), ("E", "Share of Hours"), ("F", "Views"),
                       ("G", "Average Views"), ("H", "Median Views"), ("J", "Average Runtime (h)")]:
            rows.append(("Type", t, m, f"={S2}!{col}{rr}"))
    for key, g in [("origin_netflix", "Netflix Title"), ("origin_other", "Non-Netflix")]:
        rr = k[key]
        for col, m in [("C", "Titles"), ("D", "Hours Viewed"), ("E", "Share of Hours"),
                       ("F", "Average Views"), ("G", "Median Views")]:
            rows.append(("Origin", g, m, f"={S2}!{col}{rr}"))
    for g in ("Yes", "No"):
        rr = k[f"global_{g}"]
        for col, m in [("C", "Titles"), ("D", "Hours Viewed"), ("E", "Share of Hours"),
                       ("F", "Average Views"), ("G", "Median Views")]:
            rows.append(("Available Globally", g, m, f"={S2}!{col}{rr}"))
    for rr in k["runtime"]:
        for col, m in [("F", "Titles"), ("G", "Average Views"), ("H", "Median Views")]:
            rows.append(("Runtime Group", f"={S2}!B{rr}", m, f"={S2}!{col}{rr}"))
    for rr in k["year"]:
        for col, m in [("E", "Titles"), ("F", "Hours Viewed"), ("G", "Average Views")]:
            rows.append(("Release Year", f"={S2}!B{rr}", m, f"={S2}!{col}{rr}"))
    for rr in k["pareto"]:
        rows.append(("Pareto", f"={S2}!B{rr}", "Share of Hours", f"={S2}!E{rr}"))
    for t, s in top_start.items():
        for i in range(10):
            rr = s + i
            rows.append((f"Top10 {t}", f"={S3}!C{rr}", "Views", f"={S3}!D{rr}"))
    for r in rows:
        ws.append(list(r))
    for col, w in zip("ABCD", [20, 50, 22, 18]):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2):
        row[3].number_format = "#,##0.00"
    ws.freeze_panes = "A2"


def main(src, dst):
    rows = read_source(src)
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = S1
    write_sheet1(ws1, rows)
    for name, col in NAMES.items():
        wb.defined_names[name] = DefinedName(name, attr_text=f"{S1}!${col}$2:${col}${MAX_ROW}")
    refs = build_sheet2(wb.create_sheet(S2))
    top_start = build_sheet3(wb.create_sheet(S3))
    build_sheet4(wb.create_sheet(S4), refs, top_start)
    wb.calculation.fullCalcOnLoad = True  # ให้ Excel คำนวณใหม่ทุกครั้งที่เปิด
    wb.save(dst)
    print(f"wrote {dst}: {len(rows)} rows")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
