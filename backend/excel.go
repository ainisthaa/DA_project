package main

import (
	"fmt"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	SheetRaw     = "sheet1_RawData"
	SheetSummary = "sheet4_TableauSummary"
)

// ลำดับคอลัมน์ที่ backend เขียนลง sheet1 (ตอนอ่านจะหาจากชื่อหัวคอลัมน์ จึงสลับคอลัมน์ใน Excel ได้)
var rawHeaders = []string{"ID", "Type", "Title", "Available Globally", "Release Date", "Hours Viewed", "Runtime", "Runtime (Hours)", "Views"}

type RowError struct {
	Row     int    `json:"row"`
	Message string `json:"message"`
}

// ReadExcel อ่าน sheet1 ทั้งหมด คืนแถวที่ถูกต้องและรายการแถวที่ผิด
func ReadExcel(path string) ([]Title, []RowError, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, nil, fmt.Errorf("เปิดไฟล์ Excel ไม่ได้: %w", err)
	}
	defer f.Close()
	if idx, _ := f.GetSheetIndex(SheetRaw); idx < 0 {
		return nil, nil, fmt.Errorf("ไม่พบ sheet ชื่อ %s", SheetRaw)
	}
	rows, err := f.GetRows(SheetRaw, excelize.Options{RawCellValue: true})
	if err != nil {
		return nil, nil, err
	}
	if len(rows) == 0 {
		return nil, nil, fmt.Errorf("sheet %s ว่าง", SheetRaw)
	}
	col := map[string]int{}
	for i, h := range rows[0] {
		col[strings.ToLower(strings.TrimSpace(h))] = i
	}
	for _, h := range []string{"type", "title", "available globally", "release date", "hours viewed", "runtime", "views"} {
		if _, ok := col[h]; !ok {
			return nil, nil, fmt.Errorf("sheet1 ไม่มีคอลัมน์ %q (แถวแรกต้องเป็นชื่อคอลัมน์)", h)
		}
	}
	get := func(r []string, name string) string {
		i, ok := col[name]
		if !ok || i >= len(r) {
			return ""
		}
		return strings.TrimSpace(r[i])
	}
	var out []Title
	var errs []RowError
	seenID := map[int64]int{}
	for n, r := range rows[1:] {
		excelRow := n + 2
		if strings.Join(r, "") == "" {
			continue // แถวว่าง
		}
		t := Title{Type: get(r, "type"), Title: get(r, "title"), AvailableGlobally: get(r, "available globally")}
		bad := func(msg string) { errs = append(errs, RowError{excelRow, msg}) }
		if s := get(r, "id"); s != "" {
			id, err := strconv.ParseFloat(s, 64)
			if err != nil || id <= 0 || id != math.Trunc(id) {
				bad(fmt.Sprintf("ID ไม่ถูกต้อง (%q) — ปล่อยว่างได้ถ้าเป็นแถวใหม่", s))
				continue
			}
			t.ID = int64(id)
			if prev, dup := seenID[t.ID]; dup {
				bad(fmt.Sprintf("ID %d ซ้ำกับแถว %d", t.ID, prev))
				continue
			}
			seenID[t.ID] = excelRow
		}
		if s := get(r, "release date"); s != "" {
			ds, err := excelDate(s)
			if err != nil {
				bad(err.Error())
				continue
			}
			t.ReleaseDate = &ds
		}
		var err error
		if t.HoursViewed, err = parseInt(get(r, "hours viewed")); err != nil {
			bad("Hours Viewed: " + err.Error())
			continue
		}
		if t.Views, err = parseInt(get(r, "views")); err != nil {
			bad("Views: " + err.Error())
			continue
		}
		if s := get(r, "runtime"); s != "" {
			rs := excelRuntime(s)
			t.Runtime = &rs
		}
		if err := t.Normalize(); err != nil {
			bad(err.Error())
			continue
		}
		out = append(out, t)
	}
	return out, errs, nil
}

func parseInt(s string) (int64, error) {
	s = strings.ReplaceAll(strings.ReplaceAll(s, ",", ""), " ", "")
	if s == "" {
		return 0, fmt.Errorf("ต้องไม่ว่าง")
	}
	v, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return 0, fmt.Errorf("ต้องเป็นตัวเลข (ได้ %q)", s)
	}
	return int64(math.Round(v)), nil
}

// excelDate รับได้ทั้งเลข serial ของ Excel (เช่น 46087) และข้อความ YYYY-MM-DD
func excelDate(s string) (string, error) {
	if v, err := strconv.ParseFloat(s, 64); err == nil {
		d, err := excelize.ExcelDateToTime(v, false)
		if err != nil {
			return "", fmt.Errorf("Release Date ไม่ถูกต้อง (%q)", s)
		}
		return d.Format("2006-01-02"), nil
	}
	d, err := parseDate(s)
	if err != nil {
		return "", err
	}
	return d.Format("2006-01-02"), nil
}

// excelRuntime: ถ้าผู้ใช้พิมพ์ 1:30 ใน Excel โดยตรง Excel จะแปลงเป็นเวลา (เศษของวัน เช่น 0.0625) — แปลงกลับเป็น H:MM
func excelRuntime(s string) string {
	if v, err := strconv.ParseFloat(s, 64); err == nil {
		mins := int(math.Round(v * 24 * 60))
		return fmt.Sprintf("%d:%02d", mins/60, mins%60)
	}
	return s
}

// WriteExcel เขียนข้อมูลทั้งหมดลง sheet1 โดยไม่แตะ sheet อื่น (สูตรใน sheet2–4 อ้างอิง Named Range จึงยังทำงาน)
func WriteExcel(path string, titles []Title) error {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return fmt.Errorf("เปิดไฟล์ Excel ไม่ได้: %w", err)
	}
	defer f.Close()
	if idx, _ := f.GetSheetIndex(SheetRaw); idx < 0 {
		return fmt.Errorf("ไม่พบ sheet ชื่อ %s", SheetRaw)
	}
	hdr, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"B20710"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "center"},
	})
	dateFmt := "yyyy-mm-dd"
	dateSt, _ := f.NewStyle(&excelize.Style{CustomNumFmt: &dateFmt})
	numSt, _ := f.NewStyle(&excelize.Style{NumFmt: 3})  // #,##0
	decSt, _ := f.NewStyle(&excelize.Style{NumFmt: 2})  // 0.00

	sw, err := f.NewStreamWriter(SheetRaw)
	if err != nil {
		return err
	}
	for i, w := range []float64{8, 7, 55, 18, 14, 16, 10, 16, 14} {
		if err := sw.SetColWidth(i+1, i+1, w); err != nil {
			return err
		}
	}
	if err := sw.SetPanes(&excelize.Panes{Freeze: true, YSplit: 1, TopLeftCell: "A2", ActivePane: "bottomLeft"}); err != nil {
		return err
	}
	head := make([]any, len(rawHeaders))
	for i, h := range rawHeaders {
		head[i] = excelize.Cell{StyleID: hdr, Value: h}
	}
	if err := sw.SetRow("A1", head); err != nil {
		return err
	}
	for i, t := range titles {
		var rel, rt, rth any
		if t.ReleaseDate != nil {
			if d, err := time.Parse("2006-01-02", *t.ReleaseDate); err == nil {
				rel = excelize.Cell{StyleID: dateSt, Value: d}
			}
		}
		if t.Runtime != nil {
			rt = *t.Runtime
		}
		if t.RuntimeHours != nil {
			rth = excelize.Cell{StyleID: decSt, Value: *t.RuntimeHours}
		}
		row := []any{t.ID, t.Type, t.Title, t.AvailableGlobally, rel,
			excelize.Cell{StyleID: numSt, Value: t.HoursViewed}, rt, rth, excelize.Cell{StyleID: numSt, Value: t.Views}}
		cell, _ := excelize.CoordinatesToCellName(1, i+2)
		if err := sw.SetRow(cell, row); err != nil {
			return err
		}
	}
	if err := sw.Flush(); err != nil {
		return err
	}
	last := len(titles) + 1
	_ = f.AutoFilter(SheetRaw, fmt.Sprintf("A1:I%d", last), nil)
	full := true
	_ = f.SetCalcProps(&excelize.CalcPropsOptions{FullCalcOnLoad: &full})

	// เขียนลงไฟล์ชั่วคราวแล้ว rename เพื่อไม่ให้ไฟล์เสียถ้าเขียนไม่จบ
	tmp := filepath.Join(filepath.Dir(path), ".tmp-"+filepath.Base(path)) // ต้องลงท้าย .xlsx (excelize ตรวจนามสกุล)
	if err := f.SaveAs(tmp); err != nil {
		return err
	}
	if err := os.Rename(tmp, path); err != nil {
		os.Remove(tmp)
		return fmt.Errorf("บันทึกไฟล์ Excel ไม่ได้ (ถ้าเปิดไฟล์ค้างไว้ใน Excel ให้ปิดก่อน): %w", err)
	}
	return nil
}

// SummaryRow คือ 1 แถวจาก sheet4_TableauSummary (ค่าที่ Excel คำนวณไว้)
type SummaryRow struct {
	Section string   `json:"section"`
	Group   string   `json:"group"`
	Metric  string   `json:"metric"`
	Value   *float64 `json:"value"`
}

// ReadSummary อ่านค่าที่คำนวณแล้ว (cached value) จาก sheet4
// ถ้าไฟล์ยังไม่ถูกคำนวณใหม่หลังแก้ข้อมูล ค่าจะว่าง → stale = true
func ReadSummary(path string) ([]SummaryRow, bool, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, false, err
	}
	defer f.Close()
	rows, err := f.GetRows(SheetSummary, excelize.Options{RawCellValue: true})
	if err != nil {
		return nil, false, err
	}
	out := []SummaryRow{}
	stale := false
	for _, r := range rows[1:] {
		for len(r) < 4 {
			r = append(r, "")
		}
		sr := SummaryRow{Section: r[0], Group: r[1], Metric: r[2]}
		if v, err := strconv.ParseFloat(r[3], 64); err == nil {
			sr.Value = &v
		} else {
			stale = true
		}
		if sr.Group == "" {
			stale = true
		}
		out = append(out, sr)
	}
	return out, stale, nil
}
