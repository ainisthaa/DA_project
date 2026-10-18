package main

import (
	"fmt"
	"math"
	"regexp"
	"strconv"
	"strings"
	"time"
)

// Title คือ 1 แถวของข้อมูล (1 เรื่อง) — ตรงกับ 1 แถวใน sheet1_RawData และ 1 แถวในตาราง titles
type Title struct {
	ID                int64    `json:"id"`
	Type              string   `json:"type"`              // "TV" หรือ "Film"
	Title             string   `json:"title"`             // ชื่อเรื่อง
	AvailableGlobally string   `json:"availableGlobally"` // "Yes" หรือ "No"
	ReleaseDate       *string  `json:"releaseDate"`       // YYYY-MM-DD หรือ null
	HoursViewed       int64    `json:"hoursViewed"`
	Runtime           *string  `json:"runtime"`      // H:MM หรือ null
	RuntimeHours      *float64 `json:"runtimeHours"` // คำนวณจาก Runtime อัตโนมัติ
	Views             int64    `json:"views"`
	UpdatedAt         string   `json:"updatedAt,omitempty"`
}

var runtimeRe = regexp.MustCompile(`^(\d{1,3}):([0-5]\d)$`)

// Normalize ตรวจความถูกต้องและจัดรูปแบบข้อมูลก่อนบันทึก (ใช้ทั้งข้อมูลจากเว็บและจาก Excel)
func (t *Title) Normalize() error {
	t.Title = strings.TrimSpace(t.Title)
	if t.Title == "" {
		return fmt.Errorf("Title ต้องไม่ว่าง")
	}
	if len([]rune(t.Title)) > 500 {
		return fmt.Errorf("Title ยาวเกิน 500 ตัวอักษร")
	}
	switch strings.ToLower(strings.TrimSpace(t.Type)) {
	case "tv", "show", "shows", "series":
		t.Type = "TV"
	case "film", "movie", "movies":
		t.Type = "Film"
	default:
		return fmt.Errorf("Type ต้องเป็น TV หรือ Film (ได้ %q)", t.Type)
	}
	switch strings.ToLower(strings.TrimSpace(t.AvailableGlobally)) {
	case "yes", "y", "true", "1":
		t.AvailableGlobally = "Yes"
	case "no", "n", "false", "0", "":
		t.AvailableGlobally = "No"
	default:
		return fmt.Errorf("Available Globally ต้องเป็น Yes หรือ No (ได้ %q)", t.AvailableGlobally)
	}
	if t.ReleaseDate != nil {
		s := strings.TrimSpace(*t.ReleaseDate)
		if s == "" {
			t.ReleaseDate = nil
		} else {
			d, err := parseDate(s)
			if err != nil {
				return err
			}
			ds := d.Format("2006-01-02")
			t.ReleaseDate = &ds
		}
	}
	if t.HoursViewed < 0 || t.Views < 0 {
		return fmt.Errorf("Hours Viewed และ Views ต้องไม่ติดลบ")
	}
	t.RuntimeHours = nil
	if t.Runtime != nil {
		s := strings.TrimSpace(strings.ReplaceAll(*t.Runtime, "*", ""))
		if s == "" {
			t.Runtime = nil
		} else {
			m := runtimeRe.FindStringSubmatch(s)
			if m == nil {
				return fmt.Errorf("Runtime ต้องอยู่ในรูปแบบ H:MM เช่น 1:45 (ได้ %q)", s)
			}
			h, _ := strconv.Atoi(m[1])
			mi, _ := strconv.Atoi(m[2])
			rs := fmt.Sprintf("%d:%02d", h, mi)
			t.Runtime = &rs
			rh := math.Round((float64(h)+float64(mi)/60)*100) / 100
			t.RuntimeHours = &rh
		}
	}
	return nil
}

func parseDate(s string) (time.Time, error) {
	for _, layout := range []string{"2006-01-02", "2006/01/02", "02/01/2006", "2/1/2006", "2006-01-02 15:04:05", time.RFC3339} {
		if d, err := time.Parse(layout, s); err == nil {
			return d, nil
		}
	}
	return time.Time{}, fmt.Errorf("Release Date ต้องเป็นรูปแบบ YYYY-MM-DD (ได้ %q)", s)
}

func strPtr(s string) *string { return &s }
