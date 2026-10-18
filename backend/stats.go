package main

import (
	"fmt"
	"math"
	"sort"
)

// Stats คือผลวิเคราะห์ชุดเดียวกับ sheet2_Analysis แต่คำนวณสดจากฐานข้อมูล
// (ใช้ตรวจเทียบกับค่าที่ Excel คำนวณ และแสดงบนหน้าเว็บได้ทันทีแม้ Excel ยังไม่คำนวณใหม่)
type Stats struct {
	KPI      KPI          `json:"kpi"`
	ByType   []GroupStat  `json:"byType"`
	ByOrigin []GroupStat  `json:"byOrigin"`
	ByGlobal []GroupStat  `json:"byGlobal"`
	Runtime  []GroupStat  `json:"runtime"`
	ByYear   []GroupStat  `json:"byYear"`
	Pareto   []ParetoStat `json:"pareto"`
	Top10    TopLists     `json:"top10"`
}

type KPI struct {
	Titles       int     `json:"titles"`
	HoursViewed  int64   `json:"hoursViewed"`
	Views        int64   `json:"views"`
	AvgViews     float64 `json:"avgViews"`
	MedianViews  float64 `json:"medianViews"`
	AvgHours     float64 `json:"avgHours"`
	StdDevViews  float64 `json:"stdDevViews"`
}

type GroupStat struct {
	Group       string  `json:"group"`
	Titles      int     `json:"titles"`
	HoursViewed int64   `json:"hoursViewed"`
	ShareHours  float64 `json:"shareHours"`
	Views       int64   `json:"views"`
	AvgViews    float64 `json:"avgViews"`
	MedianViews float64 `json:"medianViews"`
	MaxViews    int64   `json:"maxViews"`
	AvgRuntime  float64 `json:"avgRuntime"`
}

type ParetoStat struct {
	Label      string  `json:"label"`
	Titles     int     `json:"titles"`
	ShareHours float64 `json:"shareHours"`
}

type TopLists struct {
	Film []Title `json:"film"`
	TV   []Title `json:"tv"`
}

func median(v []int64) float64 {
	if len(v) == 0 {
		return 0
	}
	s := append([]int64(nil), v...)
	sort.Slice(s, func(i, j int) bool { return s[i] < s[j] })
	n := len(s)
	if n%2 == 1 {
		return float64(s[n/2])
	}
	return float64(s[n/2-1]+s[n/2]) / 2
}

func group(name string, ts []Title, totalHours int64) GroupStat {
	g := GroupStat{Group: name, Titles: len(ts)}
	views := make([]int64, 0, len(ts))
	var rtSum float64
	rtN := 0
	for _, t := range ts {
		g.HoursViewed += t.HoursViewed
		g.Views += t.Views
		views = append(views, t.Views)
		if t.Views > g.MaxViews {
			g.MaxViews = t.Views
		}
		if t.RuntimeHours != nil {
			rtSum += *t.RuntimeHours
			rtN++
		}
	}
	if len(ts) > 0 {
		g.AvgViews = float64(g.Views) / float64(len(ts))
	}
	if totalHours > 0 {
		g.ShareHours = float64(g.HoursViewed) / float64(totalHours)
	}
	if rtN > 0 {
		g.AvgRuntime = rtSum / float64(rtN)
	}
	g.MedianViews = median(views)
	return g
}

func filter(ts []Title, keep func(Title) bool) []Title {
	out := []Title{}
	for _, t := range ts {
		if keep(t) {
			out = append(out, t)
		}
	}
	return out
}

func ComputeStats(ts []Title) Stats {
	var st Stats
	views := make([]int64, 0, len(ts))
	var total int64
	for _, t := range ts {
		total += t.HoursViewed
		st.KPI.Views += t.Views
		views = append(views, t.Views)
	}
	st.KPI.Titles = len(ts)
	st.KPI.HoursViewed = total
	if len(ts) > 0 {
		st.KPI.AvgViews = float64(st.KPI.Views) / float64(len(ts))
		st.KPI.AvgHours = float64(total) / float64(len(ts))
	}
	st.KPI.MedianViews = median(views)
	if len(ts) > 1 { // sample standard deviation (เหมือน STDEV ของ Excel)
		var ss float64
		for _, v := range views {
			d := float64(v) - st.KPI.AvgViews
			ss += d * d
		}
		st.KPI.StdDevViews = math.Sqrt(ss / float64(len(ts)-1))
	}

	for _, typ := range []string{"TV", "Film"} {
		typ := typ
		st.ByType = append(st.ByType, group(typ, filter(ts, func(t Title) bool { return t.Type == typ }), total))
	}
	st.ByOrigin = []GroupStat{
		group("Netflix Title", filter(ts, func(t Title) bool { return t.ReleaseDate != nil }), total),
		group("Non-Netflix", filter(ts, func(t Title) bool { return t.ReleaseDate == nil }), total),
	}
	for _, g := range []string{"Yes", "No"} {
		g := g
		st.ByGlobal = append(st.ByGlobal, group(g, filter(ts, func(t Title) bool { return t.AvailableGlobally == g }), total))
	}
	// ช่วงเดียวกับ sheet2 ข้อ 5
	buckets := []struct {
		name, typ string
		lo, hi    float64
	}{
		{"Film <1.5h", "Film", 0, 1.4999}, {"Film 1.5-2h", "Film", 1.5, 2}, {"Film >2h", "Film", 2.0001, 999},
		{"TV <5h", "TV", 0, 4.9999}, {"TV 5-10h", "TV", 5, 10}, {"TV >10h", "TV", 10.0001, 9999},
	}
	for _, b := range buckets {
		b := b
		st.Runtime = append(st.Runtime, group(b.name, filter(ts, func(t Title) bool {
			return t.Type == b.typ && t.RuntimeHours != nil && *t.RuntimeHours >= b.lo && *t.RuntimeHours <= b.hi
		}), total))
	}
	// ปีที่ออกฉาย (เฉพาะ Netflix Title) — เหมือน sheet2 ข้อ 6
	yearOf := func(t Title) int {
		var y int
		fmt.Sscanf(*t.ReleaseDate, "%d", &y)
		return y
	}
	st.ByYear = append(st.ByYear, group("≤2015", filter(ts, func(t Title) bool { return t.ReleaseDate != nil && yearOf(t) <= 2015 }), total))
	for y := 2016; y <= 2026; y++ {
		y := y
		st.ByYear = append(st.ByYear, group(fmt.Sprint(y), filter(ts, func(t Title) bool { return t.ReleaseDate != nil && yearOf(t) == y }), total))
	}
	// Pareto — นับแบบเดียวกับสูตร SUMIF(">="&LARGE(...)) ใน Excel
	hours := make([]int64, len(ts))
	for i, t := range ts {
		hours[i] = t.HoursViewed
	}
	sort.Slice(hours, func(i, j int) bool { return hours[i] > hours[j] })
	for _, p := range []float64{0.01, 0.05, 0.10, 0.20, 0.50} {
		k := int(math.Ceil(float64(len(hours)) * p))
		ps := ParetoStat{Label: fmt.Sprintf("Top %d%%", int(p*100)), Titles: k}
		if k > 0 && total > 0 {
			cut := hours[k-1]
			var s int64
			for _, h := range hours {
				if h >= cut {
					s += h
				}
			}
			ps.ShareHours = float64(s) / float64(total)
		}
		st.Pareto = append(st.Pareto, ps)
	}
	top := func(typ string) []Title {
		x := filter(ts, func(t Title) bool { return t.Type == typ })
		sort.SliceStable(x, func(i, j int) bool { return x[i].Views > x[j].Views })
		if len(x) > 10 {
			x = x[:10]
		}
		return x
	}
	st.Top10 = TopLists{Film: top("Film"), TV: top("TV")}
	return st
}
