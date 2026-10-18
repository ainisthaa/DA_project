package main

import (
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"
)

// ไฟล์ Tableau อยู่โฟลเดอร์เดียวกับไฟล์ Excel (data/):
//   data/tableau_netflix2.twbx ไฟล์ workbook ที่ทำใน Tableau (ให้ดาวน์โหลดจากหน้าเว็บ)
//   data/tableau/*.png|.jpg   ภาพ Dashboard ที่ Export จาก Tableau (แสดงบนเว็บได้แม้ยังไม่ publish / ไม่มีอินเทอร์เน็ต)

type tableauFiles struct {
	workbook  string
	imagesDir string
}

func newTableauFiles(excelPath string) tableauFiles {
	dir := filepath.Dir(excelPath)
	return tableauFiles{
		workbook:  filepath.Join(dir, env("TABLEAU_WORKBOOK", "tableau_netflix2.twbx")),
		imagesDir: filepath.Join(dir, "tableau"),
	}
}

var imageTypes = map[string]string{".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}

type tableauImage struct {
	Name string `json:"name"` // ชื่อแท็บ (จากชื่อไฟล์ เช่น 01_Dashboard_Overview.png → Dashboard Overview)
	URL  string `json:"url"`
}

// ชื่อไฟล์ขึ้นต้นด้วยเลขลำดับได้ (01_, 02_) เพื่อกำหนดลำดับแท็บ — ตัดเลขและ _ ออกตอนแสดง
func imageLabel(file string) string {
	s := strings.TrimSuffix(file, filepath.Ext(file))
	s = strings.TrimLeft(s, "0123456789")
	s = strings.TrimLeft(s, "_- ")
	if s == "" {
		s = file
	}
	return strings.ReplaceAll(s, "_", " ")
}

func (a *API) tableauInfo(w http.ResponseWriter, r *http.Request) {
	out := map[string]any{"workbook": nil, "images": []tableauImage{}}
	if st, err := os.Stat(a.tableau.workbook); err == nil && !st.IsDir() {
		out["workbook"] = map[string]any{
			"name": filepath.Base(a.tableau.workbook), "size": st.Size(),
			"modified": st.ModTime().Format(time.RFC3339),
		}
	}
	entries, _ := os.ReadDir(a.tableau.imagesDir)
	var imgs []tableauImage
	for _, e := range entries {
		n := e.Name()
		if e.IsDir() || strings.HasPrefix(n, ".") || imageTypes[strings.ToLower(filepath.Ext(n))] == "" {
			continue
		}
		imgs = append(imgs, tableauImage{Name: imageLabel(n), URL: "/api/tableau/images/" + url.PathEscape(n)})
	}
	sort.Slice(imgs, func(i, j int) bool { return imgs[i].URL < imgs[j].URL })
	if imgs != nil {
		out["images"] = imgs
	}
	writeJSON(w, 200, out)
}

func (a *API) tableauImage(w http.ResponseWriter, r *http.Request) {
	name := r.PathValue("file")
	ctype := imageTypes[strings.ToLower(filepath.Ext(name))]
	// รับเฉพาะชื่อไฟล์ในโฟลเดอร์ data/tableau เท่านั้น กันการอ่านไฟล์นอกโฟลเดอร์ (../)
	if ctype == "" || name != filepath.Base(name) || strings.HasPrefix(name, ".") {
		writeErr(w, 404, "ไม่พบภาพ")
		return
	}
	data, err := os.ReadFile(filepath.Join(a.tableau.imagesDir, name))
	if err != nil {
		writeErr(w, 404, "ไม่พบภาพ")
		return
	}
	w.Header().Set("Content-Type", ctype)
	w.Header().Set("Cache-Control", "no-cache")
	_, _ = w.Write(data)
}

func (a *API) tableauWorkbook(w http.ResponseWriter, r *http.Request) {
	data, err := os.ReadFile(a.tableau.workbook)
	if err != nil {
		writeErr(w, 404, "ไม่พบไฟล์ "+filepath.Base(a.tableau.workbook)+" ในโฟลเดอร์ data")
		return
	}
	name := filepath.Base(a.tableau.workbook)
	w.Header().Set("Content-Type", "application/octet-stream")
	w.Header().Set("Content-Disposition", "attachment; filename=\""+name+"\"; filename*=UTF-8''"+url.PathEscape(name))
	_, _ = w.Write(data)
}
