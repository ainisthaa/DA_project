package main

import (
	"database/sql"
	"encoding/json"
	"errors"
	"io"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

type API struct {
	store   *Store
	sync    *SyncManager
	tableau tableauFiles
}

func writeJSON(w http.ResponseWriter, code int, v any) {
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.WriteHeader(code)
	_ = json.NewEncoder(w).Encode(v)
}

func writeErr(w http.ResponseWriter, code int, msg string) {
	writeJSON(w, code, map[string]string{"error": msg})
}

func (a *API) Routes(mux *http.ServeMux) {
	mux.HandleFunc("GET /api/health", func(w http.ResponseWriter, r *http.Request) {
		writeJSON(w, 200, map[string]string{"status": "ok"})
	})
	mux.HandleFunc("GET /api/titles", a.listTitles)
	mux.HandleFunc("GET /api/titles/{id}", a.getTitle)
	mux.HandleFunc("POST /api/titles", a.createTitle)
	mux.HandleFunc("PUT /api/titles/{id}", a.updateTitle)
	mux.HandleFunc("DELETE /api/titles/{id}", a.deleteTitle)
	mux.HandleFunc("GET /api/stats", a.stats)
	mux.HandleFunc("GET /api/excel/summary", a.excelSummary)
	mux.HandleFunc("GET /api/excel/download", a.download)
	mux.HandleFunc("POST /api/excel/upload", a.upload)
	mux.HandleFunc("GET /api/sync/status", a.syncStatus)
	mux.HandleFunc("POST /api/sync/import", a.syncImport)
	mux.HandleFunc("POST /api/sync/export", a.syncExport)
	mux.HandleFunc("GET /api/settings", a.getSettings)
	mux.HandleFunc("PUT /api/settings", a.putSettings)
	mux.HandleFunc("GET /api/tableau", a.tableauInfo)
	mux.HandleFunc("GET /api/tableau/workbook", a.tableauWorkbook)
	mux.HandleFunc("GET /api/tableau/images/{file}", a.tableauImage)
}

func pathID(r *http.Request) (int64, bool) {
	id, err := strconv.ParseInt(r.PathValue("id"), 10, 64)
	return id, err == nil && id > 0
}

func (a *API) listTitles(w http.ResponseWriter, r *http.Request) {
	q := r.URL.Query()
	page, _ := strconv.Atoi(q.Get("page"))
	size, _ := strconv.Atoi(q.Get("size"))
	items, total, err := a.store.List(ListFilter{
		Q: strings.TrimSpace(q.Get("q")), Type: q.Get("type"), Global: q.Get("global"),
		Origin: q.Get("origin"), Sort: q.Get("sort"), Page: page, Size: size,
	})
	if err != nil {
		writeErr(w, 500, err.Error())
		return
	}
	writeJSON(w, 200, map[string]any{"items": items, "total": total})
}

func (a *API) getTitle(w http.ResponseWriter, r *http.Request) {
	id, ok := pathID(r)
	if !ok {
		writeErr(w, 400, "id ไม่ถูกต้อง")
		return
	}
	t, err := a.store.Get(id)
	if errors.Is(err, sql.ErrNoRows) {
		writeErr(w, 404, "ไม่พบข้อมูล")
		return
	} else if err != nil {
		writeErr(w, 500, err.Error())
		return
	}
	writeJSON(w, 200, t)
}

func decodeTitle(r *http.Request) (Title, error) {
	var t Title
	dec := json.NewDecoder(io.LimitReader(r.Body, 1<<20))
	if err := dec.Decode(&t); err != nil {
		return t, errors.New("ข้อมูล JSON ไม่ถูกต้อง: " + err.Error())
	}
	return t, t.Normalize()
}

func (a *API) createTitle(w http.ResponseWriter, r *http.Request) {
	t, err := decodeTitle(r)
	if err != nil {
		writeErr(w, 400, err.Error())
		return
	}
	id, err := a.store.Create(t)
	if err != nil {
		writeErr(w, 500, err.Error())
		return
	}
	a.sync.RequestExport()
	saved, _ := a.store.Get(id)
	writeJSON(w, 201, saved)
}

func (a *API) updateTitle(w http.ResponseWriter, r *http.Request) {
	id, ok := pathID(r)
	if !ok {
		writeErr(w, 400, "id ไม่ถูกต้อง")
		return
	}
	t, err := decodeTitle(r)
	if err != nil {
		writeErr(w, 400, err.Error())
		return
	}
	t.ID = id
	found, err := a.store.Update(t)
	if err != nil {
		writeErr(w, 500, err.Error())
		return
	}
	if !found {
		writeErr(w, 404, "ไม่พบข้อมูล")
		return
	}
	a.sync.RequestExport()
	saved, _ := a.store.Get(id)
	writeJSON(w, 200, saved)
}

func (a *API) deleteTitle(w http.ResponseWriter, r *http.Request) {
	id, ok := pathID(r)
	if !ok {
		writeErr(w, 400, "id ไม่ถูกต้อง")
		return
	}
	found, err := a.store.Delete(id)
	if err != nil {
		writeErr(w, 500, err.Error())
		return
	}
	if !found {
		writeErr(w, 404, "ไม่พบข้อมูล")
		return
	}
	a.sync.RequestExport()
	writeJSON(w, 200, map[string]any{"deleted": id})
}

func (a *API) stats(w http.ResponseWriter, r *http.Request) {
	ts, err := a.store.All()
	if err != nil {
		writeErr(w, 500, err.Error())
		return
	}
	writeJSON(w, 200, ComputeStats(ts))
}

func (a *API) excelSummary(w http.ResponseWriter, r *http.Request) {
	rows, stale, err := ReadSummary(a.sync.path)
	if err != nil {
		writeErr(w, 500, err.Error())
		return
	}
	st := a.sync.Status()
	writeJSON(w, 200, map[string]any{"rows": rows, "stale": stale || st.PendingExport, "excelModified": st.ExcelModified, "lastRecalc": st.LastRecalc})
}

func (a *API) download(w http.ResponseWriter, r *http.Request) {
	a.sync.mu.Lock()
	data, err := os.ReadFile(a.sync.path)
	a.sync.mu.Unlock()
	if err != nil {
		writeErr(w, 500, err.Error())
		return
	}
	name := filepath.Base(a.sync.path)
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", "attachment; filename=\""+name+"\"; filename*=UTF-8''"+url.PathEscape(name))
	_, _ = w.Write(data)
}

func (a *API) upload(w http.ResponseWriter, r *http.Request) {
	r.Body = http.MaxBytesReader(w, r.Body, 50<<20)
	file, hdr, err := r.FormFile("file")
	if err != nil {
		writeErr(w, 400, "ไม่พบไฟล์ (field ชื่อ file)")
		return
	}
	defer file.Close()
	if !strings.HasSuffix(strings.ToLower(hdr.Filename), ".xlsx") {
		writeErr(w, 400, "รองรับเฉพาะไฟล์ .xlsx")
		return
	}
	tmp, err := os.CreateTemp("", "upload-*.xlsx")
	if err != nil {
		writeErr(w, 500, err.Error())
		return
	}
	defer os.Remove(tmp.Name())
	if _, err := io.Copy(tmp, file); err != nil {
		tmp.Close()
		writeErr(w, 500, err.Error())
		return
	}
	tmp.Close()
	res, err := a.sync.ImportFile(tmp.Name(), r.URL.Query().Get("force") == "1")
	if err != nil {
		writeErr(w, 400, err.Error())
		return
	}
	writeJSON(w, 200, res)
}

func (a *API) syncStatus(w http.ResponseWriter, r *http.Request) {
	logs, err := a.store.Logs(30)
	if err != nil {
		writeErr(w, 500, err.Error())
		return
	}
	n, _ := a.store.Count()
	writeJSON(w, 200, map[string]any{"status": a.sync.Status(), "logs": logs, "dbRows": n})
}

func (a *API) syncImport(w http.ResponseWriter, r *http.Request) {
	res, err := a.sync.Import(r.URL.Query().Get("force") == "1", "manual")
	if err != nil {
		writeErr(w, 400, err.Error())
		return
	}
	writeJSON(w, 200, res)
}

func (a *API) syncExport(w http.ResponseWriter, r *http.Request) {
	if err := a.sync.Export("manual"); err != nil {
		writeErr(w, 500, err.Error())
		return
	}
	writeJSON(w, 200, map[string]string{"status": "ok"})
}

// ---------- settings: รายการ Tableau dashboard ที่ฝังบนเว็บ ----------

type TableauView struct {
	Name string `json:"name"`
	URL  string `json:"url"`
}

type Settings struct {
	TableauViews []TableauView `json:"tableauViews"`
}

// Dashboard ที่ publish ไว้บน Tableau Public — ใช้เป็นค่าเริ่มต้นจนกว่าจะบันทึกในหน้าตั้งค่า
const defaultTableauViews = `[
 {"name":"ภาพรวม (Tableau)","url":"https://public.tableau.com/views/tableau_netflix2/Dashboard_Overview"},
 {"name":"Top 10 และ Runtime (Tableau)","url":"https://public.tableau.com/views/tableau_netflix2/Dashboard_Details"}
]`

func (a *API) loadSettings() Settings {
	var s Settings
	_ = json.Unmarshal([]byte(a.store.GetSetting("tableau_views", defaultTableauViews)), &s.TableauViews)
	if s.TableauViews == nil {
		s.TableauViews = []TableauView{}
	}
	return s
}

func (a *API) getSettings(w http.ResponseWriter, r *http.Request) {
	writeJSON(w, 200, a.loadSettings())
}

func (a *API) putSettings(w http.ResponseWriter, r *http.Request) {
	var s Settings
	if err := json.NewDecoder(io.LimitReader(r.Body, 1<<20)).Decode(&s); err != nil {
		writeErr(w, 400, "ข้อมูลไม่ถูกต้อง")
		return
	}
	clean := []TableauView{}
	for _, v := range s.TableauViews {
		v.Name, v.URL = strings.TrimSpace(v.Name), strings.TrimSpace(v.URL)
		if v.URL == "" {
			continue
		}
		u, err := url.Parse(v.URL)
		if err != nil || (u.Scheme != "https" && u.Scheme != "http") {
			writeErr(w, 400, "URL ไม่ถูกต้อง: "+v.URL)
			return
		}
		if v.Name == "" {
			v.Name = "Dashboard"
		}
		clean = append(clean, v)
	}
	b, _ := json.Marshal(clean)
	if err := a.store.SetSetting("tableau_views", string(b)); err != nil {
		writeErr(w, 500, err.Error())
		return
	}
	writeJSON(w, 200, Settings{TableauViews: clean})
}
