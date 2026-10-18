package main

import (
	"context"
	"fmt"
	"log"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"sync"
	"time"
)

// SyncManager ดูแลการซิงก์สองทางระหว่าง MySQL กับไฟล์ Excel
//
//	เว็บ → DB → Excel : ทุกครั้งที่เพิ่ม/แก้/ลบบนเว็บ จะเขียน sheet1 ใหม่ (หน่วง 0.5 วิ เพื่อรวมหลายคำสั่ง)
//	Excel → DB        : ตรวจไฟล์ทุก POLL_SECONDS วินาที ถ้าไฟล์ถูกแก้/บันทึกจากภายนอก จะดึงเข้า DB อัตโนมัติ
type SyncManager struct {
	store       *Store
	path        string
	interval    time.Duration
	recalc      bool
	soffice     string
	mu          sync.Mutex // กันไม่ให้ import/export ชนกัน
	stMu        sync.Mutex // ป้องกัน status
	lastMod     time.Time
	lastSize    int64
	pending     bool
	pendingAt   time.Time
	status      SyncStatus
}

type SyncStatus struct {
	ExcelPath      string     `json:"excelPath"`
	ExcelModified  string     `json:"excelModified"`
	LastImport     string     `json:"lastImport"`
	LastExport     string     `json:"lastExport"`
	LastRecalc     string     `json:"lastRecalc"`
	PendingExport  bool       `json:"pendingExport"`
	LastError      string     `json:"lastError"`
	RowErrors      []RowError `json:"rowErrors"`
	RecalcEnabled  bool       `json:"recalcEnabled"`
	PollSeconds    int        `json:"pollSeconds"`
}

type ImportResult struct {
	Inserted  int        `json:"inserted"`
	Updated   int        `json:"updated"`
	Deleted   int        `json:"deleted"`
	Rows      int        `json:"rows"`
	RowErrors []RowError `json:"rowErrors"`
}

func NewSyncManager(store *Store, path string, interval time.Duration, recalc bool, soffice string) *SyncManager {
	m := &SyncManager{store: store, path: path, interval: interval, recalc: recalc, soffice: soffice}
	m.status = SyncStatus{ExcelPath: path, RecalcEnabled: recalc, PollSeconds: int(interval.Seconds()), RowErrors: []RowError{}}
	return m
}

func (m *SyncManager) Status() SyncStatus {
	m.stMu.Lock()
	defer m.stMu.Unlock()
	s := m.status
	s.PendingExport = m.pending
	if fi, err := os.Stat(m.path); err == nil {
		s.ExcelModified = fi.ModTime().Format("2006-01-02 15:04:05")
	}
	return s
}

func (m *SyncManager) setStatus(fn func(s *SyncStatus)) {
	m.stMu.Lock()
	fn(&m.status)
	m.stMu.Unlock()
}

func now() string { return time.Now().Format("2006-01-02 15:04:05") }

// Startup ตัดสินใจว่าจะซิงก์ทิศไหนตอนเปิดโปรแกรม โดยดูสถานะที่บันทึกไว้ในตาราง settings
//   - DB ว่าง                             → ดึงข้อมูลทั้งหมดจาก Excel
//   - ไฟล์ Excel ถูกแก้ตอนโปรแกรมปิดอยู่   → ดึงจาก Excel (ถ้าเว็บมีงานค้างด้วย จะบันทึกคำเตือนไว้)
//   - เว็บมีการแก้ที่ยังไม่ได้เขียนลง Excel  → เขียนลง Excel
func (m *SyncManager) Startup() {
	n, err := m.store.Count()
	if err != nil {
		log.Printf("count titles: %v", err)
	}
	fi, err := os.Stat(m.path)
	if err != nil {
		m.fail("startup", "startup", fmt.Errorf("ไม่พบไฟล์ Excel: %s", m.path))
		return
	}
	m.loadLastTimes()
	dirty := m.store.GetSetting("db_dirty", "0") == "1"
	fileChanged := m.store.GetSetting("excel_sig", "") != fileSig(fi)
	switch {
	case n == 0:
		_, err = m.Import(true, "startup")
	case fileChanged:
		if dirty {
			m.store.AddLog("startup", "warning", "ไฟล์ Excel ถูกแก้ตอนโปรแกรมปิดอยู่ และเว็บมีการแก้ที่ยังไม่ได้เขียนลง Excel — ใช้ข้อมูลจาก Excel")
		}
		_, err = m.Import(false, "startup")
	case dirty:
		m.rememberFile() // ไฟล์ไม่ได้ถูกแก้ จำสถานะไว้ก่อน ไม่ให้ Export เข้าใจผิดว่าต้องดึงจาก Excel
		err = m.Export("startup")
	default:
		m.rememberFile()
		m.store.AddLog("startup", "ok", fmt.Sprintf("ข้อมูลตรงกันแล้ว (%d แถว)", n))
	}
	if err != nil {
		log.Printf("startup sync: %v", err)
	}
}

func fileSig(fi os.FileInfo) string {
	return fmt.Sprintf("%d:%d", fi.ModTime().UnixNano(), fi.Size())
}

// Run วนตรวจไฟล์และส่งออกที่ค้างอยู่ จนกว่า ctx จะถูกยกเลิก
func (m *SyncManager) Run(ctx context.Context) {
	tick := time.NewTicker(500 * time.Millisecond)
	defer tick.Stop()
	lastPoll := time.Now()
	for {
		select {
		case <-ctx.Done():
			return
		case <-tick.C:
		}
		m.stMu.Lock()
		due := m.pending && time.Since(m.pendingAt) >= 500*time.Millisecond
		m.stMu.Unlock()
		if due {
			if err := m.Export("web"); err != nil {
				// ส่งออกไม่สำเร็จ (เช่นไฟล์ถูกล็อก) → ลองใหม่รอบหน้า
				m.stMu.Lock()
				m.pending = true
				m.pendingAt = time.Now().Add(10 * time.Second)
				m.stMu.Unlock()
			}
			continue
		}
		if time.Since(lastPoll) < m.interval {
			continue
		}
		lastPoll = time.Now()
		if m.changedOnDisk() {
			time.Sleep(time.Second) // รอให้โปรแกรมที่บันทึกไฟล์เขียนเสร็จ
			if _, err := m.Import(false, "excel-watch"); err != nil {
				log.Printf("watch import: %v", err)
			}
		}
	}
}

// RequestExport ถูกเรียกหลังเพิ่ม/แก้/ลบบนเว็บ
func (m *SyncManager) RequestExport() {
	_ = m.store.SetSetting("db_dirty", "1") // จำไว้ เผื่อโปรแกรมปิดก่อนเขียน Excel เสร็จ
	m.stMu.Lock()
	m.pending = true
	m.pendingAt = time.Now()
	m.stMu.Unlock()
}

func (m *SyncManager) changedOnDisk() bool {
	fi, err := os.Stat(m.path)
	if err != nil {
		return false
	}
	m.stMu.Lock()
	defer m.stMu.Unlock()
	return !fi.ModTime().Equal(m.lastMod) || fi.Size() != m.lastSize
}

// rememberFile จำลายเซ็นไฟล์ (เวลาแก้ไข + ขนาด) ที่ซิงก์ล่าสุด ทั้งในหน่วยความจำและในฐานข้อมูล
func (m *SyncManager) rememberFile() {
	if fi, err := os.Stat(m.path); err == nil {
		m.stMu.Lock()
		m.lastMod, m.lastSize = fi.ModTime(), fi.Size()
		m.stMu.Unlock()
		_ = m.store.SetSetting("excel_sig", fileSig(fi))
	}
}

// Import ดึง sheet1 ของ Excel เข้า MySQL
func (m *SyncManager) Import(force bool, reason string) (ImportResult, error) {
	m.mu.Lock()
	defer m.mu.Unlock()
	return m.importLocked(force, reason, m.path)
}

// ImportFile นำเข้าจากไฟล์ที่อัปโหลดผ่านเว็บ แล้วเขียนลงไฟล์ Excel หลัก (sheet2–4 ของไฟล์หลักยังอยู่ครบ)
func (m *SyncManager) ImportFile(src string, force bool) (ImportResult, error) {
	m.mu.Lock()
	defer m.mu.Unlock()
	return m.importLocked(force, "upload", src)
}

func (m *SyncManager) importLocked(force bool, reason, src string) (ImportResult, error) {
	var res ImportResult
	if src == m.path {
		m.rememberFile()
	}
	rows, rowErrs, err := ReadExcel(src)
	if err != nil {
		m.fail("import", reason, err)
		return res, err
	}
	res.Rows, res.RowErrors = len(rows), rowErrs
	if res.RowErrors == nil {
		res.RowErrors = []RowError{}
	}
	dbCount, _ := m.store.Count()
	if !force && dbCount >= 100 && len(rows) < dbCount/2 {
		err := fmt.Errorf("ยกเลิกการนำเข้า: Excel มีข้อมูลที่ถูกต้อง %d แถว แต่ในฐานข้อมูลมี %d แถว (หายไปเกินครึ่ง) "+
			"— ถ้าตั้งใจลบจริง ให้กด \"นำเข้าจาก Excel (บังคับ)\" ในหน้าซิงก์ข้อมูล", len(rows), dbCount)
		m.fail("import", reason, err)
		return res, err
	}
	ins, upd, del, newIDs, err := m.store.ApplyExcel(rows)
	if err != nil {
		m.fail("import", reason, err)
		return res, err
	}
	res.Inserted, res.Updated, res.Deleted = ins, upd, del
	msg := fmt.Sprintf("Excel → DB (%s): อ่าน %d แถว, เพิ่ม %d, แก้ %d, ลบ %d, แถวที่ผิด %d", reason, len(rows), ins, upd, del, len(rowErrs))
	for i, e := range rowErrs {
		if i == 20 {
			msg += fmt.Sprintf("\n... และอีก %d แถว", len(rowErrs)-20)
			break
		}
		msg += fmt.Sprintf("\nแถว %d: %s", e.Row, e.Message)
	}
	status := "ok"
	if len(rowErrs) > 0 {
		status = "warning"
	}
	if ins+upd+del > 0 || len(rowErrs) > 0 || reason != "excel-watch" {
		m.store.AddLog("import", status, msg)
	}
	m.setStatus(func(s *SyncStatus) {
		s.LastImport, s.LastError, s.RowErrors = now(), "", res.RowErrors
	})
	// มีแถวใหม่ใน Excel ที่ยังไม่มี ID → เขียนกลับเพื่อใส่ ID ที่ฐานข้อมูลสร้างให้
	// (ถ้าแก้ใน Excel เอง Excel คำนวณสูตรให้แล้ว จึงไม่ต้องเขียนทับไฟล์ของผู้ใช้)
	if newIDs || reason == "upload" {
		if err := m.exportLocked("after-import"); err != nil {
			log.Printf("export after import: %v", err)
		}
	}
	return res, nil
}

// Export เขียนข้อมูลจาก MySQL ลง Excel sheet1
func (m *SyncManager) Export(reason string) error {
	m.mu.Lock()
	defer m.mu.Unlock()
	// ถ้าไฟล์ถูกแก้จากภายนอกแต่ยังไม่ได้ดึงเข้า ให้ดึงก่อนเพื่อไม่ทับงานของผู้ใช้
	if m.changedOnDisk() {
		if _, err := m.importLocked(false, "before-export", m.path); err != nil {
			return err
		}
	}
	return m.exportLocked(reason)
}

func (m *SyncManager) exportLocked(reason string) error {
	m.stMu.Lock()
	m.pending = false
	m.stMu.Unlock()
	titles, err := m.store.All()
	if err != nil {
		m.markPending()
		m.fail("export", reason, err)
		return err
	}
	if err := WriteExcel(m.path, titles); err != nil {
		m.markPending()
		m.fail("export", reason, err)
		return err
	}
	m.setStatus(func(s *SyncStatus) { s.LastExport, s.LastError = now(), "" })
	msg := fmt.Sprintf("DB → Excel (%s): เขียน %d แถวลง %s", reason, len(titles), SheetRaw)
	if m.recalc {
		if err := m.recalcExcel(); err != nil {
			msg += "\nคำนวณสูตรด้วย LibreOffice ไม่สำเร็จ: " + err.Error() + " (เปิดไฟล์ใน Excel แล้วสูตรจะคำนวณเอง)"
		} else {
			msg += " และคำนวณสูตร sheet2–4 ใหม่แล้ว"
			m.setStatus(func(s *SyncStatus) { s.LastRecalc = now() })
		}
	}
	m.rememberFile()
	_ = m.store.SetSetting("db_dirty", "0")
	m.store.AddLog("export", "ok", msg)
	return nil
}

func (m *SyncManager) markPending() {
	m.stMu.Lock()
	m.pending = true
	m.stMu.Unlock()
}

// recalcExcel ให้ LibreOffice (headless) เปิดไฟล์ คำนวณสูตรทั้งหมด แล้วบันทึกกลับ
// เพื่อให้ค่าใน sheet2–4 เป็นปัจจุบันทันที (Tableau และหน้าเว็บอ่านค่าที่คำนวณแล้วนี้)
func (m *SyncManager) recalcExcel() error {
	outDir, err := os.MkdirTemp("", "recalc")
	if err != nil {
		return err
	}
	defer os.RemoveAll(outDir)
	ctx, cancel := context.WithTimeout(context.Background(), 120*time.Second)
	defer cancel()
	// LibreOffice ต้องการ path แบบ URL: file:///tmp/... (Linux/Mac) หรือ file:///C:/... (Windows)
	profile := "file://" + filepath.ToSlash(filepath.Join(os.TempDir(), "lo_profile_netflix"))
	if !strings.HasPrefix(profile, "file:///") {
		profile = strings.Replace(profile, "file://", "file:///", 1)
	}
	cmd := exec.CommandContext(ctx, m.soffice, "-env:UserInstallation="+profile, "--headless", "--norestore",
		"--calc", "--convert-to", "xlsx:Calc Office Open XML", "--outdir", outDir, m.path)
	out, err := cmd.CombinedOutput()
	if err != nil {
		return fmt.Errorf("%v: %s", err, strings.TrimSpace(string(out)))
	}
	produced := filepath.Join(outDir, filepath.Base(m.path))
	if _, err := os.Stat(produced); err != nil {
		return fmt.Errorf("LibreOffice ไม่ได้สร้างไฟล์: %s", strings.TrimSpace(string(out)))
	}
	return copyReplace(produced, m.path)
}

func copyReplace(src, dst string) error {
	data, err := os.ReadFile(src)
	if err != nil {
		return err
	}
	tmp := filepath.Join(filepath.Dir(dst), "."+filepath.Base(dst)+".recalc")
	if err := os.WriteFile(tmp, data, 0o644); err != nil {
		return err
	}
	return os.Rename(tmp, dst)
}

func (m *SyncManager) fail(direction, reason string, err error) {
	m.store.AddLog(direction, "error", fmt.Sprintf("(%s) %v", reason, err))
	m.setStatus(func(s *SyncStatus) { s.LastError = err.Error() })
}

// loadLastTimes ดึงเวลาซิงก์ล่าสุดจากประวัติ เพื่อให้หน้าเว็บแสดงได้แม้เพิ่งเปิดโปรแกรม
func (m *SyncManager) loadLastTimes() {
	logs, err := m.store.Logs(200)
	if err != nil {
		return
	}
	m.setStatus(func(s *SyncStatus) {
		for _, l := range logs {
			if l.Status == "error" {
				continue
			}
			if l.Direction == "import" && s.LastImport == "" {
				s.LastImport = l.At
			}
			if l.Direction == "export" && s.LastExport == "" {
				s.LastExport = l.At
				if m.recalc && strings.Contains(l.Message, "คำนวณสูตร sheet2–4 ใหม่แล้ว") {
					s.LastRecalc = l.At
				}
			}
		}
	})
}
