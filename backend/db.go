package main

import (
	"database/sql"
	"fmt"
	"log"
	"strings"
	"time"

	_ "github.com/go-sql-driver/mysql"
)

type Store struct{ db *sql.DB }

func OpenStore(dsn string) (*Store, error) {
	db, err := sql.Open("mysql", dsn)
	if err != nil {
		return nil, err
	}
	// รอ MySQL พร้อม (ตอนเปิดด้วย Docker Compose MySQL อาจเริ่มช้ากว่า backend)
	for i := 0; i < 60; i++ {
		if err = db.Ping(); err == nil {
			break
		}
		log.Printf("รอ MySQL... (%v)", err)
		time.Sleep(2 * time.Second)
	}
	if err != nil {
		return nil, fmt.Errorf("เชื่อม MySQL ไม่ได้: %w", err)
	}
	db.SetMaxOpenConns(10)
	s := &Store{db: db}
	return s, s.migrate()
}

func (s *Store) migrate() error {
	stmts := []string{
		`CREATE TABLE IF NOT EXISTS titles (
			id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
			type VARCHAR(10) NOT NULL,
			title VARCHAR(500) NOT NULL,
			available_globally VARCHAR(3) NOT NULL DEFAULT 'No',
			release_date DATE NULL,
			hours_viewed BIGINT NOT NULL DEFAULT 0,
			runtime VARCHAR(10) NULL,
			runtime_hours DECIMAL(8,2) NULL,
			views BIGINT NOT NULL DEFAULT 0,
			updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
			INDEX idx_type (type),
			INDEX idx_views (views),
			INDEX idx_title (title(100))
		) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci`,
		`CREATE TABLE IF NOT EXISTS settings (
			k VARCHAR(64) NOT NULL PRIMARY KEY,
			v TEXT NOT NULL
		) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4`,
		`CREATE TABLE IF NOT EXISTS sync_log (
			id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
			at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
			direction VARCHAR(20) NOT NULL,
			status VARCHAR(10) NOT NULL,
			message TEXT NOT NULL
		) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4`,
	}
	for _, q := range stmts {
		if _, err := s.db.Exec(q); err != nil {
			return fmt.Errorf("migrate: %w", err)
		}
	}
	return nil
}

const titleCols = "id, type, title, available_globally, DATE_FORMAT(release_date,'%Y-%m-%d'), hours_viewed, runtime, runtime_hours, views, DATE_FORMAT(updated_at,'%Y-%m-%d %H:%i:%s')"

type scanner interface{ Scan(dest ...any) error }

func scanTitle(r scanner) (Title, error) {
	var t Title
	var rel, rt sql.NullString
	var rth sql.NullFloat64
	err := r.Scan(&t.ID, &t.Type, &t.Title, &t.AvailableGlobally, &rel, &t.HoursViewed, &rt, &rth, &t.Views, &t.UpdatedAt)
	if rel.Valid {
		t.ReleaseDate = strPtr(rel.String)
	}
	if rt.Valid {
		t.Runtime = strPtr(rt.String)
	}
	if rth.Valid {
		v := rth.Float64
		t.RuntimeHours = &v
	}
	return t, err
}

type ListFilter struct {
	Q, Type, Global, Origin, Sort string
	Page, Size                    int
}

var sortMap = map[string]string{
	"views_desc": "views DESC, id", "views_asc": "views ASC, id",
	"hours_desc": "hours_viewed DESC, id", "hours_asc": "hours_viewed ASC, id",
	"title_asc": "title ASC, id", "title_desc": "title DESC, id",
	"date_desc": "release_date IS NULL, release_date DESC, id", "date_asc": "release_date IS NULL, release_date ASC, id",
	"id_desc": "id DESC", "id_asc": "id ASC", "updated_desc": "updated_at DESC, id DESC",
}

func (s *Store) List(f ListFilter) ([]Title, int, error) {
	where := []string{"1=1"}
	args := []any{}
	if f.Q != "" {
		where = append(where, "title LIKE ?")
		args = append(args, "%"+f.Q+"%")
	}
	if f.Type != "" {
		where = append(where, "type = ?")
		args = append(args, f.Type)
	}
	if f.Global != "" {
		where = append(where, "available_globally = ?")
		args = append(args, f.Global)
	}
	switch f.Origin {
	case "netflix":
		where = append(where, "release_date IS NOT NULL")
	case "other":
		where = append(where, "release_date IS NULL")
	}
	w := strings.Join(where, " AND ")
	var total int
	if err := s.db.QueryRow("SELECT COUNT(*) FROM titles WHERE "+w, args...).Scan(&total); err != nil {
		return nil, 0, err
	}
	order, ok := sortMap[f.Sort]
	if !ok {
		order = sortMap["views_desc"]
	}
	if f.Size <= 0 || f.Size > 200 {
		f.Size = 50
	}
	if f.Page <= 0 {
		f.Page = 1
	}
	q := fmt.Sprintf("SELECT %s FROM titles WHERE %s ORDER BY %s LIMIT %d OFFSET %d", titleCols, w, order, f.Size, (f.Page-1)*f.Size)
	rows, err := s.db.Query(q, args...)
	if err != nil {
		return nil, 0, err
	}
	defer rows.Close()
	out := []Title{}
	for rows.Next() {
		t, err := scanTitle(rows)
		if err != nil {
			return nil, 0, err
		}
		out = append(out, t)
	}
	return out, total, rows.Err()
}

func (s *Store) All() ([]Title, error) {
	rows, err := s.db.Query("SELECT " + titleCols + " FROM titles ORDER BY id")
	if err != nil {
		return nil, err
	}
	defer rows.Close()
	out := []Title{}
	for rows.Next() {
		t, err := scanTitle(rows)
		if err != nil {
			return nil, err
		}
		out = append(out, t)
	}
	return out, rows.Err()
}

func (s *Store) Get(id int64) (Title, error) {
	return scanTitle(s.db.QueryRow("SELECT "+titleCols+" FROM titles WHERE id = ?", id))
}

func (s *Store) Count() (int, error) {
	var n int
	err := s.db.QueryRow("SELECT COUNT(*) FROM titles").Scan(&n)
	return n, err
}

func titleArgs(t Title) []any {
	return []any{t.Type, t.Title, t.AvailableGlobally, t.ReleaseDate, t.HoursViewed, t.Runtime, t.RuntimeHours, t.Views}
}

func (s *Store) Create(t Title) (int64, error) {
	res, err := s.db.Exec(`INSERT INTO titles (type, title, available_globally, release_date, hours_viewed, runtime, runtime_hours, views)
		VALUES (?,?,?,?,?,?,?,?)`, titleArgs(t)...)
	if err != nil {
		return 0, err
	}
	return res.LastInsertId()
}

func (s *Store) Update(t Title) (bool, error) {
	res, err := s.db.Exec(`UPDATE titles SET type=?, title=?, available_globally=?, release_date=?, hours_viewed=?, runtime=?, runtime_hours=?, views=?
		WHERE id=?`, append(titleArgs(t), t.ID)...)
	if err != nil {
		return false, err
	}
	// MySQL คืน 0 ถ้าค่าเหมือนเดิม จึงเช็กว่ามีแถวอยู่จริงแยกอีกครั้ง
	if n, _ := res.RowsAffected(); n > 0 {
		return true, nil
	}
	_, err = s.Get(t.ID)
	return err == nil, nil
}

func (s *Store) Delete(id int64) (bool, error) {
	res, err := s.db.Exec("DELETE FROM titles WHERE id = ?", id)
	if err != nil {
		return false, err
	}
	n, _ := res.RowsAffected()
	return n > 0, nil
}

// ApplyExcel ทำให้ตาราง titles ตรงกับข้อมูลใน Excel (ใน transaction เดียว)
//   - แถวที่มี ID ตรงกับใน DB  → อัปเดต (เฉพาะถ้าค่าต่าง)
//   - แถวที่ไม่มี ID / ID ใหม่  → เพิ่ม
//   - ID ที่อยู่ใน DB แต่ไม่อยู่ใน Excel → ลบ
func (s *Store) ApplyExcel(rows []Title) (ins, upd, del int, newIDs bool, err error) {
	existing, err := s.All()
	if err != nil {
		return
	}
	byID := make(map[int64]Title, len(existing))
	for _, t := range existing {
		byID[t.ID] = t
	}
	tx, err := s.db.Begin()
	if err != nil {
		return
	}
	defer func() {
		if err != nil {
			tx.Rollback()
		}
	}()
	upStmt, err := tx.Prepare(`UPDATE titles SET type=?, title=?, available_globally=?, release_date=?, hours_viewed=?, runtime=?, runtime_hours=?, views=? WHERE id=?`)
	if err != nil {
		return
	}
	defer upStmt.Close()
	insID, err := tx.Prepare(`INSERT INTO titles (id, type, title, available_globally, release_date, hours_viewed, runtime, runtime_hours, views) VALUES (?,?,?,?,?,?,?,?,?)`)
	if err != nil {
		return
	}
	defer insID.Close()
	insAuto, err := tx.Prepare(`INSERT INTO titles (type, title, available_globally, release_date, hours_viewed, runtime, runtime_hours, views) VALUES (?,?,?,?,?,?,?,?)`)
	if err != nil {
		return
	}
	defer insAuto.Close()

	seen := map[int64]bool{}
	for _, t := range rows {
		if t.ID > 0 {
			seen[t.ID] = true
			if old, ok := byID[t.ID]; ok {
				if !sameTitle(old, t) {
					if _, err = upStmt.Exec(append(titleArgs(t), t.ID)...); err != nil {
						return
					}
					upd++
				}
				continue
			}
			if _, err = insID.Exec(append([]any{t.ID}, titleArgs(t)...)...); err != nil {
				return
			}
			ins++
			continue
		}
		if _, err = insAuto.Exec(titleArgs(t)...); err != nil {
			return
		}
		ins++
		newIDs = true
	}
	for id := range byID {
		if !seen[id] {
			if _, err = tx.Exec("DELETE FROM titles WHERE id=?", id); err != nil {
				return
			}
			del++
		}
	}
	err = tx.Commit()
	return
}

func sameTitle(a, b Title) bool {
	eqS := func(x, y *string) bool { return (x == nil && y == nil) || (x != nil && y != nil && *x == *y) }
	return a.Type == b.Type && a.Title == b.Title && a.AvailableGlobally == b.AvailableGlobally &&
		eqS(a.ReleaseDate, b.ReleaseDate) && a.HoursViewed == b.HoursViewed && eqS(a.Runtime, b.Runtime) && a.Views == b.Views
}

// ---------- settings ----------

func (s *Store) GetSetting(k, def string) string {
	var v string
	if err := s.db.QueryRow("SELECT v FROM settings WHERE k=?", k).Scan(&v); err != nil {
		return def
	}
	return v
}

func (s *Store) SetSetting(k, v string) error {
	_, err := s.db.Exec("INSERT INTO settings (k, v) VALUES (?, ?) ON DUPLICATE KEY UPDATE v = VALUES(v)", k, v)
	return err
}

// ---------- sync log ----------

type SyncLog struct {
	ID        int64  `json:"id"`
	At        string `json:"at"`
	Direction string `json:"direction"`
	Status    string `json:"status"`
	Message   string `json:"message"`
}

func (s *Store) AddLog(direction, status, msg string) {
	if len(msg) > 60000 {
		msg = msg[:60000]
	}
	if _, err := s.db.Exec("INSERT INTO sync_log (direction, status, message) VALUES (?,?,?)", direction, status, msg); err != nil {
		log.Printf("write sync_log: %v", err)
	}
	log.Printf("[sync %s] %s: %s", direction, status, firstLine(msg))
}

func (s *Store) Logs(n int) ([]SyncLog, error) {
	rows, err := s.db.Query("SELECT id, DATE_FORMAT(at,'%Y-%m-%d %H:%i:%s'), direction, status, message FROM sync_log ORDER BY id DESC LIMIT ?", n)
	if err != nil {
		return nil, err
	}
	defer rows.Close()
	out := []SyncLog{}
	for rows.Next() {
		var l SyncLog
		if err := rows.Scan(&l.ID, &l.At, &l.Direction, &l.Status, &l.Message); err != nil {
			return nil, err
		}
		out = append(out, l)
	}
	return out, rows.Err()
}

func firstLine(s string) string {
	if i := strings.IndexByte(s, '\n'); i >= 0 {
		return s[:i]
	}
	return s
}
