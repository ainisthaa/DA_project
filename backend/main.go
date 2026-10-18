package main

import (
	"context"
	"log"
	"net/http"
	"os"
	"os/exec"
	"os/signal"
	"path/filepath"
	"strconv"
	"strings"
	"syscall"
	"time"
)

func env(k, def string) string {
	if v := os.Getenv(k); v != "" {
		return v
	}
	return def
}

func main() {
	dsn := env("DB_DSN", "netflix:netflix@tcp(127.0.0.1:3306)/netflix?parseTime=false&charset=utf8mb4")
	excelPath, _ := filepath.Abs(env("EXCEL_PATH", "../data/netflix_data.xlsx"))
	port := env("PORT", "8080")
	staticDir := env("STATIC_DIR", "")
	poll, _ := strconv.Atoi(env("POLL_SECONDS", "3"))
	if poll < 1 {
		poll = 3
	}
	soffice := env("SOFFICE_PATH", "soffice")
	recalc := env("RECALC", "auto")
	recalcOn := false
	switch strings.ToLower(recalc) {
	case "true", "1", "yes":
		recalcOn = true
	case "auto":
		_, err := exec.LookPath(soffice)
		recalcOn = err == nil
	}

	store, err := OpenStore(dsn)
	if err != nil {
		log.Fatal(err)
	}
	sm := NewSyncManager(store, excelPath, time.Duration(poll)*time.Second, recalcOn, soffice)
	log.Printf("Excel: %s | ตรวจไฟล์ทุก %d วินาที | คำนวณสูตรด้วย LibreOffice: %v", excelPath, poll, recalcOn)
	sm.Startup()

	ctx, stop := signal.NotifyContext(context.Background(), os.Interrupt, syscall.SIGTERM)
	defer stop()
	go sm.Run(ctx)

	mux := http.NewServeMux()
	(&API{store: store, sync: sm, tableau: newTableauFiles(excelPath)}).Routes(mux)
	if staticDir != "" {
		mux.Handle("/", spaHandler(staticDir))
	}
	srv := &http.Server{Addr: ":" + port, Handler: withCORS(withLog(mux))}
	go func() {
		<-ctx.Done()
		sctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
		defer cancel()
		_ = srv.Shutdown(sctx)
	}()
	log.Printf("เปิดเว็บที่ http://localhost:%s", port)
	if err := srv.ListenAndServe(); err != nil && err != http.ErrServerClosed {
		log.Fatal(err)
	}
}

// spaHandler เสิร์ฟไฟล์ React ที่ build แล้ว ถ้าไม่พบไฟล์ให้ส่ง index.html (สำหรับ client-side routing)
func spaHandler(dir string) http.Handler {
	fs := http.FileServer(http.Dir(dir))
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		p := filepath.Join(dir, filepath.Clean("/"+r.URL.Path))
		if fi, err := os.Stat(p); err != nil || fi.IsDir() {
			http.ServeFile(w, r, filepath.Join(dir, "index.html"))
			return
		}
		fs.ServeHTTP(w, r)
	})
}

func withCORS(h http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type")
		if r.Method == http.MethodOptions {
			w.WriteHeader(204)
			return
		}
		h.ServeHTTP(w, r)
	})
}

func withLog(h http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		start := time.Now()
		h.ServeHTTP(w, r)
		if strings.HasPrefix(r.URL.Path, "/api/") && r.Method != http.MethodGet {
			log.Printf("%s %s (%v)", r.Method, r.URL.Path, time.Since(start).Round(time.Millisecond))
		}
	})
}
