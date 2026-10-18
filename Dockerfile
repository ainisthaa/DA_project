# ---------- 1) build หน้าเว็บ React ----------
FROM node:22-bookworm-slim AS web
WORKDIR /web
COPY frontend/package.json frontend/package-lock.json ./
RUN npm ci
COPY frontend/ ./
RUN npm run build

# ---------- 2) build backend Go ----------
FROM golang:1.24-bookworm AS api
WORKDIR /src
COPY backend/ ./
RUN go mod tidy && CGO_ENABLED=0 go build -o /out/netflix-backend .

# ---------- 3) image สำหรับรันจริง (มี LibreOffice ไว้คำนวณสูตร Excel อัตโนมัติ) ----------
FROM debian:bookworm-slim
RUN apt-get update \
 && apt-get install -y --no-install-recommends libreoffice-calc-nogui fonts-thai-tlwg ca-certificates tzdata \
 && rm -rf /var/lib/apt/lists/*
ENV TZ=Asia/Bangkok \
    PORT=8080 \
    STATIC_DIR=/app/web \
    EXCEL_PATH=/data/netflix_data.xlsx \
    RECALC=auto \
    POLL_SECONDS=3
WORKDIR /app
COPY --from=api /out/netflix-backend /app/netflix-backend
COPY --from=web /web/dist /app/web
EXPOSE 8080
CMD ["/app/netflix-backend"]
