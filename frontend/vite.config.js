import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// ตอนพัฒนา (npm run dev) ส่งคำขอ /api ไปที่ backend Go พอร์ต 8080
export default defineConfig({
  plugins: [react()],
  server: { port: 5173, proxy: { '/api': 'http://localhost:8080' } },
})
