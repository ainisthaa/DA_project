import { useEffect, useState } from 'react'
import Dashboard from './pages/Dashboard.jsx'
import Charts from './pages/Charts.jsx'
import DataTable from './pages/DataTable.jsx'
import SyncPage from './pages/SyncPage.jsx'
import SettingsPage from './pages/SettingsPage.jsx'
import { api } from './api.js'

const PAGES = [
  { key: 'dashboard', label: 'ภาพรวม', comp: Dashboard },
  { key: 'charts', label: 'กราฟ Tableau', comp: Charts },
  { key: 'data', label: 'จัดการข้อมูล', comp: DataTable },
  { key: 'sync', label: 'ซิงก์ Excel', comp: SyncPage },
  { key: 'settings', label: 'ตั้งค่า', comp: SettingsPage },
]

// ใช้ hash (#data, #sync, ...) เป็นตัวเลือกหน้า จึงกด back/refresh ได้โดยไม่ต้องมี router
function currentPage() {
  const k = window.location.hash.replace('#', '')
  return PAGES.find((p) => p.key === k) ? k : 'dashboard'
}

export default function App() {
  const [page, setPage] = useState(currentPage())
  const [sync, setSync] = useState(null)

  useEffect(() => {
    const onHash = () => setPage(currentPage())
    window.addEventListener('hashchange', onHash)
    return () => window.removeEventListener('hashchange', onHash)
  }, [])

  // สถานะซิงก์แสดงที่มุมขวาบนทุกหน้า (อัปเดตทุก 3 วินาที)
  useEffect(() => {
    let alive = true
    const load = () => api.syncStatus().then((d) => alive && setSync(d)).catch(() => alive && setSync(null))
    load()
    const t = setInterval(load, 3000)
    return () => { alive = false; clearInterval(t) }
  }, [])

  const Page = PAGES.find((p) => p.key === page).comp
  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="logo">N</span>
          <div>
            <div className="brand-title">Netflix Data Explorer</div>
            <div className="brand-sub">What We Watched · ม.ค.–มิ.ย. 2026</div>
          </div>
        </div>
        <nav className="tabs">
          {PAGES.map((p) => (
            <a key={p.key} href={'#' + p.key} className={p.key === page ? 'active' : ''}>{p.label}</a>
          ))}
        </nav>
        <SyncBadge sync={sync} />
      </header>
      <main className="content">
        <Page sync={sync} />
      </main>
      <footer className="foot">
        ที่มาข้อมูล: Netflix, <i>What We Watched: A Netflix Engagement Report (January–June 2026)</i>
      </footer>
    </div>
  )
}

function SyncBadge({ sync }) {
  if (!sync) return <a href="#sync" className="badge bad">● ติดต่อ backend ไม่ได้</a>
  const s = sync.status
  if (s.lastError) return <a href="#sync" className="badge bad" title={s.lastError}>● ซิงก์มีปัญหา</a>
  if (s.pendingExport) return <a href="#sync" className="badge warn">● กำลังเขียน Excel…</a>
  return <a href="#sync" className="badge ok" title={'DB ' + sync.dbRows + ' แถว'}>● Excel ↔ DB ตรงกัน</a>
}
