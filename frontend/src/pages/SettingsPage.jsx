import { useEffect, useState } from 'react'
import { api } from '../api.js'
import { Card } from './Dashboard.jsx'

// ตั้งค่าลิงก์ Tableau Public ที่จะฝังในหน้า "กราฟ Tableau" (ใส่ได้หลาย dashboard)
export default function SettingsPage() {
  const [views, setViews] = useState(null)
  const [msg, setMsg] = useState(null)

  useEffect(() => { api.getSettings().then((s) => setViews(s.tableauViews)) }, [])

  if (!views) return <div className="loading">กำลังโหลด…</div>
  const update = (i, k, v) => setViews(views.map((x, j) => (j === i ? { ...x, [k]: v } : x)))

  async function save() {
    try {
      const s = await api.saveSettings({ tableauViews: views })
      setViews(s.tableauViews)
      setMsg({ kind: 'ok', text: 'บันทึกแล้ว — ดูกราฟได้ที่หน้า "กราฟ Tableau"' })
    } catch (e) {
      setMsg({ kind: 'bad', text: e.message })
    }
  }

  return (
    <Card title="ลิงก์ Tableau Dashboard" note="วางลิงก์จาก Tableau Public (ปุ่ม Share → Copy Link) ใส่ได้หลายอัน แต่ละอันจะเป็นแท็บในหน้ากราฟ">
      {msg && <div className={'alert ' + msg.kind}>{msg.text}</div>}
      {views.map((v, i) => (
        <div className="setting-row" key={i}>
          <input placeholder="ชื่อแท็บ เช่น ภาพรวม" value={v.name} onChange={(e) => update(i, 'name', e.target.value)} />
          <input className="grow" placeholder="https://public.tableau.com/views/..." value={v.url} onChange={(e) => update(i, 'url', e.target.value)} />
          <button className="btn small danger" onClick={() => setViews(views.filter((_, j) => j !== i))}>ลบ</button>
        </div>
      ))}
      <div className="modal-actions left">
        <button className="btn ghost" onClick={() => setViews([...views, { name: '', url: '' }])}>+ เพิ่มลิงก์</button>
        <button className="btn" onClick={save}>บันทึก</button>
      </div>
    </Card>
  )
}
