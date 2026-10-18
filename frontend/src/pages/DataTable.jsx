import { useEffect, useState } from 'react'
import { api, fmt, fmtDec } from '../api.js'
import { Card } from './Dashboard.jsx'

const EMPTY = { type: 'Film', title: '', availableGlobally: 'Yes', releaseDate: '', hoursViewed: '', runtime: '', views: '' }

// หน้าจัดการข้อมูล: ค้นหา / กรอง / เพิ่ม / แก้ / ลบ — ทุกการเปลี่ยนแปลงจะถูกเขียนลง Excel อัตโนมัติ
export default function DataTable() {
  const [q, setQ] = useState('')
  const [filters, setFilters] = useState({ type: '', global: '', origin: '', sort: 'views_desc' })
  const [page, setPage] = useState(1)
  const [data, setData] = useState({ items: [], total: 0 })
  const [loading, setLoading] = useState(false)
  const [editing, setEditing] = useState(null) // null | object
  const [msg, setMsg] = useState(null)
  const [reload, setReload] = useState(0)
  const size = 25

  useEffect(() => {
    const t = setTimeout(() => {
      setLoading(true)
      api.listTitles({ q, ...filters, page, size })
        .then(setData)
        .catch((e) => setMsg({ kind: 'bad', text: e.message }))
        .finally(() => setLoading(false))
    }, 250) // หน่วงเล็กน้อยระหว่างพิมพ์ค้นหา
    return () => clearTimeout(t)
  }, [q, filters, page, reload])

  const setF = (k, v) => { setFilters((f) => ({ ...f, [k]: v })); setPage(1) }
  const pages = Math.max(1, Math.ceil(data.total / size))

  async function onDelete(t) {
    if (!confirm(`ลบ "${t.title}" (ID ${t.id})?\nข้อมูลจะถูกลบจากฐานข้อมูลและไฟล์ Excel`)) return
    try {
      await api.deleteTitle(t.id)
      setMsg({ kind: 'ok', text: `ลบ "${t.title}" แล้ว — กำลังอัปเดตไฟล์ Excel` })
      setReload((r) => r + 1)
    } catch (e) {
      setMsg({ kind: 'bad', text: e.message })
    }
  }

  function onSaved(t, isNew) {
    setEditing(null)
    setMsg({ kind: 'ok', text: `${isNew ? 'เพิ่ม' : 'บันทึก'} "${t.title}" (ID ${t.id}) แล้ว — กำลังอัปเดตไฟล์ Excel` })
    setReload((r) => r + 1)
  }

  return (
    <Card
      title="จัดการข้อมูล"
      note={`ทั้งหมด ${fmt(data.total)} เรื่อง · แก้ไขที่นี่แล้ว sheet1_RawData ใน Excel จะอัปเดตเอง`}
      actions={<button className="btn" onClick={() => setEditing({ ...EMPTY })}>+ เพิ่มข้อมูล</button>}
    >
      {msg && <div className={'alert ' + msg.kind} onClick={() => setMsg(null)}>{msg.text}</div>}
      <div className="toolbar">
        <input className="search" placeholder="ค้นหาชื่อเรื่อง…" value={q} onChange={(e) => { setQ(e.target.value); setPage(1) }} />
        <select value={filters.type} onChange={(e) => setF('type', e.target.value)}>
          <option value="">ทุกประเภท</option><option value="TV">TV</option><option value="Film">Film</option>
        </select>
        <select value={filters.global} onChange={(e) => setF('global', e.target.value)}>
          <option value="">ทุกพื้นที่</option><option value="Yes">ฉายทั่วโลก</option><option value="No">ไม่ทั่วโลก</option>
        </select>
        <select value={filters.origin} onChange={(e) => setF('origin', e.target.value)}>
          <option value="">ทุกแหล่ง</option><option value="netflix">Netflix Title</option><option value="other">Non-Netflix</option>
        </select>
        <select value={filters.sort} onChange={(e) => setF('sort', e.target.value)}>
          <option value="views_desc">ยอดดู มาก→น้อย</option>
          <option value="views_asc">ยอดดู น้อย→มาก</option>
          <option value="hours_desc">ชั่วโมงดู มาก→น้อย</option>
          <option value="date_desc">ออกฉาย ใหม่→เก่า</option>
          <option value="title_asc">ชื่อ A→Z</option>
          <option value="updated_desc">แก้ไขล่าสุด</option>
          <option value="id_desc">ID ล่าสุด</option>
        </select>
      </div>

      <div className="scroll-x">
        <table className={'tbl data' + (loading ? ' dim' : '')}>
          <thead>
            <tr>
              <th className="r">ID</th><th>Type</th><th>Title</th><th>Global</th><th>Release Date</th>
              <th className="r">Hours Viewed</th><th className="r">Runtime</th><th className="r">Runtime (ชม.)</th><th className="r">Views</th><th></th>
            </tr>
          </thead>
          <tbody>
            {data.items.map((t) => (
              <tr key={t.id}>
                <td className="r muted">{t.id}</td>
                <td><span className={'pill ' + t.type.toLowerCase()}>{t.type}</span></td>
                <td className="title-cell">{t.title}</td>
                <td>{t.availableGlobally}</td>
                <td>{t.releaseDate || <span className="muted">–</span>}</td>
                <td className="r">{fmt(t.hoursViewed)}</td>
                <td className="r">{t.runtime || '–'}</td>
                <td className="r">{fmtDec(t.runtimeHours)}</td>
                <td className="r">{fmt(t.views)}</td>
                <td className="actions">
                  <button className="btn small ghost" onClick={() => setEditing({ ...t, releaseDate: t.releaseDate || '', runtime: t.runtime || '' })}>แก้ไข</button>
                  <button className="btn small danger" onClick={() => onDelete(t)}>ลบ</button>
                </td>
              </tr>
            ))}
            {data.items.length === 0 && !loading && (
              <tr><td colSpan="10" className="empty">ไม่พบข้อมูล</td></tr>
            )}
          </tbody>
        </table>
      </div>

      <div className="pager">
        <button className="btn small ghost" disabled={page <= 1} onClick={() => setPage(1)}>«</button>
        <button className="btn small ghost" disabled={page <= 1} onClick={() => setPage(page - 1)}>‹ ก่อนหน้า</button>
        <span>หน้า {page} / {fmt(pages)}</span>
        <button className="btn small ghost" disabled={page >= pages} onClick={() => setPage(page + 1)}>ถัดไป ›</button>
        <button className="btn small ghost" disabled={page >= pages} onClick={() => setPage(pages)}>»</button>
      </div>

      {editing && <TitleForm initial={editing} onClose={() => setEditing(null)} onSaved={onSaved} />}
    </Card>
  )
}

function TitleForm({ initial, onClose, onSaved }) {
  const [f, setF] = useState(initial)
  const [err, setErr] = useState('')
  const [saving, setSaving] = useState(false)
  const isNew = !initial.id
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value })

  async function submit(e) {
    e.preventDefault()
    setErr('')
    const hours = Number(String(f.hoursViewed).replace(/,/g, ''))
    const views = Number(String(f.views).replace(/,/g, ''))
    if (!f.title.trim()) return setErr('กรุณากรอกชื่อเรื่อง')
    if (!Number.isFinite(hours) || hours < 0 || String(f.hoursViewed).trim() === '') return setErr('Hours Viewed ต้องเป็นตัวเลขไม่ติดลบ')
    if (!Number.isFinite(views) || views < 0 || String(f.views).trim() === '') return setErr('Views ต้องเป็นตัวเลขไม่ติดลบ')
    if (f.runtime && !/^\d{1,3}:[0-5]\d$/.test(f.runtime.trim())) return setErr('Runtime ต้องเป็นรูปแบบ H:MM เช่น 1:45')
    const body = {
      type: f.type, title: f.title, availableGlobally: f.availableGlobally,
      releaseDate: f.releaseDate || null, hoursViewed: Math.round(hours),
      runtime: f.runtime.trim() || null, views: Math.round(views),
    }
    setSaving(true)
    try {
      const saved = isNew ? await api.createTitle(body) : await api.updateTitle(initial.id, body)
      onSaved(saved, isNew)
    } catch (e2) {
      setErr(e2.message)
    } finally {
      setSaving(false)
    }
  }

  return (
    <div className="modal-bg" onMouseDown={onClose}>
      <form className="modal" onSubmit={submit} onMouseDown={(e) => e.stopPropagation()}>
        <h2>{isNew ? 'เพิ่มข้อมูลใหม่' : `แก้ไขข้อมูล ID ${initial.id}`}</h2>
        <label>Title *<input value={f.title} onChange={set('title')} autoFocus maxLength={500} /></label>
        <div className="row2">
          <label>Type *
            <select value={f.type} onChange={set('type')}><option>Film</option><option>TV</option></select>
          </label>
          <label>Available Globally *
            <select value={f.availableGlobally} onChange={set('availableGlobally')}><option>Yes</option><option>No</option></select>
          </label>
        </div>
        <div className="row2">
          <label>Release Date<input type="date" value={f.releaseDate} onChange={set('releaseDate')} /></label>
          <label>Runtime (H:MM)<input placeholder="1:45" value={f.runtime} onChange={set('runtime')} /></label>
        </div>
        <div className="row2">
          <label>Hours Viewed *<input inputMode="numeric" placeholder="เช่น 1500000" value={f.hoursViewed} onChange={set('hoursViewed')} /></label>
          <label>Views *<input inputMode="numeric" placeholder="เช่น 800000" value={f.views} onChange={set('views')} /></label>
        </div>
        <p className="note">ปล่อย Release Date ว่างถ้าไม่ใช่คอนเทนต์ของ Netflix · Runtime (ชม.) คำนวณให้อัตโนมัติ</p>
        {err && <div className="alert bad">{err}</div>}
        <div className="modal-actions">
          <button type="button" className="btn ghost" onClick={onClose}>ยกเลิก</button>
          <button className="btn" disabled={saving}>{saving ? 'กำลังบันทึก…' : 'บันทึก'}</button>
        </div>
      </form>
    </div>
  )
}
