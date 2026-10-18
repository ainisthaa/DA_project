import { useRef, useState } from 'react'
import { api, fmt } from '../api.js'
import { Card } from './Dashboard.jsx'

// หน้าซิงก์: ดูสถานะการเชื่อม Excel ↔ MySQL, สั่งซิงก์เอง, ดาวน์โหลด/อัปโหลดไฟล์
export default function SyncPage({ sync }) {
  const [busy, setBusy] = useState('')
  const [msg, setMsg] = useState(null)
  const fileRef = useRef()

  async function run(label, fn) {
    setBusy(label)
    setMsg(null)
    try {
      const r = await fn()
      setMsg({ kind: 'ok', text: label + ' สำเร็จ' + (r && r.rows != null ? ` — อ่าน ${fmt(r.rows)} แถว, เพิ่ม ${r.inserted}, แก้ ${r.updated}, ลบ ${r.deleted}, แถวที่ผิด ${r.rowErrors.length}` : '') })
    } catch (e) {
      setMsg({ kind: 'bad', text: e.message })
    } finally {
      setBusy('')
    }
  }

  function onUpload(e) {
    const f = e.target.files[0]
    e.target.value = ''
    if (!f) return
    const force = confirm('ถ้าไฟล์ที่อัปโหลดมีข้อมูลน้อยกว่าครึ่งของในระบบ ต้องการบังคับนำเข้าหรือไม่?\n(กด Cancel = นำเข้าแบบปลอดภัย)')
    run('อัปโหลดและนำเข้า', () => api.uploadExcel(f, force))
  }

  if (!sync) return <div className="alert bad">ติดต่อ backend ไม่ได้ — ตรวจว่าโปรแกรม backend ทำงานอยู่</div>
  const s = sync.status

  return (
    <>
      <div className="grid2">
        <Card title="สถานะการเชื่อม Excel ↔ MySQL">
          <dl className="kv">
            <dt>ไฟล์ Excel</dt><dd><code>{s.excelPath}</code></dd>
            <dt>ไฟล์แก้ไขล่าสุด</dt><dd>{s.excelModified || '–'}</dd>
            <dt>จำนวนแถวในฐานข้อมูล</dt><dd>{fmt(sync.dbRows)}</dd>
            <dt>นำเข้าจาก Excel ล่าสุด</dt><dd>{s.lastImport || '–'}</dd>
            <dt>เขียนลง Excel ล่าสุด</dt><dd>{s.lastExport || '–'}</dd>
            <dt>คำนวณสูตรใหม่ล่าสุด</dt><dd>{s.recalcEnabled ? (s.lastRecalc || '–') : 'ปิดอยู่ (เปิดไฟล์ใน Excel เพื่อคำนวณ)'}</dd>
            <dt>ตรวจไฟล์ทุก</dt><dd>{s.pollSeconds} วินาที</dd>
            <dt>สถานะ</dt>
            <dd>{s.lastError ? <span className="txt-bad">{s.lastError}</span> : s.pendingExport ? 'กำลังเขียน Excel…' : <span className="txt-ok">ปกติ</span>}</dd>
          </dl>
        </Card>
        <Card title="คำสั่ง">
          {msg && <div className={'alert ' + msg.kind}>{msg.text}</div>}
          <div className="btn-col">
            <a className="btn" href="/api/excel/download">⬇ ดาวน์โหลดไฟล์ Excel</a>
            <button className="btn ghost" disabled={!!busy} onClick={() => fileRef.current.click()}>⬆ อัปโหลดไฟล์ Excel ที่แก้แล้ว</button>
            <input ref={fileRef} type="file" accept=".xlsx" hidden onChange={onUpload} />
            <button className="btn ghost" disabled={!!busy} onClick={() => run('นำเข้าจาก Excel', () => api.importExcel(false))}>Excel → ฐานข้อมูล</button>
            <button className="btn ghost" disabled={!!busy} onClick={() => run('เขียนลง Excel', () => api.exportExcel())}>ฐานข้อมูล → Excel</button>
            <button
              className="btn danger" disabled={!!busy}
              onClick={() => confirm('บังคับนำเข้า: ข้อมูลในฐานข้อมูลจะถูกแทนที่ด้วยข้อมูลใน Excel ทั้งหมด (แถวที่ไม่มีใน Excel จะถูกลบ) ยืนยัน?') && run('นำเข้าแบบบังคับ', () => api.importExcel(true))}
            >นำเข้าจาก Excel (บังคับ)</button>
          </div>
          {busy && <p className="note">กำลัง{busy}…</p>}
          <p className="note">
            การซิงก์ทำงานอัตโนมัติ: แก้บนเว็บ → Excel อัปเดตใน ~1 วินาที · บันทึกไฟล์ใน Excel → ฐานข้อมูลอัปเดตภายใน {s.pollSeconds} วินาที
          </p>
        </Card>
      </div>

      {s.rowErrors && s.rowErrors.length > 0 && (
        <Card title={`แถวใน Excel ที่นำเข้าไม่ได้ (${s.rowErrors.length})`} note="แก้ไขแถวเหล่านี้ใน sheet1_RawData แล้วบันทึกไฟล์อีกครั้ง">
          <table className="tbl">
            <thead><tr><th className="r">แถวที่</th><th>สาเหตุ</th></tr></thead>
            <tbody>{s.rowErrors.slice(0, 100).map((e) => <tr key={e.row}><td className="r">{e.row}</td><td>{e.message}</td></tr>)}</tbody>
          </table>
        </Card>
      )}

      <Card title="ประวัติการซิงก์">
        <table className="tbl log">
          <thead><tr><th>เวลา</th><th>ประเภท</th><th>ผล</th><th>รายละเอียด</th></tr></thead>
          <tbody>
            {sync.logs.map((l) => (
              <tr key={l.id}>
                <td className="nowrap">{l.at}</td><td>{l.direction}</td>
                <td><span className={'pill ' + l.status}>{l.status}</span></td>
                <td className="pre">{l.message}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>
    </>
  )
}
