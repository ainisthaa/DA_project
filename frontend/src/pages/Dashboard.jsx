import { useEffect, useState } from 'react'
import { api, fmt, fmtPct, fmtShort, fmtDec } from '../api.js'

// หน้าภาพรวม: แสดงค่าที่คำนวณได้ (ชุดเดียวกับ sheet2_Analysis ใน Excel)
export default function Dashboard({ sync }) {
  const [stats, setStats] = useState(null)
  const [excel, setExcel] = useState(null)
  const [err, setErr] = useState('')
  const version = sync ? `${sync.status.lastImport}|${sync.status.lastExport}|${sync.status.lastRecalc}|${sync.status.pendingExport}|${sync.dbRows}` : ''

  useEffect(() => {
    api.stats().then(setStats).catch((e) => setErr(e.message))
    api.excelSummary().then(setExcel).catch(() => setExcel(null))
  }, [version])

  if (err) return <div className="alert bad">{err}</div>
  if (!stats) return <div className="loading">กำลังโหลด…</div>
  const k = stats.kpi

  return (
    <>
      <ExcelCheck stats={stats} excel={excel} />
      <section className="kpis">
        <Kpi label="จำนวนเรื่องทั้งหมด" value={fmt(k.titles)} />
        <Kpi label="ชั่วโมงดูรวม" value={fmtShort(k.hoursViewed)} sub={fmt(k.hoursViewed) + ' ชม.'} />
        <Kpi label="ยอดดูรวม (Views)" value={fmtShort(k.views)} sub={fmt(k.views)} />
        <Kpi label="ยอดดูเฉลี่ยต่อเรื่อง" value={fmtShort(k.avgViews)} sub={'มัธยฐาน ' + fmtShort(k.medianViews)} />
      </section>

      <div className="grid2">
        <Card title="TV กับ Film" note="TV มีจำนวนเรื่องน้อยกว่าแต่ครองชั่วโมงดูส่วนใหญ่">
          <GroupTable rows={stats.byType} cols={['titles', 'hoursViewed', 'shareHours', 'avgViews', 'medianViews', 'maxViews', 'avgRuntime']} />
        </Card>
        <Card title="Netflix Title กับ Non-Netflix" note="เกณฑ์: มี Release Date = คอนเทนต์ของ Netflix (เกณฑ์แทน ไม่ใช่ป้ายทางการ)">
          <GroupTable rows={stats.byOrigin} cols={['titles', 'hoursViewed', 'shareHours', 'avgViews', 'medianViews']} />
        </Card>
        <Card title="ฉายทั่วโลก (Available Globally)">
          <GroupTable rows={stats.byGlobal} cols={['titles', 'hoursViewed', 'shareHours', 'avgViews', 'medianViews']} />
        </Card>
        <Card title="ความยาวกับยอดดู" note="TV ใช้ความยาวรวมทั้งซีซัน, Film ใช้ความยาวหนัง">
          <GroupTable rows={stats.runtime} cols={['titles', 'avgViews', 'medianViews']} bars="avgViews" />
        </Card>
        <Card title="Netflix Title ตามปีที่ออกฉาย">
          <GroupTable rows={stats.byYear} cols={['titles', 'hoursViewed', 'avgViews']} bars="hoursViewed" />
        </Card>
        <Card title="การกระจุกตัวของยอดดู (Pareto)" note="เรื่องที่ยอดดูสูงสุด X% ครองชั่วโมงดูรวมกี่ %">
          <table className="tbl">
            <thead><tr><th>กลุ่ม</th><th className="r">จำนวนเรื่อง</th><th>% ของชั่วโมงดู</th></tr></thead>
            <tbody>
              {stats.pareto.map((p) => (
                <tr key={p.label}>
                  <td>{p.label}</td><td className="r">{fmt(p.titles)}</td>
                  <td><Bar v={p.shareHours} max={1} label={fmtPct(p.shareHours)} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>
      </div>

      <div className="grid2">
        <Card title="Top 10 Film"><TopTable rows={stats.top10.film} /></Card>
        <Card title="Top 10 TV"><TopTable rows={stats.top10.tv} /></Card>
      </div>
    </>
  )
}

// เทียบค่าที่ Excel คำนวณ (sheet4) กับค่าที่คำนวณจากฐานข้อมูล
function ExcelCheck({ stats, excel }) {
  if (!excel) return null
  const get = (m) => excel.rows.find((r) => r.section === 'KPI' && r.metric === m)?.value
  const titles = get('Total Titles'), avg = get('Average Views')
  const match = titles === stats.kpi.titles && avg != null && Math.abs(avg - stats.kpi.avgViews) < 0.01
  if (excel.stale || !match) {
    return (
      <div className="alert warn">
        ⏳ Excel ยังคำนวณสูตรไม่เสร็จ — ค่าบนหน้านี้คำนวณจากฐานข้อมูลโดยตรง
        (sheet2 จะอัปเดตเมื่อระบบคำนวณใหม่หรือเปิดไฟล์ใน Excel)
      </div>
    )
  }
  return (
    <div className="alert ok">
      ✓ ค่าที่ Excel คำนวณ (sheet2–4) ตรงกับฐานข้อมูล · คำนวณล่าสุด {excel.lastRecalc || excel.excelModified}
    </div>
  )
}

const COLS = {
  titles: ['จำนวนเรื่อง', fmt],
  hoursViewed: ['ชั่วโมงดูรวม', fmtShort],
  shareHours: ['% ชั่วโมงดู', fmtPct],
  views: ['ยอดดูรวม', fmtShort],
  avgViews: ['ยอดดูเฉลี่ย', fmtShort],
  medianViews: ['มัธยฐาน', fmtShort],
  maxViews: ['สูงสุด', fmtShort],
  avgRuntime: ['Runtime (ชม.)', (v) => fmtDec(v)],
}

function GroupTable({ rows, cols, bars }) {
  const max = bars ? Math.max(...rows.map((r) => r[bars]), 1) : 0
  return (
    <div className="scroll-x">
      <table className="tbl">
        <thead>
          <tr><th>กลุ่ม</th>{cols.map((c) => <th key={c} className="r">{COLS[c][0]}</th>)}{bars && <th></th>}</tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.group}>
              <td>{r.group}</td>
              {cols.map((c) => <td key={c} className="r">{COLS[c][1](r[c])}</td>)}
              {bars && <td className="barcell"><Bar v={r[bars]} max={max} /></td>}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function TopTable({ rows }) {
  return (
    <table className="tbl">
      <thead><tr><th>#</th><th>Title</th><th className="r">Views</th><th className="r">Hours</th></tr></thead>
      <tbody>
        {rows.map((t, i) => (
          <tr key={t.id}>
            <td className="muted">{i + 1}</td><td>{t.title}</td>
            <td className="r">{fmtShort(t.views)}</td><td className="r">{fmtShort(t.hoursViewed)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  )
}

function Bar({ v, max, label }) {
  const w = Math.max(0, Math.min(100, (v / max) * 100))
  return (
    <div className="bar"><div className="bar-fill" style={{ width: w + '%' }} />{label && <span>{label}</span>}</div>
  )
}

function Kpi({ label, value, sub }) {
  return (
    <div className="kpi">
      <div className="kpi-label">{label}</div>
      <div className="kpi-value">{value}</div>
      {sub && <div className="kpi-sub">{sub}</div>}
    </div>
  )
}

export function Card({ title, note, children, actions }) {
  return (
    <section className="card">
      <div className="card-head">
        <div>
          <h2>{title}</h2>
          {note && <p className="note">{note}</p>}
        </div>
        {actions && <div className="card-actions">{actions}</div>}
      </div>
      {children}
    </section>
  )
}
