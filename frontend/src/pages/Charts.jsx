import { useEffect, useState } from 'react'
import { api } from '../api.js'
import { Card } from './Dashboard.jsx'

// ทำลิงก์ Tableau Public ให้อยู่ในรูปแบบที่ฝังใน iframe ได้
export function embedUrl(raw) {
  try {
    const u = new URL(raw)
    if (u.hostname.includes('tableau.com')) {
      // ลิงก์แชร์มักมี ?:language=...&:display_count=...&:origin=viz_share_link
      u.searchParams.set(':embed', 'y')
      u.searchParams.set(':showVizHome', 'no')
      u.searchParams.set(':toolbar', 'yes')
      u.searchParams.delete(':origin')
    }
    return u.toString()
  } catch {
    return raw
  }
}

// แท็บในหน้านี้มี 2 แบบ:
//  - live  = ลิงก์ Tableau Public ที่ตั้งในหน้าตั้งค่า (ฝัง iframe โต้ตอบได้ ต้องมีอินเทอร์เน็ต)
//  - image = ภาพ Dashboard ที่ Export จาก tableau_netflix2 ไว้ใน data/tableau/ (แสดงได้เสมอ แม้ออฟไลน์)
export default function Charts() {
  const [tabs, setTabs] = useState(null)
  const [workbook, setWorkbook] = useState(null)
  const [active, setActive] = useState(0)

  useEffect(() => {
    Promise.all([
      api.getSettings().then((s) => s.tableauViews).catch(() => []),
      api.tableau().catch(() => ({ workbook: null, images: [] })),
    ]).then(([views, tb]) => {
      setWorkbook(tb.workbook)
      setTabs([
        ...views.map((v) => ({ kind: 'live', name: v.name, url: v.url })),
        ...tb.images.map((img) => ({ kind: 'image', name: img.name, url: img.url })),
      ])
    })
  }, [])

  if (tabs === null) return <div className="loading">กำลังโหลด…</div>
  const download = workbook && (
    <a className="btn ghost" href="/api/tableau/workbook">ดาวน์โหลด {workbook.name}</a>
  )
  if (tabs.length === 0) {
    return (
      <Card title="ยังไม่มี Tableau Dashboard ให้แสดง" actions={download}>
        <ol className="steps">
          <li>
            ทำกราฟใน Tableau จากไฟล์ <code>netflix_data.xlsx</code>
            {workbook ? <> (พบไฟล์ <code>{workbook.name}</code> แล้ว)</> : <> แล้วบันทึกเป็น <code>data/tableau_netflix2.twbx</code></>}
          </li>
          <li>แสดงแบบออนไลน์: Publish ขึ้น Tableau Public → Share → คัดลอกลิงก์ → วางในหน้า <a href="#settings">ตั้งค่า</a></li>
          <li>แสดงแบบออฟไลน์: บันทึกภาพ Dashboard (.png) ไว้ในโฟลเดอร์ <code>data/tableau/</code> — จากหน้า viz บน Tableau Public กด Download → Image (Tableau Desktop ใช้ Dashboard → Export Image…)</li>
        </ol>
      </Card>
    )
  }
  const v = tabs[Math.min(active, tabs.length - 1)]
  const actions = (
    <>
      {v.kind === 'live'
        ? <a className="btn ghost" href={v.url} target="_blank" rel="noreferrer">เปิดใน Tableau ↗</a>
        : <a className="btn ghost" href={v.url} target="_blank" rel="noreferrer">ดูภาพขนาดเต็ม ↗</a>}
      {download}
    </>
  )
  return (
    <>
      {tabs.length > 1 && (
        <div className="subtabs">
          {tabs.map((x, i) => (
            <button key={i} className={i === active ? 'active' : ''} onClick={() => setActive(i)}>
              {x.name}{x.kind === 'image' ? ' (ภาพ)' : ''}
            </button>
          ))}
        </div>
      )}
      <Card title={v.name} actions={actions}>
        {v.kind === 'live' ? (
          <div className="tableau-frame">
            <iframe key={v.url} title={v.name} src={embedUrl(v.url)} allowFullScreen loading="lazy" />
          </div>
        ) : (
          <div className="tableau-image">
            <img src={v.url} alt={v.name} />
          </div>
        )}
        <p className="note">
          {v.kind === 'live'
            ? 'กราฟอ่านข้อมูลจากไฟล์ Excel ที่ publish ขึ้น Tableau Public — เมื่อข้อมูลเปลี่ยน ให้เปิด Tableau กด Refresh data source แล้ว Save ไปที่ Tableau Public อีกครั้ง'
            : 'ภาพนิ่งจาก Tableau (tableau_netflix2) ใช้ได้แม้ไม่มีอินเทอร์เน็ต — เมื่อข้อมูลเปลี่ยน ให้ Refresh แล้ว Save to Tableau Public จากนั้นบันทึกภาพใหม่ทับไฟล์เดิมใน data/tableau/'}
        </p>
      </Card>
    </>
  )
}
