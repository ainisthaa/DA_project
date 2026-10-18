// เรียก REST API ของ backend Go (ทุก path ขึ้นต้นด้วย /api)
async function request(method, path, body) {
  const opts = { method, headers: {} }
  if (body instanceof FormData) {
    opts.body = body
  } else if (body !== undefined) {
    opts.headers['Content-Type'] = 'application/json'
    opts.body = JSON.stringify(body)
  }
  const res = await fetch(path, opts)
  const data = await res.json().catch(() => ({}))
  if (!res.ok) throw new Error(data.error || `เกิดข้อผิดพลาด (${res.status})`)
  return data
}

export const api = {
  listTitles: (params) => request('GET', '/api/titles?' + new URLSearchParams(params)),
  createTitle: (t) => request('POST', '/api/titles', t),
  updateTitle: (id, t) => request('PUT', `/api/titles/${id}`, t),
  deleteTitle: (id) => request('DELETE', `/api/titles/${id}`),
  stats: () => request('GET', '/api/stats'),
  excelSummary: () => request('GET', '/api/excel/summary'),
  syncStatus: () => request('GET', '/api/sync/status'),
  importExcel: (force) => request('POST', '/api/sync/import' + (force ? '?force=1' : '')),
  exportExcel: () => request('POST', '/api/sync/export'),
  uploadExcel: (file, force) => {
    const fd = new FormData()
    fd.append('file', file)
    return request('POST', '/api/excel/upload' + (force ? '?force=1' : ''), fd)
  },
  getSettings: () => request('GET', '/api/settings'),
  saveSettings: (s) => request('PUT', '/api/settings', s),
  tableau: () => request('GET', '/api/tableau'),
}

// ---------- ตัวช่วยจัดรูปแบบตัวเลข ----------
export const fmt = (n) => (n == null ? '–' : Math.round(n).toLocaleString('en-US'))
export const fmtPct = (n) => (n == null ? '–' : (n * 100).toFixed(1) + '%')
export const fmtDec = (n, d = 2) => (n == null ? '–' : Number(n).toFixed(d))
export function fmtShort(n) {
  if (n == null) return '–'
  const a = Math.abs(n)
  if (a >= 1e9) return (n / 1e9).toFixed(2) + 'B'
  if (a >= 1e6) return (n / 1e6).toFixed(2) + 'M'
  if (a >= 1e3) return (n / 1e3).toFixed(1) + 'K'
  return String(Math.round(n))
}
