import { useState, useEffect } from 'react'

export default function ProjectList({ onLoadProject }) {
  const [grouped, setGrouped] = useState([])
  const [search, setSearch] = useState('')
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    fetch('/api/projects/list')
      .then(r => r.json())
      .then(data => {
        // Group by year→month
        const result = []
        const seenY = new Set(), seenM = new Set()
        data.forEach(p => {
          const date = p.date || ''
          const year = date.slice(0,4)
          const ym = date.slice(0,7)
          if (year && !seenY.has(year)) { seenY.add(year); result.push({type:'year',label:year}) }
          if (ym && !seenM.has(ym)) { seenM.add(ym); result.push({type:'month',label:ym.slice(5,7).replace(/^0/,'')+'月',count:data.filter(pp=>(pp.date||'').slice(0,7)===ym).length}) }
          result.push({type:'item',data:p,ym})
        })
        setGrouped(result)
        setLoading(false)
      })
      .catch(() => setLoading(false))
  }, [])

  const filtered = search ? grouped.filter(g => g.type!=='item' || (g.data.name||'').toLowerCase().includes(search.toLowerCase())) : grouped

  return (
    <div className="container">
      <h1 className="page-title">工程單列表</h1>
      <p style={{color:'var(--text-secondary)',fontSize:13,marginBottom:14}}>同步自 Google Drive</p>
      <input className="search-input" type="text" placeholder="搜尋工程單…" value={search} onChange={e=>setSearch(e.target.value)} />

      {loading ? <div className="empty-hint">載入中…</div> : (
        <div className="project-list" style={{background:'none',boxShadow:'none',border:'none'}}>
          {filtered.map((g, i) => {
            if (g.type === 'year') return (
              <details key={g.label} className="year-card" open>
                <summary className="year-header">{g.label}</summary>
              </details>
            )
            if (g.type === 'month') return (
              <div key={g.label+i} className="month-header">
                <span className="month-label">{g.label}</span>
                <span className="month-count">{g.count} 張單</span>
              </div>
            )
            return (
              <div key={g.data.id} className="project-item" onClick={async () => {
                try {
                  const resp = await fetch('/api/projects/'+g.data.id)
                  if (!resp.ok) throw new Error('載入失敗')
                  const data = await resp.json()
                  onLoadProject(data)
                } catch(e) { alert('載入失敗：'+e.message) }
              }}>
                <span style={{fontSize:18}}>📄</span>
                <div style={{flex:1,minWidth:0}}>
                  <div className="project-item-name">{g.data.name}</div>
                  <div style={{fontSize:11,color:'var(--text-secondary)'}}>{g.data.size_kb}KB · {g.data.date}</div>
                </div>
              </div>
            )
          })}
        </div>
      )}
    </div>
  )
}
