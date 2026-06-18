import { useState, useEffect, useCallback } from 'react'

// Category rules (same as original)
const CATEGORY_RULES = [
  {name:'安裝代工',keywords:['安裝.*人工','代裝','代工','人工安裝','裝工','連裝工','安裝費','安裝人工']},
  {name:'清拆工程',keywords:['清拆','清折','清走','拆牆','拆門','拆櫃','拆磚','垃圾','清場','搭棚','棚架','保險費','保護']},
  {name:'水電工程',keywords:['水喉','來去水','去水喉','電制','燈制','插蘇','蘇位','供電','菲士箱','煤氣','冷氣','抽氣扇','熱水爐']},
  {name:'泥水工程',keywords:['鋪磚','磁磚','瓷磚','包底料','地台','英泥沙','防水','地板','企缸','鋁窗','窗台']},
  {name:'油漆工程',keywords:['油漆','批灰','油油','起底','剷底','ICI','乳膠漆','牆紙']},
  {name:'木工工程',keywords:['門','櫃','床','天花','腳線','廚櫃','衣櫃','雲石','枱面']},
  {name:'雜項',keywords:['代購','代付','吸咀','浴屏','潔具','廁所','浴室','清潔','搬運']},
]

function autoCategorize(desc) {
  for (const r of CATEGORY_RULES) {
    for (const k of r.keywords) {
      try { if (new RegExp(k, 'i').test(desc)) return r.name }
      catch(e) { if (desc.includes(k)) return r.name }
    }
  }
  return '雜項'
}

function parseBlock(text, isAdd) {
  const items = []
  const lines = text.split('\n')
  for (const line of lines) {
    const t = line.trim(); if (!t) continue
    const m = t.match(/^(\d+)[.\s、．]*\s*(.+?)\s+(\d+)\s*$/)
    if (m) { items.push({is_additional:isAdd, category:autoCategorize(m[2].trim()), description:m[2].trim(), quantity:1, unit:'項', unit_price:parseInt(m[3],10), remark:''}); continue }
    const lm = t.match(/^(.+?)\s+(\d+)\s*$/)
    if (lm) { const d=lm[1].replace(/^\d+[.\s、．]*\s*/,'').trim(); if(d&&parseInt(lm[2],10)>=0) { items.push({is_additional:isAdd, category:autoCategorize(d), description:d, quantity:1, unit:'項', unit_price:parseInt(lm[2],10), remark:''}); continue } }
    const nm = t.match(/^(\d+)[.\s、．]*\s*(.+)$/)
    if (nm) { const d=nm[2].trim(); if(d) { items.push({is_additional:isAdd, category:autoCategorize(d), description:d, quantity:1, unit:'項', unit_price:0, remark:''}); continue } }
    if (t.length>1) items.push({is_additional:isAdd, category:autoCategorize(t), description:t, quantity:1, unit:'項', unit_price:0, remark:''})
  }
  return items
}

export default function QuotationForm({ loadData }) {
  const [basicText, setBasicText] = useState('')
  const [additionalText, setAdditionalText] = useState('')
  const [items, setItems] = useState([])
  const [projectName, setProjectName] = useState('')
  const [ownerName, setOwnerName] = useState('')
  const [address, setAddress] = useState('')
  const [companyName, setCompanyName] = useState('')
  const [quotationNo, setQuotationNo] = useState('')
  const [date, setDate] = useState(new Date().toISOString().split('T')[0])
  const [validity, setValidity] = useState('')
  const [version, setVersion] = useState('')
  const [deposit, setDeposit] = useState(0)
  const [showPayment, setShowPayment] = useState(false)
  const [showTerms, setShowTerms] = useState(false)
  const [payments, setPayments] = useState([])
  const [terms, setTerms] = useState([])
  const [msg, setMsg] = useState('')
  const [dlg, setDlg] = useState(null) // {title, filename, downloadUrl}

  // Load data from project list
  useEffect(() => {
    if (!loadData) return
    const d = loadData
    if (d.project_name) setProjectName(d.project_name)
    if (d.owner_name) setOwnerName(d.owner_name)
    if (d.address) setAddress(d.address)
    if (d.company_name) setCompanyName(d.company_name)
    if (d.quotation_no) setQuotationNo(d.quotation_no)
    if (d.date) setDate(d.date)
    if (d.validity) setValidity(d.validity)
    if (d.version) setVersion(d.version)
    if (d.deposit) setDeposit(d.deposit)
    if (d.items) {
      const bl = [], al = []
      d.items.forEach(it => {
        const line = `${it.description} ${it.unit_price}`
        if (it.is_additional) al.push(line); else bl.push(line)
      })
      setBasicText(bl.join('\n'))
      setAdditionalText(al.join('\n'))
      const parsed = [...parseBlock(bl.join('\n'), false), ...parseBlock(al.join('\n'), true)]
      setItems(parsed)
    }
  }, [loadData])

  const handleParse = useCallback(() => {
    const parsed = [...parseBlock(basicText, false), ...parseBlock(additionalText, true)]
    setItems(parsed)
  }, [basicText, additionalText])

  const updateItem = (idx, field, value) => {
    setItems(prev => prev.map((it, i) => i === idx ? {...it, [field]: value} : it))
  }

  const deleteItem = (idx) => {
    setItems(prev => prev.filter((_, i) => i !== idx))
  }

  const addEmptyItem = () => {
    setItems(prev => [...prev, {is_additional:false, category:'雜項', description:'', quantity:1, unit:'項', unit_price:0, remark:''}])
  }

  const grandTotal = items.reduce((sum, it) => sum + (it.quantity||1)*(it.unit_price||0), 0)

  const handleGenerate = async (type) => {
    if (items.length === 0) { setMsg('請至少填寫一個工程項目！'); return }
    setMsg('正在生成中…')
    try {
      const data = {
        project_name: projectName||'-', owner_name: ownerName, address, company_name: companyName,
        quotation_no: quotationNo, date, validity, version, deposit,
        items, show_payment: showPayment, payments, show_terms: showTerms, terms,
      }
      const resp = await fetch('/generate', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({type, data})})
      if (!resp.ok) throw new Error((await resp.json()).error||'生成失敗')
      const result = await resp.json()
      if (result.preview_id) {
        const tname = type==='invoice'?'發票':'報價單'
        const parts = []; if(quotationNo) parts.push(quotationNo); if(address) parts.push(address); if(ownerName) parts.push(ownerName)
        const ds = date.replace(/-/g,''); if(ds) parts.push(ds); if(!parts.length) parts.push(ds||'-')
        setDlg({title: tname+'已生成', filename: parts.join('_')+'_'+tname+'.xlsx', downloadUrl:'/download/'+result.preview_id+'/excel'})
        setMsg('')
      }
    } catch(err) { setMsg('生成失敗：'+err.message) }
  }

  return (
    <div className="container">
      <h1 className="page-title">新增報價單</h1>

      {/* Upload bar */}
      <div className="card" style={{textAlign:'center',padding:'12px 18px'}}>
        <input type="file" id="uploadFile" accept=".xlsx" style={{display:'none'}} onChange={async (e) => {
          const file = e.target.files[0]; if(!file) return
          const fd = new FormData(); fd.append('file', file)
          try {
            const resp = await fetch('/upload', {method:'POST', body:fd})
            if (!resp.ok) throw new Error((await resp.json()).error)
            const data = await resp.json()
            if (data.project_name) setProjectName(data.project_name||'')
            if (data.owner_name) setOwnerName(data.owner_name||'')
            if (data.address) setAddress(data.address||'')
            if (data.company_name) setCompanyName(data.company_name||'')
            if (data.quotation_no) setQuotationNo(data.quotation_no||'')
            if (data.date) setDate(data.date||'')
            if (data.validity) setValidity(data.validity||'')
            if (data.version) setVersion(data.version||'')
            if (data.deposit) setDeposit(data.deposit||0)
            if (data.items) {
              const bl=[], al=[]
              data.items.forEach(it => { const line=`${it.description} ${it.unit_price}`; if(it.is_additional) al.push(line); else bl.push(line) })
              setBasicText(bl.join('\n')); setAdditionalText(al.join('\n'))
              setItems([...parseBlock(bl.join('\n'),false), ...parseBlock(al.join('\n'),true)])
            }
            if (data.payments) setPayments(data.payments)
            if (data.terms) setTerms(data.terms)
          } catch(err) { alert('匯入失敗：'+err.message) }
        }} />
        <label htmlFor="uploadFile" className="btn btn-secondary btn-sm" style={{cursor:'pointer'}}>匯入 Excel（反向填表）</label>
      </div>

      {/* Basic Info */}
      <div className="card">
        <div className="card-header">基本資料</div>
        <div className="info-grid">
          <div className="form-group"><label>工程名稱</label><input type="text" value={projectName} onChange={e=>setProjectName(e.target.value)} placeholder="工程名稱…" /></div>
          <div className="form-group"><label>客戶姓名</label><input type="text" value={ownerName} onChange={e=>setOwnerName(e.target.value)} placeholder="客戶姓名…" /></div>
          <div className="form-group"><label>工程地址</label><input type="text" value={address} onChange={e=>setAddress(e.target.value)} placeholder="工程地址…" /></div>
          <div className="form-group"><label>裝修公司名 + BR</label><input type="text" value={companyName} onChange={e=>setCompanyName(e.target.value)} placeholder="公司名 + BR…" /></div>
          <div className="form-group"><label>報價單號</label><input type="text" value={quotationNo} onChange={e=>setQuotationNo(e.target.value)} placeholder="報價單號…" /></div>
          <div className="form-group"><label>報價日期</label><input type="date" value={date} onChange={e=>setDate(e.target.value)} /></div>
          <div className="form-group"><label>有效期</label><input type="text" value={validity} onChange={e=>setValidity(e.target.value)} placeholder="有效期…" /></div>
          <div className="form-group"><label>版本</label><input type="text" value={version} onChange={e=>setVersion(e.target.value)} placeholder="版本…" /></div>
          <div className="form-group"><label>已付訂金 (HKD)</label><input type="number" value={deposit} onChange={e=>setDeposit(parseFloat(e.target.value)||0)} min="0" /></div>
        </div>
      </div>

      {/* Items */}
      <div className="card">
        <div className="card-header">工程項目</div>
        <div className="textarea-row">
          <div className="textarea-group">
            <label>基本工程 <span className="hint">— {'{序號}{描述} {價錢}'}（價錢可省略）</span></label>
            <textarea value={basicText} onChange={e=>setBasicText(e.target.value)} placeholder="1兩間浴室清拆磁磚 1000&#10;2全屋地板清拆 2000&#10;…" />
          </div>
          <div className="textarea-group">
            <label>後加工程 <span className="hint">— 冇後加就留空</span></label>
            <textarea value={additionalText} onChange={e=>setAdditionalText(e.target.value)} placeholder="1全屋新造水喉 21000&#10;…" />
          </div>
        </div>
        <button className="btn btn-primary btn-cta" onClick={handleParse} style={{display:'block',margin:'0 auto'}}>解析項目（自動分類）</button>
        <div className="parse-count">{items.length > 0 ? `已解析 ${items.length} 個項目` : ''}</div>

        <div className="table-wrap">
          <table>
            <thead><tr>
              <th style={{width:48}}>序號</th><th style={{width:72}}>階段</th><th style={{width:100}}>分類</th>
              <th style={{minWidth:220}}>項目描述</th><th style={{width:55}}>數量</th><th style={{width:55}}>單位</th>
              <th style={{width:85}}>單價(HKD)</th><th style={{width:95}}>金額(HKD)</th>
              <th style={{width:140}}>備註</th><th style={{width:45}}></th>
            </tr></thead>
            <tbody>
              {items.length === 0 ? (
                <tr><td colSpan={10} className="empty-hint">請喺上方文字框貼入項目，然後點擊「解析項目」</td></tr>
              ) : items.map((it, idx) => {
                const amt = (it.quantity||1)*(it.unit_price||0)
                return (
                  <tr key={idx} className={idx%2===1?'odd-row':''}>
                    <td style={{textAlign:'center'}}>{idx+1}</td>
                    <td><select value={it.is_additional?'additional':'basic'} onChange={e=>updateItem(idx,'is_additional',e.target.value==='additional')}><option value="basic">基本</option><option value="additional">後加</option></select></td>
                    <td><select value={it.category} onChange={e=>updateItem(idx,'category',e.target.value)}>{CATEGORY_RULES.map(r=><option key={r.name} value={r.name}>{r.name}</option>)}</select></td>
                    <td><input type="text" value={it.description} onChange={e=>updateItem(idx,'description',e.target.value)} style={{textAlign:'left'}} /></td>
                    <td><input type="number" value={it.quantity||1} onChange={e=>updateItem(idx,'quantity',parseFloat(e.target.value)||1)} min="0" style={{width:'100%'}} /></td>
                    <td><input type="text" value={it.unit||'項'} onChange={e=>updateItem(idx,'unit',e.target.value)} /></td>
                    <td><input type="number" value={it.unit_price||0} onChange={e=>updateItem(idx,'unit_price',parseFloat(e.target.value)||0)} min="0" /></td>
                    <td style={{fontWeight:'bold',color:'#002F5597'}}>${amt.toLocaleString('en-US')}</td>
                    <td><input type="text" value={it.remark||''} onChange={e=>updateItem(idx,'remark',e.target.value)} style={{textAlign:'left'}} /></td>
                    <td><button className="btn btn-danger btn-sm btn-del" onClick={()=>deleteItem(idx)}>✕</button></td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
        <button className="btn btn-secondary btn-sm" onClick={addEmptyItem} style={{marginTop:8}}>＋ 手動新增項目</button>
        <div className="summary-row">
          <span className="summary-label">總計 (HKD)：</span>
          <span className="summary-value">${grandTotal.toLocaleString('en-US')}</span>
        </div>
      </div>

      {/* Payment + Terms */}
      <div className="card">
        <div className="card-header">付款階段及備註條款</div>
        <details>
          <summary><label><input type="checkbox" checked={showTerms} onChange={e=>setShowTerms(e.target.checked)} style={{marginRight:6}} />備註條款</label></summary>
          <div>{terms.map((t,i)=><div key={i} className="term-row"><span style={{fontWeight:'bold',minWidth:24}}>{i+1}.</span><input value={t} onChange={e=>{const n=[...terms];n[i]=e.target.value;setTerms(n)}} placeholder="條款內容" /><button className="btn btn-danger btn-sm btn-del" onClick={()=>{const n=terms.filter((_,j)=>j!==i);setTerms(n)}}>✕</button></div>)}</div>
          <button className="btn btn-secondary btn-sm" onClick={()=>setTerms([...terms,''])} style={{marginTop:6}}>＋ 新增條款</button>
        </details>
        <details style={{marginTop:8}}>
          <summary><label><input type="checkbox" checked={showPayment} onChange={e=>setShowPayment(e.target.checked)} style={{marginRight:6}} />付款階段</label></summary>
          <div>{payments.map((p,i)=><div key={i} className="pay-row"><input value={p.label||''} onChange={e=>{const n=[...payments];n[i]={...n[i],label:e.target.value};setPayments(n)}} placeholder="期數名" /><input className="pct" type="number" value={p.pct||''} onChange={e=>{const n=[...payments];n[i]={...n[i],pct:parseFloat(e.target.value)||0};setPayments(n)}} placeholder="%" />%<input value={p.label_pct||''} onChange={e=>{const n=[...payments];n[i]={...n[i],label_pct:e.target.value};setPayments(n)}} placeholder="比例說明" /><input value={p.desc||''} onChange={e=>{const n=[...payments];n[i]={...n[i],desc:e.target.value};setPayments(n)}} placeholder="付款條件說明" /><button className="btn btn-danger btn-sm btn-del" onClick={()=>{const n=payments.filter((_,j)=>j!==i);setPayments(n)}}>✕</button></div>)}</div>
          <button className="btn btn-secondary btn-sm" onClick={()=>setPayments([...payments,{label:'',pct:25,label_pct:'',desc:''}])} style={{marginTop:6}}>＋ 新增付款期</button>
        </details>
      </div>

      {/* Generate */}
      <div className="card" style={{textAlign:'center'}}>
        <button className="btn btn-primary btn-lg" onClick={()=>handleGenerate('quotation')}>生成報價單</button>
        <button className="btn btn-primary btn-lg" onClick={()=>handleGenerate('invoice')} style={{marginLeft:8}}>生成發票</button>
        {msg && <div className={`msg ${msg.includes('失敗')?'error':'success'}`}>{msg}</div>}
      </div>

      {/* Modal */}
      {dlg && (
        <div className="modal-overlay show" onClick={()=>setDlg(null)}>
          <div className="modal" onClick={e=>e.stopPropagation()}>
            <h3>{dlg.title}</h3>
            <p>{dlg.filename}</p>
            <a className="btn btn-primary btn-lg" href={dlg.downloadUrl} style={{display:'block'}}>下載 Excel</a>
            <button className="btn btn-secondary btn-lg" onClick={()=>{window.open('https://mini-software.github.io/MiniPdf/','_blank')}} style={{width:'100%',marginTop:8}}>MiniPdf 轉換 PDF</button>
            <button className="btn btn-secondary" onClick={()=>setDlg(null)} style={{width:'100%',marginTop:8}}>關閉</button>
          </div>
        </div>
      )}
    </div>
  )
}
