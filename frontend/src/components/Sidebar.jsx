const VERSION = '3.3.0'

export default function Sidebar({ page, onNavigate }) {
  return (
    <>
      <button className="menu-toggle" onClick={() => document.querySelector('.sidebar').classList.toggle('open')}>☰</button>
      <div className="sidebar">
        <div className="sidebar-logo">工程單助手<span>{VERSION}</span></div>
        <a href="#" className={page === 'form' ? 'active' : ''} onClick={(e) => { e.preventDefault(); onNavigate('form') }}>
          <span className="icon">＋</span>新增報價單
        </a>
        <a href="#" className={page === 'projects' ? 'active' : ''} onClick={(e) => { e.preventDefault(); onNavigate('projects') }}>
          <span className="icon">☰</span>工程單列表
        </a>
        <a href="#" onClick={(e) => { e.preventDefault(); document.getElementById('uploadFile')?.click() }}>
          <span className="icon">↑</span>匯入 Excel
        </a>
        <a href="#" className={page === 'minipdf' ? 'active' : ''} onClick={(e) => { e.preventDefault(); onNavigate('minipdf') }}>
          <span className="icon">⇄</span>MiniPdf 轉換器
        </a>
        <div style={{marginTop:'auto',padding:'8px 14px'}}>
          <button className="theme-toggle" onClick={() => {
            const html = document.documentElement
            if (html.classList.contains('dark')) { html.classList.remove('dark'); localStorage.setItem('theme','light') }
            else { html.classList.add('dark'); localStorage.setItem('theme','dark') }
          }} title="切換模式">🌓</button>
        </div>
      </div>
    </>
  )
}
