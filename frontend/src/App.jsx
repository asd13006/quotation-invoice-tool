import { useState, useEffect } from 'react'
import Sidebar from './components/Sidebar.jsx'
import QuotationForm from './components/QuotationForm.jsx'
import ProjectList from './components/ProjectList.jsx'
import MiniPdfPage from './components/MiniPdfPage.jsx'

export default function App() {
  const [page, setPage] = useState('form')
  const [loadData, setLoadData] = useState(null)

  // Theme init
  useEffect(() => {
    const saved = localStorage.getItem('theme')
    if (saved === 'dark' || (!saved && window.matchMedia('(prefers-color-scheme:dark)').matches)) {
      document.documentElement.classList.add('dark')
    }
  }, [])

  const handleNavigate = (p) => {
    setPage(p)
    if (p !== 'form') setLoadData(null)
  }

  const handleLoadProject = (data) => {
    setLoadData(data)
    setPage('form')
  }

  return (
    <>
      <Sidebar page={page} onNavigate={handleNavigate} />
      <div className="main" style={{marginLeft:210,flex:1,padding:'28px 32px',minWidth:0}}>
        {page === 'form' && <QuotationForm loadData={loadData} />}
        {page === 'projects' && <ProjectList onLoadProject={handleLoadProject} />}
        {page === 'minipdf' && <MiniPdfPage />}
      </div>
    </>
  )
}
