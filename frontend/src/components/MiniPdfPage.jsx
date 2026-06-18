export default function MiniPdfPage() {
  return (
    <div>
      <h1 className="page-title">MiniPdf 轉換器</h1>
      <p style={{color:'var(--text-secondary)',fontSize:13,marginBottom:14}}>上傳 Excel → 完美 PDF 轉換（由 MiniPdf 引擎提供）</p>
      <iframe src="https://mini-software.github.io/MiniPdf/" style={{width:'100%',minHeight:'75vh',border:'none',borderRadius:'var(--radius)',background:'#fff'}} />
    </div>
  )
}
