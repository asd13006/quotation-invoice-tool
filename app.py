"""
裝修報價單/發票助手 — Flask 網頁後端
"""
import io, re, os, uuid, tempfile, subprocess, base64
from flask import Flask, render_template, request, send_file, jsonify
from openpyxl import Workbook, load_workbook
from generator import generate_quotation
from styles import SECTIONS

app = Flask(__name__)
_preview_cache = {}

_MINIPDF = os.path.expanduser('~/.dotnet/tools/minipdf')
if os.name == 'nt': _MINIPDF += '.exe'

# 跨平台字型路徑
if os.name == 'nt':
    _FONT_DIR = 'C:/Windows/Fonts'
elif os.path.exists('/System/Library/Fonts'):
    _FONT_DIR = '/System/Library/Fonts'
else:
    _FONT_DIR = '/usr/share/fonts'


def _has_minipdf():
    return os.path.exists(_MINIPDF)


def _xlsx_to_pdf(xlsx_bytes):
    """用 MiniPdf 將 xlsx 轉 PDF，返回 bytes"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as xf:
        xf.write(xlsx_bytes)
        xlsx_path = xf.name
    pdf_path = xlsx_path.replace('.xlsx', '.pdf')
    try:
        subprocess.run([_MINIPDF, 'convert', xlsx_path, '-o', pdf_path,
                       '--fonts', _FONT_DIR],
                       capture_output=True, timeout=30, check=True)
        with open(pdf_path, 'rb') as pf:
            return pf.read()
    finally:
        if os.path.exists(xlsx_path): os.unlink(xlsx_path)
        if os.path.exists(pdf_path): os.unlink(pdf_path)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    body = request.get_json(silent=True)
    if not body: return jsonify({'error': '無效的請求資料'}), 400
    data = body.get('data', {})
    if not data.get('items'): return jsonify({'error': '請至少填寫一個工程項目'}), 400

    try:
        gen_type = body.get('type', 'quotation')
        doc_title = '發票' if gen_type == 'invoice' else '報價單'

        wb = Workbook()
        generate_quotation(wb.active, data, title=doc_title)
        xlsx_buf = io.BytesIO()
        wb.save(xlsx_buf)
        xlsx_bytes = xlsx_buf.getvalue()

        # MiniPdf: xlsx -> PDF
        pdf_bytes = b''
        if _has_minipdf():
            pdf_bytes = _xlsx_to_pdf(xlsx_bytes)

        pid = uuid.uuid4().hex[:8]
        pdf_b64 = base64.b64encode(pdf_bytes).decode() if pdf_bytes else ''

        html = _build_preview_html(pid, pdf_b64, doc_title, data)
        fname = _make_filename(data, doc_title)
        _preview_cache[pid] = {'html': html, 'xlsx': xlsx_bytes,
                               '_filename': fname + '.xlsx'}

        return jsonify({'preview_id': pid, 'status': 'ok'})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': f'生成失敗：{str(e)}'}), 500


def _make_filename(data, title):
    """(報價單號_)?(工程地址_)?(客戶姓名_)?(報價日期)_報價單/發票"""
    parts = []
    qn = data.get('quotation_no', '') or ''
    addr = data.get('address', '') or ''
    owner = data.get('owner_name', '') or ''
    date = (data.get('date', '') or '').replace('-', '')
    if qn.strip(): parts.append(qn.strip())
    if addr.strip(): parts.append(addr.strip())
    if owner.strip(): parts.append(owner.strip())
    if date.strip(): parts.append(date.strip())
    if not parts: parts.append(date or 'output')
    return '_'.join(parts) + '_' + title


def _build_preview_html(pid, pdf_b64, title, data):
    addr = data.get('address','') or data.get('project_name','output')
    date_str = (data.get('date','') or '').replace('-','')
    return f'''<!DOCTYPE html><html lang="zh-HK"><head><meta charset="UTF-8"><title>{title}</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{background:#525659;font-family:'Microsoft JhengHei',sans-serif}}
.bar{{position:fixed;top:0;left:0;right:0;background:#323639;padding:8px 16px;display:flex;gap:10px;z-index:99;align-items:center}}
.bar button{{padding:8px 16px;border:none;border-radius:4px;font-size:13px;cursor:pointer;font-weight:bold;color:#fff}}
.btn-excel{{background:#1F4E78}} .btn-pdf{{background:#2E7D32}} .btn-jpg{{background:#E65100}}
.bar span{{color:#aaa;font-size:12px;margin-left:auto}}
iframe{{border:none;width:100%;height:calc(100vh - 44px);margin-top:44px}}
</style></head><body>
<div class="bar">
<button class="btn-excel" onclick="downloadExcel()">下載 Excel</button>
<button class="btn-pdf" onclick="downloadPDF()">下載 PDF</button>
<button class="btn-jpg" onclick="downloadJPG()">下載 JPG</button>
<span>預覽跟 Excel/PDF 一致</span>
</div>
<iframe src="data:application/pdf;base64,{pdf_b64}" id="pdfFrame"></iframe>
<script>
var PID="{pid}";
var ADDR="{addr}";
var TITLE="{title}";
var DATE="{date_str}";
function downloadExcel(){{window.location.href="/download/"+PID+"/excel";}}
function downloadPDF(){{window.location.href="/download/"+PID+"/pdf";}}
function downloadJPG(){{window.location.href="/download/"+PID+"/jpg";}}
</script></body></html>'''


@app.route('/preview/<pid>')
def preview(pid):
    entry = _preview_cache.get(pid)
    return (entry['html'] if entry else '<h1>預覽已過期</h1>')


@app.route('/download/<pid>/excel')
def download_excel(pid):
    entry = _preview_cache.get(pid)
    if not entry or not entry.get('xlsx'): return 'Not found', 404
    fname = entry.get('_filename', '報價單.xlsx')
    return send_file(io.BytesIO(entry['xlsx']),
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=fname)


@app.route('/download/<pid>/pdf')
def download_pdf(pid):
    entry = _preview_cache.get(pid)
    if not entry or not entry.get('xlsx'): return 'Not found', 404
    if not _has_minipdf():
        return '需要安裝 MiniPdf 先可以下載 PDF。請參考 README.md 安裝。', 500
    try:
        pdf_bytes = _xlsx_to_pdf(entry['xlsx'])
        fname = entry.get('_filename', '報價單.xlsx').replace('.xlsx', '.pdf')
        return send_file(io.BytesIO(pdf_bytes), mimetype='application/pdf',
                         as_attachment=True, download_name=fname)
    except Exception as e:
        return f'PDF 轉換失敗：{str(e)}', 500


@app.route('/download/<pid>/jpg')
def download_jpg(pid):
    entry = _preview_cache.get(pid)
    if not entry or not entry.get('xlsx'): return 'Not found', 404
    if not _has_minipdf():
        return '需要安裝 MiniPdf 先可以下載 JPG。請參考 README.md 安裝。', 500
    try:
        pdf_bytes = _xlsx_to_pdf(entry['xlsx'])
        import fitz
        doc = fitz.open(stream=pdf_bytes, filetype='pdf')
        page = doc[0]
        mat = fitz.Matrix(2, 2)
        pix = page.get_pixmap(matrix=mat)
        jpg_bytes = pix.tobytes('jpg')
        doc.close()
        fname = entry.get('_filename', '報價單.xlsx').replace('.xlsx', '.jpg')
        return send_file(io.BytesIO(jpg_bytes), mimetype='image/jpeg',
                         as_attachment=True, download_name=fname)
    except Exception as e:
        return f'JPG 轉換失敗：{str(e)}', 500


@app.route('/upload', methods=['POST'])
def upload():
    file = request.files.get('file')
    if not file: return jsonify({'error':'請選擇 .xlsx 檔案'}),400
    try:
        wb=load_workbook(io.BytesIO(file.read()),data_only=True)
        ws=wb.active
        if ws is None: return jsonify({'error':'無法讀取 Excel'}),400
        result={'items':[],'payments':[],'terms':[],'deposit':0,'_filename':file.filename}
        _parse_header(ws,result)
        if _detect_current_format(ws): _parse_items_current(ws,result)
        else:
            n=_parse_items_universal(ws,result)
            if n==0:
                for sn in wb.sheetnames:
                    if sn!=ws.title:
                        ws2=wb[sn]
                        if _detect_current_format(ws2): _parse_items_current(ws2,result)
                        else: n=_parse_items_universal(ws2,result)
                        if n>0: _parse_header(ws2,result); break
        _parse_payments_and_terms(ws,result)
        _find_deposit(ws,result)
        result.setdefault('show_payment',len(result.get('payments',[]))>0)
        result.setdefault('show_terms',len(result.get('terms',[]))>0)
        return jsonify(result)
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error':f'解析失敗：{str(e)}'}),500

def _parse_header(ws,result):
    hm={'工程名稱：':'project_name','客戶姓名：':'owner_name','工程地址：':'address','裝修公司：':'company_name','報價單號：':'quotation_no','報價日期：':'date','有效期：':'validity','版本：':'version'}
    for r in range(1,11):
        for ca,cb in [('A','B'),('E','F')]:
            l=ws[f'{ca}{r}'].value; v=ws[f'{cb}{r}'].value
            if isinstance(l,str) and l in hm:
                k=hm[l]
                if isinstance(v,str) and v!='-': result[k]=v
                elif isinstance(v,(int,float)): result[k]=str(v)

def _detect_current_format(ws):
    for row in ws.iter_rows(min_row=8,max_row=min(ws.max_row,80)):
        a=str(row[0].value or '')
        if '、 ' in a and len(a)>3: return True
    return False

def _parse_items_current(ws,result):
    sbt={s['title']:s['cat'] for s in SECTIONS}; cat='雜項'
    for row in ws.iter_rows(min_row=8,max_row=ws.max_row):
        if not row or row[0] is None: continue
        r=row[0].row; a=str(ws[f'A{r}'].value or ''); b=str(ws[f'B{r}'].value or '')
        if '、 ' in a:
            st=a.split('、 ',1)[1]
            if st in sbt: cat=sbt[st]
            continue
        if re.match(r'^\d+\.\d+$',a) and b and b!='None': _add_item(ws,r,cat,result); continue
        if a.isdigit() and b and b!='None' and len(b)>2:
            if ws[f'C{r}'].value is not None or ws[f'E{r}'].value is not None: _add_item(ws,r,cat,result)

def _parse_items_universal(ws,result):
    cat='雜項'; sbt={s['title']:s['cat'] for s in SECTIONS}
    before=len(result.get('items',[]))
    cd,cq,cp=_detect_columns(ws)
    sk=['項目','編號','描述','總計','TOTAL','小計','合計','訂金','尾款','備註','地址','日期','名稱','電話','傳真','公司','客戶','No.','Item','Qty','Unit','Price','Amount','Subtotal','工程付款','條款','付款期數','付款條件','Attn','To','Address','Date','INVOICE','RECEIPT','Phone','Fax','TEL','FAX']
    for row in ws.iter_rows(min_row=1,max_row=ws.max_row):
        if not row or row[0] is None: continue
        r=row[0].row; a=str(ws[f'A{r}'].value or '').strip(); b=str(ws[f'B{r}'].value or '').strip()
        if not a.isdigit() and 2<=len(a)<=6:
            if not any(k in a for k in sk):
                for t,c in sbt.items():
                    if a in t or t in a: cat=c; break
        if b and not b[0].isdigit() and 2<=len(b)<=6:
            if not any(k in b for k in sk):
                for t,c in sbt.items():
                    if b in t or t in b: cat=c; break
        dv=str(ws.cell(row=r,column=cd).value or '').strip()
        if not dv or dv=='None' or len(dv)<2:
            at=str(ws[f'A{r}'].value or '').strip()
            if len(at)>3 and not at[0].isdigit(): dv=at
            if cp==5:
                dv2=ws[f'D{r}'].value
                if isinstance(dv2,(int,float)) and float(dv2)>0: cp=4
        if not dv or dv=='None' or len(dv)<2: continue
        if any(k in dv for k in sk): continue
        pv=ws.cell(row=r,column=cp).value
        sq=str(ws[f'A{r}'].value or '').strip()
        is_seq=bool(re.match(r'^\d+[\)\.\s、．]*$',sq))
        is_hier=bool(re.match(r'^\d+\.\d+$',sq))
        has_price=isinstance(pv,(int,float)) and float(pv)>=0
        is_item=(is_seq or is_hier) and len(dv)>=2
        if not is_item and has_price and len(dv)>2: is_item=True
        if is_item and not dv.startswith('='):
            try: qty=int(float(ws.cell(row=r,column=cq).value or 1))
            except: qty=1
            try: price=int(float(pv)) if pv else 0
            except: price=0
            result['items'].append({'category':_auto_categorize(dv) or cat,'description':dv,'quantity':qty,'unit':'項','unit_price':price,'remark':'','is_additional':False})
    return len(result.get('items',[]))-before

def _auto_categorize(d):
    for c,k in {'清拆工程':['清拆','拆','垃圾','棚架','保險'],'水電工程':['水喉','電制','插蘇','煤氣','冷氣','菲士','供電','燈'],'泥水工程':['鋪磚','磁磚','瓷磚','英泥沙','防水','地板','地台','企缸'],'油漆工程':['油漆','批灰','油油','起底','鏟底','ICI','乳膠漆'],'木工工程':['門','櫃','床','天花','腳線','廚櫃','衣櫃','雲石'],'安裝代工':['安裝','代裝','人工']}.items():
        for kw in k:
            if kw in d: return c
    return '雜項'

def _detect_columns(ws):
    cd,cq,cp=2,3,5; nc,tc={},{}
    for row in ws.iter_rows(min_row=1,max_row=min(ws.max_row,150)):
        r=row[0].row
        for c in range(1,min(ws.max_column+1,12)):
            v=ws.cell(row=r,column=c).value
            if isinstance(v,(int,float)) and float(v)>0: nc[c]=nc.get(c,0)+1
            if isinstance(v,str) and len(v)>3: tc[c]=tc.get(c,0)+1
    for c,_ in sorted(nc.items(),key=lambda x:-x[0]):
        if nc.get(c,0)>=2: cp=c; break
    if tc:
        al={}
        for c in tc:
            ls=[]
            for row in ws.iter_rows(min_row=1,max_row=min(ws.max_row,30)):
                v=str(ws.cell(row=row[0].row,column=c).value or '')
                if len(v)>1: ls.append(len(v))
            al[c]=sum(ls)/max(len(ls),1)
        for c,_ in sorted(tc.items(),key=lambda x:(-al.get(x[0],0),-x[1])):
            if c<cp and c!=1 and tc.get(c,0)>=2 and al.get(c,0)>5: cd=c; break
        if cd==2 and not tc.get(2,0)>=2:
            for c,_ in sorted(tc.items(),key=lambda x:-x[1]):
                if tc.get(c,0)>=2 and c!=1: cd=c; break
    for c in range(cd+1,cp):
        if nc.get(c,0)>=2: cq=c; break
    return cd,cq,cp

def _add_item(ws,r,cat,result):
    b=str(ws[f'B{r}'].value or '')
    try: qty=int(float(str(ws[f'C{r}'].value or 1)))
    except: qty=1
    try: price=int(float(str(ws[f'E{r}'].value or 0)))
    except: price=0
    rk=str(ws[f'G{r}'].value or '')
    if rk in ('None','-',''): rk=''
    result['items'].append({'category':cat,'description':b,'quantity':qty,'unit':str(ws[f'D{r}'].value or '項'),'unit_price':price,'remark':rk,'is_additional':False})

def _find_deposit(ws,result):
    for row in ws.iter_rows(min_row=1,max_row=ws.max_row):
        for cell in row:
            if cell.value and '訂金' in str(cell.value):
                fv=ws[f'F{cell.row}'].value
                if fv and isinstance(fv,(int,float)): result['deposit']=int(fv); return

def _parse_payments_and_terms(ws,result):
    for row in ws.iter_rows(min_row=1,max_row=ws.max_row):
        a=str(ws[f'A{row[0].row}'].value or '')
        if '工程付款階段說明' in a or '付款階段' in a:
            pr=row[0].row+1; cnt=0
            while pr<=ws.max_row and cnt<10:
                pa=str(ws[f'A{pr}'].value or ''); pb=str(ws[f'B{pr}'].value or ''); pd=str(ws[f'D{pr}'].value or '')
                if pa not in ('付款期數','None','備註','條款','') and len(pa)>1:
                    pm=re.search(r'(\d+)%',pb)
                    result['payments'].append({'label':pa,'pct':int(pm.group(1)) if pm else 25,'label_pct':pb if pb not in ('None','') else '','desc':pd if pd not in ('None','') else ''}); cnt+=1
                pr+=1
        if '備註及條款說明' in a or '備註及條款' in a:
            tr=row[0].row+1; cnt=0
            while tr<=ws.max_row and cnt<20:
                ta=str(ws[f'A{tr}'].value or '')
                if ta and ta!='None' and '付款' not in ta:
                    ta=re.sub(r'^\d+\.\s*','',ta).strip()
                    if len(ta)>3: result['terms'].append(ta); cnt+=1
                tr+=1

if __name__=='__main__':
    print('='*50); print('裝修報價單/發票助手已啟動'); print('http://localhost:5000'); print('='*50)
    app.run(debug=True,host='127.0.0.1',port=5000)
