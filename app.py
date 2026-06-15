"""
裝修報價單/發票助手 — Vercel 兼容
"""
import io, re, uuid, base64
from flask import Flask, render_template, request, send_file, jsonify
from openpyxl import Workbook, load_workbook
from generator import generate_quotation
from styles import SECTIONS

app = Flask(__name__)
_preview_cache = {}


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
        xlsx_b64 = base64.b64encode(xlsx_buf.getvalue()).decode()

        pid = uuid.uuid4().hex[:8]
        fname = _make_filename(data, doc_title)
        html = _build_preview(pid, xlsx_b64, doc_title, data, fname)
        _preview_cache[pid] = {'html': html, 'xlsx': xlsx_buf.getvalue(), '_filename': fname + '.xlsx'}

        return jsonify({'preview_id': pid, 'status': 'ok'})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': f'生成失敗：{str(e)}'}), 500


def _make_filename(data, title):
    parts = []
    for key in ['quotation_no', 'address', 'owner_name']:
        v = (data.get(key, '') or '').strip()
        if v: parts.append(v)
    date = (data.get('date', '') or '').replace('-', '')
    if date: parts.append(date)
    if not parts: parts.append('output')
    return '_'.join(parts) + '_' + title


def _build_preview(pid, xlsx_b64, title, data, fname):
    from styles import SECTIONS as SEC
    CN = ['一','二','三','四','五','六','七']

    def esc(s):
        if s is None: return '-'
        return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

    def fmt(n):
        try: return '${:,}'.format(int(n))
        except: return '$0'

    # Items
    si = {s['num']: [] for s in SEC}
    for it in data.get('items', []):
        cat = it.get('category', '')
        for s in SEC:
            if cat == s['cat']: si[s['num']].append(it); break
        else: si[7].append(it)

    rows = ''; sc = 0; gt = 0
    for s in SEC:
        items = si[s['num']]
        if not items: continue
        sc += 1
        rows += f'<tr class="sec"><td colspan="7">{CN[sc-1]}、 {esc(s["title"])}</td></tr>'
        st = 0
        for i, it in enumerate(items):
            seq = f'{sc}.{i+1}'
            q = it.get('quantity',1) or 1; p = it.get('unit_price',0) or 0; a = q*p
            st += a
            rows += f'<tr><td class="tc">{seq}</td><td>{esc(it["description"])}</td><td class="tc">{q}</td><td class="tc">{esc(it.get("unit","項"))}</td><td class="tr">{fmt(p)}</td><td class="tr">{fmt(a)}</td><td>{esc(it.get("remark","") or "-")}</td></tr>'
        gt += st
        rows += f'<tr class="sub"><td colspan="4"></td><td class="tr">小計：</td><td class="tr">{fmt(st)}</td><td></td></tr>'

    # Deposit
    dep = ''
    deposit = data.get('deposit',0)
    if title == '發票' and deposit > 0:
        bal = gt - deposit
        dep = f'<tr class="tot"><td colspan="5">訂金 (Deposit)：</td><td class="tr">{fmt(deposit)}</td><td></td></tr><tr class="tot"><td colspan="5">應付尾款 (Balance Due)：</td><td class="tr">{fmt(bal)}</td><td></td></tr>'

    # Payment
    pay = ''
    if data.get('show_payment',True):
        payments = data.get('payments',[])
        if payments:
            pay = '<tr class="pt"><td colspan="4">工程付款階段說明：</td></tr><tr class="th2"><td>付款期數</td><td>比例</td><td class="tr">金額 (HKD)</td><td colspan="4">付款條件說明</td></tr>'
            for p in payments:
                pct = p.get('pct',0)
                pay += f'<tr><td>{esc(p.get("label",""))}</td><td class="tc">{esc(p.get("label_pct",""))}</td><td class="tr">{fmt(int(gt*pct))}</td><td colspan="4">{esc(p.get("desc",""))}</td></tr>'

    # Terms
    terms = ''
    if data.get('show_terms',True):
        ts = data.get('terms',[])
        if ts:
            terms = '<tr class="pt"><td colspan="7">備註及條款說明：</td></tr>'
            for i, t in enumerate(ts, 1):
                terms += f'<tr class="tm"><td colspan="7">{i}. {esc(t)}</td></tr>'

    return f'''<!DOCTYPE html><html lang="zh-HK"><head><meta charset="UTF-8"><title>{esc(title)}</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Microsoft JhengHei','PMingLiu',sans-serif;color:#1a1a1a;background:#e8e8e8;display:flex;justify-content:center;padding:12px}}
.page{{width:190mm;background:#fff;padding:8mm 10mm;box-shadow:0 2px 8px rgba(0,0,0,.1)}}
h1{{text-align:center;font-size:20pt;font-weight:bold;padding-bottom:3px;border-bottom:1px solid #ccc;margin-bottom:3mm}}
.info td{{border:none;font-size:9pt;padding:2px 4px;line-height:1.6}}
.il{{font-weight:bold;border:none}}
table{{width:100%;border-collapse:collapse;font-size:9pt;table-layout:fixed;margin:2mm 0}}
th,td{{padding:2px 4px;border:1px solid #e0e0e0;vertical-align:middle}}
th{{font-weight:bold;text-align:center;border-bottom:2px solid #999}}
.tc{{text-align:center}}.tr{{text-align:right}}
.sec td{{background:#F2F2F2;font-weight:bold;font-size:9.5pt;border:none;padding:3px 4px}}
.sub td{{background:#F2F2F2;font-weight:bold;border:none}}
.tot td{{background:#F2F2F2;font-weight:bold;font-size:9.5pt;border:none;padding:3px 4px}}
.pt td{{font-weight:bold;border:none;padding-top:6px}}
.tm td{{border:none;font-size:8.5pt;color:#555;padding:1px 4px}}
.th2 td{{font-weight:bold;border:1px solid #e0e0e0;text-align:center}}
.bar{{position:fixed;top:8px;right:8px;display:flex;gap:6px;z-index:99}}
.bar button{{padding:8px 14px;border:none;border-radius:4px;font-size:13px;cursor:pointer;font-weight:bold;color:#fff}}
.b1{{background:#1F4E78}}.b2{{background:#2E7D32}}.b3{{background:#E65100}}
@media print{{@page{{size:A4;margin:10mm}}body{{background:#fff;padding:0}}.page{{box-shadow:none;margin:0;padding:5mm 8mm;max-width:none;width:100%}}.bar{{display:none}}}}
</style></head><body>
<div class="bar"><button class="b1" onclick="d(1)">Excel</button><button class="b2" onclick="window.print()">PDF</button><button class="b3" onclick="d(2)">JPG</button></div>
<div class="page" id="capture">
<h1>{esc(title)}</h1>
<table class="info"><colgroup><col style="width:15%"><col style="width:35%"><col style="width:15%"><col style="width:35%"></colgroup>
<tr><td class="il">工程名稱：</td><td>{esc(data.get("project_name","-"))}</td><td class="il">報價單號：</td><td>{esc(data.get("quotation_no","-"))}</td></tr>
<tr><td class="il">客戶姓名：</td><td>{esc(data.get("owner_name","-"))}</td><td class="il">報價日期：</td><td>{esc(data.get("date","-"))}</td></tr>
<tr><td class="il">工程地址：</td><td>{esc(data.get("address","-"))}</td><td class="il">有效期：</td><td>{esc(data.get("validity","-"))}</td></tr>
<tr><td class="il">裝修公司：</td><td>{esc(data.get("company_name","-"))}</td><td class="il">版本：</td><td>{esc(data.get("version","-"))}</td></tr>
</table>
<table><colgroup><col style="width:10%"><col style="width:34%"><col style="width:6%"><col style="width:5%"><col style="width:10%"><col style="width:11%"><col style="width:24%"></colgroup>
<thead><tr><th>項目編號</th><th style="text-align:left">工程項目及說明</th><th>數量</th><th>單位</th><th style="text-align:right">單價(HKD)</th><th style="text-align:right">複價(HKD)</th><th style="text-align:left">備註</th></tr></thead>
<tbody>{rows}</tbody></table>
<div style="display:flex;justify-content:flex-end;padding:4px 8px;font-weight:bold;font-size:10pt;background:#F2F2F2;margin-top:1mm">總工程預算總計 (HKD)：{fmt(gt)}</div>
{dep}{pay}{terms}
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<script>
var X="{xlsx_b64}";var N="{fname}";
function d(t){{if(t===1){{var b=atob(X);var a=new Uint8Array(b.length);for(var i=0;i<b.length;i++)a[i]=b.charCodeAt(i);var bl=new Blob([a],{{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}});var u=URL.createObjectURL(bl);var l=document.createElement("a");l.href=u;l.download=N+".xlsx";l.click()}}else{{var el=document.getElementById("capture");html2canvas(el,{{scale:2,backgroundColor:"#ffffff"}}).then(function(c){{c.toBlob(function(b){{var u=URL.createObjectURL(b);var l=document.createElement("a");l.href=u;l.download=N+".jpg";l.click()}},"image/jpeg",0.92)}})}}}}
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


# ── Upload (unchanged) ──

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
