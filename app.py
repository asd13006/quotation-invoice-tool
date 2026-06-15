"""
裝修報價單/發票助手 — Flask 後端（Vercel 相容，純 Python）
"""
import io, re, uuid, base64
from flask import Flask, render_template, request, send_file, jsonify
from openpyxl import Workbook, load_workbook
from generator import generate_quotation
from styles import SECTIONS
from xlsx2pdf import convert_xlsx_to_pdf

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

        # Excel
        wb = Workbook()
        generate_quotation(wb.active, data, title=doc_title)
        xlsx_buf = io.BytesIO()
        wb.save(xlsx_buf)
        xlsx_bytes = xlsx_buf.getvalue()

        # PDF (pure Python xlsx -> PDF)
        pdf_bytes = convert_xlsx_to_pdf(xlsx_bytes)
        pdf_b64 = base64.b64encode(pdf_bytes).decode()

        pid = uuid.uuid4().hex[:8]
        fname = _make_filename(data, doc_title)
        html = _build_preview_html(pid, pdf_b64, doc_title, data, xlsx_bytes)

        _preview_cache[pid] = {'html': html, 'xlsx': xlsx_bytes, '_filename': fname + '.xlsx'}

        return jsonify({'preview_id': pid, 'status': 'ok'})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': f'生成失敗：{str(e)}'}), 500


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
    try:
        pdf_bytes = convert_xlsx_to_pdf(entry['xlsx'])
        fname = entry.get('_filename', '報價單.xlsx').replace('.xlsx', '.pdf')
        return send_file(io.BytesIO(pdf_bytes), mimetype='application/pdf',
                         as_attachment=True, download_name=fname)
    except Exception as e:
        return f'PDF 轉換失敗：{str(e)}', 500


# ── Helpers ──

def _make_filename(data, title):
    """(報價單號_)?(工程地址_)?(客戶姓名_)?(報價日期)_報價單/發票"""
    parts = []
    for key in ['quotation_no', 'address', 'owner_name']:
        v = (data.get(key, '') or '').strip()
        if v: parts.append(v)
    date = (data.get('date', '') or '').replace('-', '')
    if date.strip(): parts.append(date.strip())
    if not parts: parts.append(date or 'output')
    return '_'.join(parts) + '_' + title


def _build_preview_html(pid, pdf_b64, title, data, xlsx_bytes):
    addr = data.get('address','') or data.get('project_name','output')
    date_str = (data.get('date','') or '').replace('-','')
    capture_html = _build_capture_html(xlsx_bytes)

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
<span>v1.1.0 — 純 Python PDF，Vercel 相容</span>
</div>
<div style="display:flex;align-items:center;justify-content:center;height:calc(100vh - 44px);margin-top:44px;flex-direction:column;gap:16px">
<div style="color:#fff;font-size:18px">PDF 已生成，請下載查看（瀏覽器內嵌 PDF 可能顯示不完整）</div>
<div style="display:flex;gap:12px">
<a href="/download/{pid}/excel" style="background:#1F4E78;color:#fff;padding:12px 24px;border-radius:6px;text-decoration:none;font-weight:bold">下載 Excel</a>
<a href="/download/{pid}/pdf" style="background:#2E7D32;color:#fff;padding:12px 24px;border-radius:6px;text-decoration:none;font-weight:bold">下載 PDF</a>
<button onclick="downloadJPG()" style="background:#E65100;color:#fff;padding:12px 24px;border:none;border-radius:6px;font-weight:bold;cursor:pointer;font-size:14px">下載 JPG</button>
</div>
<iframe src="data:application/pdf;base64,{pdf_b64}" id="pdfFrame" style="width:100%;flex:1;border:none"></iframe>
</div>
<div id="capture" style="position:absolute;left:-9999px;top:0;width:190mm;background:#fff;padding:8mm">{capture_html}</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<script>
var PID="{pid}";var ADDR="{addr}";var TITLE="{title}";var DATE="{date_str}";
function downloadExcel(){{window.location.href="/download/"+PID+"/excel";}}
function downloadPDF(){{window.location.href="/download/"+PID+"/pdf";}}
async function downloadJPG(){{var el=document.getElementById("capture");var canvas=await html2canvas(el,{{scale:2,backgroundColor:"#ffffff"}});canvas.toBlob(function(blob){{var url=URL.createObjectURL(blob);var link=document.createElement("a");link.href=url;link.download=ADDR+"_"+TITLE+"_"+DATE+".jpg";link.click();URL.revokeObjectURL(url);}},"image/jpeg",0.92);}}
</script></body></html>'''


def _build_capture_html(xlsx_bytes):
    """由 xlsx 生成簡單 HTML table 畀 html2canvas 截圖"""
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb.active
    h = '<table style="border-collapse:collapse;font-size:9pt;width:100%;font-family:Microsoft JhengHei,sans-serif">'
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        h += '<tr>'
        for cell in row:
            val = str(cell.value or '').replace('&','&amp;').replace('<','&lt;')
            if val.startswith('='): val = ''
            style = 'border:1px solid #d9d9d9;padding:2px 4px;'
            try:
                if cell.fill.patternType == 'solid':
                    rgb = cell.fill.fgColor.rgb
                    if rgb and rgb not in ('00000000','0'): style += 'background:#' + rgb[2:] + ';'
            except: pass
            try:
                if cell.font and cell.font.bold: style += 'font-weight:bold;'
            except: pass
            ha = 'left'
            try:
                if cell.alignment.horizontal: ha = cell.alignment.horizontal
            except: pass
            style += 'text-align:' + ha + ';'
            h += f'<td style="{style}">{val}</td>'
        h += '</tr>'
    h += '</table>'
    return h


if __name__ == '__main__':
    print('=' * 50)
    print('裝修報價單/發票助手 v1.1.0（Vercel 相容）')
    print('http://localhost:5000')
    print('=' * 50)
    app.run(debug=True, host='127.0.0.1', port=5000)
