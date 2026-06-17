"""
裝修報價單/發票助手 — Flask 後端（Vercel 相容）
"""
import io, os, re, uuid
from flask import Flask, render_template, request, send_file, jsonify
from openpyxl import Workbook, load_workbook
from generator import generate_quotation
from drive_sync import list_projects, get_project_data

app = Flask(__name__)
_preview_cache = {}

with open(os.path.join(os.path.dirname(__file__), 'VERSION')) as f:
    _VERSION = f.read().strip()


@app.route('/')
def index():
    return render_template('index.html', version=_VERSION)


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
        pid = uuid.uuid4().hex[:8]
        fname = _make_filename(data, doc_title)
        html = _build_download_page(pid, doc_title, data, fname)
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


@app.route('/upload', methods=['POST'])
def upload():
    file = request.files.get('file')
    if not file: return jsonify({'error': '請選擇 .xlsx 檔案'}), 400
    try:
        wb = load_workbook(io.BytesIO(file.read()), data_only=True)
        ws = wb.active
        result = {'items': [], 'payments': [], 'terms': [], 'deposit': 0,
                  '_filename': file.filename, 'show_payment': False, 'show_terms': False}

        # Parse header
        header_map = {'工程名稱': 'project_name', '工程地址': 'address',
                      '客戶姓名': 'owner_name', '業主姓名': 'owner_name',
                      '裝修公司': 'company_name', '報價單號': 'quotation_no',
                      '報價日期': 'date', '有效期': 'validity', '版本': 'version'}
        for r in range(1, 11):
            for c in range(1, 8):
                label = str(ws.cell(row=r, column=c).value or '')
                for kw, key in header_map.items():
                    if kw in label and '：' in label:
                        val = str(ws.cell(row=r, column=c+1).value or '')
                        if val and val != 'None' and val != '-':
                            result[key] = val

        # Scan for item rows: seq number + description + price
        items = []
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            if not row or row[0] is None: continue
            r = row[0].row
            a = str(ws.cell(row=r, column=1).value or '').strip()
            b = str(ws.cell(row=r, column=2).value or '').strip()
            # Look for sequence number: "1", "1)", "1.1", "2.3" etc
            is_seq = bool(re.match(r'^\d+[\.\)]?\s*$', a) or re.match(r'^\d+\.\d+$', a))
            if is_seq and len(b) > 2:
                # Unit price is in column E (5th column)
                e = ws.cell(row=r, column=5).value
                price = int(e) if isinstance(e, (int, float)) else 0
                # qty from column C (3rd column)
                c = ws.cell(row=r, column=3).value
                qty = int(c) if isinstance(c, (int, float)) and float(c) > 0 else 1
                items.append({'category': '雜項', 'description': b, 'quantity': qty,
                              'unit': '項', 'unit_price': price, 'remark': '',
                              'is_additional': False})
        result['items'] = items

        # Parse payments & terms
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            if not row or row[0] is None: continue
            a = str(ws.cell(row=row[0].row, column=1).value or '')
            r = row[0].row

            # Payment section: look for header "付款期數" then read following rows
            if '付款期數' in a:
                pr = r + 1
                while pr <= ws.max_row:
                    pa = str(ws.cell(row=pr, column=1).value or '')
                    pb = str(ws.cell(row=pr, column=2).value or '')
                    pd = str(ws.cell(row=pr, column=4).value or '')
                    if pa and pa != 'None' and '備註' not in pa and '條款' not in pa and '付款期數' not in pa:
                        pct_m = re.search(r'(\d+)%', pb)
                        pct = int(pct_m.group(1)) if pct_m else 25
                        result['payments'].append({'label': pa, 'pct': pct,
                            'label_pct': pb if pb not in ('None','') else f'總金額之 {pct}%',
                            'desc': pd if pd not in ('None','') else ''})
                        pr += 1
                    else:
                        break
                result['show_payment'] = len(result['payments']) > 0

            # Terms section
            if '備註及條款' in a:
                tr = r + 1
                while tr <= ws.max_row:
                    ta = str(ws.cell(row=tr, column=1).value or '')
                    if ta and ta != 'None' and '付款' not in ta:
                        ta = re.sub(r'^\d+\.\s*', '', ta).strip()
                        if len(ta) > 3:
                            result['terms'].append(ta)
                        tr += 1
                    else:
                        tr += 1
                        if tr - r > 20: break
                result['show_terms'] = len(result['terms']) > 0

        return jsonify(result)
    except Exception as e:
        return jsonify({'error': f'解析失敗：{str(e)}'}), 500


@app.route('/projects')
def projects_page():
    """工程單管理列表頁"""
    try:
        projects = list_projects()
    except Exception as e:
        return f'<h1>無法連接 Google Drive</h1><p>{e}</p><a href="/">返回</a>', 500
    return render_template('projects.html', projects=projects, version=_VERSION)


@app.route('/api/projects/<file_id>')
def api_load_project(file_id):
    """載入指定工程單 data（供前端 AJAX call）"""
    try:
        data = get_project_data(file_id)
        data['_filename'] = file_id
        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def _make_filename(data, title):
    parts = []
    for key in ['quotation_no', 'address', 'owner_name']:
        v = (data.get(key, '') or '').strip()
        if v: parts.append(v)
    date = (data.get('date', '') or '').replace('-', '')
    if date.strip(): parts.append(date.strip())
    if not parts: parts.append(date or '-')
    return '_'.join(parts) + '_' + title


def _build_download_page(pid, title, data, fname):
    addr = data.get('address','') or data.get('project_name','output')
    date_str = (data.get('date','') or '').replace('-','')
    return f"""<!DOCTYPE html><html lang="zh-HK"><head><meta charset="UTF-8"><title>{title}</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{background:#f0f2f5;font-family:'Microsoft JhengHei',sans-serif;display:flex;justify-content:center;align-items:center;min-height:100vh}}
.card{{background:#fff;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,.1);padding:40px;text-align:center;max-width:420px}}
h1{{font-size:22px;color:#002F5597;margin-bottom:10px}}
p{{color:#666;margin-bottom:24px;font-size:14px}}
.btn{{display:block;width:100%;padding:14px;border:none;border-radius:8px;font-size:16px;cursor:pointer;font-weight:bold;color:#fff;margin-bottom:10px;text-decoration:none}}
.btn-excel{{background:#1F4E78}}
.btn-back{{background:#888;font-size:13px;padding:10px}}
.ver{{color:#aaa;font-size:12px;margin-top:16px}}
</style></head><body>
<div class="card">
<h1>{title}已生成</h1>
<p>{fname}.xlsx</p>
<a class="btn btn-excel" href="/download/{pid}/excel">下載 Excel</a>
<a class="btn btn-back" href="/">返回主頁</a>
<div class="ver">{_VERSION}</div>
</div>
</body></html>"""


if __name__ == '__main__':
    print('=' * 50)
    print(f'裝修報價單/發票助手 {_VERSION}')
    print('http://localhost:5000')
    print('=' * 50)
    app.run(debug=True, host='127.0.0.1', port=5000)
