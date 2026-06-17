"""
裝修報價單/發票助手 — Flask 後端（Vercel 相容）
"""
import io, os, re, uuid
from flask import Flask, render_template, request, send_file, jsonify
from openpyxl import Workbook, load_workbook
from generator import generate_quotation

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
        items = []
        # Scan for item rows: seq number + description + price
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            if not row or row[0] is None: continue
            r = row[0].row
            a = str(ws.cell(row=r, column=1).value or '').strip()
            b = str(ws.cell(row=r, column=2).value or '').strip()
            # Look for sequence number pattern
            if re.match(r'^\d+[\.\)]?\s*$', a) and len(b) > 2:
                # Unit price is in column E (5th column)
                e = ws.cell(row=r, column=5).value
                price = int(e) if isinstance(e, (int, float)) else 0
                # qty from column C (3rd column)
                c = ws.cell(row=r, column=3).value
                qty = int(c) if isinstance(c, (int, float)) and float(c) > 0 else 1
                items.append({'category': '雜項', 'description': b, 'quantity': qty,
                              'unit': '項', 'unit_price': price, 'remark': '',
                              'is_additional': False})
        return jsonify({'items': items, 'payments': [], 'terms': [], 'deposit': 0,
                        '_filename': file.filename, 'show_payment': False, 'show_terms': False})
    except Exception as e:
        return jsonify({'error': f'解析失敗：{str(e)}'}), 500


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
