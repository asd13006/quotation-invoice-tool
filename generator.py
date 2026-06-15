"""
裝修報價單 Excel 生成器 — 跟目標格式
"""
from collections import OrderedDict
from openpyxl.utils import get_column_letter
from styles import *

CN_NUMS = ['一', '二', '三', '四', '五', '六', '七']


def generate_quotation(ws, data, title='報價單'):
    ws.title = '裝修發票' if title == '發票' else SHEET_NAME
    set_col_widths(ws, COL_WIDTHS)

    # 按 category 歸入 section
    section_items = {sec['num']: [] for sec in SECTIONS}
    for it in data.get('items', []):
        cat = it.get('category', '')
        for sec in SECTIONS:
            if cat == sec['cat']:
                section_items[sec['num']].append(it)
                break
        else:
            section_items[7].append(it)

    r = 1

    # ═══ Row 1: 標題 ═══
    title_border = Border(bottom=Side(style='thin'))
    # 先对所有格加底線，再合併
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
        apply_cell(ws, f'{col}{r}', border=title_border)
    ws.merge_cells(f'A{r}:G{r}')
    apply_cell(ws, f'A{r}', value=title, font=FONT_TITLE,
               alignment=Alignment(horizontal='center', vertical='center'))
    ws.row_dimensions[r].height = 40.0
    r = 3  # skip row 2

    # ═══ Rows 3-6: 資訊 ═══
    info_left = [
        ('工程名稱：', 'project_name'),
        ('客戶姓名：', 'owner_name'),
        ('工程地址：', 'address'),
        ('裝修公司：', 'company_name'),
    ]
    info_right = [
        ('報價單號：', 'quotation_no'),
        ('報價日期：', 'date'),
        ('有效期：', 'validity'),
        ('版本：', 'version'),
    ]
    for i in range(4):
        row = r + i
        apply_cell(ws, f'A{row}', value=info_left[i][0], font=FONT_INFO_LABEL)
        apply_cell(ws, f'B{row}', value=data.get(info_left[i][1], '') or '-', font=FONT_INFO_VALUE)
        ws.merge_cells(f'B{row}:D{row}')
        apply_cell(ws, f'E{row}', value=info_right[i][0], font=FONT_INFO_LABEL)
        apply_cell(ws, f'F{row}', value=data.get(info_right[i][1], '') or '-', font=FONT_INFO_VALUE)
        ws.merge_cells(f'F{row}:G{row}')
        ws.row_dimensions[row].height = ROW_HEIGHT_INFO
    r += 4

    # ═══ Row 8: 全頁唯一表頭 ═══
    r += 1  # spacer row 7
    _write_table_header(ws, r)
    r += 1

    # ═══ Sections ═══
    subtotal_cells = []
    section_counter = 0

    for sec in SECTIONS:
        items = section_items[sec['num']]
        if not items:
            continue

        section_title = f'{CN_NUMS[section_counter]}、 {sec["title"]}'
        section_counter += 1

        # Section header (merged A:G, 冇框, 淺灰底)
        ws.merge_cells(f'A{r}:G{r}')
        apply_cell(ws, f'A{r}', value=section_title, font=FONT_SECTION,
                   fill=FILL_SECTION, alignment=ALIGN_LEFT)
        ws.row_dimensions[r].height = 20.5
        r += 1

        # Items
        item_start = r
        for i, it in enumerate(items):
            seq = f'{section_counter}.{i + 1}'
            apply_cell(ws, f'A{r}', value=seq, font=FONT_ITEM, alignment=ALIGN_CENTER, border=BORDER_GRAY)
            apply_cell(ws, f'B{r}', value=it['description'], font=FONT_ITEM, alignment=ALIGN_LEFT, border=BORDER_GRAY)
            apply_cell(ws, f'C{r}', value=it.get('quantity', 1), font=FONT_ITEM, alignment=ALIGN_CENTER, border=BORDER_GRAY)
            apply_cell(ws, f'D{r}', value=it.get('unit', '項'), font=FONT_ITEM, alignment=ALIGN_CENTER, border=BORDER_GRAY)
            apply_cell(ws, f'E{r}', value=it.get('unit_price', 0), font=FONT_ITEM, alignment=ALIGN_RIGHT, border=BORDER_GRAY, number_format=FMT_CURRENCY)
            apply_cell(ws, f'F{r}', value=f'=C{r}*E{r}', font=FONT_ITEM, alignment=ALIGN_RIGHT, border=BORDER_GRAY, number_format=FMT_CURRENCY)
            apply_cell(ws, f'G{r}', value=it.get('remark', '') or '-', font=FONT_ITEM, alignment=ALIGN_LEFT, border=BORDER_GRAY)
            ws.row_dimensions[r].height = ROW_HEIGHT_ITEM
            r += 1
        item_end = r - 1

        # Subtotal: "小計：" in E, formula in F — thin 全包
        apply_cell(ws, f'E{r}', value='小計：', font=FONT_SUBTOTAL,
                   fill=FILL_SECTION, alignment=ALIGN_RIGHT)
        apply_cell(ws, f'F{r}', value=f'=SUM(F{item_start}:F{item_end})',
                   font=FONT_SUBTOTAL, fill=FILL_SECTION, alignment=ALIGN_RIGHT,
                   number_format=FMT_CURRENCY)
        ws.row_dimensions[r].height = ROW_HEIGHT_ITEM
        subtotal_cells.append(f'F{r}')
        r += 2  # spacer

    # ═══ 總計 ═══
    ws.merge_cells(f'A{r}:E{r}')
    apply_cell(ws, f'A{r}', value='總工程預算總計 (HKD)：',
               font=FONT_TOTAL, fill=FILL_SECTION, alignment=ALIGN_RIGHT)
    apply_cell(ws, f'F{r}', value='=' + '+'.join(subtotal_cells),
               font=FONT_TOTAL, fill=FILL_SECTION, alignment=ALIGN_RIGHT, number_format=FMT_CURRENCY)
    ws.row_dimensions[r].height = ROW_HEIGHT_TOTAL
    grand_total_row = r
    r += 1

    # 訂金 + 尾款（只限發票，且訂金 > 0）
    deposit = data.get('deposit', 0)
    if title == '發票' and deposit > 0:
        apply_cell(ws, f'A{r}', value='訂金 (Deposit)：', font=FONT_DEPOSIT_LABEL, alignment=ALIGN_RIGHT)
        ws.merge_cells(f'A{r}:E{r}')
        apply_cell(ws, f'F{r}', value=deposit, font=FONT_DEPOSIT_VALUE,
                   alignment=ALIGN_RIGHT, number_format=FMT_CURRENCY)
        ws.row_dimensions[r].height = ROW_HEIGHT_ITEM
        deposit_row = r
        r += 1

        apply_cell(ws, f'A{r}', value='應付尾款 (Balance Due)：', font=FONT_DEPOSIT_LABEL, alignment=ALIGN_RIGHT)
        ws.merge_cells(f'A{r}:E{r}')
        apply_cell(ws, f'F{r}', value=f'=F{grand_total_row}-F{deposit_row}',
                   font=FONT_BALANCE, alignment=ALIGN_RIGHT, number_format=FMT_CURRENCY)
        ws.row_dimensions[r].height = ROW_HEIGHT_ITEM
        r += 1

    r += 1

    # ═══ 付款 ═══
    show_payment = data.get('show_payment', True)
    show_terms = data.get('show_terms', True)

    if show_payment:
        apply_cell(ws, f'A{r}', value='工程付款階段說明：', font=FONT_TERMS_LABEL)
        r += 1

        # Payment header
        for col, hdr in [('A', '付款期數'), ('B', '比例'), ('C', '金額 (HKD)'), ('D', '付款條件說明')]:
            apply_cell(ws, f'{col}{r}', value=hdr, font=FONT_PAYMENT_HEADER, border=BORDER_GRAY)
        ws.merge_cells(f'D{r}:G{r}')
        ws.row_dimensions[r].height = ROW_HEIGHT_PAYMENT
        r += 1

        payments = data.get('payments', [
            {'label': '第一期 (簽約訂金)', 'pct': 0.2, 'desc': '確認施工圖則、清拆及工程保護進場前付清'},
            {'label': '第二期 (泥水水電)', 'pct': 0.4, 'desc': '水電管線暗埋及泥水工程完工、驗收後付清'},
            {'label': '第三期 (木工油漆)', 'pct': 0.3, 'desc': '木工現場結構完成、油漆進場及訂造傢俬進場前付清'},
            {'label': '第四期 (工程尾數)', 'pct': 0.1, 'desc': '全屋工程完工、清潔交吉及驗收完成後付清'},
        ])

        for p in payments:
            apply_cell(ws, f'A{r}', value=p['label'], font=FONT_PAYMENT, border=BORDER_GRAY)
            apply_cell(ws, f'B{r}', value=p.get('label_pct', f'總金額之 {int(p["pct"]*100)}%'),
                       font=FONT_PAYMENT, alignment=ALIGN_CENTER, border=BORDER_GRAY)
            apply_cell(ws, f'C{r}', value=f'=F{grand_total_row}*{p["pct"]}',
                       font=FONT_PAYMENT_AMT, alignment=ALIGN_RIGHT, number_format=FMT_CURRENCY, border=BORDER_GRAY)
            apply_cell(ws, f'D{r}', value=p['desc'], font=FONT_PAYMENT, border=BORDER_GRAY)
            ws.merge_cells(f'D{r}:G{r}')
            ws.row_dimensions[r].height = ROW_HEIGHT_PAYMENT
            r += 1
        r += 1

    # ═══ 條款 ═══
    if show_terms:
        apply_cell(ws, f'A{r}', value='備註及條款說明：', font=FONT_TERMS_LABEL)
        r += 1

        terms = data.get('terms', [
            '本報價單所列工程，於完工後均享有 1 年結構及防漏水工程保養期。',
            '報價已包含工程期間之第三者責任保險及僱員補償保險（勞保）。',
            '所有加減工程（Variation Order）必須經雙方書面確認金額及工期變更後方可施工，口頭協議一律無效。',
            '電位計算定義：單位插座計1個位，雙位插座計1.5個位，燈制與燈位分開計算，現場放線時由雙方核對確認。',
            '預計工期為 90 個工作天（不計星期日及公眾假期），若因公司原因無故延誤，每逾期一日補償業主 HKD $500。',
        ])

        for i, t in enumerate(terms, 1):
            apply_cell(ws, f'A{r}', value=f'{i}. {t}', font=FONT_TERMS, alignment=ALIGN_WRAP)
            ws.merge_cells(f'A{r}:G{r}')
            ws.row_dimensions[r].height = ROW_HEIGHT_TERMS
            r += 1

    setup_a4_print(ws)

def render_html(data, title='報價單'):
    import base64, io as io_mod
    from openpyxl import Workbook as Wb

    section_items = {sec['num']: [] for sec in SECTIONS}
    for it in data.get('items', []):
        cat = it.get('category', '')
        for sec in SECTIONS:
            if cat == sec['cat']:
                section_items[sec['num']].append(it)
                break
        else:
            section_items[7].append(it)

    def esc(s):
        if s is None: return '-'
        return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')

    def fmt(n):
        try: return '${:,}'.format(int(n))
        except: return '$0'

    css = (
        "@page{size:A4;margin:0.5cm}"
        "body{font-family:'Microsoft JhengHei','微軟正黑體','PMingLiu','新細明體',serif;color:#000;margin:0;padding:8px;background:#e8e8e8}"
        ".page{width:auto;max-width:190mm;margin:0 auto;background:#fff;padding:0.4cm 0.5cm;box-shadow:0 2px 8px rgba(0,0,0,.1)}"
        "h1{text-align:center;font-size:16pt;font-weight:bold;margin:0 0 4px 0;padding-bottom:3px;border-bottom:1px solid #000}"
        "table.main{width:100%;border-collapse:collapse;font-size:10pt;table-layout:fixed}"
        "table.main td,table.main th{white-space:nowrap;padding:2px 4px;vertical-align:middle;border:1px solid #d9d9d9}"
        ".info-td{font-size:10pt;border:none!important;white-space:nowrap}"
        ".info-lbl{font-weight:bold;border:none!important}"
        ".info-val{border:none!important}"
        ".section{background:#F2F2F2;font-weight:bold;font-size:10pt;border:none!important}"
        ".th{font-weight:bold;border-bottom:2px solid #999!important;text-align:center;font-size:10pt;background:#fff}"
        ".th-l{text-align:left}.th-r{text-align:right}"
        ".td-c{text-align:center}.td-r{text-align:right}"
        ".subtotal td{border:none!important;background:#F2F2F2;font-weight:bold;font-size:10pt}"
        ".total-row td{border:none!important;font-size:10pt;font-weight:bold;background:#F2F2F2;text-align:right;padding:4px}"
        ".pay-title td,.terms-title td{border:none!important;font-weight:bold;font-size:10pt;padding-top:8px}"
        ".pay-th{font-weight:bold;border:1px solid #d9d9d9;text-align:center;font-size:10pt}"
        ".term-text td{border:none!important;font-size:9pt;color:#555;padding:1px 4px}"
        ".btn-bar{position:fixed;top:8px;right:8px;display:flex;gap:6px;z-index:99}"
        ".btn-bar button{padding:8px 14px;border:none;border-radius:4px;font-size:13px;cursor:pointer;font-weight:bold;color:#fff}"
        ".btn-excel{background:#1F4E78}.btn-pdf{background:#2E7D32}.btn-jpg{background:#E65100}.btn-all{background:#6A1B9A}"
        "@media print{@page{size:A4;margin:0.5cm}body{background:#fff;padding:0}.page{box-shadow:none;margin:0;max-width:none;padding:0.3cm 0.4cm}.btn-bar{display:none!important}}"
    )

    # Build items HTML
    items_html = ''
    section_counter = 0
    grand_total = 0

    for sec in SECTIONS:
        items = section_items[sec['num']]
        if not items: continue
        section_counter += 1
        items_html += '<tr><td colspan="7" class="section">' + CN_NUMS[section_counter-1] + '、 ' + esc(sec['title']) + '</td></tr>'
        sec_total = 0
        for i, it in enumerate(items):
            seq = str(section_counter) + '.' + str(i+1)
            qty = it.get('quantity', 1) or 1
            price = it.get('unit_price', 0) or 0
            amt = qty * price
            sec_total += amt
            items_html += '<tr><td class="td-c">' + seq + '</td><td>' + esc(it['description']) + '</td><td class="td-c">' + str(qty) + '</td><td class="td-c">' + esc(it.get('unit','項')) + '</td><td class="td-r">' + fmt(price) + '</td><td class="td-r">' + fmt(amt) + '</td><td>' + esc(it.get('remark','') or '-') + '</td></tr>'
        grand_total += sec_total
        items_html += '<tr class="subtotal"><td colspan="4"></td><td class="td-r">小計：</td><td class="td-r">' + fmt(sec_total) + '</td><td></td></tr>'

    deposit_html = ''
    deposit = data.get('deposit', 0)
    if title == '發票' and deposit > 0:
        balance = grand_total - deposit
        deposit_html = '<tr class="total-row"><td>訂金 (Deposit)：</td><td class="td-r">' + fmt(deposit) + '</td></tr>'
        deposit_html += '<tr class="total-row"><td>應付尾款 (Balance Due)：</td><td class="td-r">' + fmt(balance) + '</td></tr>'

    payment_html = ''
    if data.get('show_payment', True):
        payments = data.get('payments', [])
        if payments:
            payment_html = '<table class="main"><colgroup><col style="width:25%"><col style="width:15%"><col style="width:20%"><col style="width:40%"></colgroup>'
            payment_html += '<tr class="pay-title"><td colspan="4">工程付款階段說明：</td></tr>'
            payment_html += '<tr><th class="pay-th">付款期數</th><th class="pay-th">比例</th><th class="pay-th th-r">金額 (HKD)</th><th class="pay-th th-l">付款條件說明</th></tr>'
            for p in payments:
                pct = p.get('pct', 0)
                payment_html += '<tr><td>' + esc(p.get('label','')) + '</td><td class="td-c">' + esc(p.get('label_pct','')) + '</td><td class="td-r">' + fmt(int(grand_total*pct)) + '</td><td>' + esc(p.get('desc','')) + '</td></tr>'
            payment_html += '</table>'

    terms_html = ''
    if data.get('show_terms', True):
        terms = data.get('terms', [])
        if terms:
            terms_html = '<table class="main"><colgroup><col style="width:100%"></colgroup>'
            terms_html += '<tr class="terms-title"><td>備註及條款說明：</td></tr>'
            for i, t in enumerate(terms, 1):
                terms_html += '<tr class="term-text"><td>' + str(i) + '. ' + esc(t) + '</td></tr>'
            terms_html += '</table>'

    date_str = (data.get('date', '') or '').replace('-', '')
    addr = (data.get('address', '') or data.get('project_name', 'output'))

    wb = Wb()
    ws = wb.active
    generate_quotation(ws, data, title)
    xlsx_buf = io_mod.BytesIO()
    wb.save(xlsx_buf)
    xlsx_b64 = base64.b64encode(xlsx_buf.getvalue()).decode()

    html = '<!DOCTYPE html><html lang="zh-HK"><head><meta charset="UTF-8"><title>' + esc(title) + '</title><style>' + css + '</style></head><body>'
    html += '<div class="btn-bar"><button class="btn-excel" onclick="downloadExcel()">下載 Excel</button><button class="btn-pdf" onclick="downloadPDF()">下載 PDF</button><button class="btn-jpg" onclick="downloadJPG()">下載 JPG</button></div>'
    html += '<div class="page" id="capture">'
    html += '<h1>' + esc(title) + '</h1>'

    # Info table
    html += '<table class="main"><colgroup><col style="width:15%"><col style="width:35%"><col style="width:15%"><col style="width:35%"></colgroup>'
    html += '<tr><td class="info-lbl">工程名稱：</td><td class="info-val">' + esc(data.get('project_name','-')) + '</td><td class="info-lbl">報價單號：</td><td class="info-val">' + esc(data.get('quotation_no','-')) + '</td></tr>'
    html += '<tr><td class="info-lbl">客戶姓名：</td><td class="info-val">' + esc(data.get('owner_name','-')) + '</td><td class="info-lbl">報價日期：</td><td class="info-val">' + esc(data.get('date','-')) + '</td></tr>'
    html += '<tr><td class="info-lbl">工程地址：</td><td class="info-val">' + esc(data.get('address','-')) + '</td><td class="info-lbl">有效期：</td><td class="info-val">' + esc(data.get('validity','-')) + '</td></tr>'
    html += '<tr><td class="info-lbl">裝修公司：</td><td class="info-val">' + esc(data.get('company_name','-')) + '</td><td class="info-lbl">版本：</td><td class="info-val">' + esc(data.get('version','-')) + '</td></tr>'
    html += '</table><br>'

    # Items table
    html += '<table class="main"><colgroup><col style="width:8%"><col style="width:31%"><col style="width:7%"><col style="width:6%"><col style="width:11%"><col style="width:13%"><col style="width:24%"></colgroup>'
    html += '<tr><th class="th">項目編號</th><th class="th th-l">工程項目及說明</th><th class="th">數量</th><th class="th">單位</th><th class="th th-r">單價 (HKD)</th><th class="th th-r">複價 (HKD)</th><th class="th th-l">備註</th></tr>'
    html += items_html
    html += '</table>'

    # Total table
    html += '<table class="main"><colgroup><col style="width:79%"><col style="width:21%"></colgroup>'
    html += '<tr class="total-row"><td>總工程預算總計 (HKD)：</td><td class="td-r">' + fmt(grand_total) + '</td></tr>'
    if deposit_html:
        html += deposit_html
    html += '</table>'

    # Payment table
    if payment_html:
        html += '<br>' + payment_html

    # Terms table
    if terms_html:
        html += '<br>' + terms_html

    html += '</div>'

    # JS
    js = '<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script><script>'
    js += 'var XLSX_B64="' + xlsx_b64 + '";'
    js += 'var TITLE="' + esc(title) + '";'
    js += 'var ADDR="' + esc(addr) + '";'
    js += 'var DATE="' + date_str + '";'
    js += 'var PID="' + esc(data.get('_preview_id', '')) + '";'
    js += 'function downloadExcel(){window.location.href="/download/"+PID+"/excel";}'
    js += 'function downloadPDF(){window.location.href="/download/"+PID+"/pdf";}'
    js += 'async function downloadJPG(){var el=document.getElementById("capture");var w=el.offsetWidth;var canvas=await html2canvas(el,{width:w,scale:2,backgroundColor:"#ffffff",windowWidth:w});canvas.toBlob(function(blob){var url=URL.createObjectURL(blob);var link=document.createElement("a");link.href=url;link.download=ADDR+"_"+TITLE+"_"+DATE+".jpg";link.click();URL.revokeObjectURL(url);},"image/jpeg",0.92);}'
    js += '</script></body></html>'
    html += js
    return html

def _write_table_header(ws, row):
    headers = ['項目編號', '工程項目及說明', '數量', '單位', '單價 (HKD)', '複價 (HKD)', '備註']
    alignments = [ALIGN_CENTER, ALIGN_LEFT, ALIGN_CENTER, ALIGN_CENTER,
                  ALIGN_RIGHT, ALIGN_RIGHT, ALIGN_LEFT]
    for i, (hdr, al) in enumerate(zip(headers, alignments)):
        col = get_column_letter(i + 1)
        apply_cell(ws, f'{col}{row}', value=hdr, font=FONT_TABLE_HEADER,
                   alignment=al, border=BORDER_HEADER_BOTTOM)
    ws.row_dimensions[row].height = 22.0
