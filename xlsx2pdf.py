"""
xlsx → PDF 轉換器
優先: LibreOffice（完美）
Fallback: 純 Python renderer（近似）
"""
import io, os, subprocess, shutil, tempfile

# 偵測 LibreOffice
_LIBREOFFICE = None
for path in [
    'libreoffice',
    '/usr/bin/libreoffice',
    'C:/Program Files/LibreOffice/program/soffice.exe',
    'C:/Program Files (x86)/LibreOffice/program/soffice.exe',
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
]:
    if shutil.which(path) or os.path.exists(path):
        _LIBREOFFICE = path if os.path.exists(path) else shutil.which(path)
        break


def convert_xlsx_to_pdf(xlsx_bytes):
    """xlsx bytes → PDF bytes"""
    # 方法 1: LibreOffice（完美）
    if _LIBREOFFICE:
        return _convert_libreoffice(xlsx_bytes)

    # 方法 2: 純 Python fallback（近似）
    return _convert_pure_python(xlsx_bytes)


def _convert_libreoffice(xlsx_bytes):
    """用 LibreOffice headless 轉換"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as xf:
        xf.write(xlsx_bytes)
        xlsx_path = xf.name

    out_dir = tempfile.mkdtemp()

    try:
        subprocess.run([
            _LIBREOFFICE, '--headless', '--convert-to', 'pdf',
            '--outdir', out_dir, xlsx_path
        ], capture_output=True, timeout=30, check=True)

        pdf_name = os.path.basename(xlsx_path).replace('.xlsx', '.pdf')
        pdf_path = os.path.join(out_dir, pdf_name)

        with open(pdf_path, 'rb') as pf:
            return pf.read()
    finally:
        if os.path.exists(xlsx_path): os.unlink(xlsx_path)
        if os.path.exists(out_dir):
            for f in os.listdir(out_dir):
                os.unlink(os.path.join(out_dir, f))
            os.rmdir(out_dir)


def _convert_pure_python(xlsx_bytes):
    """純 Python fallback — 精簡版"""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter, column_index_from_string

    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.platypus.flowables import HRFlowable
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont

    PAGE_W, PAGE_H = A4
    MARGIN = 8 * mm

    # 字型
    FONT = 'Helvetica'; FONT_BOLD = 'Helvetica-Bold'
    for path, name in [
        ('C:/Windows/Fonts/mingliu.ttc', 'MingLiU'),
        ('C:/Windows/Fonts/msjh.ttc', 'MSJH'),
        ('/System/Library/Fonts/PingFang.ttc', 'PingFang'),
    ]:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(name, path, subfontIndex=0))
                pdfmetrics.registerFont(TTFont(name+'-B', path, subfontIndex=1))
                FONT = name; FONT_BOLD = name+'-B'; break
            except: pass
    if FONT == 'Helvetica':
        try:
            pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
            pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
            FONT = 'STSong-Light'; FONT_BOLD = 'HeiseiKakuGo-W5'
        except: pass

    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active

    # Merged cells
    merged = {}
    for m in ws.merged_cells.ranges:
        r1, r2 = m.min_row, m.max_row
        c1 = column_index_from_string(m.min_col) - 1 if isinstance(m.min_col, str) else m.min_col - 1
        c2 = column_index_from_string(m.max_col) - 1 if isinstance(m.max_col, str) else m.max_col - 1
        for ri in range(r1, r2+1):
            for ci in range(c1, c2+1):
                merged[(ri, ci)] = (r1, c1, r2, c2)

    # Column widths
    col_widths = []
    for ci in range(ws.max_column):
        letter = get_column_letter(ci+1)
        w = ws.column_dimensions.get(letter)
        col_widths.append((w.width or 10)*6.5 if w and w.width else 65)

    total_w = sum(col_widths)
    scale = (PAGE_W - 2*MARGIN) / max(total_w, 1)
    col_w = [w*scale for w in col_widths]

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=MARGIN, rightMargin=MARGIN,
                            topMargin=MARGIN, bottomMargin=MARGIN)

    # Build table
    table_data = []
    row_map = []
    style_cmds = [
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
    ]

    _cache = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        r = row[0].row
        rcells = []
        has = False
        for cell in row:
            ci = cell.column - 1
            if ci >= len(col_w): continue
            if (r, ci) in merged:
                mr1, mc1, mr2, mc2 = merged[(r,ci)]
                if (r, ci) != (mr1, mc1): rcells.append(''); continue

            val = str(cell.value or '')
            if val.startswith('='): continue  # skip formulas in fallback
            if val in ('None', ''): rcells.append(''); continue

            # Style
            fn, fs, fc, ha = FONT, 9, colors.black, TA_LEFT
            try:
                if cell.font:
                    if cell.font.size: fs = max(min(cell.font.size, 20), 6)
                    if cell.font.bold: fn = FONT_BOLD
            except: pass
            try:
                if cell.alignment and cell.alignment.horizontal:
                    ha = {'center': TA_CENTER, 'right': TA_RIGHT}.get(
                        cell.alignment.horizontal, TA_LEFT)
            except: pass

            key = (fn, fs, ha)
            if key not in _cache:
                _cache[key] = ParagraphStyle(f's{len(_cache)}', fontName=fn,
                    fontSize=fs, leading=fs+3, textColor=fc, alignment=ha)
            rcells.append(Paragraph(val.replace('&','&amp;').replace('<','&lt;'), _cache[key]))
            has = True

            # BG/Border
            try:
                fill = cell.fill
                if fill.patternType == 'solid':
                    rgb = fill.fgColor.rgb
                    if rgb and rgb not in ('00000000','0'):
                        style_cmds.append(('BACKGROUND', (ci, len(table_data)), (ci, len(table_data)),
                                          colors.HexColor('#'+rgb[2:])))
            except: pass
            try:
                if any(getattr(cell.border, s).style for s in ['left','right','top','bottom']):
                    style_cmds.append(('BOX', (ci, len(table_data)), (ci, len(table_data)),
                                      0.3, colors.HexColor('#D9D9D9')))
            except: pass

        if has:
            while len(rcells) < len(col_w): rcells.append('')
            table_data.append(rcells)
            row_map.append(r)

    # Merged cell SPAN
    for ti, r in enumerate(row_map):
        for cell in ws.iter_rows(min_row=r, max_row=r):
            for cell in cell:
                ci = cell.column - 1
                if (r, ci) in merged:
                    mr1, mc1, mr2, mc2 = merged[(r,ci)]
                    if (r, ci) == (mr1, mc1):
                        end_ti = next((tj for tj, rj in enumerate(row_map) if rj == mr2), ti)
                        style_cmds.append(('SPAN', (mc1, ti), (mc2, end_ti)))

    t = Table(table_data, colWidths=col_w)
    t.setStyle(TableStyle(style_cmds))
    doc.build([t, Spacer(1, 5*mm),
              HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#999999'))])
    buf.seek(0)
    return buf.read()


# ── CLI ──
if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print('xlsx2pdf — xlsx → PDF 轉換器')
        print(f'引擎: {"LibreOffice" if _LIBREOFFICE else "純 Python"}')
        print('用法: python xlsx2pdf.py <input.xlsx> [-o output.pdf]')
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = input_path.replace('.xlsx', '.pdf')
    i = 2
    while i < len(sys.argv):
        if sys.argv[i] == '-o' and i + 1 < len(sys.argv):
            output_path = sys.argv[i + 1]; i += 2
        else: i += 1

    with open(input_path, 'rb') as f:
        pdf = convert_xlsx_to_pdf(f.read())
    with open(output_path, 'wb') as f:
        f.write(pdf)

    print(f'OK: {input_path} → {output_path} ({len(pdf)} bytes) [{("LibreOffice" if _LIBREOFFICE else "純 Python")}]')
