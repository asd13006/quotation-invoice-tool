"""
純 Python xlsx → PDF 渲染器（追 MiniPdf 100% 品質）
Canvas-based，逐格精確渲染，跟足 cell font/size/bold/color/border
"""
import io, os, re
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

PAGE_W, PAGE_H = A4
MARGIN = 8 * mm

# ── CJK 字型 ──
FONT = 'Helvetica'
FONT_BOLD = 'Helvetica-Bold'

for path, name in [
    ('C:/Windows/Fonts/mingliu.ttc', 'MingLiU'),
    ('C:/Windows/Fonts/msjh.ttc', 'MSJH'),
    ('/System/Library/Fonts/PingFang.ttc', 'PingFang'),
]:
    if os.path.exists(path):
        try:
            pdfmetrics.registerFont(TTFont(name, path, subfontIndex=0))
            pdfmetrics.registerFont(TTFont(name+'-Bold', path, subfontIndex=1))
            FONT = name; FONT_BOLD = name+'-Bold'
            break
        except:
            pass

if FONT == 'Helvetica':
    try:
        pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
        pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
        FONT = 'STSong-Light'; FONT_BOLD = 'HeiseiKakuGo-W5'
    except:
        pass


def col_letter_to_index(letter):
    return column_index_from_string(letter) - 1


def _eval_formula(formula, ws, row, col):
    if not formula or not formula.startswith('='):
        return formula
    try:
        expr = formula[1:]
        def cv(ref):
            m = re.match(r'\$?([A-G])\$?(\d+)', ref)
            if not m: return 0
            v = ws[f'{m.group(1)}{m.group(2)}'].value
            if v is None: return 0
            if isinstance(v, str) and v.startswith('='):
                return float(_eval_formula(v, ws, int(m.group(2)), 0) or 0)
            try: return float(v)
            except: return 0
        expr = re.sub(r'\$?([A-G])\$?(\d+)', lambda m: str(cv(m.group(0))), expr)
        expr = re.sub(r'SUM\(([A-G]\d+):([A-G]\d+)\)',
                      lambda m: str(_sum_range(ws, m.group(1), m.group(2))),
                      expr, flags=re.IGNORECASE)
        r = eval(expr)
        if isinstance(r, float) and r == int(r): r = int(r)
        return str(r)
    except:
        return '0'


def _sum_range(ws, s, e):
    m1 = re.match(r'\$?([A-G])\$?(\d+)', s)
    m2 = re.match(r'\$?([A-G])\$?(\d+)', e)
    if not m1 or not m2: return 0
    col, r1, r2 = m1.group(1), int(m1.group(2)), int(m2.group(2))
    t = 0
    for r in range(r1, r2 + 1):
        v = ws[f'{col}{r}'].value
        if v is None: continue
        if isinstance(v, str) and v.startswith('='): v = _eval_formula(v, ws, r, 0)
        try: t += float(v)
        except: pass
    return t


def _fmt_currency(val):
    try:
        n = int(float(val))
        return f'${n:,}'
    except:
        return str(val)


def _is_currency(cell):
    try:
        nf = cell.number_format or ''
        return '$' in nf or '#,##0' in nf
    except:
        return False


def convert_xlsx_to_pdf(xlsx_bytes):
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active

    buf = io.BytesIO()
    c = pdfcanvas.Canvas(buf, pagesize=A4)

    # ── Merged cells ──
    merged = {}
    for m in ws.merged_cells.ranges:
        r1, r2 = m.min_row, m.max_row
        c1 = col_letter_to_index(m.min_col) if isinstance(m.min_col, str) else m.min_col - 1
        c2 = col_letter_to_index(m.max_col) if isinstance(m.max_col, str) else m.max_col - 1
        for ri in range(r1, r2 + 1):
            for ci in range(c1, c2 + 1):
                merged[(ri, ci)] = (r1, c1, r2, c2)

    # ── Column widths ──
    col_widths = []
    for ci in range(ws.max_column):
        letter = get_column_letter(ci + 1)
        w = ws.column_dimensions.get(letter)
        col_widths.append((w.width or 10) * 6.5 if w and w.width else 65)

    total_w = sum(col_widths)
    scale = (PAGE_W - 2 * MARGIN) / max(total_w, 1)

    col_x = []
    x = MARGIN
    for w in col_widths:
        col_x.append(x)
        x += w * scale

    # ── Row heights ──
    row_heights = {}
    for ri in range(1, ws.max_row + 1):
        h = ws.row_dimensions.get(ri)
        row_heights[ri] = (h.height * 0.75) if (h and h.height) else 14

    # ── Pass 1: collect all cell render data ──
    cells = []  # list of dicts
    y = PAGE_H - MARGIN

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        r = row[0].row
        rh = row_heights.get(r, 14)

        for cell in row:
            ci = cell.column - 1
            if ci >= len(col_x) or cell.value is None:
                continue

            # Merged cell
            if (r, ci) in merged:
                mr1, mc1, mr2, mc2 = merged[(r, ci)]
                if (r, ci) != (mr1, mc1):
                    continue
                cw = sum(col_widths[mc1:mc2+1]) * scale
                ch = sum(row_heights.get(ri, 14) for ri in range(mr1, mr2+1))
                cx = col_x[mc1]
                cy = y - sum(row_heights.get(ri, 14) for ri in range(r, mr2+1))
            else:
                cw = col_widths[ci] * scale
                ch = rh
                cx = col_x[ci]
                cy = y - rh

            # Cell value
            val = str(cell.value)
            if val.startswith('='):
                val = _eval_formula(val, ws, r, ci)

            # Currency format
            if _is_currency(cell) and val.lstrip('-').replace('.','',1).isdigit():
                val = _fmt_currency(val)

            # Font properties
            font_name = FONT
            font_size = 8
            font_color = colors.black
            is_bold = False
            try:
                if cell.font:
                    if cell.font.size: font_size = cell.font.size
                    if cell.font.bold: is_bold = True
                    if cell.font.color and cell.font.color.rgb:
                        rgb = cell.font.color.rgb
                        if rgb not in ('00000000', '0'):
                            font_color = colors.HexColor('#' + rgb[2:])
            except: pass

            font_name = FONT_BOLD if is_bold else FONT
            font_size = max(min(font_size, 24), 5)

            # Alignment
            ha = 'left'
            try:
                if cell.alignment and cell.alignment.horizontal:
                    ha = cell.alignment.horizontal
            except: pass

            # Background
            bg_color = None
            try:
                fill = cell.fill
                if fill.patternType == 'solid':
                    rgb = fill.fgColor.rgb
                    if rgb and rgb not in ('00000000', '0'):
                        bg_color = colors.HexColor('#' + rgb[2:])
            except: pass

            # Border
            border_color = None
            try:
                b = cell.border
                for side in [b.left, b.right, b.top, b.bottom]:
                    if side and side.style:
                        border_color = colors.HexColor('#D9D9D9')
                        break
            except: pass

            cells.append({
                'cx': cx, 'cy': cy, 'cw': cw, 'ch': ch,
                'val': val, 'ha': ha,
                'font_name': font_name, 'font_size': font_size, 'font_color': font_color,
                'bg_color': bg_color, 'border_color': border_color,
            })

        y -= rh

    # ── Pass 2: backgrounds (bottom layer) ──
    for cell in cells:
        if cell['bg_color']:
            c.setFillColor(cell['bg_color'])
            c.rect(cell['cx'], cell['cy'], cell['cw'], cell['ch'], fill=1, stroke=0)

    # ── Pass 3: borders (middle layer) ──
    for cell in cells:
        if cell['border_color']:
            c.setStrokeColor(cell['border_color'])
            c.setLineWidth(0.3)
            c.rect(cell['cx'], cell['cy'], cell['cw'], cell['ch'])

    # ── Pass 4: text (top layer) ──
    # Reset to fill/stroke defaults
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)

    for cell in cells:
        c.setFont(cell['font_name'], cell['font_size'])
        c.setFillColor(cell['font_color'])
        val = cell['val']
        cx, cy, cw, ch = cell['cx'], cell['cy'], cell['cw'], cell['ch']
        padding = 2

        if cell['ha'] == 'center':
            c.drawCentredString(cx + cw/2, cy + padding, val)
        elif cell['ha'] == 'right':
            c.drawRightString(cx + cw - padding, cy + padding, val)
        else:
            c.drawString(cx + padding, cy + padding, val)

    # ── 頁尾底線 ──
    c.setStrokeColor(colors.HexColor('#999999'))
    c.setLineWidth(0.5)
    c.line(MARGIN, MARGIN + 5, PAGE_W - MARGIN, MARGIN + 5)

    c.save()
    buf.seek(0)
    return buf.read()


# ── CLI ──
if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print('xlsx2pdf — 純 Python xlsx → PDF 渲染器')
        print('用法: python xlsx2pdf.py <input.xlsx> [-o output.pdf]')
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = input_path.replace('.xlsx', '.pdf')

    i = 2
    while i < len(sys.argv):
        if sys.argv[i] == '-o' and i + 1 < len(sys.argv):
            output_path = sys.argv[i + 1]; i += 2
        else:
            i += 1

    with open(input_path, 'rb') as f:
        pdf = convert_xlsx_to_pdf(f.read())

    with open(output_path, 'wb') as f:
        f.write(pdf)

    print(f'OK: {input_path} → {output_path} ({len(pdf)} bytes)')
