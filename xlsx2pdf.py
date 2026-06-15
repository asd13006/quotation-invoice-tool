"""
純 Python xlsx → PDF 渲染器（Vercel 相容）
準確處理 merged cells、欄闊、行高、字體大小
"""
import io, os
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
for path, name in [
    ('C:/Windows/Fonts/msjh.ttc', 'MSJH'),
    ('C:/Windows/Fonts/mingliu.ttc', 'MingLiU'),
    ('/System/Library/Fonts/PingFang.ttc', 'PingFang'),
    ('/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc', 'WQY'),
]:
    if os.path.exists(path):
        try:
            pdfmetrics.registerFont(TTFont(name, path))
            FONT = name; break
        except: pass

if FONT == 'Helvetica':
    try:
        pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
        FONT = 'STSong-Light'
    except: pass


def col_letter_to_index(letter):
    return column_index_from_string(letter) - 1


def _eval_formula(formula, ws, row, col):
    """簡單公式求值：=C9*E9 → qty * price"""
    if not formula.startswith('='):
        return formula
    try:
        # =C*E pattern
        expr = formula[1:]
        for letter in ['A','B','C','D','E','F','G']:
            if letter in expr:
                cell_val = ws[f'{letter}{row}'].value
                if cell_val is None: cell_val = 0
                expr = expr.replace(letter + str(row), str(cell_val))
        # Also handle $letter$row pattern
        import re
        expr = re.sub(r'\$?([A-G])\$?(\d+)', lambda m: str(ws[f'{m.group(1)}{m.group(2)}'].value or 0), expr)
        result = eval(expr)
        return str(int(result)) if isinstance(result, float) and result == int(result) else str(result)
    except:
        return '0'


def convert_xlsx_to_pdf(xlsx_bytes):
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active

    buf = io.BytesIO()
    c = pdfcanvas.Canvas(buf, pagesize=A4)

    # ── 收集 merged cell ranges ──
    merged = {}  # (row, col_idx) → (r1, c1, r2, c2)
    for m in ws.merged_cells.ranges:
        r1, r2 = m.min_row, m.max_row
        c1 = col_letter_to_index(m.min_col) if isinstance(m.min_col, str) else m.min_col - 1
        c2 = col_letter_to_index(m.max_col) if isinstance(m.max_col, str) else m.max_col - 1
        for ri in range(r1, r2 + 1):
            for ci in range(c1, c2 + 1):
                merged[(ri, ci)] = (r1, c1, r2, c2)

    # ── 收集所有有值嘅 column ──
    col_set = set()
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for cell in row:
            if cell.value is not None:
                col_set.add(cell.column - 1)

    # ── 欄闊計算 ──
    col_widths = []  # list of width in points
    for ci in range(ws.max_column):
        letter = get_column_letter(ci + 1)
        w = ws.column_dimensions.get(letter)
        if w and w.width:
            col_widths.append(w.width * 7.0)  # 1 unit ≈ 7pt
        elif ci in col_set:
            col_widths.append(60)  # default
        else:
            col_widths.append(0)

    total_w = sum(col_widths)
    scale = (PAGE_W - 2 * MARGIN) / max(total_w, 1)

    # column x positions
    col_x = []
    x = MARGIN
    for w in col_widths:
        col_x.append(x)
        x += w * scale

    # ── 行高 ──
    row_heights = {}
    for ri in range(1, ws.max_row + 1):
        h = ws.row_dimensions.get(ri)
        if h and h.height:
            row_heights[ri] = h.height * 0.75  # pt conversion
        else:
            row_heights[ri] = 16

    # ── Render each cell ──
    y = PAGE_H - MARGIN
    rendered = set()  # (row, col) to skip merged sub-cells

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        r = row[0].row
        row_h = row_heights.get(r, 16)

        if y < MARGIN + 30:
            c.showPage()
            y = PAGE_H - MARGIN

        for cell in row:
            if cell.value is None: continue

            ci = cell.column - 1
            if ci >= len(col_x): continue

            # Skip merged sub-cells
            if (r, ci) in merged:
                mr1, mc1, mr2, mc2 = merged[(r, ci)]
                if (r, ci) != (mr1, mc1):  # not top-left
                    continue
                # Calculate merged cell dimensions
                cell_w = sum(col_widths[mc1:mc2 + 1]) * scale
                cell_h = sum(row_heights.get(ri, 16) for ri in range(mr1, mr2 + 1))
                cx = col_x[mc1]
                cy = y - sum(row_heights.get(ri, 16) for ri in range(r, mr2 + 1))
            else:
                cell_w = col_widths[ci] * scale if ci < len(col_widths) else 60 * scale
                cell_h = row_h
                cx = col_x[ci]
                cy = y - row_h

            val = str(cell.value)
            if val.startswith('='):
                val = _eval_formula(val, ws, r, ci)
            val = val[:100]

            # Background
            try:
                fill = cell.fill
                if fill.patternType == 'solid':
                    rgb = fill.fgColor.rgb
                    if rgb and rgb not in ('00000000', '0'):
                        c.setFillColor(colors.HexColor('#' + rgb[2:]))
                        c.rect(cx, cy, cell_w, cell_h, fill=1, stroke=0)
                        c.setFillColor(colors.black)
            except: pass

            # Border
            try:
                b = cell.border
                has_b = any(getattr(b, s).style for s in ['left', 'right', 'top', 'bottom'])
                if has_b:
                    c.setStrokeColor(colors.HexColor('#D9D9D9'))
                    c.setLineWidth(0.3)
                    c.rect(cx, cy, cell_w, cell_h)
            except: pass

            # Font size
            fs = 8
            try:
                if cell.font and cell.font.size:
                    fs = min(cell.font.size * 0.75, 11)
            except: pass

            c.setFont(FONT, fs)

            # Alignment
            ha = cell.alignment.horizontal or 'left'
            va = cell.alignment.vertical or 'bottom'
            padding = 3

            if ha == 'center':
                tx = cx + cell_w / 2
                c.drawCentredString(tx, cy + padding, val)
            elif ha == 'right':
                tx = cx + cell_w - padding
                c.drawRightString(tx, cy + padding, val)
            else:
                c.drawString(cx + padding, cy + padding, val)

        y -= row_h

    c.save()
    buf.seek(0)
    return buf.read()
