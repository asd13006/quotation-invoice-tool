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
FONT_BOLD = 'Helvetica-Bold'
_use_cid = False

# 方法 1: 試系統 TTF/TTC 字型
for path, name in [
    ('C:/Windows/Fonts/msjh.ttc', 'MSJH'),     # 微軟正黑體
    ('C:/Windows/Fonts/mingliu.ttc', 'MingLiU'), # 細明體
    ('/System/Library/Fonts/PingFang.ttc', 'PingFang'),
    ('/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc', 'WQY'),
]:
    if os.path.exists(path):
        try:
            pdfmetrics.registerFont(TTFont(name, path))
            FONT = name
            FONT_BOLD = name  # TTFont supports bold via setFont(..., bold)
            break
        except:
            pass

# 方法 2: reportlab 內建 CID 字型（Vercel fallback）
if FONT == 'Helvetica':
    try:
        pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
        pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
        FONT = 'STSong-Light'
        FONT_BOLD = 'HeiseiKakuGo-W5'
        _use_cid = True
    except:
        pass


def col_letter_to_index(letter):
    return column_index_from_string(letter) - 1


def _eval_formula(formula, ws, row, col):
    """公式求值：支援 =C*E, =SUM(F9:F12), =F13+F18, =F29*0.2 等"""
    if not formula.startswith('='):
        return formula
    try:
        expr = formula[1:]
        import re

        # Helper: get cell value as number
        def cell_val(ref):
            m = re.match(r'\$?([A-G])\$?(\d+)', ref)
            if m:
                v = ws[f'{m.group(1)}{m.group(2)}'].value
                if v is None: return 0
                if isinstance(v, str) and v.startswith('='):
                    return float(_eval_formula(v, ws, int(m.group(2)), 0) or 0)
                return float(v) if v else 0
            return 0

        # Replace cell references with values
        expr = re.sub(r'\$?([A-G])\$?(\d+)', lambda m: str(cell_val(m.group(0))), expr)

        # Handle SUM(range) → sum of range
        expr = re.sub(r'SUM\(([A-G]\d+):([A-G]\d+)\)', lambda m: str(_sum_range(ws, m.group(1), m.group(2))), expr, flags=re.IGNORECASE)

        result = eval(expr)
        if isinstance(result, float) and result == int(result):
            result = int(result)
        return str(result)
    except Exception as e:
        return '0'


def _sum_range(ws, start_ref, end_ref):
    """計算 SUM(F9:F12) 嘅值"""
    import re
    m1 = re.match(r'\$?([A-G])\$?(\d+)', start_ref)
    m2 = re.match(r'\$?([A-G])\$?(\d+)', end_ref)
    if not m1 or not m2: return 0
    col = m1.group(1)
    r1, r2 = int(m1.group(2)), int(m2.group(2))
    total = 0
    for r in range(r1, r2 + 1):
        v = ws[f'{col}{r}'].value
        if v is None: continue
        if isinstance(v, str) and v.startswith('='):
            v = _eval_formula(v, ws, r, 0)
        try: total += float(v)
        except: pass
    return total


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

            # Font size + bold (跟足 xlsx，font.size 已係 point size)
            fs = 8 * scale  # default scaled
            is_bold = False
            try:
                if cell.font:
                    if cell.font.size:
                        fs = cell.font.size * scale  # 直接用 xlsx point size，配合 scale
                    if cell.font.bold:
                        is_bold = True
            except:
                pass

            fs = max(min(fs, 16), 5)  # clamp

            # Use bold font if needed
            font_name = FONT_BOLD if is_bold and not _use_cid else FONT
            c.setFont(font_name, fs)

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
