"""
純 Python xlsx → PDF 渲染器（Vercel 相容，追 MiniPdf 品質）
用 reportlab platypus.Table + Paragraph 精確渲染
"""
import io, os, re
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.platypus.flowables import HRFlowable
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

PAGE_W, PAGE_H = A4
MARGIN = 10 * mm

# ── CJK 字型 ──
FONT = 'Helvetica'
FONT_BOLD = 'Helvetica-Bold'
FONT_SIZE = 9

for path, name, bold_name in [
    ('C:/Windows/Fonts/mingliu.ttc', 'MingLiU', 'MingLiU-B'),
    ('C:/Windows/Fonts/msjh.ttc', 'MSJH', 'MSJH-B'),
    ('/System/Library/Fonts/PingFang.ttc', 'PingFang', 'PingFang'),
]:
    if os.path.exists(path):
        try:
            pdfmetrics.registerFont(TTFont(name, path, subfontIndex=0))
            FONT = name
            try:
                pdfmetrics.registerFont(TTFont(bold_name, path, subfontIndex=1))
                FONT_BOLD = bold_name
            except:
                FONT_BOLD = name
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
    if not formula.startswith('='): return formula
    try:
        expr = formula[1:]

        def cell_val(ref):
            m = re.match(r'\$?([A-G])\$?(\d+)', ref)
            if not m: return 0
            v = ws[f'{m.group(1)}{m.group(2)}'].value
            if v is None: return 0
            if isinstance(v, str) and v.startswith('='):
                return float(_eval_formula(v, ws, int(m.group(2)), 0) or 0)
            try: return float(v)
            except: return 0

        expr = re.sub(r'\$?([A-G])\$?(\d+)', lambda m: str(cell_val(m.group(0))), expr)
        expr = re.sub(r'SUM\(([A-G]\d+):([A-G]\d+)\)',
                      lambda m: str(_sum_range(ws, m.group(1), m.group(2))),
                      expr, flags=re.IGNORECASE)
        result = eval(expr)
        if isinstance(result, float) and result == int(result): result = int(result)
        return str(result)
    except:
        return '0'


def _sum_range(ws, start_ref, end_ref):
    m1 = re.match(r'\$?([A-G])\$?(\d+)', start_ref)
    m2 = re.match(r'\$?([A-G])\$?(\d+)', end_ref)
    if not m1 or not m2: return 0
    col, r1, r2 = m1.group(1), int(m1.group(2)), int(m2.group(2))
    total = 0
    for r in range(r1, r2 + 1):
        v = ws[f'{col}{r}'].value
        if v is None: continue
        if isinstance(v, str) and v.startswith('='):
            v = _eval_formula(v, ws, r, 0)
        try: total += float(v)
        except: pass
    return total


def _fmt_currency(val):
    """Format number as $X,XXX"""
    try:
        n = int(float(val))
        return f'${n:,}'
    except:
        return str(val)


def convert_xlsx_to_pdf(xlsx_bytes):
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=MARGIN, rightMargin=MARGIN,
                            topMargin=MARGIN, bottomMargin=MARGIN)
    story = []

    # ── Collect merged cells ──
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
        col_widths.append((w.width or 10) * 7.0 if w and w.width else 60)

    total_w = sum(col_widths)
    scale = (PAGE_W - 2 * MARGIN) / max(total_w, 1)
    col_w = [w * scale for w in col_widths]

    # ── Styles ──
    s_normal = ParagraphStyle('n', fontName=FONT, fontSize=FONT_SIZE, leading=FONT_SIZE+3)
    s_bold = ParagraphStyle('b', fontName=FONT_BOLD, fontSize=FONT_SIZE, leading=FONT_SIZE+3)
    s_center = ParagraphStyle('c', fontName=FONT, fontSize=FONT_SIZE, leading=FONT_SIZE+3, alignment=1)
    s_bold_center = ParagraphStyle('bc', fontName=FONT_BOLD, fontSize=FONT_SIZE, leading=FONT_SIZE+3, alignment=1)
    s_right = ParagraphStyle('r', fontName=FONT, fontSize=FONT_SIZE, leading=FONT_SIZE+3, alignment=2)
    s_bold_right = ParagraphStyle('br', fontName=FONT_BOLD, fontSize=FONT_SIZE, leading=FONT_SIZE+3, alignment=2)
    s_title = ParagraphStyle('t', fontName=FONT_BOLD, fontSize=16, leading=20, alignment=1)
    s_section = ParagraphStyle('sec', fontName=FONT_BOLD, fontSize=FONT_SIZE+1, leading=FONT_SIZE+5)
    s_small = ParagraphStyle('sm', fontName=FONT, fontSize=FONT_SIZE-1, leading=FONT_SIZE+1, textColor=colors.HexColor('#555555'))

    def p(val, style=s_normal, num_fmt=None):
        if val is None or val == '': return Paragraph('-', style)
        v = str(val)
        if v.startswith('='):
            v = _eval_formula(v, ws, 0, 0)  # row/col not used for simple eval
        if num_fmt == 'currency':
            v = _fmt_currency(v)
        return Paragraph(v.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'), style)

    def get_style(cell, bold_style=None, center_style=None, right_style=None):
        """Get appropriate ParagraphStyle based on cell formatting"""
        is_bold = False
        ha = 'left'
        try:
            if cell.font and cell.font.bold: is_bold = True
            if cell.alignment and cell.alignment.horizontal: ha = cell.alignment.horizontal
        except: pass

        if ha == 'center':
            return s_bold_center if is_bold else s_center
        elif ha == 'right':
            return s_bold_right if is_bold else s_right
        else:
            return s_bold if is_bold else s_normal

    # ── Build table data ──
    table_data = []
    row_spans = {}  # track row merge info

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        r = row[0].row
        row_cells = []

        for cell in row:
            ci = cell.column - 1
            if ci >= len(col_w): continue

            # Check if this cell is a merged sub-cell
            if (r, ci) in merged:
                mr1, mc1, mr2, mc2 = merged[(r, ci)]
                if (r, ci) != (mr1, mc1):
                    row_cells.append('')  # placeholder for merged sub-cell
                    continue

            if cell.value is None:
                row_cells.append('')
                continue

            val = str(cell.value)
            style = get_style(cell)
            num_fmt = 'currency' if cell.number_format and '$' in (cell.number_format or '') else None
            row_cells.append(p(val, style, num_fmt))

        if any(c != '' for c in row_cells):
            table_data.append(row_cells)

    # ── Table style ──
    bg_style_cmds = []
    border_cmds = []
    row_idx = 0

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        r = row[0].row
        has_content = False
        for cell in row:
            if cell.value is not None:
                has_content = True
                ci = cell.column - 1
                if ci >= len(col_w): continue

                # Merged cell span
                if (r, ci) in merged:
                    mr1, mc1, mr2, mc2 = merged[(r, ci)]
                    if (r, ci) == (mr1, mc1):
                        bg_style_cmds.append(('SPAN', (ci, row_idx), (mc2, row_idx + (mr2 - mr1))))
                # Background
                try:
                    fill = cell.fill
                    if fill.patternType == 'solid':
                        rgb = fill.fgColor.rgb
                        if rgb and rgb not in ('00000000', '0'):
                            bg_style_cmds.append(('BACKGROUND', (ci, row_idx), (ci, row_idx),
                                                  colors.HexColor('#' + rgb[2:])))
                except: pass
                # Border
                try:
                    b = cell.border
                    if any(getattr(b, s).style for s in ['left','right','top','bottom']):
                        border_cmds.append(('BOX', (ci, row_idx), (ci, row_idx), 0.3, colors.HexColor('#D9D9D9')))
                except: pass

        if has_content:
            row_idx += 1

    # ── Build table ──
    t = Table(table_data, colWidths=col_w)
    style_cmds = [
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]
    style_cmds.extend(bg_style_cmds)
    style_cmds.extend(border_cmds)
    t.setStyle(TableStyle(style_cmds))

    story.append(t)

    # 頁尾底線
    story.append(Spacer(1, 5*mm))
    story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#999999')))

    doc.build(story)
    buf.seek(0)
    return buf.read()
