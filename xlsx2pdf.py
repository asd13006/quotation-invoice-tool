"""
純 Python xlsx → PDF 渲染器（同 MiniPdf 做法一樣）
讀取 xlsx 每個 cell 嘅位置、樣式、內容 → 精確渲染 PDF
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

PAGE_W, PAGE_H = A4
MARGIN = 8 * mm

# 註冊 CJK 字型
FONT = 'Helvetica'
FONT_CJK = None

# 方法1：嘗試系統 TTF/TTC 字型
for path, name in [
    ('C:/Windows/Fonts/msjh.ttc', 'MSJH'),
    ('C:/Windows/Fonts/mingliu.ttc', 'MingLiU'),
    ('/System/Library/Fonts/PingFang.ttc', 'PingFang'),
    ('/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc', 'WQY'),
]:
    if os.path.exists(path):
        try:
            pdfmetrics.registerFont(TTFont(name, path))
            FONT = name
            break
        except:
            pass

# 方法2：用 reportlab 內建 CID 字型（唔使外部檔案，Vercel 相容）
if FONT == 'Helvetica':
    try:
        pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
        FONT = 'STSong-Light'
        FONT_CJK = True
    except:
        pass

# 欄闊轉換：openpyxl column width unit ≈ 7pt at 10pt CJK font
CHAR_W = 7.0


def convert_xlsx_to_pdf(xlsx_bytes):
    """xlsx bytes → PDF bytes"""
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active

    buf = io.BytesIO()
    c = pdfcanvas.Canvas(buf, pagesize=A4)

    # ── 讀取欄闊 ──
    col_x = {}  # letter → (x, width) in points
    col_letters = []
    for col_letter, dim in ws.column_dimensions.items():
        if dim.width and col_letter not in col_x:
            col_letters.append(col_letter)
            col_x[col_letter] = dim.width * CHAR_W

    # Sort by column letter
    col_letters.sort()

    # Calculate total width and scale
    total_w = sum(col_x.values())
    scale = (PAGE_W - 2 * MARGIN) / max(total_w, 1)

    # Column positions
    col_pos = {}
    x = MARGIN
    for letter in col_letters:
        col_pos[letter] = x
        x += col_x[letter] * scale

    # ── 字型 size ──
    font_size = min(8, 7 * scale) if scale < 1 else 8

    # ── Render each row ──
    y = PAGE_H - MARGIN

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        r = row[0].row

        # Row height
        row_h = 16  # default pt
        if r in ws.row_dimensions and ws.row_dimensions[r].height:
            row_h = ws.row_dimensions[r].height * 0.75

        # Page break if needed
        if y < MARGIN + 50:
            c.showPage()
            y = PAGE_H - MARGIN

        for cell in row:
            if cell.value is None:
                continue

            letter = cell.column_letter
            if letter not in col_pos:
                continue

            cx = col_pos[letter]
            cw = col_x[letter] * scale
            val = str(cell.value)

            # Skip formulas
            if val.startswith('='):
                continue

            # Cell background
            try:
                fill = cell.fill
                if fill.patternType == 'solid':
                    rgb = fill.fgColor.rgb
                    if rgb and rgb not in ('00000000', '0'):
                        c.setFillColor(colors.HexColor('#' + rgb[2:]))
                        c.rect(cx, y - row_h + 1, cw, row_h, fill=1, stroke=0)
                        c.setFillColor(colors.black)
            except:
                pass

            # Cell border
            try:
                b = cell.border
                has_border = any(getattr(b, s).style for s in ['left', 'right', 'top', 'bottom'])
                if has_border:
                    c.setStrokeColor(colors.HexColor('#D9D9D9'))
                    c.setLineWidth(0.3)
                    c.rect(cx, y - row_h, cw, row_h)
                    c.setStrokeColor(colors.black)
            except:
                pass

            # Cell text
            c.setFont(FONT, font_size)

            # Alignment
            halign = cell.alignment.horizontal or 'left'
            text = val[:100]  # truncate for safety
            text_y = y - row_h + 3

            if halign == 'center':
                c.drawCentredString(cx + cw / 2, text_y, text)
            elif halign == 'right':
                c.drawRightString(cx + cw - 3, text_y, text)
            else:
                c.drawString(cx + 3, text_y, text)

        y -= row_h

    c.save()
    buf.seek(0)
    return buf.read()
