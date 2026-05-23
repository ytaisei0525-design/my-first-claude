#!/usr/bin/env python3
"""Build slides_en.pptx from scratch (matching HTML design, paper figures as placeholders)"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

# ── Color definitions (matching HTML CSS variables) ───────────────────────
INK    = RGBColor(0x1e, 0x29, 0x3b)
INK2   = RGBColor(0x33, 0x41, 0x55)
INK3   = RGBColor(0x47, 0x55, 0x69)
MUTED  = RGBColor(0x64, 0x74, 0x8b)
LINE   = RGBColor(0xe2, 0xe8, 0xf0)
SOFT   = RGBColor(0xf8, 0xfa, 0xfc)
CARD_C = RGBColor(0xf1, 0xf5, 0xf9)
HEADER = RGBColor(0x1e, 0x29, 0x3b)
ACCENT = RGBColor(0x3b, 0x51, 0x74)
AMBER  = RGBColor(0xb0, 0x84, 0x42)
GOLD   = RGBColor(0xd4, 0xa5, 0x74)
WHITE  = RGBColor(0xff, 0xff, 0xff)
PH_BG  = RGBColor(0xf5, 0xf8, 0xfc)
PH_BD  = RGBColor(0xc0, 0xd0, 0xe0)

D_BLUE = RGBColor(0x5b, 0x7a, 0x99)
D_TEAL = RGBColor(0x4a, 0x7c, 0x8c)
D_AMBR = RGBColor(0x9f, 0x76, 0x39)
D_GRN  = RGBColor(0x5b, 0x8c, 0x5a)
D_RED  = RGBColor(0xa0, 0x56, 0x56)
GRAY_L = RGBColor(0xaa, 0xb4, 0xc0)

# ── サイズ ───────────────────────────────────────────────────
W = Inches(13.333)
H = Inches(7.5)
PAD = Inches(0.40)
HDR_H = Inches(0.72)
FTR_H = Inches(0.06)
BODY_Y = HDR_H + Inches(0.22)
BODY_H = H - BODY_Y - FTR_H - Inches(0.20)
BODY_X = PAD
BODY_W = W - PAD * 2


# ── ユーティリティ ───────────────────────────────────────────
def rect(sl, x, y, w, h, fill, border=None, bpt=0.5):
    s = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if border:
        s.line.color.rgb = border
        s.line.width = Pt(bpt)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s

def rrect(sl, x, y, w, h, fill, border=None, bpt=0.5, radius=0.05):
    s = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    s.adjustments[0] = radius
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if border:
        s.line.color.rgb = border
        s.line.width = Pt(bpt)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s

def text(sl, x, y, w, h, content, size=11, bold=False,
         color=INK2, align=PP_ALIGN.LEFT, italic=False, anchor=MSO_ANCHOR.TOP):
    if content is None:
        return None
    s = (str(content)).strip()
    if not s:
        return None
    tb = sl.shapes.add_textbox(x, y, w, h)
    tb.word_wrap = True
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Emu(36000)
    tf.margin_right = Emu(36000)
    tf.margin_top = Emu(18000)
    tf.margin_bottom = Emu(18000)
    for i, line in enumerate(s.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        run.font.name = 'Noto Sans JP'
    return tb

def text_runs(sl, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    """Add multiple formatted runs to a single paragraph"""
    tb = sl.shapes.add_textbox(x, y, w, h)
    tb.word_wrap = True
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Emu(36000); tf.margin_right = Emu(36000)
    tf.margin_top = Emu(18000); tf.margin_bottom = Emu(18000)
    p = tf.paragraphs[0]
    p.alignment = align
    for r in runs:
        run = p.add_run()
        run.text = r['text']
        run.font.size = Pt(r.get('size', 11))
        run.font.bold = r.get('bold', False)
        run.font.italic = r.get('italic', False)
        run.font.color.rgb = r.get('color', INK2)
        run.font.name = 'Noto Sans JP'
    return tb


# ── ヘッダー／フッター描画 ──────────────────────────────────
def draw_header(sl, tag, title):
    rect(sl, 0, 0, W, HDR_H, HEADER)
    rect(sl, 0, HDR_H - Inches(0.04), W, Inches(0.04), ACCENT)
    if tag:
        text(sl, PAD, Inches(0.14), Inches(2.4), Inches(0.4),
             tag, size=11, color=RGBColor(0x88, 0xa4, 0xc4), bold=True,
             anchor=MSO_ANCHOR.MIDDLE)
        # divider line
        rect(sl, Inches(2.85), Inches(0.20), Inches(0.012), Inches(0.32),
             RGBColor(0x44, 0x58, 0x70))
    if title:
        x_t = Inches(3.0) if tag else PAD
        text(sl, x_t, Inches(0.10), W - x_t - PAD, Inches(0.55),
             title, size=24, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)

def draw_footer(sl, num):
    rect(sl, 0, H - FTR_H, W, FTR_H, ACCENT)
    if num:
        text(sl, W - Inches(1.4), H - Inches(0.34), Inches(1.25), Inches(0.26),
             num, size=10, color=MUTED, align=PP_ALIGN.RIGHT)


# ── カード・stats・図プレースホルダー ───────────────────────
def card(sl, x, y, w, h, title=None, body=None, bullets=None,
         tag=None, ct_size=15, cb_size=12, li_size=12, dark=False,
         icon=None, accent_top=True, border_color=None):
    bg = HEADER if dark else CARD_C
    bd = None if dark else (border_color or LINE)
    rrect(sl, x, y, w, h, bg, bd, 0.5, radius=0.04)
    if accent_top and not dark:
        rect(sl, x, y, w, Inches(0.04), ACCENT)

    cy = y + Inches(0.18)
    pad_x = Inches(0.20)
    cw = w - pad_x * 2
    title_c = WHITE if dark else INK
    body_c  = RGBColor(0xc8, 0xd8, 0xe8) if dark else INK3
    tag_c   = RGBColor(0xa0, 0xb8, 0xd0) if dark else MUTED

    if tag:
        text(sl, x + pad_x, cy, cw, Inches(0.24),
             tag, size=10, bold=True, color=tag_c)
        cy += Inches(0.26)

    if title:
        title_t = title
        if icon:
            title_t = f"◆ {title}"
        text(sl, x + pad_x, cy, cw, Inches(0.36),
             title_t, size=ct_size, bold=True, color=title_c)
        cy += Inches(0.40)

    if body:
        text(sl, x + pad_x, cy, cw, h - (cy - y) - Inches(0.12),
             body, size=cb_size, color=body_c)
    elif bullets:
        for b in bullets:
            if cy + Inches(0.28) > y + h - Inches(0.05):
                break
            text(sl, x + pad_x, cy, Inches(0.20), Inches(0.30),
                 "•", size=li_size, color=ACCENT if not dark else GOLD)
            text(sl, x + pad_x + Inches(0.22), cy, cw - Inches(0.22), Inches(0.30),
                 b, size=li_size, color=body_c)
            cy += Inches(0.30)


def stat(sl, x, y, w, h, num, unit=None, label=None, sub=None,
         num_size=44, lbl_size=14, sub_size=11, num_color=INK):
    rrect(sl, x, y, w, h, CARD_C, LINE, 0.5, radius=0.04)
    cy = y + Inches(0.20)

    if unit:
        text_runs(sl, x, cy, w, Inches(0.85), [
            {'text': num, 'size': num_size, 'bold': True, 'color': num_color},
            {'text': unit, 'size': int(num_size * 0.45), 'bold': False, 'color': INK3},
        ], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    else:
        text(sl, x, cy, w, Inches(0.85),
             num, size=num_size, bold=True, color=num_color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cy += Inches(0.90)

    if label:
        text(sl, x + Inches(0.1), cy, w - Inches(0.2), Inches(0.32),
             label, size=lbl_size, bold=True, color=INK2, align=PP_ALIGN.CENTER)
        cy += Inches(0.32)
    if sub:
        text(sl, x + Inches(0.1), cy, w - Inches(0.2), Inches(0.30),
             sub, size=sub_size, color=MUTED, align=PP_ALIGN.CENTER)


def callout(sl, x, y, w, h, content, dark=False, icon='▶'):
    if dark:
        rrect(sl, x, y, w, h, HEADER, None, radius=0.04)
        rect(sl, x, y, Inches(0.06), h, GOLD)
        text(sl, x + Inches(0.20), y, Inches(0.40), h,
             icon, size=14, color=GOLD, anchor=MSO_ANCHOR.MIDDLE)
        text(sl, x + Inches(0.65), y, w - Inches(0.75), h,
             content, size=13, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    else:
        rrect(sl, x, y, w, h, SOFT, None, radius=0.04)
        rect(sl, x, y, Inches(0.06), h, ACCENT)
        text(sl, x + Inches(0.20), y, Inches(0.40), h,
             icon, size=14, color=ACCENT, anchor=MSO_ANCHOR.MIDDLE)
        text(sl, x + Inches(0.65), y, w - Inches(0.75), h,
             content, size=13, color=INK, bold=True, anchor=MSO_ANCHOR.MIDDLE)


def fig_placeholder(sl, x, y, w, h, fig_num=None, caption=None):
    """Placeholder for paper figures (user will paste their own images later)"""
    # Background (solid border, light color)
    s = rrect(sl, x, y, w, h, PH_BG, PH_BD, 1.0, radius=0.02)
    # Figure number badge (top-left)
    if fig_num:
        tag_w = Inches(0.95)
        rect(sl, x + Inches(0.12), y + Inches(0.12), tag_w, Inches(0.32),
             HEADER)
        text(sl, x + Inches(0.12), y + Inches(0.12), tag_w, Inches(0.32),
             fig_num, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Center placeholder text
    text(sl, x, y, w, h,
         "(Paste paper figure here)", size=12, color=GRAY_L,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Caption
    if caption:
        text(sl, x, y + h + Inches(0.04), w, Inches(0.30),
             caption, size=10, color=MUTED, italic=True,
             align=PP_ALIGN.CENTER)


def section_label(sl, x, y, w, label):
    """Label text (meta information above a group of cards)"""
    text(sl, x, y, w, Inches(0.30),
         label, size=11, bold=True, color=MUTED)


def mini(sl, x, y, w, h, label, value, sub=None, value_color=INK):
    rrect(sl, x, y, w, h, SOFT, LINE, 0.5, radius=0.03)
    cy = y + Inches(0.12)
    text(sl, x + Inches(0.14), cy, w - Inches(0.28), Inches(0.25),
         label, size=10, color=MUTED, bold=True)
    cy += Inches(0.28)
    text(sl, x + Inches(0.14), cy, w - Inches(0.28), Inches(0.45),
         value, size=20, bold=True, color=value_color)
    if sub:
        text(sl, x + Inches(0.14), cy + Inches(0.46), w - Inches(0.28), Inches(0.24),
             sub, size=9, color=MUTED)


def bar_row(sl, x, y, w, label, fill_ratio, value, fill_color=ACCENT,
            lbl_w=Inches(2.0), val_w=Inches(1.1)):
    text(sl, x, y, lbl_w, Inches(0.30),
         label, size=12, color=INK2, anchor=MSO_ANCHOR.MIDDLE)
    bx = x + lbl_w + Inches(0.10)
    bw = w - lbl_w - val_w - Inches(0.20)
    # Background bar
    rrect(sl, bx, y + Inches(0.08), bw, Inches(0.16),
          CARD_C, LINE, 0.3, radius=0.5)
    # Fill
    fill_w = bw * fill_ratio
    if fill_w > Inches(0.05):
        rrect(sl, bx, y + Inches(0.08), fill_w, Inches(0.16),
              fill_color, None, radius=0.5)
    # Value
    text(sl, bx + bw + Inches(0.10), y, val_w, Inches(0.30),
         value, size=11, color=INK3, bold=True, anchor=MSO_ANCHOR.MIDDLE)


def simple_table(sl, x, y, w, h, headers, rows, col_widths=None):
    """Simple table"""
    n_cols = len(headers)
    if col_widths is None:
        col_widths = [w / n_cols] * n_cols
    else:
        # Ratio specification
        total = sum(col_widths)
        col_widths = [w * (cw / total) for cw in col_widths]

    n_rows = len(rows) + 1
    row_h = h / n_rows

    # Header
    rect(sl, x, y, w, row_h, HEADER)
    cx = x
    for i, hd in enumerate(headers):
        text(sl, cx + Inches(0.10), y, col_widths[i] - Inches(0.10), row_h,
             hd, size=11, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[i]

    # Data rows
    for ri, row in enumerate(rows):
        ry = y + row_h * (ri + 1)
        bg = SOFT if ri % 2 == 0 else WHITE
        rect(sl, x, ry, w, row_h, bg, LINE, 0.3)
        cx = x
        for ci, cell in enumerate(row):
            cell_color = INK2
            cell_text = str(cell)
            cell_bold = False
            if isinstance(cell, dict):
                cell_text = cell.get('text', '')
                cell_color = cell.get('color', INK2)
                cell_bold = cell.get('bold', False)
            text(sl, cx + Inches(0.10), ry, col_widths[ci] - Inches(0.10), row_h,
                 cell_text, size=11, color=cell_color, bold=cell_bold,
                 anchor=MSO_ANCHOR.MIDDLE)
            cx += col_widths[ci]


def flow_step(sl, x, y, w, h, n, title, desc, dark=False, arrow=True):
    bg = HEADER if dark else CARD_C
    bd = None if dark else LINE
    title_c = WHITE if dark else INK
    desc_c = RGBColor(0xc0,0xd0,0xe0) if dark else INK3
    n_c = RGBColor(0xa0,0xb8,0xd0) if dark else MUTED
    rrect(sl, x, y, w, h, bg, bd, 0.5, radius=0.04)
    text(sl, x + Inches(0.14), y + Inches(0.14), w - Inches(0.28), Inches(0.24),
         n, size=10, bold=True, color=n_c)
    text(sl, x + Inches(0.14), y + Inches(0.40), w - Inches(0.28), Inches(0.34),
         title, size=14, bold=True, color=title_c)
    text(sl, x + Inches(0.14), y + Inches(0.76), w - Inches(0.28), h - Inches(0.85),
         desc, size=11, color=desc_c)


# ── スライド構築 ───────────────────────────────────────────
def new_slide(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    return sl


# ===== SLIDE 1: COVER =====
def slide_01_cover(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, HEADER)
    rect(sl, 0, 0, Inches(0.12), H, ACCENT)
    rect(sl, 0, H - Inches(0.07), W, Inches(0.07), ACCENT)

    # Top tag
    text(sl, Inches(0.5), Inches(0.5), Inches(8), Inches(0.4),
         "RESEARCH PRESENTATION  ·  2024", size=11, bold=True,
         color=RGBColor(0x88, 0x9a, 0xb0))
    text(sl, W - Inches(5.5), Inches(0.5), Inches(5), Inches(0.4),
         "Adv. Mater.  ·  Dong et al., 2024", size=11,
         color=RGBColor(0x88, 0x9a, 0xb0), align=PP_ALIGN.RIGHT)

    # Title
    text(sl, Inches(0.8), Inches(1.7), Inches(11.7), Inches(0.45),
         "WATER-ENHANCING GELS  ·  ADVANCED MATERIALS 2024", size=12, bold=True,
         color=RGBColor(0x88, 0x9a, 0xb0), align=PP_ALIGN.CENTER)

    text(sl, Inches(0.5), Inches(2.3), Inches(12.3), Inches(1.2),
         "Protecting Against", size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    text(sl, Inches(0.5), Inches(3.2), Inches(12.3), Inches(1.0),
         "Heat-Activated Silica Aerogels", size=44, bold=True, color=GOLD, align=PP_ALIGN.CENTER)

    text(sl, Inches(1.0), Inches(4.5), Inches(11.3), Inches(1.0),
         "Next-generation fire-retardant gel of cellulosic polymers × colloidal silica particles\nthat self-foams upon heating to form an aerogel insulating layer",
         size=15, color=RGBColor(0xb8, 0xcc, 0xe2), align=PP_ALIGN.CENTER)

    # Meta information
    meta = "Materials Chemistry / Fire Protection Eng.    ·    Stanford University, Appel Lab    ·    Adv. Mater. 36, 2407375"
    text(sl, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.4),
         meta, size=12, color=RGBColor(0x88, 0x9c, 0xb4), align=PP_ALIGN.CENTER)

    text(sl, Inches(0.5), Inches(6.6), Inches(8), Inches(0.3),
         "DOI: 10.1002/adma.202407375", size=10, color=RGBColor(0x66, 0x77, 0x88))
    text(sl, W - Inches(1.4), H - Inches(0.36), Inches(1.2), Inches(0.26),
         "1 / 36", size=10, color=RGBColor(0x66, 0x77, 0x88), align=PP_ALIGN.RIGHT)


# ===== SLIDE 2: TOC =====
def slide_02_toc(prs):
    sl = new_slide(prs)
    draw_header(sl, "Contents", "Presentation Outline")
    draw_footer(sl, "2 / 36")

    parts = [
        ("PART 1", "Background", "Current state of wildfires, limitations of existing technologies (Phos-Chek, AquaGel-K), and innovation of this research", "Slides 4–7"),
        ("PART 2", "Materials & Methods", "Cellulosic polymers, CSP, formulations (5 types), evaluation methods", "Slides 9–14"),
        ("PART 3", "Results", "Rheology, combustion tests, foaming index, SEM observation, mechanism", "Slides 16–33"),
        ("PART 4", "Discussion & Conclusion", "Implementation scenarios, comparison with existing products, conclusions and future prospects", "Slides 35–36"),
    ]
    cw = (BODY_W - Inches(0.42)) / 4
    cy = BODY_Y
    ch = Inches(4.2)
    for i, (tag, title, body, sl_range) in enumerate(parts):
        cx = BODY_X + i * (cw + Inches(0.14))
        rrect(sl, cx, cy, cw, ch, CARD_C, LINE, 0.5, radius=0.04)
        rect(sl, cx, cy, cw, Inches(0.04), ACCENT)
        # Icon (using text glyph)
        text(sl, cx + Inches(0.20), cy + Inches(0.25), cw - Inches(0.4), Inches(0.5),
             "●", size=22, color=ACCENT)
        text(sl, cx + Inches(0.20), cy + Inches(0.85), cw - Inches(0.4), Inches(0.30),
             tag, size=11, bold=True, color=MUTED)
        text(sl, cx + Inches(0.20), cy + Inches(1.20), cw - Inches(0.4), Inches(0.45),
             title, size=18, bold=True, color=INK)
        text(sl, cx + Inches(0.20), cy + Inches(1.75), cw - Inches(0.4), Inches(1.8),
             body, size=12, color=INK3)
        text(sl, cx + Inches(0.20), cy + ch - Inches(0.45), cw - Inches(0.4), Inches(0.30),
             sl_range, size=10, color=MUTED)

    co_y = cy + ch + Inches(0.18)
    callout(sl, BODY_X, co_y, BODY_W, Inches(0.55),
            "Key Message: Foaming upon heating + sintering of silica particles spontaneously forms an aerogel layer, achieving longer protection than commercial products",
            icon='ⓘ')


# ===== SLIDE 3: SECTION DIVIDER 01 =====
def slide_section_divider(prs, num_str, num_label, title, sub, pills, page_num):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, HEADER)
    rect(sl, 0, 0, Inches(0.10), H, AMBER)

    # Large background number
    text(sl, W - Inches(5.5), Inches(0.5), Inches(5), Inches(7.0),
         num_str, size=240, bold=True, color=RGBColor(0x24, 0x30, 0x44),
         align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.BOTTOM)

    text(sl, Inches(1.0), Inches(1.7), Inches(10), Inches(0.4),
         num_label, size=13, bold=True, color=RGBColor(0x88, 0xa4, 0xc0))
    text(sl, Inches(1.0), Inches(2.2), Inches(10), Inches(1.4),
         title, size=52, bold=True, color=WHITE)
    text(sl, Inches(1.0), Inches(3.7), Inches(10), Inches(1.0),
         sub, size=17, color=RGBColor(0x9c, 0xb4, 0xcc))

    # Pills
    px = Inches(1.0)
    py = Inches(5.0)
    for p in pills:
        pw = Inches(max(len(p) * 0.16 + 0.5, 1.6))
        rrect(sl, px, py, pw, Inches(0.42),
              RGBColor(0x2a, 0x38, 0x4e), RGBColor(0x44, 0x58, 0x70), 0.5, radius=0.4)
        text(sl, px, py, pw, Inches(0.42),
             p, size=12, color=RGBColor(0xc0, 0xd0, 0xe0),
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        px += pw + Inches(0.16)

    text(sl, W - Inches(1.4), H - Inches(0.36), Inches(1.2), Inches(0.26),
         page_num, size=10, color=RGBColor(0x66, 0x77, 0x88), align=PP_ALIGN.RIGHT)


# ===== SLIDE 4: 山火事の現状 =====
def slide_04_wildfire(prs):
    sl = new_slide(prs)
    draw_header(sl, "Background 1/4", "Growing Severity of Wildfire Damage and WUI Risk")
    draw_footer(sl, "4 / 36")

    # Left: stats + cards, Right: figure placeholder
    left_w = BODY_W * 0.62 - Inches(0.10)
    right_w = BODY_W * 0.38 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    # 3 stats
    sw = (left_w - Inches(0.20)) / 3
    sy = BODY_Y
    sh = Inches(1.7)
    stat(sl, BODY_X, sy, sw, sh, "5", unit="×",
         label="Increase in Large-Scale Wildfires", sub="Burned area expansion over past 30 years",
         num_size=46, lbl_size=12, sub_size=10)
    stat(sl, BODY_X + sw + Inches(0.10), sy, sw, sh, "400", unit="M km²",
         label="Annual Burned Area (Global)", sub="Caused by climate change and drought",
         num_size=40, lbl_size=12, sub_size=10)
    stat(sl, BODY_X + (sw + Inches(0.10)) * 2, sy, sw, sh, "45M",
         label="U.S. WUI Residents", sub="High-risk population in interface zone",
         num_size=42, lbl_size=12, sub_size=10)

    # WUI terminology card
    ty = sy + sh + Inches(0.20)
    th = Inches(1.0)
    rrect(sl, BODY_X, ty, left_w, th, WHITE, LINE, 0.5, radius=0.04)
    text(sl, BODY_X + Inches(0.20), ty + Inches(0.14), left_w - Inches(0.4), Inches(0.30),
         "Terminology", size=11, bold=True, color=MUTED)
    text(sl, BODY_X + Inches(0.20), ty + Inches(0.42), left_w - Inches(0.4), Inches(0.55),
         "WUI (Wildland-Urban Interface): The boundary zone between wildland and urban areas. A high-risk region where wildfire damage directly threatens residential structures.",
         size=13, color=INK)

    # Callout
    cy = ty + th + Inches(0.18)
    callout(sl, BODY_X, cy, left_w, Inches(0.65),
            "Climate change extends dry and high-wind periods → Conventional firefighting strategies cannot keep pace",
            icon='⚠')

    # Right: figure placeholder
    fig_h = Inches(4.6)
    fig_placeholder(sl, right_x, BODY_Y, right_w, fig_h,
                    fig_num="Fig. 3c",
                    caption="Combustion process under direct flame contact")


# ===== SLIDE 5: 4段階の保護プロセス =====
def slide_05_process(prs):
    sl = new_slide(prs)
    draw_header(sl, "Background 2/4", "Research Approach: 4-Stage Protection Process")
    draw_footer(sl, "5 / 36")

    # Top: large figure placeholder
    fig_h = Inches(3.0)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 1a",
                    caption="Fire-retardant gel application → flame contact → heat-activated aerogel formation → structural protection")

    # Bottom: 4-step cards
    sy = BODY_Y + fig_h + Inches(0.45)
    sh = Inches(2.2)
    cw = (BODY_W - Inches(0.42)) / 4
    steps = [
        ("STEP 1", "Gel Application", "Spray application of cellulosic gel to buildings and vegetation"),
        ("STEP 2", "Flame Contact", "Moisture evaporates while gel foams and particles aggregate"),
        ("STEP 3", "Aerogel Formation", "Silica particles sinter to form a porous aerogel layer"),
        ("STEP 4", "Sustained Protection", "Ultra-low thermal conductivity insulating layer continuously protects the substrate"),
    ]
    for i, (tag, title, body) in enumerate(steps):
        cx = BODY_X + i * (cw + Inches(0.14))
        card(sl, cx, sy, cw, sh, title=title, body=body, tag=tag,
             ct_size=16, cb_size=12)


# ===== SLIDE 6: 既存技術と限界 =====
def slide_06_existing(prs):
    sl = new_slide(prs)
    draw_header(sl, "Background 3/4", "Existing Wildfire Protection Methods and Their Fundamental Limitations")
    draw_footer(sl, "6 / 36")

    # Left: table, Right: figure
    left_w = BODY_W * 0.66 - Inches(0.10)
    right_w = BODY_W * 0.34 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    headers = ["Technology", "Main Component", "Feature", "Limitation"]
    rows = [
        ["Water only", "H₂O", "Cooling by latent heat of evaporation",
         {'text': "Immediately runs off and evaporates", 'color': D_RED}],
        ["Phos-Chek", "Ammonium phosphate", "Direct application to trees and ground surface",
         {'text': "Residual contamination in soil and water systems", 'color': D_RED}],
        ["AquaGel-K", "Cross-linked polyacrylate", "Retains water at high concentration",
         {'text': "Function lost after water evaporation", 'color': D_RED}],
        [{'text': "This Work WEG", 'bold': True, 'color': ACCENT},
         "Cellulosic + Silica", "Forms aerogel layer upon heating",
         {'text': "Continues protecting after water evaporation", 'color': D_GRN, 'bold': True}],
    ]
    table_h = Inches(3.2)
    simple_table(sl, BODY_X, BODY_Y, left_w, table_h, headers, rows,
                 col_widths=[1.5, 2, 2, 1.5])

    # Callout
    cy = BODY_Y + table_h + Inches(0.22)
    callout(sl, BODY_X, cy, left_w, Inches(0.70),
            "Challenge: Existing WEGs are mere \"water carriers\" — protection is lost simultaneously with moisture loss",
            dark=True, icon='→')

    # Right: figure placeholder
    fig_placeholder(sl, right_x, BODY_Y, right_w, Inches(4.2),
                    fig_num="Fig. 3d",
                    caption="Comparison at 300 s: Water vs. This Work WEG")


# ===== SLIDE 7: 研究の核心 =====
def slide_07_core(prs):
    sl = new_slide(prs)
    draw_header(sl, "Background 4/4", "Research Core: Heat-Activated Aerogel Formation")
    draw_footer(sl, "7 / 36")

    left_w = BODY_W * 0.60 - Inches(0.10)
    right_w = BODY_W * 0.40 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    # Left: 2 cards
    ch = Inches(3.4)
    cw_l = (left_w - Inches(0.14)) / 2
    card(sl, BODY_X, BODY_Y, cw_l, ch,
         title="Design Concept", tag="DESIGN CONCEPT",
         bullets=[
             "Cellulosic polymers (HEC, MC, MHEC) form water-soluble gel network",
             "Colloidal silica particles (CSP) uniformly dispersed",
             "Water foams upon heating to form porous structure",
             "Silica particles sinter and convert to aerogel layer",
         ], ct_size=15, li_size=11)
    card(sl, BODY_X + cw_l + Inches(0.14), BODY_Y, cw_l, ch,
         title="Three Innovation Points", tag="INNOVATION",
         dark=True,
         bullets=[
             "Breaking free from water dependence: heat-activated solid protective layer",
             "Sustainable: naturally derived, food-grade materials",
             "Compatible with existing spray infrastructure: sprayable fluid",
         ], ct_size=15, li_size=11)

    # Bottom callout
    cy_co = BODY_Y + ch + Inches(0.25)
    callout(sl, BODY_X, cy_co, left_w, Inches(0.80),
            '"Heat-Activated Formation of Silica Aerogels"\n— Triggered by flame contact, the gel itself transforms into an aerogel insulator',
            icon='💡')

    # Right: figure placeholder
    fig_placeholder(sl, right_x, BODY_Y, right_w, Inches(4.8),
                    fig_num="Fig. 5a",
                    caption="3-stage silica aerogel formation (preview)")


# ===== SLIDE 9: セルロース系ポリマー =====
def slide_09_polymers(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials 1/6", "Cellulosic Polymers Used")
    draw_footer(sl, "9 / 36")

    # Top: 3 polymer cards
    pw = (BODY_W - Inches(0.28)) / 3
    ph = Inches(2.1)
    polymers = [
        ("Polymer A", "HEC", "Hydroxyethyl cellulose", "→ Base viscosity agent"),
        ("Polymer B", "MC", "Methyl cellulose", "→ Gelation upon heating (thermoreversible)"),
        ("Polymer C", "MHEC", "Methyl 2-hydroxyethyl cellulose", "→ Integrates properties of A and B"),
    ]
    for i, (tag, name, full, role) in enumerate(polymers):
        cx = BODY_X + i * (pw + Inches(0.14))
        rrect(sl, cx, BODY_Y, pw, ph, SOFT, LINE, 0.5, radius=0.04)
        text(sl, cx + Inches(0.20), BODY_Y + Inches(0.14), pw - Inches(0.4), Inches(0.25),
             tag, size=11, bold=True, color=MUTED)
        text(sl, cx + Inches(0.20), BODY_Y + Inches(0.42), pw - Inches(0.4), Inches(0.50),
             name, size=26, bold=True, color=INK)
        text(sl, cx + Inches(0.20), BODY_Y + Inches(0.96), pw - Inches(0.4), Inches(0.65),
             full, size=11, color=INK3)
        text(sl, cx + Inches(0.20), BODY_Y + ph - Inches(0.42), pw - Inches(0.4), Inches(0.30),
             role, size=12, bold=True, color=ACCENT)

    # Bottom: figure placeholder
    fy = BODY_Y + ph + Inches(0.30)
    fh = BODY_H - ph - Inches(0.30)
    fig_placeholder(sl, BODY_X, fy, BODY_W, fh - Inches(0.35),
                    fig_num="Fig. 1b/c",
                    caption="Chemical structures of each polymer and mixing scheme with colloidal silica (CSP)")


# ===== SLIDE 10: MC熱ゲル化 =====
def slide_10_mc(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials 2/6", "Methyl Cellulose (MC) Thermal Gelation: Key Property for Flame Protection")
    draw_footer(sl, "10 / 36")

    # Left: description card, Right: temperature process
    left_w = BODY_W * 0.48 - Inches(0.10)
    right_w = BODY_W * 0.52 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    # Left card
    ch_l = Inches(3.5)
    card(sl, BODY_X, BODY_Y, left_w, ch_l,
         title="Inverse Thermal Response (LCST Behavior)", tag="THERMAL RESPONSE",
         bullets=[
             "MC dissolves in cold water (<20°C) to form a low-viscosity solution",
             "Gelation above LCST upon heating",
             "MC gelation onset: approx. 50–60°C (concentration-dependent)",
             "Liquid at room temperature → instantaneous gelation upon flame contact",
         ], ct_size=15, li_size=12)
    callout(sl, BODY_X, BODY_Y + ch_l + Inches(0.20), left_w, Inches(0.85),
            "MC gelates in the initial stage of flame contact → maintains foamed structure while CSP is fixed",
            icon='🔥')

    # Right: 3 mini stats
    mw = (right_w - Inches(0.20)) / 3
    mh = Inches(1.4)
    mini(sl, right_x, BODY_Y, mw, mh, "LCST (Gelation onset)", "~55°C", "1 wt% MC aqueous solution", value_color=ACCENT)
    mini(sl, right_x + mw + Inches(0.10), BODY_Y, mw, mh, "Thermal decomposition", "~300°C", "TGA measurement", value_color=D_RED)
    mini(sl, right_x + (mw + Inches(0.10)) * 2, BODY_Y, mw, mh, "CSP sintering onset", "~200°C", "Inter-particle neck formation", value_color=D_TEAL)

    # Temperature process flow (4 steps)
    fy = BODY_Y + mh + Inches(0.30)
    fh = Inches(2.7)
    # Dark card background
    rrect(sl, right_x, fy, right_w, fh, HEADER, None, radius=0.04)
    text(sl, right_x + Inches(0.20), fy + Inches(0.18), right_w - Inches(0.4), Inches(0.30),
         "Temperature-dependent Process", size=11, bold=True, color=RGBColor(0xa0, 0xb8, 0xd0))

    sub_y = fy + Inches(0.65)
    sub_h = fh - Inches(0.75)
    pw = (right_w - Inches(0.40) - Inches(0.30) * 3) / 4
    steps_t = [
        ("~55°C", "MC gelation", "Structure fixed"),
        ("~100°C", "Moisture evaporation", "Foaming & expansion"),
        ("~200°C", "CSP sintering", "Grain boundary formation"),
        (">300°C", "Organic decomposition", "Pure silica layer"),
    ]
    px = right_x + Inches(0.20)
    for i, (temp, ti, dsc) in enumerate(steps_t):
        rrect(sl, px, sub_y, pw, sub_h, RGBColor(0x2c, 0x3b, 0x52), None, radius=0.04)
        text(sl, px + Inches(0.10), sub_y + Inches(0.14), pw - Inches(0.2), Inches(0.30),
             temp, size=11, bold=True, color=GOLD)
        text(sl, px + Inches(0.10), sub_y + Inches(0.50), pw - Inches(0.2), Inches(0.40),
             ti, size=13, bold=True, color=WHITE)
        text(sl, px + Inches(0.10), sub_y + Inches(1.0), pw - Inches(0.2), Inches(0.50),
             dsc, size=11, color=RGBColor(0xb0, 0xc0, 0xd0))
        px += pw + Inches(0.10)
        if i < 3:
            text(sl, px - Inches(0.10), sub_y, Inches(0.10), sub_h,
                 "▶", size=12, color=RGBColor(0x66, 0x78, 0x90),
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ===== SLIDE 11: CSP & SDS =====
def slide_11_csp_sds(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials 3/6", "Colloidal Silica Particles (CSP) and Surfactant (SDS)")
    draw_footer(sl, "11 / 36")

    cw = (BODY_W - Inches(0.18)) / 2
    ch = Inches(4.2)

    # Left: CSP
    cx = BODY_X
    rrect(sl, cx, BODY_Y, cw, ch, CARD_C, LINE, 0.5, radius=0.04)
    rect(sl, cx, BODY_Y, cw, Inches(0.04), ACCENT)
    text(sl, cx + Inches(0.20), BODY_Y + Inches(0.20), Inches(0.6), Inches(0.5),
         "●", size=28, color=ACCENT)
    text(sl, cx + Inches(0.95), BODY_Y + Inches(0.22), cw - Inches(1.1), Inches(0.3),
         "CSP", size=11, bold=True, color=MUTED)
    text(sl, cx + Inches(0.95), BODY_Y + Inches(0.50), cw - Inches(1.1), Inches(0.5),
         "Colloidal Silica Particles", size=18, bold=True, color=INK)
    bullets = [
        "Stock solution: colloidal dispersion at 12 wt% in water",
        "Particle size: monodisperse silica of tens of nanometers",
        "Formulation concentration: 5 wt% in gel",
        "Sinters upon heating → transforms into silica aerogel",
    ]
    by = BODY_Y + Inches(1.30)
    for b in bullets:
        text(sl, cx + Inches(0.30), by, Inches(0.20), Inches(0.30),
             "•", size=14, color=ACCENT)
        text(sl, cx + Inches(0.55), by, cw - Inches(0.75), Inches(0.30),
             b, size=13, color=INK2)
        by += Inches(0.42)

    # Right: SDS
    cx2 = BODY_X + cw + Inches(0.18)
    rrect(sl, cx2, BODY_Y, cw, ch, CARD_C, LINE, 0.5, radius=0.04)
    rect(sl, cx2, BODY_Y, cw, Inches(0.04), ACCENT)
    text(sl, cx2 + Inches(0.20), BODY_Y + Inches(0.20), Inches(0.6), Inches(0.5),
         "◆", size=28, color=ACCENT)
    text(sl, cx2 + Inches(0.95), BODY_Y + Inches(0.22), cw - Inches(1.1), Inches(0.3),
         "SDS (Additive)", size=11, bold=True, color=MUTED)
    text(sl, cx2 + Inches(0.95), BODY_Y + Inches(0.50), cw - Inches(1.1), Inches(0.5),
         "Sodium Dodecyl Sulfate", size=18, bold=True, color=INK)
    bullets2 = [
        "Anionic surfactant",
        "Role: promotes foaming and enhances porous structure",
        "Addition levels: two concentrations of 0.1 wt% and 0.5 wt%",
        "SEM confirms control of bubble size and distribution",
    ]
    by = BODY_Y + Inches(1.30)
    for b in bullets2:
        text(sl, cx2 + Inches(0.30), by, Inches(0.20), Inches(0.30),
             "•", size=14, color=ACCENT)
        text(sl, cx2 + Inches(0.55), by, cw - Inches(0.75), Inches(0.30),
             b, size=13, color=INK2)
        by += Inches(0.42)

    # Bottom callout
    cy_co = BODY_Y + ch + Inches(0.30)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.80),
            "The sintering temperature range of CSP overlaps with the thermal decomposition temperature of cellulose, enabling synchronized thermal response to form the insulating layer",
            icon='💡')


# ===== SLIDE 12: 配合系 5種 =====
def slide_12_formulations(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials 4/6", "Five Formulations Evaluated")
    draw_footer(sl, "12 / 36")

    # 5 cards side by side
    fw = (BODY_W - Inches(0.4 * 4)) / 5
    fh = Inches(1.9)
    forms = [
        ("Control", "AquaGel-K", "0.5 wt%\nCommercial WEG benchmark", False),
        ("Formula 1", "HEC+MC / CSP", "HEC+MC 1 wt%\nCSP 5 wt%", True),
        ("Formula 2", "MHEC / CSP", "MHEC 1 wt%\nCSP 5 wt%", True),
        ("Formula 3", "HEC+MC / CSP / SDS", "+ SDS 0.1 wt%", True),
        ("Formula 4", "HEC+MC / CSP / SDS", "+ SDS 0.5 wt%", True),
    ]
    fy = BODY_Y
    for i, (tag, name, comp, hl) in enumerate(forms):
        fx = BODY_X + i * (fw + Inches(0.10))
        bg = SOFT if hl else WHITE
        bd = ACCENT if hl else LINE
        rrect(sl, fx, fy, fw, fh, bg, bd, 0.7 if hl else 0.5, radius=0.04)
        text(sl, fx + Inches(0.10), fy + Inches(0.16), fw - Inches(0.2), Inches(0.28),
             tag, size=10, bold=True, color=ACCENT,
             align=PP_ALIGN.CENTER)
        text(sl, fx + Inches(0.10), fy + Inches(0.55), fw - Inches(0.2), Inches(0.60),
             name, size=13, bold=True, color=INK,
             align=PP_ALIGN.CENTER)
        text(sl, fx + Inches(0.10), fy + Inches(1.20), fw - Inches(0.2), Inches(0.60),
             comp, size=10, color=MUTED, align=PP_ALIGN.CENTER)

    # Bottom 2 cards
    cy = fy + fh + Inches(0.40)
    ch = Inches(2.5)
    cw2 = (BODY_W - Inches(0.20)) / 2
    card(sl, BODY_X, cy, cw2, ch,
         title="Formulation Design Intent",
         bullets=[
             "Comparing functional differences by varying polymer type (HEC+MC vs MHEC)",
             "Verifying the effect on foam structure with two SDS concentration levels",
             "Direct comparison with commercial AquaGel-K",
         ], ct_size=15, li_size=12)
    card(sl, BODY_X + cw2 + Inches(0.20), cy, cw2, ch,
         title="Notation Guide",
         body="HEC+MC/CSP/SDS 1-5-0.1\n→ HEC+MC 1 wt% ／ CSP 5 wt% ／ SDS 0.1 wt%",
         ct_size=15, cb_size=14)


# ===== SLIDE 13: 評価手法 =====
def slide_13_methods(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials 5/6", "Overview of Evaluation Methods")
    draw_footer(sl, "13 / 36")

    methods = [
        ("Rheological Measurements", "RHEOLOGY",
         ["Oscillatory frequency sweep (G', G'')", "Steady-flow sweep (viscosity vs. shear rate)", "Herschel-Bulkley model fitting"]),
        ("Combustion Test (Time-to-char)", "BURN TEST",
         ["Wood heated with butane burner", "Time to charring onset measured", "Photographic comparison at 120 s / 300 s"]),
        ("Foaming Index Measurement", "FOAMING",
         ["Foamed layer thickness measured after combustion", "Ratio to initial thickness = Foaming Index"]),
        ("SEM Morphological Observation", "SEM",
         ["Foam structure by SDS concentration", "Sintering progression by burn time (0/1/2/4 min)"]),
        ("Spectroscopic Analysis", "SPECTROSCOPY",
         ["FT-IR (chemical bonding)", "XPS (surface composition)"]),
        ("Thermal Analysis", "THERMAL",
         ["TGA (thermogravimetric analysis)", "DSC (differential scanning calorimetry)"]),
    ]
    cw = (BODY_W - Inches(0.30)) / 3
    ch = (BODY_H - Inches(0.20)) / 2
    for i, (title, tag, bullets) in enumerate(methods):
        col = i % 3
        row = i // 3
        cx = BODY_X + col * (cw + Inches(0.15))
        cy = BODY_Y + row * (ch + Inches(0.20))
        card(sl, cx, cy, cw, ch, title=title, tag=tag,
             bullets=bullets, ct_size=14, li_size=11)


# ===== SLIDE 14: 付着性・濡れ性 =====
def slide_14_adhesion(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials 6/6", "Adhesion and Surface Wettability: Properties Essential for Implementation")
    draw_footer(sl, "14 / 36")

    # 3 cards
    ch = Inches(3.6)
    cw = (BODY_W - Inches(0.28)) / 3
    cards_data = [
        ("Contact Angle Measurement", "WETTING", [
            "Contact angle evaluation on wood, concrete, and metal",
            "WEG shows low contact angle → high wettability",
            "HEC+MC shows particularly high affinity at cellulose-wood interface",
        ]),
        ("Vertical Surface Adhesion Test", "ADHESION", [
            "Gel applied to vertically oriented substrate",
            "G' ≫ G'' → resists gravity without flowing down",
            "Vertical surface retention equal to or better than AquaGel-K",
        ]),
        ("Spray Application Suitability", "SPRAY", [
            "Viscosity drops sharply under high shear (inside nozzle)",
            "Quickly recovers high viscosity upon reaching substrate → maintains adhesion",
            "Fully compatible with existing fire hoses and nozzles",
        ]),
    ]
    for i, (title, tag, bullets) in enumerate(cards_data):
        cx = BODY_X + i * (cw + Inches(0.14))
        card(sl, cx, BODY_Y, cw, ch, title=title, tag=tag,
             bullets=bullets, ct_size=15, li_size=12)

    # Bottom callout
    cy_co = BODY_Y + ch + Inches(0.30)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.80),
            "Rheological design (low n, high G') simultaneously achieves spray suitability and adhesion retention",
            icon='✓')


# ===== SLIDE 16: G'/G'' =====
def slide_16_rheology1(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 1/11", "Rheology ①: Oscillatory Moduli G' / G''")
    draw_footer(sl, "16 / 36")

    # Left: figure, Right: observations
    left_w = BODY_W * 0.58 - Inches(0.10)
    right_w = BODY_W * 0.42 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    fig_h = Inches(4.8)
    fig_placeholder(sl, BODY_X, BODY_Y, left_w, fig_h,
                    fig_num="Fig. 2a–b",
                    caption="Angular frequency vs Storage modulus G' · Loss modulus G''")

    # Right: observation card
    ch = Inches(2.8)
    card(sl, right_x, BODY_Y, right_w, ch,
         title="Key Observation", tag="KEY OBSERVATION",
         bullets=[
             "G' > G'' for all formulations → solid-like (gel) behavior",
             "Low frequency dependence → stable network",
             "This work's WEG shows higher elasticity than AquaGel-K",
         ], ct_size=15, li_size=12)
    callout(sl, right_x, BODY_Y + ch + Inches(0.25), right_w, Inches(0.80),
            "Clearly formed gel network; does not flow off after application",
            icon='"')


# ===== SLIDE 17: shear thinning =====
def slide_17_rheology2(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 2/11", "Rheology ②: Shear Thinning and Power-law Index")
    draw_footer(sl, "17 / 36")

    left_w = BODY_W * 0.58 - Inches(0.10)
    right_w = BODY_W * 0.42 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    fig_h = Inches(4.8)
    fig_placeholder(sl, BODY_X, BODY_Y, left_w, fig_h,
                    fig_num="Fig. 2c–d",
                    caption="Viscosity vs. shear rate (shear-thinning behavior)")

    # Right: mini stats + explanation
    mh = Inches(1.20)
    mini(sl, right_x, BODY_Y, right_w, mh,
         "Power-law index n", "0.108", "HEC+MC/CSP 1-5", value_color=ACCENT)
    mini(sl, right_x, BODY_Y + mh + Inches(0.15), right_w, mh,
         "Power-law index n", "0.188", "MHEC/CSP 1-5", value_color=ACCENT)

    cy_d = BODY_Y + (mh + Inches(0.15)) * 2 + Inches(0.10)
    rrect(sl, right_x, cy_d, right_w, Inches(1.0), WHITE, LINE, 0.5, radius=0.04)
    text(sl, right_x + Inches(0.15), cy_d + Inches(0.15), right_w - Inches(0.3), Inches(0.7),
         "n < 1 = shear-thinning\nLower values → better flow at high shear → excellent sprayability",
         size=12, color=INK3)

    cy_co = BODY_Y + fig_h + Inches(0.20)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.65),
            "High viscosity at rest for adhesion, low viscosity during spraying — flow properties compatible with existing firefighting equipment",
            icon='💨')


# ===== SLIDE 18: Herschel-Bulkley =====
def slide_18_hb(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 3/11", "Rheology ③: Herschel-Bulkley Model Parameters")
    draw_footer(sl, "18 / 36")

    # Top callout
    callout(sl, BODY_X, BODY_Y, BODY_W, Inches(0.65),
            "σ = τ₀ + K · γ̇ⁿ  （τ₀：降伏応力, K：稠度係数, n：流動指数）",
            icon='ƒ')

    ty = BODY_Y + Inches(0.85)
    th = Inches(2.85)
    headers = ["Formulation", "τ₀ (Pa)", "K (Pa·sⁿ)", "n", "R²"]
    rows = [
        ["AquaGel-K (Commercial control)", "0.31", "0.85", "0.52",
         {'text': ">0.99", 'color': D_GRN}],
        [{'text': "HEC+MC/CSP 1-5", 'bold': True},
         {'text': "0.59", 'bold': True}, {'text': "2.04", 'bold': True},
         {'text': "0.108", 'bold': True, 'color': AMBER},
         {'text': ">0.99", 'color': D_GRN}],
        ["HEC+MC/CSP/SDS 1-5-0.1", "0.54", "1.87", "0.115",
         {'text': ">0.99", 'color': D_GRN}],
        ["HEC+MC/CSP/SDS 1-5-0.5", "0.48", "1.65", "0.122",
         {'text': ">0.99", 'color': D_GRN}],
        [{'text': "MHEC/CSP 1-5", 'bold': True},
         {'text': "0.41", 'bold': True}, {'text': "1.43", 'bold': True},
         {'text': "0.188", 'bold': True, 'color': AMBER},
         {'text': ">0.99", 'color': D_GRN}],
    ]
    simple_table(sl, BODY_X, ty, BODY_W, th, headers, rows,
                 col_widths=[2.8, 1.2, 1.2, 1.0, 1.0])

    # Bottom 3 cards
    cy = ty + th + Inches(0.30)
    ch = Inches(1.85)
    cw = (BODY_W - Inches(0.30)) / 3
    cards_data = [
        ("Yield Stress τ₀", "Flow start requires minimum stress. WEG is ~2× AquaGel-K → higher resistance to sagging on vertical surfaces"),
        ("Flow Index n (<< 1)", "n=0.108 for HEC+MC/CSP shows strong shear thinning. Balances low viscosity during spraying and high viscosity at rest"),
        ("HEC+MC vs MHEC", "HEC+MC system shows lower n and higher K → stronger shear thinning. MHEC system achieves similar function with single polymer"),
    ]
    for i, (title, body) in enumerate(cards_data):
        cx = BODY_X + i * (cw + Inches(0.15))
        rrect(sl, cx, cy, cw, ch, WHITE, LINE, 0.5, radius=0.04)
        text(sl, cx + Inches(0.18), cy + Inches(0.16), cw - Inches(0.36), Inches(0.35),
             title, size=12, bold=True, color=MUTED)
        text(sl, cx + Inches(0.18), cy + Inches(0.56), cw - Inches(0.36), ch - Inches(0.7),
             body, size=12, color=INK3)


# ===== SLIDE 19: dynamic yield stress / tan δ =====
def slide_19_yield(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 4/11", "Rheology ④: Dynamic Yield Stress and Viscoelasticity (tan δ)")
    draw_footer(sl, "19 / 36")

    cw = (BODY_W - Inches(0.30)) / 2
    ch = Inches(5.6)
    cx2 = BODY_X + cw + Inches(0.30)

    # Left: dynamic yield stress
    text(sl, BODY_X, BODY_Y, cw, Inches(0.30),
         "Dynamic yield stress (from Amplitude Sweep)", size=12, bold=True, color=MUTED)
    by = BODY_Y + Inches(0.45)
    bar_data = [
        ("AquaGel-K", 0.28, "~0.4 Pa", GRAY_L),
        ("HEC+MC/CSP 1-5", 0.85, "~1.2 Pa", D_TEAL),
        ("+SDS 0.1 wt%", 0.78, "~1.1 Pa", D_BLUE),
        ("MHEC/CSP 1-5", 0.65, "~0.9 Pa", D_AMBR),
    ]
    for label, ratio, val, color in bar_data:
        bar_row(sl, BODY_X, by, cw, label, ratio, val, fill_color=color,
                lbl_w=Inches(2.0), val_w=Inches(1.0))
        by += Inches(0.50)

    callout(sl, BODY_X, by + Inches(0.15), cw, Inches(0.80),
            "Stress at G' = G'' crossover = dynamic yield stress. WEG is ~3× AquaGel-K",
            icon='ⓘ')

    # Right: tan δ
    text(sl, cx2, BODY_Y, cw, Inches(0.30),
         "Loss tangent tan δ = G''/G' (elastic-dominated = <1)", size=12, bold=True, color=MUTED)
    by = BODY_Y + Inches(0.45)
    tan_data = [
        ("AquaGel-K", 0.50, "~0.25", GRAY_L),
        ("HEC+MC/CSP 1-5", 0.20, "~0.10", D_TEAL),
        ("+SDS 0.1 wt%", 0.24, "~0.12", D_BLUE),
        ("MHEC/CSP 1-5", 0.30, "~0.15", D_AMBR),
    ]
    for label, ratio, val, color in tan_data:
        bar_row(sl, cx2, by, cw, label, ratio, val, fill_color=color,
                lbl_w=Inches(2.0), val_w=Inches(1.0))
        by += Inches(0.50)
    callout(sl, cx2, by + Inches(0.15), cw, Inches(0.80),
            "tan δ << 1 → all formulations show clear gel behavior. Elastic component overwhelmingly dominant",
            dark=True, icon='✓')


# ===== SLIDE 20: HEC+MC vs MHEC =====
def slide_20_compare(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 5/11", "HEC+MC vs MHEC: Comparison of Two Systems")
    draw_footer(sl, "20 / 36")

    cw = (BODY_W - Inches(0.20)) / 2
    ch = Inches(4.8)

    # Left: HEC+MC
    rrect(sl, BODY_X, BODY_Y, cw, ch, CARD_C, LINE, 0.5, radius=0.04)
    rect(sl, BODY_X, BODY_Y, cw, Inches(0.04), ACCENT)
    text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(0.18), cw - Inches(0.4), Inches(0.30),
         "HEC + MC Mixed System", size=12, bold=True, color=MUTED)
    text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(0.55), cw - Inches(0.4), Inches(0.40),
         "HEC (non-thermal gelation) + MC (thermal gelation, LCST)",
         size=13, italic=True, color=INK3)

    by = BODY_Y + Inches(1.10)
    bullets = [
        "HEC provides room-temperature viscosity and adhesion",
        "MC maintains structure during heating",
        "Lower n=0.108 → stronger shear thinning",
        "Time to char ≈ 9–10 min",
        "Foaming Index ≈ 2.2–2.6",
    ]
    for b in bullets:
        text(sl, BODY_X + Inches(0.30), by, Inches(0.2), Inches(0.30),
             "•", size=14, color=ACCENT)
        text(sl, BODY_X + Inches(0.55), by, cw - Inches(0.75), Inches(0.30),
             b, size=12, color=INK2)
        by += Inches(0.40)
    rrect(sl, BODY_X + Inches(0.20), by + Inches(0.15), cw - Inches(0.40), Inches(0.50),
          SOFT, ACCENT, 0.5, radius=0.04)
    text(sl, BODY_X + Inches(0.30), by + Inches(0.20), cw - Inches(0.60), Inches(0.40),
         "→ Good compatibility with SDS; excellent for optimizing foam structure",
         size=12, bold=True, color=ACCENT, anchor=MSO_ANCHOR.MIDDLE)

    # Right: MHEC
    cx2 = BODY_X + cw + Inches(0.20)
    rrect(sl, cx2, BODY_Y, cw, ch, CARD_C, LINE, 0.5, radius=0.04)
    rect(sl, cx2, BODY_Y, cw, Inches(0.04), ACCENT)
    text(sl, cx2 + Inches(0.20), BODY_Y + Inches(0.18), cw - Inches(0.4), Inches(0.30),
         "MHEC Single-Polymer System", size=12, bold=True, color=MUTED)
    text(sl, cx2 + Inches(0.20), BODY_Y + Inches(0.55), cw - Inches(0.4), Inches(0.40),
         "MHEC (integrates HEC+MC functions into a single molecule)",
         size=13, italic=True, color=INK3)
    by = BODY_Y + Inches(1.10)
    bullets2 = [
        "Carries both hydroxyethyl and methyl groups on the same chain",
        "Achieves equivalent viscoelasticity and thermal response with a single polymer",
        "n=0.188 (slightly higher than HEC+MC)",
        "Time to char ≈ 10 min",
        "Advantage of simplified formulation and quality control",
    ]
    for b in bullets2:
        text(sl, cx2 + Inches(0.30), by, Inches(0.2), Inches(0.30),
             "•", size=14, color=ACCENT)
        text(sl, cx2 + Inches(0.55), by, cw - Inches(0.75), Inches(0.30),
             b, size=12, color=INK2)
        by += Inches(0.40)
    rrect(sl, cx2 + Inches(0.20), by + Inches(0.15), cw - Inches(0.40), Inches(0.50),
          SOFT, ACCENT, 0.5, radius=0.04)
    text(sl, cx2 + Inches(0.30), by + Inches(0.20), cw - Inches(0.60), Inches(0.40),
         "→ Advantageous for scale-up and manufacturing cost reduction",
         size=12, bold=True, color=ACCENT, anchor=MSO_ANCHOR.MIDDLE)

    cy_co = BODY_Y + ch + Inches(0.25)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.80),
            "Both systems significantly outperform commercial AquaGel-K. HEC+MC excels in foam optimization; MHEC excels in manufacturing simplicity",
            icon='⚖')


# ===== SLIDE 21: 燃焼試験 setup =====
def slide_21_setup(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 6/11", "Combustion Test Setup")
    draw_footer(sl, "21 / 36")

    left_w = BODY_W * 0.55 - Inches(0.10)
    right_w = BODY_W * 0.45 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    fig_placeholder(sl, BODY_X, BODY_Y, left_w, Inches(4.6),
                    fig_num="Fig. 3a",
                    caption="Experimental setup for combustion testing")

    # Right: test condition card + callout
    ch = Inches(3.5)
    card(sl, right_x, BODY_Y, right_w, ch,
         title="◆ Test Conditions", tag="EXPERIMENTAL",
         bullets=[
             "Substrate: pine wood thin board",
             "Gel application: uniformly applied",
             "Flame source: butane direct flame",
             "Measurement: time until substrate charring",
             "Photographs taken at 120 s and 300 s",
         ], ct_size=15, li_size=12)
    callout(sl, right_x, BODY_Y + ch + Inches(0.20), right_w, Inches(0.85),
            "Time to char = time required for wood surface to char (longer = better protection)",
            icon='⏱')


# ===== SLIDE 22: Time to char =====
def slide_22_ttc(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 7/11", "Time to Char: Quantitative Comparison of 5 Formulations")
    draw_footer(sl, "22 / 36")

    left_w = BODY_W * 0.62 - Inches(0.10)
    right_w = BODY_W * 0.38 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    # Left: bar chart
    text(sl, BODY_X, BODY_Y, left_w, Inches(0.3),
         "Time to char (time to charring onset)", size=12, bold=True, color=MUTED)
    by = BODY_Y + Inches(0.45)
    bar_data = [
        ("Water", 0.18, "~2 min", GRAY_L),
        ("AquaGel-K (commercial)", 0.60, "~7 min", D_AMBR),
        ("HEC+MC/CSP 1-5", 0.80, "~9 min", D_TEAL),
        ("HEC+MC/CSP/SDS 1-5-0.1", 0.90, "~10 min", D_BLUE),
        ("MHEC/CSP 1-5", 0.92, "~10 min", D_BLUE),
    ]
    for label, ratio, val, color in bar_data:
        bar_row(sl, BODY_X, by, left_w, label, ratio, val, fill_color=color,
                lbl_w=Inches(2.6), val_w=Inches(1.1))
        by += Inches(0.55)

    callout(sl, BODY_X, by + Inches(0.15), left_w, Inches(0.85),
            "This work's WEG group extends protection ~40% beyond AquaGel-K and ~5× longer than water alone",
            icon='📊')

    # Right: figure placeholder
    fig_placeholder(sl, right_x, BODY_Y, right_w, Inches(5.2),
                    fig_num="Fig. 3b",
                    caption="Time-to-char bar chart (n≥3, error bars = standard deviation)")


# ===== SLIDE 23: 燃焼時系列 =====
def slide_23_timelapse(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 8/11", "Time-lapse Observation of Combustion Process")
    draw_footer(sl, "23 / 36")

    # Top: figure placeholder
    fig_h = Inches(2.4)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 3c",
                    caption="Sequential photographs during flame contact for each formulation")

    # Bottom: 3-card comparison
    cy = BODY_Y + fig_h + Inches(0.50)
    ch = Inches(2.7)
    cw = (BODY_W - Inches(0.28)) / 3
    cards_data = [
        ("Water", D_RED, [
            "Immediately evaporates and runs off under direct flame",
            "Wood surface chars in ~2 min",
            "No protective layer formed",
        ]),
        ("AquaGel-K", D_AMBR, [
            "Retains water but ends when water evaporates",
            "Chars in ~7 min",
            "No solid protective layer formed",
        ]),
        ("This Work WEG", ACCENT, [
            "Gel foams and expands simultaneously with water evaporation",
            "Porous aerogel layer provides continued protection",
            "Charring prevented for ~10 min",
        ]),
    ]
    for i, (title, color, bullets) in enumerate(cards_data):
        cx = BODY_X + i * (cw + Inches(0.14))
        bd = color if title == "This Work WEG" else LINE
        rrect(sl, cx, cy, cw, ch, WHITE, bd, 0.8 if title == "This Work WEG" else 0.5, radius=0.04)
        text(sl, cx + Inches(0.18), cy + Inches(0.18), cw - Inches(0.36), Inches(0.32),
             title, size=13, bold=True, color=color)
        by = cy + Inches(0.65)
        for b in bullets:
            text(sl, cx + Inches(0.22), by, Inches(0.20), Inches(0.30),
                 "•", size=14, color=color)
            text(sl, cx + Inches(0.47), by, cw - Inches(0.65), Inches(0.30),
                 b, size=12, color=INK3)
            by += Inches(0.42)


# ===== SLIDE 24: 120s/300s =====
def slide_24_120300(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 8/11", "Surface Condition Comparison at 120 s / 300 s")
    draw_footer(sl, "24 / 36")

    fig_h = Inches(2.3)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 3d",
                    caption="Surface condition comparison: 120 s (top) and 300 s (bottom)")

    cy = BODY_Y + fig_h + Inches(0.50)
    ch = Inches(2.8)
    cw = (BODY_W - Inches(0.28)) / 3
    cards_data = [
        ("Water", D_RED, "Time to char: ~2 min", [
            "120 s: charring already beginning over wide area",
            "300 s: entire surface blackened, severe damage",
        ]),
        ("AquaGel-K (commercial)", D_AMBR, "Time to char: ~7 min", [
            "120 s: gel drying, partial charring",
            "300 s: no protective layer, partial charring",
        ]),
        ("This Work WEG", ACCENT, "Time to char: ~10 min", [
            "120 s: aerogel layer forming and expanding",
            "300 s: surface nearly intact, layer remains",
        ]),
    ]
    for i, (title, color, ttc, bullets) in enumerate(cards_data):
        cx = BODY_X + i * (cw + Inches(0.14))
        bd = color if i == 2 else LINE
        rrect(sl, cx, cy, cw, ch, WHITE, bd, 0.8 if i == 2 else 0.5, radius=0.04)
        text(sl, cx + Inches(0.18), cy + Inches(0.16), cw - Inches(0.36), Inches(0.32),
             title, size=13, bold=True, color=color)
        by = cy + Inches(0.60)
        for b in bullets:
            text(sl, cx + Inches(0.22), by, cw - Inches(0.44), Inches(0.32),
                 "• " + b, size=12, color=INK3)
            by += Inches(0.45)
        text(sl, cx + Inches(0.18), cy + ch - Inches(0.40), cw - Inches(0.36), Inches(0.30),
             ttc, size=12, bold=True, color=color)


# ===== SLIDE 25: HERO =====
def slide_25_hero(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, HEADER)

    text(sl, Inches(0.5), Inches(1.2), Inches(12.3), Inches(0.5),
         "KEY RESULT", size=14, bold=True, color=RGBColor(0x88, 0xa4, 0xc0),
         align=PP_ALIGN.CENTER)

    text_runs(sl, Inches(0.5), Inches(1.9), Inches(12.3), Inches(2.3), [
        {'text': "~10", 'size': 130, 'bold': True, 'color': GOLD},
        {'text': " min", 'size': 60, 'bold': False, 'color': RGBColor(0xc8, 0x9a, 0x68)},
    ], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    text(sl, Inches(0.5), Inches(4.4), Inches(12.3), Inches(0.6),
         "Time to Char of This Work's WEG", size=26, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER)
    text(sl, Inches(1.5), Inches(5.15), Inches(10.3), Inches(0.7),
         "Achieved ~5× longer protection than water alone (~2 min) and ~1.4× longer than commercial AquaGel-K (~7 min)",
         size=14, color=RGBColor(0x9c, 0xb4, 0xcc), align=PP_ALIGN.CENTER)

    # Bars
    bars = [
        ("Water", 0.20, "~2 min", GRAY_L),
        ("AquaGel-K", 0.60, "~7 min", D_AMBR),
        ("This Work WEG", 0.95, "~10 min", GOLD),
    ]
    by = Inches(6.05)
    for label, ratio, val, color in bars:
        text(sl, Inches(2.5), by, Inches(1.7), Inches(0.30),
             label, size=11, color=RGBColor(0xc0, 0xd0, 0xe0), align=PP_ALIGN.RIGHT,
             anchor=MSO_ANCHOR.MIDDLE)
        bx_bar = Inches(4.3)
        bw_bar = Inches(7.0)
        rrect(sl, bx_bar, by + Inches(0.04), bw_bar, Inches(0.18),
              RGBColor(0x2c, 0x3a, 0x52), None, radius=0.5)
        if ratio > 0.02:
            rrect(sl, bx_bar, by + Inches(0.04), bw_bar * ratio, Inches(0.18),
                  color, None, radius=0.5)
        text(sl, bx_bar + bw_bar + Inches(0.15), by, Inches(1.0), Inches(0.30),
             val, size=11, color=RGBColor(0xc0, 0xd0, 0xe0), bold=True,
             anchor=MSO_ANCHOR.MIDDLE)
        by += Inches(0.35)

    text(sl, W - Inches(1.4), H - Inches(0.36), Inches(1.2), Inches(0.26),
         "25 / 36", size=10, color=RGBColor(0x55, 0x66, 0x77), align=PP_ALIGN.RIGHT)


# ===== SLIDE 26: 発泡指数 =====
def slide_26_foam(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 9/11", "Foaming Index Comparison")
    draw_footer(sl, "26 / 36")

    # Top: 2 figure placeholders
    fig_h = Inches(2.3)
    fw = (BODY_W - Inches(0.20)) / 2
    fig_placeholder(sl, BODY_X, BODY_Y, fw, fig_h, fig_num="Fig. 4a",
                    caption="Appearance of foamed layer after combustion")
    fig_placeholder(sl, BODY_X + fw + Inches(0.20), BODY_Y, fw, fig_h,
                    fig_num="Fig. 4b", caption="Foaming Index for each formulation")

    # Bottom: left bar chart, right explanation
    cy = BODY_Y + fig_h + Inches(0.45)
    cw = (BODY_W - Inches(0.20)) / 2

    text(sl, BODY_X, cy, cw, Inches(0.30),
         "Foaming Index (post-combustion thickness / initial thickness)", size=12, bold=True, color=MUTED)
    by = cy + Inches(0.40)
    bar_data = [
        ("AquaGel-K", 0.02, "~0", GRAY_L),
        ("HEC+MC/CSP 1-5", 0.55, "~2.2×", D_TEAL),
        ("+SDS 0.1 wt%", 0.65, "~2.6×", D_BLUE),
        ("MHEC/CSP 1-5", 0.58, "~2.3×", D_AMBR),
    ]
    for label, ratio, val, color in bar_data:
        bar_row(sl, BODY_X, by, cw, label, ratio, val, fill_color=color,
                lbl_w=Inches(1.6), val_w=Inches(0.95))
        by += Inches(0.45)

    # Right: explanation
    cx2 = BODY_X + cw + Inches(0.20)
    text(sl, cx2, cy, cw, Inches(0.30),
         "Insulating Effect from Foaming", size=12, bold=True, color=MUTED)
    by = cy + Inches(0.40)
    bullets = [
        "Expands to >2× initial thickness → heat conduction path extended",
        "Air trapped in pores enhances thermal insulation",
        "AquaGel-K does not foam → collapses flat under flame",
        "SDS 0.1 wt% is optimal: forms uniform fine bubbles",
    ]
    for b in bullets:
        text(sl, cx2 + Inches(0.05), by, Inches(0.20), Inches(0.30),
             "•", size=14, color=ACCENT)
        text(sl, cx2 + Inches(0.30), by, cw - Inches(0.35), Inches(0.30),
             b, size=12, color=INK2)
        by += Inches(0.42)


# ===== SLIDE 27: SDS濃度別SEM =====
def slide_27_sem_sds(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 9/11", "Effect of SDS Concentration on Aerogel Microstructure")
    draw_footer(sl, "27 / 36")

    fig_h = Inches(3.0)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 4c",
                    caption="FE-SEM images by SDS concentration (0%, 0.1%, 0.5%)")

    cy = BODY_Y + fig_h + Inches(0.45)
    ch = Inches(2.3)
    cw = (BODY_W - Inches(0.28)) / 3
    cards_data = [
        ("0% SDS", MUTED, "Dense silica network. Low porosity."),
        ("0.1% SDS (Optimal)", ACCENT, "Uniform fine bubbles. Highest foaming index."),
        ("0.5% SDS (Excess)", MUTED, "Non-uniform and coarsened bubble size."),
    ]
    for i, (title, color, desc) in enumerate(cards_data):
        cx = BODY_X + i * (cw + Inches(0.14))
        bd = ACCENT if title.endswith("(Optimal)") else LINE
        rrect(sl, cx, cy, cw, ch, WHITE, bd, 0.8 if title.endswith("(Optimal)") else 0.5, radius=0.04)
        text(sl, cx + Inches(0.18), cy + Inches(0.22), cw - Inches(0.36), Inches(0.35),
             title, size=13, bold=True, color=color)
        text(sl, cx + Inches(0.18), cy + Inches(0.75), cw - Inches(0.36), Inches(1.40),
             desc, size=12, color=INK3)


# ===== SLIDE 28: FT-IR / XPS =====
def slide_28_ftir_xps(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 10/11", "FT-IR & XPS: Chemical Composition Changes (Before/After Combustion)")
    draw_footer(sl, "28 / 36")

    cw = (BODY_W - Inches(0.20)) / 2

    # Left: FT-IR
    text(sl, BODY_X, BODY_Y, cw, Inches(0.30),
         "FT-IR Major Peak Assignments", size=12, bold=True, color=MUTED)
    ty = BODY_Y + Inches(0.42)
    th = Inches(3.0)
    headers = ["Wavenumber (cm⁻¹)", "Assignment", "After combustion"]
    rows = [
        ["3200–3500", "O–H stretch (cellulose, water)",
         {'text': "Disappears (dehydration)", 'color': D_RED}],
        ["2850–2950", "C–H stretch (methyl groups)",
         {'text': "Disappears (organic decomposition)", 'color': D_RED}],
        ["1050–1100", "Si–O–Si stretch (silica)",
         {'text': "Intensity increases", 'color': D_GRN, 'bold': True}],
        ["800", "Si–O bending vibration",
         {'text': "Becomes distinct", 'color': D_GRN, 'bold': True}],
        ["450", "Si–O bending",
         {'text': "Sharpens", 'color': D_GRN, 'bold': True}],
    ]
    simple_table(sl, BODY_X, ty, cw, th, headers, rows,
                 col_widths=[1.2, 2.5, 1.5])
    callout(sl, BODY_X, ty + th + Inches(0.20), cw, Inches(0.85),
            "Post-combustion spectrum matches pure silica (SiO₂) → aerogel formation chemically confirmed",
            icon='🔥')

    # Right: XPS
    cx2 = BODY_X + cw + Inches(0.20)
    text(sl, cx2, BODY_Y, cw, Inches(0.30),
         "XPS Surface Elemental Composition (at%)", size=12, bold=True, color=MUTED)
    th2 = Inches(2.1)
    headers2 = ["Element", "Before combustion", "After combustion"]
    rows2 = [
        [{'text': "Si 2p", 'bold': True}, "~4",
         {'text': "~33", 'color': D_GRN, 'bold': True}],
        [{'text': "C 1s", 'bold': True}, "~55",
         {'text': "~8", 'color': D_RED, 'bold': True}],
        [{'text': "O 1s", 'bold': True}, "~41", "~59"],
    ]
    simple_table(sl, cx2, BODY_Y + Inches(0.42), cw, th2, headers2, rows2,
                 col_widths=[1.5, 1.5, 1.5])
    callout(sl, cx2, BODY_Y + Inches(0.42) + th2 + Inches(0.15), cw, Inches(0.85),
            "Si concentration increases ~8×, C concentration greatly reduced → organic matrix removed, silica exposed",
            dark=True, icon='📊')

    info_y = BODY_Y + Inches(0.42) + th2 + Inches(1.20)
    rrect(sl, cx2, info_y, cw, Inches(0.85), WHITE, LINE, 0.5, radius=0.04)
    text(sl, cx2 + Inches(0.18), info_y + Inches(0.14), cw - Inches(0.36), Inches(0.30),
         "Peak Position Verification", size=11, bold=True, color=MUTED)
    text(sl, cx2 + Inches(0.18), info_y + Inches(0.42), cw - Inches(0.36), Inches(0.45),
         "Si 2p peak: ~103.5 eV (consistent with SiO₂) / O 1s peak: ~533 eV",
         size=12, color=INK3)


# ===== SLIDE 29: TGA/DSC =====
def slide_29_tga(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 11/11", "TGA/DSC: Thermal Decomposition Profile and Role of Each Component")
    draw_footer(sl, "29 / 36")

    cw = (BODY_W - Inches(0.20)) / 2

    # Left: TGA
    text(sl, BODY_X, BODY_Y, cw, Inches(0.30),
         "TGA: Weight Loss Profile (N₂ atmosphere)", size=12, bold=True, color=MUTED)
    by = BODY_Y + Inches(0.45)
    tga_data = [
        ("50–150°C", "Evaporation of free and adsorbed water", 0.30, D_BLUE),
        ("200–350°C", "Thermal decomposition of HEC/MC organic chains (main weight loss)", 0.85, D_RED),
        ("350–600°C", "Oxidation and loss of residual carbon (in air)", 0.50, D_AMBR),
        (">600°C", "Silica residue (~5–8 wt%, no change)", 0.08, D_TEAL),
    ]
    for label, desc, ratio, color in tga_data:
        text(sl, BODY_X, by, cw, Inches(0.25),
             label + ": " + desc, size=11, color=MUTED)
        rrect(sl, BODY_X, by + Inches(0.30), cw, Inches(0.14),
              CARD_C, LINE, 0.3, radius=0.5)
        if ratio > 0.02:
            rrect(sl, BODY_X, by + Inches(0.30), cw * ratio, Inches(0.14),
                  color, None, radius=0.5)
        by += Inches(0.62)

    # Right: DSC
    cx2 = BODY_X + cw + Inches(0.20)
    text(sl, cx2, BODY_Y, cw, Inches(0.30),
         "DSC: Thermal Event Assignments", size=12, bold=True, color=MUTED)
    headers = ["Temperature range", "Event", "ΔH"]
    rows = [
        ["~100°C", "Endothermic: water evaporation", "Endothermic"],
        ["~55–80°C", "Exothermic: MC gelation transition", {'text': "Slightly exothermic", 'color': D_GRN}],
        ["~250–300°C", "Exothermic: oxidative decomposition of organic chains", {'text': "Exothermic", 'color': D_RED}],
    ]
    simple_table(sl, cx2, BODY_Y + Inches(0.45), cw, Inches(1.6), headers, rows,
                 col_widths=[1.5, 3, 1.2])

    # Bottom mini stats
    my = BODY_Y + Inches(0.45) + Inches(1.6) + Inches(0.30)
    mw = (cw - Inches(0.15)) / 2
    mini(sl, cx2, my, mw, Inches(1.2), "Final silica residue", "~5–8 wt%", "Equivalent to CSP 5 wt%", value_color=D_TEAL)
    mini(sl, cx2 + mw + Inches(0.15), my, mw, Inches(1.2),
         "Cellulose decomposition peak", "~280°C", "Common to HEC/MC", value_color=D_RED)

    callout(sl, cx2, my + Inches(1.4), cw, Inches(0.85),
            "CSP sintering (~200°C+) and cellulose decomposition (~280°C+) overlap in temperature → synchronized aerogel formation",
            icon='🌡')


# ===== SLIDE 30: エアロゲル形成メカニズム =====
def slide_30_mechanism(prs):
    sl = new_slide(prs)
    draw_header(sl, "Supplementary 1/2", "Mechanism of Silica Aerogel Formation")
    draw_footer(sl, "30 / 36")

    fig_h = Inches(2.4)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 5a",
                    caption="3-stage schematic of WEG → silica aerogel transformation by flame activation")

    cy = BODY_Y + fig_h + Inches(0.45)
    ch = Inches(2.8)
    cw = (BODY_W - Inches(0.28)) / 3
    phases = [
        ("PHASE 1", "Unsintered", [
            "Silica particles dispersed independently",
            "Water fills the interstices",
            "No inter-particle bonding",
        ]),
        ("PHASE 2", "Grain boundary", [
            "Particles approach as water evaporates",
            "Grain boundaries form",
            "Initial neck formation begins",
        ]),
        ("PHASE 3", "Necking & porosity", [
            "Necks grow and strengthen",
            "Porosity decreases, skeleton complete",
            "SiO₂ aerogel layer established",
        ]),
    ]
    for i, (tag, title, bullets) in enumerate(phases):
        cx = BODY_X + i * (cw + Inches(0.14))
        card(sl, cx, cy, cw, ch, title=title, tag=tag,
             bullets=bullets, ct_size=15, li_size=12)


# ===== SLIDE 31: 燃焼時間別SEM =====
def slide_31_burnsem(prs):
    sl = new_slide(prs)
    draw_header(sl, "Supplementary 2/2", "Sintering Progression with Burn Time (SEM)")
    draw_footer(sl, "31 / 36")

    fig_h = Inches(2.3)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 5b",
                    caption="Cross-section SEM images of HEC+MC/CSP 1-5 after 0, 1, 2, and 4 min heating")

    cy = BODY_Y + fig_h + Inches(0.45)
    ch = Inches(2.7)
    cw = (BODY_W - Inches(0.42)) / 4
    stages = [
        ("0 min", "Unheated", MUTED, [
            "Spherical particles dispersed independently",
            "No inter-particle bonding",
            "Water fills the matrix",
        ]),
        ("1 min", "Early sintering", MUTED, [
            "Grain boundaries appear between particles",
            "Initial necks form",
            "Most moisture evaporated",
        ]),
        ("2 min", "Sintering in progress", MUTED, [
            "Neck regions expand",
            "Continuous network forms",
            "Porosity decreasing",
        ]),
        ("4 min", "Aerogel complete", ACCENT, [
            "Dense SiO₂ skeleton complete",
            "High mechanical strength",
            "Functions as insulating layer",
        ]),
    ]
    for i, (tag, title, color, bullets) in enumerate(stages):
        cx = BODY_X + i * (cw + Inches(0.14))
        bd = ACCENT if color == ACCENT else LINE
        rrect(sl, cx, cy, cw, ch, WHITE, bd, 0.8 if color == ACCENT else 0.5, radius=0.04)
        text(sl, cx + Inches(0.16), cy + Inches(0.16), cw - Inches(0.32), Inches(0.28),
             tag, size=11, bold=True, color=MUTED)
        text(sl, cx + Inches(0.16), cy + Inches(0.48), cw - Inches(0.32), Inches(0.42),
             title, size=15, bold=True, color=color)
        by = cy + Inches(1.05)
        for b in bullets:
            text(sl, cx + Inches(0.20), by, Inches(0.18), Inches(0.30),
                 "•", size=12, color=color)
            text(sl, cx + Inches(0.40), by, cw - Inches(0.56), Inches(0.30),
                 b, size=11, color=INK3)
            by += Inches(0.40)


# ===== SLIDE 33: 実装シナリオ =====
def slide_33_scenarios(prs):
    sl = new_slide(prs)
    draw_header(sl, "Discussion 1/3", "Implementation Scenarios: WEG Deployment Strategy")
    draw_footer(sl, "33 / 36")

    ch = Inches(3.6)
    cw = (BODY_W - Inches(0.28)) / 3
    cards_data = [
        ("Pre-protection of WUI Structures", "SCENARIO 1", [
            "Pre-applied to building exteriors and roofs before wildfire season",
            "Can be applied by helicopter or ground hose",
            "Protective function remains even after drying",
        ]),
        ("Protection of Critical Infrastructure", "SCENARIO 2", [
            "Power towers, substations, and communication base stations",
            "Strongly adheres to metal structures as well",
            "Prevents power outages and communication disruption",
        ]),
        ("Fire Break Formation", "SCENARIO 3", [
            "Pre-applied to vegetation along potential spread paths",
            "No environmental burden unlike Phos-Chek",
            "Cellulose and silica have low environmental persistence",
        ]),
    ]
    for i, (title, tag, bullets) in enumerate(cards_data):
        cx = BODY_X + i * (cw + Inches(0.14))
        card(sl, cx, BODY_Y, cw, ch, title=title, tag=tag,
             bullets=bullets, ct_size=15, li_size=12)

    cy = BODY_Y + ch + Inches(0.30)
    cw2 = (BODY_W - Inches(0.20)) / 2
    rrect(sl, BODY_X, cy, cw2, Inches(1.7), WHITE, LINE, 0.5, radius=0.04)
    text(sl, BODY_X + Inches(0.20), cy + Inches(0.14), cw2 - Inches(0.4), Inches(0.30),
         "Advantages Over Existing Methods", size=12, bold=True, color=MUTED)
    advs = [
        ("vs Water", "5× longer protection time"),
        ("vs AquaGel-K", "Continues protecting after water evaporation"),
        ("vs Phos-Chek", "No residual contamination in soil or water"),
        ("Spray compatibility", "Usable with existing equipment"),
    ]
    av = cy + Inches(0.45)
    for k, v in advs:
        text(sl, BODY_X + Inches(0.25), av, Inches(1.6), Inches(0.25),
             k, size=11, color=MUTED)
        text(sl, BODY_X + Inches(1.95), av, cw2 - Inches(2.2), Inches(0.25),
             v, size=11, bold=True, color=D_GRN)
        av += Inches(0.28)

    callout(sl, BODY_X + cw2 + Inches(0.20), cy + Inches(0.3), cw2, Inches(1.0),
            "Flow properties (shear thinning) compatible with both aerial and ground application are key to practical deployment",
            icon='🚁')


# ===== SLIDE 34: 課題と展望 =====
def slide_34_future(prs):
    sl = new_slide(prs)
    draw_header(sl, "Discussion 2/3", "Current Challenges and Future Research Directions")
    draw_footer(sl, "34 / 36")

    cw = (BODY_W - Inches(0.30)) / 2

    # Left: current limitations
    text(sl, BODY_X, BODY_Y, cw, Inches(0.30),
         "Current Limitations & Challenges", size=12, bold=True, color=MUTED)
    fy = BODY_Y + Inches(0.45)
    pw = (cw - Inches(0.30)) / 3
    ph = Inches(2.2)
    issues = [
        ("Scale", "No field trials", "Lab scale only. Verification in real environments required"),
        ("Durability", "Long-term storage stability", "Evaluation of phase separation and sedimentation. Seasonal storage compatibility"),
        ("Cost", "Mass production cost", "Establishing raw material costs and large-scale preparation processes"),
    ]
    for i, (tag, title, desc) in enumerate(issues):
        fx = BODY_X + i * (pw + Inches(0.15))
        flow_step(sl, fx, fy, pw, ph, tag, title, desc)
        if i < 2:
            text(sl, fx + pw, fy, Inches(0.15), ph,
                 "▶", size=12, color=MUTED,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    callout(sl, BODY_X, fy + ph + Inches(0.20), cw, Inches(0.85),
            "Behavior under real-world wind, dryness, and temperature gradients not yet evaluated",
            icon='⚠')

    # Right: future directions
    cx2 = BODY_X + cw + Inches(0.30)
    text(sl, cx2, BODY_Y, cw, Inches(0.30),
         "Future Research Directions", size=12, bold=True, color=MUTED)
    by = BODY_Y + Inches(0.55)
    items = [
        "Performance validation in large-scale combustion furnaces and outdoor simulated fire tests",
        "Adhesion evaluation on different substrates (concrete, metal, vegetation)",
        "Direct measurement of aerogel thermal conductivity (λ = 0.015–0.040 W/m·K expected)",
        "Exploration of SDS alternative surfactants (improved biodegradability)",
        "Systematic study of MC/HEC concentration optimization and CSP particle size effects",
        "Waterproofing improvement to eliminate need for reapplication after rainfall",
    ]
    for it in items:
        text(sl, cx2 + Inches(0.10), by, Inches(0.4), Inches(0.30),
             "→", size=15, bold=True, color=ACCENT)
        text(sl, cx2 + Inches(0.50), by, cw - Inches(0.60), Inches(0.30),
             it, size=12, color=INK2)
        by += Inches(0.55)


# ===== SLIDE 35: まとめ =====
def slide_35_summary(prs):
    sl = new_slide(prs)
    draw_header(sl, "Discussion 3/3", "Research Findings Summary and Outlook")
    draw_footer(sl, "35 / 36")

    cw = (BODY_W - Inches(0.42)) / 4
    ch = Inches(2.5)
    summaries = [
        ("RESULT 01", "Time to char ~10 min", "~5× water, ~40% longer than AquaGel-K. Greatly extends wood protection time under flame"),
        ("RESULT 02", "Self-forming aerogel", "CSP heat-activated sintering → porous silica layer continuously insulates even after water evaporation"),
        ("RESULT 03", "Compatible with existing spray infrastructure", "Shear-thinning fluid with n = 0.108–0.188. Sprayable through existing hoses and nozzles"),
        ("RESULT 04", "Sustainable materials", "Cellulose derivatives are food-grade; silica is a harmless inorganic material. Low environmental impact"),
    ]
    for i, (num, title, body) in enumerate(summaries):
        cx = BODY_X + i * (cw + Inches(0.14))
        rrect(sl, cx, BODY_Y, cw, ch, SOFT, LINE, 0.5, radius=0.04)
        text(sl, cx + Inches(0.18), BODY_Y + Inches(0.16), cw - Inches(0.36), Inches(0.28),
             num, size=10, bold=True, color=MUTED)
        text(sl, cx + Inches(0.18), BODY_Y + Inches(0.50), Inches(0.6), Inches(0.5),
             "●", size=22, color=ACCENT)
        text(sl, cx + Inches(0.18), BODY_Y + Inches(1.10), cw - Inches(0.36), Inches(0.45),
             title, size=15, bold=True, color=INK)
        text(sl, cx + Inches(0.18), BODY_Y + Inches(1.65), cw - Inches(0.36), ch - Inches(1.8),
             body, size=12, color=INK3)

    cy = BODY_Y + ch + Inches(0.25)
    ch2 = Inches(2.0)
    cw2 = (BODY_W - Inches(0.20)) / 2
    card(sl, BODY_X, cy, cw2, ch2,
         title="Anticipated Implementation Scenarios",
         bullets=[
             "Pre-protective application to WUI structures and critical infrastructure",
             "Fire break formation along potential spread paths",
             "Coating of metal structures such as power lines and communication towers",
         ], ct_size=14, li_size=12)
    card(sl, BODY_X + cw2 + Inches(0.20), cy, cw2, ch2,
         title="Future Challenges",
         bullets=[
             "Large-scale field spray tests (real-world benchmarking)",
             "Long-term storage stability and cost evaluation",
             "Environmental impact assessment (biodegradability testing)",
         ], ct_size=14, li_size=12)

    cy_co = cy + ch2 + Inches(0.20)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.55),
            "From \"water carrier\" to \"self-transforming fire-retardant material\" — A new WEG design paradigm",
            dark=True, icon='"')


# ===== SLIDE 36: HERO IMPACT =====
def slide_36_impact(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, HEADER)

    text(sl, Inches(0.5), Inches(0.9), Inches(12.3), Inches(0.5),
         "RESEARCH IMPACT", size=14, bold=True, color=RGBColor(0x88, 0xa4, 0xc0),
         align=PP_ALIGN.CENTER)

    text_runs(sl, Inches(0.5), Inches(1.5), Inches(12.3), Inches(2.0), [
        {'text': "5", 'size': 130, 'bold': True, 'color': GOLD},
        {'text': "×", 'size': 75, 'bold': True, 'color': RGBColor(0xc8, 0x9a, 0x68)},
    ], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    text(sl, Inches(0.5), Inches(3.6), Inches(12.3), Inches(0.5),
         "Time to Char Improvement vs. Conventional (Water)", size=22, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER)

    text(sl, Inches(1.5), Inches(4.25), Inches(10.3), Inches(1.4),
         "From \"water carrier\" to \"self-transforming fire-retardant material\"\nThe new WEG design paradigm opens a pathway to practical technology protecting critical infrastructure from wildfires",
         size=15, color=RGBColor(0xa8, 0xbc, 0xd4), align=PP_ALIGN.CENTER)

    bars = [
        ("Time to char", 0.95, "~10 min", GOLD),
        ("Foaming Index", 0.65, "~2.6×", ACCENT),
        ("Power-law n", 0.11, "0.108", GRAY_L),
    ]
    by = Inches(6.0)
    for label, ratio, val, color in bars:
        text(sl, Inches(3.0), by, Inches(2.0), Inches(0.28),
             label, size=11, color=RGBColor(0xc0, 0xd0, 0xe0), align=PP_ALIGN.RIGHT,
             anchor=MSO_ANCHOR.MIDDLE)
        bx_bar = Inches(5.1)
        bw_bar = Inches(5.2)
        rrect(sl, bx_bar, by + Inches(0.04), bw_bar, Inches(0.16),
              RGBColor(0x2c, 0x3a, 0x52), None, radius=0.5)
        if ratio > 0.02:
            rrect(sl, bx_bar, by + Inches(0.04), bw_bar * ratio, Inches(0.16),
                  color, None, radius=0.5)
        text(sl, bx_bar + bw_bar + Inches(0.15), by, Inches(1.0), Inches(0.28),
             val, size=11, color=RGBColor(0xc0, 0xd0, 0xe0), bold=True,
             anchor=MSO_ANCHOR.MIDDLE)
        by += Inches(0.32)

    text(sl, W - Inches(1.4), H - Inches(0.36), Inches(1.2), Inches(0.26),
         "36 / 36", size=10, color=RGBColor(0x55, 0x66, 0x77), align=PP_ALIGN.RIGHT)


# ── メイン ─────────────────────────────────────────────
def main():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    print("Building PPTX from scratch...")
    slide_01_cover(prs)
    slide_02_toc(prs)
    slide_section_divider(prs, "01", "Part 1", "Background",
                          "Wildfire damage under climate change and the fundamental limitations of existing fire retardant technologies",
                          ["Wildfire Status", "Existing WEG / Retardants", "Research Innovation"],
                          "3 / 36")
    slide_04_wildfire(prs)
    slide_05_process(prs)
    slide_06_existing(prs)
    slide_07_core(prs)
    slide_section_divider(prs, "02", "Part 2", "Materials & Methods",
                          "Novel gel design combining cellulosic polymers and colloidal silica",
                          ["Polymer Components", "Silica Particles & Surfactant", "Formulations (5 types)", "Evaluation Methods"],
                          "8 / 36")
    slide_09_polymers(prs)
    slide_10_mc(prs)
    slide_11_csp_sds(prs)
    slide_12_formulations(prs)
    slide_13_methods(prs)
    slide_14_adhesion(prs)
    slide_section_divider(prs, "03", "Part 3", "Results",
                          "Rheological properties, combustion performance, foam structure, and aerogel formation mechanism",
                          ["Rheology", "Combustion Test", "Foaming Index", "SEM Observation"],
                          "15 / 36")
    slide_16_rheology1(prs)
    slide_17_rheology2(prs)
    slide_18_hb(prs)
    slide_19_yield(prs)
    slide_20_compare(prs)
    slide_21_setup(prs)
    slide_22_ttc(prs)
    slide_23_timelapse(prs)
    slide_24_120300(prs)
    slide_25_hero(prs)
    slide_26_foam(prs)
    slide_27_sem_sds(prs)
    slide_28_ftir_xps(prs)
    slide_29_tga(prs)
    slide_30_mechanism(prs)
    slide_31_burnsem(prs)
    slide_section_divider(prs, "04", "Part 4", "Discussion & Conclusion",
                          "Implementation scenarios, comprehensive comparison with existing products, research significance and future prospects",
                          ["Implementation Scenarios", "Performance Comparison", "Summary"],
                          "32 / 36")
    slide_33_scenarios(prs)
    slide_34_future(prs)
    slide_35_summary(prs)
    slide_36_impact(prs)

    out_path = '/home/user/my-first-claude/slides_en.pptx'
    prs.save(out_path)
    print(f"Done: {out_path}")
    print(f"Slide count: {len(prs.slides.__iter__.__self__._sldIdLst)}")


if __name__ == '__main__':
    main()
