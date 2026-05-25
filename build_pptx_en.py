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

# ── Dimensions ───────────────────────────────────────────────
W = Inches(13.333)
H = Inches(7.5)
PAD = Inches(0.40)
HDR_H = Inches(0.72)
FTR_H = Inches(0.06)
BODY_Y = HDR_H + Inches(0.22)
BODY_H = H - BODY_Y - FTR_H - Inches(0.20)
BODY_X = PAD
BODY_W = W - PAD * 2

# Global font scale and family
FONT_SCALE = 1.12
FONT_NAME = 'Arial'


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
        run.font.size = Pt(size * FONT_SCALE)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        run.font.name = FONT_NAME
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
        run.font.size = Pt(r.get('size', 11) * FONT_SCALE)
        run.font.bold = r.get('bold', False)
        run.font.italic = r.get('italic', False)
        run.font.color.rgb = r.get('color', INK2)
        run.font.name = FONT_NAME
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


# ===== SLIDE 4: Wildfire Background =====
def slide_04_wildfire(prs):
    sl = new_slide(prs)
    draw_header(sl, "Background 1/4", "Growing Severity of Wildfire Damage and WUI Risk")
    draw_footer(sl, "4 / 36")

    # Left 42% typography stats / Right 58% photo placeholder (asymmetric)
    left_w = BODY_W * 0.42 - Inches(0.12)
    right_w = BODY_W * 0.58 - Inches(0.08)
    right_x = BODY_X + left_w + Inches(0.20)

    # Vertical accent bar (stat 1)
    rect(sl, BODY_X, BODY_Y, Inches(0.05), Inches(1.80), ACCENT)

    # Stat 1: large typography, no box
    text(sl, BODY_X + Inches(0.18), BODY_Y, left_w - Inches(0.18), Inches(1.10),
         "5×", size=68, bold=True, color=INK)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(1.08), left_w - Inches(0.18), Inches(0.32),
         "Increase in large-scale wildfires (past 30 years)", size=13, bold=True, color=INK2)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(1.40), left_w - Inches(0.18), Inches(0.28),
         "Driven by climate change and prolonged drought", size=11, color=MUTED)

    # Divider line
    rect(sl, BODY_X + Inches(0.18), BODY_Y + Inches(1.86), left_w * 0.75, Inches(0.02), LINE)

    # Vertical accent bar (stat 2)
    rect(sl, BODY_X, BODY_Y + Inches(2.08), Inches(0.05), Inches(1.60), D_AMBR)

    # Stat 2
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(2.08), left_w - Inches(0.18), Inches(0.90),
         "45 million", size=40, bold=True, color=INK)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(2.98), left_w - Inches(0.18), Inches(0.32),
         "U.S. WUI residents", size=13, bold=True, color=INK2)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(3.30), left_w - Inches(0.18), Inches(0.28),
         "High-risk population at the wildland–urban interface", size=11, color=MUTED)

    # Bottom note (horizontal rule, no callout box)
    rect(sl, BODY_X + Inches(0.18), BODY_Y + Inches(3.78), left_w * 0.90, Inches(0.02), LINE)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(3.96), left_w - Inches(0.18), Inches(0.70),
         "Extended dry and high-wind seasons mean\nconventional firefighting strategies can no longer keep pace",
         size=12, color=INK3, italic=True)

    # Right: wildfire / WUI photo placeholder
    fig_placeholder(sl, right_x, BODY_Y, right_w, BODY_H,
                    fig_num="Image",
                    caption="Wildfire / WUI boundary zone (photo)")


# ===== SLIDE 5: 4-Stage Protection Process =====
def slide_05_process(prs):
    sl = new_slide(prs)
    draw_header(sl, "Background 2/4", "Research Approach: 4-Stage Protection Process")
    draw_footer(sl, "5 / 36")

    # Top: large figure placeholder
    fig_h = Inches(3.5)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 1a",
                    caption="Gel application → flame contact → heat-activated aerogel formation → structural protection")

    # Bottom: timeline-style steps (numbered circles + text, no cards)
    sy = BODY_Y + fig_h + Inches(0.28)
    cw = (BODY_W - Inches(0.42)) / 4
    circ_d = Inches(0.44)
    steps = [
        ("1", "Gel Application", "Spray cellulosic gel\nonto buildings and vegetation"),
        ("2", "Flame Contact", "Moisture evaporates; gel\nfoams and particles aggregate"),
        ("3", "Aerogel Formation", "Silica particles sinter into\na porous insulating layer"),
        ("4", "Sustained Protection", "Ultra-low thermal conductivity\nlayer shields the substrate"),
    ]
    for i, (num, title, body) in enumerate(steps):
        cx = BODY_X + i * (cw + Inches(0.14))
        circ_x = cx + (cw - circ_d) / 2

        # Numbered circle
        rrect(sl, circ_x, sy, circ_d, circ_d, ACCENT, None, radius=0.5)
        text(sl, circ_x, sy, circ_d, circ_d, num, size=16, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Connector (not last)
        if i < 3:
            arr_x = cx + cw + Inches(0.04)
            arr_y = sy + circ_d / 2 - Inches(0.01)
            rect(sl, arr_x, arr_y, Inches(0.06), Inches(0.02), GRAY_L)

        # Text
        text(sl, cx, sy + circ_d + Inches(0.14), cw, Inches(0.30),
             title, size=13, bold=True, color=INK, align=PP_ALIGN.CENTER)
        text(sl, cx, sy + circ_d + Inches(0.46), cw, Inches(0.60),
             body, size=11, color=INK3, align=PP_ALIGN.CENTER)


# ===== SLIDE 6: Existing Technologies =====
def slide_06_existing(prs):
    sl = new_slide(prs)
    draw_header(sl, "Background 3/4", "Existing Wildfire Protection Methods and Their Fundamental Limitations")
    draw_footer(sl, "6 / 36")

    # 4 product cards in a row: photo placeholder on top, product name + limitation below
    n = 4
    gap = Inches(0.18)
    cw = (BODY_W - gap * (n - 1)) / n
    card_h = BODY_H - Inches(0.56)

    products = [
        ("Water only", "H₂O", "Runs off and evaporates\ninstantly — no persistence", False),
        ("Phos-Chek", "Ammonium phosphate", "Chemical residue persists\nin soil and water systems", False),
        ("AquaGel-K", "Cross-linked polyacrylate", "Protection lost the moment\nwater evaporates", False),
        ("WEG (This Work)", "Cellulosic + Silica", "Aerogel layer continues\nprotecting after water loss", True),
    ]

    for i, (name, ingredient, limit_text, is_hero) in enumerate(products):
        cx = BODY_X + i * (cw + gap)
        img_h = card_h * 0.50

        # Card background
        if is_hero:
            rrect(sl, cx, BODY_Y, cw, card_h, HEADER, None, radius=0.04)
            rect(sl, cx, BODY_Y, cw, Inches(0.04), GOLD)
        else:
            rrect(sl, cx, BODY_Y, cw, card_h, CARD_C, LINE, 0.5, radius=0.04)
            rect(sl, cx, BODY_Y, cw, Inches(0.04), LINE)

        # Photo placeholder (upper half)
        ph_fill = RGBColor(0x2a, 0x3c, 0x54) if is_hero else PH_BG
        ph_bord = RGBColor(0x44, 0x5e, 0x7a) if is_hero else PH_BD
        ph_tc = RGBColor(0x88, 0xa0, 0xb8) if is_hero else GRAY_L
        rrect(sl, cx + Inches(0.12), BODY_Y + Inches(0.12),
              cw - Inches(0.24), img_h - Inches(0.12), ph_fill, ph_bord, 0.7, radius=0.03)
        text(sl, cx + Inches(0.12), BODY_Y + Inches(0.12),
             cw - Inches(0.24), img_h - Inches(0.12),
             "(photo)", size=11, color=ph_tc,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Product name
        name_c = WHITE if is_hero else INK
        ingr_c = RGBColor(0xa0, 0xb8, 0xd0) if is_hero else MUTED
        ty = BODY_Y + img_h + Inches(0.14)
        text(sl, cx + Inches(0.10), ty, cw - Inches(0.20), Inches(0.34),
             name, size=14, bold=True, color=name_c, align=PP_ALIGN.CENTER)
        text(sl, cx + Inches(0.10), ty + Inches(0.34), cw - Inches(0.20), Inches(0.26),
             ingredient, size=10, color=ingr_c, align=PP_ALIGN.CENTER)

        # Limitation text
        limit_c = D_GRN if is_hero else D_RED
        text(sl, cx + Inches(0.10), ty + Inches(0.64), cw - Inches(0.20),
             card_h - (ty - BODY_Y) - Inches(0.72),
             limit_text, size=12, bold=is_hero,
             color=limit_c, align=PP_ALIGN.CENTER)

    # Bottom message (horizontal rule)
    msg_y = BODY_Y + card_h + Inches(0.14)
    rect(sl, BODY_X, msg_y, BODY_W, Inches(0.02), LINE)
    text(sl, BODY_X, msg_y + Inches(0.12), BODY_W, Inches(0.36),
         "Challenge: Existing WEGs are mere \"water carriers\" — protection vanishes the moment moisture is lost",
         size=13, bold=True, color=INK, align=PP_ALIGN.CENTER)


# ===== SLIDE 7: Core Innovation =====
def slide_07_core(prs):
    sl = new_slide(prs)
    draw_header(sl, "Background 4/4", "Research Core: Heat-Activated Aerogel Formation")
    draw_footer(sl, "7 / 36")

    # Left 38% text / Right 62% figure placeholder (figure as hero)
    left_w = BODY_W * 0.38 - Inches(0.12)
    right_w = BODY_W * 0.62 - Inches(0.08)
    right_x = BODY_X + left_w + Inches(0.20)

    # Vertical accent bar
    rect(sl, BODY_X, BODY_Y, Inches(0.05), Inches(1.52), ACCENT)

    # Hero text (quote style, no box)
    text(sl, BODY_X + Inches(0.18), BODY_Y, left_w - Inches(0.18), Inches(0.52),
         "Heat-Activated", size=22, bold=True, color=INK, italic=True)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(0.50), left_w - Inches(0.18), Inches(0.52),
         "Aerogel Formation", size=22, bold=True, color=ACCENT, italic=True)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(1.04), left_w - Inches(0.18), Inches(0.34),
         "Flame contact triggers the gel to transform into an aerogel insulator", size=11, color=INK3)

    # Divider line
    rect(sl, BODY_X + Inches(0.18), BODY_Y + Inches(1.54), left_w * 0.85, Inches(0.02), LINE)

    # 3 innovation points (numbered list, no card borders)
    points = [
        ("01", "Beyond water dependence", "Heat-activated solid protective layer"),
        ("02", "Sustainable materials", "Naturally derived, food-grade components"),
        ("03", "Drop-in compatible", "Sprayable fluid — no new infrastructure needed"),
    ]
    py = BODY_Y + Inches(1.74)
    for num, title, desc in points:
        text(sl, BODY_X, py, Inches(0.44), Inches(0.30),
             num, size=11, bold=True, color=ACCENT)
        text(sl, BODY_X + Inches(0.46), py, left_w - Inches(0.46), Inches(0.28),
             title, size=13, bold=True, color=INK)
        text(sl, BODY_X + Inches(0.46), py + Inches(0.28), left_w - Inches(0.46), Inches(0.28),
             desc, size=11, color=MUTED)
        py += Inches(0.75)

    # Right: concept schematic (intro slide — avoid forward-referencing a results figure)
    fig_placeholder(sl, right_x, BODY_Y, right_w, BODY_H,
                    fig_num="Concept",
                    caption="Heat trigger transforms the gel into a porous aerogel insulating layer")


# ===== SLIDE 9: セルロース系ポリマー =====
def slide_09_polymers(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials 1/6", "Cellulosic Polymers Used")
    draw_footer(sl, "9 / 36")

    # Left 55% polymer list (large letter badge + dividers) / Right 45% structure figure
    left_w = BODY_W * 0.55 - Inches(0.12)
    right_w = BODY_W * 0.45 - Inches(0.08)
    right_x = BODY_X + left_w + Inches(0.20)

    polymers = [
        ("A", "HEC", "Hydroxyethyl cellulose", "Base viscosity-building backbone"),
        ("B", "MC", "Methyl cellulose", "Gels on heating (thermoreversible) — key to flame protection"),
        ("C", "MHEC", "Methyl 2-hydroxyethyl cellulose", "Hybrid combining the traits of A and B"),
    ]
    row_h = BODY_H / 3
    for i, (letter, abbr, full, role) in enumerate(polymers):
        ry = BODY_Y + i * row_h
        # Large letter badge (faint, typographic accent)
        text(sl, BODY_X, ry, Inches(0.9), row_h,
             letter, size=54, bold=True, color=LINE, anchor=MSO_ANCHOR.MIDDLE)
        tx = BODY_X + Inches(0.95)
        text(sl, tx, ry + Inches(0.22), left_w - Inches(0.95), Inches(0.45),
             abbr, size=24, bold=True, color=INK)
        text(sl, tx, ry + Inches(0.70), left_w - Inches(0.95), Inches(0.28),
             full, size=11, color=MUTED)
        text(sl, tx, ry + Inches(1.00), left_w - Inches(0.95), Inches(0.34),
             role, size=12, bold=True, color=ACCENT)
        # Divider (not last)
        if i < 2:
            rect(sl, BODY_X, ry + row_h - Inches(0.02), left_w, Inches(0.015), LINE)

    # Right: structure figure placeholder
    fig_placeholder(sl, right_x, BODY_Y, right_w, BODY_H,
                    fig_num="Fig. 1b/c",
                    caption="Chemical structures and CSP mixing scheme")


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
        "LUDOX TM-50 (stock 50 wt%) → diluted to 15 wt% (pH 9)",
        "Particle size: 22 nm (monodisperse)",
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
        "Adding SDS did not improve Foaming Index",
        "Addition levels: two concentrations of 0.1 wt% and 0.5 wt%",
        "Higher SDS concentration coarsens bubble size (SEM)",
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
         ["Whitewood plywood heated with MAP-Pro torch (~2054°C)", "Time to charring onset measured", "Photographic comparison at 120 s / 300 s"]),
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
             "AquaGel-K (G'=457 Pa) shows higher G' than WEG (46/26 Pa); all formulations exhibit gel behavior",
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
         "Flow index n", "n < 1", "HEC+MC/CSP — strong shear thinning", value_color=ACCENT)
    mini(sl, right_x, BODY_Y + mh + Inches(0.15), right_w, mh,
         "Flow index n", "n < 1", "MHEC/CSP — shear thinning", value_color=ACCENT)

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
    draw_header(sl, "Results 3/11", "Rheology ③: Static Yield Stress and Long-term Stability")
    draw_footer(sl, "18 / 36")

    # Top callout
    callout(sl, BODY_X, BODY_Y, BODY_W, Inches(0.65),
            "Static yield stress: stress at the G'/G'' crossover (σ_s) from amplitude sweep — minimum stress to initiate flow",
            icon='ƒ')

    ty = BODY_Y + Inches(0.85)
    th = Inches(2.2)
    headers = ["Formulation", "Static yield stress (fresh)", "Static yield stress (455 days)", "Note"]
    rows = [
        [{'text': "HEC+MC/CSP 1-5", 'bold': True},
         {'text': "33.34 Pa", 'bold': True, 'color': ACCENT},
         {'text': "68.9 Pa", 'color': D_GRN},
         "Gel strength increases over time"],
        [{'text': "MHEC/CSP 1-5", 'bold': True},
         {'text': "3.31 Pa", 'bold': True, 'color': ACCENT},
         {'text': "4.56 Pa", 'color': D_GRN},
         "Long-term stability confirmed (455 days)"],
        ["AquaGel-K (Commercial control)", "—", "—",
         {'text': "Reference", 'color': MUTED}],
    ]
    simple_table(sl, BODY_X, ty, BODY_W, th, headers, rows,
                 col_widths=[2.8, 2.0, 2.0, 2.0])

    # Bottom 3 cards
    cy = ty + th + Inches(0.30)
    ch = Inches(1.85)
    cw = (BODY_W - Inches(0.30)) / 3
    cards_data = [
        ("Static Yield Stress", "Measured at G'/G'' crossover from amplitude sweep. HEC+MC/CSP: 33.34 Pa; MHEC/CSP: 3.31 Pa (fresh)"),
        ("Long-term Stability", "MHEC/CSP shows only minor change after 455 days (3.31→4.56 Pa). Demonstrates shelf stability for practical use"),
        ("HEC+MC vs MHEC", "HEC+MC system shows ~10× higher yield stress, better for vertical surface adhesion. MHEC simpler to manufacture"),
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
    draw_header(sl, "Results 4/11", "Rheology ④: Storage Modulus G' and Viscoelasticity (tan δ)")
    draw_footer(sl, "19 / 36")

    cw = (BODY_W - Inches(0.30)) / 2
    ch = Inches(5.6)
    cx2 = BODY_X + cw + Inches(0.30)

    # Left: dynamic yield stress
    text(sl, BODY_X, BODY_Y, cw, Inches(0.30),
         "G' Storage Modulus (1 rad/s)", size=12, bold=True, color=MUTED)
    by = BODY_Y + Inches(0.45)
    bar_data = [
        ("MHEC/CSP 1-5", 0.13, "25.6 Pa", D_AMBR),
        ("HEC+MC/CSP 1-5", 0.25, "46.4 Pa", D_TEAL),
        ("AquaGel-K", 0.95, "456.8 Pa", GRAY_L),
    ]
    for label, ratio, val, color in bar_data:
        bar_row(sl, BODY_X, by, cw, label, ratio, val, fill_color=color,
                lbl_w=Inches(2.0), val_w=Inches(1.0))
        by += Inches(0.50)

    callout(sl, BODY_X, by + Inches(0.15), cw, Inches(0.80),
            "All formulations show G'>G'' (gel behavior). AquaGel-K has the highest G' but also larger tan δ (more fluid-like)",
            icon='ⓘ')

    # Right: tan δ
    text(sl, cx2, BODY_Y, cw, Inches(0.30),
         "Loss tangent tan δ = G''/G' (elastic-dominated = <1)", size=12, bold=True, color=MUTED)
    by = BODY_Y + Inches(0.45)
    tan_data = [
        ("AquaGel-K", 0.42, "0.168", GRAY_L),
        ("HEC+MC/CSP 1-5", 0.55, "0.219", D_TEAL),
        ("MHEC/CSP 1-5", 1.00, "0.400", D_AMBR),
    ]
    for label, ratio, val, color in tan_data:
        bar_row(sl, cx2, by, cw, label, ratio, val, fill_color=color,
                lbl_w=Inches(2.0), val_w=Inches(1.0))
        by += Inches(0.50)
    callout(sl, cx2, by + Inches(0.15), cw, Inches(0.80),
            "All formulations: tan δ < 1 (gel behavior). AquaGel-K most elastic (0.168); MHEC system most fluid-like (0.400)",
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
        "Strong shear thinning (n < 1)",
        "Time to char > 7 min",
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
        "Shear thinning (n < 1)",
        "Time to char > 5 min",
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
             "Substrate: whitewood plywood",
             "Gel application: uniformly applied",
             "Flame source: MAP-Pro torch (~2054°C)",
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
        ("Water", 0.04, "~0.3 min", GRAY_L),
        ("AquaGel-K (commercial)", 0.20, "~1.5 min", D_AMBR),
        ("MHEC/CSP 1-5", 0.65, ">5 min", D_AMBR),
        ("HEC+MC/CSP 1-5", 0.90, ">7 min", D_TEAL),
        ("HEC+MC/CSP/SDS 1-5-0.1", 0.90, ">7 min", D_BLUE),
    ]
    for label, ratio, val, color in bar_data:
        bar_row(sl, BODY_X, by, left_w, label, ratio, val, fill_color=color,
                lbl_w=Inches(2.6), val_w=Inches(1.1))
        by += Inches(0.55)

    callout(sl, BODY_X, by + Inches(0.15), left_w, Inches(0.85),
            "This work's WEG achieves 3–6× longer protection than commercial AquaGel-K (n≥3 for all tests)",
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
            "Wood surface chars in ~0.3 min (~18 s)",
            "No protective layer formed",
        ]),
        ("AquaGel-K", D_AMBR, [
            "Retains water but ends when water evaporates",
            "Chars in ~1.5 min",
            "No solid protective layer formed",
        ]),
        ("This Work WEG", ACCENT, [
            "Gel foams and expands simultaneously with water evaporation",
            "Porous aerogel layer provides continued protection",
            "Charring prevented for >7 min (HEC+MC)",
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
        ("Water", D_RED, "Time to char: ~0.3 min", [
            "120 s: charring already beginning over wide area",
            "300 s: entire surface blackened, severe damage",
        ]),
        ("AquaGel-K (commercial)", D_AMBR, "Time to char: ~1.5 min", [
            "120 s: gel drying, partial charring",
            "300 s: no protective layer, partial charring",
        ]),
        ("This Work WEG", ACCENT, "Time to char: >7 min", [
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
        {'text': ">7", 'size': 130, 'bold': True, 'color': GOLD},
        {'text': " min", 'size': 60, 'bold': False, 'color': RGBColor(0xc8, 0x9a, 0x68)},
    ], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    text(sl, Inches(0.5), Inches(4.4), Inches(12.3), Inches(0.6),
         "Time to Char of This Work's WEG (HEC+MC/CSP)", size=26, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER)
    text(sl, Inches(1.5), Inches(5.15), Inches(10.3), Inches(0.7),
         "Achieved 3–6× longer protection than commercial AquaGel-K (~1.5 min) — n≥3 for all tests",
         size=14, color=RGBColor(0x9c, 0xb4, 0xcc), align=PP_ALIGN.CENTER)

    # Bars
    bars = [
        ("Water", 0.04, "~0.3 min", GRAY_L),
        ("AquaGel-K", 0.20, "~1.5 min", D_AMBR),
        ("This Work WEG", 0.95, ">7 min", GOLD),
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
        ("AquaGel-K", 0.02, "≈0", GRAY_L),
        ("MHEC/CSP 1-5", 0.53, "≈2.1×", D_AMBR),
        ("HEC+MC/CSP/SDS 1-5-0.1", 0.60, "≈2.3×", D_BLUE),
        ("HEC+MC/CSP 1-5", 0.68, "≈2.6×", D_TEAL),
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
        "HEC+MC/CSP without SDS achieves the highest foaming index (≈2.6)",
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
        ("0% SDS", ACCENT, "Highest foaming index (≈2.6). No SDS yields maximum expansion."),
        ("0.1% SDS", MUTED, "Uniform fine-bubble structure. Foaming index ≈2.3 (lower than no-SDS)."),
        ("0.5% SDS (Excess)", MUTED, "Non-uniform and coarsened bubble size. Foaming index further reduced."),
    ]
    for i, (title, color, desc) in enumerate(cards_data):
        cx = BODY_X + i * (cw + Inches(0.14))
        bd = ACCENT if color == ACCENT else LINE
        rrect(sl, cx, cy, cw, ch, WHITE, bd, 0.8 if color == ACCENT else 0.5, radius=0.04)
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
    headers2 = ["Formulation", "C 1s before (at%)", "C 1s after (at%)"]
    rows2 = [
        [{'text': "HEC+MC/CSP", 'bold': True},
         "25.6",
         {'text': "14.9", 'color': D_RED, 'bold': True}],
        [{'text': "MHEC/CSP", 'bold': True},
         "38.3",
         {'text': "4.8", 'color': D_RED, 'bold': True}],
    ]
    simple_table(sl, cx2, BODY_Y + Inches(0.42), cw, th2, headers2, rows2,
                 col_widths=[2.0, 1.8, 1.8])
    callout(sl, cx2, BODY_Y + Inches(0.42) + th2 + Inches(0.15), cw, Inches(0.85),
            "Carbon greatly reduced (MHEC: 38.3→4.8 at%) → organic matrix combusted and removed; silica remains",
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
        ("vs Water", "Far greater protection (water chars at ~0.3 min)"),
        ("vs AquaGel-K", "3–6× longer protection + continues after water evaporation"),
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
    draw_header(sl, "Summary 1/2", "Key Demonstrated Results")
    draw_footer(sl, "34 / 36")

    cw = (BODY_W - Inches(0.42)) / 4
    ch = Inches(4.6)
    cards_data = [
        ("RHEOLOGY", "Rheology", [
            "G' > G'', no crossover (gel)",
            "G': HEC+MC≈46, MHEC≈26 Pa",
            "Shear-thinning (HB) → sprayable",
            "tan δ < 1, solid-like gel",
        ]),
        ("FIRE PROTECTION", "Fire Protection", [
            "HEC+MC/CSP: >7 min to char",
            "MHEC/CSP: >5 min",
            "3–6× more effective than commercial",
            "Water ~0.3, AquaGel-K ~1.5 min",
        ]),
        ("FOAMING / SEM", "Foaming & Structure", [
            "Foaming Index up to ≈2.6",
            "AquaGel-K does not foam (≈0)",
            "Porous silica film on burning",
            "Particles sinter after 2 min (SEM)",
        ]),
        ("CHEMISTRY", "Chemical Change", [
            "FT-IR: CH/SiO ratio drops",
            "XPS: large carbon decrease",
            "MHEC/CSP C 38.3%→4.8%",
            "Silica (SiO₂) remains",
        ]),
    ]
    for i, (tag, title, bullets) in enumerate(cards_data):
        cx = BODY_X + i * (cw + Inches(0.14))
        card(sl, cx, BODY_Y, cw, ch, title=title, tag=tag,
             bullets=bullets, ct_size=14, li_size=11)

    cy_co = BODY_Y + ch + Inches(0.25)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.80),
            "On burning, water evaporates while the gel converts into a porous silica aerogel — insulating the substrate even after desiccation (a mechanism absent in commercial WEGs)",
            dark=True, icon='◆')


# ===== SLIDE 35: Conclusions =====
def slide_35_summary(prs):
    sl = new_slide(prs)
    draw_header(sl, "Summary 2/2", "Conclusions")
    draw_footer(sl, "35 / 36")

    cw = (BODY_W - Inches(0.42)) / 4
    ch = Inches(2.5)
    summaries = [
        ("RESULT 01", "3–6× more protective", "HEC+MC/CSP >7 min, MHEC/CSP >5 min to char. Protects substrates far longer under flame"),
        ("RESULT 02", "Self-forming aerogel", "Heat dehydrates the gel and sinters CSP → forms a porous silica aerogel insulating layer in situ (protection continues after drying)"),
        ("RESULT 03", "Compatible with spray infrastructure", "Shear-thinning, sprayable fluid combining strong substrate adherence and wetting"),
        ("RESULT 04", "Sustainable, safe materials", "Cellulose derivatives — Earth's most abundant biopolymer; biodegradability confirmed in prior study; silica is benign"),
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
         title="Core Mechanism",
         bullets=[
             "Water rapidly evaporates on flame contact",
             "CSP sinters, forming inter-particle necks",
             "Porous silica aerogel insulating layer forms",
         ], ct_size=14, li_size=12)
    card(sl, BODY_X + cw2 + Inches(0.20), cy, cw2, ch2,
         title="Significance Stated in the Paper",
         bullets=[
             "PP platform as a basis for diverse fire retardants",
             "Amenable to modular manufacturing & large scale",
             "Fire retardancy retained after 455-day aging (MHEC/CSP)",
         ], ct_size=14, li_size=12)

    cy_co = cy + ch2 + Inches(0.20)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.55),
            "From \"water carrier\" to \"heat-self-transforming fire-retardant material\" — a next-gen WEG that stays effective even after drying",
            dark=True, icon='"')


# ===== SLIDE 36: HERO IMPACT =====
def slide_36_impact(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, HEADER)

    text(sl, Inches(0.5), Inches(0.9), Inches(12.3), Inches(0.5),
         "RESEARCH IMPACT", size=14, bold=True, color=RGBColor(0x88, 0xa4, 0xc0),
         align=PP_ALIGN.CENTER)

    text_runs(sl, Inches(0.5), Inches(1.5), Inches(12.3), Inches(2.0), [
        {'text': "3–6", 'size': 100, 'bold': True, 'color': GOLD},
        {'text': "×", 'size': 65, 'bold': True, 'color': RGBColor(0xc8, 0x9a, 0x68)},
    ], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    text(sl, Inches(0.5), Inches(3.6), Inches(12.3), Inches(0.5),
         "Time to Char Improvement vs. Commercial WEG (AquaGel-K)", size=22, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER)

    text(sl, Inches(1.5), Inches(4.25), Inches(10.3), Inches(1.4),
         "From \"water carrier\" to \"self-transforming fire-retardant material\"\nThe new WEG design paradigm opens a pathway to practical technology protecting critical infrastructure from wildfires",
         size=15, color=RGBColor(0xa8, 0xbc, 0xd4), align=PP_ALIGN.CENTER)

    bars = [
        ("Time to char", 0.95, ">7 min", GOLD),
        ("Foaming Index", 0.68, "≈2.6×", ACCENT),
        ("Shear-thinning", 0.60, "n < 1", GRAY_L),
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
