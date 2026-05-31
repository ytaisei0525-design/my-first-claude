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

def oval(sl, x, y, w, h, fill, border=None, bpt=0.5):
    s = sl.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
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


def fig_placeholder(sl, x, y, w, h, fig_num=None, caption=None, cited=False):
    """Placeholder for paper figures (user will paste their own images later).
    cited=True marks a figure quoted from another paper (amber badge + source memo)."""
    # Background (solid border, light color)
    s = rrect(sl, x, y, w, h, PH_BG, PH_BD, 1.0, radius=0.02)
    badge_fill = AMBER if cited else HEADER
    place_text = "(Paste cited figure here)" if cited else "(Paste paper figure here)"
    # Figure number badge (top-left)
    if fig_num:
        tag_w = Inches(0.95)
        rect(sl, x + Inches(0.12), y + Inches(0.12), tag_w, Inches(0.32),
             badge_fill)
        text(sl, x + Inches(0.12), y + Inches(0.12), tag_w, Inches(0.32),
             fig_num, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Center placeholder text
    text(sl, x, y, w, h,
         place_text, size=12, color=GRAY_L,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Source memo for cited figures (inside box bottom)
    if cited:
        text(sl, x + Inches(0.12), y + h - Inches(0.38), w - Inches(0.24), Inches(0.28),
             "Source: ________________ (fill in later)", size=9, color=D_AMBR, italic=True)
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
         "1 / 43", size=10, color=RGBColor(0x66, 0x77, 0x88), align=PP_ALIGN.RIGHT)


# ===== SLIDE 2: TOC =====
def slide_02_toc(prs):
    sl = new_slide(prs)
    draw_header(sl, "Contents", "Presentation Outline")
    draw_footer(sl, "2 / 43")

    parts = [
        ("PART 1", "Background", "Current state of wildfires, limitations of existing technologies (Phos-Chek, AquaGel-K), and innovation of this research", "Slides 4–7"),
        ("PART 2", "Materials & Methods", "Cellulosic polymers, CSP, formulations (5 types), evaluation methods", "Slides 9–16"),
        ("PART 3", "Results", "Rheology, combustion tests, foaming index, SEM observation, mechanism", "Slides 18–34"),
        ("PART 4", "Discussion & Conclusion", "Implementation scenarios, comparison with existing products, conclusions and future prospects", "Slides 36–39"),
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
    draw_header(sl, "Background 1/5", "Growing Severity of Wildfire Damage and WUI Risk")
    draw_footer(sl, "4 / 43")

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
         "Driven by climate change and fuel accumulation (US/EU/AU)", size=11, color=MUTED)

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
         "Devastating loss to economy, infrastructure, natural\nresources and WUI residents — eco-friendly retardants are urgently needed",
         size=12, color=INK3, italic=True)

    # Right: top = wildfire photo / bottom = cited trend graph
    rtop_h = Inches(3.2)
    fig_placeholder(sl, right_x, BODY_Y, right_w, rtop_h,
                    fig_num="Image",
                    caption="Wildfire / WUI boundary zone (photo)")
    rbot_y = BODY_Y + rtop_h + Inches(0.46)
    rbot_h = Inches(2.0)
    fig_placeholder(sl, right_x, rbot_y, right_w, rbot_h,
                    fig_num="Cited", cited=True,
                    caption="Year-over-year trend in wildfire scale / annual burned area")


# ===== SLIDE 5: 4-Stage Protection Process =====
def slide_05_process(prs):
    sl = new_slide(prs)
    draw_header(sl, "Background 2/5", "Research Approach: 4-Stage Protection Process")
    draw_footer(sl, "5 / 43")

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


# ===== SLIDE 6: Existing Technologies (USFS 3 categories) =====
def _draw_retardant(sl, bx, by, bw, bh):
    """Long-term retardant: aircraft dropping red slurry"""
    SKY  = RGBColor(0xa8, 0xd4, 0xef)
    GRND = RGBColor(0x6a, 0xa5, 0x5e)
    PLN  = RGBColor(0x55, 0x6a, 0x7e)
    DROP = RGBColor(0xc0, 0x3c, 0x3c)

    rrect(sl, bx, by, bw, bh, SKY, None, radius=0.03)
    rect(sl, bx, by + bh * 0.78, bw, bh * 0.22, GRND)
    rect(sl, bx, by + bh * 0.78, bw, Inches(0.018), RGBColor(0x4e, 0x88, 0x42))

    bdy_w, bdy_h = bw * 0.30, bh * 0.08
    bdy_x = bx + bw * 0.10
    bdy_y = by + bh * 0.22
    rect(sl, bdy_x, bdy_y, bdy_w, bdy_h, PLN)
    rect(sl, bdy_x + bdy_w * 0.28, bdy_y - bh * 0.08, bw * 0.10, bh * 0.22, PLN)
    rect(sl, bdy_x + bw * 0.01, bdy_y - bh * 0.10, bw * 0.04, bh * 0.12, PLN)

    for fx, fy, fw, fh in [
        (0.35, 0.44, 0.040, 0.088), (0.44, 0.53, 0.036, 0.078),
        (0.52, 0.46, 0.038, 0.082), (0.42, 0.65, 0.032, 0.070),
        (0.58, 0.58, 0.034, 0.074), (0.50, 0.70, 0.030, 0.066),
        (0.62, 0.50, 0.028, 0.062),
    ]:
        oval(sl, bx + bw * fx, by + bh * fy, bw * fw, bh * fh, DROP)


def _draw_foam(sl, bx, by, bw, bh):
    """Foam suppressant: nozzle spraying overlapping foam bubbles"""
    BG    = RGBColor(0xbe, 0xe0, 0xf5)
    NOZL  = RGBColor(0x2c, 0x3e, 0x50)
    FOAM  = RGBColor(0xf2, 0xf8, 0xfe)
    FOAM2 = RGBColor(0xc8, 0xe4, 0xf5)

    rrect(sl, bx, by, bw, bh, BG, None, radius=0.03)
    rect(sl, bx + bw * 0.02, by + bh * 0.50, bw * 0.07, bh * 0.40, NOZL)
    rect(sl, bx + bw * 0.05, by + bh * 0.36, bw * 0.10, bh * 0.28, NOZL)
    rect(sl, bx + bw * 0.13, by + bh * 0.43, bw * 0.04, bh * 0.14, NOZL)

    for fx, fy, fw, fh, c in [
        (0.18, 0.22, 0.14, 0.30, FOAM),  (0.28, 0.14, 0.16, 0.32, FOAM),
        (0.40, 0.19, 0.14, 0.28, FOAM),  (0.51, 0.14, 0.15, 0.30, FOAM),
        (0.62, 0.20, 0.13, 0.26, FOAM),  (0.72, 0.16, 0.14, 0.28, FOAM),
        (0.82, 0.24, 0.12, 0.24, FOAM),
        (0.23, 0.44, 0.15, 0.30, FOAM2), (0.34, 0.38, 0.16, 0.32, FOAM2),
        (0.46, 0.44, 0.14, 0.28, FOAM),  (0.57, 0.38, 0.14, 0.28, FOAM2),
        (0.68, 0.44, 0.13, 0.26, FOAM),  (0.79, 0.36, 0.12, 0.24, FOAM2),
        (0.88, 0.30, 0.10, 0.20, FOAM),
    ]:
        oval(sl, bx + bw * fx, by + bh * fy, bw * fw, bh * fh, c)


def _draw_gel_house(sl, bx, by, bw, bh):
    """Water-enhancing gel: house with gel coating applied"""
    BG   = RGBColor(0x1e, 0x29, 0x3b)
    WALL = RGBColor(0xe8, 0xf0, 0xf8)
    ROOF = RGBColor(0xcc, 0xd8, 0xe8)
    GEL  = RGBColor(0x4a, 0xaa, 0x92)
    GEL2 = RGBColor(0x38, 0x90, 0x7a)
    WIN  = RGBColor(0x70, 0xaa, 0xcc)

    rrect(sl, bx, by, bw, bh, BG, None, radius=0.03)
    rect(sl, bx, by + bh * 0.88, bw, bh * 0.12, RGBColor(0x2e, 0x4a, 0x38))

    h_w = bw * 0.32
    h_x = bx + bw * 0.34
    w_y = by + bh * 0.48
    w_h = bh * 0.40

    oval(sl, h_x - bw * 0.03, w_y - bh * 0.06,
         h_w + bw * 0.06, w_h + bh * 0.12, GEL)
    rect(sl, h_x, w_y, h_w, w_h, WALL)

    s = sl.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
        h_x - bw * 0.02, by + bh * 0.18,
        h_w + bw * 0.04, bh * 0.32)
    s.fill.solid(); s.fill.fore_color.rgb = ROOF
    s.line.fill.background(); s.shadow.inherit = False

    for wi in range(2):
        rect(sl, h_x + h_w * 0.10 + wi * h_w * 0.50,
             w_y + w_h * 0.14, h_w * 0.26, w_h * 0.30, WIN)
    rect(sl, h_x + h_w * 0.36, w_y + w_h * 0.58,
         h_w * 0.28, w_h * 0.42, WIN)

    for fx, fy, fw, fh in [
        (0.10, 0.20, 0.028, 0.076), (0.15, 0.42, 0.024, 0.064),
        (0.20, 0.30, 0.026, 0.070), (0.80, 0.22, 0.026, 0.074),
        (0.84, 0.42, 0.024, 0.066), (0.76, 0.54, 0.028, 0.070),
    ]:
        oval(sl, bx + bw * fx, by + bh * fy, bw * fw, bh * fh, GEL2)


def slide_06_existing(prs):
    sl = new_slide(prs)
    draw_header(sl, "Background 3/5", "Existing Wildfire Suppressants and Their Limits (USFS 3 Categories)")
    draw_footer(sl, "6 / 43")

    n = 3
    gap = Inches(0.22)
    cw = (BODY_W - gap * (n - 1)) / n
    card_h = BODY_H - Inches(0.78)

    cats = [
        ("Long-term retardant", "Long-term retardant",
         "Phos-Chek (phosphate slurry)",
         "Aerial drop of red retardant slurry",
         "Chemicals such as ammonium phosphate",
         "Effective as long as residue remains",
         "Chemical residue left on the target", False),
        ("Foam suppressant", "Foam suppressant",
         "Class A foam / legacy AFFF",
         "Foam discharge from a fire hose",
         "Surfactants (historically fluorinated)",
         "Water retention 15–30 min (short-term)",
         "Loses effect when dry; fluorinated types bioaccumulate and are toxic", False),
        ("Water-enhancing gel (WEG)", "Water-enhancing gel",
         "Barricade / Thermo-Gel / AquaGel-K",
         "Pre-applied gel coating on homes",
         "Superabsorbent polymer (eco-friendly)",
         "Water retention 30–60 min; protects buildings",
         "Once dried by heat/wind, effectiveness is completely lost", True),
    ]

    for i, (name, en, example, photo_cap, comp, merit, limit, is_hero) in enumerate(cats):
        cx = BODY_X + i * (cw + gap)
        if is_hero:
            rrect(sl, cx, BODY_Y, cw, card_h, HEADER, None, radius=0.04)
            rect(sl, cx, BODY_Y, cw, Inches(0.05), GOLD)
        else:
            rrect(sl, cx, BODY_Y, cw, card_h, CARD_C, LINE, 0.5, radius=0.04)
            rect(sl, cx, BODY_Y, cw, Inches(0.05), LINE)

        pad = Inches(0.20)
        iw = cw - pad * 2
        name_c = WHITE if is_hero else INK
        body_c = RGBColor(0xd8, 0xe2, 0xee) if is_hero else INK3
        label_c = RGBColor(0x88, 0xa0, 0xb8) if is_hero else MUTED
        div_c = RGBColor(0x44, 0x58, 0x70) if is_hero else LINE
        accent_c = GOLD if is_hero else ACCENT

        # Illustration (AI-generated shapes)
        ph_y = BODY_Y + Inches(0.18)
        ph_h = Inches(1.18)
        if i == 0:
            _draw_retardant(sl, cx + pad, ph_y, iw, ph_h)
        elif i == 1:
            _draw_foam(sl, cx + pad, ph_y, iw, ph_h)
        else:
            _draw_gel_house(sl, cx + pad, ph_y, iw, ph_h)

        # Name
        ty = ph_y + ph_h + Inches(0.12)
        text(sl, cx + pad, ty, iw, Inches(0.40), name, size=15, bold=True, color=name_c)
        rect(sl, cx + pad, ty + Inches(0.56), iw, Inches(0.012), div_c)

        # Example (product)
        text(sl, cx + pad, ty + Inches(0.66), iw, Inches(0.22), "Example", size=9, bold=True, color=accent_c)
        text(sl, cx + pad, ty + Inches(0.86), iw, Inches(0.40), example, size=11, bold=True, color=body_c)

        # Composition
        text(sl, cx + pad, ty + Inches(1.32), iw, Inches(0.22), "Composition", size=9, bold=True, color=label_c)
        text(sl, cx + pad, ty + Inches(1.52), iw, Inches(0.42), comp, size=11, color=body_c)

        # Effect / retention
        text(sl, cx + pad, ty + Inches(2.00), iw, Inches(0.22), "Effect / retention", size=9, bold=True, color=label_c)
        text(sl, cx + pad, ty + Inches(2.20), iw, Inches(0.42), merit, size=11, color=body_c)

        # Limitation
        limit_label_c = GOLD if is_hero else D_RED
        text(sl, cx + pad, ty + Inches(2.68), iw, Inches(0.22), "Limitation", size=9, bold=True, color=limit_label_c)
        text(sl, cx + pad, ty + Inches(2.88), iw, card_h - (ty - BODY_Y) - Inches(3.06),
             limit, size=11, bold=True, color=(WHITE if is_hero else D_RED))

    # Bottom message
    msg_y = BODY_Y + card_h + Inches(0.16)
    callout(sl, BODY_X, msg_y, BODY_W, Inches(0.50),
            "This work overcomes the fatal flaw of eco-friendly WEGs: losing all effectiveness once dried",
            icon='→')


# ===== SLIDE 7: Predecessor work (PNP gel as APP carrier; adsorption & self-healing already known) =====
def slide_07_predecessor(prs):
    sl = new_slide(prs)
    draw_header(sl, "Background 4/5", "Predecessor Work: PNP Gel Platform (Yu et al., PNAS 2016)")
    draw_footer(sl, "7 / 43")

    callout(sl, BODY_X, BODY_Y, BODY_W, Inches(0.62),
            "The same HEC+MC/CSP platform was already established in 2016 as a "
            "carrier for ammonium polyphosphate fire retardants (Phos-Chek LC95A) — "
            "this work builds on that foundation",
            icon='📚')

    cw = (BODY_W - Inches(0.30)) / 2
    cx2 = BODY_X + cw + Inches(0.30)
    ctop = BODY_Y + Inches(0.78)
    ch = Inches(4.20)

    # Left: already established
    rrect(sl, BODY_X, ctop, cw, ch, CARD_C, LINE, 0.5, radius=0.04)
    rect(sl, BODY_X, ctop, cw, Inches(0.05), ACCENT)
    text(sl, BODY_X + Inches(0.24), ctop + Inches(0.18), cw - Inches(0.48), Inches(0.30),
         "Established in Yu et al., PNAS 2016", size=11, bold=True, color=MUTED)
    text(sl, BODY_X + Inches(0.24), ctop + Inches(0.50), cw - Inches(0.48), Inches(0.40),
         "Foundational properties of PNP gel (HEC+MC + CSP)", size=14, bold=True, color=INK)

    items_old = [
        ("PP interactions", "Cellulose chains selectively adsorb onto CSP (noncovalent, multivalent)"),
        ("Shear thinning",  "Viscosity drops at high shear (suitable for spraying and pipeline pumping)"),
        ("Self-healing",    "Gel structure rapidly recovers after stress release"),
        ("Scalability",     "Linearly scalable from 0.5 mL to 15 L"),
        ("APP carrier use", "Carries Phos-Chek LC95A; improves adhesion & rain resistance"),
    ]
    by = ctop + Inches(1.05)
    for label, body in items_old:
        rrect(sl, BODY_X + Inches(0.24), by + Inches(0.05),
              Inches(1.55), Inches(0.32),
              WHITE, ACCENT, 0.7, radius=0.5)
        text(sl, BODY_X + Inches(0.24), by + Inches(0.05),
             Inches(1.55), Inches(0.32),
             label, size=9, bold=True, color=ACCENT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(sl, BODY_X + Inches(1.88), by, cw - Inches(2.08), Inches(0.42),
             body, size=10, color=INK3, anchor=MSO_ANCHOR.MIDDLE)
        by += Inches(0.56)

    # Right: this work's novelty
    rrect(sl, cx2, ctop, cw, ch, HEADER, None, radius=0.04)
    rect(sl, cx2, ctop, cw, Inches(0.05), GOLD)
    text(sl, cx2 + Inches(0.24), ctop + Inches(0.18), cw - Inches(0.48), Inches(0.30),
         "New capabilities gained in this work", size=11, bold=True, color=GOLD)
    text(sl, cx2 + Inches(0.24), ctop + Inches(0.50), cw - Inches(0.48), Inches(0.40),
         "The gel itself transforms into an insulating layer", size=14, bold=True, color=WHITE)

    items_new = [
        ("Stand-alone",     "Suppresses fire without depending on APP"),
        ("Heat activation", "Flame contact → water evaporation → CSP sintering → porous silica"),
        ("Aerogel formation","Forms a highly insulating porous silica layer in situ"),
        ("Dry resistance",  "Overcomes the fatal flaw of conventional WEGs (\"dry = useless\")"),
        ("Long-term stability", "Yield stress retained after 455 days (see Supplement 3/3)"),
    ]
    by = ctop + Inches(1.05)
    for label, body in items_new:
        rrect(sl, cx2 + Inches(0.24), by + Inches(0.05),
              Inches(1.55), Inches(0.32),
              HEADER, GOLD, 0.7, radius=0.5)
        text(sl, cx2 + Inches(0.24), by + Inches(0.05),
             Inches(1.55), Inches(0.32),
             label, size=9, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(sl, cx2 + Inches(1.88), by, cw - Inches(2.08), Inches(0.42),
             body, size=10, color=RGBColor(0xd8, 0xe2, 0xee),
             anchor=MSO_ANCHOR.MIDDLE)
        by += Inches(0.56)

    # POINT band
    band_y = ctop + ch + Inches(0.18)
    band_h = H - band_y - FTR_H - Inches(0.18)
    rrect(sl, BODY_X, band_y, BODY_W, band_h, SOFT, LINE, 0.5, radius=0.05)
    rect(sl, BODY_X, band_y, Inches(0.08), band_h, ACCENT)
    text(sl, BODY_X + Inches(0.30), band_y, Inches(2.4), band_h,
         "POINT", size=13, bold=True, color=ACCENT, anchor=MSO_ANCHOR.MIDDLE)
    text(sl, BODY_X + Inches(2.0), band_y, BODY_W - Inches(2.3), band_h,
         "The innovation: from \"carrier of chemicals\" to \"material that itself "
         "becomes a thermal barrier\" — redefining the PP gel platform",
         size=13, color=INK2, anchor=MSO_ANCHOR.MIDDLE)


# ===== SLIDE 8: Core Innovation =====
def slide_07_core(prs):
    sl = new_slide(prs)
    draw_header(sl, "Background 5/5", "Research Core: Heat-Activated Aerogel Formation")
    draw_footer(sl, "8 / 43")

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

    # Proposed approach and mechanism (numbered list, no card borders)
    points = [
        ("01", "Material design", "Cellulosic biopolymer × colloidal silica, physically crosslinked"),
        ("02", "Aerogel formation", "Flame-driven water loss forms a porous silica network in situ"),
        ("03", "Sustained protection", "Insulating barrier prevents ignition even after water is gone"),
        ("04", "Applicability & scale", "Tunable flow for spray/pump; suited to large-scale manufacturing"),
    ]
    py = BODY_Y + Inches(1.74)
    for num, title, desc in points:
        text(sl, BODY_X, py, Inches(0.44), Inches(0.30),
             num, size=11, bold=True, color=ACCENT)
        text(sl, BODY_X + Inches(0.46), py, left_w - Inches(0.46), Inches(0.28),
             title, size=13, bold=True, color=INK)
        text(sl, BODY_X + Inches(0.46), py + Inches(0.28), left_w - Inches(0.46), Inches(0.40),
             desc, size=11, color=MUTED)
        py += Inches(0.70)

    # Right: top = concept schematic / bottom = 2 cited figures (aerogel SEM + insulation demo)
    rtop_h = Inches(3.1)
    fig_placeholder(sl, right_x, BODY_Y, right_w, rtop_h,
                    fig_num="Concept",
                    caption="Heat trigger transforms the gel into a porous aerogel insulating layer")
    rbot_y = BODY_Y + rtop_h + Inches(0.42)
    rbot_h = Inches(2.2)
    sub_w = (right_w - Inches(0.18)) / 2
    fig_placeholder(sl, right_x, rbot_y, sub_w, rbot_h,
                    fig_num="Cited", cited=True,
                    caption="SEM microstructure of silica aerogel")
    fig_placeholder(sl, right_x + sub_w + Inches(0.18), rbot_y, sub_w, rbot_h,
                    fig_num="Cited", cited=True,
                    caption="Aerogel thermal-insulation demo (sample over flame, etc.)")


# ===== SLIDE 9 (new): Material Composition: Sustainable & High-Performance =====
def slide_mat_composition(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials Overview 1/2", "Material Composition: Sustainable & High-Performance")
    draw_footer(sl, "10 / 43")

    # Lead sentence
    lead_h = Inches(0.50)
    text(sl, BODY_X, BODY_Y, BODY_W, lead_h,
         "Three building blocks — skeleton, function, and bonding — combine biodegradability with high performance.",
         size=14, color=INK2, anchor=MSO_ANCHOR.MIDDLE)

    n = 3
    gap = Inches(0.30)
    cw = (BODY_W - gap * (n - 1)) / n
    cards_y = BODY_Y + lead_h + Inches(0.14)
    card_h = H - cards_y - FTR_H - Inches(0.24)

    items = [
        ("🌿", "Cellulose Derivatives", "Plant-based biopolymers",
         "Plant-derived biopolymers that form the backbone of the gel network.",
         [("3 types", "HEC, MC, MHEC"),
          ("Bio.", "Highly biodegradable"),
          ("Role", "Viscosity & structure")],
         ACCENT),
        ("◎", "Colloidal Silica (CSP)", "Colloidal Silica Particles",
         "The key component that interacts dynamically with polymers and forms aerogel under heat.",
         [("Size", "~22 nm monodisperse"),
          ("Sinter", "Forms silica skeleton"),
          ("Role", "Source of aerogel")],
         D_TEAL),
        ("⬡", "PP Interaction", "Polymer–Particle Interaction",
         "Builds the network via dynamic, multivalent hydrogen bonds — no covalent bonds required.",
         [("Bonds", "Dynamic H-bonds"),
          ("Self-heal", "Gel recovers after strain"),
          ("Sprayable", "Easy spray application")],
         D_AMBR),
    ]

    for i, (icon, title, en, body, keys, accent_c) in enumerate(items):
        cx = BODY_X + i * (cw + gap)
        rrect(sl, cx, cards_y, cw, card_h, CARD_C, LINE, 0.5, radius=0.035)
        rect(sl, cx, cards_y, cw, Inches(0.06), accent_c)

        pad_x = Inches(0.24)
        inner_w = cw - pad_x * 2

        icon_y = cards_y + Inches(0.30)
        text(sl, cx, icon_y, cw, Inches(0.56),
             icon, size=30, color=accent_c,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        title_y = icon_y + Inches(0.62)
        text(sl, cx + pad_x, title_y, inner_w, Inches(0.42),
             title, size=16, bold=True, color=INK, align=PP_ALIGN.CENTER)
        text(sl, cx + pad_x, title_y + Inches(0.42), inner_w, Inches(0.28),
             en, size=10, italic=True, color=MUTED, align=PP_ALIGN.CENTER)

        div_y = title_y + Inches(0.78)
        rect(sl, cx + Inches(0.32), div_y, cw - Inches(0.64), Inches(0.015), LINE)

        body_y = div_y + Inches(0.14)
        text(sl, cx + pad_x, body_y, inner_w, Inches(1.10),
             body, size=12, color=INK3, align=PP_ALIGN.CENTER)

        ky = body_y + Inches(1.22)
        row_h = Inches(0.74)
        for j, (label, detail) in enumerate(keys):
            ry = ky + j * row_h
            chip_w = Inches(1.05)
            rrect(sl, cx + pad_x, ry + Inches(0.04), chip_w, Inches(0.32),
                  WHITE, accent_c, 0.7, radius=0.5)
            text(sl, cx + pad_x, ry + Inches(0.04), chip_w, Inches(0.32),
                 label, size=10, bold=True, color=accent_c,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            text(sl, cx + pad_x + chip_w + Inches(0.10), ry,
                 inner_w - chip_w - Inches(0.10), Inches(0.40),
                 detail, size=11, color=INK2, anchor=MSO_ANCHOR.MIDDLE)


# ===== SLIDE 10: Mechanism: Spray, Adhesion, and Flame Protection =====
def slide_self_protection(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials Overview 2/2", "Mechanism: Spray, Adhesion, and Flame Protection")
    draw_footer(sl, "11 / 43")

    # Lead sentence
    lead_h = Inches(0.52)
    text(sl, BODY_X, BODY_Y, BODY_W, lead_h,
         "Not temperature, but shear and self-healing govern the spray-to-adhesion behavior.",
         size=15, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)

    # Four large process cards (full width)
    n = 4
    arrow_w = Inches(0.30)
    cards_y = BODY_Y + lead_h + Inches(0.14)
    cw = (BODY_W - arrow_w * (n - 1)) / n
    card_h = Inches(3.50)

    stages = [
        ("01", "At rest", "Gel", "High-viscosity elastic gel. G' > G'' holds structure.",
         "Yield stress ~33 Pa", ACCENT),
        ("02", "Spraying", "Sol-like", "Shear strain lowers viscosity; the gel flows.",
         "Shear-thinning (viscosity drops)", D_TEAL),
        ("03", "On wall", "Self-healing gel", "Strain release instantly restores the gel network.",
         "G' recovery ~90%", D_GRN),
        ("04", "Flame", "Aerogel", "CSP sinters into a porous silica layer.",
         "Acts as thermal barrier", D_AMBR),
    ]

    for i, (no, phase, state, desc, metric, c) in enumerate(stages):
        cx = BODY_X + i * (cw + arrow_w)
        rrect(sl, cx, cards_y, cw, card_h, CARD_C, LINE, 0.5, radius=0.04)
        rect(sl, cx, cards_y, cw, Inches(0.06), c)

        pad_x = Inches(0.20)
        inner_w = cw - pad_x * 2

        text(sl, cx + pad_x, cards_y + Inches(0.18), inner_w, Inches(0.34),
             no, size=13, bold=True, color=c)
        text(sl, cx + pad_x, cards_y + Inches(0.54), inner_w, Inches(0.30),
             phase, size=12, bold=True, color=MUTED)
        text(sl, cx + pad_x, cards_y + Inches(0.92), inner_w, Inches(0.56),
             state, size=20, bold=True, color=INK)
        rect(sl, cx + pad_x, cards_y + Inches(1.58), inner_w, Inches(0.015), LINE)
        text(sl, cx + pad_x, cards_y + Inches(1.72), inner_w, Inches(1.10),
             desc, size=12, color=INK3)

        badge_h = Inches(0.46)
        badge_y = cards_y + card_h - badge_h - Inches(0.18)
        rrect(sl, cx + pad_x, badge_y, inner_w, badge_h, WHITE, c, 0.7, radius=0.06)
        text(sl, cx + pad_x, badge_y, inner_w, badge_h,
             metric, size=10, bold=True, color=c,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        if i < n - 1:
            text(sl, cx + cw, cards_y, arrow_w, card_h,
                 "→", size=20, color=GRAY_L,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Bottom POINT band
    band_y = cards_y + card_h + Inches(0.20)
    band_h = H - band_y - FTR_H - Inches(0.22)
    rrect(sl, BODY_X, band_y, BODY_W, band_h, HEADER, None, radius=0.05)
    rect(sl, BODY_X, band_y, Inches(0.08), band_h, GOLD)
    text(sl, BODY_X + Inches(0.30), band_y, Inches(2.4), band_h,
         "POINT", size=13, bold=True, color=GOLD, anchor=MSO_ANCHOR.MIDDLE)
    text(sl, BODY_X + Inches(2.0), band_y, BODY_W - Inches(2.3), band_h,
         "Rather than relying on MC's thermal response (LCST), rheology alone achieves \"spray → stick → set\" — the core of this material.",
         size=13, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)


# ===== SLIDE 9: セルロース系ポリマー =====
def slide_09_polymers(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials 1/6", "Cellulosic Polymers Used")
    draw_footer(sl, "12 / 43")

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

    # Right: top = paper structure figure / bottom = cited cellulose molecular structure
    rtop_h = Inches(3.3)
    fig_placeholder(sl, right_x, BODY_Y, right_w, rtop_h,
                    fig_num="Fig. 1b/c",
                    caption="Chemical structures and CSP mixing scheme")
    rbot_y = BODY_Y + rtop_h + Inches(0.46)
    rbot_h = Inches(2.0)
    fig_placeholder(sl, right_x, rbot_y, right_w, rbot_h,
                    fig_num="Cited", cited=True,
                    caption="Molecular structure of cellulose (reference)")


# ===== SLIDE 12: Rheological Design (Shear-Thinning & Self-Healing) =====
def slide_10_mc(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials 2/6", "Rheological Design: Shear-Thinning and Self-Healing")
    draw_footer(sl, "13 / 43")

    left_w = BODY_W * 0.48 - Inches(0.10)
    right_w = BODY_W * 0.52 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    # Left top: viscosity vs shear rate graph (shear-thinning behavior)
    img_h = Inches(2.7)
    fig_placeholder(sl, BODY_X, BODY_Y, left_w, img_h,
                    fig_num="Fig. 2",
                    caption="Viscosity vs. shear rate — shear-thinning behavior")

    # Left bottom: key properties
    ty = BODY_Y + img_h + Inches(0.20)
    rect(sl, BODY_X, ty, left_w * 0.9, Inches(0.015), LINE)
    ty += Inches(0.18)
    props = [
        ("Shear-thinning", "Viscosity drops sharply at high shear — fluidizes during spraying"),
        ("Yield stress", "~33 Pa (dynamic, HB fit) — gel structure maintained at rest"),
        ("Self-healing", "G' recovers ~90% after strain removal"),
    ]
    for label, detail in props:
        text(sl, BODY_X, ty, Inches(1.15), Inches(0.30),
             label, size=10, bold=True, color=ACCENT)
        text(sl, BODY_X + Inches(1.19), ty, left_w - Inches(1.19), Inches(0.30),
             detail, size=12, color=INK2)
        ty += Inches(0.40)

    # Right: 3 mini stats
    mw = (right_w - Inches(0.20)) / 3
    mh = Inches(1.4)
    mini(sl, right_x, BODY_Y, mw, mh, "Viscosity response", "reversible", "Low at high shear → recovers at rest", value_color=ACCENT)
    mini(sl, right_x + mw + Inches(0.10), BODY_Y, mw, mh, "Dynamic yield stress", "~33 Pa", "HEC+MC/CSP (HB fit)", value_color=D_TEAL)
    mini(sl, right_x + (mw + Inches(0.10)) * 2, BODY_Y, mw, mh, "G' recovery", "~90%", "After 1000 s", value_color=D_GRN)

    # Spray-to-flame process flow (4 steps)
    fy = BODY_Y + mh + Inches(0.30)
    fh = Inches(2.7)
    rrect(sl, right_x, fy, right_w, fh, HEADER, None, radius=0.04)
    text(sl, right_x + Inches(0.20), fy + Inches(0.18), right_w - Inches(0.4), Inches(0.30),
         "Spray-to-Flame Process", size=11, bold=True, color=RGBColor(0xa0, 0xb8, 0xd0))

    sub_y = fy + Inches(0.65)
    sub_h = fh - Inches(0.75)
    pw = (right_w - Inches(0.40) - Inches(0.30) * 3) / 4
    steps_t = [
        ("At rest", "Gel state", "G' > G''"),
        ("Spraying", "Shear-thinning", "Low viscosity"),
        ("On wall", "Self-healing", "G' recovery"),
        ("Flame", "Aerogel", "Porous silica"),
    ]
    px = right_x + Inches(0.20)
    for i, (phase, ti, dsc) in enumerate(steps_t):
        rrect(sl, px, sub_y, pw, sub_h, RGBColor(0x2c, 0x3b, 0x52), None, radius=0.04)
        text(sl, px + Inches(0.10), sub_y + Inches(0.14), pw - Inches(0.2), Inches(0.30),
             phase, size=11, bold=True, color=GOLD)
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
    draw_footer(sl, "14 / 43")

    cw = (BODY_W - Inches(0.18)) / 2
    callout_h = Inches(0.72)
    img_h = Inches(2.4)
    text_top = BODY_Y + img_h + Inches(0.18)

    # Left: CSP — TEM photo placeholder on top, text below
    cx = BODY_X
    fig_placeholder(sl, cx, BODY_Y, cw, img_h,
                    fig_num="TEM",
                    caption="CSP particles (22 nm, monodisperse)")
    text(sl, cx, text_top, cw, Inches(0.34),
         "CSP — Colloidal Silica Particles", size=14, bold=True, color=INK)
    bullets_csp = [
        "LUDOX TM-50 (stock 50 wt%) → diluted to 15 wt% (pH 9)",
        "Particle size: 22 nm (monodisperse)",
        "Formulation concentration: 5 wt% in gel",
        "Sinters upon heating → transforms into silica aerogel",
    ]
    by = text_top + Inches(0.38)
    for b in bullets_csp:
        text(sl, cx, by, Inches(0.20), Inches(0.28), "•", size=12, color=ACCENT)
        text(sl, cx + Inches(0.22), by, cw - Inches(0.22), Inches(0.28), b, size=12, color=INK2)
        by += Inches(0.34)

    # Right: SDS — SEM foam structure photo on top, text below
    cx2 = BODY_X + cw + Inches(0.18)
    fig_placeholder(sl, cx2, BODY_Y, cw, img_h,
                    fig_num="SEM",
                    caption="Foam structure variation by SDS concentration")
    text(sl, cx2, text_top, cw, Inches(0.34),
         "SDS — Sodium Dodecyl Sulfate", size=14, bold=True, color=INK)
    bullets_sds = [
        "Anionic surfactant (additive)",
        "Adding SDS did not improve Foaming Index",
        "Addition levels: 0.1 wt% and 0.5 wt%",
        "Higher SDS concentration coarsens bubble size (SEM)",
    ]
    by2 = text_top + Inches(0.38)
    for b in bullets_sds:
        text(sl, cx2, by2, Inches(0.20), Inches(0.28), "•", size=12, color=ACCENT)
        text(sl, cx2 + Inches(0.22), by2, cw - Inches(0.22), Inches(0.28), b, size=12, color=INK2)
        by2 += Inches(0.34)

    # Bottom callout
    cy_co = H - FTR_H - Inches(0.10) - callout_h
    callout(sl, BODY_X, cy_co, BODY_W, callout_h,
            "The sintering temperature range of CSP overlaps with cellulose thermal decomposition — synchronized response forms the insulating layer",
            icon='💡')


# ===== SLIDE 12: 配合系 5種 =====
def slide_12_formulations(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials 4/6", "Five Formulations Evaluated")
    draw_footer(sl, "15 / 43")

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

    # Bottom: 2 cards + formulation appearance photo
    cy = fy + fh + Inches(0.40)
    ch = Inches(2.5)
    cw3 = (BODY_W - Inches(0.36)) / 3
    card(sl, BODY_X, cy, cw3, ch,
         title="Formulation Design Intent",
         bullets=[
             "Compare functional differences by polymer type",
             "Verify foam structure at two SDS levels",
             "Direct comparison with AquaGel-K",
         ], ct_size=14, li_size=11)
    card(sl, BODY_X + cw3 + Inches(0.18), cy, cw3, ch,
         title="Notation Guide",
         body="HEC+MC/CSP/SDS 1-5-0.1\n→ HEC+MC 1 wt%\n／ CSP 5 wt%\n／ SDS 0.1 wt%",
         ct_size=14, cb_size=13)
    fig_placeholder(sl, BODY_X + (cw3 + Inches(0.18)) * 2, cy, cw3, ch - Inches(0.34),
                    fig_num="Photo",
                    caption="Appearance of the five formulations")


# ===== SLIDE 13: 評価手法 =====
def slide_13_methods(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials 5/6", "Overview of Evaluation Methods")
    draw_footer(sl, "16 / 43")

    # Left: burn-test setup figure / Right: 6 method cards (2 cols × 3 rows)
    left_w = BODY_W * 0.34 - Inches(0.10)
    right_w = BODY_W * 0.66 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)
    fig_placeholder(sl, BODY_X, BODY_Y, left_w, BODY_H - Inches(0.34),
                    fig_num="Setup",
                    caption="Combustion test setup (MAP-Pro torch / plywood substrate)")

    methods = [
        ("Rheological Measurements", "RHEOLOGY",
         ["Oscillatory sweep (G', G'')", "Steady-flow sweep (viscosity)", "Herschel-Bulkley fitting"]),
        ("Combustion Test", "BURN TEST",
         ["MAP-Pro torch (~2054°C)", "Time to charring onset", "Compare at 120/300 s"]),
        ("Foaming Index", "FOAMING",
         ["Foamed layer thickness", "Ratio to initial = Foaming Index"]),
        ("SEM Observation", "SEM",
         ["Foam structure by SDS level", "Sintering by burn time"]),
        ("Spectroscopy", "SPECTROSCOPY",
         ["FT-IR (chemical bonding)", "XPS (surface composition)"]),
        ("Thermal Analysis", "THERMAL",
         ["TGA (thermogravimetric)", "DSC (calorimetry)"]),
    ]
    cw = (right_w - Inches(0.15)) / 2
    ch = (BODY_H - Inches(0.40)) / 3
    for i, (title, tag, bullets) in enumerate(methods):
        col = i % 2
        row = i // 2
        cx = right_x + col * (cw + Inches(0.15))
        cy = BODY_Y + row * (ch + Inches(0.20))
        card(sl, cx, cy, cw, ch, title=title, tag=tag,
             bullets=bullets, ct_size=13, li_size=10)


# ===== SLIDE 14: 付着性・濡れ性 =====
def slide_14_adhesion(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials 6/6", "Rheological Design Translates Directly into Field Performance")
    draw_footer(sl, "17 / 43")

    # Layout: left (measured rheology) → arrow → (expected field performance), 3 rows
    left_w = Inches(4.5)
    arrow_w = Inches(0.7)
    gap = Inches(0.18)
    right_x = BODY_X + left_w + gap + arrow_w + gap
    right_w = BODY_W - left_w - arrow_w - gap * 2

    # Column headers
    lbl_y = BODY_Y
    text(sl, BODY_X, lbl_y, left_w, Inches(0.30),
         "Rheological Design (measured properties)", size=12, bold=True, color=MUTED)
    text(sl, right_x, lbl_y, right_w, Inches(0.30),
         "Expected Field Performance (behavior in use)", size=12, bold=True, color=MUTED)

    # 3-row causal mapping: (left title, left desc, right title, right desc, accent color)
    rows = [
        ("Shear-thinning behavior", "Viscosity drops sharply at high shear",
         "Sprayable", "Viscosity drops under nozzle's high shear\n→ Deployable with existing hoses/nozzles", D_TEAL),
        ("High Storage Modulus G' (G' ≫ G'')", "Holds gel state at rest",
         "No run-off on vertical surfaces", "Retains shape and adheres after application\n→ Holds on vertical substrates (wood, walls)", ACCENT),
        ("Self-healing", "G' recovers instantly after step-strain",
         "Forms protective layer on landing", "Re-gels immediately after spraying\n→ Fire layer maintained against wind/rain/gravity", GOLD),
    ]

    row_y = BODY_Y + Inches(0.42)
    row_h = Inches(1.32)
    row_gap = Inches(0.20)
    for ltitle, ldesc, rtitle, rdesc, c in rows:
        # Left: measured property card
        rrect(sl, BODY_X, row_y, left_w, row_h, CARD_C, LINE, 0.5, radius=0.05)
        rect(sl, BODY_X, row_y, Inches(0.06), row_h, c)
        text(sl, BODY_X + Inches(0.22), row_y + Inches(0.18), left_w - Inches(0.34), Inches(0.40),
             ltitle, size=14, bold=True, color=INK)
        text(sl, BODY_X + Inches(0.22), row_y + Inches(0.70), left_w - Inches(0.34), Inches(0.50),
             ldesc, size=12, color=INK2)

        # Center: arrow
        ar = sl.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                 BODY_X + left_w + gap, row_y + row_h/2 - Inches(0.22),
                                 arrow_w, Inches(0.44))
        ar.fill.solid(); ar.fill.fore_color.rgb = c
        ar.line.fill.background(); ar.shadow.inherit = False

        # Right: field performance card
        rrect(sl, right_x, row_y, right_w, row_h, WHITE, c, 1.0, radius=0.05)
        text(sl, right_x + Inches(0.22), row_y + Inches(0.16), right_w - Inches(0.34), Inches(0.40),
             rtitle, size=14, bold=True, color=c)
        text(sl, right_x + Inches(0.22), row_y + Inches(0.62), right_w - Inches(0.34), Inches(0.62),
             rdesc, size=12, color=INK2)

        row_y += row_h + row_gap

    # Bottom: honest note (no independent adhesion/contact-angle test in this paper)
    cy_co = H - FTR_H - Inches(0.16) - Inches(0.62)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.62),
            "These field performances are inferred from measured rheology. No independent contact-angle or adhesion-force tests were performed in this paper",
            icon='ⓘ')


# ===== SLIDE 16: G'/G'' =====
def slide_16_rheology1(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 1/15", "Rheology ①: Viscoelasticity & Shear-Thinning (Fig. 2a–c)")
    draw_footer(sl, "19 / 43")

    # Two figures side by side (left = viscoelasticity 2a–b, right = shear-thinning 2c)
    gap = Inches(0.24)
    fw = (BODY_W - gap) / 2
    fig_h = Inches(3.5)

    fig_placeholder(sl, BODY_X, BODY_Y, fw, fig_h,
                    fig_num="Fig. 2a–b",
                    caption="Angular frequency vs G' / G'' (viscoelasticity at rest)")
    cx2 = BODY_X + fw + gap
    fig_placeholder(sl, cx2, BODY_Y, fw, fig_h,
                    fig_num="Fig. 2c",
                    caption="Shear rate vs viscosity (shear-thinning)")

    # Two observation cards
    cy = BODY_Y + fig_h + Inches(0.22)
    ch = Inches(1.0)
    # Left card: gel behavior
    rrect(sl, BODY_X, cy, fw, ch, CARD_C, D_TEAL, 1.2, radius=0.05)
    rect(sl, BODY_X, cy, Inches(0.06), ch, D_TEAL)
    text(sl, BODY_X + Inches(0.20), cy + Inches(0.12), fw - Inches(0.34), Inches(0.30),
         "At rest: G' > G'' → solid-like (gel) behavior", size=12.5, bold=True, color=D_TEAL)
    text(sl, BODY_X + Inches(0.20), cy + Inches(0.48), fw - Inches(0.34), Inches(0.45),
         "Low frequency dependence, stable network → does not flow off after application",
         size=11.5, color=INK2)
    # Right card: shear-thinning
    rrect(sl, cx2, cy, fw, ch, CARD_C, ACCENT, 1.2, radius=0.05)
    rect(sl, cx2, cy, Inches(0.06), ch, ACCENT)
    text(sl, cx2 + Inches(0.20), cy + Inches(0.12), fw - Inches(0.34), Inches(0.30),
         "Under high shear: viscosity drops sharply (shear-thinning)", size=12.5, bold=True, color=ACCENT)
    text(sl, cx2 + Inches(0.20), cy + Inches(0.48), fw - Inches(0.34), Inches(0.45),
         "High shear through the nozzle lowers viscosity → enables spray application",
         size=11.5, color=INK2)

    # Bottom callout
    cy_co = cy + ch + Inches(0.18)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.56),
            "Gel at rest (no dripping) + low viscosity when sprayed (deliverable) — the basis of spray application",
            dark=True, icon='✓')


# ===== SLIDE 17: shear thinning =====
def slide_17_rheology2(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 2/15", "Rheology ②: Static & Dynamic Yield Stress vs AquaGel-K (Fig. 2d, S3)")
    draw_footer(sl, "20 / 43")

    # Top callout: what is yield stress
    callout(sl, BODY_X, BODY_Y, BODY_W, Inches(0.62),
            "Yield stress = stress needed to initiate / maintain flow. Evaluated as static (onset = adhesion) and dynamic (maintain = spraying)",
            icon='ƒ')

    top = BODY_Y + Inches(0.82)
    left_w = BODY_W * 0.40 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.24)
    right_w = BODY_W - left_w - Inches(0.24)

    # Left: two figures stacked (static = S3, dynamic = 2d)
    fig_h = Inches(1.70)
    fig_placeholder(sl, BODY_X, top, left_w, fig_h,
                    fig_num="SI Fig. S3",
                    caption="Amplitude sweep: G'/G'' crossover (static yield stress)")
    fig_placeholder(sl, BODY_X, top + fig_h + Inches(0.15), left_w, fig_h,
                    fig_num="Fig. 2d",
                    caption="Dynamic yield stress via HB fit (bar chart)")

    # Right: two yield-stress cards
    card_h = Inches(1.70)
    # Static yield stress
    rrect(sl, right_x, top, right_w, card_h, CARD_C, D_TEAL, 1.2, radius=0.05)
    rect(sl, right_x, top, Inches(0.06), card_h, D_TEAL)
    text(sl, right_x + Inches(0.22), top + Inches(0.12), right_w - Inches(0.34), Inches(0.30),
         "Static yield stress σ_static (onset = adhesion)", size=13, bold=True, color=D_TEAL)
    text(sl, right_x + Inches(0.22), top + Inches(0.50), right_w - Inches(0.34), Inches(0.30),
         "Measured: stress at G' / G'' crossover in an amplitude (LAOS) sweep (S3)", size=11.5, color=INK2)
    text(sl, right_x + Inches(0.22), top + Inches(0.84), right_w - Inches(0.34), Inches(0.80),
         "WEG shows higher static yield stress than AquaGel-K\n→ reluctant to flow off vertical/elevated fuel surfaces, adheres strongly",
         size=11.5, color=INK3)

    # Dynamic yield stress
    top2 = top + card_h + Inches(0.15)
    rrect(sl, right_x, top2, right_w, card_h, CARD_C, ACCENT, 1.2, radius=0.05)
    rect(sl, right_x, top2, Inches(0.06), card_h, ACCENT)
    text(sl, right_x + Inches(0.22), top2 + Inches(0.12), right_w - Inches(0.34), Inches(0.30),
         "Dynamic yield stress σ_dynamic (maintain = spraying)", size=13, bold=True, color=ACCENT)
    text(sl, right_x + Inches(0.22), top2 + Inches(0.50), right_w - Inches(0.34), Inches(0.30),
         "Measured: HB-model fit of the flow sweep (100→0.01 s⁻¹) (2d)", size=11.5, color=INK2)
    text(sl, right_x + Inches(0.22), top2 + Inches(0.84), right_w - Inches(0.34), Inches(0.80),
         "Orders of magnitude above AquaGel-K (0.05 Pa reference line)\n→ sprayable within pump pressure yet resistant to flowing off",
         size=11.5, color=INK3)

    # Bottom callout
    cy_co = top + fig_h * 2 + Inches(0.15) + Inches(0.22)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.62),
            "For both static and dynamic yield stress, WEG greatly exceeds commercial AquaGel-K. "
            "Their small difference (weak thixotropy) lets the network re-form instantly, maintaining adhesion self-healingly",
            dark=True, icon='✓')


# ===== SLIDE 18: Herschel-Bulkley =====
def slide_18_hb(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 3/15", "Rheology ③: Aging of Dynamic Yield Stress and Long-term Stability")
    draw_footer(sl, "21 / 43")

    # Top callout
    callout(sl, BODY_X, BODY_Y, BODY_W, Inches(0.65),
            "Dynamic yield stress: obtained by fitting the steady-shear flow sweep to the Herschel–Bulkley model (σ_d) — stress to maintain flow (Fig. S7)",
            icon='ƒ')

    ty = BODY_Y + Inches(0.85)
    th = Inches(2.2)
    headers = ["Formulation", "Dynamic yield stress (Day 1)", "Dynamic yield stress (Day 455)", "Note"]
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
        ("Dynamic Yield Stress (HB fit)", "Obtained by fitting the steady-shear flow sweep to the HB model. HEC+MC/CSP: 33.34 Pa; MHEC/CSP: 3.31 Pa (Day 1)"),
        ("Long-term Stability", "MHEC/CSP shows only minor change after 455 days (3.31→4.56 Pa). Demonstrates shelf stability for practical use"),
        ("HEC+MC vs MHEC", "HEC+MC stiffens with aging (33.34→68.9 Pa). MHEC changes gently and is simpler to manufacture (single polymer)"),
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
    draw_header(sl, "Results 4/15", "Rheology ④: Storage Modulus G' and Viscoelasticity (tan δ)")
    draw_footer(sl, "22 / 43")

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


# ===== SLIDE 19b: Self-healing (step-strain recovery) =====
def slide_self_healing(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 5/15", "Self-healing: Instant G' / G'' Recovery via Step-strain")
    draw_footer(sl, "23 / 43")

    cw = (BODY_W - Inches(0.30)) / 2
    cx2 = BODY_X + cw + Inches(0.30)

    # ===== Left: Step-strain protocol concept =====
    text(sl, BODY_X, BODY_Y, cw, Inches(0.30),
         "Step-strain Protocol (ω = 1 rad/s)", size=12, bold=True, color=MUTED)

    steps = [
        ("Step 1", "Low strain (γ = 1%)", "Network preserved\nG' > G'' (gel state)", D_TEAL),
        ("Step 2", "High strain (γ = 500%)", "Network broken\nG' < G'' (liquid state)", D_AMBR),
        ("Step 3", "Return to low strain (γ = 1%)", "Particle-polymer interactions reform\nG' recovers instantly", ACCENT),
    ]
    sy = BODY_Y + Inches(0.50)
    for tag, cond, desc, c in steps:
        rrect(sl, BODY_X, sy, cw, Inches(1.05), CARD_C, LINE, 0.5, radius=0.04)
        rect(sl, BODY_X, sy, Inches(0.05), Inches(1.05), c)
        text(sl, BODY_X + Inches(0.18), sy + Inches(0.08), Inches(0.80), Inches(0.28),
             tag, size=10, bold=True, color=c)
        text(sl, BODY_X + Inches(1.00), sy + Inches(0.08), cw - Inches(1.10), Inches(0.28),
             cond, size=12, bold=True, color=INK)
        text(sl, BODY_X + Inches(0.18), sy + Inches(0.40), cw - Inches(0.28), Inches(0.60),
             desc, size=10.5, color=INK2)
        sy += Inches(1.15)

    # ===== Right: G'/G'' recovery curve placeholder =====
    text(sl, cx2, BODY_Y, cw, Inches(0.30),
         "G' / G'' Time-series (cyclic step-strain)", size=12, bold=True, color=MUTED)
    fig_placeholder(sl, cx2, BODY_Y + Inches(0.45), cw, Inches(3.4),
                    fig_num="SI (step-strain)",
                    caption="γ=1%→500%→1% cycle: G' recovers immediately")

    # Right-bottom: key point
    kp_y = BODY_Y + Inches(0.45) + Inches(3.4) + Inches(0.20)
    rrect(sl, cx2, kp_y, cw, Inches(1.10), CARD_C, LINE, 0.5, radius=0.04)
    rect(sl, cx2, kp_y, Inches(0.05), Inches(1.10), ACCENT)
    text(sl, cx2 + Inches(0.18), kp_y + Inches(0.10), cw - Inches(0.28), Inches(0.30),
         "POINT: Guarantees instant adhesion after spraying", size=12, bold=True, color=ACCENT)
    text(sl, cx2 + Inches(0.18), kp_y + Inches(0.42), cw - Inches(0.28), Inches(0.62),
         "Viscosity drops at nozzle's high shear → G' recovers instantly on substrate\n→ Holds on vertical surfaces without dripping (self-healing = key to deployability)",
         size=10.5, color=INK2)

    # Bottom callout
    cy_co = H - FTR_H - Inches(0.16) - Inches(0.60)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.60),
            "Reversible network via noncovalent bonds (H-bonds, multivalent interactions) → Self-healing demonstrated",
            icon='✓')


# ===== SLIDE 20: HEC+MC vs MHEC =====
def slide_20_compare(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 6/15", "HEC+MC vs MHEC: Comparison of Two Systems")
    draw_footer(sl, "24 / 43")

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
        "Shear-thinning (low viscosity at high shear)",
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
        "Shear-thinning behavior",
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


# ===== SLIDE SDS RHEOLOGY: Fig. 2e–h (SDS Addition Effect on Rheology) =====
def slide_sds_rheology(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 7/15", "SDS Addition: Goal Was Better Foaming & Wetting — But Results Were Mixed (Fig. 2e–h)")
    draw_footer(sl, "25 / 43")

    # ── Top: purpose 2 cards (left = purpose, right = expectation) ──
    hdr_h = Inches(0.90)
    cw2 = (BODY_W - Inches(0.18)) / 2
    # Left: purpose
    rrect(sl, BODY_X, BODY_Y, cw2, hdr_h, CARD_C, D_TEAL, 1.2, radius=0.05)
    rect(sl, BODY_X, BODY_Y, Inches(0.06), hdr_h, D_TEAL)
    text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(0.08), cw2 - Inches(0.30), Inches(0.28),
         "Why SDS was added", size=12, bold=True, color=D_TEAL)
    text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(0.40), cw2 - Inches(0.30), Inches(0.45),
         "Surfactant SDS (0 / 0.1 / 0.5 wt%) was added to HEC+MC/CSP to improve wetting "
         "on fuel surfaces and promote stable foam-layer formation during combustion",
         size=11, color=INK2)
    # Right: expectation
    cx2_hdr = BODY_X + cw2 + Inches(0.18)
    rrect(sl, cx2_hdr, BODY_Y, cw2, hdr_h, CARD_C, ACCENT, 1.2, radius=0.05)
    rect(sl, cx2_hdr, BODY_Y, Inches(0.06), hdr_h, ACCENT)
    text(sl, cx2_hdr + Inches(0.20), BODY_Y + Inches(0.08), cw2 - Inches(0.30), Inches(0.28),
         "Expected outcome", size=12, bold=True, color=ACCENT)
    text(sl, cx2_hdr + Inches(0.20), BODY_Y + Inches(0.40), cw2 - Inches(0.30), Inches(0.45),
         "① Lower gel viscosity to improve sprayability  "
         "② Stabilise bubbles to increase Foaming Index",
         size=11, color=INK2)

    # ── Center: 2×2 grid (Fig. 2e–h) ──
    gap = Inches(0.18)
    cw = (BODY_W - gap) / 2
    row1_y = BODY_Y + hdr_h + Inches(0.16)
    row_h  = Inches(2.05)
    row2_y = row1_y + row_h + gap

    panels = [
        (BODY_X,            row1_y, "Fig. 2e", "G' / G'' vs angular frequency (SAOS, 3 SDS levels)"),
        (BODY_X + cw + gap, row1_y, "Fig. 2f", "G' and tan δ bar chart (at 1 rad/s)"),
        (BODY_X,            row2_y, "Fig. 2g", "Viscosity vs shear rate (Herschel–Bulkley fit)"),
        (BODY_X + cw + gap, row2_y, "Fig. 2h", "Dynamic yield stress by SDS level (0.05 Pa reference)"),
    ]
    for px, py, fig_num, cap in panels:
        fig_placeholder(sl, px, py, cw, row_h, fig_num=fig_num, caption=cap)

    # ── Bottom: result assessment (2 columns) ──
    band_y = row2_y + row_h + Inches(0.16)
    band_h = H - band_y - FTR_H - Inches(0.16)
    res_w = (BODY_W - Inches(0.18)) / 2

    # Left: rheology (limited improvement)
    rrect(sl, BODY_X, band_y, res_w, band_h, CARD_C, D_AMBR, 1.2, radius=0.04)
    rect(sl, BODY_X, band_y, Inches(0.06), band_h, D_AMBR)
    text(sl, BODY_X + Inches(0.20), band_y + Inches(0.06), res_w - Inches(0.30), Inches(0.26),
         "Rheology: concentration-dependent but limited benefit", size=11.5, bold=True, color=D_AMBR)
    text(sl, BODY_X + Inches(0.20), band_y + Inches(0.34), res_w - Inches(0.30), band_h - Inches(0.38),
         "G' rises slightly at 0.1 wt% but drops again at 0.5 wt%; yield stress also decreases. "
         "Overall, SDS addition negatively affects gel stiffness and adhesion",
         size=11, color=INK3)

    # Right: foaming (counterproductive)
    cx2_res = BODY_X + res_w + Inches(0.18)
    rrect(sl, cx2_res, band_y, res_w, band_h, CARD_C, D_RED, 1.2, radius=0.04)
    rect(sl, cx2_res, band_y, Inches(0.06), band_h, D_RED)
    text(sl, cx2_res + Inches(0.20), band_y + Inches(0.06), res_w - Inches(0.30), Inches(0.26),
         "Foaming: Foaming Index actually decreased (Fig. 4)", size=11.5, bold=True, color=D_RED)
    text(sl, cx2_res + Inches(0.20), band_y + Inches(0.34), res_w - Inches(0.30), band_h - Inches(0.38),
         "SDS caused bubble coarsening and non-uniform foam structure, lowering the Foaming Index. "
         "The original goal of enhancing foam formation was not achieved",
         size=11, color=INK3)


# ===== SLIDE 21: 燃焼試験 setup =====
def slide_21_setup(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 8/15", "Combustion Test Setup")
    draw_footer(sl, "26 / 43")

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
    draw_header(sl, "Results 9/15", "Time to Char: Quantitative Comparison of 5 Formulations")
    draw_footer(sl, "27 / 43")

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
    draw_header(sl, "Results 10/15", "Time-lapse Observation of Combustion Process")
    draw_footer(sl, "28 / 43")

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
    draw_header(sl, "Results 11/15", "Surface Condition Comparison at 120 s / 300 s")
    draw_footer(sl, "29 / 43")

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
         "30 / 43", size=10, color=RGBColor(0x55, 0x66, 0x77), align=PP_ALIGN.RIGHT)


# ===== SLIDE 26: Foaming Index =====
def slide_26_foam(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 12/15", "SDS Addition: Expected to Improve Foaming — But the Opposite Occurred (Fig. 4)")
    draw_footer(sl, "31 / 43")

    # ── Top: Hypothesis → Experiment → Unexpected Result (3 boxes) ──
    bw = Inches(3.6)
    bh = Inches(1.10)
    arrow_w = Inches(0.70)
    gap = (BODY_W - bw * 3 - arrow_w * 2) / 2
    box_y = BODY_Y

    # Hypothesis box
    rrect(sl, BODY_X, box_y, bw, bh, CARD_C, D_TEAL, 1.2, radius=0.05)
    rect(sl, BODY_X, box_y, Inches(0.06), bh, D_TEAL)
    text(sl, BODY_X + Inches(0.22), box_y + Inches(0.08), bw - Inches(0.30), Inches(0.24),
         "Hypothesis (purpose of SDS)", size=10, bold=True, color=D_TEAL)
    text(sl, BODY_X + Inches(0.22), box_y + Inches(0.38), bw - Inches(0.30), Inches(0.64),
         "Surfactant SDS promotes bubble formation\n→ Foaming Index increases\n→ Thicker aerogel insulation layer",
         size=11, color=INK2)

    # → Arrow 1
    ax1 = BODY_X + bw + gap * 0.5
    ar1 = sl.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, ax1, box_y + bh/2 - Inches(0.22),
                              arrow_w, Inches(0.44))
    ar1.fill.solid(); ar1.fill.fore_color.rgb = MUTED
    ar1.line.fill.background(); ar1.shadow.inherit = False

    # Experiment box
    bx2 = BODY_X + bw + gap * 0.5 + arrow_w + gap * 0.5
    rrect(sl, bx2, box_y, bw, bh, CARD_C, LINE, 0.8, radius=0.05)
    text(sl, bx2 + Inches(0.22), box_y + Inches(0.08), bw - Inches(0.30), Inches(0.24),
         "Experiment", size=10, bold=True, color=MUTED)
    text(sl, bx2 + Inches(0.22), box_y + Inches(0.38), bw - Inches(0.30), Inches(0.64),
         "SDS 0 / 0.1 / 0.5 wt% added to\nHEC+MC/CSP; combustion tested\n→ Foaming Index quantified (Fig. 4a/4b)",
         size=11, color=INK2)

    # → Arrow 2 (amber: signals unexpected)
    ax2 = bx2 + bw + gap * 0.5
    ar2 = sl.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, ax2, box_y + bh/2 - Inches(0.22),
                              arrow_w, Inches(0.44))
    ar2.fill.solid(); ar2.fill.fore_color.rgb = D_AMBR
    ar2.line.fill.background(); ar2.shadow.inherit = False

    # Result box (unexpected)
    bx3 = ax2 + arrow_w + gap * 0.5
    rrect(sl, bx3, box_y, bw, bh, RGBColor(0xff, 0xf7, 0xed), D_AMBR, 1.4, radius=0.05)
    rect(sl, bx3, box_y, Inches(0.06), bh, D_AMBR)
    text(sl, bx3 + Inches(0.22), box_y + Inches(0.08), bw - Inches(0.30), Inches(0.24),
         "Result (opposite of expected)", size=10, bold=True, color=D_AMBR)
    text(sl, bx3 + Inches(0.22), box_y + Inches(0.38), bw - Inches(0.30), Inches(0.64),
         "SDS addition reduced Foaming Index\n2.6 (0%) → 2.3 (0.1%) → even lower (0.5%)\nSDS did not improve foaming",
         size=11, color=INK2)

    # ── Middle: Fig. 4a / 4b placeholders ──
    fig_y = box_y + bh + Inches(0.28)
    fig_h = Inches(2.10)
    fw = (BODY_W - Inches(0.20)) / 2
    fig_placeholder(sl, BODY_X, fig_y, fw, fig_h,
                    fig_num="Fig. 4a", caption="Foamed layer appearance after combustion (by SDS level)")
    fig_placeholder(sl, BODY_X + fw + Inches(0.20), fig_y, fw, fig_h,
                    fig_num="Fig. 4b", caption="Foaming Index per formulation (quantified)")

    # ── Bottom: Foaming Index bar chart ──
    bar_y = fig_y + fig_h + Inches(0.28)
    bar_data = [
        ("HEC+MC/CSP 1-5 (SDS 0%)", 1.00, "≈2.6×", D_TEAL),
        ("HEC+MC/CSP/SDS 1-5-0.1", 0.88, "≈2.3×", MUTED),
        ("HEC+MC/CSP/SDS 1-5-0.5", 0.72, "even lower", D_AMBR),
        ("AquaGel-K (reference)", 0.02, "≈0", GRAY_L),
    ]
    text(sl, BODY_X, bar_y, BODY_W * 0.5, Inches(0.26),
         "Foaming Index (post-combustion thickness / initial thickness)", size=11, bold=True, color=MUTED)
    by = bar_y + Inches(0.32)
    for label, ratio, val, color in bar_data:
        bar_row(sl, BODY_X, by, BODY_W * 0.62, label, ratio, val,
                fill_color=color, lbl_w=Inches(3.2), val_w=Inches(1.0))
        by += Inches(0.40)

    # ── Bottom callout ──
    cy_co = H - FTR_H - Inches(0.14) - Inches(0.58)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.58),
            "SDS did not improve foaming — instead coarsened bubble structure (Fig. 4c SEM) → explains the Foaming Index decrease",
            icon='!', dark=True)


# ===== SLIDE 27: SDS濃度別SEM =====
def slide_27_sem_sds(prs):
    sl = new_slide(prs)
    draw_header(sl, "Results 13/15", "Effect of SDS Concentration on Aerogel Microstructure")
    draw_footer(sl, "32 / 43")

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
    draw_header(sl, "Results 14/15", "FT-IR & XPS: Chemical Composition Changes (Before/After Combustion)")
    draw_footer(sl, "33 / 43")

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
    draw_header(sl, "Results 15/15", "TGA/DSC: Thermal Decomposition Profile and Role of Each Component")
    draw_footer(sl, "34 / 43")

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
    draw_header(sl, "Supplementary 1/3", "Mechanism of Silica Aerogel Formation")
    draw_footer(sl, "35 / 43")

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
    draw_header(sl, "Supplementary 2/3", "Sintering Progression with Burn Time (SEM)")
    draw_footer(sl, "36 / 43")

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


# ===== SLIDE 32: Aging & Long-term Stability (Supplementary 3/3) =====
def slide_32_aging(prs):
    sl = new_slide(prs)
    draw_header(sl, "Supplementary 3/3", "Aging & Long-Term Stability: Performance Retained Beyond One Year")
    draw_footer(sl, "38 / 43")

    left_w = BODY_W * 0.36 - Inches(0.12)
    right_w = BODY_W * 0.64 - Inches(0.08)
    right_x = BODY_X + left_w + Inches(0.20)

    # Left: aging mechanism
    rect(sl, BODY_X, BODY_Y, Inches(0.05), Inches(1.02), ACCENT)
    text(sl, BODY_X + Inches(0.18), BODY_Y, left_w - Inches(0.18), Inches(0.34),
         "Aging Mechanism", size=15, bold=True, color=INK)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(0.40), left_w - Inches(0.18), Inches(0.66),
         "Transition from reversible (hydrogen) bonding\nto irreversible covalent (Si–O) conjugation", size=12, color=INK2)

    # Divider line
    rect(sl, BODY_X + Inches(0.18), BODY_Y + Inches(1.24), left_w * 0.85, Inches(0.015), LINE)

    # Two systems compared
    rows = [
        ("HEC+MC / CSP", "Extensive covalent conjugation → stiffens", ACCENT),
        ("MHEC / CSP", "Stays largely reversible → subtle change", D_TEAL),
    ]
    ry = BODY_Y + Inches(1.46)
    for name, desc, c in rows:
        text(sl, BODY_X + Inches(0.18), ry, left_w - Inches(0.18), Inches(0.30),
             name, size=13, bold=True, color=c)
        text(sl, BODY_X + Inches(0.18), ry + Inches(0.30), left_w - Inches(0.18), Inches(0.46),
             desc, size=11, color=INK3)
        ry += Inches(0.92)

    text(sl, BODY_X + Inches(0.18), ry, left_w - Inches(0.18), Inches(0.30),
         "Fig. S4 / S6 (interfacial chemistry)", size=10, color=MUTED, italic=True)

    # Right top: two mini stats
    mw = (right_w - Inches(0.16)) / 2
    mh = Inches(1.5)
    mini(sl, right_x, BODY_Y, mw, mh, "Modulus (20-day aging)", "~1000 Pa",
         "HEC+MC/CSP · Fig S5", value_color=ACCENT)
    mini(sl, right_x + mw + Inches(0.16), BODY_Y, mw, mh, "Yield stress Day1→Day455", "33 → 69 Pa",
         "HEC+MC/CSP · Fig S7", value_color=AMBER)

    # Right bottom: aging rheology placeholder (split into Day 1 / Day 455)
    fy = BODY_Y + mh + Inches(0.28)
    fh = Inches(2.3)
    sub_w = (right_w - Inches(0.16)) / 2
    fig_placeholder(sl, right_x, fy, sub_w, fh,
                    fig_num="Fig. S7 (Day 1)",
                    caption="Flow sweep · HB fit (fresh)")
    fig_placeholder(sl, right_x + sub_w + Inches(0.16), fy, sub_w, fh,
                    fig_num="Fig. S7 (Day 455)",
                    caption="Flow sweep · HB fit (455 days aged)")

    # Bottom: hero callout
    cy_co = H - FTR_H - Inches(0.16) - Inches(0.82)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.82),
            "After Day 455 (~15 months), combustion behavior and aerogel transition remain unchanged — proven shelf stability (Video S5)",
            dark=True, icon='✓')


# ===== SLIDE 33: 実装シナリオ =====
def slide_33_scenarios(prs):
    sl = new_slide(prs)
    draw_header(sl, "Discussion 1/3", "Implementation Scenarios: WEG Deployment Strategy")
    draw_footer(sl, "40 / 43")

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
    draw_header(sl, "Summary 1/2", "Comparison with Existing Products: Only WEG Meets Every Requirement")
    draw_footer(sl, "41 / 43")

    # Top callout
    callout(sl, BODY_X, BODY_Y, BODY_W, Inches(0.56),
            "This WEG (PP hydrogel) simultaneously satisfies requirements that water, commercial WEG, and Phos-Chek each meet only partially",
            icon='◆')

    # ── ○× comparison matrix ──
    ty = BODY_Y + Inches(0.78)
    crit = ["Long flame\nprotection", "Protects after\ndrying", "Sprayable", "Low environ.\npersistence", "Forms foam /\ninsulating layer"]
    G = D_GRN
    A = D_AMBR
    R = D_RED
    rows = [
        ("WEG (this work)", True,  [("◎", G), ("◎", G), ("○", G), ("○", G), ("◎", G)]),
        ("AquaGel-K (commercial)", False, [("△", A), ("×", R), ("○", G), ("○", G), ("×", R)]),
        ("Phos-Chek", False,     [("○", G), ("○", G), ("○", G), ("×", R), ("×", R)]),
        ("Water only", False,    [("×", R), ("×", R), ("◎", G), ("◎", G), ("×", R)]),
    ]

    name_w = Inches(2.9)
    grid_w = BODY_W - name_w
    col_w = grid_w / len(crit)
    hdr_h = Inches(0.95)
    row_h = Inches(0.66)
    table_h = hdr_h + row_h * len(rows)

    # Header row (criteria)
    rect(sl, BODY_X, ty, name_w, hdr_h, HEADER)
    text(sl, BODY_X + Inches(0.12), ty, name_w - Inches(0.24), hdr_h,
         "Product ＼ Requirement", size=11, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    for ci, c in enumerate(crit):
        cx = BODY_X + name_w + col_w * ci
        rect(sl, cx, ty, col_w, hdr_h, HEADER, WHITE, 0.5)
        text(sl, cx + Inches(0.06), ty + Inches(0.05), col_w - Inches(0.12), hdr_h - Inches(0.10),
             c, size=10, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

    # Data rows
    for ri, (name, emph, marks) in enumerate(rows):
        ry = ty + hdr_h + row_h * ri
        name_bg = RGBColor(0xec, 0xf2, 0xf8) if emph else (SOFT if ri % 2 == 0 else WHITE)
        rect(sl, BODY_X, ry, name_w, row_h, name_bg, LINE, 0.3)
        if emph:
            rect(sl, BODY_X, ry, Inches(0.06), row_h, D_GRN)
        text(sl, BODY_X + Inches(0.20), ry, name_w - Inches(0.30), row_h,
             name, size=11.5, bold=emph, color=(INK if emph else INK2),
             anchor=MSO_ANCHOR.MIDDLE)
        for ci, (sym, col) in enumerate(marks):
            cx = BODY_X + name_w + col_w * ci
            rect(sl, cx, ry, col_w, row_h, name_bg, LINE, 0.3)
            text(sl, cx, ry, col_w, row_h, sym, size=18, bold=True,
                 color=col, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

    # Legend
    leg_y = ty + table_h + Inches(0.16)
    legend = [("◎", "excellent", D_GRN), ("○", "capable", D_GRN), ("△", "limited", D_AMBR), ("×", "not feasible", D_RED)]
    lx = BODY_X
    for sym, lbl, col in legend:
        text(sl, lx, leg_y, Inches(0.4), Inches(0.30), sym, size=14, bold=True,
             color=col, anchor=MSO_ANCHOR.MIDDLE)
        text(sl, lx + Inches(0.38), leg_y, Inches(1.7), Inches(0.30), lbl, size=11,
             color=INK3, anchor=MSO_ANCHOR.MIDDLE)
        lx += Inches(2.1)


# ===== SLIDE 35: Conclusions =====
def slide_35_summary(prs):
    sl = new_slide(prs)
    draw_header(sl, "Summary 2/2", "Conclusions: Core Mechanism and Four Outcomes")
    draw_footer(sl, "42 / 43")

    # ── Top: core mechanism as a 3-step horizontal flow ──
    text(sl, BODY_X, BODY_Y, BODY_W, Inches(0.30),
         "Core mechanism: the gel \"self-transforms\" upon flame contact", size=13, bold=True, color=INK)
    fy = BODY_Y + Inches(0.42)
    fh = Inches(1.65)
    steps = [
        ("STEP 1", "Water evaporates", "Flame contact vaporises water in the gel (endothermic)\n→ suppresses substrate heating", D_TEAL),
        ("STEP 2", "CSP sinters", "Silica particles sinter under heat,\nforming inter-particle necks (bridges)", D_AMBR),
        ("STEP 3", "Aerogel insulation", "A porous silica aerogel forms in situ\n→ insulates the substrate even after drying", ACCENT),
    ]
    n = len(steps)
    arrow_w = Inches(0.45)
    sw = (BODY_W - arrow_w * (n - 1)) / n
    for i, (tag, title, desc, c) in enumerate(steps):
        sx = BODY_X + i * (sw + arrow_w)
        rrect(sl, sx, fy, sw, fh, CARD_C, LINE, 0.5, radius=0.04)
        rect(sl, sx, fy, sw, Inches(0.05), c)
        text(sl, sx + Inches(0.18), fy + Inches(0.14), sw - Inches(0.36), Inches(0.26),
             tag, size=10, bold=True, color=c)
        text(sl, sx + Inches(0.18), fy + Inches(0.44), sw - Inches(0.36), Inches(0.36),
             title, size=15, bold=True, color=INK)
        text(sl, sx + Inches(0.18), fy + Inches(0.86), sw - Inches(0.36), fh - Inches(0.95),
             desc, size=11, color=INK3)
        if i < n - 1:
            ax = sx + sw
            text(sl, ax, fy, arrow_w, fh, "▶", size=18, bold=True,
                 color=GRAY_L, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

    # ── Middle: four key outcomes ──
    ry = fy + fh + Inches(0.28)
    text(sl, BODY_X, ry, BODY_W, Inches(0.30),
         "Four key outcomes", size=13, bold=True, color=INK)
    ry += Inches(0.40)
    results = [
        ("3–6×", "protection time vs commercial", D_GRN),
        ("Self-forming", "aerogel insulation in situ", ACCENT),
        ("Sprayable", "compatible with existing infra", D_TEAL),
        ("Sustainable", "cellulose-based, safe materials", D_AMBR),
    ]
    rcw = (BODY_W - Inches(0.42)) / 4
    rch = Inches(1.30)
    for i, (big, lbl, c) in enumerate(results):
        cx = BODY_X + i * (rcw + Inches(0.14))
        rrect(sl, cx, ry, rcw, rch, SOFT, LINE, 0.5, radius=0.04)
        rect(sl, cx, ry, Inches(0.06), rch, c)
        text(sl, cx + Inches(0.20), ry + Inches(0.16), rcw - Inches(0.34), Inches(0.50),
             big, size=20, bold=True, color=c)
        text(sl, cx + Inches(0.20), ry + Inches(0.74), rcw - Inches(0.34), rch - Inches(0.84),
             lbl, size=11.5, color=INK3)

    # ── Bottom: key message ──
    cy_co = ry + rch + Inches(0.22)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.62),
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
        ("Shear-thinning", 0.60, "demonstrated", GRAY_L),
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
         "42 / 43", size=10, color=RGBColor(0x55, 0x66, 0x77), align=PP_ALIGN.RIGHT)


# ===== STANDALONE: MHEC Advantages (Aging Stability + Cost) =====
def slide_mhec_advantage(prs):
    sl = new_slide(prs)
    draw_header(sl, "Supplementary", "MHEC Advantages: Long-term Stability and Cost")
    draw_footer(sl, "37 / 43")

    cw = (BODY_W - Inches(0.28)) / 2
    cx2 = BODY_X + cw + Inches(0.28)

    # ===== Left column: Aging stability (SI Section 2-3 / Fig. S7) =====
    rect(sl, BODY_X, BODY_Y, Inches(0.06), Inches(2.80), D_TEAL)
    text(sl, BODY_X + Inches(0.20), BODY_Y, cw - Inches(0.20), Inches(0.36),
         "① Gentle aging behavior (SI Fig. S7)", size=14, bold=True, color=INK)
    text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(0.42), cw - Inches(0.20), Inches(0.30),
         "Yield stress change: Day 1 → Day 455", size=11, bold=True, color=MUTED)

    bar_data = [
        ("MHEC / CSP", 0.35, "+38 %\n(3.31 → 4.56 Pa)", D_TEAL),
        ("HEC+MC / CSP", 1.00, "+107 %\n(33.3 → 68.9 Pa)", D_AMBR),
    ]
    by = BODY_Y + Inches(0.88)
    for label, ratio, val, c in bar_data:
        text(sl, BODY_X + Inches(0.20), by, cw - Inches(0.20), Inches(0.28),
             label, size=11, bold=True, color=INK2)
        bx = BODY_X + Inches(0.20)
        bw = cw - Inches(0.50)
        rrect(sl, bx, by + Inches(0.30), bw, Inches(0.22), LINE, None, radius=0.4)
        rrect(sl, bx, by + Inches(0.30), bw * ratio, Inches(0.22), c, None, radius=0.4)
        text(sl, bx + bw * ratio + Inches(0.12), by + Inches(0.22),
             Inches(1.4), Inches(0.40), val, size=10, bold=True, color=c)
        by += Inches(0.90)

    rrect(sl, BODY_X + Inches(0.10), by + Inches(0.10), cw - Inches(0.20), Inches(0.90),
          SOFT, D_TEAL, 0.8, radius=0.04)
    text(sl, BODY_X + Inches(0.24), by + Inches(0.22), cw - Inches(0.46), Inches(0.72),
         "MHEC system remains primarily noncovalent → properties stable over 1+ year\n"
         "(HEC+MC system undergoes Si–O covalent condensation → stiffness increases)",
         size=11, color=INK2)
    text(sl, BODY_X + Inches(0.20), by + Inches(1.18), cw - Inches(0.40), Inches(0.22),
         "Source: SI Fig. S7 (flow sweep · HB model fit)", size=9, italic=True, color=MUTED)

    rect(sl, BODY_X, BODY_Y + Inches(2.94), cw, Inches(0.015), LINE)

    text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(3.06), cw - Inches(0.20), Inches(0.26),
         "Why stable: Si–O covalent condensation is suppressed", size=11, bold=True, color=D_TEAL)
    text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(3.38), cw - Inches(0.20), Inches(0.60),
         "In HEC+MC/CSP, Si–O condensation between CSP and polymer\n"
         "progresses over time, increasing stiffness. In MHEC/CSP,\n"
         "this condensation is suppressed, giving gentle property drift.",
         size=10.5, color=INK2)

    # ===== Right column: Cost & manufacturing =====
    rect(sl, cx2, BODY_Y, Inches(0.06), Inches(4.20), D_AMBR)
    text(sl, cx2 + Inches(0.20), BODY_Y, cw - Inches(0.20), Inches(0.36),
         "② Manufacturing cost & quality control", size=14, bold=True, color=INK)

    cost_items = [
        ("Single polymer",
         "No need to source and blend HEC + MC separately\n→ Fewer raw material types"),
        ("Simplified formulation",
         "No mixing ratio tuning or homogenization of 2 components\n→ Reduced batch-to-batch variability"),
        ("Lower QC burden",
         "Only one raw material specification to manage\n→ Cost advantage becomes clear at scale-up"),
        ("Logistics / storage",
         "Only one material to store and transport\n→ Streamlined supply chain"),
    ]
    iy = BODY_Y + Inches(0.52)
    for title, desc in cost_items:
        rrect(sl, cx2 + Inches(0.10), iy, cw - Inches(0.20), Inches(0.96),
              CARD_C, LINE, 0.5, radius=0.04)
        text(sl, cx2 + Inches(0.26), iy + Inches(0.12), cw - Inches(0.50), Inches(0.28),
             title, size=12, bold=True, color=D_AMBR)
        text(sl, cx2 + Inches(0.26), iy + Inches(0.44), cw - Inches(0.50), Inches(0.46),
             desc, size=10.5, color=INK2)
        iy += Inches(1.06)

    cy_co = H - FTR_H - Inches(0.16) - Inches(0.80)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.80),
            "Note: No quantitative cost comparison data is reported in this paper. "
            "MHEC's advantage is based on its single-polymer nature and measured aging stability (SI Fig. S7).",
            icon='ⓘ')


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
                          "3 / 43")
    slide_04_wildfire(prs)
    slide_05_process(prs)
    slide_06_existing(prs)
    slide_07_predecessor(prs)
    slide_07_core(prs)
    slide_section_divider(prs, "02", "Part 2", "Materials & Methods",
                          "Novel gel design combining cellulosic polymers and colloidal silica",
                          ["Polymer Components", "Silica Particles & Surfactant", "Formulations (5 types)", "Evaluation Methods"],
                          "9 / 43")
    slide_mat_composition(prs)
    slide_self_protection(prs)
    slide_09_polymers(prs)
    slide_10_mc(prs)
    slide_11_csp_sds(prs)
    slide_12_formulations(prs)
    slide_13_methods(prs)
    slide_14_adhesion(prs)
    slide_section_divider(prs, "03", "Part 3", "Results",
                          "Rheological properties, combustion performance, foam structure, and aerogel formation mechanism",
                          ["Rheology", "Combustion Test", "Foaming Index", "SEM Observation"],
                          "18 / 43")
    slide_16_rheology1(prs)
    slide_17_rheology2(prs)
    slide_18_hb(prs)
    slide_19_yield(prs)
    slide_self_healing(prs)
    slide_20_compare(prs)
    slide_sds_rheology(prs)
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
    slide_32_aging(prs)
    slide_mhec_advantage(prs)
    slide_section_divider(prs, "04", "Part 4", "Discussion & Conclusion",
                          "Implementation scenarios, comprehensive comparison with existing products, research significance and future prospects",
                          ["Implementation Scenarios", "Performance Comparison", "Summary"],
                          "39 / 43")
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
