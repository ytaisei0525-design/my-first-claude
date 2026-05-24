#!/usr/bin/env python3
"""slides.pptx を1から構築（HTMLデザイン準拠、論文図はプレースホルダー）"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

# ── 色定義（HTMLのCSS変数と一致） ─────────────────────────────
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
    """複数の書式付きランを1段落に追加"""
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
    """論文の図のプレースホルダー（ユーザーが後で画像を貼り付ける用）"""
    # 背景（ダッシュではなく実線、薄い色で）
    s = rrect(sl, x, y, w, h, PH_BG, PH_BD, 1.0, radius=0.02)
    # 図番号バッジ（左上）
    if fig_num:
        tag_w = Inches(0.95)
        rect(sl, x + Inches(0.12), y + Inches(0.12), tag_w, Inches(0.32),
             HEADER)
        text(sl, x + Inches(0.12), y + Inches(0.12), tag_w, Inches(0.32),
             fig_num, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 中央プレースホルダーテキスト
    text(sl, x, y, w, h,
         "（論文の図を貼り付け）", size=12, color=GRAY_L,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # キャプション
    if caption:
        text(sl, x, y + h + Inches(0.04), w, Inches(0.30),
             caption, size=10, color=MUTED, italic=True,
             align=PP_ALIGN.CENTER)


def section_label(sl, x, y, w, label):
    """ラベル文字（カードグループの上のメタ情報）"""
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
    # 背景バー
    rrect(sl, bx, y + Inches(0.08), bw, Inches(0.16),
          CARD_C, LINE, 0.3, radius=0.5)
    # フィル
    fill_w = bw * fill_ratio
    if fill_w > Inches(0.05):
        rrect(sl, bx, y + Inches(0.08), fill_w, Inches(0.16),
              fill_color, None, radius=0.5)
    # 値
    text(sl, bx + bw + Inches(0.10), y, val_w, Inches(0.30),
         value, size=11, color=INK3, bold=True, anchor=MSO_ANCHOR.MIDDLE)


def simple_table(sl, x, y, w, h, headers, rows, col_widths=None):
    """シンプルなテーブル"""
    n_cols = len(headers)
    if col_widths is None:
        col_widths = [w / n_cols] * n_cols
    else:
        # 比率指定
        total = sum(col_widths)
        col_widths = [w * (cw / total) for cw in col_widths]

    n_rows = len(rows) + 1
    row_h = h / n_rows

    # ヘッダー
    rect(sl, x, y, w, row_h, HEADER)
    cx = x
    for i, hd in enumerate(headers):
        text(sl, cx + Inches(0.10), y, col_widths[i] - Inches(0.10), row_h,
             hd, size=11, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[i]

    # データ行
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

    # トップタグ
    text(sl, Inches(0.5), Inches(0.5), Inches(8), Inches(0.4),
         "RESEARCH PRESENTATION  ·  2024", size=11, bold=True,
         color=RGBColor(0x88, 0x9a, 0xb0))
    text(sl, W - Inches(5.5), Inches(0.5), Inches(5), Inches(0.4),
         "Adv. Mater.  ·  Dong et al., 2024", size=11,
         color=RGBColor(0x88, 0x9a, 0xb0), align=PP_ALIGN.RIGHT)

    # タイトル
    text(sl, Inches(0.8), Inches(1.7), Inches(11.7), Inches(0.45),
         "WATER-ENHANCING GELS  ·  ADVANCED MATERIALS 2024", size=12, bold=True,
         color=RGBColor(0x88, 0x9a, 0xb0), align=PP_ALIGN.CENTER)

    text(sl, Inches(0.5), Inches(2.3), Inches(12.3), Inches(1.2),
         "山火事から守る", size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    text(sl, Inches(0.5), Inches(3.2), Inches(12.3), Inches(1.0),
         "熱活性化シリカエアロゲル", size=44, bold=True, color=GOLD, align=PP_ALIGN.CENTER)

    text(sl, Inches(1.0), Inches(4.5), Inches(11.3), Inches(1.0),
         "セルロース系ポリマー × コロイダルシリカ粒子で実現した、\n加熱時に自己発泡してエアロゲル層を形成する次世代難燃ゲル",
         size=15, color=RGBColor(0xb8, 0xcc, 0xe2), align=PP_ALIGN.CENTER)

    # メタ情報
    meta = "材料化学 / 防火工学    ·    Stanford University, Appel Lab    ·    Adv. Mater. 36, 2407375"
    text(sl, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.4),
         meta, size=12, color=RGBColor(0x88, 0x9c, 0xb4), align=PP_ALIGN.CENTER)

    text(sl, Inches(0.5), Inches(6.6), Inches(8), Inches(0.3),
         "DOI: 10.1002/adma.202407375", size=10, color=RGBColor(0x66, 0x77, 0x88))
    text(sl, W - Inches(1.4), H - Inches(0.36), Inches(1.2), Inches(0.26),
         "1 / 36", size=10, color=RGBColor(0x66, 0x77, 0x88), align=PP_ALIGN.RIGHT)


# ===== SLIDE 2: TOC =====
def slide_02_toc(prs):
    sl = new_slide(prs)
    draw_header(sl, "Contents", "発表の構成")
    draw_footer(sl, "2 / 36")

    parts = [
        ("PART 1", "研究背景", "山火事の現状、既存技術（Phos-Chek, AquaGel-K）の限界、本研究の革新点", "Slides 4–7"),
        ("PART 2", "材料・方法", "セルロース系ポリマー、CSP、配合系（5種）、評価手法", "Slides 9–14"),
        ("PART 3", "実験結果", "レオロジー、燃焼試験、発泡指数、SEM観察、メカニズム", "Slides 16–33"),
        ("PART 4", "考察・まとめ", "実装シナリオ、既存品との比較、結論と今後の展望", "Slides 35–36"),
    ]
    cw = (BODY_W - Inches(0.42)) / 4
    cy = BODY_Y
    ch = Inches(4.2)
    for i, (tag, title, body, sl_range) in enumerate(parts):
        cx = BODY_X + i * (cw + Inches(0.14))
        rrect(sl, cx, cy, cw, ch, CARD_C, LINE, 0.5, radius=0.04)
        rect(sl, cx, cy, cw, Inches(0.04), ACCENT)
        # icon (using text glyph)
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
            "キーメッセージ：加熱時の発泡＋シリカ粒子の焼結によりエアロゲル層が自己形成され、市販品を超える長時間保護を実現",
            icon='ⓘ')


# ===== SLIDE 3: SECTION DIVIDER 01 =====
def slide_section_divider(prs, num_str, num_label, title, sub, pills, page_num):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, HEADER)
    rect(sl, 0, 0, Inches(0.10), H, AMBER)

    # 大きな背景数字
    text(sl, W - Inches(5.5), Inches(0.5), Inches(5), Inches(7.0),
         num_str, size=240, bold=True, color=RGBColor(0x24, 0x30, 0x44),
         align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.BOTTOM)

    text(sl, Inches(1.0), Inches(1.7), Inches(10), Inches(0.4),
         num_label, size=13, bold=True, color=RGBColor(0x88, 0xa4, 0xc0))
    text(sl, Inches(1.0), Inches(2.2), Inches(10), Inches(1.4),
         title, size=52, bold=True, color=WHITE)
    text(sl, Inches(1.0), Inches(3.7), Inches(10), Inches(1.0),
         sub, size=17, color=RGBColor(0x9c, 0xb4, 0xcc))

    # ピル
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
    draw_header(sl, "背景 1/4", "山火事被害の深刻化とWUIリスク")
    draw_footer(sl, "4 / 36")

    # 左：stats + cards、右：図プレースホルダー
    left_w = BODY_W * 0.62 - Inches(0.10)
    right_w = BODY_W * 0.38 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    # stats 3つ
    sw = (left_w - Inches(0.20)) / 3
    sy = BODY_Y
    sh = Inches(1.7)
    stat(sl, BODY_X, sy, sw, sh, "5", unit="倍",
         label="大規模山火事の増加", sub="過去30年で焼失面積拡大",
         num_size=46, lbl_size=12, sub_size=10)
    stat(sl, BODY_X + sw + Inches(0.10), sy, sw, sh, "400", unit="万km²",
         label="年間焼失面積（世界）", sub="気候変動・干ばつが要因",
         num_size=40, lbl_size=12, sub_size=10)
    stat(sl, BODY_X + (sw + Inches(0.10)) * 2, sy, sw, sh, "4500", unit="万人",
         label="米国WUI居住者", sub="境界帯のリスク層",
         num_size=42, lbl_size=12, sub_size=10)

    # WUI用語カード
    ty = sy + sh + Inches(0.20)
    th = Inches(1.0)
    rrect(sl, BODY_X, ty, left_w, th, WHITE, LINE, 0.5, radius=0.04)
    text(sl, BODY_X + Inches(0.20), ty + Inches(0.14), left_w - Inches(0.4), Inches(0.30),
         "用語", size=11, bold=True, color=MUTED)
    text(sl, BODY_X + Inches(0.20), ty + Inches(0.42), left_w - Inches(0.4), Inches(0.55),
         "WUI（Wildland-Urban Interface）：野生地と都市の境界帯。山火事被害が住宅に直接及ぶ高リスク領域",
         size=13, color=INK)

    # callout
    cy = ty + th + Inches(0.18)
    callout(sl, BODY_X, cy, left_w, Inches(0.65),
            "気候変動により乾燥・強風期間が延長 → 従来の消火戦略では対応が追いつかない",
            icon='⚠')

    # 右：図プレースホルダー
    fig_h = Inches(4.6)
    fig_placeholder(sl, right_x, BODY_Y, right_w, fig_h,
                    fig_num="Fig. 3c",
                    caption="直炎接触下での燃焼過程")


# ===== SLIDE 5: 4段階の保護プロセス =====
def slide_05_process(prs):
    sl = new_slide(prs)
    draw_header(sl, "背景 2/4", "本研究のアプローチ：4段階の保護プロセス")
    draw_footer(sl, "5 / 36")

    # 上部：大きな図プレースホルダー
    fig_h = Inches(3.0)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 1a",
                    caption="難燃ゲル散布 → 火炎接触 → 熱活性化エアロゲル形成 → 構造物の保護")

    # 下部：4ステップカード
    sy = BODY_Y + fig_h + Inches(0.45)
    sh = Inches(2.2)
    cw = (BODY_W - Inches(0.42)) / 4
    steps = [
        ("STEP 1", "ゲル散布", "セルロース系ゲルを建物・植生にスプレー塗布"),
        ("STEP 2", "火炎接触", "水分が蒸発し、ゲルが発泡しながら粒子が凝集"),
        ("STEP 3", "エアロゲル形成", "シリカ粒子が焼結し多孔質エアロゲル層が生成"),
        ("STEP 4", "継続保護", "超低熱伝導率の断熱層が基材を継続的に守る"),
    ]
    for i, (tag, title, body) in enumerate(steps):
        cx = BODY_X + i * (cw + Inches(0.14))
        card(sl, cx, sy, cw, sh, title=title, body=body, tag=tag,
             ct_size=16, cb_size=12)


# ===== SLIDE 6: 既存技術と限界 =====
def slide_06_existing(prs):
    sl = new_slide(prs)
    draw_header(sl, "背景 3/4", "既存の山火事対策と、その本質的な限界")
    draw_footer(sl, "6 / 36")

    # 左：テーブル、右：図
    left_w = BODY_W * 0.66 - Inches(0.10)
    right_w = BODY_W * 0.34 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    headers = ["技術", "主成分", "特徴", "限界"]
    rows = [
        ["水のみ", "H₂O", "蒸発潜熱による冷却",
         {'text': "即座に流失・蒸発", 'color': D_RED}],
        ["Phos-Chek", "リン酸アンモニウム", "樹木・地表に直接散布",
         {'text': "土壌・水系への残留", 'color': D_RED}],
        ["AquaGel-K", "架橋ポリアクリレート", "水を高濃度に保持",
         {'text': "水蒸発後は機能消失", 'color': D_RED}],
        [{'text': "本研究 WEG", 'bold': True, 'color': ACCENT},
         "セルロース系＋シリカ", "加熱でエアロゲル層形成",
         {'text': "水蒸発後も継続保護", 'color': D_GRN, 'bold': True}],
    ]
    table_h = Inches(3.2)
    simple_table(sl, BODY_X, BODY_Y, left_w, table_h, headers, rows,
                 col_widths=[1.5, 2, 2, 1.5])

    # callout
    cy = BODY_Y + table_h + Inches(0.22)
    callout(sl, BODY_X, cy, left_w, Inches(0.70),
            "課題：既存WEGは「水のキャリア」止まりで、水分喪失と同時に保護機能が失われる",
            dark=True, icon='→')

    # 右：図プレースホルダー
    fig_placeholder(sl, right_x, BODY_Y, right_w, Inches(4.2),
                    fig_num="Fig. 3d",
                    caption="300秒後の比較：Water vs 本研究WEG")


# ===== SLIDE 7: 研究の核心 =====
def slide_07_core(prs):
    sl = new_slide(prs)
    draw_header(sl, "背景 4/4", "本研究の核心：熱活性化エアロゲル形成")
    draw_footer(sl, "7 / 36")

    left_w = BODY_W * 0.60 - Inches(0.10)
    right_w = BODY_W * 0.40 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    # 左：2カード
    ch = Inches(3.4)
    cw_l = (left_w - Inches(0.14)) / 2
    card(sl, BODY_X, BODY_Y, cw_l, ch,
         title="設計コンセプト", tag="DESIGN CONCEPT",
         bullets=[
             "セルロースポリマー(HEC, MC, MHEC)が水溶性ゲル骨格を形成",
             "コロイダルシリカ粒子(CSP)を均一分散",
             "加熱時に水が発泡し多孔構造を形成",
             "シリカ粒子が焼結しエアロゲル層に転換",
         ], ct_size=15, li_size=11)
    card(sl, BODY_X + cw_l + Inches(0.14), BODY_Y, cw_l, ch,
         title="3つの革新ポイント", tag="INNOVATION",
         dark=True,
         bullets=[
             "水分依存からの脱却：熱で活性化する固体保護層",
             "持続可能：天然由来・食品添加物グレード",
             "既存散布インフラと互換：噴霧可能な流体",
         ], ct_size=15, li_size=11)

    # 下部callout
    cy_co = BODY_Y + ch + Inches(0.25)
    callout(sl, BODY_X, cy_co, left_w, Inches(0.80),
            '"Heat-Activated Formation of Silica Aerogels"\n— 火炎接触をトリガーに、ゲル自身がエアロゲル断熱材へ変態する',
            icon='💡')

    # 右：図プレースホルダー
    fig_placeholder(sl, right_x, BODY_Y, right_w, Inches(4.8),
                    fig_num="Fig. 5a",
                    caption="シリカエアロゲル形成の3段階（プレビュー）")


# ===== SLIDE 9: セルロース系ポリマー =====
def slide_09_polymers(prs):
    sl = new_slide(prs)
    draw_header(sl, "材料 1/6", "使用したセルロース系ポリマー")
    draw_footer(sl, "9 / 36")

    # 上部：3つのpolymerカード
    pw = (BODY_W - Inches(0.28)) / 3
    ph = Inches(2.1)
    polymers = [
        ("Polymer A", "HEC", "Hydroxyethyl cellulose\nヒドロキシエチルセルロース", "→ ベース粘性付与"),
        ("Polymer B", "MC", "Methyl cellulose\nメチルセルロース", "→ 加熱時にゲル化（熱可逆性）"),
        ("Polymer C", "MHEC", "Methyl 2-hydroxyethyl cellulose\nメチル2-ヒドロキシエチルセルロース", "→ AとBの特性を統合"),
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

    # 下部：図プレースホルダー
    fy = BODY_Y + ph + Inches(0.30)
    fh = BODY_H - ph - Inches(0.30)
    fig_placeholder(sl, BODY_X, fy, BODY_W, fh - Inches(0.35),
                    fig_num="Fig. 1b/c",
                    caption="各ポリマーの化学構造とコロイダルシリカ(CSP)との混合スキーム")


# ===== SLIDE 10: MC熱ゲル化 =====
def slide_10_mc(prs):
    sl = new_slide(prs)
    draw_header(sl, "材料 2/6", "メチルセルロース(MC)の熱ゲル化：火炎保護の鍵となる特性")
    draw_footer(sl, "10 / 36")

    # 左：説明カード、右：温度プロセス
    left_w = BODY_W * 0.48 - Inches(0.10)
    right_w = BODY_W * 0.52 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    # 左カード
    ch_l = Inches(3.5)
    card(sl, BODY_X, BODY_Y, left_w, ch_l,
         title="逆熱応答（LCST挙動）", tag="THERMAL RESPONSE",
         bullets=[
             "MCは冷水（<20°C）に溶解し低粘度溶液を形成",
             "加熱でLCSTを超えてゲル化",
             "MCのゲル化開始温度：約 50–60°C（濃度依存）",
             "常温では液体 → 火炎接触で即座にゲル",
         ], ct_size=15, li_size=12)
    callout(sl, BODY_X, BODY_Y + ch_l + Inches(0.20), left_w, Inches(0.85),
            "火炎接触初期にMCがゲル化 → CSPを固定したまま発泡構造を維持",
            icon='🔥')

    # 右：3つのmini stats
    mw = (right_w - Inches(0.20)) / 3
    mh = Inches(1.4)
    mini(sl, right_x, BODY_Y, mw, mh, "LCST（ゲル化開始）", "~55°C", "1 wt% MC水溶液", value_color=ACCENT)
    mini(sl, right_x + mw + Inches(0.10), BODY_Y, mw, mh, "熱分解温度", "~300°C", "TGA測定", value_color=D_RED)
    mini(sl, right_x + (mw + Inches(0.10)) * 2, BODY_Y, mw, mh, "CSP焼結開始", "~200°C", "粒子間ネック形成", value_color=D_TEAL)

    # 温度プロセスフロー (4段)
    fy = BODY_Y + mh + Inches(0.30)
    fh = Inches(2.7)
    # 暗いカード背景
    rrect(sl, right_x, fy, right_w, fh, HEADER, None, radius=0.04)
    text(sl, right_x + Inches(0.20), fy + Inches(0.18), right_w - Inches(0.4), Inches(0.30),
         "温度別プロセス", size=11, bold=True, color=RGBColor(0xa0, 0xb8, 0xd0))

    sub_y = fy + Inches(0.65)
    sub_h = fh - Inches(0.75)
    pw = (right_w - Inches(0.40) - Inches(0.30) * 3) / 4
    steps_t = [
        ("~55°C", "MCゲル化", "構造を固定"),
        ("~100°C", "水分蒸発", "発泡・膨張"),
        ("~200°C", "CSP焼結", "粒界形成"),
        (">300°C", "有機分解", "純シリカ層"),
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
    draw_header(sl, "材料 3/6", "コロイダルシリカ粒子(CSP)と界面活性剤(SDS)")
    draw_footer(sl, "11 / 36")

    cw = (BODY_W - Inches(0.18)) / 2
    ch = Inches(4.2)

    # 左：CSP
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
        "原液：水中で12 wt%濃度のコロイド分散液",
        "粒子サイズ：数十nmの単分散シリカ",
        "配合濃度：ゲル中 5 wt%",
        "加熱時に焼結 → silica aerogel に変態",
    ]
    by = BODY_Y + Inches(1.30)
    for b in bullets:
        text(sl, cx + Inches(0.30), by, Inches(0.20), Inches(0.30),
             "•", size=14, color=ACCENT)
        text(sl, cx + Inches(0.55), by, cw - Inches(0.75), Inches(0.30),
             b, size=13, color=INK2)
        by += Inches(0.42)

    # 右：SDS
    cx2 = BODY_X + cw + Inches(0.18)
    rrect(sl, cx2, BODY_Y, cw, ch, CARD_C, LINE, 0.5, radius=0.04)
    rect(sl, cx2, BODY_Y, cw, Inches(0.04), ACCENT)
    text(sl, cx2 + Inches(0.20), BODY_Y + Inches(0.20), Inches(0.6), Inches(0.5),
         "◆", size=28, color=ACCENT)
    text(sl, cx2 + Inches(0.95), BODY_Y + Inches(0.22), cw - Inches(1.1), Inches(0.3),
         "SDS（添加剤）", size=11, bold=True, color=MUTED)
    text(sl, cx2 + Inches(0.95), BODY_Y + Inches(0.50), cw - Inches(1.1), Inches(0.5),
         "Sodium Dodecyl Sulfate", size=18, bold=True, color=INK)
    bullets2 = [
        "陰イオン性界面活性剤",
        "役割：発泡を促進し多孔構造を強化",
        "添加濃度：0.1 wt% ・ 0.5 wt% の2水準",
        "SEMで気泡サイズ・分布を制御確認",
    ]
    by = BODY_Y + Inches(1.30)
    for b in bullets2:
        text(sl, cx2 + Inches(0.30), by, Inches(0.20), Inches(0.30),
             "•", size=14, color=ACCENT)
        text(sl, cx2 + Inches(0.55), by, cw - Inches(0.75), Inches(0.30),
             b, size=13, color=INK2)
        by += Inches(0.42)

    # 下部callout
    cy_co = BODY_Y + ch + Inches(0.30)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.80),
            "CSPの焼結温度域とセルロースの熱分解温度が重なるため、熱応答が同期して断熱層を形成",
            icon='💡')


# ===== SLIDE 12: 配合系 5種 =====
def slide_12_formulations(prs):
    sl = new_slide(prs)
    draw_header(sl, "材料 4/6", "評価した5種類の配合系")
    draw_footer(sl, "12 / 36")

    # 5つのカード横並び
    fw = (BODY_W - Inches(0.4 * 4)) / 5
    fh = Inches(1.9)
    forms = [
        ("対照", "AquaGel-K", "0.5 wt%\n市販WEG基準", False),
        ("配合 1", "HEC+MC / CSP", "HEC+MC 1 wt%\nCSP 5 wt%", True),
        ("配合 2", "MHEC / CSP", "MHEC 1 wt%\nCSP 5 wt%", True),
        ("配合 3", "HEC+MC / CSP / SDS", "+ SDS 0.1 wt%", True),
        ("配合 4", "HEC+MC / CSP / SDS", "+ SDS 0.5 wt%", True),
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

    # 下部2カード
    cy = fy + fh + Inches(0.40)
    ch = Inches(2.5)
    cw2 = (BODY_W - Inches(0.20)) / 2
    card(sl, BODY_X, cy, cw2, ch,
         title="配合設計の意図",
         bullets=[
             "ポリマー種を変えて(HEC+MC vs MHEC)機能差を比較",
             "SDS濃度を2水準で発泡構造への影響を検証",
             "市販品AquaGel-Kとの直接比較",
         ], ct_size=15, li_size=12)
    card(sl, BODY_X + cw2 + Inches(0.20), cy, cw2, ch,
         title="記法の読み方",
         body="HEC+MC/CSP/SDS 1-5-0.1\n→ HEC+MC 1 wt% ／ CSP 5 wt% ／ SDS 0.1 wt%",
         ct_size=15, cb_size=14)


# ===== SLIDE 13: 評価手法 =====
def slide_13_methods(prs):
    sl = new_slide(prs)
    draw_header(sl, "材料 5/6", "評価手法の全体像")
    draw_footer(sl, "13 / 36")

    methods = [
        ("レオロジー測定", "RHEOLOGY",
         ["振動周波数掃引（G', G''）", "定常流動掃引（粘度 vs 剪断速度）", "Herschel-Bulkleyモデル適合"]),
        ("燃焼試験（Time-to-char）", "BURN TEST",
         ["ブタンバーナーで木材を加熱", "炭化開始までの時間を計測", "120 s / 300 s時点を写真比較"]),
        ("発泡指数測定", "FOAMING",
         ["燃焼後の発泡層厚さを計測", "初期厚さに対する比 = Foaming Index"]),
        ("SEM形態観察", "SEM",
         ["SDS濃度別の発泡構造", "燃焼時間別(0/1/2/4分)の焼結進行"]),
        ("分光分析", "SPECTROSCOPY",
         ["FT-IR（化学結合）", "XPS（表面組成）"]),
        ("熱分析", "THERMAL",
         ["TGA（熱重量分析）", "DSC（示差走査熱量）"]),
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
    draw_header(sl, "材料 6/6", "付着性・表面濡れ性：実装に不可欠な特性")
    draw_footer(sl, "14 / 36")

    # 3カード
    ch = Inches(3.6)
    cw = (BODY_W - Inches(0.28)) / 3
    cards_data = [
        ("接触角測定", "WETTING", [
            "木材・コンクリート・金属板で接触角評価",
            "WEGは低接触角 → 高い濡れ性",
            "HEC+MCはセルロース-木材間で特に親和性が高い",
        ]),
        ("垂直面付着試験", "ADHESION", [
            "垂直に立てた基板にゲルを塗布",
            "G' ≫ G'' → 重力に抵抗して流れ落ちない",
            "AquaGel-Kと同等以上の垂直面保持性",
        ]),
        ("スプレー散布適性", "SPRAY", [
            "高せん断（ノズル内）で粘度が急低下",
            "基材到達後すぐに高粘度を回復 → 付着維持",
            "既存消防ホース・ノズルと完全互換",
        ]),
    ]
    for i, (title, tag, bullets) in enumerate(cards_data):
        cx = BODY_X + i * (cw + Inches(0.14))
        card(sl, cx, BODY_Y, cw, ch, title=title, tag=tag,
             bullets=bullets, ct_size=15, li_size=12)

    # 下部 callout
    cy_co = BODY_Y + ch + Inches(0.30)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.80),
            "流動学的設計（低 n、高 G'）により、散布適性と付着保持を同時に達成",
            icon='✓')


# ===== SLIDE 16: G'/G'' =====
def slide_16_rheology1(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 1/11", "レオロジー①：振動弾性率 G' / G''")
    draw_footer(sl, "16 / 36")

    # 左：図、右：観察
    left_w = BODY_W * 0.58 - Inches(0.10)
    right_w = BODY_W * 0.42 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    fig_h = Inches(4.8)
    fig_placeholder(sl, BODY_X, BODY_Y, left_w, fig_h,
                    fig_num="Fig. 2a–b",
                    caption="角周波数 vs Storage modulus G' ・ Loss modulus G''")

    # 右：観察カード
    ch = Inches(2.8)
    card(sl, right_x, BODY_Y, right_w, ch,
         title="主な観察", tag="KEY OBSERVATION",
         bullets=[
             "全配合系で G' > G'' → 固体的（ゲル）挙動",
             "周波数依存性が小さい → 安定したネットワーク",
             "AquaGel-Kよりも本研究WEGの方が高弾性",
         ], ct_size=15, li_size=12)
    callout(sl, right_x, BODY_Y + ch + Inches(0.25), right_w, Inches(0.80),
            "ゲル骨格が明確に形成されており、塗布後に流れ落ちない",
            icon='"')


# ===== SLIDE 17: shear thinning =====
def slide_17_rheology2(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 2/11", "レオロジー②：剪断希薄化と Power-law 指数")
    draw_footer(sl, "17 / 36")

    left_w = BODY_W * 0.58 - Inches(0.10)
    right_w = BODY_W * 0.42 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    fig_h = Inches(4.8)
    fig_placeholder(sl, BODY_X, BODY_Y, left_w, fig_h,
                    fig_num="Fig. 2c–d",
                    caption="粘度 vs 剪断速度（shear-thinning挙動）")

    # 右：mini stats + 解説
    mh = Inches(1.20)
    mini(sl, right_x, BODY_Y, right_w, mh,
         "Power-law index n", "0.108", "HEC+MC/CSP 1-5", value_color=ACCENT)
    mini(sl, right_x, BODY_Y + mh + Inches(0.15), right_w, mh,
         "Power-law index n", "0.188", "MHEC/CSP 1-5", value_color=ACCENT)

    cy_d = BODY_Y + (mh + Inches(0.15)) * 2 + Inches(0.10)
    rrect(sl, right_x, cy_d, right_w, Inches(1.0), WHITE, LINE, 0.5, radius=0.04)
    text(sl, right_x + Inches(0.15), cy_d + Inches(0.15), right_w - Inches(0.3), Inches(0.7),
         "n < 1 = 剪断希薄化（shear-thinning）\n低いほど高せん断でよく流れる → スプレー性◎",
         size=12, color=INK3)

    cy_co = BODY_Y + fig_h + Inches(0.20)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.65),
            "静止時は高粘度で付着、噴霧時は低粘度で散布 — 既存の消火機材でそのまま使える流動特性",
            icon='💨')


# ===== SLIDE 18: Herschel-Bulkley =====
def slide_18_hb(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 3/11", "レオロジー③：Herschel-Bulkley モデルパラメータ")
    draw_footer(sl, "18 / 36")

    # 上部 callout
    callout(sl, BODY_X, BODY_Y, BODY_W, Inches(0.65),
            "σ = τ₀ + K · γ̇ⁿ  （τ₀：降伏応力, K：稠度係数, n：流動指数）",
            icon='ƒ')

    ty = BODY_Y + Inches(0.85)
    th = Inches(2.85)
    headers = ["配合系", "τ₀ (Pa)", "K (Pa·sⁿ)", "n", "R²"]
    rows = [
        ["AquaGel-K（市販対照）", "0.31", "0.85", "0.52",
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

    # 下部 3カード
    cy = ty + th + Inches(0.30)
    ch = Inches(1.85)
    cw = (BODY_W - Inches(0.30)) / 3
    cards_data = [
        ("降伏応力 τ₀", "流れ始めるのに必要な最小応力。WEGはAquaGel-Kの約2倍 → 垂直面での流れ落ち抵抗が高い"),
        ("流動指数 n （<< 1）", "HEC+MC/CSP の n=0.108 は強い剪断希薄化を示す。スプレー時の低粘度と静止時の高粘度を両立"),
        ("HEC+MC vs MHEC", "HEC+MC系はより低い n と高い K → より強い剪断希薄化。MHEC系は単一ポリマーで類似機能を実現"),
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
    draw_header(sl, "結果 4/11", "レオロジー④：動的降伏応力・粘弾性（tan δ）")
    draw_footer(sl, "19 / 36")

    cw = (BODY_W - Inches(0.30)) / 2
    ch = Inches(5.6)
    cx2 = BODY_X + cw + Inches(0.30)

    # 左：動的降伏応力
    text(sl, BODY_X, BODY_Y, cw, Inches(0.30),
         "動的降伏応力（Amplitude Sweep より）", size=12, bold=True, color=MUTED)
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
            "G' = G'' クロスオーバー点での応力 = 動的降伏応力。WEGはAquaGel-Kの約3倍",
            icon='ⓘ')

    # 右：tan δ
    text(sl, cx2, BODY_Y, cw, Inches(0.30),
         "損失正接 tan δ = G'' / G'（弾性支配 = <1）", size=12, bold=True, color=MUTED)
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
            "tan δ << 1 → 全配合系が明確なゲル挙動。弾性成分が圧倒的に支配的",
            dark=True, icon='✓')


# ===== SLIDE 20: HEC+MC vs MHEC =====
def slide_20_compare(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 5/11", "HEC+MC vs MHEC：二系統の比較")
    draw_footer(sl, "20 / 36")

    cw = (BODY_W - Inches(0.20)) / 2
    ch = Inches(4.8)

    # 左：HEC+MC
    rrect(sl, BODY_X, BODY_Y, cw, ch, CARD_C, LINE, 0.5, radius=0.04)
    rect(sl, BODY_X, BODY_Y, cw, Inches(0.04), ACCENT)
    text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(0.18), cw - Inches(0.4), Inches(0.30),
         "HEC + MC 混合系", size=12, bold=True, color=MUTED)
    text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(0.55), cw - Inches(0.4), Inches(0.40),
         "HEC（非熱ゲル化） + MC（熱ゲル化・LCST）",
         size=13, italic=True, color=INK3)

    by = BODY_Y + Inches(1.10)
    bullets = [
        "HECで常温粘度・付着性を確保",
        "MCで加熱時の構造維持を担当",
        "より低い n=0.108 → 強い剪断希薄化",
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
         "→ SDSとの相性も良く、発泡構造の最適化に優れる",
         size=12, bold=True, color=ACCENT, anchor=MSO_ANCHOR.MIDDLE)

    # 右：MHEC
    cx2 = BODY_X + cw + Inches(0.20)
    rrect(sl, cx2, BODY_Y, cw, ch, CARD_C, LINE, 0.5, radius=0.04)
    rect(sl, cx2, BODY_Y, cw, Inches(0.04), ACCENT)
    text(sl, cx2 + Inches(0.20), BODY_Y + Inches(0.18), cw - Inches(0.4), Inches(0.30),
         "MHEC 単独系", size=12, bold=True, color=MUTED)
    text(sl, cx2 + Inches(0.20), BODY_Y + Inches(0.55), cw - Inches(0.4), Inches(0.40),
         "MHEC（HEC+MCの機能を一分子に統合）",
         size=13, italic=True, color=INK3)
    by = BODY_Y + Inches(1.10)
    bullets2 = [
        "ヒドロキシエチル基とメチル基を同一鎖上に有する",
        "単一ポリマーで同等の粘弾性・熱応答を実現",
        "n=0.188（HEC+MC より若干高い）",
        "Time to char ≈ 10 min",
        "配合・品質管理が単純化される利点",
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
         "→ スケールアップ・製造コスト低減に有利",
         size=12, bold=True, color=ACCENT, anchor=MSO_ANCHOR.MIDDLE)

    cy_co = BODY_Y + ch + Inches(0.25)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.80),
            "両系統とも市販AquaGel-Kを大幅に上回る性能。HEC+MC系は発泡最適化、MHEC系は製造簡便性に優れる",
            icon='⚖')


# ===== SLIDE 21: 燃焼試験 setup =====
def slide_21_setup(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 6/11", "燃焼試験のセットアップ")
    draw_footer(sl, "21 / 36")

    left_w = BODY_W * 0.55 - Inches(0.10)
    right_w = BODY_W * 0.45 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    fig_placeholder(sl, BODY_X, BODY_Y, left_w, Inches(4.6),
                    fig_num="Fig. 3a",
                    caption="燃焼試験の実験セットアップ")

    # 右：試験条件カード + callout
    ch = Inches(3.5)
    card(sl, right_x, BODY_Y, right_w, ch,
         title="◆ 試験条件", tag="EXPERIMENTAL",
         bullets=[
             "基板：松材（pine wood）薄板",
             "ゲル散布厚：均一に塗布",
             "火炎源：ブタン直炎",
             "計測：基材が炭化するまでの時間",
             "120 s時点・300 s時点で外観撮影",
         ], ct_size=15, li_size=12)
    callout(sl, right_x, BODY_Y + ch + Inches(0.20), right_w, Inches(0.85),
            "Time to char = 木材表面が炭化するまでに要した時間（長いほど高保護）",
            icon='⏱')


# ===== SLIDE 22: Time to char =====
def slide_22_ttc(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 7/11", "Time to char：5配合系の定量比較")
    draw_footer(sl, "22 / 36")

    left_w = BODY_W * 0.62 - Inches(0.10)
    right_w = BODY_W * 0.38 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    # 左：バーチャート
    text(sl, BODY_X, BODY_Y, left_w, Inches(0.3),
         "Time to char（炭化開始までの時間）", size=12, bold=True, color=MUTED)
    by = BODY_Y + Inches(0.45)
    bar_data = [
        ("Water", 0.18, "~2 min", GRAY_L),
        ("AquaGel-K（市販品）", 0.60, "~7 min", D_AMBR),
        ("HEC+MC/CSP 1-5", 0.80, "~9 min", D_TEAL),
        ("HEC+MC/CSP/SDS 1-5-0.1", 0.90, "~10 min", D_BLUE),
        ("MHEC/CSP 1-5", 0.92, "~10 min", D_BLUE),
    ]
    for label, ratio, val, color in bar_data:
        bar_row(sl, BODY_X, by, left_w, label, ratio, val, fill_color=color,
                lbl_w=Inches(2.6), val_w=Inches(1.1))
        by += Inches(0.55)

    callout(sl, BODY_X, by + Inches(0.15), left_w, Inches(0.85),
            "本研究WEG群はAquaGel-K比で約40%延長、水のみと比べて約5倍の保護時間",
            icon='📊')

    # 右：図プレースホルダー
    fig_placeholder(sl, right_x, BODY_Y, right_w, Inches(5.2),
                    fig_num="Fig. 3b",
                    caption="Time-to-char バーチャート（n≧3、エラーバーは標準偏差）")


# ===== SLIDE 23: 燃焼時系列 =====
def slide_23_timelapse(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 8/11", "燃焼過程の時系列観察")
    draw_footer(sl, "23 / 36")

    # 上部：図プレースホルダー
    fig_h = Inches(2.4)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 3c",
                    caption="各配合系の火炎接触時の連続写真")

    # 下部：3カード比較
    cy = BODY_Y + fig_h + Inches(0.50)
    ch = Inches(2.7)
    cw = (BODY_W - Inches(0.28)) / 3
    cards_data = [
        ("Water", D_RED, [
            "直炎で即座に蒸発・流失",
            "~2分で木材表面が炭化",
            "保護層ゼロ",
        ]),
        ("AquaGel-K", D_AMBR, [
            "水を保持するが水蒸発で終了",
            "~7分で炭化",
            "固体保護層を形成しない",
        ]),
        ("本研究 WEG", ACCENT, [
            "水蒸発と同時にゲルが発泡・膨張",
            "多孔質エアロゲル層が継続保護",
            "~10分間、炭化を阻止",
        ]),
    ]
    for i, (title, color, bullets) in enumerate(cards_data):
        cx = BODY_X + i * (cw + Inches(0.14))
        bd = color if title == "本研究 WEG" else LINE
        rrect(sl, cx, cy, cw, ch, WHITE, bd, 0.8 if title == "本研究 WEG" else 0.5, radius=0.04)
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
    draw_header(sl, "結果 8/11", "120秒 / 300秒時点の表面状態比較")
    draw_footer(sl, "24 / 36")

    fig_h = Inches(2.3)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 3d",
                    caption="120秒時点（上段）と300秒時点（下段）の表面状態比較")

    cy = BODY_Y + fig_h + Inches(0.50)
    ch = Inches(2.8)
    cw = (BODY_W - Inches(0.28)) / 3
    cards_data = [
        ("Water", D_RED, "Time to char：~2 min", [
            "120 s：既に広範囲が炭化開始",
            "300 s：表面全体が黒化・損傷甚大",
        ]),
        ("AquaGel-K（市販品）", D_AMBR, "Time to char：~7 min", [
            "120 s：ゲルが乾燥し一部炭化",
            "300 s：保護層なし、部分炭化",
        ]),
        ("本研究 WEG", ACCENT, "Time to char：~10 min", [
            "120 s：エアロゲル層が形成・膨張中",
            "300 s：表面ほぼ無傷、層が残存",
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
         "本研究WEGの Time to char", size=26, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER)
    text(sl, Inches(1.5), Inches(5.15), Inches(10.3), Inches(0.7),
         "水のみ（~2分）の約5倍、市販AquaGel-K（~7分）の約1.4倍の保護時間を達成",
         size=14, color=RGBColor(0x9c, 0xb4, 0xcc), align=PP_ALIGN.CENTER)

    # bars
    bars = [
        ("Water", 0.20, "~2 min", GRAY_L),
        ("AquaGel-K", 0.60, "~7 min", D_AMBR),
        ("本研究 WEG", 0.95, "~10 min", GOLD),
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
    draw_header(sl, "結果 9/11", "発泡指数（Foaming Index）の比較")
    draw_footer(sl, "26 / 36")

    # 上部：2図プレースホルダー
    fig_h = Inches(2.3)
    fw = (BODY_W - Inches(0.20)) / 2
    fig_placeholder(sl, BODY_X, BODY_Y, fw, fig_h, fig_num="Fig. 4a",
                    caption="燃焼後の発泡層外観")
    fig_placeholder(sl, BODY_X + fw + Inches(0.20), BODY_Y, fw, fig_h,
                    fig_num="Fig. 4b", caption="各配合系の Foaming Index")

    # 下部：左バーチャート、右解説
    cy = BODY_Y + fig_h + Inches(0.45)
    cw = (BODY_W - Inches(0.20)) / 2

    text(sl, BODY_X, cy, cw, Inches(0.30),
         "Foaming Index（発泡後厚 / 初期厚）", size=12, bold=True, color=MUTED)
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

    # 右側：解説
    cx2 = BODY_X + cw + Inches(0.20)
    text(sl, cx2, cy, cw, Inches(0.30),
         "発泡がもたらす断熱効果", size=12, bold=True, color=MUTED)
    by = cy + Inches(0.40)
    bullets = [
        "初期厚の2倍以上に膨張 → 熱伝導経路が延長",
        "気孔内に空気が閉じ込められ断熱性が向上",
        "AquaGel-Kは発泡せず → 火炎で平坦に崩壊",
        "SDS 0.1 wt%が最適：均一な微細気泡を形成",
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
    draw_header(sl, "結果 9/11", "SDS濃度がエアロゲル微細構造に与える影響")
    draw_footer(sl, "27 / 36")

    fig_h = Inches(3.0)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 4c",
                    caption="SDS濃度別のFE-SEM観察像（0%, 0.1%, 0.5%）")

    cy = BODY_Y + fig_h + Inches(0.45)
    ch = Inches(2.3)
    cw = (BODY_W - Inches(0.28)) / 3
    cards_data = [
        ("0% SDS", MUTED, "緻密なシリカネットワーク。気孔少。"),
        ("0.1% SDS（最適）", ACCENT, "均一な微細気泡。最高の発泡指数。"),
        ("0.5% SDS（過剰）", MUTED, "気泡サイズ不均一・粗大化。"),
    ]
    for i, (title, color, desc) in enumerate(cards_data):
        cx = BODY_X + i * (cw + Inches(0.14))
        bd = ACCENT if title.endswith("（最適）") else LINE
        rrect(sl, cx, cy, cw, ch, WHITE, bd, 0.8 if title.endswith("（最適）") else 0.5, radius=0.04)
        text(sl, cx + Inches(0.18), cy + Inches(0.22), cw - Inches(0.36), Inches(0.35),
             title, size=13, bold=True, color=color)
        text(sl, cx + Inches(0.18), cy + Inches(0.75), cw - Inches(0.36), Inches(1.40),
             desc, size=12, color=INK3)


# ===== SLIDE 28: FT-IR / XPS =====
def slide_28_ftir_xps(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 10/11", "FT-IR・XPS：化学組成の変化（燃焼前後）")
    draw_footer(sl, "28 / 36")

    cw = (BODY_W - Inches(0.20)) / 2

    # 左：FT-IR
    text(sl, BODY_X, BODY_Y, cw, Inches(0.30),
         "FT-IR 主要ピーク帰属", size=12, bold=True, color=MUTED)
    ty = BODY_Y + Inches(0.42)
    th = Inches(3.0)
    headers = ["波数 (cm⁻¹)", "帰属", "燃焼後"]
    rows = [
        ["3200–3500", "O–H 伸縮（セルロース・水）",
         {'text': "消失（脱水）", 'color': D_RED}],
        ["2850–2950", "C–H 伸縮（メチル基）",
         {'text': "消失（有機分解）", 'color': D_RED}],
        ["1050–1100", "Si–O–Si 伸縮（シリカ）",
         {'text': "強度増大", 'color': D_GRN, 'bold': True}],
        ["800", "Si–O 変角振動",
         {'text': "明確化", 'color': D_GRN, 'bold': True}],
        ["450", "Si–O 曲げ振動",
         {'text': "シャープ化", 'color': D_GRN, 'bold': True}],
    ]
    simple_table(sl, BODY_X, ty, cw, th, headers, rows,
                 col_widths=[1.2, 2.5, 1.5])
    callout(sl, BODY_X, ty + th + Inches(0.20), cw, Inches(0.85),
            "燃焼後のスペクトルは純シリカ（SiO₂）と一致 → エアロゲル化を化学的に確認",
            icon='🔥')

    # 右：XPS
    cx2 = BODY_X + cw + Inches(0.20)
    text(sl, cx2, BODY_Y, cw, Inches(0.30),
         "XPS 表面元素組成（at%）", size=12, bold=True, color=MUTED)
    th2 = Inches(2.1)
    headers2 = ["元素", "燃焼前", "燃焼後"]
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
            "Si濃度が約8倍に増加、C濃度が激減 → 有機マトリクスが除去されシリカが露出",
            dark=True, icon='📊')

    info_y = BODY_Y + Inches(0.42) + th2 + Inches(1.20)
    rrect(sl, cx2, info_y, cw, Inches(0.85), WHITE, LINE, 0.5, radius=0.04)
    text(sl, cx2 + Inches(0.18), info_y + Inches(0.14), cw - Inches(0.36), Inches(0.30),
         "ピーク位置の確認", size=11, bold=True, color=MUTED)
    text(sl, cx2 + Inches(0.18), info_y + Inches(0.42), cw - Inches(0.36), Inches(0.45),
         "Si 2p ピーク：~103.5 eV（SiO₂と一致）／ O 1s ピーク：~533 eV",
         size=12, color=INK3)


# ===== SLIDE 29: TGA/DSC =====
def slide_29_tga(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 11/11", "TGA/DSC：熱分解プロファイルと各成分の役割")
    draw_footer(sl, "29 / 36")

    cw = (BODY_W - Inches(0.20)) / 2

    # 左：TGA
    text(sl, BODY_X, BODY_Y, cw, Inches(0.30),
         "TGA：重量損失プロファイル（N₂雰囲気）", size=12, bold=True, color=MUTED)
    by = BODY_Y + Inches(0.45)
    tga_data = [
        ("50–150°C", "自由水・吸着水の蒸発", 0.30, D_BLUE),
        ("200–350°C", "HEC・MC有機鎖の熱分解（主要重量損失）", 0.85, D_RED),
        ("350–600°C", "残留炭素の酸化・消失（空気中）", 0.50, D_AMBR),
        (">600°C", "シリカ残渣（~5–8 wt%、変化なし）", 0.08, D_TEAL),
    ]
    for label, desc, ratio, color in tga_data:
        text(sl, BODY_X, by, cw, Inches(0.25),
             label + "：" + desc, size=11, color=MUTED)
        rrect(sl, BODY_X, by + Inches(0.30), cw, Inches(0.14),
              CARD_C, LINE, 0.3, radius=0.5)
        if ratio > 0.02:
            rrect(sl, BODY_X, by + Inches(0.30), cw * ratio, Inches(0.14),
                  color, None, radius=0.5)
        by += Inches(0.62)

    # 右：DSC
    cx2 = BODY_X + cw + Inches(0.20)
    text(sl, cx2, BODY_Y, cw, Inches(0.30),
         "DSC：熱イベントの帰属", size=12, bold=True, color=MUTED)
    headers = ["温度域", "イベント", "ΔH"]
    rows = [
        ["~100°C", "吸熱：水の蒸発", "吸熱"],
        ["~55–80°C", "発熱：MCゲル化転移", {'text': "微発熱", 'color': D_GRN}],
        ["~250–300°C", "発熱：有機鎖の酸化分解", {'text': "発熱", 'color': D_RED}],
    ]
    simple_table(sl, cx2, BODY_Y + Inches(0.45), cw, Inches(1.6), headers, rows,
                 col_widths=[1.5, 3, 1.2])

    # 下部mini stats
    my = BODY_Y + Inches(0.45) + Inches(1.6) + Inches(0.30)
    mw = (cw - Inches(0.15)) / 2
    mini(sl, cx2, my, mw, Inches(1.2), "最終シリカ残渣", "~5–8 wt%", "CSP 5 wt% に相当", value_color=D_TEAL)
    mini(sl, cx2 + mw + Inches(0.15), my, mw, Inches(1.2),
         "セルロース分解ピーク", "~280°C", "HEC/MC共通", value_color=D_RED)

    callout(sl, cx2, my + Inches(1.4), cw, Inches(0.85),
            "CSP焼結（~200°C〜）とセルロース分解（~280°C〜）が温度的に重複 → エアロゲル化が同期",
            icon='🌡')


# ===== SLIDE 30: エアロゲル形成メカニズム =====
def slide_30_mechanism(prs):
    sl = new_slide(prs)
    draw_header(sl, "補足 1/2", "シリカエアロゲル形成のメカニズム")
    draw_footer(sl, "30 / 36")

    fig_h = Inches(2.4)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 5a",
                    caption="火炎活性化によるWEG → シリカエアロゲル変態の3段階模式図")

    cy = BODY_Y + fig_h + Inches(0.45)
    ch = Inches(2.8)
    cw = (BODY_W - Inches(0.28)) / 3
    phases = [
        ("PHASE 1", "Unsintered", [
            "シリカ粒子が独立分散",
            "間隙に水が満たされている",
            "粒子間結合なし",
        ]),
        ("PHASE 2", "Grain boundary", [
            "脱水・加熱で粒子が近接",
            "粒界（grain boundary）形成",
            "初期ネック形成開始",
        ]),
        ("PHASE 3", "Necking & porosity", [
            "ネックが拡大・強化",
            "空隙率低下し骨格が完成",
            "SiO₂エアロゲル層が確立",
        ]),
    ]
    for i, (tag, title, bullets) in enumerate(phases):
        cx = BODY_X + i * (cw + Inches(0.14))
        card(sl, cx, cy, cw, ch, title=title, tag=tag,
             bullets=bullets, ct_size=15, li_size=12)


# ===== SLIDE 31: 燃焼時間別SEM =====
def slide_31_burnsem(prs):
    sl = new_slide(prs)
    draw_header(sl, "補足 2/2", "燃焼時間に伴う焼結進行（SEM）")
    draw_footer(sl, "31 / 36")

    fig_h = Inches(2.3)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 5b",
                    caption="HEC+MC/CSP 1-5 を 0, 1, 2, 4分加熱後のクロスセクションSEM像")

    cy = BODY_Y + fig_h + Inches(0.45)
    ch = Inches(2.7)
    cw = (BODY_W - Inches(0.42)) / 4
    stages = [
        ("0 min", "未加熱", MUTED, [
            "球状粒子が独立分散",
            "粒子間結合なし",
            "水がマトリクスを充填",
        ]),
        ("1 min", "焼結初期", MUTED, [
            "粒子間で粒界が出現",
            "初期ネックが形成",
            "水分の大半が蒸発済み",
        ]),
        ("2 min", "焼結進行", MUTED, [
            "ネック領域が拡大",
            "連続ネットワーク形成",
            "空隙率が低下",
        ]),
        ("4 min", "エアロゲル完成", ACCENT, [
            "緻密なSiO₂骨格完成",
            "高い機械的強度",
            "断熱層として機能",
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
    draw_header(sl, "考察 1/3", "実装シナリオ：WEGの展開戦略")
    draw_footer(sl, "33 / 36")

    ch = Inches(3.6)
    cw = (BODY_W - Inches(0.28)) / 3
    cards_data = [
        ("WUI住宅の事前保護", "SCENARIO 1", [
            "山火事シーズン前に建物外壁・屋根に散布",
            "ヘリコプター・地上ホース双方で対応可能",
            "乾燥しても保護機能が残る",
        ]),
        ("重要インフラの保護", "SCENARIO 2", [
            "送電塔・変電所・通信基地局",
            "金属構造物にも強固に付着",
            "断電・通信遮断を予防",
        ]),
        ("防火線（Fire Break）形成", "SCENARIO 3", [
            "延焼経路となる植生帯に事前散布",
            "Phos-Chek の環境負荷なし",
            "セルロース・シリカは残留性が低い",
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
         "既存手法との優位点", size=12, bold=True, color=MUTED)
    advs = [
        ("vs 水", "保護時間 5× 向上"),
        ("vs AquaGel-K", "水蒸発後も継続保護"),
        ("vs Phos-Chek", "土壌・水系への残留なし"),
        ("散布互換性", "既存機材そのまま使用可"),
    ]
    av = cy + Inches(0.45)
    for k, v in advs:
        text(sl, BODY_X + Inches(0.25), av, Inches(1.6), Inches(0.25),
             k, size=11, color=MUTED)
        text(sl, BODY_X + Inches(1.95), av, cw2 - Inches(2.2), Inches(0.25),
             v, size=11, bold=True, color=D_GRN)
        av += Inches(0.28)

    callout(sl, BODY_X + cw2 + Inches(0.20), cy + Inches(0.3), cw2, Inches(1.0),
            "航空散布・地上散布ともに対応可能な流動特性（剪断希薄化）が実用化の鍵",
            icon='🚁')


# ===== SLIDE 34: 実証結果のまとめ =====
def slide_34_future(prs):
    sl = new_slide(prs)
    draw_header(sl, "まとめ 1/2", "実証された主要結果")
    draw_footer(sl, "34 / 36")

    cw = (BODY_W - Inches(0.42)) / 4
    ch = Inches(4.6)
    cards_data = [
        ("RHEOLOGY", "レオロジー特性", [
            "全系で G' > G''、crossoverなし",
            "G': HEC+MC≈46, MHEC≈26 Pa",
            "剪断希薄化（HB適合）→ 噴霧可",
            "tan δ < 1 の固体的ゲル",
        ]),
        ("FIRE PROTECTION", "火炎保護性能", [
            "HEC+MC/CSP：7分超 char遅延",
            "MHEC/CSP：5分超",
            "市販品の3〜6倍の効果",
            "水~0.3分, AquaGel-K~1.5分",
        ]),
        ("FOAMING / SEM", "発泡・微細構造", [
            "Foaming Index 最大≈2.6",
            "AquaGel-Kは発泡せず≈0",
            "燃焼で多孔質シリカ層を形成",
            "2分以降に粒子が焼結（SEM）",
        ]),
        ("CHEMISTRY", "化学的変態", [
            "FT-IR：CH/SiO比が低下",
            "XPS：燃焼で炭素が大幅減",
            "MHEC/CSP C 38.3%→4.8%",
            "残存はシリカ（SiO₂）",
        ]),
    ]
    for i, (tag, title, bullets) in enumerate(cards_data):
        cx = BODY_X + i * (cw + Inches(0.14))
        card(sl, cx, BODY_Y, cw, ch, title=title, tag=tag,
             bullets=bullets, ct_size=14, li_size=11)

    cy_co = BODY_Y + ch + Inches(0.25)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.80),
            "火炎で水が蒸発しつつゲルが多孔質シリカエアロゲルへ転換 → 乾燥後も基材を断熱保護（市販WEGにない機構）",
            dark=True, icon='◆')


# ===== SLIDE 35: 結論 =====
def slide_35_summary(prs):
    sl = new_slide(prs)
    draw_header(sl, "まとめ 2/2", "研究の結論")
    draw_footer(sl, "35 / 36")

    cw = (BODY_W - Inches(0.42)) / 4
    ch = Inches(2.5)
    summaries = [
        ("RESULT 01", "市販品の3〜6倍の保護", "HEC+MC/CSPは7分超、MHEC/CSPは5分超 char遅延。火炎下で基材を長く保護"),
        ("RESULT 02", "エアロゲルの自己形成", "加熱で脱水・CSPが焼結 → in situで多孔質シリカエアロゲル断熱層を形成（乾燥後も保護継続）"),
        ("RESULT 03", "既存散布インフラ互換", "剪断希薄化を示す噴霧可能な流体。基材への高い付着性・濡れ性を両立"),
        ("RESULT 04", "持続可能・安全な原料", "地球上最も豊富なセルロース誘導体。前研究で生分解性を確認、シリカは無害"),
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
         title="核心メカニズム",
         bullets=[
             "火炎接触で水が急速に蒸発・脱水",
             "CSPが焼結し粒子間ネックを形成",
             "多孔質シリカエアロゲル断熱層が完成",
         ], ct_size=14, li_size=12)
    card(sl, BODY_X + cw2 + Inches(0.20), cy, cw2, ch2,
         title="論文が示す意義",
         bullets=[
             "PPプラットフォームは多様な難燃材料の基盤",
             "modular製造・大規模適用に展開可能",
             "455日経時後も難燃性を維持（MHEC/CSP）",
         ], ct_size=14, li_size=12)

    cy_co = cy + ch2 + Inches(0.20)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.55),
            "「水のキャリア」から「加熱で自己変態する難燃材料」へ ― 乾燥後も効力を保つ次世代WEG",
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
         "従来品（水）比 Time to char 向上率", size=22, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER)

    text(sl, Inches(1.5), Inches(4.25), Inches(10.3), Inches(1.4),
         "「水のキャリア」から「自己変態する難燃材料」へ\nWEGの新しい設計パラダイムは山火事から重要インフラを守る実用技術への道を開く",
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

    print("PPTX を1から構築中...")
    slide_01_cover(prs)
    slide_02_toc(prs)
    slide_section_divider(prs, "01", "Part 1", "研究背景",
                          "気候変動下の山火事被害と、既存難燃技術が抱える本質的な限界",
                          ["山火事の現状", "既存WEG / 難燃剤", "本研究の革新"],
                          "3 / 36")
    slide_04_wildfire(prs)
    slide_05_process(prs)
    slide_06_existing(prs)
    slide_07_core(prs)
    slide_section_divider(prs, "02", "Part 2", "材料・実験方法",
                          "セルロース系ポリマーとコロイダルシリカの組み合わせによる新規ゲル設計",
                          ["ポリマー成分", "シリカ粒子・界面活性剤", "配合系（5種）", "評価手法"],
                          "8 / 36")
    slide_09_polymers(prs)
    slide_10_mc(prs)
    slide_11_csp_sds(prs)
    slide_12_formulations(prs)
    slide_13_methods(prs)
    slide_14_adhesion(prs)
    slide_section_divider(prs, "03", "Part 3", "実験結果",
                          "レオロジー特性、燃焼性能、発泡構造、エアロゲル形成メカニズムまで",
                          ["レオロジー", "燃焼試験", "発泡指数", "SEM観察"],
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
    slide_section_divider(prs, "04", "Part 4", "考察・まとめ",
                          "実装シナリオ、既存品との総合比較、研究の意義と今後の展望",
                          ["実装シナリオ", "性能比較", "まとめ"],
                          "32 / 36")
    slide_33_scenarios(prs)
    slide_34_future(prs)
    slide_35_summary(prs)
    slide_36_impact(prs)

    out_path = '/home/user/my-first-claude/slides.pptx'
    prs.save(out_path)
    print(f"完了: {out_path}")
    print(f"スライド数: {len(prs.slides.__iter__.__self__._sldIdLst)}")


if __name__ == '__main__':
    main()
