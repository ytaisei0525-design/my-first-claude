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


def fig_placeholder(sl, x, y, w, h, fig_num=None, caption=None, cited=False):
    """論文の図のプレースホルダー（ユーザーが後で画像を貼り付ける用）
    cited=True で他論文からの引用図（アンバーのバッジ＋出典メモ欄）"""
    # 背景（ダッシュではなく実線、薄い色で）
    s = rrect(sl, x, y, w, h, PH_BG, PH_BD, 1.0, radius=0.02)
    badge_fill = AMBER if cited else HEADER
    place_text = "（引用図を貼り付け）" if cited else "（論文の図を貼り付け）"
    # 図番号バッジ（左上）
    if fig_num:
        tag_w = Inches(0.95)
        rect(sl, x + Inches(0.12), y + Inches(0.12), tag_w, Inches(0.32),
             badge_fill)
        text(sl, x + Inches(0.12), y + Inches(0.12), tag_w, Inches(0.32),
             fig_num, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 中央プレースホルダーテキスト
    text(sl, x, y, w, h,
         place_text, size=12, color=GRAY_L,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 引用図は出典メモ欄をボックス下部内側に
    if cited:
        text(sl, x + Inches(0.12), y + h - Inches(0.38), w - Inches(0.24), Inches(0.28),
             "出典: ________________（後で記入）", size=9, color=D_AMBR, italic=True)
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
         "1 / 43", size=10, color=RGBColor(0x66, 0x77, 0x88), align=PP_ALIGN.RIGHT)


# ===== SLIDE 2: TOC =====
def slide_02_toc(prs):
    sl = new_slide(prs)
    draw_header(sl, "Contents", "発表の構成")
    draw_footer(sl, "2 / 43")

    parts = [
        ("PART 1", "研究背景", "山火事の現状、既存技術（Phos-Chek, AquaGel-K）の限界、本研究の革新点", "Slides 4–7"),
        ("PART 2", "材料・方法", "セルロース系ポリマー、CSP、配合系（5種）、評価手法", "Slides 9–16"),
        ("PART 3", "実験結果", "レオロジー、燃焼試験、発泡指数、SEM観察、メカニズム", "Slides 18–34"),
        ("PART 4", "考察・まとめ", "実装シナリオ、既存品との比較、結論と今後の展望", "Slides 36–39"),
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
    draw_header(sl, "背景 1/5", "山火事被害の深刻化とWUIリスク")
    draw_footer(sl, "4 / 43")

    # 左42% テキスト（タイポグラフィ主体） / 右58% 写真プレースホルダー
    left_w = BODY_W * 0.42 - Inches(0.12)
    right_w = BODY_W * 0.58 - Inches(0.08)
    right_x = BODY_X + left_w + Inches(0.20)

    # 縦アクセントバー（第1統計）
    rect(sl, BODY_X, BODY_Y, Inches(0.05), Inches(1.80), ACCENT)

    # 第1統計：大タイポグラフィ（ボックスなし）
    text(sl, BODY_X + Inches(0.18), BODY_Y, left_w - Inches(0.18), Inches(1.10),
         "5×", size=68, bold=True, color=INK)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(1.08), left_w - Inches(0.18), Inches(0.32),
         "大規模山火事の増加（過去30年）", size=13, bold=True, color=INK2)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(1.40), left_w - Inches(0.18), Inches(0.28),
         "気候変動・可燃物の蓄積が主要因（米・欧・豪）", size=11, color=MUTED)

    # 区切り線
    rect(sl, BODY_X + Inches(0.18), BODY_Y + Inches(1.86), left_w * 0.75, Inches(0.02), LINE)

    # 縦アクセントバー（第2統計）
    rect(sl, BODY_X, BODY_Y + Inches(2.08), Inches(0.05), Inches(1.60), D_AMBR)

    # 第2統計
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(2.08), left_w - Inches(0.18), Inches(0.90),
         "4500万人", size=40, bold=True, color=INK)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(2.98), left_w - Inches(0.18), Inches(0.32),
         "米国WUI居住者", size=13, bold=True, color=INK2)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(3.30), left_w - Inches(0.18), Inches(0.28),
         "野生地と都市の境界帯（WUI）の高リスク層", size=11, color=MUTED)

    # 下部メモ（横線区切り・calloutなし）
    rect(sl, BODY_X + Inches(0.18), BODY_Y + Inches(3.78), left_w * 0.90, Inches(0.02), LINE)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(3.96), left_w - Inches(0.18), Inches(0.70),
         "経済・インフラ・自然資源・WUI住民に甚大な被害 ─\n環境に優しい新たな火災遅延剤の開発が急務",
         size=12, color=INK3, italic=True)

    # 右：上=山火事写真／下=引用トレンドグラフ
    rtop_h = Inches(3.2)
    fig_placeholder(sl, right_x, BODY_Y, right_w, rtop_h,
                    fig_num="Image",
                    caption="山火事・WUI境界帯（写真）")
    rbot_y = BODY_Y + rtop_h + Inches(0.46)
    rbot_h = Inches(2.0)
    fig_placeholder(sl, right_x, rbot_y, right_w, rbot_h,
                    fig_num="引用", cited=True,
                    caption="山火事の大規模化・年間焼失面積の経年トレンド")


# ===== SLIDE 5: 4段階の保護プロセス =====
def slide_05_process(prs):
    sl = new_slide(prs)
    draw_header(sl, "背景 2/5", "本研究のアプローチ：4段階の保護プロセス")
    draw_footer(sl, "5 / 43")

    # 上部：大きな図プレースホルダー
    fig_h = Inches(3.5)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 1a",
                    caption="難燃ゲル散布 → 火炎接触 → 熱活性化エアロゲル形成 → 構造物の保護")

    # 下部：タイムライン風ステップ（数字円 + テキスト、カードなし）
    sy = BODY_Y + fig_h + Inches(0.28)
    cw = (BODY_W - Inches(0.42)) / 4
    circ_d = Inches(0.44)
    steps = [
        ("1", "ゲル散布", "セルロース系ゲルを\n建物・植生にスプレー塗布"),
        ("2", "火炎接触", "水分蒸発、ゲルが発泡\nしながら粒子が凝集"),
        ("3", "エアロゲル形成", "シリカ粒子が焼結し\n多孔質断熱層を形成"),
        ("4", "継続保護", "超低熱伝導の断熱層が\n基材を継続的に守る"),
    ]
    for i, (num, title, body) in enumerate(steps):
        cx = BODY_X + i * (cw + Inches(0.14))
        circ_x = cx + (cw - circ_d) / 2

        # 数字円
        rrect(sl, circ_x, sy, circ_d, circ_d, ACCENT, None, radius=0.5)
        text(sl, circ_x, sy, circ_d, circ_d, num, size=16, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # 矢印コネクター（最後以外）
        if i < 3:
            arr_x = cx + cw + Inches(0.04)
            arr_y = sy + circ_d / 2 - Inches(0.01)
            rect(sl, arr_x, arr_y, Inches(0.06), Inches(0.02), GRAY_L)

        # テキスト
        text(sl, cx, sy + circ_d + Inches(0.14), cw, Inches(0.30),
             title, size=13, bold=True, color=INK, align=PP_ALIGN.CENTER)
        text(sl, cx, sy + circ_d + Inches(0.46), cw, Inches(0.60),
             body, size=11, color=INK3, align=PP_ALIGN.CENTER)


# ===== SLIDE 6: 既存技術と限界（米国森林局の3分類） =====
def _draw_retardant(sl, bx, by, bw, bh):
    """長期遅延剤：航空機が赤色スラリーを空中散布するイラスト"""
    SKY  = RGBColor(0xa8, 0xd4, 0xef)
    GRND = RGBColor(0x6a, 0xa5, 0x5e)
    PLN  = RGBColor(0x55, 0x6a, 0x7e)
    DROP = RGBColor(0xc0, 0x3c, 0x3c)

    rrect(sl, bx, by, bw, bh, SKY, None, radius=0.03)
    rect(sl, bx, by + bh * 0.78, bw, bh * 0.22, GRND)
    rect(sl, bx, by + bh * 0.78, bw, Inches(0.018), RGBColor(0x4e, 0x88, 0x42))

    # 胴体
    bdy_w, bdy_h = bw * 0.30, bh * 0.08
    bdy_x = bx + bw * 0.10
    bdy_y = by + bh * 0.22
    rect(sl, bdy_x, bdy_y, bdy_w, bdy_h, PLN)
    # 主翼
    rect(sl, bdy_x + bdy_w * 0.28, bdy_y - bh * 0.08, bw * 0.10, bh * 0.22, PLN)
    # 垂直尾翼
    rect(sl, bdy_x + bw * 0.01, bdy_y - bh * 0.10, bw * 0.04, bh * 0.12, PLN)

    # 赤色スラリー散布（楕円）
    for fx, fy, fw, fh in [
        (0.35, 0.44, 0.040, 0.088), (0.44, 0.53, 0.036, 0.078),
        (0.52, 0.46, 0.038, 0.082), (0.42, 0.65, 0.032, 0.070),
        (0.58, 0.58, 0.034, 0.074), (0.50, 0.70, 0.030, 0.066),
        (0.62, 0.50, 0.028, 0.062),
    ]:
        oval(sl, bx + bw * fx, by + bh * fy, bw * fw, bh * fh, DROP)


def _draw_foam(sl, bx, by, bw, bh):
    """泡消火剤：消防ノズルから泡が広がるイラスト"""
    BG    = RGBColor(0xbe, 0xe0, 0xf5)
    NOZL  = RGBColor(0x2c, 0x3e, 0x50)
    FOAM  = RGBColor(0xf2, 0xf8, 0xfe)
    FOAM2 = RGBColor(0xc8, 0xe4, 0xf5)

    rrect(sl, bx, by, bw, bh, BG, None, radius=0.03)

    # ホース＋ノズル
    rect(sl, bx + bw * 0.02, by + bh * 0.50, bw * 0.07, bh * 0.40, NOZL)
    rect(sl, bx + bw * 0.05, by + bh * 0.36, bw * 0.10, bh * 0.28, NOZL)
    rect(sl, bx + bw * 0.13, by + bh * 0.43, bw * 0.04, bh * 0.14, NOZL)

    # 泡（扇状に広がる重なり合う楕円）
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
    """水強化ゲル（WEG）：ゲルを塗布された家のイラスト"""
    BG   = RGBColor(0x1e, 0x29, 0x3b)
    WALL = RGBColor(0xe8, 0xf0, 0xf8)
    ROOF = RGBColor(0xcc, 0xd8, 0xe8)
    GEL  = RGBColor(0x4a, 0xaa, 0x92)
    GEL2 = RGBColor(0x38, 0x90, 0x7a)
    WIN  = RGBColor(0x70, 0xaa, 0xcc)

    rrect(sl, bx, by, bw, bh, BG, None, radius=0.03)

    # 地面
    rect(sl, bx, by + bh * 0.88, bw, bh * 0.12, RGBColor(0x2e, 0x4a, 0x38))

    # 家の中心位置
    h_w  = bw * 0.32
    h_x  = bx + bw * 0.34
    w_y  = by + bh * 0.48   # 壁の上端
    w_h  = bh * 0.40        # 壁の高さ

    # ゲルの輝き（後ろ）
    oval(sl, h_x - bw * 0.03, w_y - bh * 0.06,
         h_w + bw * 0.06, w_h + bh * 0.12, GEL)

    # 壁
    rect(sl, h_x, w_y, h_w, w_h, WALL)

    # 屋根（三角形）
    s = sl.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
        h_x - bw * 0.02, by + bh * 0.18,
        h_w + bw * 0.04, bh * 0.32)
    s.fill.solid(); s.fill.fore_color.rgb = ROOF
    s.line.fill.background(); s.shadow.inherit = False

    # 窓
    for wi in range(2):
        rect(sl, h_x + h_w * 0.10 + wi * h_w * 0.50,
             w_y + w_h * 0.14, h_w * 0.26, w_h * 0.30, WIN)

    # ドア
    rect(sl, h_x + h_w * 0.36, w_y + w_h * 0.58,
         h_w * 0.28, w_h * 0.42, WIN)

    # ゲル滴（周囲に散布）
    for fx, fy, fw, fh in [
        (0.10, 0.20, 0.028, 0.076), (0.15, 0.42, 0.024, 0.064),
        (0.20, 0.30, 0.026, 0.070), (0.80, 0.22, 0.026, 0.074),
        (0.84, 0.42, 0.024, 0.066), (0.76, 0.54, 0.028, 0.070),
    ]:
        oval(sl, bx + bw * fx, by + bh * fy, bw * fw, bh * fh, GEL2)


def slide_06_existing(prs):
    sl = new_slide(prs)
    draw_header(sl, "背景 3/5", "既存の山火事用消火剤と、その限界（米国森林局の3分類）")
    draw_footer(sl, "6 / 43")

    n = 3
    gap = Inches(0.22)
    cw = (BODY_W - gap * (n - 1)) / n
    card_h = BODY_H - Inches(0.78)

    cats = [
        ("長期遅延剤", "Long-term retardant",
         "Phos-Chek（リン酸塩系スラリー）",
         "航空機からの赤色スラリー散布",
         "リン酸アンモニウム等の化学物質",
         "残留物がある限り効果が持続",
         "化学物質が対象物に残留する", False),
        ("泡消火剤", "Foam suppressant",
         "Class A フォーム／旧 AFFF",
         "消防ホースからの泡放水",
         "界面活性剤（旧来はフッ素系）",
         "水分保持 15〜30分（短期的）",
         "乾燥で効果消失／フッ素系は生体蓄積・環境毒性", False),
        ("水強化ゲル（WEG）", "Water-enhancing gel",
         "Barricade／Thermo-Gel／AquaGel-K",
         "住宅・植生へのゲル事前塗布",
         "高吸水性ポリマー（環境配慮型）",
         "水分保持 30〜60分・建物保護に有効",
         "高温・強風で乾燥すると効果を完全に喪失", True),
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
        en_c = RGBColor(0xa0, 0xb8, 0xd0) if is_hero else MUTED
        body_c = RGBColor(0xd8, 0xe2, 0xee) if is_hero else INK3
        label_c = RGBColor(0x88, 0xa0, 0xb8) if is_hero else MUTED
        div_c = RGBColor(0x44, 0x58, 0x70) if is_hero else LINE
        accent_c = GOLD if is_hero else ACCENT

        # イラスト（AIが図形で生成）
        ph_y = BODY_Y + Inches(0.18)
        ph_h = Inches(1.18)
        if i == 0:
            _draw_retardant(sl, cx + pad, ph_y, iw, ph_h)
        elif i == 1:
            _draw_foam(sl, cx + pad, ph_y, iw, ph_h)
        else:
            _draw_gel_house(sl, cx + pad, ph_y, iw, ph_h)

        # 名称＋英名
        ty = ph_y + ph_h + Inches(0.12)
        text(sl, cx + pad, ty, iw, Inches(0.36), name, size=16, bold=True, color=name_c)
        text(sl, cx + pad, ty + Inches(0.36), iw, Inches(0.22), en, size=9, italic=True, color=en_c)
        rect(sl, cx + pad, ty + Inches(0.64), iw, Inches(0.012), div_c)

        # 例（製品名）
        text(sl, cx + pad, ty + Inches(0.74), iw, Inches(0.22), "例", size=9, bold=True, color=accent_c)
        text(sl, cx + pad, ty + Inches(0.94), iw, Inches(0.40), example, size=11, bold=True, color=body_c)

        # 主成分
        text(sl, cx + pad, ty + Inches(1.40), iw, Inches(0.22), "主成分", size=9, bold=True, color=label_c)
        text(sl, cx + pad, ty + Inches(1.60), iw, Inches(0.40), comp, size=11, color=body_c)

        # 効果・保持
        text(sl, cx + pad, ty + Inches(2.06), iw, Inches(0.22), "効果・保持", size=9, bold=True, color=label_c)
        text(sl, cx + pad, ty + Inches(2.26), iw, Inches(0.40), merit, size=11, color=body_c)

        # 課題
        limit_label_c = GOLD if is_hero else D_RED
        text(sl, cx + pad, ty + Inches(2.72), iw, Inches(0.22), "課題", size=9, bold=True, color=limit_label_c)
        text(sl, cx + pad, ty + Inches(2.92), iw, card_h - (ty - BODY_Y) - Inches(3.10),
             limit, size=11, bold=True, color=(WHITE if is_hero else D_RED))

    # 下部メッセージ
    msg_y = BODY_Y + card_h + Inches(0.16)
    callout(sl, BODY_X, msg_y, BODY_W, Inches(0.50),
            "本研究は、環境配慮型WEGが持つ「乾燥すると無効になる」致命的な弱点を克服する",
            icon='→')


# ===== SLIDE 7: 先行研究（PNP ゲル：APP キャリアとして使用、吸着・自己修復は既知）=====
def slide_07_predecessor(prs):
    sl = new_slide(prs)
    draw_header(sl, "背景 4/5", "先行研究：PNP ゲル基盤の確立（Yu et al., PNAS 2016）")
    draw_footer(sl, "7 / 43")

    # 上部リード
    callout(sl, BODY_X, BODY_Y, BODY_W, Inches(0.62),
            "同じ HEC+MC/CSP プラットフォームは、リン酸アンモニウム剤（Phos-Chek LC95A）の "
            "「キャリア」として 2016 年に確立済み — 本研究はこの基盤の上に立つ",
            icon='📚')

    # 2列レイアウト：左＝先行研究で確立されたこと、右＝本研究の新規性
    cw = (BODY_W - Inches(0.30)) / 2
    cx2 = BODY_X + cw + Inches(0.30)
    ctop = BODY_Y + Inches(0.78)
    ch = Inches(4.20)

    # 左：先行研究で確立済み
    rrect(sl, BODY_X, ctop, cw, ch, CARD_C, LINE, 0.5, radius=0.04)
    rect(sl, BODY_X, ctop, cw, Inches(0.05), ACCENT)
    text(sl, BODY_X + Inches(0.24), ctop + Inches(0.18), cw - Inches(0.48), Inches(0.30),
         "Yu et al., PNAS 2016 で確立済み", size=11, bold=True, color=MUTED)
    text(sl, BODY_X + Inches(0.24), ctop + Inches(0.50), cw - Inches(0.48), Inches(0.40),
         "PNP ゲル（HEC+MC + CSP）の基盤特性", size=15, bold=True, color=INK)

    items_old = [
        ("PP 相互作用", "セルロース鎖が CSP 表面に選択吸着（非共有・多価）"),
        ("剪断希薄化", "高せん断で粘度が低下（噴霧・パイプ送液に適合）"),
        ("自己修復性", "応力解放で即座にゲル構造を回復"),
        ("スケーラビリティ", "0.5 mL → 15 L まで線形にスケール"),
        ("APP キャリア用途", "Phos-Chek LC95A を担持し、付着・耐降雨性を向上"),
    ]
    by = ctop + Inches(1.05)
    for label, body in items_old:
        rrect(sl, BODY_X + Inches(0.24), by + Inches(0.05),
              Inches(1.55), Inches(0.32),
              WHITE, ACCENT, 0.7, radius=0.5)
        text(sl, BODY_X + Inches(0.24), by + Inches(0.05),
             Inches(1.55), Inches(0.32),
             label, size=10, bold=True, color=ACCENT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(sl, BODY_X + Inches(1.88), by, cw - Inches(2.08), Inches(0.42),
             body, size=11, color=INK3, anchor=MSO_ANCHOR.MIDDLE)
        by += Inches(0.56)

    # 右：本研究の新規性
    rrect(sl, cx2, ctop, cw, ch, HEADER, None, radius=0.04)
    rect(sl, cx2, ctop, cw, Inches(0.05), GOLD)
    text(sl, cx2 + Inches(0.24), ctop + Inches(0.18), cw - Inches(0.48), Inches(0.30),
         "本研究で新たに獲得した機能", size=11, bold=True, color=GOLD)
    text(sl, cx2 + Inches(0.24), ctop + Inches(0.50), cw - Inches(0.48), Inches(0.40),
         "ゲル自体が「断熱層」へと変態する", size=15, bold=True, color=WHITE)

    items_new = [
        ("自立型保護", "APP に依存せず、ゲル単独で延焼を抑制"),
        ("熱活性化", "炎接触で水分蒸発 → CSP 焼結 → 多孔質シリカ"),
        ("エアロゲル化", "断熱性の高い多孔質シリカ層を in situ で形成"),
        ("乾燥耐性", "従来 WEG の致命的弱点「乾けば無効」を克服"),
        ("長期安定性", "455 日後も降伏応力を維持（補足 3/3 参照）"),
    ]
    by = ctop + Inches(1.05)
    for label, body in items_new:
        rrect(sl, cx2 + Inches(0.24), by + Inches(0.05),
              Inches(1.55), Inches(0.32),
              HEADER, GOLD, 0.7, radius=0.5)
        text(sl, cx2 + Inches(0.24), by + Inches(0.05),
             Inches(1.55), Inches(0.32),
             label, size=10, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(sl, cx2 + Inches(1.88), by, cw - Inches(2.08), Inches(0.42),
             body, size=11, color=RGBColor(0xd8, 0xe2, 0xee),
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
         "本研究の革新点：「化学物質を運ぶ容器」から「自ら断熱層になる材料」へ — "
         "PP ゲルの再定義",
         size=13, color=INK2, anchor=MSO_ANCHOR.MIDDLE)


# ===== SLIDE 8: 研究の核心 =====
def slide_07_core(prs):
    sl = new_slide(prs)
    draw_header(sl, "背景 5/5", "本研究の核心：熱活性化エアロゲル形成")
    draw_footer(sl, "8 / 43")

    # 左38% テキスト / 右62% 図プレースホルダー（図を主役に）
    left_w = BODY_W * 0.38 - Inches(0.12)
    right_w = BODY_W * 0.62 - Inches(0.08)
    right_x = BODY_X + left_w + Inches(0.20)

    # 縦アクセントバー
    rect(sl, BODY_X, BODY_Y, Inches(0.05), Inches(1.52), ACCENT)

    # ヒーロー文（引用スタイル、ボックスなし）
    text(sl, BODY_X + Inches(0.18), BODY_Y, left_w - Inches(0.18), Inches(0.52),
         "Heat-Activated", size=22, bold=True, color=INK, italic=True)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(0.50), left_w - Inches(0.18), Inches(0.52),
         "Aerogel Formation", size=22, bold=True, color=ACCENT, italic=True)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(1.04), left_w - Inches(0.18), Inches(0.34),
         "火炎接触をトリガーに、ゲル自体が断熱材へ変態", size=11, color=INK3)

    # 区切り線
    rect(sl, BODY_X + Inches(0.18), BODY_Y + Inches(1.54), left_w * 0.85, Inches(0.02), LINE)

    # 提案手法とメカニズム（番号付きリスト、カードなし）
    points = [
        ("01", "材料構成", "セルロース系バイオポリマー × コロイダルシリカの物理架橋"),
        ("02", "エアロゲル化", "炎で水分が失われると多孔質シリカ網がその場形成"),
        ("03", "持続的防護", "水分が枯渇した後も断熱バリアが発火を防ぐ"),
        ("04", "応用性・拡張性", "噴霧・ポンプ送液に適した流動特性／大規模製造も可能"),
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

    # 右：上=概念図／下=引用2図（エアロゲル微細構造＋断熱デモ）
    rtop_h = Inches(3.1)
    fig_placeholder(sl, right_x, BODY_Y, right_w, rtop_h,
                    fig_num="概念図",
                    caption="加熱トリガーでゲルが多孔質エアロゲル断熱層へ変態")
    rbot_y = BODY_Y + rtop_h + Inches(0.42)
    rbot_h = Inches(2.2)
    sub_w = (right_w - Inches(0.18)) / 2
    fig_placeholder(sl, right_x, rbot_y, sub_w, rbot_h,
                    fig_num="引用", cited=True,
                    caption="シリカエアロゲルのSEM微細構造")
    fig_placeholder(sl, right_x + sub_w + Inches(0.18), rbot_y, sub_w, rbot_h,
                    fig_num="引用", cited=True,
                    caption="エアロゲルの断熱性デモ（炎上の試料等）")


# ===== SLIDE 9 (new): 材料構成：持続可能かつ高性能 =====
def slide_mat_composition(prs):
    sl = new_slide(prs)
    draw_header(sl, "材料 概要 1/2", "材料構成：持続可能かつ高性能")
    draw_footer(sl, "10 / 43")

    # リード文
    lead_h = Inches(0.50)
    text(sl, BODY_X, BODY_Y, BODY_W, lead_h,
         "本ゲルは「骨格・機能・結合」を担う3要素で構成され、生分解性と高性能を両立します。",
         size=14, color=INK2, anchor=MSO_ANCHOR.MIDDLE)

    n = 3
    gap = Inches(0.30)
    cw = (BODY_W - gap * (n - 1)) / n
    cards_y = BODY_Y + lead_h + Inches(0.14)
    card_h = H - cards_y - FTR_H - Inches(0.24)

    items = [
        ("🌿", "セルロース誘導体", "Cellulose derivatives",
         "植物由来のバイオポリマー。ゲルネットワークの骨格を形成します。",
         [("3種", "HEC・MC・MHEC を使用"),
          ("生分解性", "環境負荷が低い"),
          ("役割", "粘性・構造の土台")],
         ACCENT),
        ("◎", "コロイダルシリカ（CSP）", "Colloidal Silica Particles",
         "ポリマーと動的に相互作用し、熱でエアロゲルを形成する主役成分です。",
         [("直径", "約 22 nm の単分散粒子"),
          ("焼結", "加熱でシリカ骨格を形成"),
          ("役割", "断熱エアロゲルの源")],
         D_TEAL),
        ("⬡", "PP 相互作用", "Polymer–Particle Interaction",
         "動的な多価水素結合でネットワークを構築。共有結合に依存しません。",
         [("結合", "動的・多価の水素結合"),
          ("自己修復", "ひずみ後にゲルが回復"),
          ("噴霧適性", "スプレー塗布が容易")],
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
             title, size=17, bold=True, color=INK, align=PP_ALIGN.CENTER)
        text(sl, cx + pad_x, title_y + Inches(0.42), inner_w, Inches(0.28),
             en, size=10, italic=True, color=MUTED, align=PP_ALIGN.CENTER)

        div_y = title_y + Inches(0.78)
        rect(sl, cx + Inches(0.32), div_y, cw - Inches(0.64), Inches(0.015), LINE)

        body_y = div_y + Inches(0.14)
        text(sl, cx + pad_x, body_y, inner_w, Inches(1.00),
             body, size=12, color=INK3, align=PP_ALIGN.CENTER)

        ky = body_y + Inches(1.22)
        row_h = Inches(0.74)
        for j, (label, detail) in enumerate(keys):
            ry = ky + j * row_h
            chip_w = Inches(0.92)
            rrect(sl, cx + pad_x, ry + Inches(0.04), chip_w, Inches(0.32),
                  WHITE, accent_c, 0.7, radius=0.5)
            text(sl, cx + pad_x, ry + Inches(0.04), chip_w, Inches(0.32),
                 label, size=10, bold=True, color=accent_c,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            text(sl, cx + pad_x + chip_w + Inches(0.10), ry,
                 inner_w - chip_w - Inches(0.10), Inches(0.40),
                 detail, size=11, color=INK2, anchor=MSO_ANCHOR.MIDDLE)


# ===== SLIDE 10: 噴霧・付着・耐火のメカニズム =====
def slide_self_protection(prs):
    sl = new_slide(prs)
    draw_header(sl, "材料 概要 2/2", "噴霧・付着・耐火のメカニズム")
    draw_footer(sl, "11 / 43")

    # リード文
    lead_h = Inches(0.52)
    text(sl, BODY_X, BODY_Y, BODY_W, lead_h,
         "温度ではなく「せん断力」と「自己修復性」が、噴霧から付着までを支配します。",
         size=15, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)

    # 4段階の大きなプロセスカード（全幅）
    n = 4
    arrow_w = Inches(0.30)
    cards_y = BODY_Y + lead_h + Inches(0.14)
    cw = (BODY_W - arrow_w * (n - 1)) / n
    card_h = Inches(3.50)

    stages = [
        ("01", "静止時", "ゲル", "高粘度の弾性ゲル。G′ > G″ で構造を保持。",
         "降伏応力 ~33 Pa", ACCENT),
        ("02", "噴霧時", "ゾル様", "せん断ひずみで粘度が低下し、流動化。",
         "せん断希薄化（粘度低下）", D_TEAL),
        ("03", "付着後", "自己修復ゲル", "ひずみ解放で即座にゲル構造を回復。",
         "G′ 回復率 ~90%", D_GRN),
        ("04", "炎接触", "エアロゲル", "CSP が焼結し多孔質シリカ層を形成。",
         "断熱バリアとして機能", D_AMBR),
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
             state, size=22, bold=True, color=INK)
        rect(sl, cx + pad_x, cards_y + Inches(1.58), inner_w, Inches(0.015), LINE)
        text(sl, cx + pad_x, cards_y + Inches(1.72), inner_w, Inches(1.10),
             desc, size=12, color=INK3)

        badge_h = Inches(0.46)
        badge_y = cards_y + card_h - badge_h - Inches(0.18)
        rrect(sl, cx + pad_x, badge_y, inner_w, badge_h, WHITE, c, 0.7, radius=0.06)
        text(sl, cx + pad_x, badge_y, inner_w, badge_h,
             metric, size=11, bold=True, color=c,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        if i < n - 1:
            text(sl, cx + cw, cards_y, arrow_w, card_h,
                 "→", size=20, color=GRAY_L,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 下部：ポイント帯
    band_y = cards_y + card_h + Inches(0.20)
    band_h = H - band_y - FTR_H - Inches(0.22)
    rrect(sl, BODY_X, band_y, BODY_W, band_h, HEADER, None, radius=0.05)
    rect(sl, BODY_X, band_y, Inches(0.08), band_h, GOLD)
    text(sl, BODY_X + Inches(0.30), band_y, Inches(2.4), band_h,
         "POINT", size=13, bold=True, color=GOLD, anchor=MSO_ANCHOR.MIDDLE)
    text(sl, BODY_X + Inches(2.0), band_y, BODY_W - Inches(2.3), band_h,
         "MC の温度応答性（LCST）に頼らず、レオロジー特性だけで「飛ばす→留める→固める」を実現する点が本材料の核心です。",
         size=13, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)


# ===== SLIDE 9: セルロース系ポリマー =====
def slide_09_polymers(prs):
    sl = new_slide(prs)
    draw_header(sl, "材料 1/6", "使用したセルロース系ポリマー")
    draw_footer(sl, "12 / 43")

    # 左55% ポリマー縦リスト（大文字バッジ＋区切り線）/ 右45% 構造図
    left_w = BODY_W * 0.55 - Inches(0.12)
    right_w = BODY_W * 0.45 - Inches(0.08)
    right_x = BODY_X + left_w + Inches(0.20)

    polymers = [
        ("A", "HEC", "Hydroxyethyl cellulose", "ベース粘性を付与する骨格成分"),
        ("B", "MC", "Methyl cellulose", "加熱でゲル化（熱可逆性）── 火炎保護の鍵"),
        ("C", "MHEC", "Methyl 2-hydroxyethyl cellulose", "AとBの特性を併せ持つ複合型"),
    ]
    row_h = BODY_H / 3
    for i, (letter, abbr, full, role) in enumerate(polymers):
        ry = BODY_Y + i * row_h
        # 大きな文字バッジ（薄色・タイポグラフィ装飾）
        text(sl, BODY_X, ry, Inches(0.9), row_h,
             letter, size=54, bold=True, color=LINE, anchor=MSO_ANCHOR.MIDDLE)
        tx = BODY_X + Inches(0.95)
        text(sl, tx, ry + Inches(0.22), left_w - Inches(0.95), Inches(0.45),
             abbr, size=24, bold=True, color=INK)
        text(sl, tx, ry + Inches(0.70), left_w - Inches(0.95), Inches(0.28),
             full, size=11, color=MUTED)
        text(sl, tx, ry + Inches(1.00), left_w - Inches(0.95), Inches(0.34),
             role, size=12, bold=True, color=ACCENT)
        # 区切り線（最後以外）
        if i < 2:
            rect(sl, BODY_X, ry + row_h - Inches(0.02), left_w, Inches(0.015), LINE)

    # 右：上=本論文の構造図／下=引用セルロース分子構造
    rtop_h = Inches(3.3)
    fig_placeholder(sl, right_x, BODY_Y, right_w, rtop_h,
                    fig_num="Fig. 1b/c",
                    caption="各ポリマーの化学構造とCSP混合スキーム")
    rbot_y = BODY_Y + rtop_h + Inches(0.46)
    rbot_h = Inches(2.0)
    fig_placeholder(sl, right_x, rbot_y, right_w, rbot_h,
                    fig_num="引用", cited=True,
                    caption="セルロースの分子構造（参考）")


# ===== SLIDE 12: レオロジー設計（せん断希薄化・自己修復性）=====
def slide_10_mc(prs):
    sl = new_slide(prs)
    draw_header(sl, "材料 2/6", "レオロジー設計：せん断希薄化と自己修復性")
    draw_footer(sl, "13 / 43")

    left_w = BODY_W * 0.48 - Inches(0.10)
    right_w = BODY_W * 0.52 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    # 左上：粘度 vs せん断速度グラフ（シアシニング挙動）
    img_h = Inches(2.7)
    fig_placeholder(sl, BODY_X, BODY_Y, left_w, img_h,
                    fig_num="Fig. 2",
                    caption="粘度 vs せん断速度（シアシニング挙動）")

    # 左下：主要特性
    ty = BODY_Y + img_h + Inches(0.20)
    rect(sl, BODY_X, ty, left_w * 0.9, Inches(0.015), LINE)
    ty += Inches(0.18)
    props = [
        ("せん断希薄化", "高せん断で粘度が急低下（噴霧時に低粘度化）"),
        ("降伏応力", "~33 Pa（静止時のゲル構造を維持）"),
        ("自己修復性", "ひずみ除去後 G' が ~90% 回復"),
    ]
    for label, detail in props:
        text(sl, BODY_X, ty, Inches(1.10), Inches(0.30),
             label, size=10, bold=True, color=ACCENT)
        text(sl, BODY_X + Inches(1.14), ty, left_w - Inches(1.14), Inches(0.30),
             detail, size=12, color=INK2)
        ty += Inches(0.40)

    # 右：3つのmini stats
    mw = (right_w - Inches(0.20)) / 3
    mh = Inches(1.4)
    mini(sl, right_x, BODY_Y, mw, mh, "粘度挙動", "可逆", "高せん断で低粘度→静止で回復", value_color=ACCENT)
    mini(sl, right_x + mw + Inches(0.10), BODY_Y, mw, mh, "動的降伏応力", "~33 Pa", "HEC+MC/CSP（HB fit）", value_color=D_TEAL)
    mini(sl, right_x + (mw + Inches(0.10)) * 2, BODY_Y, mw, mh, "G' 回復率", "~90%", "1000秒後", value_color=D_GRN)

    # 噴霧〜耐火プロセスフロー (4段)
    fy = BODY_Y + mh + Inches(0.30)
    fh = Inches(2.7)
    rrect(sl, right_x, fy, right_w, fh, HEADER, None, radius=0.04)
    text(sl, right_x + Inches(0.20), fy + Inches(0.18), right_w - Inches(0.4), Inches(0.30),
         "噴霧〜耐火プロセス", size=11, bold=True, color=RGBColor(0xa0, 0xb8, 0xd0))

    sub_y = fy + Inches(0.65)
    sub_h = fh - Inches(0.75)
    pw = (right_w - Inches(0.40) - Inches(0.30) * 3) / 4
    steps_t = [
        ("静止時", "ゲル状態", "G' > G''"),
        ("噴霧時", "せん断希薄化", "低粘度・ゾル様"),
        ("付着後", "自己修復", "G' 回復"),
        ("炎接触", "エアロゲル化", "多孔質シリカ層"),
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
    draw_header(sl, "材料 3/6", "コロイダルシリカ粒子(CSP)と界面活性剤(SDS)")
    draw_footer(sl, "14 / 43")

    cw = (BODY_W - Inches(0.18)) / 2
    callout_h = Inches(0.72)
    img_h = Inches(2.4)
    text_top = BODY_Y + img_h + Inches(0.18)

    # 左：CSP — 上部TEM写真 + 下部テキスト
    cx = BODY_X
    fig_placeholder(sl, cx, BODY_Y, cw, img_h,
                    fig_num="TEM",
                    caption="CSP粒子（22 nm、単分散）")
    text(sl, cx, text_top, cw, Inches(0.34),
         "CSP — Colloidal Silica Particles", size=14, bold=True, color=INK)
    bullets_csp = [
        "LUDOX TM-50（原液 50 wt%）→ 15 wt%に希釈（pH 9）",
        "粒子サイズ：22 nm（単分散）",
        "配合濃度：ゲル中 5 wt%",
        "加熱時に焼結 → silica aerogel に変態",
    ]
    by = text_top + Inches(0.38)
    for b in bullets_csp:
        text(sl, cx, by, Inches(0.20), Inches(0.28), "•", size=12, color=ACCENT)
        text(sl, cx + Inches(0.22), by, cw - Inches(0.22), Inches(0.28), b, size=12, color=INK2)
        by += Inches(0.34)

    # 右：SDS — 上部SEM写真（発泡構造）+ 下部テキスト
    cx2 = BODY_X + cw + Inches(0.18)
    fig_placeholder(sl, cx2, BODY_Y, cw, img_h,
                    fig_num="SEM",
                    caption="SDS添加量による発泡構造の変化")
    text(sl, cx2, text_top, cw, Inches(0.34),
         "SDS — Sodium Dodecyl Sulfate", size=14, bold=True, color=INK)
    bullets_sds = [
        "陰イオン性界面活性剤（添加剤）",
        "SDS添加でも Foaming Index は改善しなかった",
        "添加濃度：0.1 wt% ・ 0.5 wt% の2水準",
        "SDS量増加で気泡が粗大化（SEM確認）",
    ]
    by2 = text_top + Inches(0.38)
    for b in bullets_sds:
        text(sl, cx2, by2, Inches(0.20), Inches(0.28), "•", size=12, color=ACCENT)
        text(sl, cx2 + Inches(0.22), by2, cw - Inches(0.22), Inches(0.28), b, size=12, color=INK2)
        by2 += Inches(0.34)

    # 下部callout
    cy_co = H - FTR_H - Inches(0.10) - callout_h
    callout(sl, BODY_X, cy_co, BODY_W, callout_h,
            "CSPの焼結温度域とセルロースの熱分解温度が重なるため、熱応答が同期して断熱層を形成",
            icon='💡')


# ===== SLIDE 12: 配合系 5種 =====
def slide_12_formulations(prs):
    sl = new_slide(prs)
    draw_header(sl, "材料 4/6", "評価した5種類の配合系")
    draw_footer(sl, "15 / 43")

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

    # 下部：2カード ＋ 配合ゲル外観写真
    cy = fy + fh + Inches(0.40)
    ch = Inches(2.5)
    cw3 = (BODY_W - Inches(0.36)) / 3
    card(sl, BODY_X, cy, cw3, ch,
         title="配合設計の意図",
         bullets=[
             "ポリマー種を変えて機能差を比較",
             "SDS濃度を2水準で発泡構造を検証",
             "市販品AquaGel-Kと直接比較",
         ], ct_size=14, li_size=11)
    card(sl, BODY_X + cw3 + Inches(0.18), cy, cw3, ch,
         title="記法の読み方",
         body="HEC+MC/CSP/SDS 1-5-0.1\n→ HEC+MC 1 wt%\n／ CSP 5 wt%\n／ SDS 0.1 wt%",
         ct_size=14, cb_size=13)
    fig_placeholder(sl, BODY_X + (cw3 + Inches(0.18)) * 2, cy, cw3, ch - Inches(0.34),
                    fig_num="Photo",
                    caption="5配合系のゲル外観")


# ===== SLIDE 13: 評価手法 =====
def slide_13_methods(prs):
    sl = new_slide(prs)
    draw_header(sl, "材料 5/6", "評価手法の全体像")
    draw_footer(sl, "16 / 43")

    # 左：燃焼試験セットアップ図 ／ 右：6手法カード(2列×3行)
    left_w = BODY_W * 0.34 - Inches(0.10)
    right_w = BODY_W * 0.66 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)
    fig_placeholder(sl, BODY_X, BODY_Y, left_w, BODY_H - Inches(0.34),
                    fig_num="Setup",
                    caption="燃焼試験セットアップ（MAP-Proトーチ／合板基板）")

    methods = [
        ("レオロジー測定", "RHEOLOGY",
         ["振動周波数掃引（G', G''）", "定常流動掃引（粘度）", "Herschel-Bulkley適合"]),
        ("燃焼試験", "BURN TEST",
         ["MAP-Proトーチ(~2054°C)で加熱", "炭化開始までの時間を計測", "120/300 s時点を比較"]),
        ("発泡指数測定", "FOAMING",
         ["燃焼後の発泡層厚さを計測", "初期厚さ比 = Foaming Index"]),
        ("SEM形態観察", "SEM",
         ["SDS濃度別の発泡構造", "燃焼時間別の焼結進行"]),
        ("分光分析", "SPECTROSCOPY",
         ["FT-IR（化学結合）", "XPS（表面組成）"]),
        ("熱分析", "THERMAL",
         ["TGA（熱重量）", "DSC（熱量）"]),
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
    draw_header(sl, "材料 6/6", "レオロジー設計が実装性能に直結する")
    draw_footer(sl, "17 / 43")

    # レイアウト：左（測定されたレオロジー特性）→ 矢印 →（期待される実装性能）の3行
    left_w = Inches(4.5)
    arrow_w = Inches(0.7)
    gap = Inches(0.18)
    right_x = BODY_X + left_w + gap + arrow_w + gap
    right_w = BODY_W - left_w - arrow_w - gap * 2

    # 列見出し
    lbl_y = BODY_Y
    text(sl, BODY_X, lbl_y, left_w, Inches(0.30),
         "レオロジー設計（実測された物性）", size=12, bold=True, color=MUTED)
    text(sl, right_x, lbl_y, right_w, Inches(0.30),
         "期待される実装性能（現場での挙動）", size=12, bold=True, color=MUTED)

    # 3行の因果マッピング：(左タイトル, 左説明, 右タイトル, 右説明, アクセント色)
    rows = [
        ("せん断希薄化挙動", "高せん断で粘度が急低下",
         "噴霧できる", "ノズル通過時の高せん断で粘度が急低下\n→ 既存ホース・ノズルで散布可能", D_TEAL),
        ("高い貯蔵弾性率 G'（G' ≫ G''）", "静止時はゲル状態を保持",
         "垂直面で流れ落ちない", "塗布後は形状を保持して付着\n→ 木材・壁面など垂直基材にも保持", ACCENT),
        ("自己修復性", "step-strain で G' が即時回復",
         "到達後すぐ保護膜を形成", "散布直後に再ゲル化\n→ 風雨・自重で流出せず防火層を維持", GOLD),
    ]

    row_y = BODY_Y + Inches(0.42)
    row_h = Inches(1.32)
    row_gap = Inches(0.20)
    for ltitle, ldesc, rtitle, rdesc, c in rows:
        # 左：実測物性カード
        rrect(sl, BODY_X, row_y, left_w, row_h, CARD_C, LINE, 0.5, radius=0.05)
        rect(sl, BODY_X, row_y, Inches(0.06), row_h, c)
        text(sl, BODY_X + Inches(0.22), row_y + Inches(0.18), left_w - Inches(0.34), Inches(0.40),
             ltitle, size=15, bold=True, color=INK)
        text(sl, BODY_X + Inches(0.22), row_y + Inches(0.70), left_w - Inches(0.34), Inches(0.50),
             ldesc, size=12, color=INK2)

        # 中央：矢印
        ar = sl.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                 BODY_X + left_w + gap, row_y + row_h/2 - Inches(0.22),
                                 arrow_w, Inches(0.44))
        ar.fill.solid(); ar.fill.fore_color.rgb = c
        ar.line.fill.background(); ar.shadow.inherit = False

        # 右：実装性能カード
        rrect(sl, right_x, row_y, right_w, row_h, WHITE, c, 1.0, radius=0.05)
        text(sl, right_x + Inches(0.22), row_y + Inches(0.16), right_w - Inches(0.34), Inches(0.40),
             rtitle, size=15, bold=True, color=c)
        text(sl, right_x + Inches(0.22), row_y + Inches(0.62), right_w - Inches(0.34), Inches(0.62),
             rdesc, size=12, color=INK2)

        row_y += row_h + row_gap

    # 下部：正直な注記（独立した付着・接触角試験は本論文では未実施）
    cy_co = H - FTR_H - Inches(0.16) - Inches(0.62)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.62),
            "これらの実装性能はレオロジー実測値からの帰結。接触角・付着力の独立した測定試験は本論文では行われていない",
            icon='ⓘ')


# ===== SLIDE 16: G'/G'' =====
def slide_16_rheology1(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 1/15", "レオロジー①：粘弾性とせん断希薄化（Fig. 2a–c）")
    draw_footer(sl, "19 / 43")

    # 2図を横並び（左=粘弾性 2a–b、右=せん断希薄化 2c）
    gap = Inches(0.24)
    fw = (BODY_W - gap) / 2
    fig_h = Inches(3.5)

    fig_placeholder(sl, BODY_X, BODY_Y, fw, fig_h,
                    fig_num="Fig. 2a–b",
                    caption="角周波数 vs G' / G''（静止時の粘弾性）")
    cx2 = BODY_X + fw + gap
    fig_placeholder(sl, cx2, BODY_Y, fw, fig_h,
                    fig_num="Fig. 2c",
                    caption="せん断速度 vs 粘度（せん断希薄化）")

    # 観察カード 2枚
    cy = BODY_Y + fig_h + Inches(0.22)
    ch = Inches(1.0)
    # 左カード：ゲル挙動
    rrect(sl, BODY_X, cy, fw, ch, CARD_C, D_TEAL, 1.2, radius=0.05)
    rect(sl, BODY_X, cy, Inches(0.06), ch, D_TEAL)
    text(sl, BODY_X + Inches(0.20), cy + Inches(0.12), fw - Inches(0.34), Inches(0.30),
         "静止時：G' > G'' → 固体的（ゲル）挙動", size=12.5, bold=True, color=D_TEAL)
    text(sl, BODY_X + Inches(0.20), cy + Inches(0.48), fw - Inches(0.34), Inches(0.45),
         "周波数依存性が小さく安定なネットワーク → 塗布後に流れ落ちない",
         size=11.5, color=INK2)
    # 右カード：シアシニング
    rrect(sl, cx2, cy, fw, ch, CARD_C, ACCENT, 1.2, radius=0.05)
    rect(sl, cx2, cy, Inches(0.06), ch, ACCENT)
    text(sl, cx2 + Inches(0.20), cy + Inches(0.12), fw - Inches(0.34), Inches(0.30),
         "高せん断時：粘度が急低下（せん断希薄化）", size=12.5, bold=True, color=ACCENT)
    text(sl, cx2 + Inches(0.20), cy + Inches(0.48), fw - Inches(0.34), Inches(0.45),
         "ノズル通過時の高せん断で粘度低下 → スプレー噴霧が可能",
         size=11.5, color=INK2)

    # 下部 callout
    cy_co = cy + ch + Inches(0.18)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.56),
            "静止時はゲル（垂れない）／噴霧時は低粘度（飛ばせる）の両立がスプレー散布の基盤",
            dark=True, icon='✓')


# ===== SLIDE 17: shear thinning =====
def slide_17_rheology2(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 2/15", "レオロジー②：静的・動的降伏応力と AquaGel-K 比較（Fig. 2d・S3）")
    draw_footer(sl, "20 / 43")

    # 上部 callout：降伏応力とは
    callout(sl, BODY_X, BODY_Y, BODY_W, Inches(0.62),
            "降伏応力＝ゲルが流れ出す／流れ続けるのに必要な応力。静的（流れ始め＝付着）と動的（流動維持＝噴霧）の2定義で評価",
            icon='ƒ')

    top = BODY_Y + Inches(0.82)
    left_w = BODY_W * 0.40 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.24)
    right_w = BODY_W - left_w - Inches(0.24)

    # 左：2図（静的=S3、動的=2d）を縦に
    fig_h = Inches(1.70)
    fig_placeholder(sl, BODY_X, top, left_w, fig_h,
                    fig_num="SI Fig. S3",
                    caption="振幅掃引：G'/G'' クロスオーバー（静的降伏応力）")
    fig_placeholder(sl, BODY_X, top + fig_h + Inches(0.15), left_w, fig_h,
                    fig_num="Fig. 2d",
                    caption="HB fit による動的降伏応力（棒グラフ）")

    # 右：2つの降伏応力カード
    card_h = Inches(1.70)
    # 静的降伏応力
    rrect(sl, right_x, top, right_w, card_h, CARD_C, D_TEAL, 1.2, radius=0.05)
    rect(sl, right_x, top, Inches(0.06), card_h, D_TEAL)
    text(sl, right_x + Inches(0.22), top + Inches(0.12), right_w - Inches(0.34), Inches(0.30),
         "静的降伏応力 σ_static（流れ始め＝付着）", size=13, bold=True, color=D_TEAL)
    text(sl, right_x + Inches(0.22), top + Inches(0.50), right_w - Inches(0.34), Inches(0.30),
         "測定：振幅掃引（LAOS）の G' / G'' クロスオーバー応力（S3）", size=11.5, color=INK2)
    text(sl, right_x + Inches(0.22), top + Inches(0.84), right_w - Inches(0.34), Inches(0.80),
         "PPハイドロゲルは AquaGel-K より高い静的降伏応力\n→ 垂直・高所の燃料面でも流れ落ちにくく、強く付着する",
         size=11.5, color=INK3)

    # 動的降伏応力
    top2 = top + card_h + Inches(0.15)
    rrect(sl, right_x, top2, right_w, card_h, CARD_C, ACCENT, 1.2, radius=0.05)
    rect(sl, right_x, top2, Inches(0.06), card_h, ACCENT)
    text(sl, right_x + Inches(0.22), top2 + Inches(0.12), right_w - Inches(0.34), Inches(0.30),
         "動的降伏応力 σ_dynamic（流動維持＝噴霧）", size=13, bold=True, color=ACCENT)
    text(sl, right_x + Inches(0.22), top2 + Inches(0.50), right_w - Inches(0.34), Inches(0.30),
         "測定：フロースイープ（100→0.01 s⁻¹）の HB モデル fit（2d）", size=11.5, color=INK2)
    text(sl, right_x + Inches(0.22), top2 + Inches(0.84), right_w - Inches(0.34), Inches(0.80),
         "AquaGel-K（基準線 0.05 Pa）に対し PPハイドロゲルは桁違いに高い\n→ ポンプ圧内で噴霧でき、かつ流れにくい設計を両立",
         size=11.5, color=INK3)

    # 下部 callout
    cy_co = top + fig_h * 2 + Inches(0.15) + Inches(0.22)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.62),
            "静的・動的いずれの降伏応力でも PPハイドロゲルは市販 AquaGel-K を大きく上回る。"
            "両者の差が小さい（弱いチキソトロピー）ため構造が即再形成され、自己修復的に付着を保つ",
            dark=True, icon='✓')


# ===== SLIDE 18: Herschel-Bulkley =====
def slide_18_hb(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 3/15", "レオロジー③：動的降伏応力の経時変化と長期安定性")
    draw_footer(sl, "21 / 43")

    # 上部 callout
    callout(sl, BODY_X, BODY_Y, BODY_W, Inches(0.65),
            "動的降伏応力：定常フロースイープを HB モデルにフィットして算出（σ_d）。流動状態を維持するのに必要な応力（Fig. S7）",
            icon='ƒ')

    ty = BODY_Y + Inches(0.85)
    th = Inches(2.2)
    headers = ["配合系", "動的降伏応力（Day 1）", "動的降伏応力（Day 455）", "備考"]
    rows = [
        [{'text': "HEC+MC/CSP 1-5", 'bold': True},
         {'text': "33.34 Pa", 'bold': True, 'color': ACCENT},
         {'text': "68.9 Pa", 'color': D_GRN},
         "ゲル強度が経時で増大"],
        [{'text': "MHEC/CSP 1-5", 'bold': True},
         {'text': "3.31 Pa", 'bold': True, 'color': ACCENT},
         {'text': "4.56 Pa", 'color': D_GRN},
         "長期安定性を確認（455日）"],
        ["AquaGel-K（市販対照）", "—", "—",
         {'text': "比較対照", 'color': MUTED}],
    ]
    simple_table(sl, BODY_X, ty, BODY_W, th, headers, rows,
                 col_widths=[2.8, 2.0, 2.0, 2.0])

    # 下部 3カード
    cy = ty + th + Inches(0.30)
    ch = Inches(1.85)
    cw = (BODY_W - Inches(0.30)) / 3
    cards_data = [
        ("動的降伏応力（HB fit）", "定常フロースイープを HB モデルにフィットして測定。HEC+MC/CSPは33.34 Pa、MHEC/CSPは3.31 Pa（Day 1）"),
        ("長期安定性", "MHEC/CSPは455日経時後も3.31→4.56 Paとわずかな変化のみ。実用配合の長期安定性を実証"),
        ("HEC+MC vs MHEC", "HEC+MC系は経時で剛性化（33.34→68.9 Pa）。MHEC系は変化が穏やかで単一ポリマーにより製造も簡便"),
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
    draw_header(sl, "結果 4/15", "レオロジー④：貯蔵弾性率 G' と粘弾性（tan δ）")
    draw_footer(sl, "22 / 43")

    cw = (BODY_W - Inches(0.30)) / 2
    ch = Inches(5.6)
    cx2 = BODY_X + cw + Inches(0.30)

    # 左：動的降伏応力
    text(sl, BODY_X, BODY_Y, cw, Inches(0.30),
         "G' 貯蔵弾性率（1 rad/s）", size=12, bold=True, color=MUTED)
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
            "全系でG'>G''（ゲル挙動）。AquaGel-KのG'が最高だがtan δ も大きく流動的",
            icon='ⓘ')

    # 右：tan δ
    text(sl, cx2, BODY_Y, cw, Inches(0.30),
         "損失正接 tan δ = G'' / G'（弾性支配 = <1）", size=12, bold=True, color=MUTED)
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
            "全系でtan δ < 1 → ゲル挙動。AquaGel-K が最も弾性的（0.168）。MHEC系は流動性が高め（0.400）",
            dark=True, icon='✓')


# ===== SLIDE 19b: 自己修復性（step-strain recovery） =====
def slide_self_healing(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 5/15", "自己修復性：Step-strain による G' / G'' 即時回復")
    draw_footer(sl, "23 / 43")

    # 2カラム構成
    cw = (BODY_W - Inches(0.30)) / 2
    cx2 = BODY_X + cw + Inches(0.30)

    # ===== 左：試験プロトコル概念図 =====
    text(sl, BODY_X, BODY_Y, cw, Inches(0.30),
         "Step-strain プロトコル（ω = 1 rad/s）", size=12, bold=True, color=MUTED)

    # プロトコル説明（3ステップを縦に並べる）
    steps = [
        ("Step 1", "低歪み（γ = 1%）", "ネットワーク構造を保持\nG' > G''（ゲル状態）", D_TEAL),
        ("Step 2", "高歪み（γ = 500%）", "ネットワーク破壊\nG' < G''（液体状態）", D_AMBR),
        ("Step 3", "低歪みへ復帰（γ = 1%）", "粒子-高分子相互作用が再形成\nG' が即時回復", ACCENT),
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

    # ===== 右：G'/G'' 回復曲線プレースホルダー =====
    text(sl, cx2, BODY_Y, cw, Inches(0.30),
         "G' / G'' の時系列変化（cyclic step-strain）", size=12, bold=True, color=MUTED)
    fig_placeholder(sl, cx2, BODY_Y + Inches(0.45), cw, Inches(3.4),
                    fig_num="SI（step-strain）",
                    caption="γ=1%→500%→1% サイクルで G' が即時回復")

    # 右下：キーポイント
    kp_y = BODY_Y + Inches(0.45) + Inches(3.4) + Inches(0.20)
    rrect(sl, cx2, kp_y, cw, Inches(1.10), CARD_C, LINE, 0.5, radius=0.04)
    rect(sl, cx2, kp_y, Inches(0.05), Inches(1.10), ACCENT)
    text(sl, cx2 + Inches(0.18), kp_y + Inches(0.10), cw - Inches(0.28), Inches(0.30),
         "POINT：噴霧後の即時付着を保証", size=12, bold=True, color=ACCENT)
    text(sl, cx2 + Inches(0.18), kp_y + Inches(0.42), cw - Inches(0.28), Inches(0.62),
         "ノズル通過時の高せん断で粘度低下 → 基材到達後に G' を即時回復\n→ 垂直面でも流れ落ちず付着保持（自己修復性が実装可能性の鍵）",
         size=10.5, color=INK2)

    # 下部 callout
    cy_co = H - FTR_H - Inches(0.16) - Inches(0.60)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.60),
            "非共有結合（水素結合・多価相互作用）による可逆ネットワーク → 自己修復性を実証",
            icon='✓')


# ===== SLIDE 20: HEC+MC vs MHEC =====
def slide_20_compare(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 6/15", "HEC+MC vs MHEC：二系統の比較")
    draw_footer(sl, "24 / 43")

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
        "せん断希薄化（高せん断で低粘度）",
        "Time to char ＞ 7 min",
        "Foaming Index 最大≈2.6（SDS無し）",
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
        "せん断希薄化挙動",
        "Time to char ＞ 5 min",
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


# ===== SLIDE SDS RHEOLOGY: Fig. 2e–h（SDS 添加によるレオロジー調整）=====
def slide_sds_rheology(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 7/15", "SDS 添加：目的は発泡促進・濡れ性改善だったが…（Fig. 2e–h）")
    draw_footer(sl, "25 / 43")

    # ── 上部：目的 2カード（左＝目的、右＝期待した効果）──
    hdr_h = Inches(0.90)
    cw2 = (BODY_W - Inches(0.18)) / 2
    # 左：目的
    rrect(sl, BODY_X, BODY_Y, cw2, hdr_h, CARD_C, D_TEAL, 1.2, radius=0.05)
    rect(sl, BODY_X, BODY_Y, Inches(0.06), hdr_h, D_TEAL)
    text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(0.08), cw2 - Inches(0.30), Inches(0.28),
         "SDS 添加の目的", size=12, bold=True, color=D_TEAL)
    text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(0.40), cw2 - Inches(0.30), Inches(0.45),
         "界面活性剤 SDS（0 / 0.1 / 0.5 wt%）を HEC+MC/CSP に添加し、"
         "燃料表面への濡れ性改善と発泡層形成の促進を狙った",
         size=11, color=INK2)
    # 右：期待
    cx2_hdr = BODY_X + cw2 + Inches(0.18)
    rrect(sl, cx2_hdr, BODY_Y, cw2, hdr_h, CARD_C, ACCENT, 1.2, radius=0.05)
    rect(sl, cx2_hdr, BODY_Y, Inches(0.06), hdr_h, ACCENT)
    text(sl, cx2_hdr + Inches(0.20), BODY_Y + Inches(0.08), cw2 - Inches(0.30), Inches(0.28),
         "期待した効果", size=12, bold=True, color=ACCENT)
    text(sl, cx2_hdr + Inches(0.20), BODY_Y + Inches(0.40), cw2 - Inches(0.30), Inches(0.45),
         "①ゲル粘度を下げてスプレー性を向上　②気泡を安定化して Foaming Index を増大",
         size=11, color=INK2)

    # ── 中央：2×2 グリッド（Fig. 2e–h）──
    gap = Inches(0.18)
    cw = (BODY_W - gap) / 2
    row1_y = BODY_Y + hdr_h + Inches(0.16)
    row_h  = Inches(2.05)
    row2_y = row1_y + row_h + gap

    panels = [
        (BODY_X,            row1_y, "Fig. 2e", "G' / G'' vs 角周波数（SDS 3濃度）"),
        (BODY_X + cw + gap, row1_y, "Fig. 2f", "G' と tan δ（1 rad/s、濃度別棒グラフ）"),
        (BODY_X,            row2_y, "Fig. 2g", "粘度 vs せん断速度（HB fit）"),
        (BODY_X + cw + gap, row2_y, "Fig. 2h", "動的降伏応力（SDS 濃度別・0.05 Pa 基準線）"),
    ]
    for px, py, fig_num, cap in panels:
        fig_placeholder(sl, px, py, cw, row_h, fig_num=fig_num, caption=cap)

    # ── 下部：結果評価（2列）──
    band_y = row2_y + row_h + Inches(0.16)
    band_h = H - band_y - FTR_H - Inches(0.16)
    res_w = (BODY_W - Inches(0.18)) / 2

    # 左：レオロジーへの影響（限定的改善）
    rrect(sl, BODY_X, band_y, res_w, band_h, CARD_C, D_AMBR, 1.2, radius=0.04)
    rect(sl, BODY_X, band_y, Inches(0.06), band_h, D_AMBR)
    text(sl, BODY_X + Inches(0.20), band_y + Inches(0.06), res_w - Inches(0.30), Inches(0.26),
         "レオロジー：濃度依存的だが限定的改善", size=11.5, bold=True, color=D_AMBR)
    text(sl, BODY_X + Inches(0.20), band_y + Inches(0.34), res_w - Inches(0.30), band_h - Inches(0.38),
         "0.1 wt% で G' がわずかに上昇するが、0.5 wt% ではむしろ軟化し降伏応力も低下。"
         "付着性・ゲル強度の観点では SDS 添加は全体的にマイナスに働く",
         size=11, color=INK3)

    # 右：発泡への影響（逆効果）
    cx2_res = BODY_X + res_w + Inches(0.18)
    rrect(sl, cx2_res, band_y, res_w, band_h, CARD_C, D_RED, 1.2, radius=0.04)
    rect(sl, cx2_res, band_y, Inches(0.06), band_h, D_RED)
    text(sl, cx2_res + Inches(0.20), band_y + Inches(0.06), res_w - Inches(0.30), Inches(0.26),
         "発泡：期待に反して Foaming Index が低下", size=11.5, bold=True, color=D_RED)
    text(sl, cx2_res + Inches(0.20), band_y + Inches(0.34), res_w - Inches(0.30), band_h - Inches(0.38),
         "SDS 添加により気泡が粗大化・不均一化し、Foaming Index は低下（Fig. 4 参照）。"
         "発泡促進という当初の目的は達成されなかった",
         size=11, color=INK3)


# ===== SLIDE 21: 燃焼試験 setup =====
def slide_21_setup(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 8/15", "燃焼試験のセットアップ")
    draw_footer(sl, "26 / 43")

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
             "基板：白木合板（whitewood plywood）",
             "ゲル散布厚：均一に塗布",
             "火炎源：MAP-Pro トーチ（~2054°C）",
             "計測：基材が炭化するまでの時間",
             "120 s時点・300 s時点で外観撮影",
         ], ct_size=15, li_size=12)
    callout(sl, right_x, BODY_Y + ch + Inches(0.20), right_w, Inches(0.85),
            "Time to char = 木材表面が炭化するまでに要した時間（長いほど高保護）",
            icon='⏱')


# ===== SLIDE 22: Time to char =====
def slide_22_ttc(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 9/15", "Time to char：5配合系の定量比較")
    draw_footer(sl, "27 / 43")

    left_w = BODY_W * 0.62 - Inches(0.10)
    right_w = BODY_W * 0.38 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    # 左：バーチャート
    text(sl, BODY_X, BODY_Y, left_w, Inches(0.3),
         "Time to char（炭化開始までの時間）", size=12, bold=True, color=MUTED)
    by = BODY_Y + Inches(0.45)
    bar_data = [
        ("Water", 0.04, "~0.3 min", GRAY_L),
        ("AquaGel-K（市販品）", 0.20, "~1.5 min", D_AMBR),
        ("MHEC/CSP 1-5", 0.65, ">5 min", D_AMBR),
        ("HEC+MC/CSP 1-5", 0.90, ">7 min", D_TEAL),
        ("HEC+MC/CSP/SDS 1-5-0.1", 0.90, ">7 min", D_BLUE),
    ]
    for label, ratio, val, color in bar_data:
        bar_row(sl, BODY_X, by, left_w, label, ratio, val, fill_color=color,
                lbl_w=Inches(2.6), val_w=Inches(1.1))
        by += Inches(0.55)

    callout(sl, BODY_X, by + Inches(0.15), left_w, Inches(0.85),
            "本研究WEGは市販AquaGel-Kの3〜6倍の保護時間を達成（全試験でn≥3）",
            icon='📊')

    # 右：図プレースホルダー
    fig_placeholder(sl, right_x, BODY_Y, right_w, Inches(5.2),
                    fig_num="Fig. 3b",
                    caption="Time-to-char バーチャート（n≧3、エラーバーは標準偏差）")


# ===== SLIDE 23: 燃焼時系列 =====
def slide_23_timelapse(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 10/15", "燃焼過程の時系列観察")
    draw_footer(sl, "28 / 43")

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
            "~0.3分（約18秒）で炭化",
            "保護層ゼロ",
        ]),
        ("AquaGel-K", D_AMBR, [
            "水を保持するが水蒸発で終了",
            "~1.5分で炭化",
            "固体保護層を形成しない",
        ]),
        ("本研究 WEG", ACCENT, [
            "水蒸発と同時にゲルが発泡・膨張",
            "多孔質エアロゲル層が継続保護",
            ">7分間（HEC+MC）炭化を阻止",
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
    draw_header(sl, "結果 11/15", "120秒 / 300秒時点の表面状態比較")
    draw_footer(sl, "29 / 43")

    fig_h = Inches(2.3)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 3d",
                    caption="120秒時点（上段）と300秒時点（下段）の表面状態比較")

    cy = BODY_Y + fig_h + Inches(0.50)
    ch = Inches(2.8)
    cw = (BODY_W - Inches(0.28)) / 3
    cards_data = [
        ("Water", D_RED, "Time to char：~0.3 min", [
            "120 s：既に広範囲が炭化開始",
            "300 s：表面全体が黒化・損傷甚大",
        ]),
        ("AquaGel-K（市販品）", D_AMBR, "Time to char：~1.5 min", [
            "120 s：ゲルが乾燥し一部炭化",
            "300 s：保護層なし、部分炭化",
        ]),
        ("本研究 WEG", ACCENT, "Time to char：>7 min", [
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
        {'text': ">7", 'size': 130, 'bold': True, 'color': GOLD},
        {'text': " min", 'size': 60, 'bold': False, 'color': RGBColor(0xc8, 0x9a, 0x68)},
    ], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    text(sl, Inches(0.5), Inches(4.4), Inches(12.3), Inches(0.6),
         "本研究WEGの Time to char（HEC+MC/CSP）", size=26, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER)
    text(sl, Inches(1.5), Inches(5.15), Inches(10.3), Inches(0.7),
         "市販AquaGel-K（~1.5分）の3〜6倍の保護時間を達成（全試験でn≥3）",
         size=14, color=RGBColor(0x9c, 0xb4, 0xcc), align=PP_ALIGN.CENTER)

    # bars
    bars = [
        ("Water", 0.04, "~0.3 min", GRAY_L),
        ("AquaGel-K", 0.20, "~1.5 min", D_AMBR),
        ("本研究 WEG", 0.95, ">7 min", GOLD),
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


# ===== SLIDE 26: 発泡指数 =====
def slide_26_foam(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 12/15", "SDS 添加：発泡性向上を期待したが、逆の結果に（Fig. 4）")
    draw_footer(sl, "31 / 43")

    # ── 上部：仮説 → ？ → 結果 の3ボックス ──
    bw = Inches(3.6)
    bh = Inches(1.10)
    arrow_w = Inches(0.70)
    gap = (BODY_W - bw * 3 - arrow_w * 2) / 2
    box_y = BODY_Y

    # 仮説ボックス
    rrect(sl, BODY_X, box_y, bw, bh, CARD_C, D_TEAL, 1.2, radius=0.05)
    rect(sl, BODY_X, box_y, Inches(0.06), bh, D_TEAL)
    text(sl, BODY_X + Inches(0.22), box_y + Inches(0.08), bw - Inches(0.30), Inches(0.24),
         "仮説（SDS 添加の目的）", size=10, bold=True, color=D_TEAL)
    text(sl, BODY_X + Inches(0.22), box_y + Inches(0.38), bw - Inches(0.30), Inches(0.64),
         "界面活性剤 SDS が発泡を促進\n→ Foaming Index が上昇\n→ エアロゲル断熱層が厚くなる",
         size=11, color=INK2)

    # → 矢印1
    ax1 = BODY_X + bw + gap * 0.5
    ar1 = sl.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, ax1, box_y + bh/2 - Inches(0.22),
                              arrow_w, Inches(0.44))
    ar1.fill.solid(); ar1.fill.fore_color.rgb = MUTED
    ar1.line.fill.background(); ar1.shadow.inherit = False

    # 実験ボックス
    bx2 = BODY_X + bw + gap * 0.5 + arrow_w + gap * 0.5
    rrect(sl, bx2, box_y, bw, bh, CARD_C, LINE, 0.8, radius=0.05)
    text(sl, bx2 + Inches(0.22), box_y + Inches(0.08), bw - Inches(0.30), Inches(0.24),
         "実験", size=10, bold=True, color=MUTED)
    text(sl, bx2 + Inches(0.22), box_y + Inches(0.38), bw - Inches(0.30), Inches(0.64),
         "SDS 0 / 0.1 / 0.5 wt% を添加した\nHEC+MC/CSP の燃焼試験\n→ Foaming Index を定量（Fig. 4a/4b）",
         size=11, color=INK2)

    # → 矢印2（赤系：予想外を示す）
    ax2 = bx2 + bw + gap * 0.5
    ar2 = sl.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, ax2, box_y + bh/2 - Inches(0.22),
                              arrow_w, Inches(0.44))
    ar2.fill.solid(); ar2.fill.fore_color.rgb = D_AMBR
    ar2.line.fill.background(); ar2.shadow.inherit = False

    # 結果ボックス（予想外）
    bx3 = ax2 + arrow_w + gap * 0.5
    rrect(sl, bx3, box_y, bw, bh, RGBColor(0xff, 0xf7, 0xed), D_AMBR, 1.4, radius=0.05)
    rect(sl, bx3, box_y, Inches(0.06), bh, D_AMBR)
    text(sl, bx3 + Inches(0.22), box_y + Inches(0.08), bw - Inches(0.30), Inches(0.24),
         "結果（予想と逆）", size=10, bold=True, color=D_AMBR)
    text(sl, bx3 + Inches(0.22), box_y + Inches(0.38), bw - Inches(0.30), Inches(0.64),
         "SDS 添加で Foaming Index が低下\n2.6（SDS 0%） → 2.3（0.1%）→ さらに低下（0.5%）\n発泡性向上には寄与せず",
         size=11, color=INK2)

    # ── 中央：Fig. 4a / 4b プレースホルダー ──
    fig_y = box_y + bh + Inches(0.28)
    fig_h = Inches(2.10)
    fw = (BODY_W - Inches(0.20)) / 2
    fig_placeholder(sl, BODY_X, fig_y, fw, fig_h,
                    fig_num="Fig. 4a", caption="燃焼後の発泡層外観（SDS 濃度別）")
    fig_placeholder(sl, BODY_X + fw + Inches(0.20), fig_y, fw, fig_h,
                    fig_num="Fig. 4b", caption="各配合系の Foaming Index（定量）")

    # ── 下部：Foaming Index 棒グラフ ──
    bar_y = fig_y + fig_h + Inches(0.28)
    bar_data = [
        ("HEC+MC/CSP 1-5（SDS 0%）", 1.00, "≈2.6×", D_TEAL),
        ("HEC+MC/CSP/SDS 1-5-0.1", 0.88, "≈2.3×", MUTED),
        ("HEC+MC/CSP/SDS 1-5-0.5", 0.72, "さらに低下", D_AMBR),
        ("AquaGel-K（参考）", 0.02, "≈0", GRAY_L),
    ]
    text(sl, BODY_X, bar_y, BODY_W * 0.5, Inches(0.26),
         "Foaming Index（発泡後厚 / 初期厚）", size=11, bold=True, color=MUTED)
    by = bar_y + Inches(0.32)
    for label, ratio, val, color in bar_data:
        bar_row(sl, BODY_X, by, BODY_W * 0.62, label, ratio, val,
                fill_color=color, lbl_w=Inches(3.2), val_w=Inches(1.0))
        by += Inches(0.40)

    # ── 下部 callout ──
    cy_co = H - FTR_H - Inches(0.14) - Inches(0.58)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.58),
            "SDS は発泡に寄与せず、むしろ気泡を粗大化・不均一化（Fig. 4c SEM）→ 発泡指数低下の原因",
            icon='!', dark=True)


# ===== SLIDE 27: SDS濃度別SEM =====
def slide_27_sem_sds(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 13/15", "SDS濃度がエアロゲル微細構造に与える影響")
    draw_footer(sl, "32 / 43")

    fig_h = Inches(3.0)
    fig_placeholder(sl, BODY_X, BODY_Y, BODY_W, fig_h,
                    fig_num="Fig. 4c",
                    caption="SDS濃度別のFE-SEM観察像（0%, 0.1%, 0.5%）")

    cy = BODY_Y + fig_h + Inches(0.45)
    ch = Inches(2.3)
    cw = (BODY_W - Inches(0.28)) / 3
    cards_data = [
        ("0% SDS", ACCENT, "最高の発泡指数（≈2.6）。SDSなしが膨張率最大。"),
        ("0.1% SDS", MUTED, "均一な微細気泡構造。発泡指数≈2.3（SDSなしより低い）。"),
        ("0.5% SDS（過剰）", MUTED, "気泡サイズ不均一・粗大化。発泡指数がさらに低下。"),
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
    draw_header(sl, "結果 14/15", "FT-IR・XPS：化学組成の変化（燃焼前後）")
    draw_footer(sl, "33 / 43")

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
    headers2 = ["配合系", "C 1s 燃焼前 (at%)", "C 1s 燃焼後 (at%)"]
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
            "炭素濃度が大幅減（MHEC: 38.3→4.8 at%）→ 有機マトリクスが燃焼除去、シリカが残存・露出",
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
    draw_header(sl, "結果 15/15", "TGA/DSC：熱分解プロファイルと各成分の役割")
    draw_footer(sl, "34 / 43")

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
    draw_header(sl, "補足 1/3", "シリカエアロゲル形成のメカニズム")
    draw_footer(sl, "35 / 43")

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
    draw_header(sl, "補足 2/3", "燃焼時間に伴う焼結進行（SEM）")
    draw_footer(sl, "36 / 43")

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


# ===== SLIDE 32: 経時変化・長期安定性（補足 3/3） =====
def slide_32_aging(prs):
    sl = new_slide(prs)
    draw_header(sl, "補足 3/3", "経時変化と長期安定性：1年以上経過しても性能を維持")
    draw_footer(sl, "38 / 43")

    left_w = BODY_W * 0.36 - Inches(0.12)
    right_w = BODY_W * 0.64 - Inches(0.08)
    right_x = BODY_X + left_w + Inches(0.20)

    # 左：熟成メカニズム
    rect(sl, BODY_X, BODY_Y, Inches(0.05), Inches(1.02), ACCENT)
    text(sl, BODY_X + Inches(0.18), BODY_Y, left_w - Inches(0.18), Inches(0.34),
         "熟成メカニズム", size=15, bold=True, color=INK)
    text(sl, BODY_X + Inches(0.18), BODY_Y + Inches(0.40), left_w - Inches(0.18), Inches(0.66),
         "可逆的な結合（水素結合）から、\n不可逆な共有結合（Si–O縮合）へ転移", size=12, color=INK2)

    # 区切り線
    rect(sl, BODY_X + Inches(0.18), BODY_Y + Inches(1.24), left_w * 0.85, Inches(0.015), LINE)

    # 2系統の比較
    rows = [
        ("HEC+MC / CSP", "共有結合化が進行 → 剛性が上昇", ACCENT),
        ("MHEC / CSP", "ほぼ可逆結合のまま → 穏やかな変化", D_TEAL),
    ]
    ry = BODY_Y + Inches(1.46)
    for name, desc, c in rows:
        text(sl, BODY_X + Inches(0.18), ry, left_w - Inches(0.18), Inches(0.30),
             name, size=13, bold=True, color=c)
        text(sl, BODY_X + Inches(0.18), ry + Inches(0.30), left_w - Inches(0.18), Inches(0.46),
             desc, size=11, color=INK3)
        ry += Inches(0.92)

    text(sl, BODY_X + Inches(0.18), ry, left_w - Inches(0.18), Inches(0.30),
         "Fig. S4 / S6（界面化学スキーム）", size=10, color=MUTED, italic=True)

    # 右上：2つのmini stats
    mw = (right_w - Inches(0.16)) / 2
    mh = Inches(1.5)
    mini(sl, right_x, BODY_Y, mw, mh, "弾性率（熟成20日）", "~1000 Pa",
         "HEC+MC/CSP・Fig S5", value_color=ACCENT)
    mini(sl, right_x + mw + Inches(0.16), BODY_Y, mw, mh, "降伏応力 Day1→Day455", "33 → 69 Pa",
         "HEC+MC/CSP・Fig S7", value_color=AMBER)

    # 右下：熟成レオロジー図プレースホルダー（Day 1 / Day 455 を左右に並べる）
    fy = BODY_Y + mh + Inches(0.28)
    fh = Inches(2.3)
    sub_w = (right_w - Inches(0.16)) / 2
    fig_placeholder(sl, right_x, fy, sub_w, fh,
                    fig_num="Fig. S7 (Day 1)",
                    caption="フロースイープ・HB fit（作製直後）")
    fig_placeholder(sl, right_x + sub_w + Inches(0.16), fy, sub_w, fh,
                    fig_num="Fig. S7 (Day 455)",
                    caption="フロースイープ・HB fit（455日経過後）")

    # 下部：ヒーローcallout
    cy_co = H - FTR_H - Inches(0.16) - Inches(0.82)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.82),
            "Day 455（約15か月）経過後も燃焼挙動とエアロゲル化は不変 — 実用的な保存安定性を実証（Video S5）",
            dark=True, icon='✓')


# ===== SLIDE 33: 実装シナリオ =====
def slide_33_scenarios(prs):
    sl = new_slide(prs)
    draw_header(sl, "考察 1/3", "実装シナリオ：WEGの展開戦略")
    draw_footer(sl, "40 / 43")

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
    box_h = Inches(1.95)
    rrect(sl, BODY_X, cy, cw2, box_h, WHITE, LINE, 0.5, radius=0.04)
    text(sl, BODY_X + Inches(0.20), cy + Inches(0.12), cw2 - Inches(0.4), Inches(0.28),
         "既存手法との優位点", size=12, bold=True, color=MUTED)
    advs = [
        ("vs 水", "保護時間が大幅向上（水は~0.3分で炭化）"),
        ("vs AquaGel-K", "3〜6倍の保護時間＋水蒸発後も継続"),
        ("vs Phos-Chek", "土壌・水系への残留性が低い"),
        ("散布互換性", "既存機材をそのまま使用可"),
    ]
    av = cy + Inches(0.50)
    lbl_x = BODY_X + Inches(0.22)
    val_x = BODY_X + Inches(1.70)
    for k, v in advs:
        text(sl, lbl_x, av, Inches(1.45), Inches(0.32),
             k, size=10.5, color=MUTED, anchor=MSO_ANCHOR.MIDDLE)
        text(sl, val_x, av, cw2 - Inches(1.90), Inches(0.32),
             v, size=10.5, bold=True, color=D_GRN, anchor=MSO_ANCHOR.MIDDLE)
        av += Inches(0.33)

    callout(sl, BODY_X + cw2 + Inches(0.20), cy + Inches(0.20), cw2, Inches(1.0),
            "航空散布・地上散布ともに対応可能な流動特性（剪断希薄化）が実用化の鍵",
            icon='🚁')

    # 出典注記
    note_y = cy + Inches(1.30)
    text(sl, BODY_X + cw2 + Inches(0.20), note_y, cw2, Inches(0.55),
         "※ 水・AquaGel-K との保護時間比較は実測（Fig. 3）。Phos-Chek は論文の燃焼直接比較ではなく、"
         "環境残留性に関する定性的な位置づけ",
         size=9.5, italic=True, color=MUTED)


# ===== SLIDE 34: 実証結果のまとめ（既存品との○×比較）=====
def slide_34_future(prs):
    sl = new_slide(prs)
    draw_header(sl, "まとめ 1/2", "既存品との比較：WEG だけが満たす要件")
    draw_footer(sl, "41 / 43")

    # 上部 callout
    callout(sl, BODY_X, BODY_Y, BODY_W, Inches(0.56),
            "本研究の WEG（PPハイドロゲル）は、水・市販WEG・Phos-Chek が個別にしか満たせない要件をすべて両立する",
            icon='◆')

    # ── ○×比較マトリクス ──
    ty = BODY_Y + Inches(0.78)
    # 列：製品、横：5つの評価軸
    crit = ["長時間の火炎保護", "乾燥後も保護継続", "噴霧散布が可能", "環境残留が低い", "発泡・断熱層を形成"]
    # 各行：(製品名, 強調か, [評価記号...])  記号: '◎'/'○'/'△'/'×'
    G = D_GRN      # ◎/○ 緑
    A = D_AMBR     # △ 琥珀
    R = D_RED      # × 赤
    rows = [
        ("WEG（本研究）", True,  [("◎", G), ("◎", G), ("○", G), ("○", G), ("◎", G)]),
        ("市販 AquaGel-K", False, [("△", A), ("×", R), ("○", G), ("○", G), ("×", R)]),
        ("Phos-Chek", False,     [("○", G), ("○", G), ("○", G), ("×", R), ("×", R)]),
        ("水のみ", False,         [("×", R), ("×", R), ("◎", G), ("◎", G), ("×", R)]),
    ]

    name_w = Inches(2.6)
    grid_w = BODY_W - name_w
    col_w = grid_w / len(crit)
    hdr_h = Inches(0.95)
    row_h = Inches(0.66)
    table_h = hdr_h + row_h * len(rows)

    # ヘッダー行（評価軸）
    rect(sl, BODY_X, ty, name_w, hdr_h, HEADER)
    text(sl, BODY_X + Inches(0.12), ty, name_w - Inches(0.24), hdr_h,
         "製品 ＼ 要件", size=11, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    for ci, c in enumerate(crit):
        cx = BODY_X + name_w + col_w * ci
        rect(sl, cx, ty, col_w, hdr_h, HEADER, WHITE, 0.5)
        text(sl, cx + Inches(0.06), ty + Inches(0.05), col_w - Inches(0.12), hdr_h - Inches(0.10),
             c, size=10.5, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

    # データ行
    for ri, (name, emph, marks) in enumerate(rows):
        ry = ty + hdr_h + row_h * ri
        name_bg = RGBColor(0xec, 0xf2, 0xf8) if emph else (SOFT if ri % 2 == 0 else WHITE)
        rect(sl, BODY_X, ry, name_w, row_h, name_bg, LINE, 0.3)
        if emph:
            rect(sl, BODY_X, ry, Inches(0.06), row_h, D_GRN)
        text(sl, BODY_X + Inches(0.20), ry, name_w - Inches(0.30), row_h,
             name, size=12, bold=emph, color=(INK if emph else INK2),
             anchor=MSO_ANCHOR.MIDDLE)
        for ci, (sym, col) in enumerate(marks):
            cx = BODY_X + name_w + col_w * ci
            cell_bg = name_bg
            rect(sl, cx, ry, col_w, row_h, cell_bg, LINE, 0.3)
            text(sl, cx, ry, col_w, row_h, sym, size=18, bold=True,
                 color=col, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

    # 凡例
    leg_y = ty + table_h + Inches(0.16)
    legend = [("◎", "優れる", D_GRN), ("○", "可能", D_GRN), ("△", "限定的", D_AMBR), ("×", "不可・課題", D_RED)]
    lx = BODY_X
    for sym, lbl, col in legend:
        text(sl, lx, leg_y, Inches(0.4), Inches(0.30), sym, size=14, bold=True,
             color=col, anchor=MSO_ANCHOR.MIDDLE)
        text(sl, lx + Inches(0.38), leg_y, Inches(1.5), Inches(0.30), lbl, size=11,
             color=INK3, anchor=MSO_ANCHOR.MIDDLE)
        lx += Inches(2.0)

    # 出典注記
    text(sl, BODY_X, leg_y + Inches(0.40), BODY_W, Inches(0.50),
         "※ 水・AquaGel-K との火炎保護・発泡の比較は本論文の実測（Fig. 3・4）に基づく。"
         "Phos-Chek 列は燃焼の直接測定ではなく、文献に基づく環境残留性・適用性の定性的評価",
         size=9.5, italic=True, color=MUTED)


# ===== SLIDE 35: 結論 =====
def slide_35_summary(prs):
    sl = new_slide(prs)
    draw_header(sl, "まとめ 2/2", "研究の結論：核心メカニズムと4つの成果")
    draw_footer(sl, "42 / 43")

    # ── 上段：核心メカニズムを3ステップの横並びフロー図で視覚化 ──
    text(sl, BODY_X, BODY_Y, BODY_W, Inches(0.30),
         "核心メカニズム：火炎接触でゲルが「自己変態」する", size=13, bold=True, color=INK)
    fy = BODY_Y + Inches(0.42)
    fh = Inches(1.65)
    steps = [
        ("STEP 1", "水が急速に蒸発", "火炎接触でゲル中の水が気化・吸熱\n→ 基材表面の温度上昇を抑制", D_TEAL),
        ("STEP 2", "CSP が焼結", "シリカ粒子が加熱で焼結し\n粒子間にネック（架橋）を形成", D_AMBR),
        ("STEP 3", "エアロゲル断熱層", "多孔質シリカエアロゲルが in situ で完成\n→ 乾燥後も基材を断熱保護", ACCENT),
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
        # 矢印
        if i < n - 1:
            ax = sx + sw
            text(sl, ax, fy, arrow_w, fh, "▶", size=18, bold=True,
                 color=GRAY_L, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

    # ── 中段：4つの成果（1行＝アイコン＋見出し＋一言）──
    ry = fy + fh + Inches(0.28)
    text(sl, BODY_X, ry, BODY_W, Inches(0.30),
         "4つの主要成果", size=13, bold=True, color=INK)
    ry += Inches(0.40)
    results = [
        ("3–6×", "市販品比の保護時間", D_GRN),
        ("自己形成", "エアロゲル断熱層を in situ 生成", ACCENT),
        ("噴霧可能", "既存散布インフラと互換", D_TEAL),
        ("持続可能", "セルロース系・安全な原料", D_AMBR),
    ]
    rcw = (BODY_W - Inches(0.42)) / 4
    rch = Inches(1.30)
    for i, (big, lbl, c) in enumerate(results):
        cx = BODY_X + i * (rcw + Inches(0.14))
        rrect(sl, cx, ry, rcw, rch, SOFT, LINE, 0.5, radius=0.04)
        rect(sl, cx, ry, Inches(0.06), rch, c)
        text(sl, cx + Inches(0.20), ry + Inches(0.16), rcw - Inches(0.34), Inches(0.50),
             big, size=22, bold=True, color=c)
        text(sl, cx + Inches(0.20), ry + Inches(0.74), rcw - Inches(0.34), rch - Inches(0.84),
             lbl, size=11.5, color=INK3)

    # ── 下段：キーメッセージ ──
    cy_co = ry + rch + Inches(0.22)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.62),
            "「水のキャリア」から「加熱で自己変態する難燃材料」へ ― 乾燥後も効力を保つ次世代 WEG",
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
         "市販WEG（AquaGel-K）比 Time to char 向上率", size=22, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER)

    text(sl, Inches(1.5), Inches(4.25), Inches(10.3), Inches(1.4),
         "「水のキャリア」から「自己変態する難燃材料」へ\nWEGの新しい設計パラダイムは山火事から重要インフラを守る実用技術への道を開く",
         size=15, color=RGBColor(0xa8, 0xbc, 0xd4), align=PP_ALIGN.CENTER)

    bars = [
        ("Time to char", 0.95, ">7 min", GOLD),
        ("Foaming Index", 0.68, "≈2.6×", ACCENT),
        ("Shear-thinning", 0.60, "実証", GRAY_L),
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


# ===== STANDALONE: MHEC の優位性（経時安定性＋コスト） =====
def slide_mhec_advantage(prs):
    """
    独立スライド：MHEC が HEC+MC に対して優れる2点
    ① 経時変化が穏やか（SI Fig. S7 データ）
    ② 単一ポリマーによるコスト・製造簡便性
    main() から任意の位置で呼び出して挿入する。
    """
    sl = new_slide(prs)
    draw_header(sl, "補足", "MHEC 系の優位性：経時安定性とコスト")
    draw_footer(sl, "37 / 43")   # 挿入位置に応じてページ番号を変更する

    cw = (BODY_W - Inches(0.28)) / 2
    cx2 = BODY_X + cw + Inches(0.28)

    # ===== 左列：経時安定性（SI Section 2-3 / Fig. S7） =====
    rect(sl, BODY_X, BODY_Y, Inches(0.06), Inches(2.80), D_TEAL)
    text(sl, BODY_X + Inches(0.20), BODY_Y, cw - Inches(0.20), Inches(0.36),
         "① 経時変化が穏やか（SI Fig. S7）", size=14, bold=True, color=INK)
    text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(0.42), cw - Inches(0.20), Inches(0.30),
         "Day 1 → Day 455 の降伏応力変化", size=11, bold=True, color=MUTED)

    # Bar chart: MHEC vs HEC+MC の変化率
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
        rrect(sl, bx, by + Inches(0.30), bw, Inches(0.22),
              LINE, None, radius=0.4)
        rrect(sl, bx, by + Inches(0.30), bw * ratio, Inches(0.22),
              c, None, radius=0.4)
        text(sl, bx + bw * ratio + Inches(0.12), by + Inches(0.22),
             Inches(1.4), Inches(0.40), val, size=10, bold=True, color=c)
        by += Inches(0.90)

    # 解釈テキスト
    rrect(sl, BODY_X + Inches(0.10), by + Inches(0.10), cw - Inches(0.20), Inches(0.90),
          SOFT, D_TEAL, 0.8, radius=0.04)
    text(sl, BODY_X + Inches(0.24), by + Inches(0.22), cw - Inches(0.46), Inches(0.72),
         "MHEC系は非共有結合が主体のまま経時変化\n→ 製品特性が1年以上にわたって安定\n（HEC+MC系は共有結合化が進み剛性が変動）",
         size=11, color=INK2)

    # SI Fig. S7 参照
    text(sl, BODY_X + Inches(0.20), by + Inches(1.18), cw - Inches(0.40), Inches(0.22),
         "出典：SI Fig. S7（フロースイープ・HB モデル fit）", size=9, italic=True, color=MUTED)

    # 区切り線
    rect(sl, BODY_X, BODY_Y + Inches(2.94), cw, Inches(0.015), LINE)

    # 小見出し（Si-O 縮合の不在）
    text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(3.06), cw - Inches(0.20), Inches(0.26),
         "なぜ安定か：Si–O 共有結合縮合が起こりにくい", size=11, bold=True, color=D_TEAL)
    text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(3.38), cw - Inches(0.20), Inches(0.60),
         "HEC+MC/CSP では熟成とともに CSP と高分子の間で\n"
         "Si–O 縮合が進行し剛性が上昇する。MHEC/CSP では\n"
         "この縮合が抑制され、特性が緩やかに変化する。",
         size=10.5, color=INK2)

    # ===== 右列：コスト・製造簡便性 =====
    rect(sl, cx2, BODY_Y, Inches(0.06), Inches(4.20), D_AMBR)
    text(sl, cx2 + Inches(0.20), BODY_Y, cw - Inches(0.20), Inches(0.36),
         "② 製造コスト・品質管理で有利", size=14, bold=True, color=INK)

    cost_items = [
        ("単一ポリマー",
         "HEC と MC を別途調達・混合する必要がない\n→ 原料種数を削減"),
        ("配合工程の簡略化",
         "2成分の混合比調整・均一分散の管理が不要\n→ 製造バッチ間のばらつきを低減"),
        ("品質管理の負担軽減",
         "1種類の原料規格を管理するだけで良い\n→ スケールアップ時のコスト優位性が顕在化"),
        ("廃棄・物流コスト",
         "保管・輸送する原料が1種類に統合\n→ サプライチェーンを簡素化"),
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

    # 下部：限界の注記
    cy_co = H - FTR_H - Inches(0.16) - Inches(0.80)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.80),
            "注：コスト比較の定量データは本論文には掲載されていない。MHEC の優位性は「単一ポリマー」であることと"
            "「経時安定性」の実測（SI Fig. S7）に基づく定性的評価である。",
            icon='ⓘ')


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
                          "3 / 43")
    slide_04_wildfire(prs)
    slide_05_process(prs)
    slide_06_existing(prs)
    slide_07_predecessor(prs)
    slide_07_core(prs)
    slide_section_divider(prs, "02", "Part 2", "材料・実験方法",
                          "セルロース系ポリマーとコロイダルシリカの組み合わせによる新規ゲル設計",
                          ["ポリマー成分", "シリカ粒子・界面活性剤", "配合系（5種）", "評価手法"],
                          "9 / 43")
    slide_mat_composition(prs)
    slide_self_protection(prs)
    slide_09_polymers(prs)
    slide_10_mc(prs)
    slide_11_csp_sds(prs)
    slide_12_formulations(prs)
    slide_13_methods(prs)
    slide_14_adhesion(prs)
    slide_section_divider(prs, "03", "Part 3", "実験結果",
                          "レオロジー特性、燃焼性能、発泡構造、エアロゲル形成メカニズムまで",
                          ["レオロジー", "燃焼試験", "発泡指数", "SEM観察"],
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
    slide_section_divider(prs, "04", "Part 4", "考察・まとめ",
                          "実装シナリオ、既存品との総合比較、研究の意義と今後の展望",
                          ["実装シナリオ", "性能比較", "まとめ"],
                          "39 / 43")
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
