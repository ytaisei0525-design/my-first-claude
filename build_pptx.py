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
         "1 / 37", size=10, color=RGBColor(0x66, 0x77, 0x88), align=PP_ALIGN.RIGHT)


# ===== SLIDE 2: TOC =====
def slide_02_toc(prs):
    sl = new_slide(prs)
    draw_header(sl, "Contents", "発表の構成")
    draw_footer(sl, "2 / 37")

    parts = [
        ("PART 1", "研究背景", "山火事の現状、既存技術（Phos-Chek, AquaGel-K）の限界、本研究の革新点", "Slides 4–7"),
        ("PART 2", "材料・方法", "セルロース系ポリマー、CSP、配合系（5種）、評価手法", "Slides 9–14"),
        ("PART 3", "実験結果", "レオロジー、燃焼試験、発泡指数、SEM観察、メカニズム", "Slides 16–32"),
        ("PART 4", "考察・まとめ", "実装シナリオ、既存品との比較、結論と今後の展望", "Slides 34–37"),
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
    draw_footer(sl, "4 / 37")

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
         "気候変動・干ばつが主要因", size=11, color=MUTED)

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
         "乾燥・強風期間の長期化により、\n従来の消火・防火戦略だけでは対応が困難",
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
    draw_header(sl, "背景 2/4", "本研究のアプローチ：4段階の保護プロセス")
    draw_footer(sl, "5 / 37")

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


# ===== SLIDE 6: 既存技術と限界 =====
def slide_06_existing(prs):
    sl = new_slide(prs)
    draw_header(sl, "背景 3/4", "既存の山火事対策と、その本質的な限界")
    draw_footer(sl, "6 / 37")

    # 4製品カード横一列：上部に写真プレースホルダー、下部に製品名＋限界テキスト
    n = 4
    gap = Inches(0.18)
    cw = (BODY_W - gap * (n - 1)) / n
    card_h = BODY_H - Inches(0.56)

    products = [
        ("水のみ", "H₂O", "蒸発・流失で\n即座に効果消失", False),
        ("Phos-Chek", "リン酸アンモニウム", "土壌・水系への\n化学汚染が残留", False),
        ("AquaGel-K", "架橋ポリアクリレート", "水蒸発と同時に\n保護機能が消失", False),
        ("WEG（本研究）", "セルロース＋シリカ", "水蒸発後も\nエアロゲル層が継続保護", True),
    ]

    for i, (name, ingredient, limit_text, is_hero) in enumerate(products):
        cx = BODY_X + i * (cw + gap)
        img_h = card_h * 0.50

        # カード背景
        if is_hero:
            rrect(sl, cx, BODY_Y, cw, card_h, HEADER, None, radius=0.04)
            rect(sl, cx, BODY_Y, cw, Inches(0.04), GOLD)
        else:
            rrect(sl, cx, BODY_Y, cw, card_h, CARD_C, LINE, 0.5, radius=0.04)
            rect(sl, cx, BODY_Y, cw, Inches(0.04), LINE)

        # 写真プレースホルダー（上部）
        ph_fill = RGBColor(0x2a, 0x3c, 0x54) if is_hero else PH_BG
        ph_bord = RGBColor(0x44, 0x5e, 0x7a) if is_hero else PH_BD
        ph_tc = RGBColor(0x88, 0xa0, 0xb8) if is_hero else GRAY_L
        rrect(sl, cx + Inches(0.12), BODY_Y + Inches(0.12),
              cw - Inches(0.24), img_h - Inches(0.12), ph_fill, ph_bord, 0.7, radius=0.03)
        text(sl, cx + Inches(0.12), BODY_Y + Inches(0.12),
             cw - Inches(0.24), img_h - Inches(0.12),
             "（写真）", size=11, color=ph_tc,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # 製品名
        name_c = WHITE if is_hero else INK
        ingr_c = RGBColor(0xa0, 0xb8, 0xd0) if is_hero else MUTED
        ty = BODY_Y + img_h + Inches(0.14)
        text(sl, cx + Inches(0.10), ty, cw - Inches(0.20), Inches(0.34),
             name, size=14, bold=True, color=name_c, align=PP_ALIGN.CENTER)
        text(sl, cx + Inches(0.10), ty + Inches(0.34), cw - Inches(0.20), Inches(0.26),
             ingredient, size=10, color=ingr_c, align=PP_ALIGN.CENTER)

        # 限界テキスト
        limit_c = D_GRN if is_hero else D_RED
        text(sl, cx + Inches(0.10), ty + Inches(0.64), cw - Inches(0.20),
             card_h - (ty - BODY_Y) - Inches(0.72),
             limit_text, size=12, bold=is_hero,
             color=limit_c, align=PP_ALIGN.CENTER)

    # 下部メッセージ（横線区切り）
    msg_y = BODY_Y + card_h + Inches(0.14)
    rect(sl, BODY_X, msg_y, BODY_W, Inches(0.02), LINE)
    text(sl, BODY_X, msg_y + Inches(0.12), BODY_W, Inches(0.36),
         "課題：既存WEGは「水のキャリア」止まり ─ 水分喪失と同時に保護機能が失われる",
         size=13, bold=True, color=INK, align=PP_ALIGN.CENTER)


# ===== SLIDE 7: 研究の核心 =====
def slide_07_core(prs):
    sl = new_slide(prs)
    draw_header(sl, "背景 4/4", "本研究の核心：熱活性化エアロゲル形成")
    draw_footer(sl, "7 / 37")

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

    # 3つの革新ポイント（番号付きリスト、カードなし）
    points = [
        ("01", "水分依存からの脱却", "熱で活性化する固体保護層を形成"),
        ("02", "持続可能な素材", "天然由来・食品添加物グレード素材"),
        ("03", "既存インフラと互換", "噴霧可能な流体として散布可能"),
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


# ===== SLIDE 9: セルロース系ポリマー =====
def slide_09_polymers(prs):
    sl = new_slide(prs)
    draw_header(sl, "材料 1/6", "使用したセルロース系ポリマー")
    draw_footer(sl, "9 / 37")

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


# ===== SLIDE 10: MC熱ゲル化 =====
def slide_10_mc(prs):
    sl = new_slide(prs)
    draw_header(sl, "材料 2/6", "メチルセルロース(MC)の熱ゲル化：火炎保護の鍵となる特性")
    draw_footer(sl, "10 / 37")

    left_w = BODY_W * 0.48 - Inches(0.10)
    right_w = BODY_W * 0.52 - Inches(0.10)
    right_x = BODY_X + left_w + Inches(0.20)

    # 左上：ゾル→ゲル外観写真プレースホルダー
    img_h = Inches(2.7)
    fig_placeholder(sl, BODY_X, BODY_Y, left_w, img_h,
                    fig_num="Photo",
                    caption="MCゾル（室温）→ ゲル（加熱後）の外観比較")

    # 左下：主要特性（横線区切り、カードなし）
    ty = BODY_Y + img_h + Inches(0.20)
    rect(sl, BODY_X, ty, left_w * 0.9, Inches(0.015), LINE)
    ty += Inches(0.18)
    props = [
        ("LCST", "~55°C でゲル化開始（1 wt% 水溶液）"),
        ("熱分解", "~300°C で有機成分が分解"),
        ("設計意図", "常温では液体、火炎接触で即ゲル → CSPを固定"),
    ]
    for label, detail in props:
        text(sl, BODY_X, ty, Inches(0.88), Inches(0.30),
             label, size=10, bold=True, color=ACCENT)
        text(sl, BODY_X + Inches(0.92), ty, left_w - Inches(0.92), Inches(0.30),
             detail, size=12, color=INK2)
        ty += Inches(0.40)

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
    draw_footer(sl, "11 / 37")

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
    draw_footer(sl, "12 / 37")

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
    draw_footer(sl, "13 / 37")

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
    draw_header(sl, "材料 6/6", "付着性・表面濡れ性：実装に不可欠な特性")
    draw_footer(sl, "14 / 37")

    # 3列：各列に写真プレースホルダー（上）＋テキスト（下）
    cw = (BODY_W - Inches(0.28)) / 3
    img_h = Inches(2.4)
    photos = [
        ("接触角測定", "WETTING", "木材上での接触角", [
            "木材・コンクリート・金属で評価",
            "WEGは低接触角 → 高い濡れ性",
            "HEC+MCは木材との親和性が高い",
        ]),
        ("垂直面付着試験", "ADHESION", "垂直基板でのゲル保持", [
            "垂直基板に塗布し保持性を評価",
            "G' ≫ G'' → 流れ落ちない",
            "AquaGel-K同等以上の保持性",
        ]),
        ("スプレー散布適性", "SPRAY", "ノズル噴霧の様子", [
            "高せん断で粘度が急低下",
            "到達後すぐ高粘度を回復",
            "既存ホース・ノズルと互換",
        ]),
    ]
    for i, (title, tag, ph_cap, bullets) in enumerate(photos):
        cx = BODY_X + i * (cw + Inches(0.14))
        fig_placeholder(sl, cx, BODY_Y, cw, img_h, fig_num="Photo", caption=ph_cap)
        ty = BODY_Y + img_h + Inches(0.38)
        text(sl, cx, ty, cw, Inches(0.26), tag, size=10, bold=True, color=MUTED)
        text(sl, cx, ty + Inches(0.26), cw, Inches(0.34), title, size=14, bold=True, color=INK)
        by = ty + Inches(0.68)
        for b in bullets:
            text(sl, cx, by, Inches(0.20), Inches(0.30), "•", size=11, color=ACCENT)
            text(sl, cx + Inches(0.22), by, cw - Inches(0.22), Inches(0.30), b, size=11, color=INK2)
            by += Inches(0.34)

    # 下部 callout
    cy_co = H - FTR_H - Inches(0.18) - Inches(0.72)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.72),
            "流動学的設計（低 n、高 G'）により、散布適性と付着保持を同時に達成",
            icon='✓')


# ===== SLIDE 16: G'/G'' =====
def slide_16_rheology1(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 1/11", "レオロジー①：振動弾性率 G' / G''")
    draw_footer(sl, "16 / 37")

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
             "AquaGel-K（G'=457 Pa）より低い G'（46/26 Pa）だが全系でゲル挙動",
         ], ct_size=15, li_size=12)
    callout(sl, right_x, BODY_Y + ch + Inches(0.25), right_w, Inches(0.80),
            "ゲル骨格が明確に形成されており、塗布後に流れ落ちない",
            icon='"')


# ===== SLIDE 17: shear thinning =====
def slide_17_rheology2(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 2/11", "レオロジー②：剪断希薄化と Power-law 指数")
    draw_footer(sl, "17 / 37")

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
         "流動指数 n", "n < 1", "HEC+MC/CSP — 強い剪断希薄化", value_color=ACCENT)
    mini(sl, right_x, BODY_Y + mh + Inches(0.15), right_w, mh,
         "流動指数 n", "n < 1", "MHEC/CSP — 同様に剪断希薄化", value_color=ACCENT)

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
    draw_header(sl, "結果 3/11", "レオロジー③：静的降伏応力と長期安定性")
    draw_footer(sl, "18 / 37")

    # 上部 callout
    callout(sl, BODY_X, BODY_Y, BODY_W, Inches(0.65),
            "静的降伏応力：振幅掃引のG'/G'' クロスオーバー点（σ_s）。流れ始めに必要な最小応力",
            icon='ƒ')

    ty = BODY_Y + Inches(0.85)
    th = Inches(2.2)
    headers = ["配合系", "静的降伏応力（新鮮）", "静的降伏応力（455日後）", "備考"]
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
        ("静的降伏応力", "振幅掃引のG'/G''クロスオーバー点で測定。HEC+MC/CSPは33.34 Pa、MHEC/CSPは3.31 Pa（新鮮配合）"),
        ("長期安定性", "MHEC/CSPは455日経時後も3.31→4.56 Paとわずかな変化のみ。実用配合の長期安定性を実証"),
        ("HEC+MC vs MHEC", "HEC+MC系はより高い降伏応力で垂直面付着に有利。MHEC系は単一ポリマーで製造が簡便"),
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
    draw_header(sl, "結果 4/11", "レオロジー④：貯蔵弾性率 G' と粘弾性（tan δ）")
    draw_footer(sl, "19 / 37")

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


# ===== SLIDE 20: HEC+MC vs MHEC =====
def slide_20_compare(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 5/11", "HEC+MC vs MHEC：二系統の比較")
    draw_footer(sl, "20 / 37")

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
        "強い剪断希薄化（n < 1）",
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
        "剪断希薄化（n < 1）",
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


# ===== SLIDE 21: 燃焼試験 setup =====
def slide_21_setup(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 6/11", "燃焼試験のセットアップ")
    draw_footer(sl, "21 / 37")

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
    draw_header(sl, "結果 7/11", "Time to char：5配合系の定量比較")
    draw_footer(sl, "22 / 37")

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
    draw_header(sl, "結果 8/11", "燃焼過程の時系列観察")
    draw_footer(sl, "23 / 37")

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
    draw_header(sl, "結果 8/11", "120秒 / 300秒時点の表面状態比較")
    draw_footer(sl, "24 / 37")

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
         "25 / 37", size=10, color=RGBColor(0x55, 0x66, 0x77), align=PP_ALIGN.RIGHT)


# ===== SLIDE 26: 発泡指数 =====
def slide_26_foam(prs):
    sl = new_slide(prs)
    draw_header(sl, "結果 9/11", "発泡指数（Foaming Index）の比較")
    draw_footer(sl, "26 / 37")

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
        ("AquaGel-K", 0.02, "≈0", GRAY_L),
        ("MHEC/CSP 1-5", 0.53, "≈2.1×", D_AMBR),
        ("HEC+MC/CSP/SDS 1-5-0.1", 0.60, "≈2.3×", D_BLUE),
        ("HEC+MC/CSP 1-5", 0.68, "≈2.6×", D_TEAL),
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
        "SDS無しのHEC+MC/CSPが最高発泡指数（≈2.6）を達成",
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
    draw_footer(sl, "27 / 37")

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
    draw_header(sl, "結果 10/11", "FT-IR・XPS：化学組成の変化（燃焼前後）")
    draw_footer(sl, "28 / 37")

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
    draw_header(sl, "結果 11/11", "TGA/DSC：熱分解プロファイルと各成分の役割")
    draw_footer(sl, "29 / 37")

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
    draw_footer(sl, "30 / 37")

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
    draw_footer(sl, "31 / 37")

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
    draw_footer(sl, "32 / 37")

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

    # 右下：熟成レオロジー図プレースホルダー
    fy = BODY_Y + mh + Inches(0.28)
    fh = Inches(2.3)
    fig_placeholder(sl, right_x, fy, right_w, fh,
                    fig_num="Fig. S5/S7",
                    caption="熟成に伴う弾性率・粘度の変化（Day 1 vs Day 455）")

    # 下部：ヒーローcallout
    cy_co = H - FTR_H - Inches(0.16) - Inches(0.82)
    callout(sl, BODY_X, cy_co, BODY_W, Inches(0.82),
            "Day 455（約15か月）経過後も燃焼挙動とエアロゲル化は不変 — 実用的な保存安定性を実証（Video S5）",
            dark=True, icon='✓')


# ===== SLIDE 33: 実装シナリオ =====
def slide_33_scenarios(prs):
    sl = new_slide(prs)
    draw_header(sl, "考察 1/3", "実装シナリオ：WEGの展開戦略")
    draw_footer(sl, "34 / 37")

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
        ("vs 水", "保護時間大幅向上（水は~0.3分で炭化）"),
        ("vs AquaGel-K", "3〜6倍の保護時間＋水蒸発後も継続"),
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
    draw_footer(sl, "35 / 37")

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
    draw_footer(sl, "36 / 37")

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
         "37 / 37", size=10, color=RGBColor(0x55, 0x66, 0x77), align=PP_ALIGN.RIGHT)


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
                          "3 / 37")
    slide_04_wildfire(prs)
    slide_05_process(prs)
    slide_06_existing(prs)
    slide_07_core(prs)
    slide_section_divider(prs, "02", "Part 2", "材料・実験方法",
                          "セルロース系ポリマーとコロイダルシリカの組み合わせによる新規ゲル設計",
                          ["ポリマー成分", "シリカ粒子・界面活性剤", "配合系（5種）", "評価手法"],
                          "8 / 37")
    slide_09_polymers(prs)
    slide_10_mc(prs)
    slide_11_csp_sds(prs)
    slide_12_formulations(prs)
    slide_13_methods(prs)
    slide_14_adhesion(prs)
    slide_section_divider(prs, "03", "Part 3", "実験結果",
                          "レオロジー特性、燃焼性能、発泡構造、エアロゲル形成メカニズムまで",
                          ["レオロジー", "燃焼試験", "発泡指数", "SEM観察"],
                          "15 / 37")
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
    slide_32_aging(prs)
    slide_section_divider(prs, "04", "Part 4", "考察・まとめ",
                          "実装シナリオ、既存品との総合比較、研究の意義と今後の展望",
                          ["実装シナリオ", "性能比較", "まとめ"],
                          "33 / 37")
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
