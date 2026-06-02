#!/usr/bin/env python3
"""slides.html → slides.pptx 変換スクリプト v2"""

from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import re, os

# ── 色定義（CSS変数と対応） ──────────────────────────────────
INK    = RGBColor(0x1e, 0x29, 0x3b)
INK2   = RGBColor(0x33, 0x41, 0x55)
INK3   = RGBColor(0x47, 0x55, 0x69)
MUTED  = RGBColor(0x64, 0x74, 0x8b)
CARD   = RGBColor(0xf1, 0xf5, 0xf9)
SOFT   = RGBColor(0xf8, 0xfa, 0xfc)
LINE   = RGBColor(0xe2, 0xe8, 0xf0)
HEADER = RGBColor(0x1e, 0x29, 0x3b)
ACCENT = RGBColor(0x3b, 0x51, 0x74)
AMBER  = RGBColor(0xb0, 0x84, 0x42)
WHITE  = RGBColor(0xff, 0xff, 0xff)
D_BLUE = RGBColor(0x5b, 0x7a, 0x99)
D_TEAL = RGBColor(0x4a, 0x7c, 0x8c)
D_RED  = RGBColor(0xa0, 0x56, 0x56)
D_GRN  = RGBColor(0x5b, 0x8c, 0x5a)
D_AMBR = RGBColor(0x9f, 0x76, 0x39)

FIGS_DIR = '/home/user/my-first-claude/figs'
W = Inches(13.333)
H = Inches(7.5)
HDR_H = Inches(0.72)
FTR_H = Inches(0.08)
PAD   = Inches(0.40)
BODY_Y = HDR_H + Inches(0.22)
BODY_H = H - BODY_Y - FTR_H - Inches(0.18)
BODY_X = PAD
BODY_W = W - PAD * 2


# ── ユーティリティ ──────────────────────────────────────────
def add_rect(sl, x, y, w, h, fill, border=None, bpt=0.75):
    s = sl.shapes.add_shape(1, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if border:
        s.line.color.rgb = border
        s.line.width = Pt(bpt)
    else:
        s.line.fill.background()
    return s

def add_text(sl, x, y, w, h, text, size=11, bold=False,
             color=INK2, align=PP_ALIGN.LEFT, italic=False):
    text = (text or '').strip()
    if not text:
        return None
    tb = sl.shapes.add_textbox(x, y, w, max(h, Inches(0.18)))
    tb.word_wrap = True
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(text.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
    return tb

def add_image(sl, src, x, y, w, h):
    fname = os.path.basename(src or '')
    fpath = os.path.join(FIGS_DIR, fname)
    if os.path.exists(fpath):
        try:
            return sl.shapes.add_picture(fpath, x, y, w, h)
        except Exception as e:
            print(f"  [画像エラー] {fname}: {e}")
    # プレースホルダー
    add_rect(sl, x, y, w, h, RGBColor(0xec,0xf1,0xf8), RGBColor(0xb0,0xc4,0xde), 1)
    add_text(sl, x+Inches(0.1), y+h*0.38, w-Inches(0.2), Inches(0.36),
             fname or '(図)', size=9, color=MUTED, align=PP_ALIGN.CENTER)

def t(el):
    return el.get_text(separator=' ', strip=True) if el else ''

def classes(el):
    return el.get('class', []) if el else []

def has_cls(el, *names):
    c = classes(el)
    return any(n in c for n in names)


# ── ヘッダーを描画（全コンテンツスライド共通） ────────────────
def draw_header(sl, div):
    add_rect(sl, 0, 0, W, HDR_H, HEADER)
    add_rect(sl, 0, HDR_H - Inches(0.05), W, Inches(0.05), ACCENT)
    sh = div.find(class_='sh')
    if sh:
        tag_el = sh.find(class_='sh-tag')
        h2_el  = sh.find('h2')
        tag = t(tag_el)
        title = t(h2_el)
        if tag:
            add_text(sl, PAD, Inches(0.16), Inches(2.2), Inches(0.42),
                     tag, size=9, color=RGBColor(0x88,0xa4,0xc4))
        if title:
            add_text(sl, Inches(2.8), Inches(0.10), W - Inches(3.1), Inches(0.56),
                     title, size=22, bold=True, color=WHITE)
    # フッター
    add_rect(sl, 0, H - FTR_H, W, FTR_H, ACCENT)
    sn = div.find(class_='sn')
    if sn:
        add_text(sl, W-Inches(1.3), H-Inches(0.30), Inches(1.2), Inches(0.24),
                 t(sn), size=9, color=MUTED, align=PP_ALIGN.RIGHT)


# ── カード群を描画 ────────────────────────────────────────────
def draw_cards(sl, cards, x, y, w, h, cols=None):
    n = len(cards)
    if n == 0:
        return
    ncols = cols or min(n, 4)
    nrows = (n + ncols - 1) // ncols
    gap = Inches(0.14)
    cw = (w - gap * (ncols - 1)) / ncols
    ch = (h - gap * (nrows - 1)) / nrows

    for i, card in enumerate(cards):
        col = i % ncols
        row = i // ncols
        cx = x + col * (cw + gap)
        cy = y + row * (ch + gap)
        dark = 'dark' in classes(card)
        bg = HEADER if dark else CARD
        add_rect(sl, cx, cy, cw, ch, bg, LINE if not dark else None, 0.5)
        # アクセントトップバー
        add_rect(sl, cx, cy, cw, Inches(0.04), ACCENT)

        # タイトル
        ct_el = card.find(class_='ct')
        cb_el = card.find(class_='cb')
        li_els = card.find_all('li')

        iy = cy + Inches(0.10)
        txt_color = WHITE if dark else INK
        sub_color = RGBColor(0xcc,0xdd,0xee) if dark else INK3

        if ct_el:
            ct_text = t(ct_el)
            add_text(sl, cx+Inches(0.12), iy, cw-Inches(0.2), Inches(0.32),
                     ct_text, size=12, bold=True, color=txt_color)
            iy += Inches(0.34)

        if cb_el:
            add_text(sl, cx+Inches(0.12), iy, cw-Inches(0.2), ch-(iy-cy)-Inches(0.08),
                     t(cb_el), size=10, color=sub_color)
        elif li_els:
            bullets = '\n'.join('• ' + t(li) for li in li_els[:5])
            add_text(sl, cx+Inches(0.12), iy, cw-Inches(0.2), ch-(iy-cy)-Inches(0.08),
                     bullets, size=10, color=sub_color)
        else:
            body = card.get_text(separator='\n', strip=True)
            lines = [ln for ln in body.split('\n') if ln.strip()]
            ct_text2 = t(ct_el) if ct_el else ''
            body_lines = [ln for ln in lines if ln != ct_text2][:6]
            if body_lines:
                add_text(sl, cx+Inches(0.12), iy, cw-Inches(0.2), ch-(iy-cy)-Inches(0.08),
                         '\n'.join(body_lines), size=10, color=sub_color)


# ── stat ボックスを描画 ────────────────────────────────────────
def draw_stats(sl, stats, x, y, w, h):
    n = len(stats)
    if n == 0:
        return
    gap = Inches(0.14)
    sw = (w - gap * (n - 1)) / n
    for i, stat in enumerate(stats):
        sx = x + i * (sw + gap)
        add_rect(sl, sx, y, sw, h, CARD, LINE, 0.5)
        num_el = stat.find(class_='stat-num')
        lbl_el = stat.find(class_='stat-lbl')
        sub_el = stat.find(class_='stat-sub')
        num = t(num_el) if num_el else stat.get_text(strip=True)[:10]
        lbl = t(lbl_el)
        sub = t(sub_el)
        add_text(sl, sx+Inches(0.08), y+Inches(0.12), sw-Inches(0.16), Inches(0.55),
                 num, size=28, bold=True, color=INK, align=PP_ALIGN.CENTER)
        if lbl:
            add_text(sl, sx+Inches(0.08), y+Inches(0.68), sw-Inches(0.16), Inches(0.28),
                     lbl, size=11, color=INK2, align=PP_ALIGN.CENTER)
        if sub:
            add_text(sl, sx+Inches(0.08), y+Inches(0.96), sw-Inches(0.16), Inches(0.24),
                     sub, size=9, color=MUTED, align=PP_ALIGN.CENTER)


# ── テーブルを描画 ────────────────────────────────────────────
def draw_table(sl, tbl, x, y, w, h):
    rows = tbl.find_all('tr')
    if not rows:
        return
    n = min(len(rows), 10)
    row_h = min(h / n, Inches(0.40))
    for ri, row in enumerate(rows[:n]):
        cells = row.find_all(['th', 'td'])
        if not cells:
            continue
        cw = w / len(cells)
        ry = y + ri * row_h
        bg = HEADER if ri == 0 else (SOFT if ri % 2 == 0 else WHITE)
        add_rect(sl, x, ry, w, row_h, bg, LINE, 0.3)
        for ci, cell in enumerate(cells):
            cx2 = x + ci * cw
            fg = WHITE if ri == 0 else INK2
            add_text(sl, cx2 + Inches(0.1), ry + Inches(0.05),
                     cw - Inches(0.12), row_h - Inches(0.06),
                     t(cell), size=10, bold=(ri == 0), color=fg)


# ── callout / key を描画 ──────────────────────────────────────
def draw_callout(sl, el, x, y, w, h, dark=False):
    bg = HEADER if dark else RGBColor(0xf0,0xf4,0xf8)
    border = ACCENT
    add_rect(sl, x, y, Inches(0.05), h, border)
    add_rect(sl, x + Inches(0.05), y, w - Inches(0.05), h, bg)
    add_text(sl, x + Inches(0.2), y + Inches(0.07), w - Inches(0.28),
             h - Inches(0.1), t(el), size=11,
             color=WHITE if dark else INK)


# ── img-box から src を取得 ────────────────────────────────────
def get_img_src(container):
    img = container.find('img') if container else None
    return img.get('src', '') if img else ''


# ── .sb 内のコンテンツを解析してレンダリング ──────────────────
def render_sb(sl, sb, x, y, w, h):
    if not sb:
        return
    cur_y = y
    rem_h = lambda: h - (cur_y - y)

    # 全 img-box を収集
    img_boxes = sb.find_all(class_='img-box')
    imgs = [(ib, get_img_src(ib)) for ib in img_boxes if get_img_src(ib)]

    # グリッドコンテナを特定
    grid_re = re.compile(r'\bg(?:2|3|4|12|21|32)\b')
    grids = [c for c in sb.children
             if hasattr(c, 'get') and any(grid_re.match(cl) for cl in c.get('class', []))]

    # ── レイアウト判定 ──────────────────────────────────
    cards_all = sb.find_all(class_='card')
    stats_all = sb.find_all(class_='stat')
    tables    = sb.find_all(class_='tbl')
    callouts  = sb.find_all(class_=re.compile(r'\b(?:co|key)\b'))
    lists     = sb.find_all(class_='list')

    # 画像がある場合とない場合で分岐
    if imgs:
        # 画像1枚 + テキスト（g21など）
        if len(imgs) == 1:
            img_box, src = imgs[0]
            # 画像を左60%か右40%に配置
            # img-boxが左側にある場合: 画像左、テキスト右
            img_w = w * 0.55
            txt_w = w * 0.42
            gap = w * 0.03

            # テキストエリアのカード・リストを収集
            text_cards = [c for c in cards_all if img_box not in c.parents and c not in [img_box]]
            # Actually always put image left, text right
            img_h = min(rem_h(), Inches(4.2))
            add_image(sl, src, x, cur_y, img_w, img_h)

            # テキスト右側
            tx = x + img_w + gap
            ty = cur_y
            if text_cards:
                draw_cards(sl, text_cards, tx, ty, txt_w, img_h - Inches(0.1))
            elif lists:
                bullets = '\n'.join('• ' + t(li) for li in lists[0].find_all('li')[:8])
                add_text(sl, tx, ty, txt_w, img_h - Inches(0.1), bullets, size=11, color=INK2)
            cur_y += img_h + Inches(0.12)

        # 画像2枚並列
        elif len(imgs) == 2:
            img_h = min(rem_h() * 0.55, Inches(3.0))
            iw = (w - Inches(0.14)) / 2
            for i, (ib, src) in enumerate(imgs[:2]):
                add_image(sl, src, x + i*(iw+Inches(0.14)), cur_y, iw, img_h)
            cur_y += img_h + Inches(0.12)

        # 残りのテキスト要素
        remaining_cards = [c for c in cards_all
                           if not any(ib in c.parents or c in [ib] for ib, _ in imgs)]
        if remaining_cards and rem_h() > Inches(0.5):
            ch = rem_h() - Inches(0.05)
            draw_cards(sl, remaining_cards, x, cur_y, w, ch)
            cur_y += ch + Inches(0.1)

    else:
        # 画像なし
        # stats
        if stats_all and rem_h() > Inches(0.5):
            sh2 = min(rem_h() * 0.45, Inches(1.4))
            draw_stats(sl, stats_all, x, cur_y, w, sh2)
            cur_y += sh2 + Inches(0.14)

        # cards
        if cards_all and rem_h() > Inches(0.4):
            ch = rem_h() - (Inches(0.5) if callouts else Inches(0.05))
            n = len(cards_all)
            cols = min(n, 4)
            draw_cards(sl, cards_all, x, cur_y, w, ch, cols=cols)
            cur_y += ch + Inches(0.12)

        # table
        if tables and rem_h() > Inches(0.5):
            th = min(rem_h() * 0.75, Inches(3.5))
            draw_table(sl, tables[0], x, cur_y, w, th)
            cur_y += th + Inches(0.12)

        # lists (no cards)
        if lists and not cards_all and rem_h() > Inches(0.3):
            for lst in lists[:1]:
                bullets = '\n'.join('• ' + t(li) for li in lst.find_all('li')[:8])
                lh = min(rem_h() - Inches(0.4), Inches(3.0))
                add_text(sl, x, cur_y, w, lh, bullets, size=12, color=INK2)
                cur_y += lh + Inches(0.12)

    # callout / key （下部に描画）
    for el in callouts[:2]:
        if rem_h() < Inches(0.3):
            break
        dark = 'key' in classes(el)
        ch = Inches(0.48)
        draw_callout(sl, el, x, cur_y, w, ch, dark=dark)
        cur_y += ch + Inches(0.08)


# ── 各スライドタイプの生成 ─────────────────────────────────────
def make_cover(prs, div):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(sl, 0, 0, W, H, HEADER)
    add_rect(sl, 0, 0, Inches(0.12), H, ACCENT)
    add_rect(sl, 0, H - Inches(0.07), W, Inches(0.07), ACCENT)

    h1 = div.find('h1')
    title = t(h1) if h1 else 'Water-Enhancing Gels'
    add_text(sl, Inches(0.5), Inches(1.6), Inches(12.4), Inches(2.2),
             title, size=34, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    sub = div.find(class_='cov-sub')
    if sub:
        add_text(sl, Inches(0.8), Inches(3.9), Inches(11.8), Inches(1.2),
                 t(sub), size=16, color=RGBColor(0xb8, 0xcc, 0xe2), align=PP_ALIGN.CENTER)

    meta = div.find(class_='cov-meta')
    if meta:
        items = [t(m) for m in meta.find_all(class_='cov-mi')]
        add_text(sl, Inches(0.8), Inches(5.2), Inches(11.8), Inches(0.5),
                 '  ·  '.join(items), size=12,
                 color=RGBColor(0x77, 0x8a, 0xa0), align=PP_ALIGN.CENTER)

    sn = div.find(class_='sn')
    if sn:
        add_text(sl, W - Inches(1.3), H - Inches(0.35), Inches(1.1), Inches(0.26),
                 t(sn), size=9, color=MUTED, align=PP_ALIGN.RIGHT)
    return sl


def make_sdiv(prs, div):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(sl, 0, 0, W, H, HEADER)
    add_rect(sl, 0, 0, Inches(0.10), H, AMBER)

    lbl = div.find(class_='sdiv-lbl')
    ttl = div.find(class_='sdiv-ttl')
    sub = div.find(class_='sdiv-sub')
    pills = div.find_all(class_='sdiv-pill')

    if lbl:
        add_text(sl, Inches(1.0), Inches(1.5), Inches(10), Inches(0.4),
                 t(lbl), size=12, color=RGBColor(0x88, 0xa4, 0xc0))
    if ttl:
        add_text(sl, Inches(1.0), Inches(1.95), Inches(10), Inches(1.3),
                 t(ttl), size=46, bold=True, color=WHITE)
    if sub:
        add_text(sl, Inches(1.0), Inches(3.35), Inches(10), Inches(0.8),
                 t(sub), size=17, color=RGBColor(0x88, 0xa8, 0xc8))
    if pills:
        pill_strs = [t(p) for p in pills]
        px, py = Inches(1.0), Inches(4.4)
        for ps in pill_strs:
            pw = Inches(max(len(ps) * 0.13, 1.4))
            add_rect(sl, px, py, pw, Inches(0.38), RGBColor(0x2a,0x38,0x50),
                     RGBColor(0x44,0x58,0x76), 0.75)
            add_text(sl, px + Inches(0.15), py + Inches(0.06), pw - Inches(0.2),
                     Inches(0.28), ps, size=11, color=RGBColor(0xbb, 0xcc, 0xde))
            px += pw + Inches(0.2)

    sn = div.find(class_='sn')
    if sn:
        add_text(sl, W - Inches(1.3), H - Inches(0.35), Inches(1.1), Inches(0.26),
                 t(sn), size=9, color=RGBColor(0x55, 0x66, 0x77), align=PP_ALIGN.RIGHT)
    return sl


def make_hero(prs, div):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(sl, 0, 0, W, H, HEADER)

    eyebrow = div.find(class_='hero-eyebrow')
    num_el  = div.find(class_='hero-num')
    label   = div.find(class_='hero-label')
    sub_el  = div.find(class_='hero-sub')

    cy = Inches(0.8)
    if eyebrow:
        add_text(sl, Inches(1.5), cy, Inches(10.3), Inches(0.4),
                 t(eyebrow), size=12, color=RGBColor(0x88, 0xa4, 0xc0), align=PP_ALIGN.CENTER)
        cy += Inches(0.45)
    if num_el:
        add_text(sl, Inches(1.5), cy, Inches(10.3), Inches(1.8),
                 t(num_el), size=72, bold=True, color=RGBColor(0xd4, 0xa5, 0x74),
                 align=PP_ALIGN.CENTER)
        cy += Inches(1.85)
    if label:
        add_text(sl, Inches(1.5), cy, Inches(10.3), Inches(0.6),
                 t(label), size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        cy += Inches(0.68)
    if sub_el:
        add_text(sl, Inches(2.0), cy, Inches(9.3), Inches(0.7),
                 t(sub_el), size=14, color=RGBColor(0x88, 0xa4, 0xc0), align=PP_ALIGN.CENTER)
        cy += Inches(0.8)

    # bars
    bars = div.find_all(class_='hero-bar-row')
    if bars:
        bx, bw = Inches(2.0), Inches(9.3)
        add_rect(sl, bx, cy, bw, Inches(0.04), RGBColor(0x44, 0x55, 0x66))
        cy += Inches(0.12)
        for row in bars:
            lbl_el = row.find(class_='hero-bar-lbl')
            val_el = row.find(class_='hero-bar-val')
            lbl_str = t(lbl_el)
            val_str = t(val_el)
            add_rect(sl, bx, cy, bw, Inches(0.30), RGBColor(0x28, 0x36, 0x4c))
            if lbl_str:
                add_text(sl, bx + Inches(0.1), cy + Inches(0.05), Inches(1.8), Inches(0.22),
                         lbl_str, size=10, color=RGBColor(0xaa, 0xbc, 0xcc), align=PP_ALIGN.RIGHT)
            if val_str:
                add_text(sl, bx + bw - Inches(1.3), cy + Inches(0.05), Inches(1.2), Inches(0.22),
                         val_str, size=10, color=RGBColor(0xaa, 0xbc, 0xcc))
            cy += Inches(0.38)

    sn = div.find(class_='sn')
    if sn:
        add_text(sl, W - Inches(1.3), H - Inches(0.35), Inches(1.1), Inches(0.26),
                 t(sn), size=9, color=MUTED, align=PP_ALIGN.RIGHT)
    return sl


def make_content(prs, div):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(sl, 0, 0, W, H, WHITE)
    draw_header(sl, div)
    sb = div.find(class_='sb')
    render_sb(sl, sb, BODY_X, BODY_Y, BODY_W, BODY_H)
    return sl


# ── メイン ────────────────────────────────────────────────────
def main():
    html_path = '/home/user/my-first-claude/slides.html'
    out_path  = '/home/user/my-first-claude/slides.pptx'

    with open(html_path, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'lxml')

    slide_divs = soup.find_all('div', class_='slide')
    print(f"変換対象スライド数: {len(slide_divs)}")

    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    for i, div in enumerate(slide_divs):
        cls = div.get('class', [])
        print(f"  [{i+1:02d}/{len(slide_divs)}] ", end='')
        if 'cover' in cls:
            print("cover")
            make_cover(prs, div)
        elif 'sdiv' in cls:
            print("section divider")
            make_sdiv(prs, div)
        elif 'hero' in cls:
            print("hero")
            make_hero(prs, div)
        else:
            h2 = div.find('h2')
            print(h2.get_text(strip=True)[:40] if h2 else '(content)')
            make_content(prs, div)

    prs.save(out_path)
    print(f"\n完了: {out_path}")

if __name__ == '__main__':
    main()
