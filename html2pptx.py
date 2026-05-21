#!/usr/bin/env python3
"""slides.html → slides.pptx 変換スクリプト"""

from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import re

# ─── 色定義 ──────────────────────────────────────────
NAVY    = RGBColor(0x06, 0x15, 0x29)   # ダークネイビー（サイドバー）
BLUE    = RGBColor(0x0c, 0x2d, 0x58)   # サイドバーグラデーション下
ACCENT  = RGBColor(0xf0, 0xa5, 0x00)   # アクセントオレンジ
WHITE   = RGBColor(0xff, 0xff, 0xff)
LTBLUE  = RGBColor(0x1a, 0x6b, 0xcc)   # カードボーダー
BGCARD  = RGBColor(0xf3, 0xf7, 0xfd)   # カード背景
GRAY    = RGBColor(0x33, 0x33, 0x44)   # 本文グレー
PLACEHOLDER_BG = RGBColor(0xf0, 0xf4, 0xfa)
PLACEHOLDER_BD = RGBColor(0xb0, 0xc4, 0xde)

# ─── スライドサイズ 16:9 ─────────────────────────────
W = Inches(13.333)  # 1280px相当 (96dpi換算)
H = Inches(7.5)     # 720px相当

SIDEBAR_W = Inches(1.97)  # 190px
TOPBAR_H  = Inches(0.44)  # ~42px
BOTBAR_H  = Inches(0.27)  # ~26px

def emu(px, dpi=96):
    """ピクセル→EMU"""
    return int(px / dpi * 914400)

def add_rect(slide, x, y, w, h, fill_color, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(1, x, y, w, h)  # MSO_SHAPE_TYPE.RECTANGLE
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape

def add_textbox(slide, x, y, w, h, text, font_size=12, bold=False,
                color=WHITE, align=PP_ALIGN.LEFT, wrap=True):
    txBox = slide.shapes.add_textbox(x, y, w, h)
    txBox.word_wrap = wrap
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox

def clean_text(element):
    """HTML要素からプレーンテキストを抽出"""
    if element is None:
        return ""
    return element.get_text(separator="\n", strip=True)

def extract_bullets(element):
    """li要素からテキストリストを抽出"""
    if element is None:
        return []
    items = element.find_all('li')
    return [li.get_text(strip=True) for li in items]

def extract_cards(element):
    """card要素からタイトル+テキストを抽出"""
    cards = []
    for card in element.find_all(class_='card'):
        h = card.find(['h3','h4','strong'])
        title = h.get_text(strip=True) if h else ""
        texts = []
        for child in card.children:
            if hasattr(child, 'name') and child.name in ('p','ul','li'):
                t = child.get_text(strip=True)
                if t:
                    texts.append(t)
            elif hasattr(child, 'name') and child.name == 'ul':
                for li in child.find_all('li'):
                    texts.append("• " + li.get_text(strip=True))
        if not texts:
            all_text = card.get_text(separator='\n', strip=True)
            lines = [l for l in all_text.split('\n') if l.strip()]
            texts = lines[1:] if len(lines) > 1 else lines
        cards.append((title, "\n".join(texts[:4])))
    return cards

def make_cover_slide(prs, slide_html):
    """表紙スライド"""
    layout = prs.slide_layouts[6]  # 空白レイアウト
    slide = prs.slides.add_slide(layout)
    slide.shapes.title  # noqa

    # 背景全体をネイビーに
    add_rect(slide, 0, 0, W, H, NAVY)

    # 左アクセントバー
    add_rect(slide, 0, 0, Inches(0.15), H, ACCENT)

    # 中央コンテンツ取得
    ct = slide_html.find(class_='ct') or slide_html
    h1 = slide_html.find('h1')
    if not h1:
        h1 = slide_html.find(class_='cover-title')

    # タイトル
    title_text = ""
    if h1:
        title_text = h1.get_text(strip=True)
    else:
        tb = slide_html.find(class_='tb')
        if tb:
            h2 = tb.find('h2')
            title_text = h2.get_text(strip=True) if h2 else "WEG Wildfire Protection"

    if not title_text:
        title_text = "WEG Wildfire Protection Gel"

    # メインタイトル
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(12.5), Inches(2.5))
    txBox.word_wrap = True
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = title_text
    run.font.size = Pt(36)
    run.font.bold = True
    run.font.color.rgb = WHITE

    # サブタイトル・著者情報
    sub_texts = []
    for tag in (ct or slide_html).find_all(['p', 'div'], limit=5):
        t = tag.get_text(strip=True)
        if t and t != title_text and len(t) < 200:
            sub_texts.append(t)
            if len(sub_texts) >= 3:
                break

    if sub_texts:
        txBox2 = slide.shapes.add_textbox(Inches(0.5), Inches(4.5), Inches(12.5), Inches(2.0))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        for i, st in enumerate(sub_texts[:3]):
            p2 = tf2.paragraphs[0] if i == 0 else tf2.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            run2 = p2.add_run()
            run2.text = st
            run2.font.size = Pt(14)
            run2.font.color.rgb = RGBColor(0xc0, 0xd4, 0xf0)

    # 下部アクセントライン
    add_rect(slide, 0, H - Inches(0.08), W, Inches(0.08), ACCENT)

    return slide

def make_content_slide(prs, slide_html):
    """コンテンツスライド（サイドバー+本文）"""
    layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(layout)

    # ─ 背景白 ─
    add_rect(slide, 0, 0, W, H, WHITE)

    # ─ トップバー（ネイビー） ─
    add_rect(slide, 0, 0, W, TOPBAR_H, NAVY)

    # トップバーテキスト
    tb = slide_html.find(class_='tb')
    section_text = ""
    title_text = ""
    if tb:
        sec = tb.find(class_='tb-sec')
        h2 = tb.find('h2')
        section_text = sec.get_text(strip=True) if sec else ""
        title_text = h2.get_text(strip=True) if h2 else ""

    if section_text:
        add_textbox(slide, Inches(0.2), Inches(0.05), Inches(2.0), TOPBAR_H,
                    section_text, font_size=8, bold=True, color=ACCENT)

    if title_text:
        add_textbox(slide, Inches(2.3), Inches(0.05), Inches(10.8), TOPBAR_H,
                    title_text, font_size=14, bold=True, color=WHITE)

    # ─ サイドバー ─
    sb_y = TOPBAR_H
    sb_h = H - TOPBAR_H - BOTBAR_H
    add_rect(slide, 0, sb_y, SIDEBAR_W, sb_h, NAVY)

    sb_el = slide_html.find(class_='sb')
    if sb_el:
        sb_title_el = sb_el.find(class_='sb-title')
        if sb_title_el:
            add_textbox(slide, Inches(0.15), sb_y + Inches(0.25), SIDEBAR_W - Inches(0.2),
                        Inches(0.3), sb_title_el.get_text(strip=True),
                        font_size=7, bold=True, color=ACCENT)

        items = sb_el.find_all(class_='sb-item')
        for j, item in enumerate(items):
            is_on = 'on' in item.get('class', [])
            y_pos = sb_y + Inches(0.65) + j * Inches(0.32)
            if is_on:
                add_rect(slide, Inches(0.08), y_pos - Inches(0.03),
                         SIDEBAR_W - Inches(0.16), Inches(0.3),
                         RGBColor(0x2a, 0x1e, 0x00))
            add_textbox(slide, Inches(0.18), y_pos,
                        SIDEBAR_W - Inches(0.25), Inches(0.28),
                        item.get_text(strip=True),
                        font_size=9, bold=is_on,
                        color=ACCENT if is_on else RGBColor(0xaa, 0xbb, 0xcc))

    # ─ コンテンツエリア ─
    ct_x = SIDEBAR_W + Inches(0.1)
    ct_y = TOPBAR_H + Inches(0.2)
    ct_w = W - SIDEBAR_W - Inches(0.4)
    ct_h = H - TOPBAR_H - BOTBAR_H - Inches(0.3)

    ct = slide_html.find(class_='ct')
    if ct:
        render_content(slide, ct, ct_x, ct_y, ct_w, ct_h)

    # ─ ボトムバー ─
    bb_y = H - BOTBAR_H
    add_rect(slide, 0, bb_y, W, BOTBAR_H, NAVY)

    bb = slide_html.find(class_='bb')
    if bb:
        spans = bb.find_all('span')
        if spans:
            add_textbox(slide, Inches(0.3), bb_y + Inches(0.03), Inches(6), BOTBAR_H,
                        spans[0].get_text(strip=True), font_size=7,
                        color=RGBColor(0x88, 0x99, 0xaa))
            if len(spans) > 1:
                add_textbox(slide, Inches(7), bb_y + Inches(0.03), Inches(6), BOTBAR_H,
                            spans[-1].get_text(strip=True), font_size=7,
                            color=RGBColor(0x88, 0x99, 0xaa), align=PP_ALIGN.RIGHT)

    # スライド番号
    sn = slide_html.find(class_='sn')
    if sn:
        add_textbox(slide, W - Inches(0.8), H - Inches(0.25), Inches(0.7), Inches(0.22),
                    sn.get_text(strip=True), font_size=7,
                    color=RGBColor(0x88, 0x88, 0x99), align=PP_ALIGN.RIGHT)

    return slide

def render_content(slide, ct, x, y, w, h):
    """コンテンツエリアを描画"""
    cur_y = y

    # 図プレースホルダー
    figs = ct.find_all(class_='fig')
    cards = ct.find_all(class_='card')
    stats = ct.find_all(class_='stat')
    tables = ct.find_all('table')
    uls = ct.find_all('ul', recursive=False)
    paras = ct.find_all('p', recursive=False)

    # figがある場合はプレースホルダーボックス追加
    if figs:
        fig_h = min(Inches(2.5), h * 0.45)
        fig_w_each = (w - Inches(0.15) * (len(figs) - 1)) / len(figs)
        for fi, fig in enumerate(figs):
            fx = x + fi * (fig_w_each + Inches(0.15))
            fig_shape = add_rect(slide, fx, cur_y, fig_w_each, fig_h,
                                 PLACEHOLDER_BG, PLACEHOLDER_BD, Pt(1.5))
            label_el = fig.find(class_='fig-label')
            note_el = fig.find(class_='fig-note')
            label = label_el.get_text(strip=True) if label_el else fig.get_text(strip=True)[:40]
            note = note_el.get_text(strip=True) if note_el else ""

            add_textbox(slide, fx + Inches(0.1), cur_y + fig_h * 0.3,
                        fig_w_each - Inches(0.2), fig_h * 0.4,
                        label, font_size=9, bold=True,
                        color=RGBColor(0x44, 0x66, 0x88), align=PP_ALIGN.CENTER)
            if note:
                add_textbox(slide, fx + Inches(0.1), cur_y + fig_h * 0.55,
                            fig_w_each - Inches(0.2), fig_h * 0.35,
                            note, font_size=8,
                            color=RGBColor(0x77, 0x88, 0x99), align=PP_ALIGN.CENTER)
        cur_y += fig_h + Inches(0.15)

    # カード
    if cards:
        card_w = (w - Inches(0.15) * (min(len(cards), 3) - 1)) / min(len(cards), 3)
        card_h = min(Inches(2.2), (h - (cur_y - y)) * 0.9)
        for ci, card in enumerate(cards[:3]):
            cx = x + ci * (card_w + Inches(0.15))
            add_rect(slide, cx, cur_y, card_w, card_h, BGCARD)
            # カードトップボーダー
            add_rect(slide, cx, cur_y, card_w, Inches(0.04), LTBLUE)

            title_el = card.find(['h3', 'h4', 'strong', 'b'])
            card_title = title_el.get_text(strip=True) if title_el else ""
            all_text = card.get_text(separator='\n', strip=True)
            lines = [l for l in all_text.split('\n') if l.strip()]

            ty = cur_y + Inches(0.08)
            if card_title:
                add_textbox(slide, cx + Inches(0.12), ty, card_w - Inches(0.2), Inches(0.35),
                            card_title, font_size=10, bold=True, color=LTBLUE)
                ty += Inches(0.35)
                body_lines = [l for l in lines if l != card_title]
            else:
                body_lines = lines

            body = "\n".join(body_lines[:6])
            if body:
                add_textbox(slide, cx + Inches(0.12), ty, card_w - Inches(0.2),
                            card_h - (ty - cur_y) - Inches(0.1),
                            body, font_size=9, color=GRAY)
        cur_y += card_h + Inches(0.15)

    # stat ボックス
    if stats:
        stat_w = (w - Inches(0.1) * (len(stats) - 1)) / len(stats)
        stat_h = Inches(1.2)
        for si, stat in enumerate(stats[:5]):
            sx = x + si * (stat_w + Inches(0.1))
            add_rect(slide, sx, cur_y, stat_w, stat_h, RGBColor(0xe8, 0xf0, 0xfc))
            val_el = stat.find(class_='stat-val') or stat.find(['strong', 'b'])
            lbl_el = stat.find(class_='stat-lbl') or stat.find('p')
            val = val_el.get_text(strip=True) if val_el else stat.get_text(strip=True)[:10]
            lbl = lbl_el.get_text(strip=True) if lbl_el else ""
            add_textbox(slide, sx + Inches(0.05), cur_y + Inches(0.1),
                        stat_w - Inches(0.1), Inches(0.6),
                        val, font_size=18, bold=True, color=LTBLUE, align=PP_ALIGN.CENTER)
            if lbl:
                add_textbox(slide, sx + Inches(0.05), cur_y + Inches(0.7),
                            stat_w - Inches(0.1), Inches(0.4),
                            lbl, font_size=8, color=GRAY, align=PP_ALIGN.CENTER)
        cur_y += stat_h + Inches(0.15)

    # テーブル
    if tables:
        for tbl in tables[:1]:
            rows = tbl.find_all('tr')
            row_h = Inches(0.32)
            tbl_h = row_h * min(len(rows), 8)
            tbl_w = w
            add_rect(slide, x, cur_y, tbl_w, tbl_h, RGBColor(0xf8, 0xf9, 0xff))

            for ri, row in enumerate(rows[:8]):
                cells = row.find_all(['th', 'td'])
                if not cells:
                    continue
                cell_w = tbl_w / max(len(cells), 1)
                ry = cur_y + ri * row_h
                bg = NAVY if ri == 0 else (RGBColor(0xea, 0xf1, 0xfc) if ri % 2 == 0 else WHITE)
                add_rect(slide, x, ry, tbl_w, row_h, bg)
                for ci, cell in enumerate(cells):
                    cx = x + ci * cell_w
                    fg = WHITE if ri == 0 else GRAY
                    add_textbox(slide, cx + Inches(0.05), ry + Inches(0.04),
                                cell_w - Inches(0.1), row_h - Inches(0.06),
                                cell.get_text(strip=True), font_size=9,
                                bold=(ri == 0), color=fg)
            cur_y += tbl_h + Inches(0.15)

    # リスト（直下のul/ol）
    for ul in ct.find_all(['ul', 'ol'], recursive=False):
        items = ul.find_all('li')
        rem_h = h - (cur_y - y)
        if rem_h < Inches(0.3):
            break
        bullet_text = "\n".join(["• " + li.get_text(strip=True) for li in items[:8]])
        if bullet_text:
            add_textbox(slide, x, cur_y, w, min(rem_h, Inches(2.5)),
                        bullet_text, font_size=10, color=GRAY)
            cur_y += Inches(0.3) + Inches(0.26) * min(len(items), 8)

    # 残りの段落
    for p in ct.find_all('p', recursive=False):
        rem_h = h - (cur_y - y)
        if rem_h < Inches(0.25):
            break
        text = p.get_text(strip=True)
        if text and len(text) > 3:
            add_textbox(slide, x, cur_y, w, Inches(0.4), text, font_size=10, color=GRAY)
            cur_y += Inches(0.4)

    # グリッドレイアウト内のコンテンツ（fig以外）
    for grid in ct.find_all(class_=re.compile(r'grid|row|flex|cols')):
        rem_h = h - (cur_y - y)
        if rem_h < Inches(0.3):
            break
        text = grid.get_text(separator='\n', strip=True)
        lines = [l for l in text.split('\n') if l.strip()][:10]
        if lines:
            add_textbox(slide, x, cur_y, w, min(rem_h, Inches(2.5)),
                        "\n".join(lines), font_size=9, color=GRAY)
            cur_y += Inches(0.25) * min(len(lines), 10) + Inches(0.1)


# ─── メイン処理 ─────────────────────────────────────
def main():
    with open('/home/user/my-first-claude/slides.html', 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'lxml')

    slides_html = soup.find_all('div', class_='slide')
    print(f"変換するスライド数: {len(slides_html)}")

    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    for i, slide_html in enumerate(slides_html):
        print(f"  スライド {i+1}/{len(slides_html)} を処理中...")
        if i == 0:
            make_cover_slide(prs, slide_html)
        else:
            make_content_slide(prs, slide_html)

    out_path = '/home/user/my-first-claude/slides.pptx'
    prs.save(out_path)
    print(f"\n完了: {out_path}")

if __name__ == '__main__':
    main()
