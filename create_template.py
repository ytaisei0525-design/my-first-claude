from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# カラー定義（ブルー系）
COLOR_PRIMARY = RGBColor(0x1A, 0x56, 0xAB)    # 濃いブルー
COLOR_ACCENT  = RGBColor(0x2E, 0x86, 0xC1)    # ミディアムブルー
COLOR_LIGHT   = RGBColor(0xD6, 0xE8, 0xF7)    # 薄いブルー
COLOR_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_DARK    = RGBColor(0x1C, 0x1C, 0x1C)
COLOR_GRAY    = RGBColor(0x55, 0x55, 0x55)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


def set_fill_solid(shape, color):
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_textbox(slide, left, top, width, height, text, font_size, bold=False,
                color=COLOR_DARK, align=PP_ALIGN.LEFT, italic=False):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox


def add_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    set_fill_solid(shape, color)
    shape.line.fill.background()
    return shape


# ===== プレゼンテーション作成 =====
prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

blank_layout = prs.slide_layouts[6]  # 完全に空白のレイアウト


# -----------------------------------------------
# 1. タイトルスライド
# -----------------------------------------------
slide = prs.slides.add_slide(blank_layout)

# 背景：濃いブルー
bg = add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_PRIMARY)

# 左側アクセントバー
add_rect(slide, 0, 0, Inches(0.18), SLIDE_H, COLOR_ACCENT)

# 中央区切り線
add_rect(slide, Inches(1.0), Inches(3.5), Inches(11.33), Inches(0.04), COLOR_ACCENT)

# タイトル
add_textbox(
    slide,
    Inches(1.0), Inches(2.0), Inches(11.0), Inches(1.2),
    "プレゼンテーションタイトル",
    font_size=40, bold=True, color=COLOR_WHITE, align=PP_ALIGN.LEFT
)

# サブタイトル
add_textbox(
    slide,
    Inches(1.0), Inches(3.7), Inches(11.0), Inches(0.6),
    "サブタイトル・概要を入力",
    font_size=20, color=COLOR_LIGHT, align=PP_ALIGN.LEFT
)

# 発表者・日付
add_textbox(
    slide,
    Inches(1.0), Inches(6.5), Inches(6.0), Inches(0.5),
    "発表者名　　|　　2026年5月22日",
    font_size=14, color=COLOR_LIGHT, align=PP_ALIGN.LEFT
)


# -----------------------------------------------
# 2. 目次スライド
# -----------------------------------------------
slide = prs.slides.add_slide(blank_layout)

# 上部ヘッダー
add_rect(slide, 0, 0, SLIDE_W, Inches(1.3), COLOR_PRIMARY)

# ヘッダータイトル
add_textbox(
    slide,
    Inches(0.6), Inches(0.3), Inches(12.0), Inches(0.8),
    "目次",
    font_size=28, bold=True, color=COLOR_WHITE, align=PP_ALIGN.LEFT
)

# 左側アクセントバー
add_rect(slide, 0, 0, Inches(0.18), SLIDE_H, COLOR_PRIMARY)

# 目次項目（番号付き）
items = [
    ("01", "はじめに・背景"),
    ("02", "課題・目的"),
    ("03", "提案・解決策"),
    ("04", "まとめ・今後の展開"),
]
for i, (num, label) in enumerate(items):
    y = Inches(1.8 + i * 1.2)

    # 番号バッジ
    add_rect(slide, Inches(0.8), y, Inches(0.7), Inches(0.7), COLOR_ACCENT)
    add_textbox(
        slide,
        Inches(0.8), y + Pt(4), Inches(0.7), Inches(0.6),
        num, font_size=18, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )

    # 項目テキスト
    add_textbox(
        slide,
        Inches(1.7), y + Pt(4), Inches(10.0), Inches(0.6),
        label, font_size=20, color=COLOR_DARK, align=PP_ALIGN.LEFT
    )

    # 区切り線
    if i < len(items) - 1:
        add_rect(slide, Inches(0.8), y + Inches(0.85), Inches(11.5), Inches(0.02), COLOR_LIGHT)


# -----------------------------------------------
# 3. コンテンツスライド
# -----------------------------------------------
slide = prs.slides.add_slide(blank_layout)

# 上部ヘッダー
add_rect(slide, 0, 0, SLIDE_W, Inches(1.3), COLOR_PRIMARY)

# 左側アクセントバー
add_rect(slide, 0, 0, Inches(0.18), SLIDE_H, COLOR_PRIMARY)

# スライドタイトル
add_textbox(
    slide,
    Inches(0.6), Inches(0.3), Inches(12.0), Inches(0.8),
    "スライドタイトルをここに入力",
    font_size=28, bold=True, color=COLOR_WHITE, align=PP_ALIGN.LEFT
)

# スライド番号
add_textbox(
    slide,
    Inches(11.5), Inches(0.4), Inches(1.5), Inches(0.5),
    "3 / 4",
    font_size=12, color=COLOR_LIGHT, align=PP_ALIGN.RIGHT
)

# 本文エリア（箇条書きプレースホルダー）
bullet_items = [
    "ここに箇条書き項目を入力します",
    "2つ目の要点をここに記載します",
    "3つ目の要点をここに記載します",
]
for i, item in enumerate(bullet_items):
    y = Inches(1.6 + i * 1.1)
    # ブレットマーク
    add_rect(slide, Inches(0.6), y + Inches(0.2), Inches(0.12), Inches(0.12), COLOR_ACCENT)
    add_textbox(
        slide,
        Inches(0.9), y, Inches(11.2), Inches(0.9),
        item, font_size=18, color=COLOR_DARK, align=PP_ALIGN.LEFT
    )

# フッター区切り線
add_rect(slide, 0, Inches(7.1), SLIDE_W, Inches(0.04), COLOR_LIGHT)

# フッターテキスト
add_textbox(
    slide,
    Inches(0.6), Inches(7.15), Inches(10.0), Inches(0.3),
    "社名 / チーム名",
    font_size=10, color=COLOR_GRAY, align=PP_ALIGN.LEFT
)


# -----------------------------------------------
# 4. まとめスライド
# -----------------------------------------------
slide = prs.slides.add_slide(blank_layout)

# 背景
add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_PRIMARY)

# 左側アクセントバー
add_rect(slide, 0, 0, Inches(0.18), SLIDE_H, COLOR_ACCENT)

# 上部ライン
add_rect(slide, Inches(1.0), Inches(1.8), Inches(11.0), Inches(0.05), COLOR_ACCENT)

# "まとめ" ラベル
add_textbox(
    slide,
    Inches(1.0), Inches(0.8), Inches(5.0), Inches(0.8),
    "まとめ",
    font_size=16, bold=True, color=COLOR_LIGHT, align=PP_ALIGN.LEFT
)

# メインメッセージ
add_textbox(
    slide,
    Inches(1.0), Inches(2.0), Inches(11.0), Inches(1.5),
    "ここに締めのメッセージを入力してください",
    font_size=34, bold=True, color=COLOR_WHITE, align=PP_ALIGN.LEFT
)

# キーポイント
key_points = [
    "重要なポイント 1：〇〇〇〇〇〇〇〇",
    "重要なポイント 2：〇〇〇〇〇〇〇〇",
    "重要なポイント 3：〇〇〇〇〇〇〇〇",
]
for i, pt in enumerate(key_points):
    add_textbox(
        slide,
        Inches(1.2), Inches(3.8 + i * 0.65), Inches(10.5), Inches(0.6),
        f"✓  {pt}",
        font_size=16, color=COLOR_LIGHT, align=PP_ALIGN.LEFT
    )

# Thank you テキスト
add_textbox(
    slide,
    Inches(1.0), Inches(6.6), Inches(11.0), Inches(0.6),
    "ご清聴ありがとうございました",
    font_size=14, italic=True, color=COLOR_LIGHT, align=PP_ALIGN.LEFT
)


# ===== 保存 =====
output_path = "/home/user/my-first-claude/template.pptx"
prs.save(output_path)
print(f"保存完了: {output_path}")
