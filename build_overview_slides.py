#!/usr/bin/env python3
"""材料概要スライド（Slide 9・10）だけを独立して生成する。
余白を減らし、カード内に具体的な情報を追加した改良版。
build_pptx.py のユーティリティ・色・サイズ定義を再利用する。"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

import build_pptx as B
from build_pptx import (
    Inches as _I,  # noqa
    rect, rrect, text, draw_header, draw_footer, new_slide,
    W, H, PAD, HDR_H, FTR_H, BODY_Y, BODY_H, BODY_X, BODY_W,
    INK, INK2, INK3, MUTED, LINE, SOFT, CARD_C, HEADER, ACCENT,
    AMBER, GOLD, WHITE, D_BLUE, D_TEAL, D_AMBR, D_GRN, D_RED, GRAY_L,
)


# ===== SLIDE 9: 材料構成（余白を詰めた改良版）=====
def slide_mat_composition(prs):
    sl = new_slide(prs)
    draw_header(sl, "材料 概要 1/2", "材料構成：持続可能かつ高性能")
    draw_footer(sl, "9 / 39")

    # リード文（カード上部の余白を活用）
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

        # アイコン
        icon_y = cards_y + Inches(0.30)
        text(sl, cx, icon_y, cw, Inches(0.56),
             icon, size=30, color=accent_c,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # タイトル＋英名
        title_y = icon_y + Inches(0.62)
        text(sl, cx + pad_x, title_y, inner_w, Inches(0.42),
             title, size=17, bold=True, color=INK, align=PP_ALIGN.CENTER)
        text(sl, cx + pad_x, title_y + Inches(0.42), inner_w, Inches(0.28),
             en, size=10, italic=True, color=MUTED, align=PP_ALIGN.CENTER)

        # 区切り線
        div_y = title_y + Inches(0.78)
        rect(sl, cx + Inches(0.32), div_y, cw - Inches(0.64), Inches(0.015), LINE)

        # 本文
        body_y = div_y + Inches(0.14)
        text(sl, cx + pad_x, body_y, inner_w, Inches(1.00),
             body, size=12, color=INK3, align=PP_ALIGN.CENTER)

        # キーポイント行（カード下部の余白を埋める）
        ky = body_y + Inches(1.22)
        row_h = Inches(0.74)
        for j, (label, detail) in enumerate(keys):
            ry = ky + j * row_h
            # ラベルチップ
            chip_w = Inches(0.92)
            rrect(sl, cx + pad_x, ry + Inches(0.04), chip_w, Inches(0.32),
                  WHITE, accent_c, 0.7, radius=0.5)
            text(sl, cx + pad_x, ry + Inches(0.04), chip_w, Inches(0.32),
                 label, size=10, bold=True, color=accent_c,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            # 詳細
            text(sl, cx + pad_x + chip_w + Inches(0.10), ry,
                 inner_w - chip_w - Inches(0.10), Inches(0.40),
                 detail, size=11, color=INK2, anchor=MSO_ANCHOR.MIDDLE)


# ===== SLIDE 10: 噴霧・付着・耐火のメカニズム（余白を詰めた改良版）=====
def slide_self_protection(prs):
    sl = new_slide(prs)
    draw_header(sl, "材料 概要 2/2", "噴霧・付着・耐火のメカニズム")
    draw_footer(sl, "10 / 39")

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
         "せん断希薄化 n < 1", D_TEAL),
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

        # 番号
        text(sl, cx + pad_x, cards_y + Inches(0.18), inner_w, Inches(0.34),
             no, size=13, bold=True, color=c)
        # フェーズ（小ラベル）
        text(sl, cx + pad_x, cards_y + Inches(0.54), inner_w, Inches(0.30),
             phase, size=12, bold=True, color=MUTED)
        # 状態（大見出し）
        text(sl, cx + pad_x, cards_y + Inches(0.92), inner_w, Inches(0.56),
             state, size=22, bold=True, color=INK)
        # 区切り
        rect(sl, cx + pad_x, cards_y + Inches(1.58), inner_w, Inches(0.015), LINE)
        # 説明
        text(sl, cx + pad_x, cards_y + Inches(1.72), inner_w, Inches(1.10),
             desc, size=12, color=INK3)
        # メトリクスバッジ（下部）
        badge_h = Inches(0.46)
        badge_y = cards_y + card_h - badge_h - Inches(0.18)
        rrect(sl, cx + pad_x, badge_y, inner_w, badge_h, WHITE, c, 0.7, radius=0.06)
        text(sl, cx + pad_x, badge_y, inner_w, badge_h,
             metric, size=11, bold=True, color=c,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # 矢印
        if i < n - 1:
            text(sl, cx + cw, cards_y, arrow_w, card_h,
                 "→", size=20, color=GRAY_L,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 下部：ポイント帯（残りの余白を埋める）
    band_y = cards_y + card_h + Inches(0.20)
    band_h = H - band_y - FTR_H - Inches(0.22)
    rrect(sl, BODY_X, band_y, BODY_W, band_h, HEADER, None, radius=0.05)
    rect(sl, BODY_X, band_y, Inches(0.08), band_h, GOLD)
    text(sl, BODY_X + Inches(0.30), band_y, Inches(2.4), band_h,
         "POINT", size=13, bold=True, color=GOLD, anchor=MSO_ANCHOR.MIDDLE)
    text(sl, BODY_X + Inches(2.0), band_y, BODY_W - Inches(2.3), band_h,
         "MC の温度応答性（LCST）に頼らず、レオロジー特性だけで「飛ばす→留める→固める」を実現する点が本材料の核心です。",
         size=13, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)


def main():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    print("材料概要スライド（Slide 9・10）を構築中...")
    slide_mat_composition(prs)
    slide_self_protection(prs)

    out_path = '/home/user/my-first-claude/overview_slides.pptx'
    prs.save(out_path)
    print(f"完了: {out_path}")
    print(f"スライド数: {len(prs.slides.__iter__.__self__._sldIdLst)}")


if __name__ == '__main__':
    main()
