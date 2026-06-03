#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""A0縦ポスターのレイアウト・イメージ図(モックアップ)を生成するスクリプト。
   ブロックごとに配色を変え、In vitro / In vivo を視覚的に区別する。"""

W, H = 841, 1189  # A0縦(mm)をユーザー単位として使用
FONT = "IPAGothic, sans-serif"

# 配色
C_HEADER = "#0d3b66"   # ヘッダー(濃紺)
C_INTRO  = "#e8eef5"   # 序論系(淡い青灰)
C_MAT    = "#fdf0d5"   # 材料合成(淡い黄)
C_VITRO  = "#dbeafe"   # In vitro(青)
C_VITRO_BAR = "#2563eb"
C_VIVO   = "#ffe3d3"   # In vivo(オレンジ)
C_VIVO_BAR = "#ea580c"
C_CONC   = "#d8f3dc"   # 結論(淡い緑)
C_BORDER = "#90a4b5"
C_FIG    = "#ffffff"   # 図プレースホルダ
C_FIGB   = "#b0bcc7"

svg = []
def add(s): svg.append(s)

def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def text(x, y, s, size=11, color="#1a1a1a", weight="normal", anchor="start"):
    add(f'<text x="{x}" y="{y}" font-family="{FONT}" font-size="{size}" '
        f'fill="{color}" font-weight="{weight}" text-anchor="{anchor}">{esc(s)}</text>')

def rect(x, y, w, h, fill, stroke=C_BORDER, rx=6, sw=1):
    add(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" '
        f'fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>')

def panel(x, y, w, h, title, lines, fill, bar=None, fig=None, fig_h=0):
    """パネル: タイトルバー + 本文 + 図プレースホルダ"""
    rect(x, y, w, h, fill, sw=1.2)
    # タイトルバー
    bar_c = bar if bar else "#4a5d6e"
    add(f'<rect x="{x}" y="{y}" width="{w}" height="22" rx="6" fill="{bar_c}"/>')
    add(f'<rect x="{x}" y="{y+12}" width="{w}" height="10" fill="{bar_c}"/>')
    text(x+8, y+16, title, size=11.5, color="#ffffff", weight="bold")
    ty = y + 38
    for ln in lines:
        text(x+8, ty, ln, size=9.2, color="#222")
        ty += 13
    # 図プレースホルダ
    if fig:
        fy = y + h - fig_h - 8
        rect(x+8, fy, w-16, fig_h, C_FIG, stroke=C_FIGB, rx=3, sw=1)
        add(f'<line x1="{x+8}" y1="{fy}" x2="{x+8+w-16}" y2="{fy+fig_h}" '
            f'stroke="{C_FIGB}" stroke-width="0.8"/>')
        add(f'<line x1="{x+8}" y1="{fy+fig_h}" x2="{x+8+w-16}" y2="{fy}" '
            f'stroke="{C_FIGB}" stroke-width="0.8"/>')
        text(x+w/2, fy+fig_h/2+4, fig, size=9, color="#6b7884", anchor="middle")

add(f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" '
    f'viewBox="0 0 {W} {H}">')
add(f'<rect width="{W}" height="{H}" fill="#f4f6f8"/>')

# ===== ヘッダー =====
M = 18
hx, hy, hw, hh = M, M, W-2*M, 96
rect(hx, hy, hw, hh, C_HEADER, stroke=C_HEADER, rx=8)
text(W/2, hy+30, "ポリグリセロールデンドリマーとグリコールキトサンの水素結合による",
     size=18, color="#ffffff", weight="bold", anchor="middle")
text(W/2, hy+54, "注入可能・自己修復型 超分子ハイドロゲルの創製",
     size=18, color="#ffffff", weight="bold", anchor="middle")
text(W/2, hy+74, "〜パクリタキセルの高可溶化と局所投与を目指して〜",
     size=12, color="#bcd0e6", anchor="middle")
text(W/2, hy+90, "〇山道 大誠¹  Cho Iksung²³  田中 賢²³   "
                 "（1 九大 工 応用化学  2 九大院 工学府  3 九大 先導研）",
     size=10, color="#dce7f2", anchor="middle")

# ===== 凡例 =====
ly = hy + hh + 8
def legend(x, c, label):
    add(f'<rect x="{x}" y="{ly}" width="16" height="12" rx="2" fill="{c}" stroke="{C_BORDER}"/>')
    text(x+20, ly+10, label, size=9.5, color="#333")
text(M, ly+10, "凡例:", size=9.5, color="#333", weight="bold")
legend(M+38, C_MAT, "材料合成")
legend(M+135, C_VITRO, "In vitro 評価")
legend(M+255, C_VIVO, "In vivo 評価")
legend(M+375, C_INTRO, "序論・戦略")
legend(M+475, C_CONC, "結論・展望")

# ===== 3カラム領域 =====
top = ly + 22
gap = 12
colw = (W - 2*M - 2*gap) / 3
x0 = M
x1 = M + colw + gap
x2 = M + 2*(colw+gap)
bottom_band_h = 168
col_bottom = H - M - bottom_band_h - gap
col_h_total = col_bottom - top

# --- 左カラム: 序論・戦略 ---
ph = (col_h_total - 2*gap) / 3
panel(x0, top, colw, ph, "1. Introduction｜化学療法の課題",
      ["・全身投与は腫瘍到達率 中央値0.7%",
       "・PTXは水溶解度<0.5µg/mL",
       "・可溶化剤Cremophor ELが重篤毒性",
       "→ 無毒可溶化+局所投与のDDSが必要"],
      C_INTRO, fig="図: PTX/CrEL構造・0.7%模式図", fig_h=70)
panel(x0, top+ph+gap, colw, ph, "2. 着眼点｜2機能を1材料で",
      ["・PGD: 単分散・中間水・生体適合性",
       "・単分子ハイドロトロープ→PTX高可溶化",
       "  (PEG400の約10倍)",
       "・shear-thinning→注入後に自己修復"],
      C_INTRO, fig="図: PGD構造・溶解度・shear機構", fig_h=80)
panel(x0, top+2*(ph+gap), colw, ph, "3. Objective & Strategy",
      ["目的: ハイドロトロープ機能と自己修復性",
       "を両立する注入ゲルの開発",
       "戦略: PGD+GCの水素結合(物理架橋)で",
       "形成する超分子ハイドロゲルを利用"],
      C_INTRO, fig="図: PGD+GC 超分子ゲル形成 (中心図)", fig_h=84)

# --- 中央カラム: 材料合成 + In vitro前半 ---
panel(x1, top, colw, ph, "A. 材料合成｜高純度 PGD G4 〔新〕",
      ["・Divergent法(全8段階)で合成",
       "・撹拌強化+反応時間2倍で最適化",
       "・G3.5→G4 収率90%, 13.0 g",
       "・MALDI単一ピーク[M+Na]+=3488"],
      C_MAT, bar="#c9962a", fig="図: MALDI 1回目 vs 2回目 / 収率表", fig_h=78)
panel(x1, top+ph+gap, colw, ph, "B-1. ゲル化時間の制御",
      ["・世代↑(G3→G4)で短縮",
       "・PGD/GC比↑で短縮",
       "・PG4-025: 2分 / PG3-1: 67分",
       "  (架橋点密度に依存)"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: ゲル化時間 vs アミノ基比", fig_h=82)
panel(x1, top+2*(ph+gap), colw, ph, "B-2. 架橋点の同定: GCアセチル化〔新〕",
      ["・GCアミノ基を部分N-アセチル化",
       "・DA↑でゲル化遅延(16.7%:1.5h→90%:80h)",
       "→ アミノ基が主架橋点と証明",
       "  アセチル化度で物性調整可"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: DA vs ゲル化時間 / NMR", fig_h=82)

# --- 右カラム: In vitro後半 ---
ph4 = (col_h_total - 3*gap) / 4
panel(x2, top, colw, ph4, "B-3. レオロジー(目玉)",
      ["・G'>G'' 安定ネットワーク",
       "・大歪500%でゾル化→低歪で即回復",
       "・自己修復を複数サイクル実証"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: Time sweep / ステップ歪み", fig_h=52)
panel(x2, top+ph4+gap, colw, ph4, "B-4. in vitro 分解性〔新〕",
      ["・37℃ PBS中の重量変化",
       "・PG3-1:3日 PG4-0125:4日 PG4-025:25日",
       "→ 滞留期間を数日〜数週間で制御"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: 分解曲線", fig_h=52)
panel(x2, top+2*(ph4+gap), colw, ph4, "B-5. 生体適合性 Live/Dead〔新〕",
      ["・D1細胞を3次元包埋し培養",
       "・1/5/9日とも死細胞ほぼ無し",
       "→ 化学架橋剤なしで高生存率"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: 蛍光像 (PG3-1/4-025/4-0125)", fig_h=52)
panel(x2, top+3*(ph4+gap), colw, ph4, "B-6. in vitro 注入試験〔新〕",
      ["・赤色化PG4-0125をシリンジ吐出",
       "・針内でゾル化→吐出後に即ゲル化",
       "→ 液だれせず形状保持"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: 注入の連続写真", fig_h=52)

# ===== 下部バンド: In vivo + Conclusion/Future/Ref =====
by = col_bottom + gap
bh = bottom_band_h
# In vivo (左1/3, 強調)
vivo_w = colw + gap*0  # 左カラム幅相当
panel(x0, by, colw+8, bh, "C. In vivo 評価｜マウス皮下注入(最重要・目玉)〔新〕",
      ["・香港大 Sang-Jin Lee先生と共同研究",
       "・PG4-0125(0.25mL)を背部皮下に注入",
       "・抵抗なく注入→生体内で即自己修復",
       "→ 局所に明瞭なドーム状デポを形成",
       "  低侵襲な局所投与キャリアを実証"],
      C_VIVO, bar=C_VIVO_BAR, fig="図: in vivo マウス皮下デポ写真 (大きく)", fig_h=72)

# Conclusion (中央〜右上)
cx = x1
cw = colw*2 + gap
panel(cx, by, cw, bh*0.62, "Conclusion",
      ["1. 材料: プロセス最適化で単分散PGD G4(13.0g, 収率90%, MALDI単一ピーク)を達成",
       "2. In vitro: 世代・混合比でゲル化時間/分解性を制御。アセチル化でアミノ基=主架橋点を証明。",
       "    レオロジーでシアシニング・自己修復を実証。D1細胞9日間高生存率。シリンジ注入で形状保持。",
       "3. In vivo: マウス皮下で抵抗なく注入→即自己修復し局所ドーム状デポを形成。実用性を実証。"],
      C_CONC, bar="#2d6a4f")

# Future + References (中央〜右下)
fy = by + bh*0.62 + 8
fh = bh - bh*0.62 - 8
panel(cx, fy, cw, fh, "Future Plans ／ References",
      ["Future: HPLCでPTX徐放定量 / 長期生分解性 / 担がんマウスで抗腫瘍効果 / PGD G5合成 / DSCで中間水評価",
       "Ref: Wilhelm Nat.Rev.Mater.2016 ; Ooya Gels 2022 ; Cho & Ooya Chem.Asian J.2018 ; "
       "Yamazaki Langmuir 2021 ; Tanaka BCSJ 2019 ほか"],
      "#eef2f4", bar="#5a6b78")

add('</svg>')

with open("poster/poster_mockup.svg", "w", encoding="utf-8") as f:
    f.write("\n".join(svg))
print("SVG生成完了: poster/poster_mockup.svg")
