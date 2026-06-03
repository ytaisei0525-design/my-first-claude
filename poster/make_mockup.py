#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""A0縦ポスターのレイアウト・イメージ図(モックアップ)。
   既存ポスターに合わせ、上部=Introduction(全幅)、その下=データ(Results)を配置。
   In vitro / In vivo はブロックごとに配色で区別する。"""

W, H = 841, 1189  # A0縦(mm)をユーザー単位として使用
FONT = "IPAGothic, sans-serif"

# 配色
C_HEADER = "#0d3b66"
C_INTRO  = "#e8eef5"   # 序論・着眼点・戦略
C_MAT    = "#fdf0d5"   # 材料合成
C_VITRO  = "#dbeafe"   # In vitro
C_VITRO_BAR = "#2563eb"
C_VIVO   = "#ffe3d3"   # In vivo
C_VIVO_BAR = "#ea580c"
C_CONC   = "#d8f3dc"   # 結論
C_BORDER = "#90a4b5"
C_FIG    = "#ffffff"
C_FIGB   = "#b0bcc7"

svg = []
def add(s): svg.append(s)
def esc(t): return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def text(x, y, s, size=11, color="#1a1a1a", weight="normal", anchor="start"):
    add(f'<text x="{x}" y="{y}" font-family="{FONT}" font-size="{size}" '
        f'fill="{color}" font-weight="{weight}" text-anchor="{anchor}">{esc(s)}</text>')

def rect(x, y, w, h, fill, stroke=C_BORDER, rx=6, sw=1):
    add(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" '
        f'fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>')

def sectionbar(x, y, w, label, color):
    """全幅のセクション見出しバー"""
    add(f'<rect x="{x}" y="{y}" width="{w}" height="26" rx="5" fill="{color}"/>')
    text(x+12, y+18, label, size=14, color="#ffffff", weight="bold")

def panel(x, y, w, h, title, lines, fill, bar="#4a5d6e", fig=None, fig_h=0, tsize=9.2):
    rect(x, y, w, h, fill, sw=1.2)
    add(f'<rect x="{x}" y="{y}" width="{w}" height="20" rx="6" fill="{bar}"/>')
    add(f'<rect x="{x}" y="{y+10}" width="{w}" height="10" fill="{bar}"/>')
    text(x+7, y+15, title, size=10.5, color="#ffffff", weight="bold")
    ty = y + 34
    for ln in lines:
        text(x+7, ty, ln, size=tsize, color="#222")
        ty += 12.5
    if fig:
        fy = y + h - fig_h - 7
        rect(x+7, fy, w-14, fig_h, C_FIG, stroke=C_FIGB, rx=3, sw=1)
        add(f'<line x1="{x+7}" y1="{fy}" x2="{x+7+w-14}" y2="{fy+fig_h}" stroke="{C_FIGB}" stroke-width="0.8"/>')
        add(f'<line x1="{x+7}" y1="{fy+fig_h}" x2="{x+7+w-14}" y2="{fy}" stroke="{C_FIGB}" stroke-width="0.8"/>')
        text(x+w/2, fy+fig_h/2+4, fig, size=8.6, color="#6b7884", anchor="middle")

add(f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">')
add(f'<rect width="{W}" height="{H}" fill="#f4f6f8"/>')

M = 18
gap = 10

# ===== ヘッダー =====
hx, hy, hw, hh = M, M, W-2*M, 88
rect(hx, hy, hw, hh, C_HEADER, stroke=C_HEADER, rx=8)
text(W/2, hy+28, "ポリグリセロールデンドリマーとグリコールキトサンの水素結合による",
     size=17, color="#ffffff", weight="bold", anchor="middle")
text(W/2, hy+50, "注入可能・自己修復型 超分子ハイドロゲルの創製",
     size=17, color="#ffffff", weight="bold", anchor="middle")
text(W/2, hy+70, "〇山道 大誠¹  Cho Iksung²³  田中 賢²³ （九州大学 工学部／工学府／先導物質化学研究所）",
     size=10.5, color="#cfe0f0", anchor="middle")

cx = M
cw = W - 2*M

# ===== Introduction(全幅・上部)=====
iy = hy + hh + gap
ih = 250
sectionbar(cx, iy, cw, "Introduction ｜ 背景・着眼点・目的", C_HEADER)
sub_y = iy + 32
sub_h = ih - 32
sub_gap = 10
sub_w = (cw - 2*sub_gap) / 3
sx0 = cx
sx1 = cx + sub_w + sub_gap
sx2 = cx + 2*(sub_w + sub_gap)
panel(sx0, sub_y, sub_w, sub_h, "化学療法の課題",
      ["・全身投与は腫瘍到達率 中央値0.7%",
       "・骨髄抑制・脱毛など重篤な副作用",
       "・PTXは水溶解度<0.5µg/mL",
       "・可溶化剤Cremophor ELが重篤毒性",
       "→ 無毒可溶化+局所投与のDDSが必要"],
      C_INTRO, bar="#4a6178", fig="図: PTX/CrEL構造・0.7%模式図", fig_h=80)
panel(sx1, sub_y, sub_w, sub_h, "着眼点: 2機能を1材料で",
      ["・PGD: 単分散・中間水・生体適合性",
       "・単分子ハイドロトロープ機能",
       "  → PTXを高可溶化(PEG400の約10倍)",
       "・shear-thinning型: 注入後に自己修復",
       "  → 液だれせず患部に留まる"],
      C_INTRO, bar="#4a6178", fig="図: PGD構造 / 溶解度 / shear機構", fig_h=80)
panel(sx2, sub_y, sub_w, sub_h, "目的・戦略",
      ["目的: ハイドロトロープ機能と自己修復性を",
       "両立した注入ゲルの開発",
       "戦略: PGD+GCの水素結合(物理架橋)で",
       "自発形成する超分子ハイドロゲルを利用",
       "→ PGDがPTXを高濃度保持し局所徐放"],
      C_INTRO, bar="#4a6178", fig="図: PGD+GC 超分子ゲル形成(中心図)", fig_h=80)

# ===== Results & Discussion(データ)=====
ry = iy + ih + gap
sectionbar(cx, ry, cw, "Results & Discussion ｜ データ", "#33506b")

# 下部バンドの位置を先に確保
band_conc_h = 96
band_vivo_h = 120
conc_y = H - M - band_conc_h
vivo_y = conc_y - gap - band_vivo_h

# データ用カラム領域
data_y = ry + 32
data_bottom = vivo_y - gap
data_h = data_bottom - data_y
colw = (cw - 2*gap) / 3
dx0 = cx
dx1 = cx + colw + gap
dx2 = cx + 2*(colw + gap)

# --- Col1: 材料合成 + In vitro ---
p1 = (data_h - gap) / 2
panel(dx0, data_y, colw, p1, "A. 材料合成: 高純度 PGD G4 〔新〕",
      ["・Divergent法(全8段階)で合成",
       "・撹拌強化+反応時間2倍で最適化",
       "・G3.5→G4 収率90%, 13.0 g 取得",
       "・MALDI単一ピーク [M+Na]+=3488",
       "  → 単分散の高品質体を確認"],
      C_MAT, bar="#c9962a", fig="図: MALDI(1回目vs2回目) / 収率表", fig_h=p1-92)
panel(dx0, data_y+p1+gap, colw, p1, "B-1. ゲル化時間の制御 (In vitro)",
      ["・世代↑(G3→G4)で短縮",
       "・PGD/GC比↑で短縮",
       "・PG4-025: 2分 / PG3-1: 67分",
       "  (架橋点密度に依存)"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: ゲル化時間 vs アミノ基比", fig_h=p1-80)

# --- Col2: In vitro ---
panel(dx1, data_y, colw, p1, "B-2. 架橋点の同定: GCアセチル化 (In vitro)〔新〕",
      ["・GCアミノ基を部分N-アセチル化",
       "・DA↑でゲル化遅延",
       "  (DA16.7%:1.5h → 90%:80h)",
       "→ アミノ基が主架橋点と証明",
       "  アセチル化度で物性調整も可"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: DA vs ゲル化時間 / NMR", fig_h=p1-92)
panel(dx1, data_y+p1+gap, colw, p1, "B-3. レオロジー (In vitro・目玉)",
      ["・G'>G'' 安定なゲルネットワーク",
       "・大歪500%でゾル化",
       "・低歪に戻すと即G'回復(自己修復)",
       "・複数サイクルで再現"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: Time sweep / ステップ歪み", fig_h=p1-80)

# --- Col3: In vitro(分解性・Live/Dead・注入)---
p3 = (data_h - 2*gap) / 3
panel(dx2, data_y, colw, p3, "B-4. in vitro 分解性〔新〕",
      ["・37℃ PBS中の重量変化で評価",
       "・PG3-1:3日/PG4-0125:4日/PG4-025:25日",
       "→ 滞留期間を数日〜数週間で制御"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: 分解曲線", fig_h=p3-58)
panel(dx2, data_y+p3+gap, colw, p3, "B-5. 生体適合性 Live/Dead〔新〕",
      ["・D1細胞を3次元包埋し培養",
       "・1/5/9日とも死細胞ほぼ無し",
       "→ 化学架橋剤なしで高生存率"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: 蛍光像(PG3-1/4-025/4-0125)", fig_h=p3-58)
panel(dx2, data_y+2*(p3+gap), colw, p3, "B-6. in vitro 注入試験〔新〕",
      ["・赤色化PG4-0125をシリンジ吐出",
       "・針内でゾル化→吐出後に即ゲル化",
       "→ 液だれせず形状を保持"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: 注入の連続写真", fig_h=p3-58)

# ===== In vivo(全幅・独立帯・色で区別)=====
rect(cx, vivo_y, cw, band_vivo_h, C_VIVO, stroke=C_VIVO_BAR, rx=8, sw=2)
add(f'<rect x="{cx}" y="{vivo_y}" width="{cw}" height="22" rx="8" fill="{C_VIVO_BAR}"/>')
add(f'<rect x="{cx}" y="{vivo_y+12}" width="{cw}" height="10" fill="{C_VIVO_BAR}"/>')
text(cx+12, vivo_y+16, "C. In vivo 評価 ｜ マウス皮下注入試験（最重要・目玉）〔新データ〕",
     size=13, color="#ffffff", weight="bold")
# 左テキスト
vtx = cx + 12
ty = vivo_y + 40
for ln in ["・香港大 Sang-Jin Lee先生との共同研究",
           "・PG4-0125(0.25mL)をマウス背部皮下へ注入",
           "・抵抗なくスムーズに注入 → 生体内で即座に自己修復",
           "→ 局所に明瞭なドーム状デポを形成。低侵襲な局所投与キャリアの実用性を実証"]:
    text(vtx, ty, ln, size=10, color="#222"); ty += 16
# 右に大きめ図枠
fvw = cw*0.34
fvx = cx + cw - fvw - 12
fvy = vivo_y + 30
fvh = band_vivo_h - 40
rect(fvx, fvy, fvw, fvh, C_FIG, stroke=C_FIGB, rx=3, sw=1)
add(f'<line x1="{fvx}" y1="{fvy}" x2="{fvx+fvw}" y2="{fvy+fvh}" stroke="{C_FIGB}" stroke-width="0.8"/>')
add(f'<line x1="{fvx}" y1="{fvy+fvh}" x2="{fvx+fvw}" y2="{fvy}" stroke="{C_FIGB}" stroke-width="0.8"/>')
text(fvx+fvw/2, fvy+fvh/2+4, "図: in vivo マウス皮下デポ写真(大きく)", size=9.5, color="#6b7884", anchor="middle")

# ===== Conclusion / Future / References(最下段・全幅)=====
rect(cx, conc_y, cw, band_conc_h, C_CONC, stroke="#2d6a4f", rx=8, sw=1.5)
add(f'<rect x="{cx}" y="{conc_y}" width="{cw}" height="20" rx="8" fill="#2d6a4f"/>')
add(f'<rect x="{cx}" y="{conc_y+10}" width="{cw}" height="10" fill="#2d6a4f"/>')
text(cx+12, conc_y+15, "Conclusion / Future Plans / References", size=12, color="#ffffff", weight="bold")
ty = conc_y + 34
for ln in ["① 材料: 最適化で単分散PGD G4(13.0g,収率90%,MALDI単一ピーク)を達成   "
           "② In vitro: ゲル化時間・分解性を制御、アミノ基=主架橋点を証明、自己修復実証、9日間高生存率",
           "③ In vivo: マウス皮下で即自己修復し局所ドーム状デポ形成 → 実用性実証",
           "Future: PTX徐放定量(HPLC) / 長期生分解性 / 担がんマウスで抗腫瘍効果 / PGD G5合成 / DSCで中間水評価",
           "Ref: Wilhelm 2016 ; Ooya Gels 2022 ; Cho&Ooya 2018 ; Yamazaki Langmuir 2021 ; Tanaka BCSJ 2019 ほか"]:
    text(cx+12, ty, ln, size=8.8, color="#1d3a2a"); ty += 13

# ===== 凡例(右下小)=====
def legend(x, y, c, label):
    add(f'<rect x="{x}" y="{y}" width="13" height="10" rx="2" fill="{c}" stroke="{C_BORDER}"/>')
    text(x+17, y+9, label, size=8.5, color="#333")
lgy = ry + 6
legend(cx+cw-330, lgy, C_MAT, "材料合成")
legend(cx+cw-250, lgy, C_VITRO, "In vitro")
legend(cx+cw-165, lgy, C_VIVO, "In vivo")

add('</svg>')
with open("poster/poster_mockup.svg","w",encoding="utf-8") as f:
    f.write("\n".join(svg))
print("SVG生成完了")
