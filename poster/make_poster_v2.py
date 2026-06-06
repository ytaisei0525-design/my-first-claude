#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""デザイン改良版ポスター(モックアップ)。
   改善点: キャッチコピー / 数値カットアウト帯 / 流れ矢印 / 配色の規律 / ヒーロー図の強調。"""

W, H = 841, 1189
FONT = "IPAGothic, sans-serif"

# 配色(3テーマ色に規律)
NAVY    = "#0d3b66"
MAT     = "#fdf0d5"; MAT_BAR = "#c08a1e"   # 合成(黄系)
VITRO   = "#dbeafe"; VITRO_BAR = "#2563eb" # In vitro(青)
VIVO    = "#ffe3d3"; VIVO_BAR = "#ea580c"  # 実用/In vivo(橙)
CONC    = "#d8f3dc"; CONC_BAR = "#2d6a4f"
SEC     = "#33506b"
BORDER  = "#9bb0c2"
FIG     = "#ffffff"; FIGB = "#c2cdd8"
GOLD    = "#f4b400"  # ヒーロー強調

svg = []
def add(s): svg.append(s)
def esc(t): return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def text(x, y, s, size=11, color="#1a1a1a", weight="normal", anchor="start", italic=False):
    st = ' font-style="italic"' if italic else ''
    add(f'<text x="{x}" y="{y}" font-family="{FONT}" font-size="{size}" '
        f'fill="{color}" font-weight="{weight}" text-anchor="{anchor}"{st}>{esc(s)}</text>')

def rect(x, y, w, h, fill, stroke=BORDER, rx=6, sw=1):
    add(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" fill="{fill}" '
        f'stroke="{stroke}" stroke-width="{sw}"/>')

def sectionbar(x, y, w, label, color, h=26, size=14):
    add(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="5" fill="{color}"/>')
    text(x+12, y+h/2+5, label, size=size, color="#ffffff", weight="bold")

def figbox(x, y, w, h, caption):
    rect(x, y, w, h, FIG, stroke=FIGB, rx=3, sw=1)
    add(f'<line x1="{x}" y1="{y}" x2="{x+w}" y2="{y+h}" stroke="{FIGB}" stroke-width="0.8"/>')
    add(f'<line x1="{x}" y1="{y+h}" x2="{x+w}" y2="{y}" stroke="{FIGB}" stroke-width="0.8"/>')
    text(x+w/2, y+h/2+4, caption, size=8.6, color="#6b7884", anchor="middle")

def panel(x, y, w, h, title, lines, fill, bar, fig=None, fig_h=0, hero=False):
    sw = 2.6 if hero else 1.2
    stroke = GOLD if hero else BORDER
    rect(x, y, w, h, fill, stroke=stroke, sw=sw)
    add(f'<rect x="{x}" y="{y}" width="{w}" height="20" rx="6" fill="{bar}"/>')
    add(f'<rect x="{x}" y="{y+10}" width="{w}" height="10" fill="{bar}"/>')
    text(x+8, y+15, title, size=10.5, color="#ffffff", weight="bold")
    if hero:
        text(x+w-8, y+15, "★目玉", size=10, color=GOLD, weight="bold", anchor="end")
    ty = y + 34
    for ln in lines:
        text(x+8, ty, ln, size=9.2, color="#222"); ty += 12.5
    if fig:
        figbox(x+8, y+h-fig_h-7, w-16, fig_h, fig)

add(f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">')
add(f'<rect width="{W}" height="{H}" fill="#f4f6f8"/>')

M = 18; gap = 10
cx = M; cw = W - 2*M

# ===== ヘッダー(タイトル+キャッチコピー+著者) =====
hh = 104
rect(cx, M, cw, hh, NAVY, stroke=NAVY, rx=8)
# ロゴ枠(左右)
rect(cx+10, M+18, 78, hh-36, "#ffffff", stroke="#cfe0f0", rx=6); text(cx+49, M+hh/2+4, "九州大学", size=11, color=NAVY, weight="bold", anchor="middle")
rect(cx+cw-88, M+18, 78, hh-36, "#ffffff", stroke="#cfe0f0", rx=6); text(cx+cw-49, M+hh/2+4, "Tanaka Lab.", size=10, color=NAVY, weight="bold", anchor="middle")
text(W/2, M+28, "ポリグリセロールデンドリマーとグリコールキトサンの水素結合による",
     size=17, color="#ffffff", weight="bold", anchor="middle")
text(W/2, M+50, "注入可能・自己修復型 超分子ハイドロゲルの創製",
     size=17, color="#ffffff", weight="bold", anchor="middle")
# キャッチコピー(リボン)
add(f'<rect x="{W/2-300}" y="{M+58}" width="600" height="20" rx="10" fill="{GOLD}"/>')
text(W/2, M+72, "可溶化剤フリーで高濃度可溶化 × 注射針で注入 × 生体内で自己修復する超分子ゲル",
     size=11, color="#3a2c00", weight="bold", anchor="middle")
text(W/2, M+95, "〇山道 大誠¹  Cho Iksung²³  田中 賢²³ （九州大学 工学部／工学府／先導物質化学研究所）",
     size=10, color="#cfe0f0", anchor="middle")

# ===== KEY FINDINGS 数値カットアウト帯 =====
ky = M + hh + gap
kh = 70
rect(cx, ky, cw, kh, "#ffffff", stroke=SEC, sw=1.5, rx=8)
add(f'<rect x="{cx}" y="{ky}" width="150" height="{kh}" rx="8" fill="{SEC}"/>')
text(cx+75, ky+kh/2-6, "KEY", size=20, color="#fff", weight="bold", anchor="middle")
text(cx+75, ky+kh/2+16, "FINDINGS", size=15, color="#fff", weight="bold", anchor="middle")
stats = [("90%", "G3.5→G4 ステップ収率", MAT_BAR),
         ("単分散", "MALDI 単一ピーク [M+Na]⁺=3488", MAT_BAR),
         ("〜25日", "皮下デポ滞留(数日〜数週で制御)", VIVO_BAR),
         ("9日間", "高細胞生存率 (Live/Dead)", VITRO_BAR)]
sx = cx + 160
sw_ = (cw - 160 - 10) / 4
for i,(num,lab,col) in enumerate(stats):
    bx = sx + i*sw_
    if i>0: add(f'<line x1="{bx}" y1="{ky+12}" x2="{bx}" y2="{ky+kh-12}" stroke="#dde4ea" stroke-width="1.5"/>')
    text(bx+sw_/2, ky+34, num, size=27, color=col, weight="bold", anchor="middle")
    text(bx+sw_/2, ky+54, lab, size=9.2, color="#333", anchor="middle")

# ===== Introduction =====
iy = ky + kh + gap
ih = 168
sectionbar(cx, iy, cw, "Introduction ｜ 背景・着眼点・目的", NAVY)
sub_y = iy + 32; sub_h = ih - 32
sw3 = (cw - 2*gap)/3
sxs = [cx, cx+sw3+gap, cx+2*(sw3+gap)]
panel(sxs[0], sub_y, sw3, sub_h, "化学療法の課題",
      ["・全身投与は腫瘍到達率 中央値0.7%",
       "・PTXは水溶解度<0.5µg/mL",
       "・可溶化剤Cremophor ELが重篤毒性",
       "→ 無毒可溶化+局所投与のDDSが必要"],
      "#e8eef5", "#4a6178", fig="図: PTX/CrEL構造・0.7%模式図", fig_h=60)
panel(sxs[1], sub_y, sw3, sub_h, "着眼点: 2機能を1材料で",
      ["・PGD: 単分散・中間水・生体適合性",
       "・単分子ハイドロトロープ→PTX高可溶化",
       "・shear-thinning型→注入後に自己修復",
       "→ 液だれせず患部に留まる"],
      "#e8eef5", "#4a6178", fig="図: PGD構造 / 溶解度 / shear機構", fig_h=60)
panel(sxs[2], sub_y, sw3, sub_h, "目的・戦略",
      ["目的: ハイドロトロープ機能と自己修復性を",
       "  両立した注入ゲルの開発",
       "戦略: PGD+GCの水素結合(物理架橋)で",
       "  超分子ハイドロゲルを形成"],
      "#e8eef5", "#4a6178", fig="図: PGD+GC 超分子ゲル形成(中心図)", fig_h=60)

# ===== Results =====
ry = iy + ih + gap
sectionbar(cx, ry, cw, "Results & Discussion", SEC)

band_conc_h = 66
conc_y = H - M - band_conc_h

colw = (cw - 2*gap)/3
dx = [cx, cx+colw+gap, cx+2*(colw+gap)]

# 列見出し + 流れ矢印
head_y = ry + 32; head_h = 18
heads = [("材料合成・インジェクタブル特性", SEC),
         ("In vitro 生体適合性評価", VITRO_BAR),
         ("実用性評価 (In vitro / In vivo)", VIVO_BAR)]
for i,(lab,col) in enumerate(heads):
    add(f'<rect x="{dx[i]}" y="{head_y}" width="{colw}" height="{head_h}" rx="4" fill="{col}"/>')
    text(dx[i]+colw/2, head_y+12.5, lab, size=11, color="#fff", weight="bold", anchor="middle")
# 流れ矢印(列間)
for i in range(2):
    axc = dx[i]+colw + gap/2
    add(f'<polygon points="{axc-5},{head_y+3} {axc+5},{head_y+head_h/2} {axc-5},{head_y+head_h-3}" fill="{GOLD}"/>')

data_y = head_y + head_h + 8
data_bottom = conc_y - gap
data_h = data_bottom - data_y
ph3 = (data_h - 2*gap)/3
ph2 = (data_h - gap)/2

# 左列
panel(dx[0], data_y, colw, ph3, "① PGD G4 の精密合成",
      ["・Divergent法で単分散PGD G4を合成",
       "・撹拌強化+反応時間2倍で最適化",
       "・収率90% / MALDI単一ピーク"],
      MAT, MAT_BAR, fig="図: MALDI(最適化前後) / 収率表", fig_h=ph3-76)
panel(dx[0], data_y+ph3+gap, colw, ph3, "② 自己修復性 (レオロジー)",
      ["・せん断除去でG'が即座に回復",
       "・元の強固なゲル構造へ再構築",
       "・複数サイクルで再現性よく回復"],
      VITRO, VITRO_BAR, fig="図: ステップ歪みサイクル (G'/G'')", fig_h=ph3-76, hero=True)
panel(dx[0], data_y+2*(ph3+gap), colw, ph3, "③ シアシニング性 (レオロジー)",
      ["・高せん断でゲルが流動化(ゾル化)",
       "・粘度が急低下→注射針を通過可能",
       "・大歪500%で G'<G''"],
      VITRO, VITRO_BAR, fig="図: せん断速度-粘度 / 大歪でのG'低下", fig_h=ph3-76)

# 中列
panel(dx[1], data_y, colw, ph2, "④ In vitro 安全性｜Live/Dead",
      ["・D1細胞を3次元ゲル内に包埋し培養",
       "・1/5/9日とも死細胞ほぼ無し",
       "・溶解後も毒性なし",
       "→ 化学架橋剤なしで高い細胞生存率"],
      VITRO, VITRO_BAR, fig="図: Live/Dead 蛍光像 (3条件×1・5・9日)", fig_h=ph2-90)
panel(dx[1], data_y+ph2+gap, colw, ph2, "⑤ In vitro 分解性",
      ["・37℃ PBS中の重量変化で分解を評価",
       "・PG3-1:3日 / PG4-0125:4日 / PG4-025:25日",
       "・世代・混合比で滞留期間を制御",
       "→ 数日〜数週間で任意に制御"],
      VITRO, VITRO_BAR, fig="図: 分解曲線 (重量比 vs 時間)", fig_h=ph2-90)

# 右列
ph_rt = (data_h - gap)*0.36; ph_rb = (data_h - gap)*0.64
panel(dx[2], data_y, colw, ph_rt, "⑥ 注入試験 (In vitro & In vivo)",
      ["・in vitro: 赤色化ゲルをシリンジ吐出→即ゲル化",
       "・in vivo: マウス背部皮下へ注入(0.25 mL)",
       "・抵抗なく注入→液漏れせずデポ形成"],
      VIVO, VIVO_BAR, fig="図: in vitro注入写真 / in vivoデポ写真", fig_h=ph_rt-76)
panel(dx[2], data_y+ph_rt+gap, colw, ph_rb, "⑦ In vivo 組織学的評価 H&E染色",
      ["・Day 0/5/14で皮下デポを採取しH&E評価",
       "・ゲル辺縁から宿主細胞が内部へ浸潤(矢印)",
       "・Day5→14で浸潤進行→緩やかに生分解・組織置換",
       "・重度の炎症・壊死なし→良好な生体適合性",
       "・in vitro分解(≈25日)と整合→局所貯留に好適"],
      VIVO, VIVO_BAR, fig="図: スキーム / Optical / H&E組織像 (Native vs PG4-025×0・5・14日)", fig_h=ph_rb-103, hero=True)

# ===== Conclusion =====
rect(cx, conc_y, cw, band_conc_h, CONC, stroke=CONC_BAR, rx=6, sw=1.5)
text(cx+10, conc_y+18, "Conclusion:", size=11, color="#1d3a2a", weight="bold")
text(cx+95, conc_y+18, "単分散PGD G4を確立 → 水素結合で超分子ゲルを形成(自己修復・分解性を制御・高生体適合性) → in vivoで局所デポを形成し緩やかに生分解",
     size=9.3, color="#1d3a2a")
text(cx+10, conc_y+38, "Future:", size=10, color=SEC, weight="bold")
text(cx+62, conc_y+38, "PTX徐放定量(HPLC) / 担がんマウスで抗腫瘍効果 / PGD G5合成 / DSCで中間水評価",
     size=8.8, color="#3a4a55")
# 謝辞 + 参考文献(最下行・小フォント)
add(f'<line x1="{cx+10}" y1="{conc_y+46}" x2="{cx+cw-10}" y2="{conc_y+46}" stroke="#bcd0bf" stroke-width="0.8"/>')
text(cx+10, conc_y+59, "謝辞:", size=8.5, color="#2d6a4f", weight="bold")
text(cx+44, conc_y+59,
     "本研究の in vivo 試験は香港大学 Sang-Jin Lee 先生との共同研究による。"
     "　　Ref: Ooya, Gels 2022 / Cho & Ooya, Chem. Asian J. 2018 / Yamazaki, Langmuir 2021 ほか",
     size=8, color="#3a4a55")

add('</svg>')
open("poster/poster_v2.svg","w",encoding="utf-8").write("\n".join(svg))
print("done")
