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
ih = 220
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
      C_INTRO, bar="#4a6178", fig="図: PTX/CrEL構造・0.7%模式図", fig_h=66)
panel(sx1, sub_y, sub_w, sub_h, "着眼点: 2機能を1材料で",
      ["・PGD: 単分散・中間水・生体適合性",
       "・単分子ハイドロトロープ機能",
       "  → PTXを高可溶化(PEG400の約10倍)",
       "・shear-thinning型: 注入後に自己修復",
       "  → 液だれせず患部に留まる"],
      C_INTRO, bar="#4a6178", fig="図: PGD構造 / 溶解度 / shear機構", fig_h=66)
panel(sx2, sub_y, sub_w, sub_h, "目的・戦略",
      ["目的: ハイドロトロープ機能と自己修復性を",
       "両立した注入ゲルの開発",
       "戦略: PGD+GCの水素結合(物理架橋)で",
       "自発形成する超分子ハイドロゲルを利用",
       "→ PGDがPTXを高濃度保持し局所徐放"],
      C_INTRO, bar="#4a6178", fig="図: PGD+GC 超分子ゲル形成(中心図)", fig_h=66)

# ===== Results & Discussion(データ)=====
ry = iy + ih + gap
sectionbar(cx, ry, cw, "Results & Discussion ｜ データ（左→右で研究の流れ）", "#33506b")

# 下部は薄い結論帯のみ(In vivo帯は廃止)
band_conc_h = 50
conc_y = H - M - band_conc_h

# データ用カラム領域
colw = (cw - 2*gap) / 3
dx0 = cx
dx1 = cx + colw + gap
dx2 = cx + 2*(colw + gap)

# --- 列ごとの見出しバー ---
head_y = ry + 32
head_h = 17
def colhead(x, label, color):
    add(f'<rect x="{x}" y="{head_y}" width="{colw}" height="{head_h}" rx="4" fill="{color}"/>')
    text(x+colw/2, head_y+12.5, label, size=12, color="#ffffff", weight="bold", anchor="middle")
colhead(dx0, "材料合成・インジェクタブル特性", "#33506b")
colhead(dx1, "In vitro 生体適合性評価", C_VITRO_BAR)
colhead(dx2, "実用性評価 (In vitro / In vivo)", C_VIVO_BAR)

data_y = head_y + head_h + 7
data_bottom = conc_y - gap
data_h = data_bottom - data_y

# 左列=3枚, 中列=2枚, 右列=2枚
ph3 = (data_h - 2*gap) / 3
ph2 = (data_h - gap) / 2

# ---- 左列: 合成 → 自己修復性 → シアシニング ----
panel(dx0, data_y, colw, ph3, "① PGD G4 の精密合成〔新〕",
      ["・Divergent法(全8段階)で単分散PGD G4を合成",
       "・撹拌強化+反応時間2倍で高世代を最適化",
       "・収率90%, MALDI単一ピーク[M+Na]+=3488"],
      C_MAT, bar="#c9962a", fig="図: MALDI(最適化前後) / 収率表", fig_h=ph3-78)
panel(dx0, data_y+ph3+gap, colw, ph3, "② 自己修復性 (レオロジー)〔目玉〕",
      ["・せん断除去で G' が即座に回復",
       "・元の強固なゲル構造へ再構築",
       "・複数サイクルで再現性よく回復"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: ステップ歪みサイクル (G'/G'')", fig_h=ph3-78)
panel(dx0, data_y+2*(ph3+gap), colw, ph3, "③ シアシニング性 (レオロジー)",
      ["・高せん断でゲルが流動化(ゾル化)",
       "・粘度が急低下 → 注射針を通過可能",
       "・大歪500%で G'<G''"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: せん断速度-粘度 / 大歪での G'低下", fig_h=ph3-78)

# ---- 中列: In vitro 安全性 → 分解性 ----
panel(dx1, data_y, colw, ph2, "④ In vitro 安全性｜Live/Dead〔新〕",
      ["・D1細胞を3次元ゲル内に包埋し培養",
       "・1/5/9日とも死細胞はほぼ観察されず",
       "・溶解後も毒性なし",
       "・→ 化学架橋剤なしで高い細胞生存率"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: Live/Dead 蛍光像 (PG3-1 / PG4-025 / PG4-0125 ×1・5・9日)", fig_h=ph2-92)
panel(dx1, data_y+ph2+gap, colw, ph2, "⑤ In vitro 分解性〔新〕",
      ["・37℃ PBS中の重量変化で分解を評価",
       "・PG3-1:3日 / PG4-0125:4日 / PG4-025:25日",
       "・世代・混合比で滞留期間を制御",
       "・→ 数日〜数週間で任意に制御可能"],
      C_VITRO, bar=C_VITRO_BAR, fig="図: 分解曲線 (重量比 vs 時間)", fig_h=ph2-92)

# ---- 右列: 注入試験(In vitro & In vivo) → In vivo 分解試験(H&E) ----
panel(dx2, data_y, colw, ph2, "⑥ 注入試験 (In vitro & In vivo)〔新〕",
      ["・in vitro: 赤色化ゲルをシリンジ吐出→即ゲル化",
       "・in vivo: マウス背部皮下へ注入(0.25 mL)",
       "・抵抗なく注入→局所にドーム状デポを形成",
       "・香港大 Sang-Jin Lee先生と共同研究"],
      C_VIVO, bar=C_VIVO_BAR, fig="図: in vitro 注入写真 / in vivo マウス皮下デポ写真", fig_h=ph2-92)
panel(dx2, data_y+ph2+gap, colw, ph2, "⑦ In vivo 分解試験 (H&E染色)〔要データ〕",
      ["・皮下デポの経時的な分解・吸収を評価",
       "・H&E染色で組織反応・分解挙動を観察",
       "・(生分解性・生体適合性を組織学的に確認)",
       "※元データは別途提供をお願いします"],
      C_VIVO, bar=C_VIVO_BAR, fig="図: H&E染色 組織像 (経時)", fig_h=ph2-92)

# ===== Conclusion(最下段・全幅・薄い帯)=====
rect(cx, conc_y, cw, band_conc_h, C_CONC, stroke="#2d6a4f", rx=6, sw=1.2)
text(cx+10, conc_y+18, "Conclusion:", size=11, color="#1d3a2a", weight="bold")
text(cx+95, conc_y+18,
     "単分散PGD G4を確立 → 水素結合で超分子ゲル形成（自己修復・分解性を制御, 高生体適合性）"
     " → in vivoで局所デポを形成し実用性を実証",
     size=9.3, color="#1d3a2a")
text(cx+10, conc_y+38, "Future:", size=10, color="#33506b", weight="bold")
text(cx+62, conc_y+38,
     "PTX徐放定量(HPLC) / 担がんマウスで抗腫瘍効果 / PGD G5合成 / DSCで中間水評価　　"
     "Ref: Ooya Gels 2022 ; Cho&Ooya 2018 ; Yamazaki Langmuir 2021 ほか",
     size=8.5, color="#3a4a55")

add('</svg>')
with open("poster/poster_mockup.svg","w",encoding="utf-8") as f:
    f.write("\n".join(svg))
print("SVG生成完了")
