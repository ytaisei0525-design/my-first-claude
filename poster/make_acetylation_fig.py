#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GCアセチル化度(DA)とゲル化時間の関係グラフ(見やすい版)。
   ゲル化時間が1.5〜80hと広範囲のため、対数軸版と線形(分割軸)版を作成する。"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
import numpy as np

# 日本語フォント設定(IPAGothic)
fp = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
font_manager.fontManager.addfont(fp)
plt.rcParams["font.family"] = font_manager.FontProperties(fname=fp).get_name()
plt.rcParams["axes.unicode_minus"] = False

# ---- データ(Table 3-2) ----
labels = ["Start(GC)", "30%", "50%", "70%", "90%"]
DA   = [16.7, 40.0, 60.0, 76.7, 90.0]   # アセチル化度 (%)
tmin = [90, 110, 120, 900, 4800]        # ゲル化時間 (min)
th   = [t/60 for t in tmin]             # ゲル化時間 (hours)

ACCENT = "#2563eb"
MARK   = "#0d3b66"
HILITE = "#ea580c"

# =========================================================
# 案1: 対数軸(1枚で全体の傾向が最も見やすい)
# =========================================================
fig, ax = plt.subplots(figsize=(6.4, 4.6), dpi=200)
ax.plot(DA, th, "-o", color=ACCENT, markerfacecolor=MARK, markeredgecolor="white",
        markersize=10, linewidth=2.5, zorder=3)
ax.set_yscale("log")
ax.set_ylim(1, 150)
ax.set_xlim(10, 96)

# 値ラベル
for x, y, t in zip(DA, th, th):
    lbl = f"{y:.1f} h" if y >= 1 else f"{y:.2f} h"
    ax.annotate(lbl, (x, y), textcoords="offset points", xytext=(0, 12),
                ha="center", fontsize=11, fontweight="bold", color=MARK)

# しきい値域を強調(DA>60%で急増)
ax.axvspan(60, 96, color=HILITE, alpha=0.07, zorder=0)
ax.annotate("DA>60%で急増\n(遊離アミノ基が減少)", (83, 38), ha="center",
            fontsize=10.5, color=HILITE, fontweight="bold")

ax.set_xlabel("アセチル化度 DA (%)", fontsize=13, fontweight="bold")
ax.set_ylabel("ゲル化時間 (hours, 対数)", fontsize=13, fontweight="bold")
ax.set_title("GCアセチル化度とゲル化時間の関係", fontsize=15, fontweight="bold", pad=12)
ax.grid(True, which="both", axis="y", ls="--", alpha=0.4)
ax.grid(True, which="major", axis="x", ls="--", alpha=0.25)
ax.tick_params(labelsize=11)
fig.text(0.5, -0.02,
         "→ アミノ基をアセチル化するほどゲル化が遅延。アミノ基が主要な架橋点であることを示す。",
         ha="center", fontsize=10, color="#444")
fig.tight_layout()
fig.savefig("poster/fig_acetylation_log.png", bbox_inches="tight", facecolor="white")
plt.close(fig)

# =========================================================
# 案2: 線形・分割軸(低時間域も潰れず、急増も見える)
# =========================================================
fig, (axt, axb) = plt.subplots(2, 1, sharex=True, figsize=(6.4, 5.0), dpi=200,
                               gridspec_kw={"height_ratios": [2, 1.2], "hspace": 0.08})
for ax in (axt, axb):
    ax.plot(DA, th, "-o", color=ACCENT, markerfacecolor=MARK, markeredgecolor="white",
            markersize=10, linewidth=2.5, zorder=3)
    ax.axvspan(60, 96, color=HILITE, alpha=0.07, zorder=0)
    ax.grid(True, ls="--", alpha=0.35)

# 上段: 高時間域(15〜80h)
axt.set_ylim(8, 90)
# 下段: 低時間域(1〜3h)
axb.set_ylim(0, 3)
axb.set_xlim(10, 96)

# 分割の波線
d = .012
kw = dict(transform=axt.transAxes, color="gray", clip_on=False, lw=1)
axt.plot((-d, +d), (-d, +d), **kw); axt.plot((1-d, 1+d), (-d, +d), **kw)
kw = dict(transform=axb.transAxes, color="gray", clip_on=False, lw=1)
axb.plot((-d, +d), (1-d*2, 1+d*2), **kw); axb.plot((1-d, 1+d), (1-d*2, 1+d*2), **kw)
axt.spines["bottom"].set_visible(False)
axb.spines["top"].set_visible(False)
axt.tick_params(labelbottom=False)

for x, y in zip(DA, th):
    ax = axt if y >= 8 else axb
    lbl = f"{y:.1f} h" if y >= 1 else f"{y:.2f} h"
    ax.annotate(lbl, (x, y), textcoords="offset points", xytext=(8, 6),
                ha="left", fontsize=11, fontweight="bold", color=MARK)

axb.set_xlabel("アセチル化度 DA (%)", fontsize=13, fontweight="bold")
fig.text(0.02, 0.5, "ゲル化時間 (hours)", va="center", rotation="vertical",
         fontsize=13, fontweight="bold")
axt.set_title("GCアセチル化度とゲル化時間の関係", fontsize=15, fontweight="bold", pad=12)
axt.annotate("DA>60%で急増", (84, 60), ha="center", fontsize=10.5,
             color=HILITE, fontweight="bold")
for ax in (axt, axb):
    ax.tick_params(labelsize=11)
fig.savefig("poster/fig_acetylation_split.png", bbox_inches="tight", facecolor="white")
plt.close(fig)

print("生成: poster/fig_acetylation_log.png, poster/fig_acetylation_split.png")
