#!/usr/bin/env python3
"""Materials Overview slides (Slide 9 & 10) generated independently.
Reduced whitespace, denser cards. Reuses helpers/colors/sizes from build_pptx_en.py."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from build_pptx_en import (
    rect, rrect, text, draw_header, draw_footer, new_slide,
    W, H, PAD, HDR_H, FTR_H, BODY_Y, BODY_H, BODY_X, BODY_W,
    INK, INK2, INK3, MUTED, LINE, SOFT, CARD_C, HEADER, ACCENT,
    AMBER, GOLD, WHITE, D_BLUE, D_TEAL, D_AMBR, D_GRN, D_RED, GRAY_L,
)


# ===== SLIDE 9: Material Composition (denser, less whitespace) =====
def slide_mat_composition(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials Overview 1/2", "Material Composition: Sustainable & High-Performance")
    draw_footer(sl, "9 / 39")

    # Lead sentence (uses the gap above the cards)
    lead_h = Inches(0.50)
    text(sl, BODY_X, BODY_Y, BODY_W, lead_h,
         "Three building blocks — skeleton, function, and bonding — combine biodegradability with high performance.",
         size=14, color=INK2, anchor=MSO_ANCHOR.MIDDLE)

    n = 3
    gap = Inches(0.30)
    cw = (BODY_W - gap * (n - 1)) / n
    cards_y = BODY_Y + lead_h + Inches(0.14)
    card_h = H - cards_y - FTR_H - Inches(0.24)

    items = [
        ("🌿", "Cellulose Derivatives", "Plant-based biopolymers",
         "Plant-derived biopolymers that form the backbone of the gel network.",
         [("3 types", "HEC, MC, MHEC"),
          ("Bio.", "Highly biodegradable"),
          ("Role", "Viscosity & structure")],
         ACCENT),
        ("◎", "Colloidal Silica (CSP)", "Colloidal Silica Particles",
         "The key component that interacts dynamically with polymers and forms aerogel under heat.",
         [("Size", "~22 nm monodisperse"),
          ("Sinter", "Forms silica skeleton"),
          ("Role", "Source of aerogel")],
         D_TEAL),
        ("⬡", "PP Interaction", "Polymer–Particle Interaction",
         "Builds the network via dynamic, multivalent hydrogen bonds — no covalent bonds required.",
         [("Bonds", "Dynamic H-bonds"),
          ("Self-heal", "Gel recovers after strain"),
          ("Sprayable", "Easy spray application")],
         D_AMBR),
    ]

    for i, (icon, title, en, body, keys, accent_c) in enumerate(items):
        cx = BODY_X + i * (cw + gap)
        rrect(sl, cx, cards_y, cw, card_h, CARD_C, LINE, 0.5, radius=0.035)
        rect(sl, cx, cards_y, cw, Inches(0.06), accent_c)

        pad_x = Inches(0.24)
        inner_w = cw - pad_x * 2

        # Icon
        icon_y = cards_y + Inches(0.30)
        text(sl, cx, icon_y, cw, Inches(0.56),
             icon, size=30, color=accent_c,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Title + sub
        title_y = icon_y + Inches(0.62)
        text(sl, cx + pad_x, title_y, inner_w, Inches(0.42),
             title, size=16, bold=True, color=INK, align=PP_ALIGN.CENTER)
        text(sl, cx + pad_x, title_y + Inches(0.42), inner_w, Inches(0.28),
             en, size=10, italic=True, color=MUTED, align=PP_ALIGN.CENTER)

        # Divider
        div_y = title_y + Inches(0.78)
        rect(sl, cx + Inches(0.32), div_y, cw - Inches(0.64), Inches(0.015), LINE)

        # Body
        body_y = div_y + Inches(0.14)
        text(sl, cx + pad_x, body_y, inner_w, Inches(1.10),
             body, size=12, color=INK3, align=PP_ALIGN.CENTER)

        # Key-point rows (fill the lower whitespace)
        ky = body_y + Inches(1.22)
        row_h = Inches(0.74)
        for j, (label, detail) in enumerate(keys):
            ry = ky + j * row_h
            chip_w = Inches(1.05)
            rrect(sl, cx + pad_x, ry + Inches(0.04), chip_w, Inches(0.32),
                  WHITE, accent_c, 0.7, radius=0.5)
            text(sl, cx + pad_x, ry + Inches(0.04), chip_w, Inches(0.32),
                 label, size=10, bold=True, color=accent_c,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            text(sl, cx + pad_x + chip_w + Inches(0.10), ry,
                 inner_w - chip_w - Inches(0.10), Inches(0.40),
                 detail, size=11, color=INK2, anchor=MSO_ANCHOR.MIDDLE)


# ===== SLIDE 10: Spray / Adhesion / Flame mechanism (denser) =====
def slide_self_protection(prs):
    sl = new_slide(prs)
    draw_header(sl, "Materials Overview 2/2", "Mechanism: Spray, Adhesion, and Flame Protection")
    draw_footer(sl, "10 / 39")

    # Lead sentence
    lead_h = Inches(0.52)
    text(sl, BODY_X, BODY_Y, BODY_W, lead_h,
         "Not temperature, but shear and self-healing govern the spray-to-adhesion behavior.",
         size=15, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)

    # Four large process cards (full width)
    n = 4
    arrow_w = Inches(0.30)
    cards_y = BODY_Y + lead_h + Inches(0.14)
    cw = (BODY_W - arrow_w * (n - 1)) / n
    card_h = Inches(3.50)

    stages = [
        ("01", "At rest", "Gel", "High-viscosity elastic gel. G' > G'' holds structure.",
         "Yield stress ~33 Pa", ACCENT),
        ("02", "Spraying", "Sol-like", "Shear strain lowers viscosity; the gel flows.",
         "Shear-thinning n < 1", D_TEAL),
        ("03", "On wall", "Self-healing gel", "Strain release instantly restores the gel network.",
         "G' recovery ~90%", D_GRN),
        ("04", "Flame", "Aerogel", "CSP sinters into a porous silica layer.",
         "Acts as thermal barrier", D_AMBR),
    ]

    for i, (no, phase, state, desc, metric, c) in enumerate(stages):
        cx = BODY_X + i * (cw + arrow_w)
        rrect(sl, cx, cards_y, cw, card_h, CARD_C, LINE, 0.5, radius=0.04)
        rect(sl, cx, cards_y, cw, Inches(0.06), c)

        pad_x = Inches(0.20)
        inner_w = cw - pad_x * 2

        text(sl, cx + pad_x, cards_y + Inches(0.18), inner_w, Inches(0.34),
             no, size=13, bold=True, color=c)
        text(sl, cx + pad_x, cards_y + Inches(0.54), inner_w, Inches(0.30),
             phase, size=12, bold=True, color=MUTED)
        text(sl, cx + pad_x, cards_y + Inches(0.92), inner_w, Inches(0.56),
             state, size=20, bold=True, color=INK)
        rect(sl, cx + pad_x, cards_y + Inches(1.58), inner_w, Inches(0.015), LINE)
        text(sl, cx + pad_x, cards_y + Inches(1.72), inner_w, Inches(1.10),
             desc, size=12, color=INK3)

        badge_h = Inches(0.46)
        badge_y = cards_y + card_h - badge_h - Inches(0.18)
        rrect(sl, cx + pad_x, badge_y, inner_w, badge_h, WHITE, c, 0.7, radius=0.06)
        text(sl, cx + pad_x, badge_y, inner_w, badge_h,
             metric, size=10, bold=True, color=c,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        if i < n - 1:
            text(sl, cx + cw, cards_y, arrow_w, card_h,
                 "→", size=20, color=GRAY_L,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Bottom POINT band (fills remaining whitespace)
    band_y = cards_y + card_h + Inches(0.20)
    band_h = H - band_y - FTR_H - Inches(0.22)
    rrect(sl, BODY_X, band_y, BODY_W, band_h, HEADER, None, radius=0.05)
    rect(sl, BODY_X, band_y, Inches(0.08), band_h, GOLD)
    text(sl, BODY_X + Inches(0.30), band_y, Inches(2.4), band_h,
         "POINT", size=13, bold=True, color=GOLD, anchor=MSO_ANCHOR.MIDDLE)
    text(sl, BODY_X + Inches(2.0), band_y, BODY_W - Inches(2.3), band_h,
         "Rather than relying on MC's thermal response (LCST), rheology alone achieves \"spray → stick → set\" — the core of this material.",
         size=13, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)


def main():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    print("Building Materials Overview slides (Slide 9 & 10)...")
    slide_mat_composition(prs)
    slide_self_protection(prs)

    out_path = '/home/user/my-first-claude/overview_slides_en.pptx'
    prs.save(out_path)
    print(f"Done: {out_path}")
    print(f"Slide count: {len(prs.slides.__iter__.__self__._sldIdLst)}")


if __name__ == '__main__':
    main()
