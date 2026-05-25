#!/usr/bin/env python3
"""Build slides2.pptx — Wang et al. 2026 (Biomaterials 331, 124114)
   Co-assembly of zwitterionic and cationic polymers for antibacterial
   and antithrombotic surfaces via weak electrostatic interactions
   Design: Clean & Minimal (white bg, teal accent, thin borders)
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ── Colors ───────────────────────────────────────────────────
WHITE   = RGBColor(0xff, 0xff, 0xff)
CARD    = RGBColor(0xf7, 0xf9, 0xfc)
CARD2   = RGBColor(0xed, 0xf2, 0xf7)
BORDER  = RGBColor(0xd4, 0xdb, 0xe4)
TEXT    = RGBColor(0x1e, 0x25, 0x2e)
TEXT2   = RGBColor(0x44, 0x4e, 0x5a)
MUTED   = RGBColor(0x82, 0x8e, 0x9c)
TEAL    = RGBColor(0x07, 0x70, 0x90)
TEAL_L  = RGBColor(0xdd, 0xf0, 0xf7)
TEAL_D  = RGBColor(0x05, 0x50, 0x68)
GREEN   = RGBColor(0x1a, 0x88, 0x55)
GREEN_L = RGBColor(0xd4, 0xf0, 0xe2)
RED     = RGBColor(0xb5, 0x28, 0x28)
RED_L   = RGBColor(0xf9, 0xe0, 0xe0)
AMBER   = RGBColor(0xc2, 0x65, 0x05)
AMBER_L = RGBColor(0xfb, 0xf0, 0xd8)
SLATE   = RGBColor(0x3a, 0x52, 0x68)
GRAY_L  = RGBColor(0xf0, 0xf2, 0xf5)
GRAY_M  = RGBColor(0xa8, 0xb4, 0xc0)

# ── Dimensions ───────────────────────────────────────────────
W      = Inches(13.333)
H      = Inches(7.5)
PAD    = Inches(0.42)
HDR_H  = Inches(0.80)
FTR_H  = Inches(0.30)
BODY_Y = HDR_H + Inches(0.24)
BODY_H = H - BODY_Y - FTR_H - Inches(0.10)
BODY_X = PAD
BODY_W = W - PAD * 2

FONT_SCALE = 1.12
FONT_NAME  = 'Arial'


# ── Primitives ───────────────────────────────────────────────
def rect(sl, x, y, w, h, fill, border=None, bpt=0.5):
    s = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if border:
        s.line.color.rgb = border; s.line.width = Pt(bpt)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s

def rrect(sl, x, y, w, h, fill, border=None, bpt=0.5, radius=0.04):
    s = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    s.adjustments[0] = radius
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if border:
        s.line.color.rgb = border; s.line.width = Pt(bpt)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s

def text(sl, x, y, w, h, content, size=11, bold=False,
         color=TEXT2, align=PP_ALIGN.LEFT, italic=False,
         anchor=MSO_ANCHOR.TOP):
    if content is None: return None
    s = str(content).strip()
    if not s: return None
    tb = sl.shapes.add_textbox(x, y, w, h)
    tb.word_wrap = True
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Emu(36000); tf.margin_right  = Emu(36000)
    tf.margin_top  = Emu(18000); tf.margin_bottom = Emu(18000)
    for i, line in enumerate(s.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.size  = Pt(size * FONT_SCALE)
        run.font.bold  = bold
        run.font.italic = italic
        run.font.color.rgb = color
        run.font.name  = FONT_NAME
    return tb


# ── Layout helpers ───────────────────────────────────────────
def draw_header(sl, section, title):
    rect(sl, 0, 0, W, HDR_H, WHITE)
    rect(sl, 0, 0, Inches(0.07), HDR_H, TEAL)
    rect(sl, 0, HDR_H - Inches(0.02), W, Inches(0.02), BORDER)
    if section:
        text(sl, Inches(0.20), Inches(0.12), Inches(6), Inches(0.26),
             section.upper(), size=9, bold=True, color=TEAL)
    if title:
        text(sl, Inches(0.20), Inches(0.38), BODY_W, Inches(0.36),
             title, size=19, bold=True, color=TEXT)

def draw_footer(sl, num):
    rect(sl, 0, H - FTR_H, W, FTR_H, GRAY_L)
    rect(sl, 0, H - FTR_H, W, Inches(0.02), BORDER)
    text(sl, PAD, H - FTR_H + Inches(0.04), Inches(8), Inches(0.22),
         "Wang et al. | Biomaterials 331 (2026) 124114", size=8, color=MUTED)
    if num:
        text(sl, W - Inches(1.5), H - FTR_H + Inches(0.04),
             Inches(1.3), Inches(0.22), str(num), size=8, color=MUTED,
             align=PP_ALIGN.RIGHT)

def slide(prs, section="", title="", num=None):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    draw_header(sl, section, title)
    draw_footer(sl, num)
    return sl

def card(sl, x, y, w, h, title=None, body=None, bullets=None,
         title_size=13, body_size=11, accent=None):
    rrect(sl, x, y, w, h, CARD, BORDER, 0.5)
    if accent:
        rect(sl, x, y, Inches(0.05), h, accent)
    cy = y + Inches(0.16)
    px = x + Inches(0.18)
    cw = w - Inches(0.36)
    if title:
        text(sl, px, cy, cw, Inches(0.34), title,
             size=title_size, bold=True, color=TEXT)
        cy += Inches(0.36)
    if body:
        text(sl, px, cy, cw, h - (cy - y) - Inches(0.10),
             body, size=body_size, color=TEXT2)
    elif bullets:
        for b in bullets:
            if cy + Inches(0.26) > y + h - Inches(0.06):
                break
            text(sl, px, cy, Inches(0.18), Inches(0.26),
                 "·", size=body_size, bold=True, color=TEAL)
            text(sl, px + Inches(0.20), cy, cw - Inches(0.20), Inches(0.26),
                 b, size=body_size, color=TEXT2)
            cy += Inches(0.27)

def stat_box(sl, x, y, w, h, value, unit="", label="", val_color=TEAL):
    rrect(sl, x, y, w, h, CARD, BORDER, 0.5)
    text(sl, x + Inches(0.10), y + Inches(0.14), w - Inches(0.20),
         Inches(0.48), value, size=30, bold=True, color=val_color,
         align=PP_ALIGN.CENTER)
    if unit:
        text(sl, x + Inches(0.10), y + Inches(0.60), w - Inches(0.20),
             Inches(0.22), unit, size=9, color=MUTED, align=PP_ALIGN.CENTER)
    if label:
        text(sl, x + Inches(0.10), y + h - Inches(0.30), w - Inches(0.20),
             Inches(0.26), label, size=9, color=TEXT2, align=PP_ALIGN.CENTER,
             bold=True)

def fig_box(sl, x, y, w, h, caption=""):
    rrect(sl, x, y, w, h, GRAY_L, BORDER, 0.5)
    text(sl, x, y + h * 0.3, w, h * 0.4,
         "[Figure Placeholder]", size=10, color=MUTED, align=PP_ALIGN.CENTER)
    if caption:
        text(sl, x + Inches(0.10), y + h - Inches(0.28),
             w - Inches(0.20), Inches(0.24), caption,
             size=8, color=MUTED, align=PP_ALIGN.CENTER, italic=True)

def label(sl, x, y, w, h, txt, bg=TEAL_L, fg=TEAL_D, size=9):
    rrect(sl, x, y, w, h, bg, None, radius=0.08)
    text(sl, x, y, w, h, txt, size=size, bold=True, color=fg,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

def bar_row(sl, x, y, w, h, label_txt, pct, bar_color=TEAL, bg=CARD,
            label_w=Inches(1.8), val_w=Inches(0.7)):
    bar_w = w - label_w - val_w - Inches(0.10)
    text(sl, x, y, label_w, h, label_txt, size=10, color=TEXT2,
         anchor=MSO_ANCHOR.MIDDLE)
    rect(sl, x + label_w + Inches(0.05), y + h * 0.25,
         bar_w, h * 0.50, BORDER)
    fill_w = max(Inches(0.02), bar_w * pct / 100)
    rect(sl, x + label_w + Inches(0.05), y + h * 0.25,
         fill_w, h * 0.50, bar_color)
    text(sl, x + label_w + bar_w + Inches(0.12), y,
         val_w, h, f"{pct:.1f}%", size=10, bold=True, color=bar_color,
         anchor=MSO_ANCHOR.MIDDLE)

def divider_slide(prs, num, section_num, section_title, subtitle=""):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    rect(sl, 0, 0, Inches(0.55), H, TEAL)
    rect(sl, Inches(0.55), 0, Inches(0.03), H, TEAL_L)
    draw_footer(sl, num)
    text(sl, Inches(0.90), Inches(2.60), Inches(10), Inches(0.40),
         section_num.upper(), size=10, bold=True, color=TEAL)
    text(sl, Inches(0.90), Inches(3.05), Inches(10), Inches(0.80),
         section_title, size=34, bold=True, color=TEXT)
    if subtitle:
        text(sl, Inches(0.90), Inches(3.90), Inches(9), Inches(0.40),
             subtitle, size=13, color=TEXT2)
    return sl


# ════════════════════════════════════════════════════════════════
# BUILD
# ════════════════════════════════════════════════════════════════
prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

# ── Slide 1: Cover ───────────────────────────────────────────
sl = prs.slides.add_slide(prs.slide_layouts[6])
rect(sl, 0, 0, W, H, WHITE)
rect(sl, 0, 0, Inches(0.10), H, TEAL)
rect(sl, 0, H - Inches(0.06), W, Inches(0.06), TEAL)
rect(sl, 0, H - FTR_H, W, Inches(0.02), BORDER)

# Teal accent block
rect(sl, Inches(0.22), Inches(1.80), Inches(9.0), Inches(0.06), TEAL)
text(sl, Inches(0.22), Inches(1.95), Inches(11.5), Inches(1.60),
     "Co-assembly of Zwitterionic and Cationic Polymers\n"
     "for Antibacterial and Antithrombotic Surfaces\n"
     "via Weak Electrostatic Interactions",
     size=28, bold=True, color=TEXT)
text(sl, Inches(0.22), Inches(3.70), Inches(10), Inches(0.36),
     "Wang et al. | Biomaterials 331 (2026) 124114", size=13, color=TEXT2)
text(sl, Inches(0.22), Inches(4.14), Inches(10), Inches(0.30),
     "LbL coating  ·  Zwitterionic polymers  ·  Antibacterial  ·  Antithrombotic",
     size=11, color=MUTED)

for i, (kw, desc) in enumerate([
    ("PA2.5", "Optimized dual-function coating"),
    (">97.6%", "Bactericidal efficiency vs S. aureus"),
    ("3.1%",   "Thrombus coverage (vs PLA 99.0%)"),
]):
    bx = Inches(0.22) + i * Inches(3.0)
    rrect(sl, bx, Inches(5.10), Inches(2.7), Inches(1.00), CARD, BORDER)
    text(sl, bx + Inches(0.15), Inches(5.22), Inches(2.4), Inches(0.38),
         kw, size=22, bold=True, color=TEAL)
    text(sl, bx + Inches(0.15), Inches(5.60), Inches(2.4), Inches(0.36),
         desc, size=9, color=TEXT2)

# ── Slide 2: Table of Contents ───────────────────────────────
sl = slide(prs, "", "Table of Contents", 2)
sections = [
    ("01", "Background & Motivation",       "Clinical need for dual-function medical device coatings"),
    ("02", "Materials & Methods",            "Polymer synthesis, LbL coating, characterization toolkit"),
    ("03", "Solution Interactions",          "ITC, turbidity, fluorescence — weak co-assembly confirmed"),
    ("04", "Coating Properties",             "QCM-d, AFM, XPS, WCA, zeta potential"),
    ("05", "Antifouling & Antithrombotic",   "Ex vivo blood contact — PA2.5: 3.1% thrombus coverage"),
    ("06", "Antibacterial Performance",      "Contact sterilization — PA2.5: >97.6% bactericidal"),
    ("07", "Biocompatibility & In Vivo",     "Hemolysis, cytotox, rat/rabbit animal models"),
]
cols = [Inches(0.50), Inches(6.80)]
col_w = Inches(6.0)
for i, (num, sec, sub) in enumerate(sections):
    row = i % 4
    col = i // 4
    bx = cols[col]
    by = BODY_Y + Inches(0.10) + row * Inches(1.35)
    rrect(sl, bx, by, col_w, Inches(1.22), CARD, BORDER)
    rect(sl, bx, by, Inches(0.06), Inches(1.22), TEAL)
    label(sl, bx + Inches(0.14), by + Inches(0.28), Inches(0.44), Inches(0.28),
          num, bg=TEAL_L, fg=TEAL_D, size=8)
    text(sl, bx + Inches(0.66), by + Inches(0.20), col_w - Inches(0.80),
         Inches(0.36), sec, size=13, bold=True, color=TEXT)
    text(sl, bx + Inches(0.66), by + Inches(0.60), col_w - Inches(0.80),
         Inches(0.50), sub, size=9, color=TEXT2)

# ── Slide 3: Section Divider — Background ────────────────────
divider_slide(prs, 3, "Section 01", "Background & Motivation",
              "Clinical need for dual-function medical device coatings")

# ── Slide 4: Clinical Problem ─────────────────────────────────
sl = slide(prs, "01 · Background", "The Dual Challenge: Infection and Thrombosis", 4)
problems = [
    ("Bacterial Infection",
     "Medical device-associated infections account for ~2 million cases per year in the USA alone. "
     "Biofilm formation on implant surfaces leads to antibiotic resistance and implant failure."),
    ("Thrombosis",
     "Thrombus formation on implanted materials triggers device failure, embolic events, and "
     "systemic complications. Platelet adhesion and activation occur within seconds of blood contact."),
    ("Antibiotic Resistance",
     "Conventional antibiotic coatings provide only short-term protection and accelerate resistance. "
     "Contact-killing surfaces offer a sustained, resistance-free antibacterial mechanism."),
    ("Conflicting Strategies",
     "Cationic surfaces kill bacteria but promote thrombosis. Zwitterionic surfaces are antifouling "
     "but lack intrinsic antibacterial activity. Achieving both simultaneously remains a key challenge."),
]
cw = (BODY_W - Inches(0.30)) / 2
for i, (ttl, body) in enumerate(problems):
    col = i % 2; row = i // 2
    bx = BODY_X + col * (cw + Inches(0.30))
    by = BODY_Y + row * Inches(2.40) + Inches(0.10)
    card(sl, bx, by, cw, Inches(2.20), title=ttl, body=body,
         title_size=13, body_size=10, accent=TEAL)

# ── Slide 5: Current Approaches & Limitations ─────────────────
sl = slide(prs, "01 · Background", "Current Approaches and Their Limitations", 5)
approaches = [
    ("Zwitterionic Coatings",
     ["Excellent antifouling & antithrombotic via hydration layer",
      "No inherent antibacterial activity",
      "Polysulfobetaine, polycarboxybetaine well-studied"], TEAL),
    ("Cationic (QA) Coatings",
     ["Membrane disruption kills bacteria on contact",
      "Cationic charge promotes platelet activation",
      "Antithrombotic performance is poor"], RED),
    ("Antibiotic-Releasing Coatings",
     ["Short-term protection only",
      "Drives antibiotic resistance",
      "Drug depletion limits long-term use"], AMBER),
    ("Dual-Function Attempts",
     ["Sequential layers sacrifice one function",
      "Blending disrupts both functional groups",
      "No prior success via co-assembly approach"], SLATE),
]
cw = (BODY_W - Inches(0.30)) / 2
for i, (ttl, buls, ac) in enumerate(approaches):
    col = i % 2; row = i // 2
    bx = BODY_X + col * (cw + Inches(0.30))
    by = BODY_Y + row * Inches(2.40) + Inches(0.10)
    rrect(sl, bx, by, cw, Inches(2.20), CARD, BORDER)
    rect(sl, bx, by, Inches(0.06), Inches(2.20), ac)
    text(sl, bx + Inches(0.18), by + Inches(0.16), cw - Inches(0.30),
         Inches(0.32), ttl, size=13, bold=True, color=TEXT)
    cy = by + Inches(0.54)
    for b in buls:
        text(sl, bx + Inches(0.18), cy, Inches(0.18), Inches(0.26),
             "·", size=10, bold=True, color=ac)
        text(sl, bx + Inches(0.38), cy, cw - Inches(0.52), Inches(0.26),
             b, size=10, color=TEXT2)
        cy += Inches(0.28)

# ── Slide 6: Design Concept ───────────────────────────────────
sl = slide(prs, "01 · Background", "Design Concept: Weak Co-assembly for Dual Function", 6)
rect(sl, BODY_X, BODY_Y + Inches(0.10), BODY_W, Inches(1.10), TEAL_L)
rect(sl, BODY_X, BODY_Y + Inches(0.10), BODY_W, Inches(0.03), TEAL)
text(sl, BODY_X + Inches(0.30), BODY_Y + Inches(0.20), BODY_W - Inches(0.60),
     Inches(0.80),
     "Key Insight: When zwitterionic and cationic polymers co-assemble via weak electrostatic "
     "interactions, both functional groups remain accessible — enabling simultaneous antifouling "
     "and antibacterial activity from a single coating layer.",
     size=12, color=TEAL_D)

components = [
    ("PA (Poly-carboxybetaine amide)",
     "Zwitterionic polymer\nCOO⁻ and N⁺ groups\nStrong hydration layer → antifouling\nWeak interaction with PQ",
     TEAL),
    ("PS (Poly-sulfobetaine)",
     "Zwitterionic polymer\nSO₃⁻ and N⁺ groups\nExcellent antifouling properties\nDifferent interaction mode with PQ",
     SLATE),
    ("PQ (Poly-quaternary ammonium)",
     "Cationic polymer\nPermanent positive charge\nMembrane disruption → antibacterial\nAccessible when interaction is weak",
     AMBER),
    ("Co-assembly Result",
     "PA-PQ or PS-PQ complex\nLbL coating on PLA substrate\nRetains both functional groups\nDual antibacterial + antithrombotic",
     GREEN),
]
cw = (BODY_W - Inches(0.45)) / 4
by2 = BODY_Y + Inches(1.35)
for i, (ttl, body, ac) in enumerate(components):
    bx = BODY_X + i * (cw + Inches(0.15))
    rrect(sl, bx, by2, cw, Inches(4.50), CARD, BORDER)
    rect(sl, bx, by2, cw, Inches(0.06), ac)
    text(sl, bx + Inches(0.14), by2 + Inches(0.18), cw - Inches(0.28),
         Inches(0.40), ttl, size=11, bold=True, color=TEXT)
    text(sl, bx + Inches(0.14), by2 + Inches(0.62), cw - Inches(0.28),
         Inches(3.70), body, size=10, color=TEXT2)

# ── Slide 7: Research Overview ────────────────────────────────
sl = slide(prs, "01 · Background", "Research Overview", 7)
overview = [
    ("Hypothesis",
     "Weak electrostatic interaction between zwitterionic (PA/PS) and cationic (PQ) polymers "
     "preserves functional group accessibility, enabling dual-function LbL coatings.",
     TEAL),
    ("Solution-Phase Characterization",
     "ITC, fluorescence spectroscopy, turbidimetry, and viscometry confirm spontaneous "
     "co-assembly (ΔG < 0) with weak interaction for PA-PQ pair.",
     SLATE),
    ("Surface Characterization",
     "QCM-d, AFM, XPS, WCA, and zeta potential quantify coating deposition, morphology, "
     "elemental composition, and hydrophilicity as a function of formulation.",
     AMBER),
    ("In Vitro Performance",
     "Ex vivo blood contact tests, MIC assays, and CFU counting evaluate antithrombotic "
     "and antibacterial performance. PA2.5 achieves 3.1% thrombus coverage and >97.6% kill.",
     GREEN),
    ("In Vivo Validation",
     "Rat subcutaneous, rabbit jugular vein, and PGLA suture implant models confirm "
     "safety and efficacy. PA2.5 achieves >99% bacterial eradication in vivo.",
     RED),
]
cw = (BODY_W - Inches(0.15)) / 3
rows = [overview[:3], overview[3:]]
for ri, row in enumerate(rows):
    by2 = BODY_Y + ri * Inches(3.05) + Inches(0.10)
    bh = Inches(2.85)
    for ci, (ttl, body, ac) in enumerate(row):
        bx = BODY_X + ci * (cw + Inches(0.075))
        rrect(sl, bx, by2, cw, bh, CARD, BORDER)
        rect(sl, bx, by2, cw, Inches(0.05), ac)
        text(sl, bx + Inches(0.16), by2 + Inches(0.16), cw - Inches(0.30),
             Inches(0.32), ttl, size=12, bold=True, color=TEXT)
        text(sl, bx + Inches(0.16), by2 + Inches(0.54), cw - Inches(0.30),
             bh - Inches(0.68), body, size=10, color=TEXT2)

# ── Slide 8: Section Divider — Methods ───────────────────────
divider_slide(prs, 8, "Section 02", "Materials & Methods",
              "Polymer synthesis, LbL coating process, and characterization toolkit")

# ── Slide 9: Polymer Components ───────────────────────────────
sl = slide(prs, "02 · Methods", "Polymer Components: Structure and Properties", 9)
polys = [
    ("PA — Poly(carboxybetaine amide)",
     ["Zwitterionic: pendant COO⁻ and N⁺ groups",
      "Amphoteric character enables weak interaction with PQ",
      "Strong hydration layer at physiological pH",
      "Synthesized via RAFT polymerization",
      "Confirmed by ¹H NMR and GPC"], TEAL),
    ("PS — Poly(sulfobetaine)",
     ["Zwitterionic: pendant SO₃⁻ and N⁺ groups",
      "Balanced dipole — no net charge",
      "Known for protein and cell resistance",
      "Synthesized via RAFT polymerization",
      "Confirmed by ¹H NMR and GPC"], SLATE),
    ("PQ — Poly(quaternary ammonium)",
     ["Permanently cationic at all pH values",
      "Membrane-disrupting antibacterial activity",
      "Interacts electrostatically with PA or PS",
      "Interaction strength tunable by ratio",
      "Confirmed by ¹H NMR and GPC"], AMBER),
]
cw = (BODY_W - Inches(0.30)) / 3
by2 = BODY_Y + Inches(0.10)
for i, (ttl, buls, ac) in enumerate(polys):
    bx = BODY_X + i * (cw + Inches(0.15))
    rrect(sl, bx, by2, cw, Inches(5.40), CARD, BORDER)
    rect(sl, bx, by2, Inches(0.06), Inches(5.40), ac)
    text(sl, bx + Inches(0.20), by2 + Inches(0.18), cw - Inches(0.34),
         Inches(0.40), ttl, size=12, bold=True, color=TEXT)
    fig_box(sl, bx + Inches(0.16), by2 + Inches(0.66), cw - Inches(0.32),
            Inches(1.60), "Chemical structure")
    cy = by2 + Inches(2.40)
    for b in buls:
        text(sl, bx + Inches(0.20), cy, Inches(0.18), Inches(0.26),
             "·", size=10, bold=True, color=ac)
        text(sl, bx + Inches(0.40), cy, cw - Inches(0.56), Inches(0.26),
             b, size=10, color=TEXT2)
        cy += Inches(0.27)

# ── Slide 10: LbL Coating & Nomenclature ─────────────────────
sl = slide(prs, "02 · Methods", "Layer-by-Layer Coating Process & Sample Nomenclature", 10)
# Process steps
steps = [
    ("1", "Surface\nPreparation", "PLA film cleaned with ethanol and DI water"),
    ("2", "Co-assembly\nSolution", "Mix zwitterionic + PQ at defined concentrations"),
    ("3", "Dip Coating", "Substrate immersed in co-assembly solution"),
    ("4", "Drying &\nCuring", "Air-dried; process repeated for multilayers"),
    ("5", "Characterization", "QCM-d, AFM, XPS, WCA, antibacterial/antithrombotic"),
]
sw = (BODY_W - Inches(0.40)) / 5
by2 = BODY_Y + Inches(0.10)
for i, (num, step, desc) in enumerate(steps):
    bx = BODY_X + i * (sw + Inches(0.10))
    rrect(sl, bx, by2, sw, Inches(2.20), CARD, BORDER)
    rrect(sl, bx + (sw - Inches(0.40)) / 2, by2 + Inches(0.14),
          Inches(0.40), Inches(0.40), TEAL, None)
    text(sl, bx + (sw - Inches(0.40)) / 2, by2 + Inches(0.14),
         Inches(0.40), Inches(0.40), num, size=14, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(sl, bx + Inches(0.10), by2 + Inches(0.65), sw - Inches(0.20),
         Inches(0.40), step, size=11, bold=True, color=TEXT, align=PP_ALIGN.CENTER)
    text(sl, bx + Inches(0.10), by2 + Inches(1.12), sw - Inches(0.20),
         Inches(0.96), desc, size=9, color=TEXT2, align=PP_ALIGN.CENTER)

# Nomenclature table
by3 = BODY_Y + Inches(2.50)
rrect(sl, BODY_X, by3, BODY_W, Inches(3.20), CARD, BORDER)
rect(sl, BODY_X, by3, BODY_W, Inches(0.05), TEAL)
text(sl, BODY_X + Inches(0.20), by3 + Inches(0.12), BODY_W - Inches(0.40),
     Inches(0.32), "Sample Nomenclature", size=12, bold=True, color=TEXT)
headers = ["Sample", "Zwitterionic polymer", "PQ conc. (mg/mL)", "Key property"]
rows2 = [
    ["PA2.0", "PA — 2 mg/mL",  "0.5",  "Antifouling-dominant"],
    ["PA2.5", "PA — 2 mg/mL",  "1.0",  "Optimized dual function ★"],
    ["PS2.0", "PS — 2 mg/mL",  "0.5",  "Antibacterial (PQ accessible)"],
    ["PS2.5", "PS — 2.5 mg/mL","1.0",  "Reduced efficacy (PQ shielded)"],
]
col_xs = [BODY_X + Inches(0.20), BODY_X + Inches(2.40),
          BODY_X + Inches(5.60), BODY_X + Inches(8.40)]
col_ws = [Inches(2.0), Inches(3.0), Inches(2.6), Inches(4.0)]
hy = by3 + Inches(0.52)
for j, h in enumerate(headers):
    text(sl, col_xs[j], hy, col_ws[j], Inches(0.26),
         h, size=9, bold=True, color=MUTED)
for ri, row in enumerate(rows2):
    ry = by3 + Inches(0.86) + ri * Inches(0.52)
    bg = TEAL_L if "★" in row[3] else (CARD2 if ri % 2 == 0 else CARD)
    rect(sl, BODY_X + Inches(0.08), ry - Inches(0.04),
         BODY_W - Inches(0.16), Inches(0.48), bg)
    for j, val in enumerate(row):
        fc = TEAL_D if "★" in val else TEXT2
        text(sl, col_xs[j], ry, col_ws[j], Inches(0.42),
             val, size=10, bold=("★" in val), color=fc, anchor=MSO_ANCHOR.MIDDLE)

# ── Slide 11: Characterization Toolkit ───────────────────────
sl = slide(prs, "02 · Methods", "Characterization Toolkit", 11)
tools = [
    ("QCM-d", "Quartz Crystal Microbalance\nwith Dissipation",
     "Real-time mass deposition\nFilm stiffness (ΔD)\nAssembly kinetics", TEAL),
    ("AFM", "Atomic Force\nMicroscopy",
     "Surface morphology\nRoughness: Rq, Ra\nNano-scale topography", SLATE),
    ("XPS", "X-ray Photoelectron\nSpectroscopy",
     "Elemental composition\nSurface chemistry\nN, S, O quantification", AMBER),
    ("WCA", "Water Contact\nAngle",
     "Surface hydrophilicity\nPA2.5: 21.58° (hydrophilic)\nPLA: 88.73° (hydrophobic)", GREEN),
    ("ITC", "Isothermal Titration\nCalorimetry",
     "Thermodynamics of co-assembly\nΔH, ΔS, ΔG, Ka\nInteraction strength", RED),
    ("Fluorescence", "Fluorescence\nSpectroscopy",
     "Complex tightness\nPS-PQ: avg 34.25 (tight)\nPA-PQ: avg 43.59 (loose)", TEAL_D),
]
cw2 = (BODY_W - Inches(0.50)) / 3
by2 = BODY_Y + Inches(0.10)
for i, (abbr, full, details, ac) in enumerate(tools):
    col = i % 3; row = i // 3
    bx = BODY_X + col * (cw2 + Inches(0.25))
    by3 = by2 + row * Inches(2.80)
    rrect(sl, bx, by3, cw2, Inches(2.60), CARD, BORDER)
    rect(sl, bx, by3, Inches(0.06), Inches(2.60), ac)
    label(sl, bx + Inches(0.16), by3 + Inches(0.16),
          Inches(0.70), Inches(0.28), abbr, bg=TEAL_L, fg=TEAL_D, size=9)
    text(sl, bx + Inches(0.16), by3 + Inches(0.52), cw2 - Inches(0.30),
         Inches(0.34), full, size=10, bold=True, color=TEXT)
    text(sl, bx + Inches(0.16), by3 + Inches(0.92), cw2 - Inches(0.30),
         Inches(1.55), details, size=10, color=TEXT2)

# ── Slide 12: Section Divider — Solution Interactions ────────
divider_slide(prs, 12, "Section 03", "Solution Interactions",
              "ITC, fluorescence, turbidity, and viscometry — characterizing the co-assembly")

# ── Slide 13: ITC & Turbidity ─────────────────────────────────
sl = slide(prs, "03 · Solution Interactions", "ITC and Turbidity: Thermodynamics of Co-assembly", 13)
# ITC results
rrect(sl, BODY_X, BODY_Y + Inches(0.10), Inches(6.0), Inches(5.50), CARD, BORDER)
rect(sl, BODY_X, BODY_Y + Inches(0.10), Inches(6.0), Inches(0.05), TEAL)
text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(0.22), Inches(5.60), Inches(0.32),
     "Isothermal Titration Calorimetry (ITC)", size=12, bold=True, color=TEXT)
fig_box(sl, BODY_X + Inches(0.20), BODY_Y + Inches(0.62), Inches(5.60),
        Inches(2.80), "ITC thermogram — PA-PQ and PS-PQ")
itc_data = [
    ("ΔG < 0",         "Both pairs: thermodynamically spontaneous co-assembly"),
    ("PA-PQ",          "Weak electrostatic interaction — carboxylate groups"),
    ("PS-PQ",          "Relatively stronger hydrophobic/electrostatic interaction"),
    ("Key difference", "PA-PQ weak binding → PQ accessible for antibacterial action"),
]
cy = BODY_Y + Inches(3.56)
for k, v in itc_data:
    rrect(sl, BODY_X + Inches(0.20), cy, Inches(5.60), Inches(0.38), WHITE, BORDER, 0.3)
    text(sl, BODY_X + Inches(0.30), cy + Inches(0.04), Inches(1.20), Inches(0.30),
         k, size=9, bold=True, color=TEAL)
    text(sl, BODY_X + Inches(1.60), cy + Inches(0.04), Inches(4.10), Inches(0.30),
         v, size=9, color=TEXT2)
    cy += Inches(0.44)

# Turbidity
bx2 = BODY_X + Inches(6.30)
rrect(sl, bx2, BODY_Y + Inches(0.10), Inches(6.52), Inches(5.50), CARD, BORDER)
rect(sl, bx2, BODY_Y + Inches(0.10), Inches(6.52), Inches(0.05), AMBER)
text(sl, bx2 + Inches(0.20), BODY_Y + Inches(0.22), Inches(6.10), Inches(0.32),
     "Turbidity & Aggregate Formation", size=12, bold=True, color=TEXT)
fig_box(sl, bx2 + Inches(0.20), BODY_Y + Inches(0.62), Inches(6.10),
        Inches(2.80), "Turbidity vs concentration — aggregate size")
turb_notes = [
    "PA-PQ: turbidity increases at higher concentrations → larger aggregates",
    "PS-PQ: more compact aggregates (tighter complex)",
    "Loose PA-PQ assembly → functional groups remain exposed on surface",
    "Concentration ratio determines coating properties and performance",
]
cy = BODY_Y + Inches(3.56)
for note in turb_notes:
    text(sl, bx2 + Inches(0.20), cy, Inches(0.18), Inches(0.32),
         "·", size=10, bold=True, color=AMBER)
    text(sl, bx2 + Inches(0.40), cy, Inches(5.90), Inches(0.32),
         note, size=10, color=TEXT2)
    cy += Inches(0.40)

# ── Slide 14: Fluorescence & Viscosity ───────────────────────
sl = slide(prs, "03 · Solution Interactions", "Fluorescence and Viscosity: Complex Tightness", 14)
fig_box(sl, BODY_X, BODY_Y + Inches(0.10), Inches(6.0), Inches(3.50),
        "Fluorescence intensity vs titration ratio — PA-PQ vs PS-PQ")
fig_box(sl, BODY_X, BODY_Y + Inches(3.70), Inches(6.0), Inches(2.00),
        "Viscosity — individual vs co-assembly solutions")

bx2 = BODY_X + Inches(6.30)
# Stats
stats = [
    ("34.25", "PS-PQ avg intensity", "Tighter complex", TEAL),
    ("43.59", "PA-PQ avg intensity", "Looser complex", AMBER),
]
for i, (val, sub, desc, ac) in enumerate(stats):
    by3 = BODY_Y + Inches(0.10) + i * Inches(1.45)
    rrect(sl, bx2, by3, Inches(6.52), Inches(1.30), CARD, BORDER)
    rect(sl, bx2, by3, Inches(0.06), Inches(1.30), ac)
    text(sl, bx2 + Inches(0.20), by3 + Inches(0.16), Inches(2.0), Inches(0.50),
         val, size=28, bold=True, color=ac)
    text(sl, bx2 + Inches(0.20), by3 + Inches(0.68), Inches(5.80), Inches(0.24),
         sub, size=10, color=MUTED)
    text(sl, bx2 + Inches(2.50), by3 + Inches(0.36), Inches(3.80), Inches(0.50),
         desc, size=11, color=TEXT2)

by4 = BODY_Y + Inches(3.10)
rrect(sl, bx2, by4, Inches(6.52), Inches(2.60), CARD, BORDER)
rect(sl, bx2, by4, Inches(6.52), Inches(0.05), TEAL)
text(sl, bx2 + Inches(0.20), by4 + Inches(0.16), Inches(6.10), Inches(0.32),
     "Mechanistic Interpretation", size=12, bold=True, color=TEXT)
interp = ("PS-PQ forms a tighter complex (lower fluorescence emission, smaller aggregates). "
          "PA-PQ forms a looser complex — the carboxylate groups of PA interact weakly with "
          "PQ's quaternary ammonium. This weaker interaction is key: it allows PQ functional "
          "groups to remain accessible at the coating surface for antibacterial activity, "
          "while PA groups maintain the hydration layer for antifouling.")
text(sl, bx2 + Inches(0.20), by4 + Inches(0.56), Inches(6.10), Inches(1.90),
     interp, size=10, color=TEXT2)

# ── Slide 15: Section Divider — Coating Properties ───────────
divider_slide(prs, 15, "Section 04", "Coating Properties",
              "QCM-d, AFM, XPS, zeta potential, and water contact angle")

# ── Slide 16: QCM-d Assembly Kinetics ────────────────────────
sl = slide(prs, "04 · Coating Properties", "QCM-d: Real-Time Coating Deposition", 16)
fig_box(sl, BODY_X, BODY_Y + Inches(0.10), Inches(7.80), Inches(5.50),
        "QCM-d frequency (Δf) and dissipation (ΔD) vs time — PA-PQ and PS-PQ")

bx2 = BODY_X + Inches(8.10)
qcm_stats = [
    ("PA-PQ", "Δf = 156 Hz", "ΔD = 559 × 10⁻⁶",
     "More mass deposited\nHighly hydrated soft film → superior antifouling", TEAL),
    ("PS-PQ", "Δf = 89 Hz",  "ΔD = 225 × 10⁻⁶",
     "Less mass, stiffer film\nCompact, denser coating", SLATE),
]
by2 = BODY_Y + Inches(0.10)
for i, (name, df, dD, note, ac) in enumerate(qcm_stats):
    by3 = by2 + i * Inches(2.90)
    rrect(sl, bx2, by3, Inches(4.80), Inches(2.70), CARD, BORDER)
    rect(sl, bx2, by3, Inches(0.06), Inches(2.70), ac)
    text(sl, bx2 + Inches(0.20), by3 + Inches(0.16), Inches(4.40), Inches(0.30),
         name, size=12, bold=True, color=TEXT)
    text(sl, bx2 + Inches(0.20), by3 + Inches(0.52), Inches(4.40), Inches(0.30),
         df, size=13, bold=True, color=ac)
    text(sl, bx2 + Inches(0.20), by3 + Inches(0.86), Inches(4.40), Inches(0.30),
         dD, size=13, bold=True, color=ac)
    text(sl, bx2 + Inches(0.20), by3 + Inches(1.28), Inches(4.40), Inches(1.30),
         note, size=10, color=TEXT2)

# ── Slide 17: Zeta Potential ──────────────────────────────────
sl = slide(prs, "04 · Coating Properties", "Zeta Potential: Surface Charge Analysis", 17)
fig_box(sl, BODY_X, BODY_Y + Inches(0.10), Inches(7.0), Inches(5.50),
        "Zeta potential after each coating layer — PLA, PS-PQ, PA-PQ series")

bx2 = BODY_X + Inches(7.30)
text(sl, bx2, BODY_Y + Inches(0.16), Inches(5.60), Inches(0.32),
     "Key Findings", size=12, bold=True, color=TEXT)
zeta_pts = [
    "Bare PLA substrate: negative zeta potential",
    "After PS-PQ coating: shifts toward near-neutral charge balance",
    "After PA-PQ coating: slightly negative — favorable for antithrombotic surfaces",
    "PA2.5 surface charge reduces electrostatic attraction with blood proteins",
    "Charge evolution consistent with layer-by-layer buildup confirmed by QCM-d",
]
cy = BODY_Y + Inches(0.60)
for pt in zeta_pts:
    rrect(sl, bx2, cy, Inches(5.60), Inches(0.86), CARD, BORDER)
    rect(sl, bx2, cy, Inches(0.06), Inches(0.86), TEAL)
    text(sl, bx2 + Inches(0.18), cy + Inches(0.12), Inches(5.22), Inches(0.62),
         pt, size=10, color=TEXT2)
    cy += Inches(0.96)

# ── Slide 18: AFM & Roughness ─────────────────────────────────
sl = slide(prs, "04 · Coating Properties", "AFM: Surface Morphology and Roughness", 18)
fig_box(sl, BODY_X, BODY_Y + Inches(0.10), Inches(7.80), Inches(3.60),
        "AFM height images — PLA, PA-PQ (layers 1–5), PS-PQ (layers 1–5)")
fig_box(sl, BODY_X, BODY_Y + Inches(3.80), Inches(7.80), Inches(1.80),
        "Rq and Ra vs number of coating layers — PA-PQ vs PS-PQ")

bx2 = BODY_X + Inches(8.10)
afm_data = [
    ("PLA (bare)", "Rq = 1.06 nm", "Ra = 0.67 nm", "Baseline roughness", MUTED),
    ("PA-PQ",      "Rq decreases", "Smoother surface", "Each layer fills nanoscale defects → uniform hydration layer", TEAL),
    ("PS-PQ",      "Rq increases", "Rougher surface",  "Compact aggregates create higher topographic variation", SLATE),
]
by2 = BODY_Y + Inches(0.10)
for i, (name, rq, ra, note, ac) in enumerate(afm_data):
    by3 = by2 + i * Inches(1.88)
    rrect(sl, bx2, by3, Inches(4.80), Inches(1.72), CARD, BORDER)
    rect(sl, bx2, by3, Inches(0.06), Inches(1.72), ac)
    text(sl, bx2 + Inches(0.20), by3 + Inches(0.14), Inches(4.40), Inches(0.28),
         name, size=11, bold=True, color=TEXT)
    text(sl, bx2 + Inches(0.20), by3 + Inches(0.48), Inches(2.0), Inches(0.26),
         rq, size=10, bold=True, color=ac)
    text(sl, bx2 + Inches(2.30), by3 + Inches(0.48), Inches(2.1), Inches(0.26),
         ra, size=10, bold=True, color=ac)
    text(sl, bx2 + Inches(0.20), by3 + Inches(0.82), Inches(4.40), Inches(0.76),
         note, size=9, color=TEXT2)

# ── Slide 19: Hydration & WCA ─────────────────────────────────
sl = slide(prs, "04 · Coating Properties", "Water Contact Angle: Surface Hydrophilicity", 19)
fig_box(sl, BODY_X, BODY_Y + Inches(0.10), Inches(7.0), Inches(4.0),
        "WCA photographs and values — PLA, PA2.0, PA2.5, PS2.0, PS2.5")

bx2 = BODY_X + Inches(7.30)
wca_data = [
    ("PLA (control)", "88.73°", "Moderately hydrophobic", RED),
    ("PA2.0",         "53.3°",  "Hydrophilic (PA coating)", TEAL),
    ("PA2.5",         "21.58°", "Highly hydrophilic ★",    TEAL),
    ("PS series",     ">60°",   "More hydrophobic than PA", SLATE),
]
by2 = BODY_Y + Inches(0.10)
for i, (name, angle, desc, ac) in enumerate(wca_data):
    by3 = by2 + i * Inches(1.38)
    rrect(sl, bx2, by3, Inches(5.60), Inches(1.24), CARD, BORDER)
    rect(sl, bx2, by3, Inches(0.06), Inches(1.24), ac)
    text(sl, bx2 + Inches(0.20), by3 + Inches(0.14), Inches(2.60), Inches(0.28),
         name, size=11, bold=True, color=TEXT)
    text(sl, bx2 + Inches(0.20), by3 + Inches(0.50), Inches(1.60), Inches(0.40),
         angle, size=20, bold=True, color=ac)
    text(sl, bx2 + Inches(2.00), by3 + Inches(0.36), Inches(3.40), Inches(0.60),
         desc, size=10, color=TEXT2)

rrect(sl, BODY_X, BODY_Y + Inches(4.26), Inches(7.0), Inches(1.40), TEAL_L, TEAL, 0.5)
text(sl, BODY_X + Inches(0.20), BODY_Y + Inches(4.42), Inches(6.60), Inches(1.04),
     "Mechanism: The zwitterionic PA groups form a dense hydration shell through "
     "hydrogen bonding and electrostatic interaction with water molecules. "
     "Higher PQ content (PA2.5) increases surface charge density, further enhancing hydrophilicity.",
     size=10, color=TEAL_D)

# ── Slide 20: Section Divider — Performance ──────────────────
divider_slide(prs, 20, "Section 05", "Antifouling & Antithrombotic Performance",
              "Ex vivo whole blood contact — protein, platelet, and thrombus evaluation")

# ── Slide 21: Protein & Platelet Adhesion ────────────────────
sl = slide(prs, "05 · Performance", "Protein Adsorption and Platelet Adhesion", 21)
fig_box(sl, BODY_X, BODY_Y + Inches(0.10), Inches(6.20), Inches(2.80),
        "Fluorescence images — protein adsorption (BSA-FITC) on PLA vs PA2.5")
fig_box(sl, BODY_X, BODY_Y + Inches(3.06), Inches(6.20), Inches(2.60),
        "SEM images — platelet adhesion and activation on PLA vs coated surfaces")

bx2 = BODY_X + Inches(6.50)
text(sl, bx2, BODY_Y + Inches(0.16), Inches(6.34), Inches(0.32),
     "Antifouling Performance", size=12, bold=True, color=TEXT)
pts = [
    ("Protein adsorption",
     "PA2.5 dramatically reduces BSA adsorption vs bare PLA. "
     "The hydration layer acts as a steric/energetic barrier."),
    ("Platelet adhesion",
     "Significantly fewer platelets adhere to PA2.5 surfaces. "
     "Adhered platelets show rounded (inactive) morphology vs activated on PLA."),
    ("Mechanism",
     "Highly hydrated PA groups (WCA 21.58°) repel hydrophobic proteins "
     "and inhibit platelet integrin engagement with the surface."),
    ("PA vs PS",
     "PA2.5 outperforms PS series due to greater hydrophilicity "
     "and more uniform coating morphology (smoother AFM topography)."),
]
cy = BODY_Y + Inches(0.60)
for ttl, body in pts:
    rrect(sl, bx2, cy, Inches(6.34), Inches(1.24), CARD, BORDER)
    rect(sl, bx2, cy, Inches(0.06), Inches(1.24), TEAL)
    text(sl, bx2 + Inches(0.20), cy + Inches(0.14), Inches(5.94), Inches(0.28),
         ttl, size=11, bold=True, color=TEXT)
    text(sl, bx2 + Inches(0.20), cy + Inches(0.48), Inches(5.94), Inches(0.68),
         body, size=10, color=TEXT2)
    cy += Inches(1.34)

# ── Slide 22: Ex Vivo Antithrombotic ─────────────────────────
sl = slide(prs, "05 · Performance", "Ex Vivo Antithrombotic: Whole Blood Contact Test", 22)
fig_box(sl, BODY_X, BODY_Y + Inches(0.10), Inches(6.50), Inches(3.80),
        "Photographs and SEM of tube surfaces after whole-blood contact — PLA, PS, PA series")

bx2 = BODY_X + Inches(6.80)
# Thrombus coverage bar chart (manual)
text(sl, bx2, BODY_Y + Inches(0.14), Inches(6.06), Inches(0.28),
     "Thrombus Coverage (%)", size=11, bold=True, color=TEXT)
thromb_data = [
    ("PLA",  99.0, RED),
    ("PS2.0", 97.1, RED),
    ("PS2.5", 89.7, AMBER),
    ("PA2.0", 30.3, TEAL),
    ("PA2.5",  3.1, GREEN),
]
by3 = BODY_Y + Inches(0.52)
for name, pct, ac in thromb_data:
    bar_row(sl, bx2, by3, Inches(6.06), Inches(0.54), name, pct,
            bar_color=ac, label_w=Inches(1.4), val_w=Inches(0.80))
    by3 += Inches(0.60)

# Mass increase
text(sl, bx2, BODY_Y + Inches(3.62), Inches(6.06), Inches(0.28),
     "Thrombus Mass Increase (%)", size=11, bold=True, color=TEXT)
mass_data = [
    ("PLA",  152.9, RED),
    ("PS2.0",  90.3, AMBER),
    ("PS2.5",  82.2, AMBER),
    ("PA2.0",  52.5, TEAL),
    ("PA2.5",  51.7, GREEN),
]
by4 = BODY_Y + Inches(4.00)
for name, pct, ac in mass_data:
    bar_row(sl, bx2, by4, Inches(6.06), Inches(0.42), name, pct,
            bar_color=ac, label_w=Inches(1.4), val_w=Inches(0.80))
    by4 += Inches(0.48)

# ── Slide 23: Section Divider — Antibacterial ────────────────
divider_slide(prs, 23, "Section 06", "Antibacterial Performance",
              "MIC assay, solution bactericidal efficiency, and contact sterilization")

# ── Slide 24: MIC & Solution Bactericidal ────────────────────
sl = slide(prs, "06 · Antibacterial", "MIC Assay and Solution Bactericidal Efficiency", 24)
fig_box(sl, BODY_X, BODY_Y + Inches(0.10), Inches(6.50), Inches(3.0),
        "MIC heatmap — PA/PQ concentration ratios vs S. aureus / E. coli growth")
fig_box(sl, BODY_X, BODY_Y + Inches(3.20), Inches(6.50), Inches(2.40),
        "Colony plates — solution bactericidal test PA-PQ vs PS-PQ")

bx2 = BODY_X + Inches(6.80)
text(sl, bx2, BODY_Y + Inches(0.14), Inches(6.06), Inches(0.30),
     "MIC Results (PA/PQ mg/mL)", size=12, bold=True, color=TEXT)
mic_rows = [
    ("PA 2 / PQ 0.5", "Effective", GREEN, "✓"),
    ("PA 2 / PQ 1.0", "Effective", GREEN, "✓"),
    ("PA 2 / PQ 2.0", "Effective", GREEN, "✓"),
    ("PA 2 / PQ 3.0", "Effective", GREEN, "✓"),
]
cy = BODY_Y + Inches(0.54)
for ratio, result, ac, mark in mic_rows:
    rrect(sl, bx2, cy, Inches(6.06), Inches(0.44), CARD, BORDER)
    text(sl, bx2 + Inches(0.16), cy + Inches(0.08), Inches(3.40), Inches(0.28),
         ratio, size=10, color=TEXT2)
    label(sl, bx2 + Inches(3.70), cy + Inches(0.08), Inches(1.10), Inches(0.28),
          result, bg=GREEN_L, fg=GREEN, size=9)
    cy += Inches(0.50)

# PS comparison
text(sl, bx2, BODY_Y + Inches(2.74), Inches(6.06), Inches(0.30),
     "PS-PQ — Formulation-Dependent Efficacy", size=12, bold=True, color=TEXT)
ps_rows = [
    ("PS2.0", "97.41%", "Effective — PQ accessible", GREEN),
    ("PS5.0", "97.04%", "Effective — PQ accessible", GREEN),
    ("PS2.5", "<10%",   "Ineffective — PQ shielded by excess PS", RED),
    ("PS5.5", "<10%",   "Ineffective — PQ shielded by excess PS", RED),
]
cy = BODY_Y + Inches(3.14)
for name, eff, note, ac in ps_rows:
    rrect(sl, bx2, cy, Inches(6.06), Inches(0.50), CARD, BORDER)
    rect(sl, bx2, cy, Inches(0.06), Inches(0.50), ac)
    text(sl, bx2 + Inches(0.20), cy + Inches(0.10), Inches(1.0), Inches(0.30),
         name, size=10, bold=True, color=TEXT)
    text(sl, bx2 + Inches(1.30), cy + Inches(0.10), Inches(1.0), Inches(0.30),
         eff, size=12, bold=True, color=ac)
    text(sl, bx2 + Inches(2.50), cy + Inches(0.10), Inches(3.40), Inches(0.30),
         note, size=9, color=TEXT2)
    cy += Inches(0.58)

# ── Slide 25: Contact Sterilization ──────────────────────────
sl = slide(prs, "06 · Antibacterial", "Contact Sterilization: In Vitro Antibacterial Surface", 25)
fig_box(sl, BODY_X, BODY_Y + Inches(0.10), Inches(7.0), Inches(3.80),
        "Colony-forming unit (CFU) plates — PLA, PA2.0, PA2.5, PS2.0, PS2.5")

# Stat boxes
stats2 = [
    ("6.87×10³", "CFU/cm²", "Bare PLA", RED),
    ("~0",       "CFU/cm²", "PA2.0 coating", GREEN),
    ("~0",       "CFU/cm²", "PA2.5 coating ★", GREEN),
]
bx0 = BODY_X
by3 = BODY_Y + Inches(4.08)
bw = (Inches(7.0) - Inches(0.30)) / 3
for i, (val, unit, lbl, ac) in enumerate(stats2):
    bx = bx0 + i * (bw + Inches(0.15))
    rrect(sl, bx, by3, bw, Inches(1.56), CARD, BORDER)
    rect(sl, bx, by3, bw, Inches(0.05), ac)
    text(sl, bx + Inches(0.10), by3 + Inches(0.16), bw - Inches(0.20),
         Inches(0.48), val, size=22, bold=True, color=ac, align=PP_ALIGN.CENTER)
    text(sl, bx + Inches(0.10), by3 + Inches(0.68), bw - Inches(0.20),
         Inches(0.22), unit, size=9, color=MUTED, align=PP_ALIGN.CENTER)
    text(sl, bx + Inches(0.10), by3 + Inches(0.98), bw - Inches(0.20),
         Inches(0.44), lbl, size=10, bold=True, color=TEXT2, align=PP_ALIGN.CENTER)

bx2 = BODY_X + Inches(7.30)
rrect(sl, bx2, BODY_Y + Inches(0.10), Inches(5.60), Inches(5.54), CARD, BORDER)
rect(sl, bx2, BODY_Y + Inches(0.10), Inches(5.60), Inches(0.05), TEAL)
text(sl, bx2 + Inches(0.20), BODY_Y + Inches(0.22), Inches(5.20), Inches(0.32),
     "Mechanism: Contact-Killing", size=12, bold=True, color=TEXT)
mech_pts = [
    "PQ quaternary ammonium groups remain surface-accessible on PA2.5",
    "Electrostatic attraction draws negatively-charged bacterial membrane",
    "Membrane disruption → intracellular contents leak → cell death",
    "PA zwitterionic groups prevent non-specific adhesion of debris",
    "No antibiotic released → no resistance pressure",
    "PA2.0 and PA2.5 both achieve near-complete (~100%) surface sterilization",
]
cy = BODY_Y + Inches(0.68)
for pt in mech_pts:
    text(sl, bx2 + Inches(0.20), cy, Inches(0.20), Inches(0.28),
         "·", size=11, bold=True, color=TEAL)
    text(sl, bx2 + Inches(0.42), cy, Inches(5.00), Inches(0.28),
         pt, size=10, color=TEXT2)
    cy += Inches(0.70)

# ── Slide 26: HERO Slide ──────────────────────────────────────
sl = prs.slides.add_slide(prs.slide_layouts[6])
rect(sl, 0, 0, W, H, WHITE)
rect(sl, 0, 0, Inches(0.10), H, TEAL)
draw_footer(sl, 26)
text(sl, Inches(0.24), Inches(0.24), Inches(9), Inches(0.28),
     "PA2.5 — OPTIMIZED DUAL-FUNCTION COATING", size=9, bold=True, color=TEAL)
text(sl, Inches(0.24), Inches(0.56), Inches(10), Inches(0.52),
     "Key Performance Results", size=26, bold=True, color=TEXT)
rect(sl, Inches(0.24), Inches(1.16), Inches(9.0), Inches(0.03), BORDER)

hero_stats = [
    ("3.1%",    "Thrombus coverage\n(vs PLA 99.0%)",    GREEN),
    (">97.6%",  "Bactericidal efficiency\nvs S. aureus", TEAL),
    ("21.58°",  "Water contact angle\n(highly hydrophilic)", TEAL_D),
    ("~0 CFU",  "CFU/cm² on surface\n(contact sterilization)", GREEN),
    (">99%",    "In vivo bacterial\neradication (rat model)", TEAL),
    ("<5%",     "Hemolysis rate\n(non-hemolytic)", GREEN),
]
sw = (W - Inches(0.24) - Inches(0.20) - Inches(0.50)) / 3
for i, (val, lbl, ac) in enumerate(hero_stats):
    col = i % 3; row = i // 3
    bx = Inches(0.24) + col * (sw + Inches(0.25))
    by3 = Inches(1.28) + row * Inches(2.75)
    rrect(sl, bx, by3, sw, Inches(2.55), CARD, BORDER)
    rect(sl, bx, by3, sw, Inches(0.06), ac)
    text(sl, bx + Inches(0.20), by3 + Inches(0.24), sw - Inches(0.40),
         Inches(0.80), val, size=36, bold=True, color=ac)
    text(sl, bx + Inches(0.20), by3 + Inches(1.12), sw - Inches(0.40),
         Inches(1.30), lbl, size=11, color=TEXT2)

# Right panel context
bx_r = Inches(0.24) + 3 * (sw + Inches(0.25)) + Inches(0.10)
bw_r = W - bx_r - Inches(0.20)
rrect(sl, bx_r, Inches(1.28), bw_r, Inches(6.0), TEAL_L, TEAL, 0.7)
text(sl, bx_r + Inches(0.20), Inches(1.44), bw_r - Inches(0.40), Inches(0.32),
     "Why PA2.5?", size=12, bold=True, color=TEAL_D)
context_pts = [
    "Weak PA-PQ electrostatic interaction preserves PQ accessibility",
    "Highly hydrated surface (WCA 21.58°) repels blood proteins",
    "Smooth morphology (AFM) enables uniform functional coverage",
    "Dual function confirmed in 3 animal models",
    "Biocompatible: hemolysis <5%, L929 viability >80%",
]
cy = Inches(1.86)
for pt in context_pts:
    text(sl, bx_r + Inches(0.20), cy, Inches(0.20), Inches(0.28),
         "✓", size=11, bold=True, color=TEAL_D)
    text(sl, bx_r + Inches(0.44), cy, bw_r - Inches(0.66), Inches(0.28),
         pt, size=10, color=TEAL_D)
    cy += Inches(0.80)

# ── Slide 27: Section Divider — Biocompatibility ─────────────
divider_slide(prs, 27, "Section 07", "Biocompatibility & In Vivo Validation",
              "Hemolysis, cytotoxicity, and three animal model studies")

# ── Slide 28: Hemolysis & Cytotoxicity ───────────────────────
sl = slide(prs, "07 · Biocompatibility", "Hemolysis and Cytotoxicity: In Vitro Safety", 28)
fig_box(sl, BODY_X, BODY_Y + Inches(0.10), Inches(6.0), Inches(2.80),
        "Hemolysis rate (%) — PLA, PA2.0, PA2.5, PS2.0, PS2.5")
fig_box(sl, BODY_X, BODY_Y + Inches(3.04), Inches(6.0), Inches(2.60),
        "L929 cell viability (%) — MTT assay — all formulations at 24/48/72 h")

bx2 = BODY_X + Inches(6.30)
safety_stats = [
    ("<5%",  "Hemolysis rate",      "All groups non-hemolytic\n(< 5% threshold)", GREEN),
    (">80%", "L929 cell viability", "All groups biocompatible\nat all time points",  GREEN),
]
by2 = BODY_Y + Inches(0.10)
for i, (val, sub, note, ac) in enumerate(safety_stats):
    by3 = by2 + i * Inches(2.90)
    rrect(sl, bx2, by3, Inches(6.52), Inches(2.70), CARD, BORDER)
    rect(sl, bx2, by3, Inches(6.52), Inches(0.05), ac)
    text(sl, bx2 + Inches(0.24), by3 + Inches(0.20), Inches(3.0), Inches(0.70),
         val, size=40, bold=True, color=ac)
    text(sl, bx2 + Inches(0.24), by3 + Inches(0.94), Inches(6.0), Inches(0.28),
         sub, size=11, bold=True, color=TEXT)
    text(sl, bx2 + Inches(0.24), by3 + Inches(1.30), Inches(6.0), Inches(1.20),
         note, size=10, color=TEXT2)

# ── Slide 29: Rat Subcutaneous Model ─────────────────────────
sl = slide(prs, "07 · Biocompatibility", "In Vivo: Rat Subcutaneous Implant Model", 29)
fig_box(sl, BODY_X, BODY_Y + Inches(0.10), Inches(7.80), Inches(3.20),
        "H&E staining and immunofluorescence (CD3, CD68) — day 1 and day 3 post-implant")
fig_box(sl, BODY_X, BODY_Y + Inches(3.40), Inches(7.80), Inches(2.20),
        "CFU recovered from tissue — PA2.5 vs control day 1 and day 3")

bx2 = BODY_X + Inches(8.10)
rat_data = [
    (">99%", "Bacterial eradication\nat day 1", "PA2.5 implant", GREEN),
    (">99%", "Bacterial eradication\nat day 3", "PA2.5 implant", GREEN),
]
by2 = BODY_Y + Inches(0.10)
for i, (val, lbl, sub, ac) in enumerate(rat_data):
    by3 = by2 + i * Inches(1.60)
    rrect(sl, bx2, by3, Inches(4.80), Inches(1.46), CARD, BORDER)
    rect(sl, bx2, by3, Inches(0.06), Inches(1.46), ac)
    text(sl, bx2 + Inches(0.20), by3 + Inches(0.16), Inches(1.60), Inches(0.52),
         val, size=28, bold=True, color=ac)
    text(sl, bx2 + Inches(0.20), by3 + Inches(0.72), Inches(4.40), Inches(0.28),
         lbl, size=10, bold=True, color=TEXT)
    text(sl, bx2 + Inches(0.20), by3 + Inches(1.04), Inches(4.40), Inches(0.28),
         sub, size=9, color=MUTED)

rrect(sl, bx2, BODY_Y + Inches(3.38), Inches(4.80), Inches(2.26), CARD, BORDER)
rect(sl, bx2, BODY_Y + Inches(3.38), Inches(0.06), Inches(2.26), TEAL)
text(sl, bx2 + Inches(0.20), BODY_Y + Inches(3.54), Inches(4.40), Inches(0.30),
     "Immune Response", size=11, bold=True, color=TEXT)
imm_pts = [
    "H&E: PA2.5 shows mild tissue reaction vs significant inflammation in control",
    "CD3+ T-cell infiltration: significantly reduced around PA2.5 implant",
    "CD68+ macrophage infiltration: reduced — low chronic inflammation",
    "PA2.5 demonstrates favorable host immune response profile",
]
cy = BODY_Y + Inches(3.92)
for pt in imm_pts:
    text(sl, bx2 + Inches(0.20), cy, Inches(0.20), Inches(0.26),
         "·", size=10, bold=True, color=TEAL)
    text(sl, bx2 + Inches(0.42), cy, Inches(4.20), Inches(0.26),
         pt, size=9, color=TEXT2)
    cy += Inches(0.38)

# ── Slide 30: Rabbit Jugular Vein Model ──────────────────────
sl = slide(prs, "07 · Biocompatibility", "In Vivo: Rabbit Jugular Vein Antithrombotic Model", 30)
fig_box(sl, BODY_X, BODY_Y + Inches(0.10), BODY_W, Inches(3.40),
        "Photographs and histology of jugular vein sections — PLA, PS2.5, PA2.5 after implantation")

bx_l = BODY_X
bx_m = BODY_X + Inches(4.20)
bx_r = BODY_X + Inches(8.40)
cw3  = Inches(3.80)
by2  = BODY_Y + Inches(3.62)
for bx, name, desc, ac in [
    (bx_l, "PLA (Control)", "Significant thrombus formation\nLumen markedly narrowed\nPlatelet aggregation extensive", RED),
    (bx_m, "PS2.5 Coating",  "Substantial thrombus\nSome improvement vs PLA\nCationic charge limits antifouling", AMBER),
    (bx_r, "PA2.5 Coating",  "Minimal thrombus\nLumen largely patent\nConfirms ex vivo results in vivo ★", GREEN),
]:
    rrect(sl, bx, by2, cw3, Inches(2.58), CARD, BORDER)
    rect(sl, bx, by2, cw3, Inches(0.06), ac)
    text(sl, bx + Inches(0.18), by2 + Inches(0.18), cw3 - Inches(0.36),
         Inches(0.32), name, size=12, bold=True, color=TEXT)
    text(sl, bx + Inches(0.18), by2 + Inches(0.58), cw3 - Inches(0.36),
         Inches(1.84), desc, size=10, color=TEXT2)

# ── Slide 31: PGLA Suture Model ───────────────────────────────
sl = slide(prs, "07 · Biocompatibility", "In Vivo: PGLA Suture Antibacterial Model (Rat)", 31)
fig_box(sl, BODY_X, BODY_Y + Inches(0.10), Inches(7.0), Inches(3.80),
        "CFU from tissue around PGLA sutures — PA2.5-coated vs uncoated — day 1 and day 3")
fig_box(sl, BODY_X, BODY_Y + Inches(4.02), Inches(7.0), Inches(1.62),
        "Fluorescence images of bacteria around suture wound — day 1 and day 3")

bx2 = BODY_X + Inches(7.30)
rrect(sl, bx2, BODY_Y + Inches(0.10), Inches(5.60), Inches(5.54), CARD, BORDER)
rect(sl, bx2, BODY_Y + Inches(0.10), Inches(0.06), Inches(5.54), TEAL)
text(sl, bx2 + Inches(0.20), BODY_Y + Inches(0.26), Inches(5.20), Inches(0.32),
     "PGLA Suture Model — Key Findings", size=12, bold=True, color=TEXT)
pgla_pts = [
    (">99%", "Bacterial eradication around PA2.5-coated suture at day 1"),
    (">99%", "Bacterial eradication around PA2.5-coated suture at day 3"),
    ("—",    "Demonstrates practical applicability: biodegradable suture material"),
    ("—",    "Coating persists through early wound healing period"),
    ("—",    "Reduced immune cell infiltration vs uncoated suture control"),
    ("—",    "Supports translation to clinical suture and implant applications"),
]
cy = BODY_Y + Inches(0.72)
for val, desc in pgla_pts:
    if val != "—":
        rrect(sl, bx2 + Inches(0.20), cy, Inches(0.80), Inches(0.38), TEAL_L, None)
        text(sl, bx2 + Inches(0.20), cy, Inches(0.80), Inches(0.38),
             val, size=9, bold=True, color=TEAL_D, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        text(sl, bx2 + Inches(1.10), cy, Inches(4.30), Inches(0.38),
             desc, size=9, color=TEXT2, anchor=MSO_ANCHOR.MIDDLE)
    else:
        text(sl, bx2 + Inches(0.20), cy, Inches(0.20), Inches(0.38),
             "·", size=10, bold=True, color=TEAL)
        text(sl, bx2 + Inches(0.42), cy, Inches(4.98), Inches(0.38),
             desc, size=9, color=TEXT2)
    cy += Inches(0.76)

# ── Slide 32: Section Divider — Conclusions ──────────────────
divider_slide(prs, 32, "Section 08", "Conclusions",
              "Mechanism, comparative summary, and future outlook")

# ── Slide 33: Mechanism Summary ───────────────────────────────
sl = slide(prs, "08 · Conclusions", "Proposed Mechanism: Weak Co-assembly Enables Dual Function", 33)
fig_box(sl, BODY_X, BODY_Y + Inches(0.10), Inches(5.80), Inches(5.54),
        "Schematic: weak PA-PQ interaction → exposed PQ (antibacterial) + hydrated PA (antifouling)")

bx2 = BODY_X + Inches(6.10)
mech_items = [
    ("Co-assembly", TEAL,
     "PA and PQ co-assemble spontaneously (ΔG < 0) via weak "
     "electrostatic interactions between carboxylate and quaternary ammonium groups."),
    ("Coating formation", SLATE,
     "LbL deposition creates a stable, smooth, hydrated film on PLA. "
     "PA2.5 produces the most uniform and hydrophilic coating (WCA 21.58°)."),
    ("Antifouling/Antithrombotic", TEAL,
     "Dense hydration layer from PA zwitterionic groups repels proteins and "
     "platelets. Low surface energy prevents thrombus initiation (3.1% coverage)."),
    ("Antibacterial", GREEN,
     "Weak interaction leaves PQ functional groups surface-accessible. "
     "Membrane disruption kills bacteria on contact (>97.6% efficiency) "
     "without antibiotic release."),
    ("Biocompatibility", GREEN,
     "Zwitterionic character minimizes non-specific interactions. "
     "Hemolysis <5%, cell viability >80% at all tested concentrations."),
]
cy = BODY_Y + Inches(0.10)
mh = Inches(5.54) / len(mech_items) - Inches(0.06)
for ttl, ac, body in mech_items:
    rrect(sl, bx2, cy, Inches(7.06), mh, CARD, BORDER)
    rect(sl, bx2, cy, Inches(0.06), mh, ac)
    text(sl, bx2 + Inches(0.20), cy + Inches(0.10), Inches(2.0), Inches(0.28),
         ttl, size=11, bold=True, color=TEXT)
    text(sl, bx2 + Inches(0.20), cy + Inches(0.40), Inches(6.66), mh - Inches(0.50),
         body, size=9, color=TEXT2)
    cy += mh + Inches(0.06)

# ── Slide 34: Comparison Table ───────────────────────────────
sl = slide(prs, "08 · Conclusions", "Performance Comparison: PA2.5 vs Controls", 34)
headers3 = ["Metric", "PLA (bare)", "PS2.0", "PS2.5", "PA2.0", "PA2.5 ★"]
col_data = [
    ["Thrombus coverage",    "99.0%",   "97.1%",   "89.7%",   "30.3%",  "3.1%"],
    ["Thrombus mass incr.",  "152.9%",  "90.3%",   "82.2%",   "52.5%",  "51.7%"],
    ["WCA (°)",              "88.73",   "~70",     "~65",     "53.3",   "21.58"],
    ["Solution bactericidal","—",       "97.41%",  "<10%",    ">97%",   ">97.6%"],
    ["Surface CFU/cm²",      "6870",    "~high",   "—",       "~0",     "~0"],
    ["Hemolysis",            "<5%",     "<5%",     "<5%",     "<5%",    "<5%"],
    ["Cell viability",       ">80%",    ">80%",    ">80%",    ">80%",   ">80%"],
    ["In vivo bacterial kill","—",      "—",       "—",       "—",      ">99%"],
]
table_w = BODY_W
col_widths = [Inches(2.40), Inches(1.50), Inches(1.50), Inches(1.50), Inches(1.50), Inches(2.00)]
col_starts = [BODY_X]
for cw4 in col_widths[:-1]:
    col_starts.append(col_starts[-1] + cw4 + Inches(0.02))

row_h = Inches(0.52)
hdr_y = BODY_Y + Inches(0.10)
# Header row
rect(sl, BODY_X, hdr_y, BODY_W, row_h, TEAL)
for j, (hdr, cs) in enumerate(zip(headers3, col_starts)):
    fc = WHITE if j < 5 else AMBER_L
    bld = j == 5
    text(sl, cs + Inches(0.08), hdr_y, col_widths[j] - Inches(0.08), row_h,
         hdr, size=10, bold=True, color=fc, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)

for ri, row_vals in enumerate(col_data):
    ry = hdr_y + row_h + ri * row_h
    bg = CARD2 if ri % 2 == 0 else WHITE
    rect(sl, BODY_X, ry, BODY_W, row_h, bg)
    rect(sl, BODY_X, ry, BODY_W, Inches(0.01), BORDER)
    for j, (val, cs) in enumerate(zip(row_vals, col_starts)):
        if j == 5:  # PA2.5 column highlight
            rect(sl, cs, ry, col_widths[j], row_h, TEAL_L)
        is_best = (j == 5 and val not in ["—", ">80%", "<5%"])
        fc = TEAL_D if j == 5 else (RED if j == 1 and ri < 2 else TEXT2)
        text(sl, cs + Inches(0.08), ry, col_widths[j] - Inches(0.08), row_h,
             val, size=10, bold=(j == 5), color=fc,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ── Slide 35: Conclusion & Impact ────────────────────────────
sl = prs.slides.add_slide(prs.slide_layouts[6])
rect(sl, 0, 0, W, H, WHITE)
rect(sl, 0, 0, Inches(0.10), H, TEAL)
draw_footer(sl, 35)
text(sl, Inches(0.24), Inches(0.24), Inches(9), Inches(0.28),
     "CONCLUSIONS & IMPACT", size=9, bold=True, color=TEAL)
text(sl, Inches(0.24), Inches(0.56), Inches(10), Inches(0.52),
     "Conclusions and Future Outlook", size=26, bold=True, color=TEXT)
rect(sl, Inches(0.24), Inches(1.16), Inches(9.0), Inches(0.03), BORDER)

conclusions = [
    ("First demonstration",
     "PA-PQ co-assembly via weak electrostatic interactions achieves genuine dual "
     "antibacterial and antithrombotic function from a single coating system.",
     TEAL),
    ("Key quantitative outcomes",
     "PA2.5 coating: thrombus coverage 3.1% (vs PLA 99.0%), bactericidal efficiency "
     ">97.6%, in vivo bacterial eradication >99%, hemolysis <5%.",
     GREEN),
    ("Mechanism validated",
     "The weak PA-PQ interaction is essential: strong interaction (as in PS-PQ) shields "
     "PQ groups and compromises antibacterial activity. Weak interaction preserves both functions.",
     SLATE),
    ("Translational potential",
     "Demonstrated on PLA films and PGLA sutures. Applicable to catheters, orthopedic implants, "
     "cardiovascular devices. LbL process is substrate-agnostic and scalable.",
     AMBER),
]
by2 = Inches(1.28)
bh2 = Inches(1.46)
cw5 = (W - Inches(0.24) - Inches(0.20) - Inches(0.45)) / 2
for i, (ttl, body, ac) in enumerate(conclusions):
    col = i % 2; row = i // 2
    bx = Inches(0.24) + col * (cw5 + Inches(0.30))
    by3 = by2 + row * (bh2 + Inches(0.20))
    rrect(sl, bx, by3, cw5, bh2, CARD, BORDER)
    rect(sl, bx, by3, Inches(0.06), bh2, ac)
    text(sl, bx + Inches(0.20), by3 + Inches(0.16), cw5 - Inches(0.30),
         Inches(0.32), ttl, size=12, bold=True, color=TEXT)
    text(sl, bx + Inches(0.20), by3 + Inches(0.54), cw5 - Inches(0.30),
         Inches(0.84), body, size=10, color=TEXT2)

# Future outlook
by5 = by2 + 2 * (bh2 + Inches(0.20))
rrect(sl, Inches(0.24), by5, W - Inches(0.44), Inches(1.68), TEAL_L, TEAL, 0.7)
text(sl, Inches(0.44), by5 + Inches(0.16), Inches(4.0), Inches(0.30),
     "Future Directions", size=12, bold=True, color=TEAL_D)
futures = [
    "Long-term stability and durability studies under physiological conditions",
    "Clinical translation: catheter and cardiovascular device coating trials",
    "Expanded pathogen spectrum: Gram-negative bacteria, fungi, viruses",
    "Biodegradable formulations for transient implant applications",
]
cx = Inches(0.44)
cy = by5 + Inches(0.58)
fw = (W - Inches(0.88)) / 2
for fi, ft in enumerate(futures):
    bx_f = cx + (fi % 2) * (fw + Inches(0.20))
    by_f = cy + (fi // 2) * Inches(0.44)
    text(sl, bx_f, by_f, Inches(0.22), Inches(0.36),
         "→", size=11, bold=True, color=TEAL_D)
    text(sl, bx_f + Inches(0.24), by_f, fw - Inches(0.28), Inches(0.36),
         ft, size=10, color=TEAL_D)


# ── Save ─────────────────────────────────────────────────────
prs.save('slides2.pptx')
print(f"slides2.pptx saved — {len(prs.slides)} slides")
