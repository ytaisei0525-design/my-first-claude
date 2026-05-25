#!/usr/bin/env python3
"""slides2.pptx — Wang et al., Biomaterials 331 (2026) 124114
   40 slides · Clean & Minimal · all paper data · varied layouts
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ── Palette ──────────────────────────────────────────────────
WHITE  = RGBColor(0xff,0xff,0xff); CARD  = RGBColor(0xf5,0xf7,0xfa)
CARD2  = RGBColor(0xeb,0xf0,0xf5); BORDER= RGBColor(0xd6,0xdf,0xe8)
TEXT   = RGBColor(0x18,0x22,0x2e); TEXT2 = RGBColor(0x3e,0x4a,0x58)
MUTED  = RGBColor(0x82,0x90,0xa0); TEAL  = RGBColor(0x07,0x6e,0x8a)
TEAL_L = RGBColor(0xe0,0xf2,0xf8); TEAL_D= RGBColor(0x05,0x50,0x68)
GREEN  = RGBColor(0x18,0x84,0x50); GREEN_L=RGBColor(0xd8,0xf0,0xe4)
RED    = RGBColor(0xb0,0x28,0x28); RED_L  =RGBColor(0xf8,0xe0,0xe0)
AMBER  = RGBColor(0xbf,0x62,0x06); AMBER_L=RGBColor(0xfb,0xef,0xd8)
SLATE  = RGBColor(0x38,0x50,0x66); GRAY_L =RGBColor(0xf0,0xf3,0xf6)
GRAY_M = RGBColor(0xb0,0xbc,0xc8)

# ── Dimensions ───────────────────────────────────────────────
W=Inches(13.333); H=Inches(7.5); PAD=Inches(0.45)
HDR_H=Inches(0.76); FTR_H=Inches(0.28)
BODY_Y=HDR_H+Inches(0.22); BODY_H=H-BODY_Y-FTR_H-Inches(0.12)
BODY_X=PAD; BODY_W=W-PAD*2
FONT_SCALE=1.12; FONT_NAME='Arial'

# ── Primitives ───────────────────────────────────────────────
def rect(sl,x,y,w,h,fill,border=None,bpt=0.5):
    s=sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,x,y,w,h)
    s.fill.solid(); s.fill.fore_color.rgb=fill
    if border: s.line.color.rgb=border; s.line.width=Pt(bpt)
    else: s.line.fill.background()
    s.shadow.inherit=False; return s

def rrect(sl,x,y,w,h,fill,border=None,bpt=0.5,rad=0.04):
    s=sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,x,y,w,h)
    s.adjustments[0]=rad; s.fill.solid(); s.fill.fore_color.rgb=fill
    if border: s.line.color.rgb=border; s.line.width=Pt(bpt)
    else: s.line.fill.background()
    s.shadow.inherit=False; return s

def text(sl,x,y,w,h,content,size=11,bold=False,color=TEXT2,
         align=PP_ALIGN.LEFT,italic=False,anchor=MSO_ANCHOR.TOP):
    if not content: return None
    s=str(content).strip()
    if not s: return None
    tb=sl.shapes.add_textbox(x,y,w,h); tb.word_wrap=True
    tf=tb.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    tf.margin_left=Emu(36000); tf.margin_right=Emu(36000)
    tf.margin_top=Emu(18000); tf.margin_bottom=Emu(18000)
    for i,line in enumerate(s.split('\n')):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment=align; run=p.add_run(); run.text=line
        run.font.size=Pt(size*FONT_SCALE); run.font.bold=bold
        run.font.italic=italic; run.font.color.rgb=color; run.font.name=FONT_NAME
    return tb

def fig(sl,x,y,w,h,cap=""):
    rrect(sl,x,y,w,h,GRAY_L,BORDER,0.4)
    text(sl,x,y+h*0.38,w,h*0.24,"[ Figure ]",size=9,color=GRAY_M,align=PP_ALIGN.CENTER)
    if cap: text(sl,x+Inches(0.1),y+h-Inches(0.22),w-Inches(0.2),Inches(0.20),
                 cap,size=7.5,color=MUTED,align=PP_ALIGN.CENTER,italic=True)

def rule(sl,x,y,w,c=BORDER):
    rect(sl,x,y,w,Inches(0.015),c)

def card(sl,x,y,w,h,title=None,body=None,bullets=None,ts=12,bs=10,ac=None):
    rrect(sl,x,y,w,h,CARD,BORDER,0.5)
    if ac: rect(sl,x,y,Inches(0.05),h,ac)
    cy=y+Inches(0.15); px=x+Inches(0.18); cw=w-Inches(0.36)
    if title:
        text(sl,px,cy,cw,Inches(0.30),title,size=ts,bold=True,color=TEXT)
        cy+=Inches(0.33)
    if body: text(sl,px,cy,cw,h-(cy-y)-Inches(0.08),body,size=bs,color=TEXT2)
    elif bullets:
        for b in bullets:
            if cy+Inches(0.23)>y+h-Inches(0.05): break
            text(sl,px,cy,Inches(0.16),Inches(0.24),"·",size=bs,bold=True,color=ac or TEAL)
            text(sl,px+Inches(0.18),cy,cw-Inches(0.18),Inches(0.24),b,size=bs,color=TEXT2)
            cy+=Inches(0.25)

def bar_row(sl,x,y,w,h,lbl,pct,c=TEAL,lw=Inches(1.8),vw=Inches(0.65)):
    bw=w-lw-vw-Inches(0.08)
    text(sl,x,y,lw,h,lbl,size=9.5,color=TEXT2,anchor=MSO_ANCHOR.MIDDLE)
    rect(sl,x+lw+Inches(0.04),y+h*0.28,bw,h*0.44,BORDER)
    rect(sl,x+lw+Inches(0.04),y+h*0.28,max(Inches(0.02),bw*pct/100),h*0.44,c)
    text(sl,x+lw+bw+Inches(0.10),y,vw,h,f"{pct:.1f}%",size=10,bold=True,
         color=c,anchor=MSO_ANCHOR.MIDDLE)

def hdr(sl,sec,title):
    rect(sl,0,0,W,HDR_H,WHITE)
    rect(sl,0,0,Inches(0.07),HDR_H,TEAL)
    rule(sl,0,HDR_H-Inches(0.015),W)
    if sec: text(sl,Inches(0.20),Inches(0.10),Inches(8),Inches(0.22),
                 sec.upper(),size=8,bold=True,color=TEAL)
    ty=Inches(0.34) if sec else Inches(0.18)
    text(sl,Inches(0.20),ty,BODY_W,Inches(0.36),title,size=18,bold=True,color=TEXT)

def ftr(sl,num):
    rect(sl,0,H-FTR_H,W,FTR_H,GRAY_L); rule(sl,0,H-FTR_H,W)
    text(sl,PAD,H-FTR_H+Inches(0.04),Inches(9),Inches(0.20),
         "Wang et al. | Biomaterials 331 (2026) 124114",size=7.5,color=MUTED)
    if num: text(sl,W-Inches(1.5),H-FTR_H+Inches(0.04),Inches(1.3),Inches(0.20),
                 str(num),size=7.5,color=MUTED,align=PP_ALIGN.RIGHT)

def sl_new(prs,sec="",title="",num=None):
    sl=prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl,0,0,W,H,WHITE); hdr(sl,sec,title); ftr(sl,num); return sl

def sl_div(prs,num,ns,title,sub=""):
    sl=prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl,0,0,W,H,WHITE)
    rect(sl,0,0,W,Inches(0.56),TEAL)
    ftr(sl,num)
    text(sl,PAD,Inches(0.90),Inches(4),Inches(0.28),ns.upper(),
         size=9,bold=True,color=MUTED)
    text(sl,PAD,Inches(1.22),W-PAD*2,Inches(1.10),title,
         size=38,bold=True,color=TEXT)
    if sub:
        rule(sl,PAD,Inches(2.56),Inches(0.04),TEAL)
        text(sl,PAD+Inches(0.20),Inches(2.48),W-PAD*2-Inches(0.24),
             Inches(0.36),sub,size=12,color=TEXT2)
    return sl


# ════════════════════════════════════════════════════════════
prs=Presentation(); prs.slide_width=W; prs.slide_height=H

# ── 01 COVER ─────────────────────────────────────────────────
sl=prs.slides.add_slide(prs.slide_layouts[6])
rect(sl,0,0,W,H,WHITE)
rect(sl,0,0,Inches(0.10),H,TEAL)
rect(sl,0,H-Inches(0.08),W,Inches(0.08),TEAL)
ftr(sl,1)
rule(sl,Inches(0.24),Inches(1.80),Inches(10.0),TEAL)
text(sl,Inches(0.24),Inches(1.96),Inches(11.0),Inches(1.70),
     "Co-assembly of Zwitterionic and Cationic Polymers\n"
     "for Antibacterial and Antithrombotic Surfaces\n"
     "via Weak Electrostatic Interactions",
     size=27,bold=True,color=TEXT)
text(sl,Inches(0.24),Inches(3.78),Inches(10),Inches(0.34),
     "Jiaqi Wang, Jingze Liu, Yuan Wei, Tong Wang, Yage Hu, Yao Xiong, Rifang Luo, Fanjun Zhang, Yunbing Wang",
     size=10,color=TEXT2)
text(sl,Inches(0.24),Inches(4.18),Inches(10),Inches(0.26),
     "Biomaterials  331 (2026) 124114  ·  Sichuan University / Chinese Academy of Medical Sciences",
     size=9,color=MUTED)
rule(sl,Inches(0.24),Inches(4.56),Inches(10.0))
# Hero stats
sv=[("3.1%","Thrombus coverage\nvs PLA 99.0%",GREEN),
    (">97.6%","Bactericidal eff.\nvs S. aureus",TEAL),
    (">99%","In vivo bacterial\neradication",TEAL_D),
    ("<2%","Hemolysis\n(non-hemolytic)",GREEN)]
sw=(Inches(10.0)-Inches(0.45)*3)/4
for i,(v,l,c) in enumerate(sv):
    bx=Inches(0.24)+i*(sw+Inches(0.45))
    text(sl,bx,Inches(4.70),sw,Inches(0.60),v,size=30,bold=True,color=c)
    text(sl,bx,Inches(5.34),sw,Inches(0.60),l,size=9,color=TEXT2)

# ── 02 TABLE OF CONTENTS ─────────────────────────────────────
sl=sl_new(prs,"","Table of Contents",2)
secs=[("01","Background & Motivation",3),
      ("02","Materials & Methods",7),
      ("03","Solution Interactions",9),
      ("04","Coating Properties",14),
      ("05","Antifouling & Antithrombotic",21),
      ("06","Antibacterial Performance",24),
      ("07","Biocompatibility & In Vivo",29),
      ("08","Conclusions",38)]
cw2=(BODY_W-Inches(0.30))/2
for i,(n,s,p) in enumerate(secs):
    col=i%2; row=i//2
    bx=BODY_X+col*(cw2+Inches(0.30))
    by=BODY_Y+Inches(0.10)+row*Inches(1.36)
    rrect(sl,bx,by,cw2,Inches(1.22),CARD,BORDER,0.5)
    rect(sl,bx,by,Inches(0.06),Inches(1.22),TEAL)
    rrect(sl,bx+Inches(0.18),by+Inches(0.30),Inches(0.44),Inches(0.26),TEAL_L,None,rad=0.10)
    text(sl,bx+Inches(0.18),by+Inches(0.30),Inches(0.44),Inches(0.26),
         n,size=8,bold=True,color=TEAL_D,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    text(sl,bx+Inches(0.74),by+Inches(0.20),cw2-Inches(0.90),Inches(0.36),
         s,size=12,bold=True,color=TEXT)
    text(sl,bx+Inches(0.74),by+Inches(0.60),cw2-Inches(0.90),Inches(0.48),
         f"→ slide {p}",size=9,color=MUTED,italic=True)

# ── 03 DIV: BACKGROUND ───────────────────────────────────────
sl_div(prs,3,"Section 01","Background & Motivation",
       "Clinical need for dual-function coatings on implantable medical devices")

# ── 04 CLINICAL PROBLEM ──────────────────────────────────────
sl=sl_new(prs,"01 · Background","The Problem: Infection and Thrombosis on Medical Devices",4)
# Left wide panel
text(sl,BODY_X,BODY_Y+Inches(0.10),Inches(7.4),Inches(0.72),
     "Device-associated infections and thrombotic complications\nremain major causes of implant failure.",
     size=16,bold=True,color=TEXT)
rule(sl,BODY_X,BODY_Y+Inches(0.92),Inches(7.4))
pts=[("Incidence","Device-associated infections exceed 20% in ICU patients; biofilm on implant surfaces resists systemic antibiotics."),
     ("Thrombosis","Platelet activation and thrombus formation begin within seconds of blood contact with foreign materials."),
     ("Resistance","Antibiotic-releasing coatings accelerate resistance; short drug-release window leaves devices unprotected."),
     ("Conflict","Cationic coatings kill bacteria but promote thrombosis; zwitterionic coatings are antifouling but not antibacterial.")]
for i,(kw,body) in enumerate(pts):
    by2=BODY_Y+Inches(1.06)+i*Inches(1.28)
    text(sl,BODY_X,by2,Inches(1.40),Inches(0.30),kw,size=10,bold=True,color=TEAL)
    text(sl,BODY_X+Inches(1.48),by2,Inches(5.80),Inches(0.50),body,size=10,color=TEXT2)
    if i<3: rule(sl,BODY_X,by2+Inches(0.62),Inches(7.4),GRAY_L)
# Right stats panel
rect(sl,Inches(8.10),BODY_Y+Inches(0.10),Inches(4.76),BODY_H,CARD,BORDER,0.5)
rect(sl,Inches(8.10),BODY_Y+Inches(0.10),Inches(4.76),Inches(0.04),TEAL)
text(sl,Inches(8.30),BODY_Y+Inches(0.22),Inches(4.36),Inches(0.28),
     "Scale of the Problem",size=11,bold=True,color=TEXT)
big=[("≥ 20%","of ICU admissions develop device-associated infections","source: Ref [11]",RED),
     ("50–70%","of implant failures linked to biofilm or thrombosis","combined pathologies",AMBER),
     ("$ billions","annual healthcare cost of implant-related infections","USA estimate",SLATE)]
for i,(v,l,s,c) in enumerate(big):
    by3=BODY_Y+Inches(0.72)+i*Inches(1.62)
    text(sl,Inches(8.30),by3,Inches(4.36),Inches(0.60),v,size=28,bold=True,color=c)
    text(sl,Inches(8.30),by3+Inches(0.62),Inches(4.36),Inches(0.30),l,size=10,color=TEXT2)
    text(sl,Inches(8.30),by3+Inches(0.96),Inches(4.36),Inches(0.22),s,size=8,color=MUTED,italic=True)
    if i<2: rule(sl,Inches(8.30),by3+Inches(1.38),Inches(4.16),BORDER)

# ── 05 POLYMER TRIO ───────────────────────────────────────────
sl=sl_new(prs,"01 · Background","Three Polymers: Structure and Roles",5)
polys=[
  ("PA","poly(SBMA)","Zwitterionic",TEAL,
   "SO₃⁻ and N⁺ within same repeat unit → electrically neutral\n"
   "Forms dense hydration shell → antifouling & antithrombotic\n"
   "Weak electrostatic interaction with cationic PQ\n"
   "PA-PQ loose complex: PQ remains functionally accessible"),
  ("PS","poly(SMPS)","Anionic",SLATE,
   "SO₃⁻ only → permanently negatively charged\n"
   "Strong electrostatic attraction to cationic PQ\n"
   "PS-PQ tight complex: PQ buried inside aggregate\n"
   "Comparison system for the strong-interaction case"),
  ("PQ","poly(QAS-8C)","Cationic",AMBER,
   "Quaternary ammonium groups → permanent positive charge\n"
   "Hydrophobic C8 chains enhance membrane disruption\n"
   "Bactericidal by disrupting bacterial cell membranes\n"
   "Interaction with PA/PS controls its surface accessibility"),
]
cw3=(BODY_W-Inches(0.40))/3
for i,(abbr,full,cat,c,body) in enumerate(polys):
    bx=BODY_X+i*(cw3+Inches(0.20))
    rect(sl,bx,BODY_Y+Inches(0.10),cw3,Inches(0.40),c)
    text(sl,bx+Inches(0.16),BODY_Y+Inches(0.10),cw3-Inches(0.32),Inches(0.40),
         abbr,size=20,bold=True,color=WHITE,anchor=MSO_ANCHOR.MIDDLE)
    rrect(sl,bx,BODY_Y+Inches(0.50),cw3,Inches(5.14),CARD,BORDER,0.5)
    text(sl,bx+Inches(0.16),BODY_Y+Inches(0.62),cw3-Inches(0.32),Inches(0.26),
         f"{full}  ·  {cat}",size=9,bold=True,color=c)
    fig(sl,bx+Inches(0.16),BODY_Y+Inches(0.98),cw3-Inches(0.32),Inches(1.40),"Chemical structure")
    text(sl,bx+Inches(0.16),BODY_Y+Inches(2.50),cw3-Inches(0.32),Inches(2.94),
         body,size=10,color=TEXT2)

# ── 06 DESIGN HYPOTHESIS ─────────────────────────────────────
sl=sl_new(prs,"01 · Background","Design Hypothesis: Weak Interaction Enables Dual Function",6)
# Two panels side by side
pw=(BODY_W-Inches(0.30))/2
# PS-PQ (strong)
rrect(sl,BODY_X,BODY_Y+Inches(0.10),pw,Inches(4.60),RED_L,RED,0.5)
text(sl,BODY_X+Inches(0.20),BODY_Y+Inches(0.22),pw-Inches(0.40),Inches(0.30),
     "PS-PQ  ·  STRONG Interaction",size=12,bold=True,color=RED)
fig(sl,BODY_X+Inches(0.20),BODY_Y+Inches(0.62),pw-Inches(0.40),Inches(2.20),"Fig 1D left — compact PS-PQ complex")
pts_s=["SO₃⁻ tightly binds quaternary N⁺ of PQ",
       "PQ buried deep inside aggregate",
       "Quaternary ammonium inaccessible to bacteria",
       "PS2.5 loses antibacterial function"]
cy=BODY_Y+Inches(3.00)
for p in pts_s:
    text(sl,BODY_X+Inches(0.20),cy,Inches(0.18),Inches(0.24),"✗",size=10,bold=True,color=RED)
    text(sl,BODY_X+Inches(0.40),cy,pw-Inches(0.60),Inches(0.24),p,size=10,color=RED)
    cy+=Inches(0.26)
# PA-PQ (weak)
bx2=BODY_X+pw+Inches(0.30)
rrect(sl,bx2,BODY_Y+Inches(0.10),pw,Inches(4.60),GREEN_L,GREEN,0.5)
text(sl,bx2+Inches(0.20),BODY_Y+Inches(0.22),pw-Inches(0.40),Inches(0.30),
     "PA-PQ  ·  WEAK Interaction",size=12,bold=True,color=GREEN)
fig(sl,bx2+Inches(0.20),BODY_Y+Inches(0.62),pw-Inches(0.40),Inches(2.20),"Fig 1D right — loose PA-PQ complex")
pts_w=["COO⁻ / SO₃⁻ of PA interacts weakly with PQ",
       "PQ retains 'freedom' within loose aggregate",
       "Quaternary ammonium accessible → bactericidal",
       "PA2.5 achieves dual antibacterial + antifouling"]
cy=BODY_Y+Inches(3.00)
for p in pts_w:
    text(sl,bx2+Inches(0.20),cy,Inches(0.18),Inches(0.24),"✓",size=10,bold=True,color=GREEN)
    text(sl,bx2+Inches(0.40),cy,pw-Inches(0.60),Inches(0.24),p,size=10,color=GREEN)
    cy+=Inches(0.26)
# Bottom key statement
rect(sl,BODY_X,BODY_Y+Inches(4.82),BODY_W,Inches(0.76),TEAL_L)
rule(sl,BODY_X,BODY_Y+Inches(4.82),BODY_W,TEAL)
text(sl,BODY_X+Inches(0.30),BODY_Y+Inches(4.96),BODY_W-Inches(0.60),Inches(0.48),
     "Key insight: the weaker the zwitterion–polycation interaction, the more "
     "freedom the cationic groups retain — preserving both bactericidal and antifouling function.",
     size=11,color=TEAL_D)

# ── 07 DIV: METHODS ──────────────────────────────────────────
sl_div(prs,7,"Section 02","Materials & Methods",
       "Free-radical polymerization · LbL self-assembly · in vitro & in vivo evaluation")

# ── 08 LbL PROCESS + NOMENCLATURE ────────────────────────────
sl=sl_new(prs,"02 · Methods","LbL Fabrication Process and Sample Nomenclature",8)
# Process steps top row
steps=[("① Pretreat","PLA or PU substrate\ncleaned with ethanol\nand DI water"),
       ("② DOPA coat","Oxidized DOPA solution\n2 mg/mL, pH 8.5\n2 h pretreatment"),
       ("③ Base layers","Alternate PEI/PS or\nPEI/PA dipping\n(4 base layers)"),
       ("④ Active layers","Alternate PA or PS +\nPQ solution dipping\n10 min each step"),
       ("⑤ Dry & store","40°C oven drying\nStored dry until\nexperiments")]
sw2=(BODY_W-Inches(0.40))/5
by2=BODY_Y+Inches(0.10)
for i,(s,d) in enumerate(steps):
    bx=BODY_X+i*(sw2+Inches(0.10))
    rrect(sl,bx,by2,sw2,Inches(1.60),CARD,BORDER,0.5)
    rect(sl,bx,by2,sw2,Inches(0.04),TEAL)
    text(sl,bx+Inches(0.10),by2+Inches(0.12),sw2-Inches(0.20),Inches(0.28),
         s,size=10,bold=True,color=TEAL)
    text(sl,bx+Inches(0.10),by2+Inches(0.46),sw2-Inches(0.20),Inches(1.06),
         d,size=9,color=TEXT2)
# Nomenclature table
by3=BODY_Y+Inches(1.88)
rule(sl,BODY_X,by3,BODY_W,TEAL)
text(sl,BODY_X,by3+Inches(0.08),BODY_W,Inches(0.28),
     "Sample Nomenclature",size=12,bold=True,color=TEXT)
cols=[BODY_X,BODY_X+Inches(1.40),BODY_X+Inches(3.20),BODY_X+Inches(5.00),
      BODY_X+Inches(7.00),BODY_X+Inches(9.00),BODY_X+Inches(10.80)]
cws=[Inches(1.30),Inches(1.70),Inches(1.70),Inches(1.90),Inches(1.90),Inches(1.70),Inches(2.00)]
hdrs=["Sample","Layers","Top layer","PA/PS conc.","PQ conc.","PA/PS type","Key character"]
rows_t=[["PA2.0","8","PQ (cationic)","2 mg/mL","0.5 mg/mL","Zwitterionic","PQ top → antibacterial"],
        ["PA2.5","9","PA (zwitterionic)","2 mg/mL","1.0 mg/mL","Zwitterionic","Dual: AB + antifouling ★"],
        ["PS2.0","8","PQ (cationic)","2 mg/mL","0.5 mg/mL","Anionic","PQ top → antibacterial"],
        ["PS2.5","9","PS (anionic)","2 mg/mL","1.0 mg/mL","Anionic","Shielded PQ → poor AB"],
        ["PA5.0","14","PQ (cationic)","5 mg/mL","1.0 mg/mL","Zwitterionic","More layers, PQ top"],
        ["PA5.5","15","PA (zwitterionic)","5 mg/mL","1.5 mg/mL","Zwitterionic","More layers, dual"],
        ["PS5.0","14","PQ (cationic)","5 mg/mL","1.0 mg/mL","Anionic","More layers, PQ top"],
        ["PS5.5","15","PS (anionic)","5 mg/mL","1.5 mg/mL","Anionic","Shielded PQ → poor AB"]]
hy2=by3+Inches(0.44)
for j,(h,cw) in enumerate(zip(hdrs,cws)):
    text(sl,cols[j],hy2,cw,Inches(0.26),h,size=8.5,bold=True,color=MUTED)
for ri,row in enumerate(rows_t):
    ry=hy2+Inches(0.30)+ri*Inches(0.38)
    bg=TEAL_L if "★" in (row[-1]) else (CARD if ri%2==0 else WHITE)
    rect(sl,BODY_X,ry-Inches(0.03),BODY_W,Inches(0.36),bg)
    for j,(val,cw) in enumerate(zip(row,cws)):
        fc=TEAL_D if "★" in val else (TEXT if j==0 else TEXT2)
        text(sl,cols[j],ry,cw,Inches(0.30),val,size=9,
             bold=("★" in val or j==0),color=fc,anchor=MSO_ANCHOR.MIDDLE)

# ── 09 DIV: SOLUTION INTERACTIONS ────────────────────────────
sl_div(prs,9,"Section 03","Solution Interactions",
       "ITC · turbidity · viscosity · DLS · fluorescence quenching · zeta potential")

# ── 10 TURBIDITY & AGGREGATION ────────────────────────────────
sl=sl_new(prs,"03 · Solution Interactions","Turbidity and Aggregate Formation",10)
fig(sl,BODY_X,BODY_Y+Inches(0.10),Inches(7.60),Inches(5.54),
    "Fig 1B,C — Absorbance at 550 nm vs PQ volume / turbidity photos before & after equilibrium")
bx2=BODY_X+Inches(7.90)
text(sl,bx2,BODY_Y+Inches(0.14),Inches(5.00),Inches(0.32),
     "Turbidity Results",size=13,bold=True,color=TEXT)
rule(sl,bx2,BODY_Y+Inches(0.52),Inches(5.00))
items=[("PS-PQ",SLATE,
        ["Absorbance ↑ to 1.79 at 1.5 eq PQ",
         "Turbidity stable after equilibrium",
         "Compact, insoluble aggregates formed",
         "Strong anion–cation interaction"]),
       ("PA-PQ",TEAL,
        ["Absorbance ↑ then drops to 1.33",
         "PA-PQ clarifies at equilibrium",
         "Loose aggregates easily disrupted",
         "Weak electrostatic interaction confirmed"])]
cy=BODY_Y+Inches(0.64)
for name,c,buls in items:
    text(sl,bx2,cy,Inches(5.00),Inches(0.28),name,size=11,bold=True,color=c)
    cy+=Inches(0.32)
    for b in buls:
        text(sl,bx2,cy,Inches(0.16),Inches(0.24),"·",size=10,bold=True,color=c)
        text(sl,bx2+Inches(0.20),cy,Inches(4.72),Inches(0.24),b,size=10,color=TEXT2)
        cy+=Inches(0.27)
    cy+=Inches(0.18)
# Interpretation box
rrect(sl,bx2,BODY_Y+Inches(3.64),Inches(5.00),Inches(2.00),TEAL_L,TEAL,0.5)
text(sl,bx2+Inches(0.18),BODY_Y+Inches(3.80),Inches(4.64),Inches(1.70),
     "The weak electrostatic interaction in PA-PQ allows water molecules to "
     "penetrate the aggregate, disrupting the complex and leaving both PA and "
     "PQ functional groups exposed — the mechanistic basis for dual function.",
     size=10,color=TEAL_D)

# ── 11 ITC THERMODYNAMICS ─────────────────────────────────────
sl=sl_new(prs,"03 · Solution Interactions","ITC: Thermodynamics of Co-assembly",11)
fig(sl,BODY_X,BODY_Y+Inches(0.10),Inches(7.60),Inches(2.60),"Fig 1E — ΔG, ΔH, −TΔS bar chart for PS-PQ and PA-PQ")
fig(sl,BODY_X,BODY_Y+Inches(2.82),Inches(7.60),Inches(2.82),"Fig 1F — Cumulative heat vs volume of PS or PA added to PQ")
bx2=BODY_X+Inches(7.90)
text(sl,bx2,BODY_Y+Inches(0.14),Inches(5.00),Inches(0.32),
     "ITC Key Findings",size=13,bold=True,color=TEXT)
rule(sl,bx2,BODY_Y+Inches(0.52),Inches(5.00))
thermo=[("ΔG < 0","Both PA-PQ and PS-PQ: spontaneous co-assembly at 25°C",TEAL),
        ("ΔH (PS-PQ)","Large negative enthalpy — enthalpically driven, strong binding",SLATE),
        ("ΔH (PA-PQ)","Smaller enthalpy change — weaker binding, entropy contribution",TEAL),
        ("−TΔS","Entropy term larger for PA-PQ, consistent with loose/dynamic complex",MUTED),
        ("PS vs PA","PS-PQ enthalpy change much larger at first drip → charge density drives interaction strength",SLATE)]
cy=BODY_Y+Inches(0.64)
for kw,desc,c in thermo:
    rrect(sl,bx2,cy,Inches(5.00),Inches(0.92),WHITE,BORDER,0.3)
    text(sl,bx2+Inches(0.14),cy+Inches(0.10),Inches(1.60),Inches(0.28),kw,size=9.5,bold=True,color=c)
    text(sl,bx2+Inches(0.14),cy+Inches(0.40),Inches(4.66),Inches(0.44),desc,size=9,color=TEXT2)
    cy+=Inches(1.02)

# ── 12 VISCOSITY + PARTICLE SIZE ─────────────────────────────
sl=sl_new(prs,"03 · Solution Interactions","Viscosity and Aggregate Particle Size",12)
pw2=(BODY_W-Inches(0.30))/2
fig(sl,BODY_X,BODY_Y+Inches(0.10),pw2,Inches(3.50),"Fig 1G — Viscosity (mPa·s) vs shear rate (1/s)")
fig(sl,BODY_X+pw2+Inches(0.30),BODY_Y+Inches(0.10),pw2,Inches(3.50),"Fig 1J — DLS particle size distribution before/after equilibrium")
# Viscosity annotations
rrect(sl,BODY_X,BODY_Y+Inches(3.74),pw2,Inches(1.90),CARD,BORDER,0.5)
text(sl,BODY_X+Inches(0.18),BODY_Y+Inches(3.88),pw2-Inches(0.36),Inches(0.28),
     "Viscosity Interpretation",size=11,bold=True,color=TEXT)
visc_pts=["PQ ~2, PS ~8, PA ~4 mPa·s (individual)",
          "PS-PQ: viscosity decreases vs PS alone",
          "PA-PQ: viscosity increases vs PA alone",
          "Gel-like PA-PQ network → loose cross-linking",
          "Confirms weak but stable co-assembly"]
cy=BODY_Y+Inches(4.22)
for p in visc_pts:
    text(sl,BODY_X+Inches(0.18),cy,Inches(0.16),Inches(0.22),"·",size=9.5,bold=True,color=TEAL)
    text(sl,BODY_X+Inches(0.36),cy,pw2-Inches(0.54),Inches(0.22),p,size=9.5,color=TEXT2); cy+=Inches(0.24)
# DLS annotations
bx2=BODY_X+pw2+Inches(0.30)
rrect(sl,bx2,BODY_Y+Inches(3.74),pw2,Inches(1.90),CARD,BORDER,0.5)
text(sl,bx2+Inches(0.18),BODY_Y+Inches(3.88),pw2-Inches(0.36),Inches(0.28),
     "Particle Size (DLS)",size=11,bold=True,color=TEXT)
dls_pts=["PS-PQ: size increases monotonically with concentration",
         "PA-PQ: size changes significantly before vs after equilibrium",
         "PA-PQ equilibrium: smaller, more stable loose aggregates",
         "Charge inversion observed in PS-PQ at high PQ (10:30 ratio)",
         "Consistent with turbidity and ITC data"]
cy=BODY_Y+Inches(4.22)
for p in dls_pts:
    text(sl,bx2+Inches(0.18),cy,Inches(0.16),Inches(0.22),"·",size=9.5,bold=True,color=SLATE)
    text(sl,bx2+Inches(0.36),cy,pw2-Inches(0.54),Inches(0.22),p,size=9.5,color=TEXT2); cy+=Inches(0.24)

# ── 13 FLUORESCENCE + ZETA (SOLUTION) ────────────────────────
sl=sl_new(prs,"03 · Solution Interactions","Fluorescence Quenching and Zeta Potential in Solution",13)
fig(sl,BODY_X,BODY_Y+Inches(0.10),Inches(6.60),Inches(2.70),"Fig 1H — Fluorescence microscopy images PS-PQ vs PA-PQ (scale 200 µm)")
fig(sl,BODY_X,BODY_Y+Inches(2.94),Inches(6.60),Inches(2.70),"Fig 1I — Zeta potential vs concentration ratio PS-PQ and PA-PQ")
bx2=BODY_X+Inches(6.90)
# Fluorescence stats as hero numbers (no box, just typography)
text(sl,bx2,BODY_Y+Inches(0.14),Inches(5.76),Inches(0.26),
     "FLUORESCENCE QUENCHING",size=8,bold=True,color=MUTED)
text(sl,bx2,BODY_Y+Inches(0.44),Inches(2.50),Inches(0.62),"34.25",size=38,bold=True,color=SLATE)
text(sl,bx2+Inches(2.60),BODY_Y+Inches(0.44),Inches(3.10),Inches(0.62),"43.59",size=38,bold=True,color=TEAL)
text(sl,bx2,BODY_Y+Inches(1.08),Inches(2.50),Inches(0.24),"PS-PQ avg intensity",size=9,color=SLATE)
text(sl,bx2+Inches(2.60),BODY_Y+Inches(1.08),Inches(3.10),Inches(0.24),"PA-PQ avg intensity",size=9,color=TEAL)
text(sl,bx2,BODY_Y+Inches(1.36),Inches(5.76),Inches(0.54),
     "Lower PS-PQ intensity → stronger fluorescence quenching → tighter complex.\n"
     "Higher PA-PQ intensity → weaker quenching → looser, more open complex.",
     size=9.5,color=TEXT2)
rule(sl,bx2,BODY_Y+Inches(2.00),Inches(5.76))
text(sl,bx2,BODY_Y+Inches(2.14),Inches(5.76),Inches(0.26),
     "ZETA POTENTIAL (SOLUTION)",size=8,bold=True,color=MUTED)
zeta_pts=["PA-PQ initially negative (PA anionic character at pH 7.4)",
          "Zeta shifts positive as PQ concentration increases",
          "At high PQ: charge inversion → positively charged complexes",
          "PA-PQ aggregates incorporate water → larger, softer at equilibrium",
          "PS-PQ: consistently more negative due to strong SO₃⁻ / N⁺ binding"]
cy=BODY_Y+Inches(2.50)
for p in zeta_pts:
    text(sl,bx2,cy,Inches(0.16),Inches(0.24),"·",size=10,bold=True,color=TEAL)
    text(sl,bx2+Inches(0.20),cy,Inches(5.50),Inches(0.24),p,size=9.5,color=TEXT2); cy+=Inches(0.52)

# ── 14 DIV: COATING PROPERTIES ───────────────────────────────
sl_div(prs,14,"Section 04","Coating Properties",
       "QCM-d · hydration · zeta potential · XPS · AFM · WCA · stability")

# ── 15 QCM-D ASSEMBLY + HYDRATION ────────────────────────────
sl=sl_new(prs,"04 · Coating Properties","QCM-d: Real-Time Assembly and Hydration Capacity",15)
pw3=Inches(7.20)
fig(sl,BODY_X,BODY_Y+Inches(0.10),pw3,Inches(2.60),"Fig 2B — Δf (×100 Hz) vs time: PS-PQ (left) and PA-PQ (right) LbL assembly")
fig(sl,BODY_X,BODY_Y+Inches(2.84),pw3,Inches(2.80),"Fig 2H — ΔF and ΔD during ethanol/water alternation (hydration test) for PS-PQ and PA-PQ")
bx2=BODY_X+pw3+Inches(0.30)
rw=W-bx2-PAD
text(sl,bx2,BODY_Y+Inches(0.14),rw,Inches(0.28),"Assembly Kinetics",size=12,bold=True,color=TEXT)
rule(sl,bx2,BODY_Y+Inches(0.48),rw)
qcm=[("PA-PQ","Δf = 156 ppm","ΔD = 559 Hz","More mass · softer film · higher hydration",TEAL),
     ("PS-PQ","Δf = 89 ppm","ΔD = 225 Hz","Less mass · stiffer film · compact structure",SLATE)]
cy=BODY_Y+Inches(0.58)
for name,df,dD,note,c in qcm:
    rrect(sl,bx2,cy,rw,Inches(1.46),CARD,BORDER,0.5)
    rect(sl,bx2,cy,Inches(0.05),Inches(1.46),c)
    text(sl,bx2+Inches(0.18),cy+Inches(0.12),rw-Inches(0.28),Inches(0.26),name,size=11,bold=True,color=TEXT)
    text(sl,bx2+Inches(0.18),cy+Inches(0.44),rw-Inches(0.28),Inches(0.26),df,size=13,bold=True,color=c)
    text(sl,bx2+Inches(0.18),cy+Inches(0.70),rw-Inches(0.28),Inches(0.26),dD,size=13,bold=True,color=c)
    text(sl,bx2+Inches(0.18),cy+Inches(1.00),rw-Inches(0.28),Inches(0.36),note,size=9,color=TEXT2)
    cy+=Inches(1.60)
rule(sl,bx2,BODY_Y+Inches(3.76),rw)
text(sl,bx2,BODY_Y+Inches(3.90),rw,Inches(0.24),"Hydration Test",size=12,bold=True,color=TEXT)
text(sl,bx2,BODY_Y+Inches(4.22),rw,Inches(1.42),
     "When liquid switched from anhydrous to 90% ethanol:\n"
     "PA-PQ: Δf first increases then decreases → rapidly absorbs water, swells.\n"
     "PS-PQ: Δf responds less → compact structure, limited water uptake.\n"
     "Confirms strong hydration capacity of PA-PQ coating.",size=9.5,color=TEXT2)

# ── 16 ZETA POTENTIAL (SURFACE) + XPS ────────────────────────
sl=sl_new(prs,"04 · Coating Properties","Surface Zeta Potential and XPS Elemental Composition",16)
pw4=Inches(5.80)
fig(sl,BODY_X,BODY_Y+Inches(0.10),pw4,Inches(2.70),"Fig 2C — Zeta potential (mV) vs assembly layer number: PS2.5 and PA2.5")
fig(sl,BODY_X+pw4+Inches(0.30),BODY_Y+Inches(0.10),BODY_W-pw4-Inches(0.30),Inches(2.70),
    "Fig 2E,F — XPS N element% and S element% per sample")
# Zeta note
rrect(sl,BODY_X,BODY_Y+Inches(2.94),pw4,Inches(2.70),CARD,BORDER,0.5)
rect(sl,BODY_X,BODY_Y+Inches(2.94),pw4,Inches(0.04),TEAL)
text(sl,BODY_X+Inches(0.18),BODY_Y+Inches(3.08),pw4-Inches(0.36),Inches(0.28),
     "Zeta Potential (Surface)",size=11,bold=True,color=TEXT)
zs=["Zeta oscillates +/− with each layer (alternating PQ/PA deposition)",
    "When PQ is top layer (PA2.0, PS2.0): positive ζ",
    "When PA is top layer (PA2.5): near-neutral → antifouling",
    "When PS is top layer (PS2.5): negative ζ",
    "Layer-by-layer buildup confirmed by alternating charge"]
cy=BODY_Y+Inches(3.44)
for p in zs:
    text(sl,BODY_X+Inches(0.18),cy,Inches(0.16),Inches(0.22),"·",size=9.5,bold=True,color=TEAL)
    text(sl,BODY_X+Inches(0.36),cy,pw4-Inches(0.54),Inches(0.22),p,size=9.5,color=TEXT2); cy+=Inches(0.26)
# XPS note
bx3=BODY_X+pw4+Inches(0.30)
rw2=BODY_W-pw4-Inches(0.30)
rrect(sl,bx3,BODY_Y+Inches(2.94),rw2,Inches(2.70),CARD,BORDER,0.5)
rect(sl,bx3,BODY_Y+Inches(2.94),rw2,Inches(0.04),AMBER)
text(sl,bx3+Inches(0.18),BODY_Y+Inches(3.08),rw2-Inches(0.36),Inches(0.28),
     "XPS Composition",size=11,bold=True,color=TEXT)
xps_pts=["PLA (bare): no N or S detected",
         "All coated surfaces: N% and S% detected → coating confirmed",
         "PA samples: higher N% (sulfobetaine N)",
         "PS samples: higher S% (sulfopropyl S)",
         "XPS verifies successful LbL assembly of both polymer types"]
cy=BODY_Y+Inches(3.44)
for p in xps_pts:
    text(sl,bx3+Inches(0.18),cy,Inches(0.16),Inches(0.22),"·",size=9.5,bold=True,color=AMBER)
    text(sl,bx3+Inches(0.36),cy,rw2-Inches(0.54),Inches(0.22),p,size=9.5,color=TEXT2); cy+=Inches(0.26)

# ── 17 AFM ROUGHNESS ─────────────────────────────────────────
sl=sl_new(prs,"04 · Coating Properties","AFM Topography and Surface Roughness",17)
fig(sl,BODY_X,BODY_Y+Inches(0.10),BODY_W,Inches(3.20),"Fig 2I — AFM 3D topography: PLA, PS2.0, PS2.5, PA2.0, PA2.5 (scale 1.8 µm × 1.8 µm)")
# Table
by3=BODY_Y+Inches(3.44)
rule(sl,BODY_X,by3,BODY_W,TEAL)
text(sl,BODY_X,by3+Inches(0.08),BODY_W,Inches(0.28),"Surface Roughness — AFM Parameters",size=11,bold=True,color=TEXT)
afm_cols=[BODY_X,BODY_X+Inches(2.0),BODY_X+Inches(4.4),BODY_X+Inches(6.8),BODY_X+Inches(9.2)]
afm_cws=[Inches(1.90),Inches(2.30),Inches(2.30),Inches(2.30),Inches(3.40)]
afm_hdrs=["Sample","Rq (nm)","Ra (nm)","Trend","Interpretation"]
afm_rows=[["PLA (bare)","1.06","0.67","—","Flat substrate baseline"],
           ["PA-PQ coatings","Decreases","Decreases","↓ per layer","Regions filled by PA+PQ; smoother"],
           ["PS-PQ coatings","Increases","Increases","↑ per layer","Island growth; uneven PQ distribution"]]
hy3=by3+Inches(0.44)
for j,(h,cw) in enumerate(zip(afm_hdrs,afm_cws)):
    text(sl,afm_cols[j],hy3,cw,Inches(0.24),h,size=9,bold=True,color=MUTED)
for ri,row in enumerate(afm_rows):
    ry=hy3+Inches(0.28)+ri*Inches(0.46)
    rect(sl,BODY_X,ry-Inches(0.03),BODY_W,Inches(0.44),CARD if ri%2==0 else WHITE)
    for j,(val,cw) in enumerate(zip(row,afm_cws)):
        c=TEAL_D if ri==0 and j in(1,2) else TEXT2
        text(sl,afm_cols[j],ry,cw,Inches(0.38),val,size=10,bold=(j==0),color=c,anchor=MSO_ANCHOR.MIDDLE)

# ── 18 WATER CONTACT ANGLE ────────────────────────────────────
sl=sl_new(prs,"04 · Coating Properties","Water Contact Angle: Static, Advancing, and Receding",18)
fig(sl,BODY_X,BODY_Y+Inches(0.10),Inches(7.20),Inches(3.90),"Fig 2G — WCA bar chart (static, advancing, receding angles) for all samples")
bx2=BODY_X+Inches(7.50)
rw3=W-bx2-PAD
text(sl,bx2,BODY_Y+Inches(0.14),rw3,Inches(0.28),"Key Values",size=12,bold=True,color=TEXT)
rule(sl,bx2,BODY_Y+Inches(0.48),rw3)
wca_d=[("PLA (bare)","88.73°","~82°","~65°",RED),
       ("PS2.0","~70°","~65°","~48°",SLATE),
       ("PS2.5","~65°","~60°","~44°",SLATE),
       ("PA2.0","53.3°","~46°","~28°",TEAL),
       ("PA2.5","21.58°","~18°","~8°",GREEN)]
wca_hdrs=["Static","Advancing","Receding"]
for j,h in enumerate(wca_hdrs):
    tx=bx2+Inches(1.40)+j*Inches(1.08)
    text(sl,tx,BODY_Y+Inches(0.54),Inches(1.04),Inches(0.22),h,size=8,bold=True,color=MUTED,align=PP_ALIGN.CENTER)
cy=BODY_Y+Inches(0.80)
for name,st,adv,rec,c in wca_d:
    rrect(sl,bx2,cy,rw3,Inches(0.80),CARD,BORDER,0.5)
    rect(sl,bx2,cy,Inches(0.05),Inches(0.80),c)
    text(sl,bx2+Inches(0.18),cy+Inches(0.24),Inches(1.10),Inches(0.36),name,size=9.5,bold=True,color=TEXT)
    for j,v in enumerate([st,adv,rec]):
        tx=bx2+Inches(1.38)+j*Inches(1.08)
        text(sl,tx,cy+Inches(0.18),Inches(1.04),Inches(0.44),v,size=14,bold=True,color=c,align=PP_ALIGN.CENTER)
    cy+=Inches(0.90)
text(sl,BODY_X,BODY_Y+Inches(4.16),Inches(7.20),Inches(1.50),
     "PA2.5 achieves WCA of 21.58° — highly hydrophilic surface. "
     "The strong hydration capacity originates from the zwitterionic sulfobetaine groups of PA, "
     "which form tight hydrogen bonds with water. PS-PQ coatings, despite anionic character, "
     "show less hydrophilicity because the strong PS-PQ interaction reduces free charged groups.",
     size=10,color=TEXT2)

# ── 19 COATING STABILITY ──────────────────────────────────────
sl=sl_new(prs,"04 · Coating Properties","Coating Stability: Fluorescence Intensity Decay in PBS",19)
fig(sl,BODY_X,BODY_Y+Inches(0.10),Inches(7.60),Inches(5.54),"Fig 2D — Fluorescence intensity of PS2.5 and PA2.5 coatings vs immersion time (days)")
bx2=BODY_X+Inches(7.90)
rw4=W-bx2-PAD
text(sl,bx2,BODY_Y+Inches(0.14),rw4,Inches(0.28),"Stability Findings",size=12,bold=True,color=TEXT)
rule(sl,bx2,BODY_Y+Inches(0.48),rw4)
stab=[("PA2.5",GREEN,
       ["Fluorescence decreases more rapidly on day 1",
        "Then stabilizes — long-term stability stronger",
        "Loose PA-PQ complex maintains structural integrity",
        "Coating does not fully fail within 10 days in PBS"]),
      ("PS2.5",SLATE,
       ["Fluorescence decreases steadily over time",
        "Less initial drop but more gradual decay",
        "Tighter complex more susceptible to dissociation?",
        "Both coatings retain function over the study period"])]
cy=BODY_Y+Inches(0.62)
for name,c,pts in stab:
    text(sl,bx2,cy,rw4,Inches(0.26),name,size=11,bold=True,color=c); cy+=Inches(0.30)
    for p in pts:
        text(sl,bx2,cy,Inches(0.16),Inches(0.22),"·",size=9.5,bold=True,color=c)
        text(sl,bx2+Inches(0.20),cy,rw4-Inches(0.20),Inches(0.22),p,size=9.5,color=TEXT2); cy+=Inches(0.26)
    cy+=Inches(0.22)
rrect(sl,bx2,BODY_Y+Inches(3.90),rw4,Inches(1.74),TEAL_L,TEAL,0.5)
text(sl,bx2+Inches(0.18),BODY_Y+Inches(4.06),rw4-Inches(0.36),Inches(1.48),
     "The PA2.5 coating demonstrates that the loose co-assembly "
     "structure does not compromise durability. The initial fluorescence "
     "decrease reflects initial rearrangement; subsequent plateau shows "
     "the coating reaches a stable equilibrium state in aqueous conditions.",
     size=10,color=TEAL_D)

# ── 20 DIV: ANTIFOULING ───────────────────────────────────────
sl_div(prs,20,"Section 05","Antifouling & Antithrombotic",
       "Protein, cell, and whole-blood adhesion · ex vivo blood circulation model")

# ── 21 PROTEIN & CELL ADHESION ───────────────────────────────
sl=sl_new(prs,"05 · Antifouling","Protein Adsorption and Cell Adhesion",21)
fig(sl,BODY_X,BODY_Y+Inches(0.10),(BODY_W-Inches(0.20))/2,Inches(2.60),"Fig 3A,B — FITC-BSA fluorescence images + quantitative area ratio (%)")
fig(sl,BODY_X+(BODY_W+Inches(0.20))/2,BODY_Y+Inches(0.10),(BODY_W-Inches(0.20))/2,Inches(2.60),
    "Fig 3C,D — L929 cell adhesion at 1d and 3d (CCK-8, scale 200 µm) + quantification")
# Interpretation below
by3=BODY_Y+Inches(2.84)
rule(sl,BODY_X,by3,BODY_W)
pts_adh=[("Protein (BSA-FITC)",TEAL,
          "PA2.5 dramatically reduces FITC-BSA adsorption. "
          "The hydration shell acts as an entropic barrier against protein adsorption. "
          "PS-PQ shows less reduction (weaker hydration)."),
         ("L929 Fibroblasts — 1 day",TEAL,
          "Both PA2.0 and PA2.5 significantly suppress fibroblast adhesion. "
          "PS2.0 shows slight reduction; PS2.5 similar to PLA."),
         ("L929 Fibroblasts — 3 days",GREEN,
          "PA2.5 maintains minimal cell adhesion at day 3. "
          "Sustained antifouling performance over extended periods. "
          "Consistent with WCA and QCM-d hydration data."),
         ("Mechanism",SLATE,
          "Zwitterionic sulfobetaine groups (PA) bind water tightly via electrostatic "
          "and H-bond interactions → dense hydration layer repels proteins and cells.")]
pw5=(BODY_W-Inches(0.30))/2
for i,(ttl,c,body) in enumerate(pts_adh):
    col=i%2; row=i//2
    bx=BODY_X+col*(pw5+Inches(0.30))
    by4=by3+Inches(0.16)+row*Inches(1.38)
    rrect(sl,bx,by4,pw5,Inches(1.24),CARD,BORDER,0.5)
    rect(sl,bx,by4,Inches(0.05),Inches(1.24),c)
    text(sl,bx+Inches(0.18),by4+Inches(0.12),pw5-Inches(0.28),Inches(0.26),ttl,size=10,bold=True,color=TEXT)
    text(sl,bx+Inches(0.18),by4+Inches(0.44),pw5-Inches(0.28),Inches(0.72),body,size=9.5,color=TEXT2)

# ── 22 WHOLE BLOOD + EX VIVO ANTITHROMBOTIC ──────────────────
sl=sl_new(prs,"05 · Antifouling","Whole-Blood Adhesion and Ex Vivo Antithrombotic Test",22)
# Top: whole blood
fig(sl,BODY_X,BODY_Y+Inches(0.10),Inches(6.60),Inches(1.90),"Fig 3E,F — Whole-blood cell adhesion fluorescence images + area ratio (%)")
# Bar charts
bx2=BODY_X+Inches(6.90)
text(sl,bx2,BODY_Y+Inches(0.10),Inches(5.96),Inches(0.26),"Thrombus Coverage (%)",size=10.5,bold=True,color=TEXT)
td=[("PLA",99.0,RED),("PS2.0",97.1,RED),("PS2.5",89.7,AMBER),("PA2.0",30.3,TEAL),("PA2.5",3.1,GREEN)]
cy=BODY_Y+Inches(0.44)
for nm,pct,c in td:
    bar_row(sl,bx2,cy,Inches(5.96),Inches(0.40),nm,pct,c,lw=Inches(1.3),vw=Inches(0.70)); cy+=Inches(0.44)
# Photo + mass
fig(sl,BODY_X,BODY_Y+Inches(2.14),Inches(6.60),Inches(1.90),"Fig 3H,I — Tube photographs after blood circulation + surface coverage ratio")
text(sl,bx2,BODY_Y+Inches(2.36),Inches(5.96),Inches(0.26),"Thrombus Mass Increase (%)",size=10.5,bold=True,color=TEXT)
md=[("PLA",152.9,RED),("PS2.0",90.3,AMBER),("PS2.5",82.2,AMBER),("PA2.0",52.5,TEAL),("PA2.5",51.7,GREEN)]
cy=BODY_Y+Inches(2.70)
for nm,pct,c in md:
    bar_row(sl,bx2,cy,Inches(5.96),Inches(0.38),nm,pct,c,lw=Inches(1.3),vw=Inches(0.70)); cy+=Inches(0.42)
# Bottom: SEM + note
fig(sl,BODY_X,BODY_Y+Inches(4.18),Inches(6.60),Inches(1.46),"Fig 3K — SEM: platelet morphology on PLA, PS2.0, PS2.5, PA2.0, PA2.5 (scale 10 µm)")
rrect(sl,bx2,BODY_Y+Inches(4.18),Inches(5.96),Inches(1.46),TEAL_L,TEAL,0.5)
text(sl,bx2+Inches(0.16),BODY_Y+Inches(4.34),Inches(5.60),Inches(1.10),
     "PA2.5 reduces thrombus coverage from 99.0% → 3.1%. "
     "SEM confirms rounded (inactive) platelets on PA2.5 vs spread/activated platelets on PLA. "
     "Mass increase reflects residual saline — actual thrombus formation is near zero.",
     size=10,color=TEAL_D)

# ── 23 DIV: ANTIBACTERIAL ────────────────────────────────────
sl_div(prs,23,"Section 06","Antibacterial Performance",
       "MIC · contact sterilization · live/dead staining · SEM morphology")

# ── 24 MIC TABLE ─────────────────────────────────────────────
sl=sl_new(prs,"06 · Antibacterial","Minimum Inhibitory Concentration (MIC) — Full Table",24)
text(sl,BODY_X,BODY_Y+Inches(0.10),BODY_W,Inches(0.30),
     "S. aureus growth inhibition (√ = inhibited, × = not inhibited) at concentrations 1/2 M through 1/64 M",
     size=10,color=TEXT2)
# Two sub-tables: PS/PQ left, PA/PQ right
tw=(BODY_W-Inches(0.30))/2
concs=["1/2 M","1/4 M","1/8 M","1/16 M","1/32 M","1/64 M"]
# PS/PQ ratios
ps_ratios=["2/0","2/0.2","2/0.5","2/1","2/2","2/3"]
ps_data=[[False]*6,[False]*6,[False]*6,[False]*6,[False]*6,
         [True,True,True,False,False,False]]
# PA/PQ ratios
pa_ratios=["2/0","2/0.2","2/0.5","2/1","2/2","2/3"]
pa_data=[[False]*6,[False]*6,
         [True,True,False,False,False,False],
         [True,True,True,True,False,False],
         [True,True,True,True,True,False],
         [True,True,True,True,True,False]]
col_w_t=Inches(0.60)
row_h_t=Inches(0.44)
for side,(name,ratios,data,c) in enumerate([("PS/PQ",ps_ratios,ps_data,SLATE),
                                              ("PA/PQ",pa_ratios,pa_data,TEAL)]):
    ox=BODY_X+side*(tw+Inches(0.30))
    rect(sl,ox,BODY_Y+Inches(0.52),tw,Inches(0.34),c)
    text(sl,ox+Inches(0.10),BODY_Y+Inches(0.52),tw-Inches(0.20),Inches(0.34),
         f"{name} — S. aureus MIC",size=10,bold=True,color=WHITE,anchor=MSO_ANCHOR.MIDDLE)
    # Header row
    rw_x=ox+Inches(1.50)
    for j,cc in enumerate(concs):
        text(sl,rw_x+j*col_w_t,BODY_Y+Inches(0.92),col_w_t,Inches(0.28),
             cc,size=7.5,bold=True,color=MUTED,align=PP_ALIGN.CENTER)
    for ri,(ratio,row) in enumerate(zip(ratios,data)):
        ry=BODY_Y+Inches(1.24)+ri*row_h_t
        bg=WHITE if ri%2==0 else GRAY_L
        rect(sl,ox,ry,tw,row_h_t,bg)
        text(sl,ox+Inches(0.10),ry,Inches(1.30),row_h_t,
             f"M {ratio}",size=9,bold=True,color=TEXT2,anchor=MSO_ANCHOR.MIDDLE)
        for j,val in enumerate(row):
            sym="√" if val else "×"
            fc=GREEN if val else GRAY_M
            bg2=GREEN_L if val else WHITE
            rrect(sl,rw_x+j*col_w_t+Inches(0.04),ry+Inches(0.06),
                  col_w_t-Inches(0.08),row_h_t-Inches(0.12),bg2,None)
            text(sl,rw_x+j*col_w_t,ry,col_w_t,row_h_t,sym,size=11,bold=True,
                 color=fc,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
# Summary note
rrect(sl,BODY_X,BODY_Y+Inches(4.04),BODY_W,Inches(1.60),AMBER_L,AMBER,0.5)
text(sl,BODY_X+Inches(0.20),BODY_Y+Inches(4.20),BODY_W-Inches(0.40),Inches(1.28),
     "PA/PQ: effective at PA/PQ = 2/0.5, 2/1, 2/2, 2/3 mg·mL⁻¹ (MIC = 1/4 M, 1/8 M, 1/16 M, 1/32 M). "
     "PS/PQ: only 2/3 concentration shows antibacterial effect — the strong PS-PQ binding severely shields PQ. "
     "Weak interaction in PA-PQ preserves the antibacterial potency of PQ at much lower PQ concentrations.",
     size=10,color=AMBER)

# ── 25 CONTACT STERILIZATION ─────────────────────────────────
sl=sl_new(prs,"06 · Antibacterial","Contact Sterilization: S. aureus on All 8 Coating Variants",25)
fig(sl,BODY_X,BODY_Y+Inches(0.10),Inches(7.0),Inches(2.80),"Fig 4B — Colony plates (PLA, PS2.0, PS2.5, PS5.0, PS5.5, PA2.0, PA2.5, PA5.0, PA5.5)")
fig(sl,BODY_X,BODY_Y+Inches(3.04),Inches(7.0),Inches(2.60),"Fig 4C — Bacterial adhesion ratio (%) all samples vs PLA=100%")
bx2=BODY_X+Inches(7.30)
rw5=W-bx2-PAD
text(sl,bx2,BODY_Y+Inches(0.10),rw5,Inches(0.26),"Surface Bactericidal Results",size=12,bold=True,color=TEXT)
rule(sl,bx2,BODY_Y+Inches(0.42),rw5)
ab_res=[("PLA","6.87 × 10³ CFU/cm²","Baseline — high colonization",RED),
        ("PA2.0","~0 CFU/cm²","Near-complete kill (PQ top)",GREEN),
        ("PA2.5","~0 CFU/cm²","Near-complete kill (PA top) ★",GREEN),
        ("PA5.0","97.66% inhibition","More layers, PQ top",GREEN),
        ("PA5.5","97.62% inhibition","More layers, PA top",GREEN),
        ("PS2.0","Good reduction","PQ exposed on top surface",TEAL),
        ("PS2.5","~100% adhesion","PQ shielded — no effect",RED),
        ("PS5.0","Good reduction","PQ exposed on top surface",TEAL),
        ("PS5.5","~100% adhesion","PQ shielded — no effect",RED)]
cy=BODY_Y+Inches(0.56)
for name,val,note,c in ab_res:
    if cy+Inches(0.54)>BODY_Y+BODY_H: break
    rrect(sl,bx2,cy,rw5,Inches(0.50),CARD,BORDER,0.3)
    if c==GREEN: rect(sl,bx2,cy,Inches(0.05),Inches(0.50),GREEN)
    elif c==RED: rect(sl,bx2,cy,Inches(0.05),Inches(0.50),RED)
    text(sl,bx2+Inches(0.14),cy+Inches(0.10),Inches(0.80),Inches(0.30),name,size=9,bold=True,color=TEXT)
    text(sl,bx2+Inches(1.00),cy+Inches(0.10),Inches(2.30),Inches(0.30),val,size=9.5,bold=True,color=c)
    text(sl,bx2+Inches(3.40),cy+Inches(0.10),rw5-Inches(3.50),Inches(0.30),note,size=8.5,color=MUTED)
    cy+=Inches(0.56)

# ── 26 LIVE/DEAD + BACTERIA SEM ──────────────────────────────
sl=sl_new(prs,"06 · Antibacterial","SYTO9/PI Live-Dead Staining and Bacterial SEM Morphology",26)
fig(sl,BODY_X,BODY_Y+Inches(0.10),(BODY_W-Inches(0.20))/2,Inches(3.20),
    "Fig 4D,E — SYTO9 (live) and PI (dead) fluorescence + adhesion ratio for PLA, PS2.0, PS2.5, PA2.0, PA2.5")
fig(sl,BODY_X+(BODY_W+Inches(0.20))/2,BODY_Y+Inches(0.10),(BODY_W-Inches(0.20))/2,Inches(3.20),
    "Fig 4F — SEM morphology of S. aureus on PLA, PS2.0, PS2.5, PA2.0, PA2.5 (scale 1 µm)")
rule(sl,BODY_X,BODY_Y+Inches(3.44),BODY_W)
text(sl,BODY_X,BODY_Y+Inches(3.58),BODY_W,Inches(0.26),"Interpretation",size=11,bold=True,color=TEXT)
items_ld=[("Live (SYTO9)","PA2.5 surface shows minimal green signal — very few viable bacteria remain on surface after 6 h contact."),
          ("Dead (PI)","Strong PI signal on PS2.0 and PA2.5 confirms membrane disruption by accessible PQ groups."),
          ("SEM morphology","Bacteria on PA2.5 show irregular shapes and cell wall rupture — characteristic of membrane-disruption killing by PQ."),
          ("PA vs PS","All PA variants effectively kill bacteria. PS2.5 and PS5.5 show intact bacteria — PQ shielded, no killing.")]
cy=BODY_Y+Inches(3.92)
pw6=(BODY_W-Inches(0.30))/2
for i,(kw,body) in enumerate(items_ld):
    col=i%2; row=i//2
    bx=BODY_X+col*(pw6+Inches(0.30))
    by4=cy+row*Inches(0.68)
    text(sl,bx,by4,Inches(1.60),Inches(0.26),kw,size=10,bold=True,color=TEAL)
    text(sl,bx+Inches(1.68),by4,pw6-Inches(1.68),Inches(0.54),body,size=9.5,color=TEXT2)

# ── 27 CONFOCAL CLSM ─────────────────────────────────────────
sl=sl_new(prs,"06 · Antibacterial","Confocal Laser Scanning Microscopy — Antimicrobial Adhesion",27)
fig(sl,BODY_X,BODY_Y+Inches(0.10),Inches(8.40),Inches(5.54),
    "Fig 4D — CLSM images: PLA, PS2.0, PS2.5, PA2.0, PA2.5 · rows: SYTO9, PI, Merged · scale bar 100 µm")
bx2=BODY_X+Inches(8.70); rw6=W-bx2-PAD
text(sl,bx2,BODY_Y+Inches(0.14),rw6,Inches(0.28),"CLSM Summary",size=12,bold=True,color=TEXT)
rule(sl,bx2,BODY_Y+Inches(0.48),rw6)
clsm=[("PLA",RED,"High SYTO9 (green): abundant live bacteria; minimal PI (red) signal"),
      ("PS2.0",TEAL,"Reduced SYTO9; moderate PI signal — PQ kills some bacteria from top layer"),
      ("PS2.5",RED,"High SYTO9 near PLA level — PQ shielded; no effective killing"),
      ("PA2.0",GREEN,"Very low SYTO9; strong PI — PQ accessible, effective membrane disruption"),
      ("PA2.5",GREEN,"Near-zero SYTO9; strong PI — best antibacterial response in PA series")]
cy=BODY_Y+Inches(0.60)
for name,c,note in clsm:
    rrect(sl,bx2,cy,rw6,Inches(0.90),CARD,BORDER,0.4)
    rect(sl,bx2,cy,Inches(0.05),Inches(0.90),c)
    text(sl,bx2+Inches(0.16),cy+Inches(0.12),rw6-Inches(0.26),Inches(0.26),name,size=10,bold=True,color=TEXT)
    text(sl,bx2+Inches(0.16),cy+Inches(0.44),rw6-Inches(0.26),Inches(0.40),note,size=9,color=TEXT2)
    cy+=Inches(1.02)

# ── 28 HERO SLIDE ────────────────────────────────────────────
sl=prs.slides.add_slide(prs.slide_layouts[6])
rect(sl,0,0,W,H,WHITE)
rect(sl,0,0,Inches(0.10),H,TEAL)
ftr(sl,28)
rule(sl,Inches(0.24),Inches(0.90),W-Inches(0.44),TEAL)
text(sl,Inches(0.24),Inches(0.38),Inches(9),Inches(0.42),
     "PA2.5 — OPTIMIZED COATING  ·  KEY RESULTS",size=10,bold=True,color=TEAL)
# 6 hero numbers — no boxes, just typography in a grid
heroes=[("3.1%","Thrombus coverage","(vs PLA 99.0%)",GREEN),
        (">97.6%","Bactericidal efficiency","S. aureus, solution",TEAL),
        ("21.58°","Water contact angle","highly hydrophilic",TEAL_D),
        ("~0 CFU/cm²","Surface CFU","contact sterilization",GREEN),
        (">99%","In vivo kill","rat & rabbit models",TEAL),
        ("<2%","Hemolysis rate","non-hemolytic",GREEN)]
sw3=(W-Inches(0.44)-Inches(0.20))/3
for i,(val,lbl,sub,c) in enumerate(heroes):
    col=i%3; row=i//3
    bx=Inches(0.24)+col*(sw3+Inches(0.10))
    by2=Inches(1.10)+row*Inches(2.96)
    rule(sl,bx,by2,sw3,c)
    text(sl,bx,by2+Inches(0.14),sw3,Inches(0.96),val,size=44,bold=True,color=c)
    text(sl,bx,by2+Inches(1.12),sw3,Inches(0.36),lbl,size=13,bold=True,color=TEXT)
    text(sl,bx,by2+Inches(1.50),sw3,Inches(0.26),sub,size=9.5,color=MUTED)
    rule(sl,bx,by2+Inches(1.80),sw3,GRAY_L)

# ── 29 DIV: BIOCOMPATIBILITY ─────────────────────────────────
sl_div(prs,29,"Section 07","Biocompatibility & In Vivo",
       "Hemolysis · cytotoxicity · rat model · rabbit jugular vein · PGLA suture")

# ── 30 HEMOLYSIS + CYTOTOXICITY ──────────────────────────────
sl=sl_new(prs,"07 · Biocompatibility","Hemolysis and Cytotoxicity (L929 Fibroblasts)",30)
fig(sl,BODY_X,BODY_Y+Inches(0.10),(BODY_W-Inches(0.20))/2,Inches(2.60),"Fig 5A,B — Hemolysis visual + bar chart (%) for all samples vs PC/NC")
fig(sl,BODY_X+(BODY_W+Inches(0.20))/2,BODY_Y+Inches(0.10),(BODY_W-Inches(0.20))/2,Inches(2.60),
    "Fig 5C,D — L929 viability (%) at 1 day and 3 day + FDA staining images")
rule(sl,BODY_X,BODY_Y+Inches(2.84),BODY_W)
# Hemolysis values
text(sl,BODY_X,BODY_Y+Inches(2.98),BODY_W,Inches(0.26),"Hemolysis Rate (all samples <2% — non-hemolytic threshold <5%)",size=11,bold=True,color=TEXT)
hemo_d=[("PLA","~1.5%",MUTED),("PS2.0","~1.5%",MUTED),("PS2.5","~1.0%",TEAL),
        ("PA2.0","~0.5%",GREEN),("PA2.5","~0.5%",GREEN)]
bw3=(BODY_W-Inches(0.40))/5
for i,(nm,val,c) in enumerate(hemo_d):
    bx=BODY_X+i*(bw3+Inches(0.10))
    rrect(sl,bx,BODY_Y+Inches(3.30),bw3,Inches(0.86),CARD,BORDER,0.5)
    text(sl,bx+Inches(0.10),BODY_Y+Inches(3.42),bw3-Inches(0.20),Inches(0.36),val,
         size=20,bold=True,color=c,align=PP_ALIGN.CENTER)
    text(sl,bx+Inches(0.10),BODY_Y+Inches(3.82),bw3-Inches(0.20),Inches(0.24),nm,
         size=9,color=TEXT2,align=PP_ALIGN.CENTER)
# Cytotox note
rule(sl,BODY_X,BODY_Y+Inches(4.30),BODY_W)
text(sl,BODY_X,BODY_Y+Inches(4.44),BODY_W,Inches(0.26),"L929 Cell Viability",size=11,bold=True,color=TEXT)
text(sl,BODY_X,BODY_Y+Inches(4.76),BODY_W,Inches(0.80),
     "All groups maintain >80% L929 viability at both 1 day and 3 days — "
     "confirming non-cytotoxicity of all polymer composite coatings. "
     "PA2.0 hemolysis rate (0.5%) is the lowest, suggesting that PS and PA shield the hemolytic potential of PQ.",
     size=10,color=TEXT2)

# ── 31 DIV: IN VIVO ──────────────────────────────────────────
sl_div(prs,31,"Section 07","In Vivo Validation",
       "Three animal models: rat subcutaneous · rabbit jugular vein · rat PGLA suture")

# ── 32 IN VIVO OVERVIEW ───────────────────────────────────────
sl=sl_new(prs,"07 · In Vivo","Overview of Three In Vivo Models",32)
models=[("Rat Subcutaneous\nImplant",TEAL,
         "Substrate: PLA sheet\nGroups: PLA, PS2.5, PA2.5\nInfection: S. aureus (10⁸ CFU/mL)\nRead-out: day 1 and day 3",
         "Assess antibacterial efficacy and tissue biocompatibility"),
        ("Rabbit Jugular\nVein Catheter",SLATE,
         "Substrate: PU catheter\nGroups: PU, PS2.5, PA2.5\nInfection: S. aureus + blood contact\nRead-out: day 1 and day 3",
         "Assess simultaneous antibacterial AND antithrombotic performance"),
        ("Rat PGLA Suture\n(Thigh Muscle)",GREEN,
         "Substrate: absorbable PGLA suture\nGroups: PGLA, PS2.5, PA2.5\nInfection: S. aureus (10⁸ CFU/mL)\nRead-out: day 1 and day 3",
         "Demonstrate applicability to absorbable surgical materials")]
cw4=(BODY_W-Inches(0.40))/3
for i,(name,c,design,aim) in enumerate(models):
    bx=BODY_X+i*(cw4+Inches(0.20))
    rrect(sl,bx,BODY_Y+Inches(0.10),cw4,Inches(5.54),CARD,BORDER,0.5)
    rect(sl,bx,BODY_Y+Inches(0.10),cw4,Inches(0.06),c)
    text(sl,bx+Inches(0.18),BODY_Y+Inches(0.24),cw4-Inches(0.36),Inches(0.50),
         name,size=13,bold=True,color=TEXT)
    rule(sl,bx+Inches(0.18),BODY_Y+Inches(0.82),cw4-Inches(0.36),c)
    text(sl,bx+Inches(0.18),BODY_Y+Inches(0.98),cw4-Inches(0.36),Inches(1.80),
         design,size=10,color=TEXT2)
    rule(sl,bx+Inches(0.18),BODY_Y+Inches(2.90),cw4-Inches(0.36))
    text(sl,bx+Inches(0.18),BODY_Y+Inches(3.06),cw4-Inches(0.36),Inches(0.26),
         "Study aim",size=9,bold=True,color=c)
    text(sl,bx+Inches(0.18),BODY_Y+Inches(3.36),cw4-Inches(0.36),Inches(1.50),
         aim,size=9.5,color=TEXT2)
    fig(sl,bx+Inches(0.18),BODY_Y+Inches(4.96),cw4-Inches(0.36),Inches(0.54),
        f"Schematic — Fig 6A / 7A / 8A")

# ── 33 RAT SUBCUT: CFU ───────────────────────────────────────
sl=sl_new(prs,"07 · In Vivo","Rat Subcutaneous Model — Bacterial Kill at Day 1 and Day 3",33)
fig(sl,BODY_X,BODY_Y+Inches(0.10),Inches(8.40),Inches(2.80),"Fig 6B,C — Visual images and colony plates: S. aureus at day 1 and day 3")
fig(sl,BODY_X,BODY_Y+Inches(3.04),Inches(8.40),Inches(2.60),"Fig 6D — CFU quantification (log scale) at day 1 and day 3: PLA, PS2.5, PA2.5")
bx2=BODY_X+Inches(8.70); rw7=W-bx2-PAD
text(sl,bx2,BODY_Y+Inches(0.14),rw7,Inches(0.28),"CFU Results",size=12,bold=True,color=TEXT)
rule(sl,bx2,BODY_Y+Inches(0.48),rw7)
cfu_r=[("PLA","~10⁴ CFU","High bacterial survival",RED),
       ("PS2.5","~10³–10⁴","Partial reduction only",AMBER),
       ("PA2.5","~10¹–10²","≥99% eradication",GREEN),
       ("Day 3 PA2.5","< Day 1","Effect enhanced over time",GREEN)]
cy=BODY_Y+Inches(0.60)
for name,val,note,c in cfu_r:
    rrect(sl,bx2,cy,rw7,Inches(1.26),CARD,BORDER,0.4)
    rect(sl,bx2,cy,Inches(0.05),Inches(1.26),c)
    text(sl,bx2+Inches(0.16),cy+Inches(0.12),rw7-Inches(0.26),Inches(0.26),name,size=10,bold=True,color=TEXT)
    text(sl,bx2+Inches(0.16),cy+Inches(0.44),rw7-Inches(0.26),Inches(0.30),val,size=14,bold=True,color=c)
    text(sl,bx2+Inches(0.16),cy+Inches(0.80),rw7-Inches(0.26),Inches(0.36),note,size=9,color=TEXT2)
    cy+=Inches(1.38)

# ── 34 RAT SUBCUT: HISTOLOGY ─────────────────────────────────
sl=sl_new(prs,"07 · In Vivo","Rat Subcutaneous — H&E and Immunofluorescence (CD3/CD68)",34)
fig(sl,BODY_X,BODY_Y+Inches(0.10),Inches(8.40),Inches(2.76),"Fig 6E,F — SEM of implant surfaces + H&E staining at day 1 and day 3 (scale 500 µm)")
fig(sl,BODY_X,BODY_Y+Inches(2.98),Inches(8.40),Inches(2.66),"Fig 6G — Immunofluorescence DAPI/CD3 and DAPI/CD68 at day 1 and day 3")
bx2=BODY_X+Inches(8.70); rw8=W-bx2-PAD
text(sl,bx2,BODY_Y+Inches(0.14),rw8,Inches(0.28),"Immune Response",size=12,bold=True,color=TEXT)
rule(sl,bx2,BODY_Y+Inches(0.48),rw8)
imm=[("H&E",TEAL,"PLA/PS2.5: significant endometrial hyperplasia, abundant inflammatory cells. PA2.5: mild tissue reaction."),
     ("CD3 (T-cells)",TEAL,"Area ratio: PLA ~4%, PS2.5 ~4%, PA2.5 ~1% at day 1; similar at day 3. PA2.5 significantly reduced."),
     ("CD68 (macrophages)",GREEN,"Macrophage infiltration significantly reduced around PA2.5 implant — low chronic inflammation."),
     ("Conclusion",GREEN,"PA2.5 elicits a milder, more favorable host immune response — superior biocompatibility in vivo.")]
cy=BODY_Y+Inches(0.60)
for kw,c,body in imm:
    rrect(sl,bx2,cy,rw8,Inches(1.14),CARD,BORDER,0.4)
    rect(sl,bx2,cy,Inches(0.05),Inches(1.14),c)
    text(sl,bx2+Inches(0.16),cy+Inches(0.10),rw8-Inches(0.26),Inches(0.26),kw,size=10,bold=True,color=TEXT)
    text(sl,bx2+Inches(0.16),cy+Inches(0.42),rw8-Inches(0.26),Inches(0.62),body,size=9,color=TEXT2)
    cy+=Inches(1.24)

# ── 35 RABBIT JUGULAR: CFU + THROMBUS ────────────────────────
sl=sl_new(prs,"07 · In Vivo","Rabbit Jugular Vein — Bacterial Kill and Antithrombotic",35)
fig(sl,BODY_X,BODY_Y+Inches(0.10),Inches(8.40),Inches(2.76),"Fig 7B,C — Colony plates + macroscopic view of catheters at day 1 and day 3")
fig(sl,BODY_X,BODY_Y+Inches(2.98),Inches(8.40),Inches(2.66),"Fig 7D,E — SEM of catheter surfaces + H&E cross-sections of jugular vein")
bx2=BODY_X+Inches(8.70); rw9=W-bx2-PAD
text(sl,bx2,BODY_Y+Inches(0.14),rw9,Inches(0.28),"Rabbit Model Results",size=12,bold=True,color=TEXT)
text(sl,bx2,BODY_Y+Inches(0.40),rw9,Inches(0.22),"Substrate: PU catheter  ·  S. aureus infection",size=8.5,color=MUTED,italic=True)
rule(sl,bx2,BODY_Y+Inches(0.68),rw9)
rabb=[("PU control",RED,"High CFU at day 1 (~10⁵) and day 3. Macroscopic thrombus visible. Severe inflammation."),
      ("PS2.5",AMBER,"CFU reduced vs PU but significant thrombus still present. Antibacterial partial; thrombosis not prevented."),
      ("PA2.5",GREEN,"CFU >99% reduction. Minimal thrombus visible. Confirms dual antibacterial + antithrombotic in vivo."),
      ("Day 3 PA2.5",GREEN,"Antibacterial effect further enhanced at day 3. Lumen largely patent. Superior to all controls.")]
cy=BODY_Y+Inches(0.82)
for name,c,body in rabb:
    rrect(sl,bx2,cy,rw9,Inches(1.10),CARD,BORDER,0.4)
    rect(sl,bx2,cy,Inches(0.05),Inches(1.10),c)
    text(sl,bx2+Inches(0.16),cy+Inches(0.10),rw9-Inches(0.26),Inches(0.24),name,size=10,bold=True,color=TEXT)
    text(sl,bx2+Inches(0.16),cy+Inches(0.40),rw9-Inches(0.26),Inches(0.58),body,size=9,color=TEXT2)
    cy+=Inches(1.18)

# ── 36 RABBIT JUGULAR: HISTOLOGY ─────────────────────────────
sl=sl_new(prs,"07 · In Vivo","Rabbit Jugular Vein — Immunofluorescence (CD31 / CD68)",36)
fig(sl,BODY_X,BODY_Y+Inches(0.10),Inches(8.40),Inches(5.54),
    "Fig 7F — DAPI/CD31 and DAPI/CD68 immunofluorescence: PU, PS2.5, PA2.5 at day 1 and day 3 (scale 500 µm)")
bx2=BODY_X+Inches(8.70); rw10=W-bx2-PAD
text(sl,bx2,BODY_Y+Inches(0.14),rw10,Inches(0.28),"Histology Results",size=12,bold=True,color=TEXT)
rule(sl,bx2,BODY_Y+Inches(0.48),rw10)
hist=[("CD31 (endothelial)",SLATE,"Strong in PU/PS2.5 → neovascularization from repair response. Weak in PA2.5 → less tissue damage, better biocompat."),
      ("CD68 (macrophages)",RED,"Abundant macrophages in PU and PS2.5 groups. Significantly reduced around PA2.5 catheter."),
      ("Quantification",GREEN,"CD31 and CD68 area ratio both significantly lower in PA2.5 at day 1 and day 3 (p<0.01)."),
      ("Interpretation",TEAL,"Reduced inflammatory infiltrate and neovascularization confirms PA2.5 reduces infection-driven tissue damage.")]
cy=BODY_Y+Inches(0.60)
for kw,c,body in hist:
    rrect(sl,bx2,cy,rw10,Inches(1.14),CARD,BORDER,0.4)
    rect(sl,bx2,cy,Inches(0.05),Inches(1.14),c)
    text(sl,bx2+Inches(0.16),cy+Inches(0.10),rw10-Inches(0.26),Inches(0.26),kw,size=10,bold=True,color=TEXT)
    text(sl,bx2+Inches(0.16),cy+Inches(0.42),rw10-Inches(0.26),Inches(0.62),body,size=9,color=TEXT2)
    cy+=Inches(1.24)

# ── 37 PGLA SUTURE: CFU ──────────────────────────────────────
sl=sl_new(prs,"07 · In Vivo","PGLA Suture Model (Rat Thigh Muscle) — Bacterial Kill",37)
fig(sl,BODY_X,BODY_Y+Inches(0.10),Inches(7.60),Inches(2.60),"Fig 8A,B — Rat thigh suture model schematic + colony plates at day 1 and day 3")
fig(sl,BODY_X,BODY_Y+Inches(2.84),Inches(7.60),Inches(2.80),"Fig 8C — CFU quantification (log scale) PGLA, PS2.5, PA2.5 at day 1 and day 3")
bx2=BODY_X+Inches(7.90); rw11=W-bx2-PAD
text(sl,bx2,BODY_Y+Inches(0.14),rw11,Inches(0.28),"PGLA Suture Results",size=12,bold=True,color=TEXT)
text(sl,bx2,BODY_Y+Inches(0.40),rw11,Inches(0.22),"Biodegradable absorbable suture model",size=8.5,color=MUTED,italic=True)
rule(sl,bx2,BODY_Y+Inches(0.68),rw11)
pgla=[("PGLA control","~10⁶ CFU/cm²","Abundant bacteria at day 1 and day 3",RED),
      ("PS2.5","~10⁴–10⁵","Partial reduction; CD3/CD68 still high",AMBER),
      ("PA2.5 day 1","~10² CFU","Bacterial killing >99% ★",GREEN),
      ("PA2.5 day 3","Further ↓","Killing enhanced; inflammation reduced",GREEN)]
cy=BODY_Y+Inches(0.82)
for name,val,note,c in pgla:
    rrect(sl,bx2,cy,rw11,Inches(1.10),CARD,BORDER,0.4)
    rect(sl,bx2,cy,Inches(0.05),Inches(1.10),c)
    text(sl,bx2+Inches(0.16),cy+Inches(0.10),rw11-Inches(0.26),Inches(0.24),name,size=10,bold=True,color=TEXT)
    text(sl,bx2+Inches(0.16),cy+Inches(0.40),Inches(1.80),Inches(0.28),val,size=12,bold=True,color=c)
    text(sl,bx2+Inches(2.06),cy+Inches(0.40),rw11-Inches(2.16),Inches(0.28),note,size=9,color=TEXT2)
    cy+=Inches(1.18)

# ── 38 PGLA SUTURE: HISTOLOGY ────────────────────────────────
sl=sl_new(prs,"07 · In Vivo","PGLA Suture Model — H&E and Immunofluorescence",38)
fig(sl,BODY_X,BODY_Y+Inches(0.10),(BODY_W-Inches(0.20))/2,Inches(2.80),"Fig 8D — H&E staining of peri-suture tissue (scale 200 µm) day 1 and day 3")
fig(sl,BODY_X+(BODY_W+Inches(0.20))/2,BODY_Y+Inches(0.10),(BODY_W-Inches(0.20))/2,Inches(2.80),
    "Fig 8E — Immunofluorescence DAPI/CD3 and DAPI/CD68 (scale 500 µm) day 1 and day 3")
rule(sl,BODY_X,BODY_Y+Inches(3.04),BODY_W)
text(sl,BODY_X,BODY_Y+Inches(3.18),BODY_W,Inches(0.26),"Histology Key Findings",size=11,bold=True,color=TEXT)
pgla_h=[("H&E — PLA/PS2.5","Significant endometrial hyperplasia and inflammatory cell infiltration at day 1 and day 3.","Acute infection-driven inflammation",RED),
        ("H&E — PA2.5","Markedly reduced inflammatory cell infiltration. Mild tissue reaction consistent with foreign body response.",  "Superior biocompatibility",GREEN),
        ("CD3/CD68 — PA2.5","CD3 and CD68 area ratios significantly lower than PGLA and PS2.5 at both time points (Fig 8F, p<0.001).", "Milder immune activation",GREEN),
        ("Overall","PA2.5 coating on absorbable suture material maintains antibacterial activity and low inflammatory response — demonstrates versatility for clinical suture applications.","Translation potential",TEAL)]
pw7=(BODY_W-Inches(0.30))/2
cy=BODY_Y+Inches(3.52)
for i,(kw,body,tag_l,c) in enumerate(pgla_h):
    col=i%2; row=i//2
    bx=BODY_X+col*(pw7+Inches(0.30))
    by4=cy+row*Inches(0.94)
    rrect(sl,bx,by4,pw7,Inches(0.86),CARD,BORDER,0.4)
    rect(sl,bx,by4,Inches(0.05),Inches(0.86),c)
    text(sl,bx+Inches(0.16),by4+Inches(0.08),pw7-Inches(0.26),Inches(0.24),kw,size=9.5,bold=True,color=TEXT)
    text(sl,bx+Inches(0.16),by4+Inches(0.38),pw7-Inches(0.26),Inches(0.40),body,size=9,color=TEXT2)

# ── 39 DIV: CONCLUSIONS ──────────────────────────────────────
sl_div(prs,39,"Section 08","Conclusions",
       "Mechanism · comparative performance · future outlook")

# ── 40 CONCLUSION + IMPACT ────────────────────────────────────
sl=prs.slides.add_slide(prs.slide_layouts[6])
rect(sl,0,0,W,H,WHITE)
rect(sl,0,0,Inches(0.10),H,TEAL)
ftr(sl,40)
text(sl,Inches(0.24),Inches(0.22),Inches(9),Inches(0.30),
     "CONCLUSIONS & OUTLOOK",size=9,bold=True,color=TEAL)
text(sl,Inches(0.24),Inches(0.56),W-Inches(0.44),Inches(0.48),
     "Weak Co-assembly as a Strategy for Dual-Function Medical Device Coatings",
     size=22,bold=True,color=TEXT)
rule(sl,Inches(0.24),Inches(1.12),W-Inches(0.44))
# Mechanism panel (left 55%)
mech_w=Inches(7.0)
text(sl,Inches(0.24),Inches(1.22),mech_w,Inches(0.26),"Mechanism",size=11,bold=True,color=TEXT)
mech_steps=[("Co-assembly","PA and PQ spontaneously co-assemble (ΔG < 0) via weak carboxylate–ammonium interaction.",TEAL),
            ("Weak interaction","Looser complex leaves PQ functional groups mobile and surface-accessible.",TEAL),
            ("Antifouling","Zwitterionic PA groups form a dense hydration shell: WCA 21.58°, thrombus 3.1%.",GREEN),
            ("Antibacterial","Accessible PQ disrupts bacterial membranes: >97.6% kill, ~0 CFU on surface.",GREEN),
            ("In vivo","PA2.5 achieves >99% bacterial eradication in three distinct animal models.",TEAL_D),
            ("Safety","Hemolysis <2%, L929 viability >80%, mild immune response — excellent biocompatibility.",GREEN)]
cy=Inches(1.56)
for kw,body,c in mech_steps:
    text(sl,Inches(0.24),cy,Inches(1.40),Inches(0.22),kw,size=9.5,bold=True,color=c)
    text(sl,Inches(1.70),cy,mech_w-Inches(1.50),Inches(0.22),body,size=9.5,color=TEXT2)
    cy+=Inches(0.76)
rule(sl,Inches(0.24),cy+Inches(0.06),mech_w)
# Future outlook
text(sl,Inches(0.24),cy+Inches(0.22),mech_w,Inches(0.24),"Future Directions",size=10,bold=True,color=TEXT)
fut=["Long-term stability and in vivo durability studies",
     "Clinical translation to central venous catheters and orthopedic implants",
     "Expanded pathogen spectrum: Gram-negative bacteria, fungi",
     "Tuning ionic strength to optimize interaction strength for different applications"]
cy2=cy+Inches(0.52)
for f in fut:
    text(sl,Inches(0.24),cy2,Inches(0.22),Inches(0.22),"→",size=9.5,bold=True,color=TEAL)
    text(sl,Inches(0.48),cy2,mech_w-Inches(0.28),Inches(0.22),f,size=9.5,color=TEXT2); cy2+=Inches(0.38)
# Right: comparison table
bx_r=Inches(0.24)+mech_w+Inches(0.30)
tw2=W-bx_r-PAD
rect(sl,bx_r,Inches(1.22),tw2,Inches(5.60),CARD)
rect(sl,bx_r,Inches(1.22),tw2,Inches(0.04),TEAL)
text(sl,bx_r+Inches(0.16),Inches(1.30),tw2-Inches(0.32),Inches(0.28),"PA2.5 vs Controls",size=11,bold=True,color=TEXT)
t_hdrs=["","PLA","PS2.5","PA2.5 ★"]
t_cws=[Inches(1.90),Inches(1.10),Inches(1.10),Inches(1.10)]
t_cols=[bx_r+Inches(0.10),bx_r+Inches(2.10),bx_r+Inches(3.26),bx_r+Inches(4.42)]
t_rows=[["Thrombus","99.0%","89.7%","3.1%"],
        ["Mass incr.","152.9%","82.2%","51.7%"],
        ["WCA (°)","88.73","~65","21.58"],
        ["Bact. kill","—","partial",">97.6%"],
        ["CFU/cm²","6870","high","~0"],
        ["Hemolysis","~1.5%","~1.0%","~0.5%"],
        ["In vivo kill","—","partial",">99%"],
        ["CD3/CD68","high","high","low"]]
for j,(h,cw) in enumerate(zip(t_hdrs,t_cws)):
    text(sl,t_cols[j],Inches(1.64),cw,Inches(0.24),h,size=8,bold=True,
         color=WHITE if j>0 else MUTED,align=PP_ALIGN.CENTER if j>0 else PP_ALIGN.LEFT)
rect(sl,bx_r,Inches(1.62),tw2,Inches(0.28),TEAL)
for ri,row in enumerate(t_rows):
    ry=Inches(1.96)+ri*Inches(0.42)
    rect(sl,bx_r,ry,tw2,Inches(0.40),CARD2 if ri%2==0 else WHITE)
    for j,(val,cw) in enumerate(zip(row,t_cws)):
        c=TEAL_D if j==3 else (RED if j==1 else TEXT2)
        if j==3: rect(sl,t_cols[j],ry,cw,Inches(0.40),TEAL_L)
        text(sl,t_cols[j],ry,cw,Inches(0.40),val,size=9,bold=(j==3),
             color=c,align=PP_ALIGN.CENTER if j>0 else PP_ALIGN.LEFT,anchor=MSO_ANCHOR.MIDDLE)


# ── SAVE ─────────────────────────────────────────────────────
prs.save('slides2.pptx')
print(f"slides2.pptx saved — {len(prs.slides)} slides")
