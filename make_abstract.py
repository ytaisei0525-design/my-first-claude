# -*- coding: utf-8 -*-
"""
雑誌会要旨（Word）を生成するスクリプト。
- 見本（島田氏の要旨）と同じ形式に揃える。
- 色は使わず、見出し・重要語は太字、補足は（かっこ）で表現。
- 図の位置には空の罫線ボックス（枠）を置き、その下にキャプションを付す。
- 対象論文：Co-assembly of zwitterionic and cationic polymers ...（Biomaterials 2026）
"""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

JP_FONT = "ＭＳ 明朝"      # 本文（明朝）
JP_FONT_G = "ＭＳ ゴシック"  # 見出し（ゴシック）

doc = Document()

# ---- ページ余白（A4・見本に近い体裁） ----
sec = doc.sections[0]
sec.top_margin = Cm(2.0)
sec.bottom_margin = Cm(2.0)
sec.left_margin = Cm(2.2)
sec.right_margin = Cm(2.2)

# ---- 既定スタイル ----
style = doc.styles["Normal"]
style.font.name = JP_FONT
style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn("w:eastAsia"), JP_FONT)


def set_run_font(run, font=JP_FONT, size=10.5, bold=False):
    run.font.name = font
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)  # 黒のみ（色は使わない）
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font)
    rfonts.set(qn("w:ascii"), font)
    rfonts.set(qn("w:hAnsi"), font)


def add_para(text="", size=10.5, bold=False, align=None, font=JP_FONT,
             space_after=4, space_before=0, line=1.15):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = line
    if text:
        r = p.add_run(text)
        set_run_font(r, font=font, size=size, bold=bold)
    return p


def add_body(segments, size=10.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=True):
    """segments: [(text, bold), ...] 太字を混在させた本文段落。"""
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_after = Pt(6)
    pf.line_spacing = 1.3
    if first_indent:
        pf.first_line_indent = Pt(10.5)
    for text, bold in segments:
        r = p.add_run(text)
        set_run_font(r, font=JP_FONT, size=size, bold=bold)
    return p


def add_heading(text, size=11.5):
    """【概要】などの大見出し（太字・ゴシック）。"""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(10)
    pf.space_after = Pt(5)
    r = p.add_run(text)
    set_run_font(r, font=JP_FONT_G, size=size, bold=True)
    return p


def add_subheading(text, size=10.5):
    """結果・考察内の小見出し（太字）。"""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(7)
    pf.space_after = Pt(3)
    r = p.add_run(text)
    set_run_font(r, font=JP_FONT_G, size=size, bold=True)
    return p


def _set_cell_border(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "8")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "000000")
        borders.append(e)
    tcPr.append(borders)


def add_figure_box(caption, height_cm=4.2):
    """図を貼るための空の罫線ボックスと、その下のキャプションを追加。"""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    # 幅を本文幅いっぱいに
    table.columns[0].width = Cm(16.5)
    cell = table.cell(0, 0)
    cell.width = Cm(16.5)
    _set_cell_border(cell)
    # 高さ確保
    tr = table.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(int(height_cm * 567)))  # cm -> twips
    trHeight.set(qn("w:hRule"), "atLeast")
    trPr.append(trHeight)
    # セル内に空行（図はここに貼る）
    cp = cell.paragraphs[0]
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rr = cp.add_run("（図　ここに貼付）")
    set_run_font(rr, font=JP_FONT, size=9, bold=False)
    # キャプション
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_before = Pt(2)
    cap.paragraph_format.space_after = Pt(8)
    rc = cap.add_run(caption)
    set_run_font(rc, font=JP_FONT, size=9.5, bold=True)


# ============================================================
# ヘッダー部
# ============================================================
add_para("雑誌会要旨", size=13, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
         font=JP_FONT_G, space_after=2)
add_para("2026/06/02", size=10.5, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=0)
add_para("M2　＿＿＿＿＿", size=10.5, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=8)

# 英語タイトル
ptitle = doc.add_paragraph()
ptitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
ptitle.paragraph_format.space_after = Pt(2)
rt = ptitle.add_run(
    "Co-assembly of zwitterionic and cationic polymers for antibacterial "
    "and antithrombotic surfaces via weak electrostatic interactions"
)
set_run_font(rt, font=JP_FONT, size=11.5, bold=True)
add_para("Biomaterials 331 (2026) 124114（J. Wang ら，四川大学）",
         size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)

# ============================================================
# 【概要】
# ============================================================
add_heading("【概要】")

add_body([
    ("埋め込み型の医療デバイス（中心静脈カテーテル、外科用縫合糸、気管挿管チューブなど）は、"
     "集中治療や外科手術、術後管理において広く用いられている。しかし、これらのデバイスに関連する"
     "細菌感染は依然として臨床上の大きな課題であり、デバイス関連感染の発生率は", False),
    ("20%を超える", True),
    ("と報告されている。重度の細菌感染は、デバイスの機能不全や再手術を引き起こすだけでなく、"
     "菌血症や敗血症、多臓器不全といった致命的な合併症を招き、患者の死亡率と医療費を大幅に増加させる"
     "（図1）。デバイス表面を", False),
    ("抗菌コーティング", True),
    ("で改質することは細菌の付着を抑制する有効な戦略であり、適切な抗菌剤の選択がその鍵となる。", False),
])

add_body([
    ("陽イオン性ポリマーは分子鎖上に正電荷を持つ抗菌剤の一種であり、広域スペクトルの抗菌活性を示す。"
     "これらは静電的相互作用によって細菌膜を破壊し、細菌耐性を誘導しにくいという利点がある。なかでも", False),
    ("ポリクオタニウム（PQ）", True),
    ("は代表的な陽イオン性ポリマーであり、第四級アンモニウムカチオンと疎水性の炭素鎖を有する。"
     "疎水性炭素鎖はさらに膜の完全性を破壊し、殺菌効果を相乗的に高める。", False),
])

add_body([
    ("PQをデバイス表面へ被覆する方法としては、ポリマーブラシ法、", False),
    ("層状自己組織化（LBL）法", True),
    ("、ドーパミン補助による共有結合グラフト法などが用いられてきた。なかでもLBL法は、正負に帯電した"
     "ポリ電解質を交互に吸着させる簡便かつ汎用性の高い手法である。しかし従来のLBL法で作製した"
     "コーティングでは、PQの正電荷が負に帯電した", False),
    ("陰イオン性ポリマー（PS）", True),
    ("によって強く遮蔽されるため抗菌性が制限され、また正負電荷の強い静電引力により電荷と水分子との"
     "結合が減少して、親水性・抗ファウリング性能も不十分となるという課題があった。", False),
])

add_body([
    ("一方、", False),
    ("両性イオン性ポリマー（ポリスルホベタインなど）", True),
    ("は、分子鎖内に等量の陰イオン（スルホン酸基など）と陽イオン（第四級アンモニウム基など）を"
     "併せ持ち、全体として電気的に中性である。両性イオン性ポリマーと陽イオン性ポリマーの間の"
     "静電的相互作用は、陰イオン−陽イオン系のそれよりも弱いことが報告されている。著者らは、この", False),
    ("「弱い静電的相互作用」", True),
    ("が陽イオン基により大きな「自由度」を与え、細菌膜を破壊する能力を保持させると同時に、"
     "荷電基と水分子との結合を維持して緻密な", False),
    ("「水和保護層」", True),
    ("の形成を促すと考えた。", False),
])

add_body([
    ("本研究では、両性イオン性ポリマー", False),
    ("PA", True),
    ("（poly(2-(methacryloyloxy)ethyl) dimethyl-(3-sulfopropyl) ammonium hydroxide）と"
     "陽イオン性ポリマー", False),
    ("PQ", True),
    ("をLBL法で組み合わせ、抗菌性と抗ファウリング性を兼ね備えた複合コーティングを構築した。"
     "溶液中での両ポリマーの相互作用を解析した後、コーティングの物理化学的特性、抗菌性、"
     "抗ファウリング性、生体適合性をin vitroおよびin vivoで評価し、さらに吸収性縫合糸や"
     "血液接触カテーテルへ応用してその有効性を検証した。", False),
])

add_figure_box("図1：弱い静電的相互作用を利用した両性イオン性／陽イオン性ポリマー複合コーティングの概念図",
               height_cm=4.5)

# ============================================================
# 【結果・考察】
# ============================================================
add_heading("【結果・考察】")

add_subheading("溶液中でのPQとPAの相互作用")
add_body([
    ("ポリマーはフリーラジカル重合により合成され、", False),
    ("1H NMRおよびXPS", True),
    ("によりその構造と元素組成が確認された。PAは純水には不溶であるが電解質存在下で速やかに"
     "溶解性が増す。塩溶液中での挙動を調べると、PS-PQおよびPA-PQ混合物はいずれも混合直後に白濁し、"
     "PS-PQの方が高い濁度を示した。さらに平衡到達後、PA-PQ混合液は", False),
    ("透明化", True),
    ("したがPS-PQには変化が見られなかった。550 nmにおける吸光度測定でも、PQ添加によりPSの吸光度は"
     "1.79まで上昇したのに対しPAは1.33にとどまり、PS-PQがより緻密な凝集体を形成することが示された。", False),
])
add_body([
    ("等温滴定型熱量測定（ITC）", True),
    ("により、反対電荷を持つポリマー間の複合体形成は自発的（ΔG＜0）に進行することが示された。"
     "同濃度でのPS初滴下時のエンタルピー変化はPA-PQより大きく、", False),
    ("PA-PQ間の静電的相互作用が非常に弱い", True),
    ("ことが裏付けられた。粘度測定では、PS-PQの粘度はPS単独より低下した一方、PA-PQの粘度は"
     "著しく増加し、両ポリマー間に「ゲル状」の架橋ネットワークが形成された。蛍光顕微鏡観察でも"
     "PS-PQ（平均強度34.25）がPA-PQ（同43.59）より低く、より強固な複合体を形成することが確認された。"
     "動的光散乱（DLS）では、陽イオン濃度の増加に伴いゼータ電位が負から正へ転じ、粒子径が増大した。", False),
])

add_figure_box("図2：コーティングの自己組織化過程と表面特性")

add_subheading("コーティングの自己組織化過程と表面特性")
add_body([
    ("QCM-d", True),
    ("による実時間モニタリングにより、ポリマーが基板表面へ逐次吸着し、弱く吸着した分子はPBS洗浄で"
     "除去される様子が観察された。ゼータ電位は積層数に応じて変化し、PAまたはPSが陽イオン性PQと"
     "PLA表面へ良好に組み立てられたことが確認された。XPSおよびSEM-EDSでは、両コーティングで"
     "窒素（N）と硫黄（S）の特性シグナルが検出され、コーティングの形成が裏付けられた。", False),
])
add_body([
    ("AFM", True),
    ("による粗さ評価では、純PLAは平坦（Rq＝1.06 nm）であったのに対し複合コーティングでは粗さが"
     "増加し、PS-PQは正電荷の不均一分布に由来する典型的な", False),
    ("「島状構造」", True),
    ("を形成した。一方PA-PQは積層数の増加に伴い粗さが連続的に減少した。これはPA中の正負電荷が"
     "後続のPQ層へのアンカー点を増やすためである。接触角測定では、PLA（88.73°）に対しPA2.0、PA2.5は"
     "それぞれ53.3°、21.58°と優れた親水性を示し、PAの", False),
    ("強い水和能", True),
    ("が反映された。エタノール置換試験でもPA-PQはより多くの液体を吸着し、より柔らかく疎な構造"
     "（＝より弱い静電的相互作用）を持つことが示された。", False),
])

add_figure_box("図3：コーティングの抗ファウリング性能と抗血栓性")

add_subheading("抗ファウリング（抗付着）性能")
add_body([
    ("病原体の宿主構造（タンパク質や細胞）への付着は感染症発症の重要な初期段階である。"
     "タンパク質・細胞・血小板の付着試験および全血循環試験により評価したところ、PLAやPS-PQと比較して", False),
    ("PA-PQでは各種生体分子の吸着が有意に減少", True),
    ("し、高い親水性に起因する優れた抗ファウリング性能が示された。", False),
])
add_body([
    ("抗血栓性を評価するため、ウサギの", False),
    ("動静脈シャント（ex vivo）モデル", True),
    ("を用いた。1時間の循環後、PLA、PS2.0、PS2.5では明らかな血栓沈着（表面被覆率99.0%、97.1%、89.7%）"
     "が観察された一方、PA-PQでは血栓被覆と質量増加が大幅に減少した。特に", False),
    ("PA2.5ではほとんど血栓が形成されず（被覆率わずか3.1%）", True),
    ("、優れた親水性がタンパク質吸着や血液・細胞付着を効果的に抑制したことが示された。", False),
])

add_figure_box("図4：コーティングの抗菌性能")

add_subheading("抗菌性能")
add_body([
    ("グラム陽性菌の代表である", False),
    ("S. aureus", True),
    ("をモデル菌とした。微量液体希釈法によるMIC測定では、PA/PQ混合物はPQ濃度に応じてMIC値が"
     "段階的に低下し、弱い静電的相互作用が抗菌効果に正の影響を与えることが示された。一方PS/PQでは"
     "PS/PQ-M2/3のみが抗菌効果を示し、単純な正負電荷の静電的相互作用は抗菌効果を発揮せず、"
     "むしろ正電荷の抗菌効果を遮蔽することが示された。", False),
])
add_body([
    ("接触殺菌試験では、PLA表面の生菌数が6.87×10³ CFU cm⁻²であったのに対し", False),
    ("PA2.0およびPA5.0表面では生菌が検出されず", True),
    ("、PA2.5、PA5.5も97.66%、97.62%と高い抗菌率を示した。これはPAが主に抗ファウリング性を、"
     "PQが抗菌性を担うことを示す。CLSM観察やSEM観察では、細菌が不規則な形態と", False),
    ("細胞壁の破裂", True),
    ("を示し、PQが細胞壁の完全性を破壊して抗菌効果を発揮することが確認された。弱い静電的相互作用により"
     "PA-PQ中のPQとPAは高い自由度を持ち、", False),
    ("PQによる殺菌とPAによる付着抑制という二重機構", True),
    ("を発揮する。", False),
])

add_figure_box("図5：コーティングの生体適合性（溶血・細胞毒性）")

add_subheading("生体適合性")
add_body([
    ("埋め込み型デバイスへの応用には良好な生体適合性が不可欠である。全サンプルの", False),
    ("溶血率は5%未満", True),
    ("であり、コーティングが赤血球に有意な損傷を与えないことが示された。特に2.5コーティングの溶血率は"
     "2.0コーティングより低く、PSおよびPAがPQの赤血球への潜在的損傷を遮蔽することが示唆された。"
     "L929線維芽細胞を用いた細胞毒性試験でも、全群で", False),
    ("細胞生存率が80%超", True),
    ("となり、細胞毒性がないことが確認された。", False),
])

add_figure_box("図6：ラット皮下埋め込みモデルにおけるin vivo抗菌性能と組織適合性の評価")

add_subheading("生体内（in vivo）での生体適合性と抗菌性能")
add_body([
    ("S. aureusを表面に付着させたサンプルをラットに埋め込み評価した。埋め込み1日後・3日後の生菌数測定では、"
     "PLAおよびPS2.5表面で有意な細菌生存が認められた一方、", False),
    ("PA2.5では細菌除去率が99%超", True),
    ("であった。H&E染色では、PLAとPS2.5周囲に顕著な炎症細胞浸潤（急性炎症反応）が見られたのに対し、"
     "PA2.5群は軽度の組織反応にとどまった。CD3（T細胞）およびCD68（マクロファージ）の免疫蛍光染色でも"
     "PA2.5の浸潤は有意に少なく、優れた抗菌性により感染関連の炎症が軽減された。", False),
])
add_body([
    ("さらに", False),
    ("ウサギ頸静脈埋め込みモデル", True),
    ("では、PA2.5表面で細菌除去率99%超とともに血栓形成が最小限に抑えられ、抗菌性と抗ファウリング性の"
     "両立が確認された。CD31（新生血管）・CD68の発現もPLA・PS2.5で強くPA2.5では弱く、弱い静電的"
     "相互作用によりPQの毒性が低減されつつ抗菌性が保持され、PAが抗ファウリング効果を発揮することが示された。", False),
])

add_figure_box("図7：ウサギ頸静脈埋め込みモデルにおける抗菌性能と組織反応の評価")

add_subheading("吸収性縫合糸への応用（ラット大腿筋損傷モデル）")
add_body([
    ("複合コーティングを", False),
    ("吸収性外科縫合糸（PGLA）", True),
    ("へ適用し、ラットの大腿筋創傷の治療に用いた。S. aureusを付着させた縫合糸で創を縫合したところ、"
     "1日後・3日後ともにPA2.5表面の生菌はごくわずかで、", False),
    ("殺菌率はいずれも99%超", True),
    ("、3日後にはさらに抗菌効果が増強した。H&E染色および免疫蛍光染色（CD3、CD68）でも、PA2.5周囲の"
     "炎症細胞浸潤はPGLAやPS2.5より顕著に少なく、優れた抗菌性により感染関連炎症が軽減され、"
     "良好な生体適合性が示された。", False),
])

add_figure_box("図8：ラット大腿筋損傷モデルにおける抗菌性能と組織反応の評価")

# ============================================================
# 【結論】
# ============================================================
add_heading("【結論】")
add_body([
    ("本研究では、陽イオン性ポリマーPQと両性イオン性ポリマーPAをLBL法で層状に自己組織化させることにより、", False),
    ("抗菌性と抗ファウリング性を兼ね備えた複合コーティングの構築に成功", True),
    ("した。複合コーティング中のPQは広域スペクトルの抗菌活性を示し、PAはその強い水和能により緻密な"
     "水和層を形成して、細菌・細胞・タンパク質の非特異的付着を効果的に抑制した。PAとPQの間の", False),
    ("弱い静電的相互作用", True),
    ("はITC、DLS、MICなどの実験により裏付けられ、この相互作用がPQの毒性を大幅に低減しつつ抗菌効果を"
     "概ね維持することを明らかにした。in vitroおよびin vivo実験により、本コーティングは優れた"
     "抗ファウリング性、抗菌性、生体適合性を示し、", False),
    ("埋め込み型デバイス関連感染の予防に大きな可能性", True),
    ("を持つことが示された。今後は、イオン強度がポリマーに与える影響をさらに検討するとともに、"
     "ポリマー比を調整して広域スペクトルの抗菌性能を確保することが課題である。", False),
])

# ============================================================
# 【実験方法】
# ============================================================
add_heading("【実験方法】")

add_subheading("ポリマーの合成")
add_body([
    ("第四級アンモニウム塩（QAS）モノマーは文献の方法に従って合成した。両性イオン性モノマー（SBMA）、"
     "陽イオン性モノマー（QAS-8C）、陰イオン性モノマー（SMPS）を適切な溶媒（SBMAとSMPSはトリフルオロ"
     "エタノール、QASはイソプロパノール）に溶解し、開始剤AIBME（モノマーの1質量%）を加えて65℃・"
     "無酸素雰囲気下で12時間フリーラジカル重合した。得られた溶液を濃縮・洗浄・乾燥し、ポリマー"
     "PA（poly-SBMA）、PQ（poly-QAS-8C）、PS（poly-SMPS）を得た。", False),
], first_indent=False)

add_subheading("溶液中での相互作用解析")
add_body([
    ("25℃で等温滴定型熱量測定（ITC）を行い、PS-PQおよびPA-PQの相互作用に伴う熱変化を測定した。"
     "各ポリマー溶液（2 mg mL⁻¹）の粘度はレオメーターで測定し、UVスペクトル、蛍光顕微鏡観察、"
     "ゼータ電位および粒子径（DLS）により混合後の挙動を解析した。", False),
], first_indent=False)

add_subheading("複合コーティングの作製")
add_body([
    ("コーティングはLBL法により作製した（8〜15層、最初の4層を基層とする）。PLA基板をDOPA溶液に2時間"
     "浸漬して前処理し、PEI、PS、PQ溶液へ順次浸漬して基層接着層を形成した後、PS/PA溶液とPQ溶液へ"
     "交互に浸漬した。各浸漬の間に脱イオン水で洗浄し40℃で乾燥した。陽イオン性を最上層とする8層・14層を"
     "PA2.0・PA5.0（PS2.0・PS5.0）、両性／陰イオン性を最上層とする9層・15層をPA2.5・PA5.5"
     "（PS2.5・PS5.5）と命名した。同手法をポリウレタン（PU）チューブおよび吸収性縫合糸（PGLA）にも適用した。", False),
], first_indent=False)

add_subheading("コーティング特性評価")
add_body([
    ("表面ゼータ電位はSurPASS 3で測定し、組み立て過程はQCM-dで、表面粗さはAFMで評価した。"
     "水和挙動はQCM-dにより周波数（ΔF）と散逸（ΔD）を同時測定して解析した。", False),
], first_indent=False)

add_subheading("抗菌性評価")
add_body([
    ("S. aureusに対するPS/PQおよびPA/PQ混合物のMICを微量液体希釈法で測定した。接触殺菌試験では、"
     "菌懸濁液をサンプル表面に滴下して培養後、剥離した菌をMueller-Hinton寒天培地に播種してコロニー数を"
     "計数した。CLSMによる生死染色およびSEMによる形態観察も行った。", False),
], first_indent=False)

add_subheading("全血循環試験")
add_body([
    ("PLA、PS2.0/2.5、PA2.0/2.5サンプルをカテーテルに入れ、ウサギ全血を充填した遠心管に接続して"
     "体外循環ループを構築し、蠕動ポンプで2時間循環させた後、SEMおよびFDA染色で表面付着を評価した。", False),
], first_indent=False)

add_subheading("in vivo動物試験")
add_body([
    ("全ての動物実験は中国動物保護委員会および四川大学の倫理基準（倫理番号KS2022756）に従って実施した。"
     "S. aureusを付着させたPGLA縫合糸をSprague-Dawleyラットの大腿内側筋に埋め込み、1日後・3日後に"
     "回収して平板計数法で生菌を計数し、周囲組織の炎症反応を病理切片で評価した。ウサギ頸静脈埋め込み"
     "モデルも同様に評価した。", False),
], first_indent=False)

add_subheading("統計解析")
add_body([
    ("全試験は3回以上反復し、結果は平均±標準偏差で示した。GraphPad Prismを用いて一元配置分散分析と"
     "LSD検定を行い、*p＜0.05、**p＜0.01、***p＜0.001を有意とした。", False),
], first_indent=False)

out = "/home/user/my-first-claude/雑誌会要旨_PA-PQコーティング.docx"
doc.save(out)
print("保存しました:", out)
