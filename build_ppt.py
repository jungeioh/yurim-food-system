#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# ---------- theme ----------
NAVY   = RGBColor(0x1F, 0x33, 0x55)
BLUE   = RGBColor(0x2E, 0x5C, 0x8A)
ACCENT = RGBColor(0x00, 0x7A, 0x6E)   # teal
RED    = RGBColor(0xC0, 0x2A, 0x2A)
GRAY   = RGBColor(0x44, 0x44, 0x44)
LGRAY  = RGBColor(0xEE, 0xF1, 0xF5)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
FONT   = "맑은 고딕"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

def set_font(run, size=18, bold=False, color=GRAY, font=FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set("typeface", font)

def box(slide, l, t, w, h, fill=None, line=None):
    sp = slide.shapes.add_shape(1, l, t, w, h)  # rectangle
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(1)
    return sp

def txt(slide, l, t, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, wrap=True):
    tb = slide.shapes.add_textbox(l, t, w, h); tf = tb.text_frame
    tf.word_wrap = wrap; tf.vertical_anchor = anchor
    tf.margin_left=Pt(4); tf.margin_right=Pt(4); tf.margin_top=Pt(2); tf.margin_bottom=Pt(2)
    first = True
    for ln in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = ln.get("align", align)
        p.space_after = Pt(ln.get("sa", 6)); p.space_before = Pt(ln.get("sb", 0))
        if "level" in ln: p.level = ln["level"]
        runs = ln["runs"] if "runs" in ln else [ln]
        for rr in runs:
            r = p.add_run(); r.text = rr["t"]
            set_font(r, rr.get("size",18), rr.get("bold",False), rr.get("color",GRAY))
    return tb

def header(slide, title, num):
    box(slide, 0, 0, SW, Inches(1.05), fill=NAVY)
    box(slide, 0, Inches(1.05), SW, Pt(4), fill=ACCENT)
    txt(slide, Inches(0.55), 0, Inches(11.5), Inches(1.05),
        [{"t":title,"size":26,"bold":True,"color":WHITE}], anchor=MSO_ANCHOR.MIDDLE)
    txt(slide, Inches(12.2), 0, Inches(0.9), Inches(1.05),
        [{"t":num,"size":14,"bold":True,"color":RGBColor(0x9F,0xB4,0xCC)}],
        align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

def bullet(t, size=18, bold=False, color=GRAY, level=0, sa=8, mark="•  "):
    pre = "" if level>0 else mark
    return {"runs":[{"t":pre,"size":size,"bold":True,"color":ACCENT},
                    {"t":t,"size":size,"bold":bold,"color":color}],
            "level":level,"sa":sa}

def sub(t, size=15, color=GRAY, sa=6):
    return {"runs":[{"t":"–  ","size":size,"bold":True,"color":BLUE},
                    {"t":t,"size":size,"color":color}],"level":1,"sa":sa}

# ================= Slide 1 : Title =================
s = prs.slides.add_slide(BLANK)
box(s,0,0,SW,SH, fill=NAVY)
box(s,0,Inches(2.55),SW,Pt(3), fill=ACCENT)
box(s,0,Inches(4.35),SW,Pt(3), fill=ACCENT)
txt(s, Inches(1), Inches(2.7), Inches(11.3), Inches(1.7),
    [{"t":"타다라필(Tadalafil) 검출 원인 분석","size":40,"bold":True,"color":WHITE,"sa":6},
     {"t":"삼지구엽초(음양곽) 고농축 남성 성기능 보조제","size":22,"bold":False,"color":RGBColor(0xB9,0xCB,0xDE)}])
txt(s, Inches(1), Inches(5.0), Inches(11.3), Inches(1.5),
    [{"t":"– 천연 생성설 검토 및 대응 설명서 –","size":18,"color":RGBColor(0x9F,0xB4,0xCC),"sa":18},
     {"t":"노마드오아시스(주)   |   2026. 06","size":16,"color":WHITE}])

# ================= Slide 2 : 핵심 결론 =================
s = prs.slides.add_slide(BLANK); header(s,"핵심 결론 요약","01")
box(s, Inches(0.55), Inches(1.35), Inches(12.2), Inches(1.15), fill=RED)
txt(s, Inches(0.8), Inches(1.35), Inches(11.7), Inches(1.15),
    [{"t":"삼지구엽초를 아무리 고농축해도 타다라필은 만들어지지 않습니다.","size":22,"bold":True,"color":WHITE}],
    anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(0.6), Inches(2.8), Inches(12.1), Inches(4.4),
    [bullet("타다라필은 식물에 존재하지 않는 100% 인공 합성 의약품입니다.",bold=True,color=NAVY),
     sub("제약사가 PDE5만 선택 억제하도록 설계한 '테트라하이드로-β-카볼린' 구조 화합물(시알리스)"),
     bullet("삼지구엽초의 천연 유효성분은 이카린(Icariin) — 타다라필과 화학 골격이 완전히 다른 별개 물질",bold=True,color=NAVY),
     sub("이카린(플라보노이드)에는 타다라필 핵심인 질소 함유 β-카볼린 고리가 아예 없음"),
     bullet("농축·가열로 이카린이 타다라필로 변환되는 화학 경로는 존재하지 않음",bold=True,color=NAVY),
     bullet("'천연 생성'을 입증한 논문은 없음 — 관련 문헌은 전부 '불법 혼입(adulteration)' 검출 사례",bold=True,color=NAVY),
     {"runs":[{"t":"⇒ 검출은 천연 유래가 아니라 공급망 어딘가의 '혼입(오염)'으로 보아야 하며,","size":18,"bold":True,"color":RED}],"sa":2,"sb":6},
     {"runs":[{"t":"     원료·제조 전반에 대한 즉각적 원인 추적 조사가 필요합니다.","size":18,"bold":True,"color":RED}],"sa":2}])

# ================= Slide 3 : 제품 개요 =================
s = prs.slides.add_slide(BLANK); header(s,"검출 대상 제품 개요 — 14종 복합 원료","02")
txt(s, Inches(0.6), Inches(1.3), Inches(12.1), Inches(0.6),
    [{"t":"시험의뢰서(한국기능식품연구원, 2026-06-12 접수) 기준, 단일 원료가 아닌 14종 한약재 복합 처방",
      "size":16,"bold":True,"color":NAVY}])
# table
rows, cols = 8, 4
tbl = s.shapes.add_table(rows, cols, Inches(0.6), Inches(2.0), Inches(12.1), Inches(4.4)).table
data = [["1  구기자","4  복분자","7  산약","10  오미자"],
        ["2  목단피","★5  사상자","8  숙지황","11  육계"],
        ["3  복령","6  산수유","★9  야관문","★12  음양곽(삼지구엽초)"],
        ["★14  토사자","13  택사","","" ]]
# fill 8 rows: pair header rows
hdr = ["원료","원료","원료","원료"]
allrows = [hdr] + data + [["","","",""]]*(rows-1-len(data))
for ci in range(cols):
    c = tbl.cell(0,ci); c.fill.solid(); c.fill.fore_color.rgb = NAVY
    p=c.text_frame.paragraphs[0]; r=p.add_run(); r.text=hdr[ci]; set_font(r,14,True,WHITE); p.alignment=PP_ALIGN.CENTER
ri=1
for drow in data:
    for ci in range(cols):
        c=tbl.cell(ri,ci); c.fill.solid()
        c.fill.fore_color.rgb = LGRAY if ri%2 else WHITE
        val=drow[ci]; star = val.startswith("★")
        p=c.text_frame.paragraphs[0]; r=p.add_run(); r.text=val.replace("★","")
        set_font(r,14,star, RED if star else GRAY); p.alignment=PP_ALIGN.CENTER
    ri+=1
for rr in range(ri, rows):
    for ci in range(cols):
        c=tbl.cell(rr,ci); c.fill.solid(); c.fill.fore_color.rgb=WHITE
txt(s, Inches(0.6), Inches(6.55), Inches(12.1), Inches(0.7),
    [{"runs":[{"t":"★ ","size":14,"bold":True,"color":RED},
              {"t":"성기능 관련 한약재(사상자·야관문·음양곽·토사자 등)는 국내외에서 타다라필 불법 혼입이 가장 빈번한 품목군 → 혼입 가능 지점이 넓음","size":14,"color":NAVY}]}])

# ================= Slide 4 : 타다라필이란 =================
s = prs.slides.add_slide(BLANK); header(s,"타다라필(Tadalafil)이란 무엇인가","03")
txt(s, Inches(0.6), Inches(1.4), Inches(12.1), Inches(5.6),
    [bullet("분류 : PDE5 억제제 — 발기부전 치료 전문의약품 (상품명 '시알리스/Cialis')",bold=True,color=NAVY),
     bullet("CAS 번호 : 171596-29-5",color=GRAY),
     bullet("화학적 본질 : 테트라하이드로-β-카볼린(tetrahydro-β-carboline) 골격 기반",bold=True,color=NAVY),
     sub("제약사(ICOS/Lilly)가 구조 기반 설계(structure-based design)로 합성·최적화한 인공 화합물"),
     bullet("입체화학 : 두 개의 키랄 탄소, 시판품은 (6R,12aR) 단일 이성질체",bold=True,color=NAVY),
     sub("자연 식물 대사에서 이런 정밀 입체선택적 합성이 우연히 일어날 가능성은 없음"),
     bullet("FDA 승인 : 2003년 12월",color=GRAY),
     {"runs":[{"t":"⇒ ","size":19,"bold":True,"color":ACCENT},
              {"t":"타다라필은 '정밀하게 합성·설계된 의약품'이며, 자연계(식물)에는 존재하지 않습니다.","size":19,"bold":True,"color":NAVY}],"sb":10}])

# ================= Slide 5 : 구조 비교 표 =================
s = prs.slides.add_slide(BLANK); header(s,"삼지구엽초 천연성분 vs 타다라필 — 구조 비교","04")
rows, cols = 6, 3
tbl = s.shapes.add_table(rows, cols, Inches(0.6), Inches(1.45), Inches(12.1), Inches(5.0)).table
tbl.columns[0].width=Inches(2.7); tbl.columns[1].width=Inches(4.7); tbl.columns[2].width=Inches(4.7)
content=[["구분","이카린 (삼지구엽초 천연성분)","타다라필 (합성 의약품)"],
         ["화학 분류","플라보노이드 배당체 (폴리페놀+당)","테트라하이드로-β-카볼린 유도체"],
         ["기본 골격","플라본 골격 + 당 + 프레닐기","인돌-피리도-피라진다이온 융합 고리"],
         ["질소 원자","없음 (플라보노이드 골격)","핵심 골격에 다수 포함"],
         ["유래","식물 천연 생합성","실험실 화학 합성"],
         ["PDE5 억제력","약함 (IC50 ~1–6 μM)","매우 강함 (선택적)"]]
for ri in range(rows):
    for ci in range(cols):
        c=tbl.cell(ri,ci); c.fill.solid()
        if ri==0: c.fill.fore_color.rgb=NAVY
        elif ci==0: c.fill.fore_color.rgb=BLUE
        else: c.fill.fore_color.rgb = LGRAY if ri%2 else WHITE
        p=c.text_frame.paragraphs[0]; r=p.add_run(); r.text=content[ri][ci]
        col = WHITE if (ri==0 or ci==0) else (ACCENT if ci==1 else RED if ci==2 else GRAY)
        set_font(r,15, ri==0 or ci==0, col); p.alignment=PP_ALIGN.CENTER if ci==0 or ri==0 else PP_ALIGN.LEFT
txt(s, Inches(0.6), Inches(6.6), Inches(12.1), Inches(0.7),
    [{"t":"⇒ 두 물질은 탄소 골격 자체가 근본적으로 다름. 이카린 농축으로 '없던 질소 골격'이 새로 만들어질 수 없음.",
      "size":15,"bold":True,"color":NAVY}])

# ================= Slide 6 : 천연생성설 검토 =================
s = prs.slides.add_slide(BLANK); header(s,"'고농축하면 타다라필이 검출된다'설 검토","05")
box(s, Inches(0.55), Inches(1.3), Inches(12.2), Inches(0.85), fill=RED)
txt(s, Inches(0.8), Inches(1.3), Inches(11.7), Inches(0.85),
    [{"t":"결론 : 근거 없음 (과학적으로 성립하지 않음)","size":20,"bold":True,"color":WHITE}], anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(0.6), Inches(2.4), Inches(12.1), Inches(4.7),
    [bullet("천연 생성을 입증한 논문은 존재하지 않음",bold=True,color=NAVY),
     sub("Epimedium + tadalafil 문헌은 전부 '불법 혼입·위조 제품에서 타다라필 검출' 내용"),
     bullet("'농축하면 검출된다'의 진짜 의미 — 생성이 아니라 '가시화'",bold=True,color=NAVY,sa=4),
     sub("원료에 (어떤 경로로든) 미량 혼입돼 있던 타다라필이 고농축 공정에서 함께 농축"),
     sub("검출한계(LOD)를 넘어서면서 비로소 '드러나는' 현상일 뿐 (없던 물질을 만드는 것이 아님)"),
     bullet("타다라필 유사체(designer analogue) 문제",bold=True,color=NAVY,sa=4),
     sub("단속 회피용으로 구조를 변형한 유사체를 혼입 → 분석기가 타다라필 계열로 보고하기도 함"),
     sub("이 역시 천연 생성이 아니라 인위적 혼입")])

# ================= Slide 7 : 한방자료(淫羊藿黄酮) 검토 =================
s = prs.slides.add_slide(BLANK); header(s,"부록 : 한방자료(淫羊藿黄酮) 검토","06")
txt(s, Inches(0.6), Inches(1.3), Inches(12.1), Inches(5.8),
    [bullet("성격 : 학술 자료라기보다 효능 표방형 마케팅·홍보성 설명문 → 규제 근거자료로는 부적합",bold=True,color=NAVY),
     bullet("맞는 부분(딱 하나) : 이카린·이카리사이드가 PDE5를 억제한다 (단, 약하고 느림 — 자료도 인정)",color=GRAY,sa=4),
     bullet("그러나 이 자료가 오히려 증명하는 것 :",bold=True,color=RED,sa=2),
     sub("자료의 비교표가 '淫羊藿黄酮(천연)'과 '他达拉非 타다라필(합성)'을 서로 다른 두 물질로 명확히 구분"),
     sub("즉 이 자료조차 '삼지구엽초 = 타다라필' 또는 '농축 시 타다라필 생성'이라고 말하지 않음"),
     {"runs":[{"t":"핵심 함정 : ","size":18,"bold":True,"color":RED},
              {"t":"'작용 통로가 같다(同源)' ≠ '같은 분자'","size":18,"bold":True,"color":NAVY}],"sb":6,"sa":2},
     sub("커피 카페인과 약국 각성제가 '잠을 쫓는' 효과는 같아도 분자는 완전히 다른 것과 동일"),
     sub("둘 다 PDE5를 친다는 사실과, 기기에 '타다라필'이라는 특정 분자가 잡힌 것은 별개의 문제"),
     {"runs":[{"t":"⇒ 이 자료는 '천연 생성' 소명의 근거가 될 수 없음.","size":17,"bold":True,"color":RED}],"sb":6}])

# ================= Slide 8 : 검사 위양성 확인 =================
s = prs.slides.add_slide(BLANK); header(s,"반드시 확인 — '검사가 이카린을 오인했나?'","07")
txt(s, Inches(0.6), Inches(1.35), Inches(12.1), Inches(2.0),
    [bullet("LC-MS/MS(질량분석) 확정시험이라면 → 오인 불가능",bold=True,color=NAVY),
     sub("타다라필은 고유 분자량·질량파편(m/z)으로 식별 / 이카린은 분자량이 완전히 달라 오인될 수 없음"),
     bullet("단순 색반응·간이키트·TLC만으로 했다면 → 위양성 가능성 검토 여지 있음",bold=True,color=NAVY)])
box(s, Inches(0.6), Inches(3.5), Inches(12.1), Inches(0.55), fill=ACCENT)
txt(s, Inches(0.8), Inches(3.5), Inches(11.7), Inches(0.55),
    [{"t":"검사기관에 반드시 확인할 3가지","size":17,"bold":True,"color":WHITE}], anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(0.6), Inches(4.25), Inches(12.1), Inches(2.8),
    [bullet("① 분석법 — LC-MS/MS 확정시험인가, 단순 스크리닝인가",bold=True,color=NAVY),
     bullet("② 검출 물질 — 타다라필 본체인가, 유사체(analogue)인가",bold=True,color=NAVY),
     bullet("③ 정량값 — 몇 mg/회분, ppm 인가 (미량=혼입/교차오염, 유효량=명백한 첨가)",bold=True,color=NAVY),
     {"runs":[{"t":"⇒ LC-MS/MS에서 본체가 유효량으로 검출됐다면, 한방자료로 설명 불가 → 혼입으로 결론.","size":16,"bold":True,"color":RED}],"sb":8}])

# ================= Slide 9 : 혼입 경로 =================
s = prs.slides.add_slide(BLANK); header(s,"실제 가능한 원인 — 혼입 경로","08")
txt(s, Inches(0.6), Inches(1.4), Inches(12.1), Inches(5.6),
    [bullet("① 원료(한약재) 공급단계 부정혼입  ★가능성 높음",bold=True,color=NAVY),
     sub("성기능 관련 한약재는 타다라필·실데나필 불법 혼입이 가장 빈번 / 일부 원료상이 분말 단계에서 첨가"),
     sub("→ 14종 원료 각각의 입고검사 성적서(CoA)·공급사 추적 필요"),
     bullet("② 제조시설 교차오염 (Cross-contamination)",bold=True,color=NAVY,sa=4),
     sub("동일 설비/라인에서 타다라필 함유 제품 취급 이력 + 세척 불충분 시 미량 혼입"),
     bullet("③ 용수·부형제·캡슐 등 부자재 오염 (가능성 낮으나 배제 불가)",bold=True,color=NAVY,sa=4),
     bullet("④ 시험 단계 오염 또는 분석 오류",bold=True,color=NAVY,sa=4),
     sub("재시험·교차검증으로 확인 (재현되면 오염, 재현 안 되면 분석오류 검토)")])

# ================= Slide 10 : Action Plan =================
s = prs.slides.add_slide(BLANK); header(s,"대응 조치 (Action Plan)","09")
txt(s, Inches(0.55), Inches(1.3), Inches(6.0), Inches(5.9),
    [{"runs":[{"t":"A. 즉시 — 분석·사실 확인","size":17,"bold":True,"color":ACCENT}],"sa":6},
     sub("검출 성적서 원본 확보(물질명·정량값·검출한계)"),
     sub("공인기관 2곳 이상 재시험·교차분석"),
     sub("14종 원료 개별 단독 분석 → 혼입 원료 핀포인트 (가장 결정적)"),
     {"runs":[{"t":"B. 공급망 추적","size":17,"bold":True,"color":ACCENT}],"sa":6,"sb":8},
     sub("원료별 공급사·로트·CoA 전수 확보"),
     sub("동일 로트 사용 타 배치도 함께 검사"),
     sub("공급사에 PDE5억제제 무첨가 시험성적서 요구")])
txt(s, Inches(6.85), Inches(1.3), Inches(6.0), Inches(5.9),
    [{"runs":[{"t":"C. 제조 환경","size":17,"bold":True,"color":ACCENT}],"sa":6},
     sub("라인 이력(타다라필 제품 취급)·세척검증(CIP) 점검"),
     sub("설비·작업대 잔류물(swab) 검사"),
     {"runs":[{"t":"D. 규제·법적 대응","size":17,"bold":True,"color":ACCENT}],"sa":6,"sb":8},
     sub("사실 기반 정직 소명 ('투입 안 함, 혼입 경로 조사 중')"),
     sub("식품·약사 전문 변호사 자문 병행"),
     sub("혼입 확인 시 자진 회수·판매중단 검토"),
     {"runs":[{"t":"E. 소비자 안전 (최우선)","size":17,"bold":True,"color":RED}],"sa":6,"sb":8},
     sub("질산염 복용자 치명적 저혈압 위험 → 회수·고지 적극 검토")])

# ================= Slide 11 : 경고 =================
s = prs.slides.add_slide(BLANK); header(s,"⚠️ 반드시 숙지 — '천연 생성' 소명 금지","10")
rows, cols = 4, 2
tbl = s.shapes.add_table(rows, cols, Inches(0.6), Inches(1.5), Inches(12.1), Inches(3.4)).table
tbl.columns[0].width=Inches(5.6); tbl.columns[1].width=Inches(6.5)
cc=[["이렇게 주장하면","규제기관/전문가의 반박"],
    ["'삼지구엽초를 농축하니 타다라필이 생겼다'","타다라필은 합성 의약품으로 식물 생성 불가. β-카볼린 구조는 이카린(플라보노이드)과 무관"],
    ["'관련 논문이 있다'","해당 문헌은 전부 '불법 혼입·위조' 검출 논문. 천연 생성 입증 논문은 없음"],
    ["결과","허위 소명·고의 은폐로 간주 → 가중 처벌 위험"]]
for ri in range(rows):
    for ci in range(cols):
        c=tbl.cell(ri,ci); c.fill.solid()
        c.fill.fore_color.rgb = NAVY if ri==0 else (RGBColor(0xF7,0xE3,0xE3) if ri==rows-1 else (LGRAY if ri%2 else WHITE))
        p=c.text_frame.paragraphs[0]; r=p.add_run(); r.text=cc[ri][ci]
        set_font(r,15, ri==0 or ri==rows-1, WHITE if ri==0 else (RED if ri==rows-1 else GRAY))
        p.alignment=PP_ALIGN.LEFT
box(s, Inches(0.6), Inches(5.25), Inches(12.1), Inches(1.3), fill=ACCENT)
txt(s, Inches(0.85), Inches(5.25), Inches(11.6), Inches(1.3),
    [{"t":"⇒ 정직한 원인조사 + 신속한 안전조치가","size":20,"bold":True,"color":WHITE,"sa":2},
     {"t":"     법적·평판 리스크를 최소화하는 유일한 길입니다.","size":20,"bold":True,"color":WHITE}],
    anchor=MSO_ANCHOR.MIDDLE)

# ================= Slide 12 : 참고문헌 =================
s = prs.slides.add_slide(BLANK); header(s,"참고 문헌 / 출처","11")
txt(s, Inches(0.6), Inches(1.4), Inches(12.1), Inches(5.7),
    [sub("Tadalafil 화학정보 (테트라하이드로-β-카볼린, CAS 171596-29-5) — ChemicalBook",sa=8),
     sub("Daugan A. et al., The discovery of tadalafil: a novel and highly selective PDE5 inhibitor — PubMed 14521414",sa=8),
     sub("Icariin이 Epimedium의 천연 유효성분이며 약한 PDE5 억제 활성 — OPSS, Horny Goat Weed",sa=8),
     sub("Identification and screening of a tadalafil analogue found in adulterated herbal products — PubMed 25402198",sa=8),
     sub("Detection of a Tadalafil Analogue as an Adulterant in a Dietary Supplement for ED — academia.edu",sa=8),
     sub("Strategies for characterizing sildenafil/vardenafil/tadalafil and their analogues in herbal supplements",sa=8),
     sub("Horny Goat Weed — LiverTox / NCBI Bookshelf NBK583203",sa=8)])
txt(s, Inches(0.6), Inches(6.65), Inches(12.1), Inches(0.7),
    [{"t":"※ 본 자료는 공개 과학문헌·화학정보에 근거하며 법적 자문이 아님. 규제대응·회수결정은 전문 변호사·공인분석기관과 협의 요망.",
      "size":12,"color":RGBColor(0x88,0x88,0x88)}])

prs.save("/home/user/yurim-food-system/타다라필_검출_원인분석_설명서.pptx")
print("saved", len(prs.slides.__iter__.__self__._sldIdLst), "slides")
