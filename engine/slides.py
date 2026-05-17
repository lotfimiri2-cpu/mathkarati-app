"""
CANVA Engine v19.1 — Dynamic Editorial  ★ TEXT UPGRADE
"""
from __future__ import annotations
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, cm, pt, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_3stop, gradient_rect, shadow, glow,
    shadow_and_glow, set_solid_alpha,
    txt, txt_hero, txt_label, txt_body, txt_stat, txt_quote,
    blank_slide,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"
def set_font(f): global _FONT; _FONT = f

def _rgb(h):
    h = h.lstrip('#')
    return RGBColor(int(h[:2],16), int(h[2:4],16), int(h[4:],16))

def _CW(): return W - 1.6

def _hdr(slide, T, title, sub=""):
    hb = rect(slide, 0, 0, W, 3.2, _rgb(T.bg2))
    if hb: gradient_fill(hb, T.grad1, T.grad2, 178)
    # دوائر زخرفية في الهيدر
    oval(slide, W-7.5, -3.8, 11, 11, _rgb(T.accent), alpha=7)
    oval(slide, -2.5, -1.2, 8, 8, _rgb(T.accent2), alpha=5)
    oval(slide, W*0.45, -4, 9, 9, _rgb(T.bg2), alpha=40)
    bar = rect(slide, 0, 3.06, W, 0.14, _rgb(T.accent))
    if bar: gradient_fill(bar, T.accent_grad1, T.accent_grad2, 0)
    sz = 25 if len(title)<18 else 20 if len(title)<30 else 16
    txt_hero(slide, title, 1.3, 0.38, W-2.6, 1.9,
             font=_FONT, size=sz, color=_rgb(T.text_light),
             align=PP_ALIGN.RIGHT, rtl=True, shadow_on=True)
    if sub:
        txt_label(slide, "◈  "+sub, 1.3, 2.12, W-2.6, 0.88,
                  font=_FONT, size=11.5, color=_rgb(T.muted),
                  align=PP_ALIGN.RIGHT, rtl=True, uppercase=False)

def _card(slide, x, y, w, h, T, radius=12):
    c = rrect(slide, x, y, w, h, _rgb(T.card), radius_pct=radius)
    if c:
        gradient_fill(c, T.card, T.bg2, 135)
        shadow(c, blur=22, dist=7, alpha=0.52)
    return c

def _pill(slide, T, x, y, w, h, text, fsize=10):
    p = rrect(slide, x, y, w, h, _rgb(T.accent), radius_pct=50)
    if p:
        gradient_fill(p, T.accent_grad1, T.accent_grad2, 0)
        shadow(p, blur=6, dist=2, alpha=0.28)
    txt_label(slide, text, x, y, w, h,
              font=_FONT, size=fsize, color=_rgb(T.text_dark),
              align=PP_ALIGN.CENTER, rtl=True, uppercase=False)

# ═════════════════════════════════════════════════════════════════════
def make_cover(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 148)
    # زخارف ضخمة
    oval(slide, -3.5, -3.5, 10, 10, _rgb(T.accent), alpha=6)
    oval(slide, W*0.52, -5, 14, 14, _rgb(T.bg2), alpha=35)
    oval(slide, W-6, H-5, 10, 10, _rgb(T.accent), alpha=5)
    # مستطيل مائل خفي
    diag = rect(slide, W*0.58, 0, W*0.48, H, _rgb(T.bg2))
    if diag:
        gradient_fill(diag, T.bg2, T.bg, 0)
        set_solid_alpha(diag, 16)
    tb2 = rect(slide, 0, 0, W, 0.55, _rgb(T.accent))
    if tb2: gradient_fill(tb2, T.accent_grad1, T.accent_grad2, 0)

    if req.institution:
        inst = req.institution.split(' — ')[0]
        ic = rrect(slide, 1.6, 0.78, W-3.2, 0.74, _rgb(T.card), radius_pct=40)
        if ic:
            shadow(ic, blur=8, dist=2, alpha=0.28)
            set_solid_alpha(ic, 68)
        txt_body(slide, inst, 1.8, 0.78, W-3.6, 0.74,
                 font=_FONT, size=11, color=_rgb(T.muted),
                 align=PP_ALIGN.CENTER, rtl=True)

    # بطاقة العنوان الرئيسية
    cy = H*0.19; ch = H*0.4
    mc = rrect(slide, 2.4, cy, W-4.8, ch, _rgb(T.card), radius_pct=14)
    if mc:
        gradient_fill(mc, T.card, T.bg2, 135)
        shadow(mc, blur=30, dist=10, alpha=0.58)
    ct = rect(slide, 2.4, cy, W-4.8, 0.65, _rgb(T.accent))
    if ct: gradient_fill(ct, T.accent_grad1, T.accent_grad2, 0)

    tsz = 28 if len(req.title_ar)<35 else 21 if len(req.title_ar)<58 else 16 if len(req.title_ar)<85 else 13
    txt_hero(slide, req.title_ar, 3.0, cy+0.85, W-6.0, ch-1.35,
             font=_FONT, size=tsz, color=_rgb(T.text_light),
             align=PP_ALIGN.CENTER, rtl=True, shadow_on=True)

    if req.title_en:
        txt(slide, req.title_en, 3.0, cy+ch-0.95, W-6.0, 0.82,
            font="Calibri", size=10.5, italic=True,
            color=_rgb(T.muted), align=PP_ALIGN.CENTER,
            letter_spacing=0.8)

    # بطاقات معلومات
    fields = [("الطالب", req.student_name)]
    if req.supervisor:     fields.append(("المشرف", req.supervisor))
    if req.co_supervisor:  fields.append(("م. مساعد", req.co_supervisor))
    if req.specialization: fields.append(("التخصص", req.specialization))
    n = len(fields)
    pw = min((W-3.0)/n-0.3, 8.4)
    tot = n*(pw+0.3)-0.3; sx = (W-tot)/2
    iy = cy+ch+0.45
    for i,(lbl,val) in enumerate(fields[:4]):
        px = sx+i*(pw+0.3)
        pc = rrect(slide, px, iy, pw, 2.05, _rgb(T.card), radius_pct=12)
        if pc:
            shadow(pc, blur=14, dist=5, alpha=0.42)
            set_solid_alpha(pc, 85)
        lb = rrect(slide, px, iy, pw, 0.58, _rgb(T.accent), radius_pct=12)
        if lb: gradient_fill(lb, T.accent_grad1, T.accent_grad2, 0)
        txt_label(slide, lbl, px, iy+0.04, pw, 0.52,
                  font=_FONT, size=10, color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER, rtl=True, uppercase=False)
        txt_body(slide, val, px+0.18, iy+0.66, pw-0.36, 1.28,
                 font=_FONT, size=12, color=_rgb(T.text_light),
                 align=PP_ALIGN.CENTER, rtl=True)

    if req.year:
        yr = rrect(slide, W/2-2.7, H-1.38, 5.4, 0.78, _rgb(T.accent), radius_pct=50)
        if yr:
            gradient_fill(yr, T.accent_grad1, T.accent_grad2, 0)
            shadow(yr, blur=8, dist=2, alpha=0.32)
        txt_label(slide, req.year, W/2-2.7, H-1.38, 5.4, 0.78,
                  font="Calibri", size=13, color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER, rtl=False, uppercase=False)
    return slide

def make_intro(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,158)
    _hdr(slide,T,"مقدمة البحث","نظرة عامة ومقاربة الدراسة")
    cy=3.28; ch=H-cy-0.42
    if req.intro_overview and req.intro_approach:
        cw1=_CW()*0.56; cw2=_CW()-cw1-0.38
        for titl,text,cw,cx in [
            ("الإطار العام",req.intro_overview,cw1,0.8),
            ("المقاربة",req.intro_approach,cw2,0.8+cw1+0.38)]:
            # بطاقة نظيفة بدون دوائر داخلية
            c=rrect(slide,cx,cy,cw,ch,_rgb(T.card),radius_pct=10)
            if c:
                gradient_fill(c,T.card,T.bg2,135)
                shadow(c,blur=20,dist=7,alpha=0.48)
            # شريط علوي مزخرف
            tp=rect(slide,cx,cy,cw,0.62,_rgb(T.accent))
            if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
            # خط جانبي يساري
            vline(slide,cx,cy,ch,_rgb(T.accent),thickness=0.22)
            txt_label(slide,titl,cx+0.4,cy+0.07,cw-0.6,0.52,
                      font=_FONT,size=14,color=_rgb(T.text_dark),
                      align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
            txt_body(slide,text,cx+0.45,cy+0.75,cw-0.75,ch-0.95,
                     font=_FONT,size=13,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
    else:
        text=req.intro_overview or req.intro_approach
        c=rrect(slide,0.8,cy,_CW(),ch,_rgb(T.card),radius_pct=10)
        if c:
            gradient_fill(c,T.card,T.bg2,135)
            shadow(c,blur=20,dist=7,alpha=0.48)
        # شريط جانبي لوني
        lb=rrect(slide,0.8,cy,0.52,ch,_rgb(T.accent),radius_pct=10)
        if lb: gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
        txt(slide,"❝",1.55,cy+0.3,2.2,2.0,
            font="Calibri",size=58,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.LEFT,txt_shadow=True)
        txt_quote(slide,text,1.6,cy+0.7,_CW()-1.2,ch-0.95,
                  font=_FONT,size=14.5,color=_rgb(T.text_light),rtl=True)
    return slide

def make_plan(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,158)
    _hdr(slide,T,"خطة البحث","محتويات الدراسة وهيكلها")
    chapters=req.chapters[:8]
    if not chapters: return slide
    cy=3.28; avail=H-cy-0.48
    n=len(chapters)
    if n<=4: cols=n; rows=1
    elif n<=6: cols=3; rows=2
    else: cols=4; rows=(n+3)//4
    gap=0.28
    cw=(_CW()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,3.85)
    for i,chap in enumerate(chapters):
        ci=i%cols; ri=i//cols
        x=0.8+ci*(cw+gap); y=cy+ri*(ch+gap)
        _card(slide,x,y,cw,ch,T,radius=11)
        tp=rect(slide,x,y,cw,0.58,_rgb(T.accent))
        if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
        txt(slide,f"{i+1:02d}",x+0.14,y+0.06,1.05,0.5,
            font="Calibri",size=18,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.LEFT,
            letter_spacing=-1.0)
        hline(slide,x+0.2,y+0.72,cw-0.4,_rgb(T.muted),thickness=0.04)
        txt_body(slide,chap.title,x+0.2,y+0.84,cw-0.4,ch-1.22,
                 font=_FONT,size=14.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
        if chap.pages:
            txt_label(slide,f"ص {chap.pages}",x+0.18,y+ch-0.52,cw-0.36,0.42,
                      font="Calibri",size=11,color=_rgb(T.muted),
                      align=PP_ALIGN.LEFT,rtl=False,uppercase=False)
    return slide

def make_problem(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,158)
    _hdr(slide,T,"إشكالية البحث","التساؤلات المحورية للدراسة")
    cy=3.28
    if req.main_problem or req.main_question:
        text=req.main_problem or req.main_question
        qh=min(H*0.31,3.05)
        qc=_card(slide,0.8,cy,_CW(),qh,T,radius=11)
        hline(slide,0.8,cy+qh-0.14,_CW(),_rgb(T.accent),thickness=0.14)
        rbar=rrect(slide,W-1.45,cy,0.65,qh,_rgb(T.accent2),radius_pct=11)
        if rbar:
            gradient_fill(rbar,T.accent_grad2,T.accent_grad1,90)
            set_solid_alpha(rbar,65)
        txt(slide,"❝",1.18,cy+0.14,2.5,qh-0.28,
            font="Calibri",size=54,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.RIGHT,txt_shadow=True)
        txt_quote(slide,text,1.4,cy+0.24,_CW()-2.2,qh-0.48,
                  font=_FONT,size=14,color=_rgb(T.text_light),rtl=True)
        cy+=qh+0.28
    if req.sub_questions:
        sq=req.sub_questions[:6]; n=len(sq)
        cols=2 if n>2 else 1
        cw=(_CW()-(cols-1)*0.28)/cols
        avail=H-cy-0.38
        sh=min(avail/((n+cols-1)//cols)-0.1,1.22)
        for i,q in enumerate(sq):
            ci=i%cols; ri=i//cols
            x=0.8+ci*(cw+0.28); y=cy+ri*(sh+0.1)
            rc=rrect(slide,x,y,cw,sh,_rgb(T.card),radius_pct=7)
            if rc: shadow(rc,blur=8,dist=2,alpha=0.22)
            lb=rrect(slide,x,y,0.4,sh,_rgb(T.accent),radius_pct=7)
            if lb: gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
            _pill(slide,T,x+cw-1.7,y+(sh-0.48)/2,1.55,0.48,str(i+1),10)
            txt_body(slide,q,x+0.58,y+0.1,cw-2.38,sh-0.2,
                     font=_FONT,size=13.5,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_objectives(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,158)
    _hdr(slide,T,"أهداف البحث وفرضياته","")
    cy=3.28; ch=H-cy-0.42; gap=0.38
    cols=[]
    if req.objectives: cols.append(("الأهداف",req.objectives))
    if req.hypotheses:  cols.append(("الفرضيات",req.hypotheses))
    if not cols: return slide
    widths=[_CW()*0.56,_CW()*0.42] if len(cols)==2 else [_CW()]
    x=0.8
    for i,(lbl,items) in enumerate(cols[:2]):
        cw=widths[i]
        _card(slide,x,cy,cw,ch,T,radius=11)
        hd=rect(slide,x,cy,cw,0.62,_rgb(T.accent))
        if hd: gradient_fill(hd,T.accent_grad1,T.accent_grad2,0)
        txt_label(slide,lbl,x+0.2,cy+0.07,cw-0.4,0.52,
                  font=_FONT,size=14.5,color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        ih=min((ch-0.7)/max(len(items),1),1.18)
        for j,item in enumerate(items[:8]):
            iy=cy+0.7+j*ih
            if iy+ih>cy+ch-0.08: break
            if j>0: hline(slide,x+0.22,iy,cw-0.44,_rgb(T.bg),thickness=0.03)
            _pill(slide,T,x+cw-1.48,iy+(ih-0.46)/2,1.32,0.46,str(j+1),10)
            txt_body(slide,item,x+0.22,iy+0.09,cw-1.9,ih-0.18,
                     font=_FONT,size=13,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
        x+=cw+gap
    return slide

def make_importance(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,158)
    _hdr(slide,T,"أهمية البحث","مبررات اختيار الموضوع")
    items=list(req.importance or [])
    if req.reasons and req.reasons not in items: items.append(req.reasons)
    items=items[:6]
    if not items: return slide
    ah_list=[0.65,0.82,0.56,0.74,0.62,0.78]
    cols=3 if len(items)>=4 else 2 if len(items)>=2 else 1
    rows=(len(items)+cols-1)//cols; gap=0.28
    cy=3.28; avail=H-cy-0.4
    cw=(_CW()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,3.72)
    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=0.8+ci*(cw+gap); y=cy+ri*(ch+gap)
        _card(slide,x,y,cw,ch,T,radius=9)
        ah=ah_list[i%len(ah_list)]
        tp=rect(slide,x,y,cw,ah,_rgb(T.accent))
        if tp:
            gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
            set_solid_alpha(tp,92)
        txt(slide,f"{i+1:02d}",x+0.16,y+0.04,1.1,ah-0.08,
            font="Calibri",size=21,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.LEFT,
            letter_spacing=-1.0,txt_shadow=False)
        txt_body(slide,item,x+0.22,y+ah+0.18,cw-0.44,ch-ah-0.3,
                 font=_FONT,size=11.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_methodology(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,158)
    _hdr(slide,T,"منهجية البحث","الإجراءات والأدوات المستخدمة")
    fields=[]
    if req.methodology: fields.append(("المنهج","م",req.methodology))
    if req.sample_type: fields.append(("العينة","ع",req.sample_type))
    if req.sample_size: fields.append(("الحجم","ن",req.sample_size))
    if req.tool:        fields.append(("الأداة","أ",req.tool))
    if not fields: return slide
    n=len(fields); cy=3.28; ch=H-cy-0.46
    cw=(_CW()-0.28*(n-1))/n
    if n>1: hline(slide,0.8+cw/2,cy+ch/2,_CW()-cw,_rgb(T.muted),thickness=0.04)
    for i,(lbl,icon,val) in enumerate(fields[:4]):
        x=0.8+i*(cw+0.28)
        _card(slide,x,cy,cw,ch,T,radius=13)
        sz=1.78; ix=x+(cw-sz)/2
        ic=oval(slide,ix,cy+0.52,sz,sz,_rgb(T.accent))
        if ic:
            gradient_fill(ic,T.accent_grad1,T.accent_grad2,45)
            shadow_and_glow(ic,s_blur=12,s_dist=4,s_alpha=0.42,
                            g_color=T.accent_grad2,g_rad=6,g_alpha=0.25)
        txt(slide,icon,ix,cy+0.52,sz,sz,
            font=_FONT,size=26,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.CENTER)
        hline(slide,x+0.28,cy+2.52,cw-0.56,_rgb(T.accent),thickness=0.08)
        txt_label(slide,lbl,x+0.14,cy+2.64,cw-0.28,0.66,
                  font=_FONT,size=13,color=_rgb(T.accent),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        txt_body(slide,val,x+0.22,cy+3.4,cw-0.44,ch-3.56,
                 font=_FONT,size=11.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.CENTER,rtl=True)
    return slide

def make_stats(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,158)
    _hdr(slide,T,"الإحصاءات والأرقام الرئيسية","مؤشرات كمية للدراسة")
    stats=req.stats[:6]
    if not stats: return slide
    n=len(stats); cols=3 if n>=3 else n
    rows=(n+cols-1)//cols; gap=0.3
    cy=3.28; avail=H-cy-0.4
    cw=(_CW()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,4.65)
    for i,st in enumerate(stats):
        ci=i%cols; ri=i//cols
        x=0.8+ci*(cw+gap); y=cy+ri*(ch+gap)
        _card(slide,x,y,cw,ch,T,radius=13)
        cs=min(cw,ch)*0.76
        oval(slide,x+(cw-cs)/2,y+(ch-cs)/2-0.32,cs,cs,_rgb(T.accent),alpha=6)
        txt_stat(slide,st.value,x+0.2,y+0.44,cw-0.4,ch*0.5,
                 font="Calibri",color=_rgb(T.accent),align=PP_ALIGN.CENTER)
        if st.unit:
            txt_label(slide,st.unit,x+0.2,y+ch*0.5+0.28,cw-0.4,0.58,
                      font=_FONT,size=11,color=_rgb(T.muted),
                      align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        lw=min(len(st.label)*0.2+1.1,cw-0.36)
        _pill(slide,T,x+(cw-lw)/2,y+ch-1.12,lw,0.58,st.label,10)
    return slide

def make_results(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,158)
    _hdr(slide,T,"نتائج البحث","أبرز ما توصلت إليه الدراسة")
    results=req.main_results[:8]
    if not results: return slide
    cy=3.28; avail=H-cy-0.4
    ih=min(avail/max(len(results),1)-0.12,1.62)
    for i,res in enumerate(results):
        y=cy+i*(ih+0.12)
        row=rrect(slide,0.8,y,_CW(),ih,_rgb(T.bg2) if i%2==0 else _rgb(T.card),radius_pct=6)
        if row and i%2==0: shadow(row,blur=5,dist=1,alpha=0.15)
        _pill(slide,T,W-3.2,y+(ih-0.54)/2,2.08,0.54,str(i+1),12)
        vline(slide,0.8,y,ih,_rgb(T.accent),thickness=0.34)
        txt_body(slide,res,1.38,y+0.12,W-5.35,ih-0.24,
                 font=_FONT,size=13.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_conclusion(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,158)
    _hdr(slide,T,"خاتمة البحث","الاستنتاج العام")
    cy=3.28; ch=H-cy-0.52
    c=_card(slide,0.8,cy,_CW(),ch,T,radius=13)
    lb=rrect(slide,0.8,cy,3.6,ch,_rgb(T.accent),radius_pct=13)
    if lb:
        gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
        set_solid_alpha(lb,18)
    tp=rrect(slide,0.8,cy,_CW(),0.44,_rgb(T.accent),radius_pct=13)
    if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
    txt(slide,"❝",1.2,cy+0.55,3.1,2.7,
        font="Calibri",size=66,bold=True,
        color=_rgb(T.accent),align=PP_ALIGN.CENTER,txt_shadow=True)
    txt_quote(slide,req.general_conclusion,4.7,cy+0.88,_CW()-4.0,ch-1.32,
              font=_FONT,size=14.5,color=_rgb(T.text_light),rtl=True)
    return slide

def make_recommendations(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,158)
    _hdr(slide,T,"توصيات البحث","المقترحات العملية")
    recs=req.recommendations[:8]
    if not recs: return slide
    cy=3.28; avail=H-cy-0.4
    ih=min(avail/max(len(recs),1)-0.12,1.52)
    for i,rec in enumerate(recs):
        y=cy+i*(ih+0.12)
        rrect(slide,0.8,y,_CW(),ih,_rgb(T.bg2) if i%2==0 else _rgb(T.card),radius_pct=6)
        tg=rrect(slide,0.8,y,0.46,ih,_rgb(T.accent),radius_pct=6)
        if tg: gradient_fill(tg,T.accent_grad1,T.accent_grad2,90)
        txt_body(slide,rec,1.52,y+0.12,_CW()-2.12,ih-0.24,
                 font=_FONT,size=13.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_future(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,158)
    _hdr(slide,T,"آفاق البحث المستقبلية","مسارات الاستكشاف القادمة")
    items=req.future_work[:6]
    if not items: return slide
    cols=3 if len(items)>=4 else 2 if len(items)>=2 else 1
    rows=(len(items)+cols-1)//cols; gap=0.28
    cy=3.28; avail=H-cy-0.4
    cw=(_CW()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,3.25)
    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=0.8+ci*(cw+gap); y=cy+ri*(ch+gap)
        _card(slide,x,y,cw,ch,T,radius=11)
        bot=rect(slide,x,y+ch-0.42,cw,0.42,_rgb(T.accent))
        if bot: gradient_fill(bot,T.accent_grad1,T.accent_grad2,0)
        txt(slide,f"{i+1:02d}",x+0.18,y+0.14,1.3,0.92,
            font="Calibri",size=25,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.LEFT,
            letter_spacing=-1.0,txt_shadow=True)
        hline(slide,x+0.2,y+1.12,cw-0.4,_rgb(T.muted),thickness=0.05)
        txt_body(slide,item,x+0.22,y+1.26,cw-0.44,ch-1.78,
                 font=_FONT,size=13.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_references(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,158)
    _hdr(slide,T,"المراجع والمصادر","قائمة المراجع المعتمدة")
    refs=req.references[:14]
    if not refs: return slide
    cy=3.28; avail=H-cy-0.36
    ih=max(min(avail/max(len(refs),1)-0.08,1.02),0.48)
    for i,ref in enumerate(refs):
        y=cy+i*(ih+0.08)
        if y+ih>H-0.26: break
        if i%2==0: rrect(slide,0.8,y,_CW(),ih,_rgb(T.bg2),radius_pct=4)
        _pill(slide,T,W-3.1,y+(ih-0.44)/2,1.9,0.44,f"[{i+1}]",9)
        txt_body(slide,ref,1.2,y+0.05,_CW()-2.85,ih-0.1,
                 font=_FONT,size=10.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_final(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,148)
    for ox,oy,os,oa in [(-7,-7,25,5),(W-16,H-14,23,5),(W*0.27,-6.5,19,3),(W*0.7,H*0.4,15,4)]:
        oval(slide,ox,oy,os,os,_rgb(T.accent),alpha=oa)
    tb=rect(slide,0,0,W,0.55,_rgb(T.accent))
    if tb: gradient_fill(tb,T.accent_grad1,T.accent_grad2,0)
    bb=rect(slide,0,H-0.55,W,0.55,_rgb(T.accent))
    if bb: gradient_fill(bb,T.accent_grad2,T.accent_grad1,0)
    cw=28.4; ch=13.65; cx=(W-cw)/2; cy=(H-ch)/2
    mc=rrect(slide,cx,cy,cw,ch,_rgb(T.card),radius_pct=14)
    if mc:
        gradient_fill(mc,T.card,T.bg2,142)
        shadow(mc,blur=36,dist=13,alpha=0.6)
    ct=rrect(slide,cx,cy,cw,0.58,_rgb(T.accent),radius_pct=14)
    if ct: gradient_fill(ct,T.accent_grad1,T.accent_grad2,0)
    txt_hero(slide,"شكراً وتقديراً",cx+1,cy+0.72,cw-2,3.3,
             font=_FONT,size=48,color=_rgb(T.text_light),
             align=PP_ALIGN.CENTER,rtl=True,shadow_on=True)
    for j in range(5):
        dx=cx+cw/2-1.5+j*0.75
        sz=0.44 if j==2 else 0.26
        dc=oval(slide,dx,cy+4.18,sz,sz,_rgb(T.accent))
        if dc:
            if j==2: glow(dc,T.accent,radius_pt=5.5,alpha=0.58)
            else: set_solid_alpha(dc,48)
    txt_hero(slide,req.student_name,cx+1,cy+4.8,cw-2,1.32,
             font=_FONT,size=21,color=_rgb(T.accent),
             align=PP_ALIGN.CENTER,rtl=True,shadow_on=False)
    short=req.title_ar[:88]+("..." if len(req.title_ar)>88 else "")
    txt_body(slide,short,cx+1.8,cy+6.3,cw-3.6,2.8,
             font=_FONT,size=12.5,color=_rgb(T.muted),
             align=PP_ALIGN.CENTER,rtl=True)
    hline(slide,cx+cw*0.15,cy+ch-1.45,cw*0.7,_rgb(T.accent),thickness=0.055)
    footer=[]
    if req.institution: footer.append(req.institution.split(" — ")[0])
    if req.year: footer.append(req.year)
    if footer:
        txt_label(slide," · ".join(footer),cx+1,cy+ch-1.22,cw-2,0.75,
                  font=_FONT,size=11.5,color=_rgb(T.muted),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
    return slide
