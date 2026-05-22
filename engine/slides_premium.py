"""
PREMIUM Engine v19.1 — Corporate Luxury / Glassmorphism  ★ TEXT UPGRADE
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

def _CW(): return W - 2.0

def _glass(slide, T, x, y, w, h, radius=14, alpha=84):
    c = rrect(slide, x, y, w, h, _rgb(T.card), radius_pct=radius)
    if c:
        gradient_fill(c, T.card, T.bg2, 135)
        shadow(c, blur=26, dist=9, alpha=0.55)
        set_solid_alpha(c, alpha)
    brd = rrect(slide, x, y, w, h, _rgb(T.muted), radius_pct=radius)
    if brd: set_solid_alpha(brd, 16)
    return c

def _pill(slide, T, x, y, w, h, text, fsize=10, alpha=100):
    p = rrect(slide, x, y, w, h, _rgb(T.accent), radius_pct=50)
    if p:
        gradient_fill(p, T.accent_grad1, T.accent_grad2, 0)
        if alpha < 100: set_solid_alpha(p, alpha)
        shadow(p, blur=6, dist=2, alpha=0.28)
    txt_label(slide, text, x, y, w, h,
              font=_FONT, size=fsize, color=_rgb(T.text_dark),
              align=PP_ALIGN.CENTER, rtl=True, uppercase=False)

def _hdr(slide, T, title, sub=""):
    top = rect(slide, 0, 0, W, 0.26, _rgb(T.accent))
    if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)
    gradient_rect(slide, 0, 0.26, W, 2.95, T.grad1, T.grad2, 88)
    # زخارف هيدر
    od = oval(slide, W-1.78, 0.45, 0.62, 0.62, _rgb(T.accent))
    if od: glow(od, T.accent, radius_pt=6, alpha=0.5)
    oval(slide, W-4.2, 1.28, 0.32, 0.32, _rgb(T.accent2), alpha=58)
    sz = 28 if len(title)<18 else 22 if len(title)<30 else 17
    txt_hero(slide, title, 1.0, 0.28, W-2.4, 2.02,
             font=_FONT, size=sz, color=_rgb(T.text_light),
             align=PP_ALIGN.RIGHT, rtl=True, shadow_on=True)
    if sub:
        txt_label(slide, sub, 1.0, 2.28, W-2.4, 0.9,
                  font=_FONT, size=11.5, color=_rgb(T.muted),
                  align=PP_ALIGN.RIGHT, rtl=True, uppercase=False)
    hline(slide, 0, 3.06, W, _rgb(T.accent), thickness=0.065)
    oc = oval(slide, W*0.5-0.32, 2.93, 0.64, 0.64, _rgb(T.accent))
    if oc: glow(oc, T.accent, radius_pt=5, alpha=0.48)

# ═════════════════════════════════════════════════════════════════════
def make_cover(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 143)
    # طبقة overlay
    ov = rect(slide, 0, 0, W, H, _rgb(T.bg))
    if ov: set_solid_alpha(ov, 20)
    # دوائر ضخمة خلفية
    oval(slide, W*0.46, -H*0.42, H*1.55, H*1.55, _rgb(T.accent), alpha=6)
    oval(slide, -H*0.32, H*0.33, H*1.18, H*1.18, _rgb(T.bg2), alpha=50)
    oval(slide, W*0.7, H*0.55, H*0.78, H*0.78, _rgb(T.accent), alpha=4)
    # مستطيل ديناميكي
    diag = rect(slide, W*0.55, 0, W*0.52, H, _rgb(T.accent))
    if diag:
        gradient_fill(diag, T.accent_grad1, T.accent_grad2, 90)
        set_solid_alpha(diag, 4)
    top = rect(slide, 0, 0, W, 0.26, _rgb(T.accent))
    if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)

    if req.institution:
        inst = req.institution.split(' — ')[0]
        _pill(slide, T, W-len(inst)*0.19-3.0, 0.5, len(inst)*0.19+2.8, 0.66, inst, 10)

    # العنوان الكبير
    tsz = 33 if len(req.title_ar)<30 else 26 if len(req.title_ar)<52 else 19 if len(req.title_ar)<82 else 14
    txt_hero(slide, req.title_ar, 1.3, H*0.1, W-2.6, H*0.46,
             font=_FONT, size=tsz, color=_rgb(T.text_light),
             align=PP_ALIGN.RIGHT, rtl=True, shadow_on=True)

    if req.title_en:
        txt(slide, req.title_en, 1.3, H*0.56-0.38, W-2.6, 0.9,
            font="Calibri", size=11.5, italic=True,
            color=_rgb(T.muted), align=PP_ALIGN.RIGHT,
            letter_spacing=0.8)

    hline(slide, 1.3, H*0.61, W-2.6, _rgb(T.accent), thickness=0.18)
    # نقطة ذهبية على الخط
    lp = oval(slide, 1.1, H*0.61-0.14, 0.36, 0.36, _rgb(T.accent))
    if lp: glow(lp, T.accent, radius_pt=5, alpha=0.5)

    # بطاقات المعلومات
    fields = [("الطالب", req.student_name)]
    if req.supervisor:     fields.append(("المشرف", req.supervisor))
    if req.co_supervisor:  fields.append(("المشرف المساعد", req.co_supervisor))
    if req.specialization: fields.append(("التخصص", req.specialization))
    n = len(fields)
    cw = min((_CW()-0.28*(n-1))/n, 8.6)
    tot = n*cw+0.28*(n-1); sx = 1.0+(W-2.0-tot)/2
    for i,(lbl,val) in enumerate(fields[:4]):
        px = sx+i*(cw+0.28); py = H*0.64
        _glass(slide, T, px, py, cw, 2.55, radius=13, alpha=82)
        lb = rrect(slide, px, py, cw, 0.55, _rgb(T.accent), radius_pct=13)
        if lb: gradient_fill(lb, T.accent_grad1, T.accent_grad2, 0)
        txt_label(slide, lbl, px+0.1, py+0.04, cw-0.2, 0.49,
                  font=_FONT, size=10, color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER, rtl=True, uppercase=False)
        txt_body(slide, val, px+0.15, py+0.62, cw-0.3, 1.8,
                 font=_FONT, size=12.5, color=_rgb(T.text_light),
                 align=PP_ALIGN.CENTER, rtl=True)

    if req.year:
        yr = rrect(slide, W/2-2.7, H-1.38, 5.4, 0.8, _rgb(T.accent), radius_pct=50)
        if yr:
            gradient_fill(yr, T.accent_grad1, T.accent_grad2, 0)
            shadow(yr, blur=8, dist=2, alpha=0.32)
            glow(yr, T.accent, radius_pt=5, alpha=0.22)
        txt_label(slide, req.year, W/2-2.7, H-1.38, 5.4, 0.8,
                  font="Calibri", size=13, color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER, rtl=False, uppercase=False)
    return slide

def make_intro(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"مقدمة البحث","الإطار العام والمقاربة المنهجية")
    cy=3.15; ch=H-cy-0.42
    if req.intro_overview and req.intro_approach:
        cw1=_CW()*0.56; cw2=_CW()-cw1-0.38
        for titl,text,cw,cx in [
            ("نظرة عامة",req.intro_overview,cw1,1.0),
            ("المقاربة",req.intro_approach,cw2,1.0+cw1+0.38)]:
            _glass(slide,T,cx,cy,cw,ch,radius=13)
            tp=rect(slide,cx,cy,cw,0.65,_rgb(T.accent))
            if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
            vline(slide,cx,cy,ch,_rgb(T.accent),thickness=0.2)
            txt_label(slide,titl,cx+0.38,cy+0.07,cw-0.58,0.55,
                      font=_FONT,size=14.5,color=_rgb(T.text_dark),
                      align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
            txt_body(slide,text,cx+0.38,cy+0.78,cw-0.68,ch-0.98,
                     font=_FONT,size=13.5,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
    else:
        text=req.intro_overview or req.intro_approach
        _glass(slide,T,1.0,cy,_CW(),ch,radius=13)
        lb=rrect(slide,1.0,cy,0.52,ch,_rgb(T.accent),radius_pct=13)
        if lb: gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
        txt(slide,"❝",1.6,cy+0.28,2.2,2.0,
            font="Calibri",size=58,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.LEFT,txt_shadow=True)
        txt_quote(slide,text,1.68,cy+0.72,_CW()-1.1,ch-0.98,
                  font=_FONT,size=14.5,color=_rgb(T.text_light),rtl=True)
    return slide

def make_plan(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"خطة البحث","هيكل ومحتويات الدراسة")
    chapters=req.chapters[:8]
    if not chapters: return slide
    cy=3.15; avail=H-cy-0.42
    n=len(chapters)
    if n<=4: cols=n; rows=1
    elif n<=6: cols=3; rows=2
    else: cols=4; rows=(n+3)//4
    gap=0.28
    cw=(_CW()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,3.88)
    for i,chap in enumerate(chapters):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(cw+gap); y=cy+ri*(ch+gap)
        _glass(slide,T,x,y,cw,ch,radius=13,alpha=88)
        tp=rect(slide,x,y,cw,0.58,_rgb(T.accent))
        if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
        txt(slide,f"{i+1:02d}",x+0.14,y+0.05,2.0,0.5,
            font="Calibri",size=18,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.LEFT,
            letter_spacing=-1.0)
        hline(slide,x+0.2,y+0.72,cw-0.4,_rgb(T.muted),thickness=0.04)
        txt_body(slide,chap.title,x+0.2,y+0.85,cw-0.4,ch-1.24,
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
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"إشكالية البحث","التساؤلات المحورية للدراسة")
    cy=3.15
    if req.main_problem or req.main_question:
        text=req.main_problem or req.main_question
        qh=min(3.05,H*0.31)
        _glass(slide,T,1.0,cy,_CW(),qh,radius=13)
        hline(slide,1.0,cy+qh-0.15,_CW(),_rgb(T.accent),thickness=0.15)
        txt(slide,"❝",1.42,cy+0.18,2.55,qh-0.36,
            font="Calibri",size=55,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.RIGHT,txt_shadow=True)
        txt_quote(slide,"  "+text,1.45,cy+0.26,_CW()-0.82,qh-0.52,
                  font=_FONT,size=14,color=_rgb(T.text_light),rtl=True)
        cy+=qh+0.26
    if req.sub_questions:
        sq=req.sub_questions[:6]; n=len(sq)
        cols=2 if n>3 else 1
        cw=(_CW()-(cols-1)*0.28)/cols
        avail=H-cy-0.4
        sh=min(avail/((n+cols-1)//cols)-0.1,1.2)
        for i,q in enumerate(sq):
            ci=i%cols; ri=i//cols
            x=1.0+ci*(cw+0.28); y=cy+ri*(sh+0.1)
            _glass(slide,T,x,y,cw,sh,radius=9,alpha=80)
            _pill(slide,T,x+cw-1.72,y+(sh-0.48)/2,1.58,0.48,str(i+1),10)
            vline(slide,x,y,sh,_rgb(T.accent),thickness=0.36)
            txt_body(slide,q,x+0.55,y+0.1,cw-2.38,sh-0.2,
                     font=_FONT,size=13.5,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_objectives(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"أهداف البحث وفرضياته","")
    cy=3.15; ch=H-cy-0.42; gap=0.4
    cols=[]
    if req.objectives: cols.append(("الأهداف",req.objectives))
    if req.hypotheses:  cols.append(("الفرضيات",req.hypotheses))
    if not cols: return slide
    widths=[_CW()*0.56,_CW()*0.42] if len(cols)==2 else [_CW()]
    x=1.0
    for i,(lbl,items) in enumerate(cols[:2]):
        cw=widths[i]
        _glass(slide,T,x,cy,cw,ch,radius=13)
        hd=rect(slide,x,cy,cw,0.66,_rgb(T.accent))
        if hd: gradient_fill(hd,T.accent_grad1,T.accent_grad2,0)
        txt_label(slide,lbl,x+0.22,cy+0.07,cw-0.44,0.56,
                  font=_FONT,size=14.5,color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        n_items=len(items)
        # ADAPTIVE: fewer items → larger rows and font
        ih=min((ch-0.75)/max(n_items,1), 1.18 if n_items>=4 else 2.2)
        body_size=13 if n_items>=4 else 15
        for j,item in enumerate(items[:8]):
            iy=cy+0.75+j*ih
            if iy+ih>cy+ch-0.08: break
            if j>0: hline(slide,x+0.22,iy,cw-0.44,_rgb(T.bg),thickness=0.03)
            _pill(slide,T,x+cw-1.48,iy+(ih-0.47)/2,1.32,0.47,str(j+1),10)
            txt_body(slide,item,x+0.22,iy+0.09,cw-1.9,ih-0.18,
                     font=_FONT,size=body_size,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
        x+=cw+gap
    return slide

def make_importance(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"أهمية البحث","مبررات اختيار الموضوع")
    items=list(req.importance or [])
    if req.reasons and req.reasons not in items: items.append(req.reasons)
    items=items[:6]
    if not items: return slide
    n=len(items)
    ah_list=[0.65,0.82,0.56,0.74,0.62,0.78]
    # ADAPTIVE layout
    if n <= 2:
        cols=n; body_size=14.5; num_size=28
    elif n == 3:
        cols=3; body_size=13; num_size=24
    else:
        cols=3; body_size=11.5; num_size=21
    rows=(n+cols-1)//cols; gap=0.28
    cy=3.15; avail=H-cy-0.42
    cw=(_CW()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap, 3.72 if n>=4 else avail-gap*0.5)
    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(cw+gap); y=cy+ri*(ch+gap)
        _glass(slide,T,x,y,cw,ch,radius=11,alpha=87)
        ah=ah_list[i%len(ah_list)]
        tp=rect(slide,x,y,cw,ah,_rgb(T.accent))
        if tp:
            gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
            set_solid_alpha(tp,90)
        txt(slide,f"{i+1:02d}",x+0.16,y+0.04,2.2,ah-0.08,
            font="Calibri",size=num_size,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.LEFT,
            letter_spacing=-1.0)
        txt_body(slide,item,x+0.22,y+ah+0.18,cw-0.44,ch-ah-0.3,
                 font=_FONT,size=body_size,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_methodology(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"منهجية البحث","الإجراءات والأدوات المستخدمة")
    fields=[]
    if req.methodology: fields.append(("المنهج","م",req.methodology))
    if req.sample_type: fields.append(("العينة","ع",req.sample_type))
    if req.sample_size: fields.append(("الحجم","ن",req.sample_size))
    if req.tool:        fields.append(("الأداة","أ",req.tool))
    if not fields: return slide
    n=len(fields); cy=3.15; ch=H-cy-0.46
    cw=(_CW()-0.28*(n-1))/n
    if n>1: hline(slide,1.0+cw/2,cy+ch/2,_CW()-cw,_rgb(T.muted),thickness=0.04)
    for i,(lbl,icon,val) in enumerate(fields[:4]):
        x=1.0+i*(cw+0.28)
        _glass(slide,T,x,cy,cw,ch,radius=14,alpha=87)
        sz=1.78; ix=x+(cw-sz)/2
        ic=oval(slide,ix,cy+0.52,sz,sz,_rgb(T.accent))
        if ic:
            gradient_fill(ic,T.accent_grad1,T.accent_grad2,45)
            shadow_and_glow(ic,s_blur=14,s_dist=4,s_alpha=0.45,
                            g_color=T.accent_grad2,g_rad=7,g_alpha=0.28)
        txt(slide,icon,ix,cy+0.52,sz,sz,
            font=_FONT,size=26,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.CENTER)
        hline(slide,x+0.28,cy+2.52,cw-0.56,_rgb(T.accent),thickness=0.08)
        txt_label(slide,lbl,x+0.14,cy+2.64,cw-0.28,0.68,
                  font=_FONT,size=13,color=_rgb(T.accent),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        txt_body(slide,val,x+0.22,cy+3.42,cw-0.44,ch-3.58,
                 font=_FONT,size=11.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.CENTER,rtl=True)
    return slide

def make_stats(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"الإحصاءات والأرقام الرئيسية","مؤشرات كمية للدراسة")
    stats=req.stats[:6]
    if not stats: return slide
    n=len(stats); cols=3 if n>=3 else n
    rows=(n+cols-1)//cols; gap=0.3
    cy=3.15; avail=H-cy-0.42
    cw=(_CW()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,4.65)
    for i,st in enumerate(stats):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(cw+gap); y=cy+ri*(ch+gap)
        _glass(slide,T,x,y,cw,ch,radius=14,alpha=90)
        cs=min(cw,ch)*0.78
        od=oval(slide,x+(cw-cs)/2,y+(ch-cs)/2-0.32,cs,cs,_rgb(T.accent),alpha=6)
        txt_stat(slide,st.value,x+0.2,y+0.44,cw-0.4,ch*0.5,
                 font="Calibri",color=_rgb(T.accent),align=PP_ALIGN.CENTER)
        if st.unit:
            txt_label(slide,st.unit,x+0.2,y+ch*0.5+0.28,cw-0.4,0.58,
                      font=_FONT,size=11,color=_rgb(T.muted),
                      align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        lw=min(len(st.label)*0.2+1.1,cw-0.36)
        _pill(slide,T,x+(cw-lw)/2,y+ch-1.14,lw,0.6,st.label,10,alpha=88)
    return slide

def make_results(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"نتائج البحث","أبرز ما توصلت إليه الدراسة")
    results=req.main_results[:8]
    if not results: return slide
    n=len(results)
    cy=3.55; avail=H-cy-0.4
    # ADAPTIVE sizing
    if n <= 3:
        ih=min(avail/n - 0.2, 3.6); fsize=16; pill_h=0.68; pill_w=2.4; pill_fsize=13; vt=0.5
    elif n <= 5:
        ih=min(avail/n - 0.14, 2.4); fsize=14.5; pill_h=0.6; pill_w=2.2; pill_fsize=12; vt=0.42
    else:
        ih=min(avail/n - 0.12, 1.62); fsize=13.5; pill_h=0.54; pill_w=2.08; pill_fsize=12; vt=0.36
    for i,res in enumerate(results):
        y=cy+i*(ih+(0.2 if n<=3 else 0.12))
        row=rrect(slide,1.0,y,_CW(),ih,_rgb(T.bg2) if i%2==0 else _rgb(T.card),radius_pct=7)
        if row: shadow(row,blur=7,dist=2,alpha=0.2)
        _pill(slide,T,W-3.25,y+(ih-pill_h)/2,pill_w,pill_h,str(i+1),pill_fsize)
        vline(slide,1.0,y,ih,_rgb(T.accent),thickness=vt)
        txt_body(slide,res,1.54,y+0.12,W-5.4,ih-0.24,
                 font=_FONT,size=fsize,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_conclusion(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"خاتمة البحث","الاستنتاج العام")
    cy=3.15; ch=H-cy-0.52
    _glass(slide,T,1.0,cy,_CW(),ch,radius=15,alpha=90)
    lb=rrect(slide,1.0,cy,3.3,ch,_rgb(T.accent),radius_pct=15)
    if lb:
        gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
        set_solid_alpha(lb,18)
    tp=rrect(slide,1.0,cy,_CW(),0.46,_rgb(T.accent),radius_pct=15)
    if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
    txt(slide,"❝",1.42,cy+0.58,3.1,2.75,
        font="Calibri",size=66,bold=True,
        color=_rgb(T.accent),align=PP_ALIGN.CENTER,txt_shadow=True)
    ctext = req.general_conclusion or ""
    conc_size = 16 if len(ctext) < 120 else 14.5 if len(ctext) < 220 else 13
    txt_quote(slide,ctext,4.72,cy+0.9,_CW()-3.98,ch-1.32,
              font=_FONT,size=conc_size,color=_rgb(T.text_light),rtl=True)
    return slide

def make_recommendations(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"توصيات البحث","المقترحات العملية")
    recs=req.recommendations[:8]
    if not recs: return slide
    n=len(recs)
    cy=3.55; avail=H-cy-0.4
    # ADAPTIVE sizing
    if n <= 3:
        ih=min(avail/n - 0.18, 3.6); fsize=16; tag_w=0.6
    elif n <= 5:
        ih=min(avail/n - 0.14, 2.4); fsize=14.5; tag_w=0.52
    else:
        ih=min(avail/n - 0.12, 1.54); fsize=13.5; tag_w=0.46
    for i,rec in enumerate(recs):
        y=cy+i*(ih+(0.18 if n<=3 else 0.12))
        row=rrect(slide,1.0,y,_CW(),ih,_rgb(T.bg2) if i%2==0 else _rgb(T.card),radius_pct=7)
        tg=rrect(slide,1.0,y,tag_w,ih,_rgb(T.accent),radius_pct=7)
        if tg: gradient_fill(tg,T.accent_grad1,T.accent_grad2,90)
        txt_body(slide,rec,1.0+tag_w+0.18,y+0.12,_CW()-tag_w-0.25,ih-0.24,
                 font=_FONT,size=fsize,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_future(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"آفاق البحث المستقبلية","مسارات الاستكشاف القادمة")
    items=req.future_work[:6]
    if not items: return slide
    n=len(items)
    cols=3 if n>=4 else 2 if n>=2 else 1
    rows=(n+cols-1)//cols; gap=0.28
    cy=3.15; avail=H-cy-0.42
    cw=(_CW()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap, 3.25 if n>=4 else avail-gap*0.5)
    num_size=32 if n<=2 else 25
    body_size=15 if n<=2 else 13.5
    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(cw+gap); y=cy+ri*(ch+gap)
        _glass(slide,T,x,y,cw,ch,radius=13,alpha=88)
        bot=rect(slide,x,y+ch-0.42,cw,0.42,_rgb(T.accent))
        if bot: gradient_fill(bot,T.accent_grad1,T.accent_grad2,0)
        txt(slide,f"{i+1:02d}",x+0.18,y+0.14,2.2,0.92,
            font="Calibri",size=num_size,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.LEFT,
            letter_spacing=-1.0,txt_shadow=True)
        hline(slide,x+0.2,y+1.12,cw-0.4,_rgb(T.muted),thickness=0.05)
        txt_body(slide,item,x+0.22,y+1.26,cw-0.44,ch-1.78,
                 font=_FONT,size=body_size,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_references(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"المراجع والمصادر","قائمة المراجع المعتمدة")
    refs=req.references[:14]
    if not refs: return slide
    n=len(refs)
    cy=3.55; avail=H-cy-0.4
    # ADAPTIVE sizing
    if n <= 4:
        ih=min(avail/n - 0.1, 2.8); body_size=13
    elif n <= 8:
        ih=min(avail/n - 0.08, 1.6); body_size=11.5
    else:
        ih=max(min(avail/n - 0.08, 1.02), 0.48); body_size=10.5
    for i,ref in enumerate(refs):
        y=cy+i*(ih+0.08)
        if y+ih>H-0.26: break
        if i%2==0: rrect(slide,1.0,y,_CW(),ih,_rgb(T.bg2),radius_pct=5)
        _pill(slide,T,W-3.2,y+(ih-0.44)/2,1.95,0.44,f"[{i+1}]",9)
        txt_body(slide,ref,1.4,y+0.05,_CW()-2.88,ih-0.1,
                 font=_FONT,size=body_size,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_final(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,148)
    for ox,oy,os,oa in [(-7.5,-7.5,27,5),(W-17,H-15,25,5),(W*0.24,-7,21,3)]:
        oval(slide,ox,oy,os,os,_rgb(T.accent),alpha=oa)
    tb=rect(slide,0,0,W,0.55,_rgb(T.accent))
    if tb: gradient_fill(tb,T.accent_grad1,T.accent_grad2,0)
    bb=rect(slide,0,H-0.55,W,0.55,_rgb(T.accent))
    if bb: gradient_fill(bb,T.accent_grad2,T.accent_grad1,0)
    cw=28.5; ch=14.1; cx=(W-cw)/2; cy=(H-ch)/2
    _glass(slide,T,cx,cy,cw,ch,radius=16,alpha=92)
    ct=rrect(slide,cx,cy,cw,0.58,_rgb(T.accent),radius_pct=16)
    if ct: gradient_fill(ct,T.accent_grad1,T.accent_grad2,0)
    txt_hero(slide,"شكراً وتقديراً",cx+1,cy+0.72,cw-2,3.35,
             font=_FONT,size=48,color=_rgb(T.text_light),
             align=PP_ALIGN.CENTER,rtl=True,shadow_on=True)
    for j in range(5):
        dx=cx+cw/2-1.5+j*0.75
        sz=0.46 if j==2 else 0.27
        dc=oval(slide,dx,cy+4.22,sz,sz,_rgb(T.accent))
        if dc:
            if j==2: glow(dc,T.accent,radius_pt=6,alpha=0.6)
            else: set_solid_alpha(dc,48)
    txt_hero(slide,req.student_name,cx+1,cy+4.88,cw-2,1.34,
             font=_FONT,size=22,color=_rgb(T.accent),
             align=PP_ALIGN.CENTER,rtl=True,shadow_on=False)
    short=req.title_ar[:88]+("..." if len(req.title_ar)>88 else "")
    txt_body(slide,short,cx+1.8,cy+6.42,cw-3.6,2.85,
             font=_FONT,size=12.5,color=_rgb(T.muted),
             align=PP_ALIGN.CENTER,rtl=True)
    hline(slide,cx+cw*0.15,cy+ch-1.48,cw*0.7,_rgb(T.accent),thickness=0.06)
    footer=[]
    if req.institution: footer.append(req.institution.split(" — ")[0])
    if req.year: footer.append(req.year)
    if footer:
        txt_label(slide," · ".join(footer),cx+1,cy+ch-1.25,cw-2,0.78,
                  font=_FONT,size=11.5,color=_rgb(T.muted),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
    return slide
