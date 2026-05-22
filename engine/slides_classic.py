"""
CLASSIC Engine v20.0 — Academic Precision  ★ ADAPTIVE FILL
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
    h = h.lstrip("#")
    return RGBColor(int(h[:2],16), int(h[2:4],16), int(h[4:],16))

_SW  = 5.8      # side panel width
_CX  = _SW + 0.85
_CW  = W - _CX - 0.65
_HH  = 2.95     # header height
_BY  = _HH + 0.35

# ── Side Panel ──────────────────────────────────────────────────────
def _side(slide, T, num=""):
    sp = rect(slide, 0, 0, _SW, H, _rgb(T.bg2))
    if sp: gradient_fill(sp, T.bg, T.bg2, 175)
    vline(slide, _SW - 0.07, 0, H, _rgb(T.accent), thickness=0.07)
    bot = rect(slide, 0, H - 1.15, _SW, 1.15, _rgb(T.accent))
    if bot:
        gradient_fill(bot, T.accent_grad1, T.accent_grad2, 175)
        shadow(bot, blur=12, dist=0, alpha=0.35)
    if num:
        txt(slide, num, 0, H - 1.12, _SW, 1.08,
            font="Calibri", size=30, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.CENTER,
            letter_spacing=0, txt_shadow=False)
    # نقطة زخرفية علوية
    od = oval(slide, _SW*0.5-0.25, 0.55, 0.5, 0.5, _rgb(T.accent))
    if od: glow(od, T.accent, radius_pt=5, alpha=0.45)

# ── Header ───────────────────────────────────────────────────────────
def _hdr(slide, T, title, sub="", num=""):
    hb = rect(slide, _SW, 0, W-_SW, _HH, _rgb(T.bg))
    if hb: gradient_fill(hb, T.bg, T.bg2, 85)
    accent_bar = rect(slide, _SW, _HH-0.14, W-_SW, 0.14, _rgb(T.accent))
    if accent_bar: gradient_fill(accent_bar, T.accent_grad1, T.accent_grad2, 0)
    _side(slide, T, num)
    sz = 26 if len(title)<20 else 21 if len(title)<32 else 16
    txt_hero(slide, title, _CX, 0.38, _CW, 1.82,
             font=_FONT, size=sz, color=_rgb(T.text_light),
             align=PP_ALIGN.RIGHT, rtl=True, shadow_on=True)
    if sub:
        hline(slide, _CX, 2.18, _CW*0.32, _rgb(T.accent), thickness=0.07)
        txt_label(slide, sub, _CX, 2.22, _CW, 0.82,
                  font=_FONT, size=11, color=_rgb(T.muted),
                  align=PP_ALIGN.RIGHT, rtl=True, uppercase=False)

def _row(slide, T, y, h, text, num="", alt=False, fsize=12.5):
    if alt:
        rb = rect(slide, _CX, y, _CW, h, _rgb(T.bg2))
        if rb: set_solid_alpha(rb, 55)
    if num:
        nb = rect(slide, _CX, y, 0.38, h, _rgb(T.accent))
        if nb: gradient_fill(nb, T.accent_grad1, T.accent_grad2, 90)
        tx, tw = _CX+0.55, _CW-0.7
    else:
        tx, tw = _CX+0.22, _CW-0.32
    txt_body(slide, text, tx, y, tw, h,
             font=_FONT, size=fsize, color=_rgb(T.text_light),
             align=PP_ALIGN.RIGHT, rtl=True, valign_center=True)

def _card_info(slide, T, x, y, w, h, label, value):
    c = rrect(slide, x, y, w, h, _rgb(T.card), radius_pct=7)
    if c: shadow(c, blur=12, dist=4, alpha=0.38)
    top = rect(slide, x, y, w, 0.4, _rgb(T.accent))
    if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)
    txt_label(slide, label, x+0.1, y+0.03, w-0.2, 0.36,
              font=_FONT, size=9, color=_rgb(T.text_dark),
              align=PP_ALIGN.CENTER, rtl=True, uppercase=False)
    txt_body(slide, value, x+0.1, y+0.46, w-0.2, h-0.58,
             font=_FONT, size=11.5, color=_rgb(T.text_light),
             align=PP_ALIGN.CENTER, rtl=True)

# ═════════════════════════════════════════════════════════════════════
def make_cover(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    _side(slide, T)

    # معلومات المؤسسة على الشريط الجانبي
    if req.institution:
        parts = req.institution.split(" — ")
        for i, p in enumerate(parts[:3]):
            txt_body(slide, p, 0.18, 1.2+i*1.05, _SW-0.36, 0.9,
                     font=_FONT, size=10, color=_rgb(T.muted),
                     align=PP_ALIGN.CENTER, rtl=True)

    # المنطقة الرئيسية
    cx = _SW+0.8; cw = W-cx-0.65
    # بطاقة العنوان
    cy = 0.55; ch = H*0.48
    mc = rrect(slide, cx, cy, cw, ch, _rgb(T.card), radius_pct=10)
    if mc:
        gradient_fill(mc, T.card, T.bg2, 135)
        shadow(mc, blur=26, dist=9, alpha=0.55)
    ct = rect(slide, cx, cy, cw, 0.58, _rgb(T.accent))
    if ct: gradient_fill(ct, T.accent_grad1, T.accent_grad2, 0)

    tsz = 24 if len(req.title_ar)<40 else 18 if len(req.title_ar)<68 else 14
    txt_hero(slide, req.title_ar, cx+0.5, cy+0.72, cw-1.0, ch-1.3,
             font=_FONT, size=tsz, color=_rgb(T.text_light),
             align=PP_ALIGN.RIGHT, rtl=True, shadow_on=True)

    if req.title_en:
        txt(slide, req.title_en, cx+0.5, cy+ch-0.88, cw-1.0, 0.78,
            font="Calibri", size=10, italic=True,
            color=_rgb(T.muted), align=PP_ALIGN.CENTER, letter_spacing=0.5)

    # بطاقات المعلومات
    fields = [("الطالب", req.student_name)]
    if req.supervisor:     fields.append(("المشرف", req.supervisor))
    if req.co_supervisor:  fields.append(("م. مساعد", req.co_supervisor))
    if req.specialization: fields.append(("التخصص", req.specialization))

    n = len(fields)
    iy = cy+ch+0.35
    ih = min((H - iy - 1.05) * 0.92, 2.1)
    iw = (cw - 0.25*(n-1)) / n
    for i, (lbl, val) in enumerate(fields[:4]):
        ix = cx + i*(iw+0.25)
        _card_info(slide, T, ix, iy, iw, ih, lbl, val)

    if req.year:
        yr = rrect(slide, cx + cw/2 - 2.5, H-1.28, 5.0, 0.72, _rgb(T.accent), radius_pct=50)
        if yr:
            gradient_fill(yr, T.accent_grad1, T.accent_grad2, 0)
            shadow(yr, blur=6, dist=2, alpha=0.28)
        txt_label(slide, req.year, cx+cw/2-2.5, H-1.28, 5.0, 0.72,
                  font="Calibri", size=12, color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER, rtl=False, uppercase=False)
    return slide

def make_intro(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"مقدمة البحث","نظرة عامة على الدراسة","01")
    cy=_BY+0.08; ch=H-cy-0.42
    if req.intro_overview and req.intro_approach:
        cw1=_CW*0.56; cw2=_CW-cw1-0.35
        for titl,text,cw,cx in [
            ("الإطار العام",req.intro_overview,cw1,_CX),
            ("المقاربة",req.intro_approach,cw2,_CX+cw1+0.35)]:
            c=rrect(slide,cx,cy,cw,ch,_rgb(T.card),radius_pct=9)
            if c:
                gradient_fill(c,T.card,T.bg2,135)
                shadow(c,blur=18,dist=6,alpha=0.45)
            tp=rect(slide,cx,cy,cw,0.58,_rgb(T.accent))
            if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
            vline(slide,cx,cy,ch,_rgb(T.accent),thickness=0.18)
            txt_label(slide,titl,cx+0.3,cy+0.06,cw-0.5,0.5,
                      font=_FONT,size=13,color=_rgb(T.text_dark),
                      align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
            txt_body(slide,text,cx+0.38,cy+0.7,cw-0.62,ch-0.88,
                     font=_FONT,size=13,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
    else:
        text=req.intro_overview or req.intro_approach or ""
        c=rrect(slide,_CX,cy,_CW,ch,_rgb(T.card),radius_pct=9)
        if c:
            gradient_fill(c,T.card,T.bg2,135)
            shadow(c,blur=18,dist=6,alpha=0.45)
        lb=rect(slide,_CX,cy,0.52,ch,_rgb(T.accent))
        if lb: gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
        txt(slide,"❝",_CX+0.68,cy+0.28,2.1,1.9,
            font="Calibri",size=52,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.LEFT,txt_shadow=True)
        txt_quote(slide,text,_CX+0.8,cy+0.62,_CW-1.1,ch-0.9,
                  font=_FONT,size=14.5,color=_rgb(T.text_light),rtl=True)
    return slide

def make_plan(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"خطة البحث","هيكل ومحتويات الدراسة","02")
    chapters=req.chapters[:8]
    if not chapters: return slide
    n=len(chapters)
    cy=_BY+0.08; avail=H-cy-0.38
    if n<=4: cols=n; rows=1
    elif n<=6: cols=3; rows=2
    else: cols=4; rows=(n+3)//4
    gap=0.28
    cw=(_CW-(cols-1)*gap)/cols
    ch=avail/rows-gap
    for i,chap in enumerate(chapters):
        ci=i%cols; ri=i//cols
        x=_CX+ci*(cw+gap); y=cy+ri*(ch+gap)
        c=rrect(slide,x,y,cw,ch,_rgb(T.card),radius_pct=9)
        if c: shadow(c,blur=12,dist=4,alpha=0.32)
        tp=rect(slide,x,y,cw,0.55,_rgb(T.accent))
        if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
        txt(slide,f"{i+1:02d}",x+0.12,y+0.05,1.8,0.47,
            font="Calibri",size=17,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.LEFT)
        hline(slide,x+0.18,y+0.68,cw-0.36,_rgb(T.muted),thickness=0.04)
        txt_body(slide,chap.title,x+0.18,y+0.8,cw-0.36,ch-1.15,
                 font=_FONT,size=14,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
        if chap.pages:
            txt_label(slide,f"ص {chap.pages}",x+0.16,y+ch-0.5,cw-0.32,0.4,
                      font="Calibri",size=10.5,color=_rgb(T.muted),
                      align=PP_ALIGN.LEFT,rtl=False,uppercase=False)
    return slide

def make_problem(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"إشكالية البحث","التساؤلات المحورية للدراسة","03")
    cy=_BY+0.08
    if req.main_problem or req.main_question:
        text=req.main_problem or req.main_question
        qh=min(H*0.3,2.9)
        qc=rrect(slide,_CX,cy,_CW,qh,_rgb(T.card),radius_pct=9)
        if qc: shadow(qc,blur=18,dist=6,alpha=0.45)
        hline(slide,_CX,cy+qh-0.12,_CW,_rgb(T.accent),thickness=0.12)
        txt(slide,"❝",_CX+0.15,cy+0.12,2.2,qh-0.24,
            font="Calibri",size=52,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.RIGHT,txt_shadow=True)
        txt_quote(slide,text,_CX+0.35,cy+0.22,_CW-0.65,qh-0.44,
                  font=_FONT,size=13.5,color=_rgb(T.text_light),rtl=True)
        cy+=qh+0.25
    if req.sub_questions:
        sq=req.sub_questions[:6]; n=len(sq)
        cols=2 if n>2 else 1
        cw=(_CW-(cols-1)*0.25)/cols
        avail=H-cy-0.35
        sh=avail/((n+cols-1)//cols) - 0.1
        for i,q in enumerate(sq):
            ci=i%cols; ri=i//cols
            x=_CX+ci*(cw+0.25); y=cy+ri*(sh+0.1)
            rc=rrect(slide,x,y,cw,sh,_rgb(T.card),radius_pct=7)
            if rc: shadow(rc,blur=7,dist=2,alpha=0.2)
            lb=rect(slide,x,y,0.38,sh,_rgb(T.accent))
            if lb: gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
            nb=oval(slide,x+cw-1.3,y+(sh-0.42)/2,0.42,0.42,_rgb(T.accent))
            if nb: set_solid_alpha(nb,88)
            txt(slide,str(i+1),x+cw-1.3,y+(sh-0.42)/2,0.42,0.42,
                font="Calibri",size=9,bold=True,
                color=_rgb(T.text_dark),align=PP_ALIGN.CENTER)
            txt_body(slide,q,x+0.55,y+0.09,cw-2.0,sh-0.18,
                     font=_FONT,size=13,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_objectives(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"أهداف البحث وفرضياته","","04")
    cy=_BY+0.08; cha=H-cy-0.42; gap=0.35
    cols=[]
    if req.objectives: cols.append(("الأهداف",req.objectives))
    if req.hypotheses:  cols.append(("الفرضيات",req.hypotheses))
    if not cols: return slide
    widths=[_CW*0.56,_CW*0.42] if len(cols)==2 else [_CW]
    x=_CX
    for i,(lbl,items) in enumerate(cols[:2]):
        cw=widths[i]
        c=rrect(slide,x,cy,cw,cha,_rgb(T.card),radius_pct=9)
        if c: shadow(c,blur=14,dist=5,alpha=0.38)
        hd=rect(slide,x,cy,cw,0.58,_rgb(T.accent))
        if hd: gradient_fill(hd,T.accent_grad1,T.accent_grad2,0)
        txt_label(slide,lbl,x+0.18,cy+0.05,cw-0.36,0.5,
                  font=_FONT,size=14,color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        n_items=len(items)
        # FILL the full height — no artificial cap
        ih=(cha-0.65)/max(n_items,1)
        body_size=15 if n_items<=3 else 13.5 if n_items<=5 else 13
        for j,item in enumerate(items[:8]):
            iy=cy+0.65+j*ih
            if iy+ih>cy+cha-0.04: break
            if j>0: hline(slide,x+0.22,iy,cw-0.44,_rgb(T.bg),thickness=0.03)
            nb=oval(slide,x+cw-1.35,iy+(ih-0.44)/2,0.44,0.44,_rgb(T.accent))
            if nb: set_solid_alpha(nb,88)
            txt(slide,str(j+1),x+cw-1.35,iy+(ih-0.44)/2,0.44,0.44,
                font="Calibri",size=9.5,bold=True,
                color=_rgb(T.text_dark),align=PP_ALIGN.CENTER)
            txt_body(slide,item,x+0.22,iy+0.09,cw-1.85,ih-0.18,
                     font=_FONT,size=body_size,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
        x+=cw+gap
    return slide

def make_importance(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"أهمية البحث ومبرراته","","05")
    items=list(req.importance or [])
    if req.reasons and req.reasons not in items: items.append(req.reasons)
    items=items[:6]
    if not items: return slide
    n=len(items)
    cy=_BY+0.08; avH=H-cy-0.38
    # ADAPTIVE columns
    cols=3 if n>=4 else (n if n<=2 else 3)
    rows=(n+cols-1)//cols
    gap=0.28
    cw=(_CW-(cols-1)*gap)/cols
    # Fill full available height
    ch=avH/rows - gap*(rows-1)/rows
    num_size=26 if n<=2 else 22 if n==3 else 18
    body_size=14 if n<=2 else 12.5 if n==3 else 11
    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=_CX+ci*(cw+gap); y=cy+ri*(ch+gap)
        c=rrect(slide,x,y,cw,ch,_rgb(T.card),radius_pct=8)
        if c: shadow(c,blur=12,dist=4,alpha=0.32)
        top_h=0.68
        top=rect(slide,x,y,cw,top_h,_rgb(T.accent))
        if top: gradient_fill(top,T.accent_grad1,T.accent_grad2,0)
        txt(slide,f"{i+1:02d}",x+0.14,y+0.04,2.2,top_h-0.06,
            font="Calibri",size=num_size,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.LEFT)
        # موضع النص في مركز المنطقة تحت الشريط العلوي
        text_area_y = y + top_h + 0.1
        text_area_h = ch - top_h - 0.2
        txt_body(slide,item,x+0.2,text_area_y,cw-0.4,text_area_h,
                 font=_FONT,size=body_size,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_methodology(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"منهجية البحث","الإجراءات والأدوات المستخدمة","06")
    fields=[]
    if req.methodology: fields.append(("المنهج المتبع","م",req.methodology))
    if req.sample_type: fields.append(("نوع العينة","ع",req.sample_type))
    if req.sample_size: fields.append(("حجم العينة","ن",req.sample_size))
    if req.tool:        fields.append(("أداة الجمع","أ",req.tool))
    if not fields: return slide
    n=len(fields); cy=_BY+0.08; ch=H-cy-0.48
    cw=(_CW-0.28*(n-1))/n
    if n>1: hline(slide,_CX+cw/2,cy+ch/2,_CW-cw,_rgb(T.muted),thickness=0.04)
    for i,(lbl,icon,val) in enumerate(fields[:4]):
        x=_CX+i*(cw+0.28)
        c=rrect(slide,x,cy,cw,ch,_rgb(T.card),radius_pct=11)
        if c: shadow(c,blur=16,dist=5,alpha=0.42)
        sz=1.75; ix=x+(cw-sz)/2
        ic=oval(slide,ix,cy+0.52,sz,sz,_rgb(T.accent))
        if ic:
            gradient_fill(ic,T.accent_grad1,T.accent_grad2,45)
            shadow_and_glow(ic,s_blur=10,s_dist=3,s_alpha=0.38,
                            g_color=T.accent_grad2,g_rad=5,g_alpha=0.22)
        txt(slide,icon,ix,cy+0.52,sz,sz,
            font=_FONT,size=26,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.CENTER)
        hline(slide,x+0.28,cy+2.48,cw-0.56,_rgb(T.accent),thickness=0.08)
        txt_label(slide,lbl,x+0.12,cy+2.58,cw-0.24,0.68,
                  font=_FONT,size=12,color=_rgb(T.accent),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        txt_body(slide,val,x+0.2,cy+3.35,cw-0.4,ch-3.52,
                 font=_FONT,size=11.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.CENTER,rtl=True)
    return slide

def make_stats(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"الإحصاءات والأرقام الرئيسية","مؤشرات كمية للدراسة","07")
    stats=req.stats[:6]
    if not stats: return slide
    n=len(stats); cols=3 if n>=3 else n
    rows=(n+cols-1)//cols; gap=0.3
    cy=_BY+0.08; avail=H-cy-0.38
    cw=(_CW-(cols-1)*gap)/cols
    ch=avail/rows - gap*(rows-1)/rows
    for i,st in enumerate(stats):
        ci=i%cols; ri=i//cols
        x=_CX+ci*(cw+gap); y=cy+ri*(ch+gap)
        c=rrect(slide,x,y,cw,ch,_rgb(T.card),radius_pct=11)
        if c: shadow(c,blur=14,dist=5,alpha=0.38)
        cs=min(cw,ch)*0.72
        oval(slide,x+(cw-cs)/2,y+(ch-cs)/2-0.3,cs,cs,_rgb(T.accent),alpha=5)
        txt_stat(slide,st.value,x+0.2,y+0.38,cw-0.4,ch*0.5,
                 font="Calibri",color=_rgb(T.accent),align=PP_ALIGN.CENTER)
        if st.unit:
            txt_label(slide,st.unit,x+0.2,y+ch*0.5+0.24,cw-0.4,0.55,
                      font=_FONT,size=11,color=_rgb(T.muted),
                      align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        lw=min(len(st.label)*0.22+1.0,cw-0.36)
        pill=rrect(slide,x+(cw-lw)/2,y+ch-1.08,lw,0.56,_rgb(T.accent),radius_pct=50)
        if pill:
            gradient_fill(pill,T.accent_grad1,T.accent_grad2,0)
            shadow(pill,blur=5,dist=2,alpha=0.25)
        txt_label(slide,st.label,x+(cw-lw)/2,y+ch-1.08,lw,0.56,
                  font=_FONT,size=9.5,color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
    return slide

def make_results(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"نتائج البحث","أبرز ما توصلت إليه الدراسة","08")
    results=req.main_results[:8]
    if not results: return slide
    n=len(results)
    cy=_BY+0.08; avail=H-cy-0.38
    gap=0.1
    # FILL full height — rows grow to fill
    ih=(avail - gap*(n-1)) / n
    fsize=16 if n<=3 else 14 if n<=5 else 12.5
    for i,res in enumerate(results):
        y=cy+i*(ih+gap)
        _row(slide,T,y,ih,res,str(i+1),i%2==0,fsize)
    return slide

def make_conclusion(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.bg2,T.grad2,178)
    _hdr(slide,T,"خاتمة البحث","الاستنتاج العام","09")
    cy=_BY+0.08; ch=H-cy-0.52
    c=rrect(slide,_CX,cy,_CW,ch,_rgb(T.card),radius_pct=11)
    if c: shadow(c,blur=22,dist=8,alpha=0.5)
    lb=rect(slide,_CX,cy,0.55,ch,_rgb(T.accent))
    if lb: gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
    txt(slide,"❝",_CX+0.72,cy+0.32,2.2,2.1,
        font="Calibri",size=58,bold=True,
        color=_rgb(T.accent),align=PP_ALIGN.LEFT,txt_shadow=True)
    ctext=req.general_conclusion or ""
    conc_size=17 if len(ctext)<100 else 15 if len(ctext)<200 else 13
    txt_quote(slide,ctext,_CX+0.85,cy+0.82,_CW-1.2,ch-1.15,
              font=_FONT,size=conc_size,color=_rgb(T.text_light),rtl=True)
    return slide

def make_recommendations(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"توصيات البحث","المقترحات العملية","10")
    recs=req.recommendations[:8]
    if not recs: return slide
    n=len(recs)
    cy=_BY+0.08; avail=H-cy-0.38
    gap=0.1
    # FILL full height
    ih=(avail - gap*(n-1)) / n
    fsize=16 if n<=3 else 14 if n<=5 else 13.5
    for i,rec in enumerate(recs):
        y=cy+i*(ih+gap)
        rb=rrect(slide,_CX,y,_CW,ih,_rgb(T.bg2) if i%2==0 else _rgb(T.card),radius_pct=5)
        if rb and i%2==0: shadow(rb,blur=5,dist=1,alpha=0.12)
        tg=rect(slide,_CX,y,0.42,ih,_rgb(T.accent))
        if tg: gradient_fill(tg,T.accent_grad1,T.accent_grad2,90)
        txt_body(slide,rec,_CX+0.62,y,_CW-0.82,ih,
                 font=_FONT,size=fsize,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True,valign_center=True)
    return slide

def make_future(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"آفاق البحث المستقبلية","","11")
    items=req.future_work[:6]
    if not items: return slide
    n=len(items)
    cols=3 if n>=4 else 2 if n>=2 else 1
    rows=(n+cols-1)//cols; gap=0.28
    cy=_BY+0.08; avail=H-cy-0.38
    cw=(_CW-(cols-1)*gap)/cols
    # FILL full height
    ch=avail/rows - gap*(rows-1)/rows
    num_size=34 if n<=2 else 28 if n==3 else 22
    body_size=15.5 if n<=2 else 14 if n==3 else 13
    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=_CX+ci*(cw+gap); y=cy+ri*(ch+gap)
        c=rrect(slide,x,y,cw,ch,_rgb(T.card),radius_pct=9)
        if c: shadow(c,blur=12,dist=4,alpha=0.32)
        bot=rect(slide,x,y+ch-0.4,cw,0.4,_rgb(T.accent))
        if bot: gradient_fill(bot,T.accent_grad1,T.accent_grad2,0)
        num_h=min(ch*0.22,1.15)
        # عرض أوسع للرقم لمنع التفاف النص
        txt(slide,f"{i+1:02d}",x+0.14,y+0.1,2.2,num_h,
            font="Calibri",size=num_size,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.LEFT,
            txt_shadow=True)
        hline(slide,x+0.16,y+num_h+0.22,cw-0.32,_rgb(T.muted),thickness=0.04)
        txt_body(slide,item,x+0.2,y+num_h+0.36,cw-0.4,ch-num_h-0.84,
                 font=_FONT,size=body_size,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True,valign_center=True)
    return slide

def make_references(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,178)
    _hdr(slide,T,"المراجع والمصادر","قائمة المراجع المعتمدة","12")
    refs=req.references[:14]
    if not refs: return slide
    n=len(refs)
    cy=_BY+0.04; avail=H-cy-0.32
    gap=0.07
    # FILL full height
    ih=(avail - gap*(n-1)) / n
    body_size=13 if n<=4 else 11.5 if n<=8 else 10.5
    for i,ref in enumerate(refs):
        y=cy+i*(ih+gap)
        if y+ih>H-0.2: break
        if i%2==0:
            rb=rrect(slide,_CX,y,_CW,ih,_rgb(T.bg2),radius_pct=4)
        nb=rrect(slide,_CX+_CW-1.85,y+(ih-0.4)/2,1.65,0.4,_rgb(T.accent),radius_pct=50)
        if nb: set_solid_alpha(nb,78)
        txt_label(slide,f"[{i+1}]",_CX+_CW-1.85,y+(ih-0.4)/2,1.65,0.4,
                  font="Calibri",size=9,color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER,rtl=False,uppercase=False)
        txt_body(slide,ref,_CX+0.18,y,_CW-2.15,ih,
                 font=_FONT,size=body_size,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True,valign_center=True)
    return slide

def make_final(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,148)
    sp=rect(slide,0,0,_SW,H,_rgb(T.bg2))
    if sp: gradient_fill(sp,T.bg,T.bg2,175)
    vline(slide,_SW-0.07,0,H,_rgb(T.accent),thickness=0.07)
    # زخارف دائرية
    for ox,oy,os,oa in [(_SW/2-6,H/2-7,14,5),(W-10,-5,16,4),(W*0.6,H*0.55,10,3)]:
        oval(slide,ox,oy,os,os,_rgb(T.accent),alpha=oa)
    cx=_SW+0.8; cw=W-cx-0.65
    cy=(H-13.5)/2; ch=13.5
    mc=rrect(slide,cx,cy,cw,ch,_rgb(T.card),radius_pct=13)
    if mc:
        gradient_fill(mc,T.card,T.bg2,142)
        shadow(mc,blur=32,dist=11,alpha=0.58)
    ct=rrect(slide,cx,cy,cw,0.56,_rgb(T.accent),radius_pct=13)
    if ct: gradient_fill(ct,T.accent_grad1,T.accent_grad2,0)
    txt_hero(slide,"شكراً وتقديراً",cx+0.8,cy+0.68,cw-1.6,3.1,
             font=_FONT,size=44,color=_rgb(T.text_light),
             align=PP_ALIGN.CENTER,rtl=True,shadow_on=True)
    # نقاط زخرفية
    for j in range(5):
        dx=cx+cw/2-1.4+j*0.7
        sz=0.42 if j==2 else 0.24
        dc=oval(slide,dx,cy+3.98,sz,sz,_rgb(T.accent))
        if dc:
            if j==2: glow(dc,T.accent,radius_pt=5,alpha=0.55)
            else: set_solid_alpha(dc,45)
    txt_hero(slide,req.student_name,cx+0.8,cy+4.55,cw-1.6,1.25,
             font=_FONT,size=20,color=_rgb(T.accent),
             align=PP_ALIGN.CENTER,rtl=True,shadow_on=False)
    short=req.title_ar[:85]+("..." if len(req.title_ar)>85 else "")
    txt_body(slide,short,cx+1.4,cy+5.98,cw-2.8,2.65,
             font=_FONT,size=12,color=_rgb(T.muted),
             align=PP_ALIGN.CENTER,rtl=True)
    hline(slide,cx+cw*0.15,cy+ch-1.38,cw*0.7,_rgb(T.accent),thickness=0.05)
    footer=[]
    if req.institution: footer.append(req.institution.split(" — ")[0])
    if req.year: footer.append(req.year)
    if footer:
        txt_label(slide," · ".join(footer),cx+0.8,cy+ch-1.18,cw-1.6,0.7,
                  font=_FONT,size=11,color=_rgb(T.muted),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
    return slide
