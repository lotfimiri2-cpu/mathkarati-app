"""CLASSIC Engine v19.2 — Academic Precision ★ VISUAL FIX"""
from __future__ import annotations
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, glow, shadow_and_glow,
    set_solid_alpha, txt, txt_hero, txt_label, txt_body, txt_stat,
    txt_quote, blank_slide,
)
from core.models import PresentationRequest
from core.themes import Theme

_FONT = "Cairo"
def set_font(f): global _FONT; _FONT = f
def _rgb(h):
    h=h.lstrip('#'); return RGBColor(int(h[:2],16),int(h[2:4],16),int(h[4:],16))

_SW=5.8; _CX=_SW+0.85; _CW=W-_CX-0.65; _HH=2.95; _BY=_HH+0.35


def _wrap_title(title, max_chars=35):
    """تقسيم ذكي للعنوان: كل سطر حوالي max_chars حرفاً"""
    if len(title) <= max_chars:
        return title
    words = title.split()
    # نقسم إلى مجموعات تقريبية
    total = len(title)
    n_lines = max(2, (total + max_chars - 1) // max_chars)
    target = total // n_lines
    lines = []; current = ""; current_len = 0
    for i, word in enumerate(words):
        if current_len >= target and len(lines) < n_lines - 1 and current:
            lines.append(current.strip())
            current = word; current_len = len(word)
        else:
            current = (current + " " + word).strip() if current else word
            current_len += len(word) + 1
    if current: lines.append(current.strip())
    return "\n".join(lines)

def _side(slide,T,num=""):
    sp=rect(slide,0,0,_SW,H,_rgb(T.bg2))
    if sp: gradient_fill(sp,T.bg,T.bg2,178)
    vline(slide,_SW-0.07,0,H,_rgb(T.accent),thickness=0.07)
    bot=rect(slide,0,H-1.15,_SW,1.15,_rgb(T.accent))
    if bot: gradient_fill(bot,T.accent_grad1,T.accent_grad2,175)
    if num:
        txt(slide,num,0,H-1.12,_SW,1.08,
            font="Calibri",size=30,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.CENTER,letter_spacing=-1.0)
    od=oval(slide,_SW*0.5-0.25,0.55,0.5,0.5,_rgb(T.accent))
    if od: glow(od,T.accent,radius_pt=5,alpha=0.45)

def _hdr(slide,T,title,sub="",num=""):
    hb=rect(slide,_SW,0,W-_SW,_HH,_rgb(T.bg))
    if hb: gradient_fill(hb,T.bg,T.bg2,88)
    bar=rect(slide,_SW,_HH-0.14,W-_SW,0.14,_rgb(T.accent))
    if bar: gradient_fill(bar,T.accent_grad1,T.accent_grad2,0)
    _side(slide,T,num)
    sz=26 if len(title)<20 else 21 if len(title)<32 else 16
    txt_hero(slide,title,_CX,0.38,_CW,1.82,
             font=_FONT,size=sz,color=_rgb(T.text_light),
             align=PP_ALIGN.RIGHT,rtl=True,shadow_on=True)
    if sub:
        hline(slide,_CX,2.08,_CW*0.32,_rgb(T.accent),thickness=0.07)
        txt_label(slide,sub,_CX,2.12,_CW,0.75,
                  font=_FONT,size=11,color=_rgb(T.muted),
                  align=PP_ALIGN.RIGHT,rtl=True,uppercase=False)

def _card(slide,x,y,w,h,T,radius=8):
    c=rrect(slide,x,y,w,h,_rgb(T.card),radius_pct=radius)
    if c: shadow(c,blur=14,dist=4,alpha=0.38)
    return c

def _info_card(slide,T,x,y,w,h,label,value):
    _card(slide,x,y,w,h,T,radius=7)
    top=rect(slide,x,y,w,0.4,_rgb(T.accent))
    if top: gradient_fill(top,T.accent_grad1,T.accent_grad2,0)
    txt_label(slide,label,x+0.1,y+0.03,w-0.2,0.36,
              font=_FONT,size=9,color=_rgb(T.text_dark),
              align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
    txt_body(slide,value,x+0.1,y+0.46,w-0.2,h-0.58,
             font=_FONT,size=12,color=_rgb(T.text_light),
             align=PP_ALIGN.CENTER,rtl=True)

def _row(slide,T,y,h,text,num="",alt=False,fsize=13.5):
    if alt:
        rb=rrect(slide,_CX,y,_CW,h,_rgb(T.bg2),radius_pct=5)
        if rb: set_solid_alpha(rb,55)
    if num:
        nb=rect(slide,_CX,y,0.38,h,_rgb(T.accent))
        if nb: gradient_fill(nb,T.accent_grad1,T.accent_grad2,90)
        tx,tw=_CX+0.55,_CW-0.72
    else:
        tx,tw=_CX+0.25,_CW-0.38
    txt_body(slide,text,tx,y+0.1,tw,h-0.2,
             font=_FONT,size=fsize,color=_rgb(T.text_light),
             align=PP_ALIGN.RIGHT,rtl=True)

# ════════════════════════════════════════════════════════════════════
def _draw_multiline_title(slide, title, x, y, w, max_h, font, color, align, rtl_on=True):
    """رسم العنوان على أسطر متعددة — 7 كلمات لكل سطر تقريباً"""
    words = title.split()
    words_per_line = 7  # عدد الكلمات لكل سطر

    lines = []
    for i in range(0, len(words), words_per_line):
        lines.append(" ".join(words[i:i + words_per_line]))
    if not lines:
        lines = [title]

    total = len(title)
    base_sz = 24 if total < 40 else 20 if total < 65 else 17 if total < 90 else 14
    
    n = len(lines)
    line_h = min(max_h / n, 2.8)
    # توسيط العنوان عمودياً
    total_used = n * line_h * 1.15
    start_y = y + max(0, (max_h - total_used) / 2)

    for i, line in enumerate(lines):
        ly = start_y + i * (line_h * 1.15)
        if ly + line_h > y + max_h + 0.3:
            break
        txt(slide, line, x, ly, w, line_h,
            font=font, size=base_sz, bold=True,
            color=color, align=align, rtl=rtl_on,
            letter_spacing=0.4)


def make_cover(prs,req,T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    # خلفية الجانب الأيمن
    mb=rect(slide,_SW,0,W-_SW,H,_rgb(T.bg))
    if mb: gradient_fill(mb,T.bg,T.bg2,132)
    # شريط علوي
    tb=rect(slide,_SW,0,W-_SW,0.2,_rgb(T.accent))
    if tb: gradient_fill(tb,T.accent_grad1,T.accent_grad2,0)
    # الشريط الجانبي
    _side(slide,T)
    if req.institution:
        parts=req.institution.split(" — ")
        for i,p in enumerate(parts[:3]):
            txt_body(slide,p,0.2,1.3+i*1.1,_SW-0.4,0.92,
                     font=_FONT,size=9.5 if i==0 else 8.5,
                     color=_rgb(T.text_light) if i==0 else _rgb(T.muted),
                     align=PP_ALIGN.CENTER,rtl=True)
    if req.specialization:
        sp=rrect(slide,0.45,H*0.42,_SW-0.9,0.56,_rgb(T.accent),radius_pct=50)
        if sp: gradient_fill(sp,T.accent_grad1,T.accent_grad2,0)
        txt_label(slide,req.specialization,0.45,H*0.42,_SW-0.9,0.56,
                  font=_FONT,size=10,color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
    # ── خط أكسنت + عنوان يملأ المنتصف ──────────────────
    hline(slide,_CX,1.8,_CW,_rgb(T.accent),thickness=0.14)
    dot=oval(slide,_CX-0.18,1.72,0.3,0.3,_rgb(T.accent))
    if dot: glow(dot,T.accent,radius_pt=6,alpha=0.5)
    # العنوان — منطقة كبيرة من أعلى إلى أسفل
    # العنوان يملأ المنتصف — حجم ثابت يضمن أسطراً متعددة
    _draw_multiline_title(slide, req.title_ar,
        _CX, 1.85, _CW, 7.8,
        font=_FONT, color=_rgb(T.text_light),
        align=PP_ALIGN.RIGHT, rtl_on=True)
    if req.title_en:
        txt(slide,req.title_en,_CX,H*0.54,_CW,0.82,
            font="Calibri",size=11,italic=True,
            color=_rgb(T.muted),align=PP_ALIGN.RIGHT,letter_spacing=0.5)
    hline(slide,_CX,10.32,_CW,_rgb(T.muted),thickness=0.05)
    # ── بطاقات المعلومات أسفل الشريحة ──────────────────
    fields=[("الطالب",req.student_name)]
    if req.supervisor:    fields.append(("المشرف",req.supervisor))
    if req.co_supervisor: fields.append(("م. مساعد",req.co_supervisor))
    n=len(fields)
    iw=min((_CW-0.22*(n-1))/n,10.5)
    tot=n*iw+0.22*(n-1); ix0=_CX+(_CW-tot)/2
    for i,(lbl,val) in enumerate(fields[:3]):
        cx2=ix0+i*(iw+0.22); cy2=H*0.57
        _info_card(slide,T,cx2,cy2,iw,2.1,lbl,val)
    if req.year:
        txt_label(slide,req.year,_CX,H-0.95,_CW,0.82,
                  font="Calibri",size=16,color=_rgb(T.accent),
                  align=PP_ALIGN.LEFT,rtl=False,uppercase=False)
    return slide


def make_intro(prs,req,T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"مقدمة البحث","نظرة عامة على الدراسة","01")
    cy=_BY+0.08; ch=H-cy-0.42
    if req.intro_overview and req.intro_approach:
        cw1=_CW*0.57; cw2=_CW-cw1-0.38
        for titl,text,cw,cx in [
            ("نظرة عامة",req.intro_overview,cw1,_CX),
            ("المقاربة المنهجية",req.intro_approach,cw2,_CX+cw1+0.38)]:
            c=rect(slide,cx,cy,cw,ch,_rgb(T.card))
            if c: shadow(c,blur=4,dist=1,alpha=0.18)
            tp=rect(slide,cx,cy,cw,0.64,_rgb(T.accent))
            if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
            txt_label(slide,titl,cx+0.32,cy+0.07,cw-0.54,0.54,
                      font=_FONT,size=14.5,color=_rgb(T.text_dark),
                      align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
            txt_body(slide,text,cx+0.32,cy+0.8,cw-0.62,ch-0.98,
                     font=_FONT,size=14,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
    else:
        text=req.intro_overview or req.intro_approach
        c=rect(slide,_CX,cy,_CW,ch,_rgb(T.card))
        if c: shadow(c,blur=4,dist=1,alpha=0.18)
        lb=rect(slide,_CX,cy,0.52,ch,_rgb(T.accent))
        if lb: gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
        txt(slide,"❝",_CX+0.65,cy+0.28,2.0,1.8,
            font="Calibri",size=55,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.LEFT,txt_shadow=True)
        txt_quote(slide,text,_CX+0.72,cy+0.72,_CW-1.08,ch-0.92,
                  font=_FONT,size=15,color=_rgb(T.text_light),rtl=True)
    return slide

def make_plan(prs,req,T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"خطة البحث","هيكل ومحتويات الدراسة","02")
    chapters=req.chapters[:8]
    if not chapters: return slide
    cy=_BY+0.08; avail=H-cy-0.42
    n=len(chapters)
    ih=min(avail/n-0.1,1.82)
    for i,ch in enumerate(chapters):
        y=cy+i*(ih+0.1)
        row=rrect(slide,_CX,y,_CW,ih,
                  _rgb(T.bg2) if i%2==0 else _rgb(T.card),radius_pct=6)
        if row and i%2==0: set_solid_alpha(row,65)
        nb=rrect(slide,_CX,y,2.4,ih,_rgb(T.accent),radius_pct=6)
        if nb: gradient_fill(nb,T.accent_grad1,T.accent_grad2,90)
        txt(slide,f"{i+1:02d}",_CX,y,2.4,ih,
            font="Calibri",size=24,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.CENTER,letter_spacing=-1.0)
        txt_body(slide,ch.title,_CX+2.55,y+0.12,_CW-5.1,ih-0.24,
                 font=_FONT,size=15,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
        if ch.pages:
            txt_label(slide,f"ص {ch.pages}",_CX+_CW-2.6,y+(ih-0.5)/2,2.4,0.5,
                      font="Calibri",size=11,color=_rgb(T.muted),
                      align=PP_ALIGN.LEFT,rtl=False,uppercase=False)
    return slide

def make_problem(prs,req,T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"إشكالية البحث","التساؤلات المحورية للدراسة","03")
    cy=_BY+0.05
    if req.main_problem or req.main_question:
        text=req.main_problem or req.main_question
        qh=min(3.4,H*0.33)
        _card(slide,_CX,cy,_CW,qh,T,radius=9)
        lb=rrect(slide,_CX,cy,0.48,qh,_rgb(T.accent),radius_pct=9)
        if lb: gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
        txt(slide,"❝",_CX+0.62,cy+0.22,2.0,qh-0.44,
            font="Calibri",size=52,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.LEFT,txt_shadow=True)
        txt_quote(slide,text,_CX+0.68,cy+0.28,_CW-0.98,qh-0.48,
                  font=_FONT,size=14.5,color=_rgb(T.text_light),rtl=True)
        cy+=qh+0.28
    if req.sub_questions:
        avail=H-cy-0.32
        sq_h=min(avail/max(len(req.sub_questions),1)-0.1,1.05)
        for i,q in enumerate(req.sub_questions[:6]):
            y=cy+i*(sq_h+0.1)
            _row(slide,T,y,sq_h,q,str(i+1),i%2==0,13.5)
    return slide

def make_objectives(prs,req,T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"أهداف البحث وفرضياته","","04")
    cy=_BY+0.08; cha=H-cy-0.42; gap=0.38
    cols=[]
    if req.objectives: cols.append(("الأهداف",req.objectives))
    if req.hypotheses:  cols.append(("الفرضيات",req.hypotheses))
    if not cols: return slide
    widths=[_CW*0.56,_CW*0.42] if len(cols)==2 else [_CW]
    x=_CX
    for i,(lbl,items) in enumerate(cols[:2]):
        cw=widths[i]
        _card(slide,x,cy,cw,cha,T,radius=9)
        hd=rect(slide,x,cy,cw,0.62,_rgb(T.accent))
        if hd: gradient_fill(hd,T.accent_grad1,T.accent_grad2,0)
        txt_label(slide,lbl,x+0.18,cy+0.06,cw-0.36,0.52,
                  font=_FONT,size=15,color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        ih=min((cha-0.68)/max(len(items),1),1.2)
        for j,item in enumerate(items[:8]):
            iy=cy+0.68+j*ih
            if iy+ih>cy+cha-0.08: break
            if j>0: hline(slide,x+0.22,iy,cw-0.44,_rgb(T.bg),thickness=0.03)
            nb=oval(slide,x+cw-1.38,iy+(ih-0.44)/2,0.44,0.44,_rgb(T.accent))
            if nb: set_solid_alpha(nb,88)
            txt(slide,str(j+1),x+cw-1.38,iy+(ih-0.44)/2,0.44,0.44,
                font="Calibri",size=9.5,bold=True,
                color=_rgb(T.text_dark),align=PP_ALIGN.CENTER)
            txt_body(slide,item,x+0.25,iy+0.09,cw-1.88,ih-0.18,
                     font=_FONT,size=13,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
        x+=cw+gap
    return slide

def make_importance(prs,req,T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"أهمية البحث ومبرراته","","05")
    items=list(req.importance or [])
    if req.reasons and req.reasons not in items: items.append(req.reasons)
    items=items[:6]
    if not items: return slide
    cy=_BY+0.08; avH=H-cy-0.38
    cols=3 if len(items)>=4 else 2 if len(items)>=2 else 1
    rows=(len(items)+cols-1)//cols; gap=0.28
    cw=(_CW-(cols-1)*gap)/cols
    ch=min(avH/rows-gap*0.9,3.8)
    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=_CX+ci*(cw+gap); y=cy+ri*(ch+gap)
        rect(slide,x,y,cw,ch,_rgb(T.card))
        top=rect(slide,x,y,cw,0.72,_rgb(T.accent))
        if top: gradient_fill(top,T.accent_grad1,T.accent_grad2,0)
        txt(slide,f"{i+1:02d}",x+0.14,y+0.04,1.1,0.65,
            font="Calibri",size=22,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.LEFT,letter_spacing=-1.0)
        txt_body(slide,item,x+0.2,y+0.85,cw-0.4,ch-1.05,
                 font=_FONT,size=13,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_methodology(prs,req,T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"منهجية البحث","الإجراءات والأدوات المستخدمة","06")
    fields=[]
    if req.methodology: fields.append(("المنهج المتبع","م",req.methodology))
    if req.sample_type: fields.append(("نوع العينة","ع",req.sample_type))
    if req.sample_size: fields.append(("حجم العينة","ن",req.sample_size))
    if req.tool:        fields.append(("أداة الجمع","أ",req.tool))
    if not fields: return slide
    n=len(fields); cy=_BY+0.08; ch=H-cy-0.48
    cw=(_CW-0.28*(n-1))/n
    if n>1: hline(slide,_CX+cw/2,cy+ch*0.5,_CW-cw,_rgb(T.muted),thickness=0.04)
    for i,(lbl,icon,val) in enumerate(fields[:4]):
        x=_CX+i*(cw+0.28)
        rect(slide,x,cy,cw,ch,_rgb(T.card))
        sz=1.72; ix=x+(cw-sz)/2
        ic=oval(slide,ix,cy+0.5,sz,sz,_rgb(T.accent))
        if ic:
            gradient_fill(ic,T.accent_grad1,T.accent_grad2,45)
            shadow_and_glow(ic,s_blur=10,s_dist=3,s_alpha=0.38,
                            g_color=T.accent_grad2,g_rad=5,g_alpha=0.22)
        txt(slide,icon,ix,cy+0.5,sz,sz,
            font=_FONT,size=26,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.CENTER)
        hline(slide,x+0.28,cy+2.42,cw-0.56,_rgb(T.accent),thickness=0.07)
        txt_label(slide,lbl,x+0.12,cy+2.54,cw-0.24,0.66,
                  font=_FONT,size=13,color=_rgb(T.accent),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        txt_body(slide,val,x+0.2,cy+3.3,cw-0.4,ch-3.46,
                 font=_FONT,size=13,color=_rgb(T.text_light),
                 align=PP_ALIGN.CENTER,rtl=True)
    return slide

def make_stats(prs,req,T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"الإحصاءات والأرقام الرئيسية","مؤشرات كمية للدراسة","07")
    stats=req.stats[:6]
    if not stats: return slide
    n=len(stats); cols=3 if n>=3 else n
    rows=(n+cols-1)//cols; gap=0.3
    cy=_BY+0.08; avail=H-cy-0.38
    cw=(_CW-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,4.8)
    for i,st in enumerate(stats):
        ci=i%cols; ri=i//cols
        x=_CX+ci*(cw+gap); y=cy+ri*(ch+gap)
        rect(slide,x,y,cw,ch,_rgb(T.card))
        top=rect(slide,x,y,cw,0.52,_rgb(T.accent))
        if top: gradient_fill(top,T.accent_grad1,T.accent_grad2,0)
        txt_stat(slide,st.value,x+0.18,y+0.58,cw-0.36,ch*0.46,
                 font="Calibri",color=_rgb(T.accent),align=PP_ALIGN.CENTER)
        if st.unit:
            txt_label(slide,st.unit,x+0.18,y+ch*0.57,cw-0.36,0.55,
                      font=_FONT,size=11.5,color=_rgb(T.muted),
                      align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        hline(slide,x+0.35,y+ch-1.1,cw-0.7,_rgb(T.accent),thickness=0.07)
        txt_label(slide,st.label,x+0.14,y+ch-0.98,cw-0.28,0.88,
                  font=_FONT,size=12,color=_rgb(T.text_light),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
    return slide

def make_results(prs,req,T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"نتائج البحث","أبرز ما توصلت إليه الدراسة","08")
    results=req.main_results[:8]
    if not results: return slide
    cy=_BY+0.08; avail=H-cy-0.38
    ih=min(avail/max(len(results),1)-0.1,1.65)
    for i,res in enumerate(results):
        y=cy+i*(ih+0.1)
        _row(slide,T,y,ih,res,str(i+1),i%2==0,13.5)
    return slide

def make_conclusion(prs,req,T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"خاتمة البحث","الاستنتاج العام","09")
    cy=_BY+0.08; ch=H-cy-0.52
    _card(slide,_CX,cy,_CW,ch,T,radius=11)
    lb=rect(slide,_CX,cy,0.55,ch,_rgb(T.accent))
    if lb: gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
    txt(slide,"❝",_CX+0.7,cy+0.32,2.0,1.8,
        font="Calibri",size=56,bold=True,
        color=_rgb(T.accent),align=PP_ALIGN.LEFT,txt_shadow=True)
    txt_quote(slide,req.general_conclusion,
              _CX+0.82,cy+1.0,_CW-1.18,ch-1.5,
              font=_FONT,size=15,color=_rgb(T.text_light),rtl=True)
    return slide

def make_recommendations(prs,req,T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"توصيات البحث","المقترحات العملية","10")
    recs=req.recommendations[:8]
    if not recs: return slide
    cy=_BY+0.08; avail=H-cy-0.38
    ih=min(avail/max(len(recs),1)-0.1,1.55)
    for i,rec in enumerate(recs):
        y=cy+i*(ih+0.1)
        rrect(slide,_CX,y,_CW,ih,
              _rgb(T.bg2) if i%2==0 else _rgb(T.card),radius_pct=5)
        tg=rect(slide,_CX,y,0.42,ih,_rgb(T.accent))
        if tg: gradient_fill(tg,T.accent_grad1,T.accent_grad2,90)
        txt_body(slide,rec,_CX+0.62,y+0.1,_CW-0.82,ih-0.2,
                 font=_FONT,size=13.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_future(prs,req,T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"آفاق البحث المستقبلية","","11")
    items=req.future_work[:6]
    if not items: return slide
    cols=3 if len(items)>=4 else 2 if len(items)>=2 else 1
    rows=(len(items)+cols-1)//cols; gap=0.28
    cy=_BY+0.08; avail=H-cy-0.38
    cw=(_CW-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,3.35)
    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=_CX+ci*(cw+gap); y=cy+ri*(ch+gap)
        rect(slide,x,y,cw,ch,_rgb(T.card))
        bot=rect(slide,x,y+ch-0.4,cw,0.4,_rgb(T.accent))
        if bot: gradient_fill(bot,T.accent_grad1,T.accent_grad2,0)
        txt(slide,f"{i+1:02d}",x+0.16,y+0.13,1.3,0.9,
            font="Calibri",size=24,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.LEFT,
            letter_spacing=-1.0,txt_shadow=True)
        hline(slide,x+0.16,y+1.1,cw-0.32,_rgb(T.muted),thickness=0.04)
        txt_body(slide,item,x+0.2,y+1.22,cw-0.42,ch-1.72,
                 font=_FONT,size=13,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_references(prs,req,T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"المراجع والمصادر","قائمة المراجع المعتمدة","12")
    refs=req.references[:14]
    if not refs: return slide
    cy=_BY+0.04; avail=H-cy-0.32
    ih=max(min(avail/max(len(refs),1)-0.07,1.05),0.5)
    for i,ref in enumerate(refs):
        y=cy+i*(ih+0.07)
        if y+ih>H-0.22: break
        if i%2==0: rrect(slide,_CX,y,_CW,ih,_rgb(T.bg2),radius_pct=4)
        nb=rrect(slide,_CX+_CW-1.88,y+(ih-0.4)/2,1.68,0.4,_rgb(T.accent),radius_pct=50)
        if nb: set_solid_alpha(nb,78)
        txt_label(slide,f"[{i+1}]",_CX+_CW-1.88,y+(ih-0.4)/2,1.68,0.4,
                  font="Calibri",size=9,color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER,rtl=False,uppercase=False)
        txt_body(slide,ref,_CX+0.18,y+0.05,_CW-2.18,ih-0.1,
                 font=_FONT,size=11,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_final(prs,req,T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,150)
    sp=rect(slide,0,0,_SW,H,_rgb(T.bg2))
    if sp: gradient_fill(sp,T.bg,T.bg2,180)
    vline(slide,_SW-0.07,0,H,_rgb(T.accent),thickness=0.07)
    oval(slide,-2,-2,8,8,_rgb(T.accent),alpha=5)
    oval(slide,W-6,H-5,9,9,_rgb(T.accent),alpha=4)
    tb=rect(slide,0,0,W,0.2,_rgb(T.accent))
    if tb: gradient_fill(tb,T.accent_grad1,T.accent_grad2,0)
    bb=rect(slide,0,H-0.2,W,0.2,_rgb(T.accent))
    if bb: gradient_fill(bb,T.accent_grad2,T.accent_grad1,0)
    cw=W-_SW-2.0; cx=_SW+1.0; cy=H*0.07; ch=H*0.86
    mc=rrect(slide,cx,cy,cw,ch,_rgb(T.card),radius_pct=11)
    if mc: shadow(mc,blur=28,dist=10,alpha=0.52)
    ct=rect(slide,cx,cy,cw,0.48,_rgb(T.accent))
    if ct: gradient_fill(ct,T.accent_grad1,T.accent_grad2,0)
    txt_hero(slide,"شكراً وتقديراً",cx+0.8,cy+0.62,cw-1.6,3.4,
             font=_FONT,size=44,color=_rgb(T.text_light),
             align=PP_ALIGN.CENTER,rtl=True,shadow_on=True)
    for j in range(5):
        dx=cx+cw/2-1.5+j*0.75; sz=0.42 if j==2 else 0.25
        dc=oval(slide,dx,cy+4.2,sz,sz,_rgb(T.accent))
        if dc:
            if j==2: glow(dc,T.accent,radius_pt=5,alpha=0.55)
            else: set_solid_alpha(dc,48)
    txt_hero(slide,req.student_name,cx+0.8,cy+4.82,cw-1.6,1.35,
             font=_FONT,size=20,color=_rgb(T.accent),
             align=PP_ALIGN.CENTER,rtl=True,shadow_on=False)
    short=req.title_ar[:88]+("..." if len(req.title_ar)>88 else "")
    txt_body(slide,short,cx+1.5,cy+6.35,cw-3.0,2.8,
             font=_FONT,size=13,color=_rgb(T.muted),
             align=PP_ALIGN.CENTER,rtl=True)
    hline(slide,cx+cw*0.15,cy+ch-1.45,cw*0.7,_rgb(T.accent),thickness=0.05)
    footer=[]
    if req.institution: footer.append(req.institution.split(" — ")[0])
    if req.year: footer.append(req.year)
    if footer:
        txt_label(slide," · ".join(footer),cx+0.8,cy+ch-1.22,cw-1.6,0.75,
                  font=_FONT,size=11,color=_rgb(T.muted),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
    return slide
