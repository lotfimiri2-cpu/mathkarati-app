"""CANVA Engine v19.2 — Dynamic Editorial ★ VISUAL FIX"""
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
def _W(): return W-1.6

# ── Primitives ────────────────────────────────────────────────────────

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

def _card(slide, x, y, w, h, T, radius=10):
    """بطاقة نظيفة — solid fill فقط، لا gradient داخلي"""
    c = rrect(slide, x, y, w, h, _rgb(T.card), radius_pct=radius)
    if c: shadow(c, blur=6, dist=2, alpha=0.25)
    return c

def _pill_accent(slide, T, x, y, w, h, text, fsize=10):
    p = rrect(slide, x, y, w, h, _rgb(T.accent), radius_pct=50)
    if p:
        gradient_fill(p, T.accent_grad1, T.accent_grad2, 0)
        shadow(p, blur=5, dist=1, alpha=0.25)
    txt_label(slide, text, x, y, w, h, font=_FONT, size=fsize,
              color=_rgb(T.text_dark), align=PP_ALIGN.CENTER, rtl=True, uppercase=False)

def _hdr(slide, T, title, sub=""):
    hb = rect(slide, 0, 0, W, 3.1, _rgb(T.bg2))
    if hb: gradient_fill(hb, T.grad1, T.grad2, 180)
    bar = rect(slide, 0, 2.98, W, 0.12, _rgb(T.accent))
    if bar: gradient_fill(bar, T.accent_grad1, T.accent_grad2, 0)
    # زخرفة بسيطة — دوائر خارج نطاق الشريحة فقط
    oval(slide, W-5, -4, 8, 8, _rgb(T.accent), alpha=8)
    oval(slide, -1, -1, 5, 5, _rgb(T.accent2), alpha=6)
    sz = 26 if len(title)<18 else 21 if len(title)<30 else 16
    txt_hero(slide, title, 1.2, 0.35, W-2.4, 1.85,
             font=_FONT, size=sz, color=_rgb(T.text_light),
             align=PP_ALIGN.RIGHT, rtl=True, shadow_on=True)
    if sub:
        txt_label(slide, "◈  "+sub, 1.2, 2.16, W-2.4, 0.76,
                  font=_FONT, size=11.5, color=_rgb(T.muted),
                  align=PP_ALIGN.RIGHT, rtl=True, uppercase=False)

# ── Cover ─────────────────────────────────────────────────────────────
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


def make_cover(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 150)
    # زخرفة خارج الحدود فقط
    oval(slide, W-3, -2, 6, 6, _rgb(T.accent), alpha=7)
    oval(slide, -1, H-3, 5, 5, _rgb(T.accent2), alpha=5)
    # شريط علوي قوي
    top = rect(slide, 0, 0, W, 0.55, _rgb(T.accent))
    if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)
    # شريط سفلي
    bot = rect(slide, 0, H-0.42, W, 0.42, _rgb(T.accent))
    if bot: gradient_fill(bot, T.accent_grad2, T.accent_grad1, 0)
    # المؤسسة
    if req.institution:
        inst = req.institution.split(' — ')[0]
        txt_body(slide, inst, 1.5, 0.72, W-3, 0.65,
                 font=_FONT, size=11, color=_rgb(T.muted),
                 align=PP_ALIGN.CENTER, rtl=True)
    # ── خط فاصل تحت المؤسسة ─────────────────────────────
    hline(slide, 1.5, 1.55, W-3, _rgb(T.accent), thickness=0.1)
    # ── العنوان الضخم يملأ المنتصف ───────────────────────
    _draw_multiline_title(slide, req.title_ar,
        1.5, 1.72, W-3, 7.8,
        font=_FONT, color=_rgb(T.text_light),
        align=PP_ALIGN.CENTER, rtl_on=True)
    if req.title_en:
        txt(slide, req.title_en, 1.5, H*0.54, W-3, 0.8,
            font="Calibri", size=11, italic=True,
            color=_rgb(T.muted), align=PP_ALIGN.CENTER, letter_spacing=0.5)
    # ── خط فاصل + بطاقات معلومات ─────────────────────────
    hline(slide, 1.5, H*0.56-0.18, W-3, _rgb(T.accent), thickness=0.12)
    fields = [("الطالب", req.student_name)]
    if req.supervisor:     fields.append(("المشرف", req.supervisor))
    if req.co_supervisor:  fields.append(("م. مساعد", req.co_supervisor))
    if req.specialization: fields.append(("التخصص", req.specialization))
    n = len(fields)
    pw = min((_W()-0.22*(n-1))/n, 9.0)
    tot = n*(pw+0.22)-0.22; sx = 0.8+(_W()-tot)/2
    for i,(lbl,val) in enumerate(fields[:4]):
        px = sx+i*(pw+0.22); py = H*0.57
        pc = rect(slide, px, py, pw, 2.0, _rgb(T.card))
        lb = rrect(slide, px, py, pw, 0.54, _rgb(T.accent), radius_pct=5)
        if lb: gradient_fill(lb, T.accent_grad1, T.accent_grad2, 0)
        txt_label(slide, lbl, px, py+0.04, pw, 0.48,
                  font=_FONT, size=10, color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER, rtl=True, uppercase=False)
        txt_body(slide, val, px+0.18, py+0.62, pw-0.36, 1.28,
                 font=_FONT, size=13, color=_rgb(T.text_light),
                 align=PP_ALIGN.CENTER, rtl=True)
    if req.year:
        yr = rrect(slide, W/2-2.8, H-1.32, 5.6, 0.72, _rgb(T.accent), radius_pct=50)
        if yr: gradient_fill(yr, T.accent_grad1, T.accent_grad2, 0)
        txt_label(slide, req.year, W/2-2.8, H-1.32, 5.6, 0.72,
                  font="Calibri", size=13, color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER, rtl=False, uppercase=False)
    return slide


# ── Intro ──────────────────────────────────────────────────────────────
def make_intro(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"مقدمة البحث","نظرة عامة ومقاربة الدراسة")
    cy=3.14; ch=H-cy-0.42
    if req.intro_overview and req.intro_approach:
        gap=0.4; cw1=_W()*0.55; cw2=_W()-cw1-gap
        for titl,text,cw,cx in [
            ("الإطار العام",req.intro_overview,cw1,0.8),
            ("المقاربة",req.intro_approach,cw2,0.8+cw1+gap)]:
            c=rect(slide,cx,cy,cw,ch,_rgb(T.card))
            if c: shadow(c,blur=4,dist=1,alpha=0.18)
            tp=rect(slide,cx,cy,cw,0.64,_rgb(T.accent))
            if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
            txt_label(slide,titl,cx+0.25,cy+0.07,cw-0.5,0.54,
                      font=_FONT,size=14.5,color=_rgb(T.text_dark),
                      align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
            txt_body(slide,text,cx+0.3,cy+0.8,cw-0.6,ch-0.98,
                     font=_FONT,size=14,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
    else:
        text=req.intro_overview or req.intro_approach
        c=rect(slide,0.8,cy,_W(),ch,_rgb(T.card))
        if c: shadow(c,blur=4,dist=1,alpha=0.18)
        lb=rect(slide,0.8,cy,0.52,ch,_rgb(T.accent))
        if lb: gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
        txt(slide,"❝",1.5,cy+0.3,2.0,1.8,font="Calibri",size=55,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.LEFT,txt_shadow=True)
        txt_quote(slide,text,1.6,cy+0.68,_W()-1.05,ch-0.85,
                  font=_FONT,size=15,color=_rgb(T.text_light),rtl=True)
    return slide

# ── Plan ───────────────────────────────────────────────────────────────
def make_plan(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"خطة البحث","محتويات الدراسة وهيكلها")
    chapters=req.chapters[:8]
    if not chapters: return slide
    n=len(chapters)
    cy=3.14; avail=H-cy-0.42
    # التخطيط المتجاوب
    if n<=3: cols=n; rows=1
    elif n<=6: cols=3; rows=(n+2)//3
    else: cols=4; rows=(n+3)//4
    gap=0.3
    cw=(_W()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap, 4.2)
    for i,chap in enumerate(chapters):
        ci=i%cols; ri=i//cols
        x=0.8+ci*(cw+gap); y=cy+ri*(ch+gap)
        _card(slide,x,y,cw,ch,T,radius=10)
        tp=rect(slide,x,y,cw,0.6,_rgb(T.accent))
        if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
        txt(slide,f"{i+1:02d}",x+0.14,y+0.06,1.1,0.5,
            font="Calibri",size=20,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.LEFT,letter_spacing=-1.0)
        hline(slide,x+0.2,y+0.74,cw-0.4,_rgb(T.muted),thickness=0.04)
        txt_body(slide,chap.title,x+0.2,y+0.86,cw-0.4,ch-1.28,
                 font=_FONT,size=14.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
        if chap.pages:
            txt_label(slide,f"ص {chap.pages}",x+0.18,y+ch-0.52,cw-0.36,0.44,
                      font="Calibri",size=11,color=_rgb(T.muted),
                      align=PP_ALIGN.LEFT,rtl=False,uppercase=False)
    return slide

# ── Problem ─────────────────────────────────────────────────────────────
def make_problem(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"إشكالية البحث","التساؤلات المحورية للدراسة")
    cy=3.14
    if req.main_problem or req.main_question:
        text=req.main_problem or req.main_question
        qh=min(3.2,H*0.32)
        qc=rrect(slide,0.8,cy,_W(),qh,_rgb(T.card),radius_pct=10)
        if qc: shadow(qc,blur=5,dist=1,alpha=0.20)
        lb=rect(slide,0.8,cy,0.48,qh,_rgb(T.accent))
        if lb: gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
        txt(slide,"❝",1.45,cy+0.18,2.0,qh-0.36,
            font="Calibri",size=50,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.RIGHT,txt_shadow=True)
        txt_quote(slide,text,1.5,cy+0.28,_W()-1.1,qh-0.56,
                  font=_FONT,size=14.5,color=_rgb(T.text_light),rtl=True)
        cy+=qh+0.28
    if req.sub_questions:
        sq=req.sub_questions[:6]; n2=len(sq)
        cols=2 if n2>2 else 1
        cw=(_W()-(cols-1)*0.28)/cols
        avail=H-cy-0.38
        sh=min(avail/((n2+cols-1)//cols)-0.1,1.25)
        for i,q in enumerate(sq):
            ci=i%cols; ri=i//cols
            x=0.8+ci*(cw+0.28); y=cy+ri*(sh+0.1)
            rc=rrect(slide,x,y,cw,sh,_rgb(T.card),radius_pct=6)
            if rc: shadow(rc,blur=8,dist=2,alpha=0.22)
            lb2=rect(slide,x,y,0.42,sh,_rgb(T.accent))
            if lb2: gradient_fill(lb2,T.accent_grad1,T.accent_grad2,90)
            _pill_accent(slide,T,x+cw-1.75,y+(sh-0.5)/2,1.6,0.5,str(i+1),11)
            txt_body(slide,q,x+0.6,y+0.1,cw-2.48,sh-0.2,
                     font=_FONT,size=13.5,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
    return slide

# ── Objectives ──────────────────────────────────────────────────────────
def make_objectives(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"أهداف البحث وفرضياته","")
    cy=3.14; ch=H-cy-0.42; gap=0.38
    cols_data=[]
    if req.objectives: cols_data.append(("الأهداف",req.objectives))
    if req.hypotheses:  cols_data.append(("الفرضيات",req.hypotheses))
    if not cols_data: return slide
    widths=[_W()*0.56,_W()*0.42] if len(cols_data)==2 else [_W()]
    x=0.8
    for i,(lbl,items) in enumerate(cols_data[:2]):
        cw=widths[i]
        rect(slide,x,cy,cw,ch,_rgb(T.card))
        hd=rect(slide,x,cy,cw,0.62,_rgb(T.accent))
        if hd: gradient_fill(hd,T.accent_grad1,T.accent_grad2,0)
        txt_label(slide,lbl,x+0.2,cy+0.07,cw-0.4,0.52,
                  font=_FONT,size=15,color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        ih=min((ch-0.7)/max(len(items),1),1.2)
        for j,item in enumerate(items[:8]):
            iy=cy+0.7+j*ih
            if iy+ih>cy+ch-0.1: break
            if j>0: hline(slide,x+0.22,iy,cw-0.44,_rgb(T.bg),thickness=0.03)
            _pill_accent(slide,T,x+cw-1.52,iy+(ih-0.48)/2,1.38,0.48,str(j+1),10)
            txt_body(slide,item,x+0.22,iy+0.09,cw-1.92,ih-0.18,
                     font=_FONT,size=13,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
        x+=cw+gap
    return slide

# ── Importance ──────────────────────────────────────────────────────────
def make_importance(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"أهمية البحث","مبررات اختيار الموضوع")
    items=list(req.importance or [])
    if req.reasons and req.reasons not in items: items.append(req.reasons)
    items=items[:6]
    if not items: return slide
    cols=3 if len(items)>=4 else 2 if len(items)>=2 else 1
    rows=(len(items)+cols-1)//cols; gap=0.28
    cy=3.14; avail=H-cy-0.4
    cw=(_W()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,3.9)
    accent_h=[0.64,0.82,0.55,0.72,0.62,0.78]
    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=0.8+ci*(cw+gap); y=cy+ri*(ch+gap)
        rect(slide,x,y,cw,ch,_rgb(T.card))
        ah=accent_h[i%len(accent_h)]
        tp=rect(slide,x,y,cw,ah,_rgb(T.accent))
        if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
        txt(slide,f"{i+1:02d}",x+0.15,y+0.04,1.1,ah-0.08,
            font="Calibri",size=22,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.LEFT,letter_spacing=-1.0)
        txt_body(slide,item,x+0.22,y+ah+0.18,cw-0.44,ch-ah-0.3,
                 font=_FONT,size=13,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

# ── Methodology ─────────────────────────────────────────────────────────
def make_methodology(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"منهجية البحث","الإجراءات والأدوات المستخدمة")
    fields=[]
    if req.methodology: fields.append(("المنهج","م",req.methodology))
    if req.sample_type: fields.append(("العينة","ع",req.sample_type))
    if req.sample_size: fields.append(("الحجم","ن",req.sample_size))
    if req.tool:        fields.append(("الأداة","أ",req.tool))
    if not fields: return slide
    n=len(fields); cy=3.14; ch=H-cy-0.46
    cw=(_W()-0.28*(n-1))/n
    if n>1: hline(slide,0.8+cw/2,cy+ch*0.5,_W()-cw,_rgb(T.muted),thickness=0.04)
    for i,(lbl,icon,val) in enumerate(fields[:4]):
        x=0.8+i*(cw+0.28)
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
        txt_label(slide,lbl,x+0.14,cy+2.54,cw-0.28,0.66,
                  font=_FONT,size=13.5,color=_rgb(T.accent),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        txt_body(slide,val,x+0.22,cy+3.3,cw-0.44,ch-3.46,
                 font=_FONT,size=13,color=_rgb(T.text_light),
                 align=PP_ALIGN.CENTER,rtl=True)
    return slide

# ── Stats ────────────────────────────────────────────────────────────────
def make_stats(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"الإحصاءات والأرقام الرئيسية","مؤشرات كمية للدراسة")
    stats=req.stats[:6]
    if not stats: return slide
    n=len(stats); cols=3 if n>=3 else n
    rows=(n+cols-1)//cols; gap=0.3
    cy=3.14; avail=H-cy-0.4
    cw=(_W()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,4.8)
    for i,st in enumerate(stats):
        ci=i%cols; ri=i//cols
        x=0.8+ci*(cw+gap); y=cy+ri*(ch+gap)
        rect(slide,x,y,cw,ch,_rgb(T.card))
        # شريط أكسنت علوي
        top=rect(slide,x,y,cw,0.52,_rgb(T.accent))
        if top: gradient_fill(top,T.accent_grad1,T.accent_grad2,0)
        vs=44 if len(st.value)<=4 else 32 if len(st.value)<=8 else 22
        txt_stat(slide,st.value,x+0.2,y+0.6,cw-0.4,ch*0.46,
                 font="Calibri",color=_rgb(T.accent),align=PP_ALIGN.CENTER)
        if st.unit:
            txt_label(slide,st.unit,x+0.2,y+ch*0.56,cw-0.4,0.55,
                      font=_FONT,size=11.5,color=_rgb(T.muted),
                      align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        hline(slide,x+0.35,y+ch-1.12,cw-0.7,_rgb(T.accent),thickness=0.06)
        txt_label(slide,st.label,x+0.15,y+ch-1.0,cw-0.3,0.9,
                  font=_FONT,size=12,color=_rgb(T.text_light),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
    return slide

# ── Results ──────────────────────────────────────────────────────────────
def make_results(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"نتائج البحث","أبرز ما توصلت إليه الدراسة")
    results=req.main_results[:8]
    if not results: return slide
    cy=3.14; avail=H-cy-0.4
    ih=min(avail/max(len(results),1)-0.12,1.65)
    for i,res in enumerate(results):
        y=cy+i*(ih+0.12)
        row=rrect(slide,0.8,y,_W(),ih,
                  _rgb(T.bg2) if i%2==0 else _rgb(T.card),radius_pct=6)
        if row: shadow(row,blur=5,dist=1,alpha=0.15)
        _pill_accent(slide,T,W-3.2,y+(ih-0.54)/2,2.1,0.54,str(i+1),12)
        vline(slide,0.8,y,ih,_rgb(T.accent),thickness=0.34)
        txt_body(slide,res,1.38,y+0.12,W-5.42,ih-0.24,
                 font=_FONT,size=13.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

# ── Conclusion ───────────────────────────────────────────────────────────
def make_conclusion(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"خاتمة البحث","الاستنتاج العام")
    cy=3.14; ch=H-cy-0.52
    _card(slide,0.8,cy,_W(),ch,T,radius=12)
    lb=rect(slide,0.8,cy,0.52,ch,_rgb(T.accent))
    if lb: gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
    top=rect(slide,0.8,cy,_W(),0.44,_rgb(T.accent))
    if top: gradient_fill(top,T.accent_grad1,T.accent_grad2,0)
    txt(slide,"❝",1.5,cy+0.55,2.2,2.0,
        font="Calibri",size=60,bold=True,
        color=_rgb(T.accent),align=PP_ALIGN.LEFT,txt_shadow=True)
    txt_quote(slide,req.general_conclusion,
              4.5,cy+0.85,_W()-3.8,ch-1.28,
              font=_FONT,size=15,color=_rgb(T.text_light),rtl=True)
    return slide

# ── Recommendations ──────────────────────────────────────────────────────
def make_recommendations(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"توصيات البحث","المقترحات العملية")
    recs=req.recommendations[:8]
    if not recs: return slide
    cy=3.14; avail=H-cy-0.4
    ih=min(avail/max(len(recs),1)-0.12,1.55)
    for i,rec in enumerate(recs):
        y=cy+i*(ih+0.12)
        rrect(slide,0.8,y,_W(),ih,
              _rgb(T.bg2) if i%2==0 else _rgb(T.card),radius_pct=6)
        tg=rect(slide,0.8,y,0.44,ih,_rgb(T.accent))
        if tg: gradient_fill(tg,T.accent_grad1,T.accent_grad2,90)
        txt_body(slide,rec,1.48,y+0.12,_W()-2.08,ih-0.24,
                 font=_FONT,size=13.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

# ── Future ────────────────────────────────────────────────────────────────
def make_future(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"آفاق البحث المستقبلية","مسارات الاستكشاف القادمة")
    items=req.future_work[:6]
    if not items: return slide
    cols=3 if len(items)>=4 else 2 if len(items)>=2 else 1
    rows=(len(items)+cols-1)//cols; gap=0.28
    cy=3.14; avail=H-cy-0.4
    cw=(_W()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,3.35)
    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=0.8+ci*(cw+gap); y=cy+ri*(ch+gap)
        rect(slide,x,y,cw,ch,_rgb(T.card))
        bot=rect(slide,x,y+ch-0.42,cw,0.42,_rgb(T.accent))
        if bot: gradient_fill(bot,T.accent_grad1,T.accent_grad2,0)
        txt(slide,f"{i+1:02d}",x+0.18,y+0.14,1.3,0.95,
            font="Calibri",size=26,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.LEFT,
            letter_spacing=-1.0,txt_shadow=True)
        hline(slide,x+0.2,y+1.15,cw-0.4,_rgb(T.muted),thickness=0.05)
        txt_body(slide,item,x+0.22,y+1.28,cw-0.44,ch-1.82,
                 font=_FONT,size=13,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

# ── References ───────────────────────────────────────────────────────────
def make_references(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"المراجع والمصادر","قائمة المراجع المعتمدة")
    refs=req.references[:14]
    if not refs: return slide
    cy=3.14; avail=H-cy-0.36
    ih=max(min(avail/max(len(refs),1)-0.08,1.05),0.5)
    for i,ref in enumerate(refs):
        y=cy+i*(ih+0.08)
        if y+ih>H-0.26: break
        if i%2==0: rrect(slide,0.8,y,_W(),ih,_rgb(T.bg2),radius_pct=4)
        _pill_accent(slide,T,W-3.15,y+(ih-0.44)/2,1.95,0.44,f"[{i+1}]",9)
        txt_body(slide,ref,1.2,y+0.05,_W()-2.92,ih-0.1,
                 font=_FONT,size=11,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

# ── Final ─────────────────────────────────────────────────────────────────
def make_final(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,150)
    oval(slide,W-5,-3,9,9,_rgb(T.accent),alpha=6)
    oval(slide,-1,H-4,7,7,_rgb(T.accent2),alpha=5)
    top=rect(slide,0,0,W,0.55,_rgb(T.accent))
    if top: gradient_fill(top,T.accent_grad1,T.accent_grad2,0)
    bot=rect(slide,0,H-0.55,W,0.55,_rgb(T.accent))
    if bot: gradient_fill(bot,T.accent_grad2,T.accent_grad1,0)
    # البطاقة المركزية — solid fill فقط
    cw=28.4; ch=13.65; cx=(W-cw)/2; cy=(H-ch)/2
    mc=rrect(slide,cx,cy,cw,ch,_rgb(T.card),radius_pct=14)
    if mc: shadow(mc,blur=32,dist=12,alpha=0.58)
    ct=rect(slide,cx,cy,cw,0.58,_rgb(T.accent))
    if ct: gradient_fill(ct,T.accent_grad1,T.accent_grad2,0)
    txt_hero(slide,"شكراً وتقديراً",cx+1,cy+0.72,cw-2,3.3,
             font=_FONT,size=48,color=_rgb(T.text_light),
             align=PP_ALIGN.CENTER,rtl=True,shadow_on=True)
    for j in range(5):
        dx=cx+cw/2-1.5+j*0.75; sz=0.44 if j==2 else 0.26
        dc=oval(slide,dx,cy+4.18,sz,sz,_rgb(T.accent))
        if dc:
            if j==2: glow(dc,T.accent,radius_pt=5,alpha=0.55)
            else: set_solid_alpha(dc,48)
    txt_hero(slide,req.student_name,cx+1,cy+4.8,cw-2,1.32,
             font=_FONT,size=22,color=_rgb(T.accent),
             align=PP_ALIGN.CENTER,rtl=True,shadow_on=False)
    short=req.title_ar[:88]+("..." if len(req.title_ar)>88 else "")
    txt_body(slide,short,cx+1.8,cy+6.32,cw-3.6,2.8,
             font=_FONT,size=13,color=_rgb(T.muted),
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
