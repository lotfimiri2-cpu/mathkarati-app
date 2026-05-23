"""PREMIUM Engine v19.2 — Corporate Luxury ★ VISUAL FIX"""
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
def _W(): return W-2.0


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

def _glass(slide, T, x, y, w, h, radius=12):
    """بطاقة glassmorphism — solid fill نظيف، بدون gradient داخلي"""
    c = rrect(slide, x, y, w, h, _rgb(T.card), radius_pct=radius)
    if c: shadow(c, blur=6, dist=2, alpha=0.25)
    return c

def _pill(slide, T, x, y, w, h, text, fsize=10):
    p = rrect(slide, x, y, w, h, _rgb(T.accent), radius_pct=50)
    if p:
        gradient_fill(p, T.accent_grad1, T.accent_grad2, 0)
        shadow(p, blur=5, dist=1, alpha=0.25)
    txt_label(slide, text, x, y, w, h, font=_FONT, size=fsize,
              color=_rgb(T.text_dark), align=PP_ALIGN.CENTER,
              rtl=True, uppercase=False)

def _hdr(slide, T, title, sub=""):
    top = rect(slide, 0, 0, W, 0.26, _rgb(T.accent))
    if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)
    gradient_rect(slide, 0, 0.26, W, 2.88, T.grad1, T.grad2, 90)
    # زخارف صغيرة فقط
    od = oval(slide, W-1.78, 0.44, 0.6, 0.6, _rgb(T.accent))
    if od: glow(od, T.accent, radius_pt=6, alpha=0.5)
    oval(slide, W-4.0, 1.22, 0.3, 0.3, _rgb(T.accent2), alpha=58)
    sz = 28 if len(title)<18 else 22 if len(title)<30 else 17
    txt_hero(slide, title, 1.0, 0.28, W-2.4, 2.05,
             font=_FONT, size=sz, color=_rgb(T.text_light),
             align=PP_ALIGN.RIGHT, rtl=True, shadow_on=True)
    if sub:
        txt_label(slide, sub, 1.0, 2.22, W-2.4, 0.82,
                  font=_FONT, size=11.5, color=_rgb(T.muted),
                  align=PP_ALIGN.RIGHT, rtl=True, uppercase=False)
    hline(slide, 0, 3.0, W, _rgb(T.accent), thickness=0.06)
    oc = oval(slide, W*0.5-0.3, 2.88, 0.6, 0.6, _rgb(T.accent))
    if oc: glow(oc, T.accent, radius_pt=5, alpha=0.48)

# ══════════════════════════════════════════════════════════════════
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
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 145)
    # شرائط جانبية مزخرفة
    rect_l=rect(slide, 0, 0, 0.55, H, _rgb(T.accent))
    if rect_l: gradient_fill(rect_l, T.accent_grad1, T.accent_grad2, 90)
    rect_r=rect(slide, W-0.55, 0, 0.55, H, _rgb(T.accent))
    if rect_r: gradient_fill(rect_r, T.accent_grad2, T.accent_grad1, 90)
    # شريط علوي
    top = rect(slide, 0, 0, W, 0.26, _rgb(T.accent))
    if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)
    # المؤسسة — pill
    if req.institution:
        inst = req.institution.split(' — ')[0]
        iw2 = min(len(inst)*0.22+3.0, W-4)
        ip = rrect(slide, (W-iw2)/2, 0.48, iw2, 0.65, _rgb(T.accent), radius_pct=50)
        if ip:
            gradient_fill(ip, T.accent_grad1, T.accent_grad2, 0)
        txt_label(slide, inst, (W-iw2)/2, 0.48, iw2, 0.65,
                  font=_FONT, size=10, color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER, rtl=True, uppercase=False)
    # ── خط أكسنت ─────────────────────────────────────────
    hline(slide, 0.85, 1.55, W-1.7, _rgb(T.accent), thickness=0.16)
    lp = oval(slide, 0.65, 1.47, 0.32, 0.32, _rgb(T.accent))
    if lp: glow(lp, T.accent, radius_pt=5, alpha=0.5)
    # ── العنوان — يملأ المنتصف ────────────────────────────
    _draw_multiline_title(slide, req.title_ar,
        0.88, 1.68, W-1.76, 7.9,
        font=_FONT, color=_rgb(T.text_light),
        align=PP_ALIGN.RIGHT, rtl_on=True)
    if req.title_en:
        txt(slide, req.title_en, 0.88, H*0.54, W-1.76, 0.82,
            font="Calibri", size=11.5, italic=True,
            color=_rgb(T.muted), align=PP_ALIGN.RIGHT, letter_spacing=0.8)
    # ── خط فاصل + بطاقات معلومات ─────────────────────────
    hline(slide, 0.88, H*0.56-0.18, W-1.76, _rgb(T.accent), thickness=0.18)
    fields = [("الطالب", req.student_name)]
    if req.supervisor:     fields.append(("المشرف", req.supervisor))
    if req.co_supervisor:  fields.append(("المشرف المساعد", req.co_supervisor))
    if req.specialization: fields.append(("التخصص", req.specialization))
    n = len(fields)
    cw = min((_W()-0.22*(n-1))/n, 8.8)
    tot = n*cw+0.22*(n-1); sx = 1.0+(_W()-tot)/2
    for i,(lbl,val) in enumerate(fields[:4]):
        px = sx+i*(cw+0.22); py = H*0.57
        gc = rect(slide, px, py, cw, 2.35, _rgb(T.card))
        lb = rrect(slide, px, py, cw, 0.54, _rgb(T.accent), radius_pct=5)
        if lb: gradient_fill(lb, T.accent_grad1, T.accent_grad2, 0)
        txt_label(slide, lbl, px+0.1, py+0.04, cw-0.2, 0.48,
                  font=_FONT, size=10, color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER, rtl=True, uppercase=False)
        txt_body(slide, val, px+0.15, py+0.62, cw-0.3, 1.62,
                 font=_FONT, size=13, color=_rgb(T.text_light),
                 align=PP_ALIGN.CENTER, rtl=True)
    if req.year:
        yr = rrect(slide, W/2-2.8, H-1.38, 5.6, 0.8, _rgb(T.accent), radius_pct=50)
        if yr:
            gradient_fill(yr, T.accent_grad1, T.accent_grad2, 0)
            glow(yr, T.accent, radius_pt=5, alpha=0.2)
        txt_label(slide, req.year, W/2-2.8, H-1.38, 5.6, 0.8,
                  font="Calibri", size=13, color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER, rtl=False, uppercase=False)
    return slide


def make_intro(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"مقدمة البحث","الإطار العام والمقاربة المنهجية")
    cy=3.08; ch=H-cy-0.42
    if req.intro_overview and req.intro_approach:
        gap=0.4; cw1=_W()*0.56; cw2=_W()-cw1-gap
        for titl,text,cw,cx in [
            ("نظرة عامة",req.intro_overview,cw1,1.0),
            ("المقاربة",req.intro_approach,cw2,1.0+cw1+gap)]:
            c=rect(slide,cx,cy,cw,ch,_rgb(T.card))
            if c: shadow(c,blur=4,dist=1,alpha=0.18)
            tp=rect(slide,cx,cy,cw,0.66,_rgb(T.accent))
            if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
            txt_label(slide,titl,cx+0.3,cy+0.07,cw-0.6,0.56,
                      font=_FONT,size=15,color=_rgb(T.text_dark),
                      align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
            txt_body(slide,text,cx+0.3,cy+0.82,cw-0.6,ch-1.0,
                     font=_FONT,size=14,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
    else:
        text=req.intro_overview or req.intro_approach
        c=rect(slide,1.0,cy,_W(),ch,_rgb(T.card))
        if c: shadow(c,blur=4,dist=1,alpha=0.18)
        lb=rrect(slide,1.0,cy,0.52,ch,_rgb(T.accent),radius_pct=13)
        if lb: gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
        txt(slide,"❝",1.62,cy+0.3,2.0,1.8,
            font="Calibri",size=58,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.LEFT,txt_shadow=True)
        txt_quote(slide,text,1.7,cy+0.72,_W()-1.15,ch-0.95,
                  font=_FONT,size=15,color=_rgb(T.text_light),rtl=True)
    return slide

def make_plan(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"خطة البحث","هيكل ومحتويات الدراسة")
    chapters=req.chapters[:8]
    if not chapters: return slide
    n=len(chapters)
    cy=3.08; avail=H-cy-0.42
    if n<=3: cols=n; rows=1
    elif n<=6: cols=3; rows=(n+2)//3
    else: cols=4; rows=(n+3)//4
    gap=0.3
    cw=(_W()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,4.25)
    for i,chap in enumerate(chapters):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(cw+gap); y=cy+ri*(ch+gap)
        _glass(slide,T,x,y,cw,ch,radius=12)
        tp=rect(slide,x,y,cw,0.62,_rgb(T.accent))
        if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
        txt(slide,f"{i+1:02d}",x+0.14,y+0.06,1.1,0.52,
            font="Calibri",size=20,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.LEFT,letter_spacing=-1.0)
        hline(slide,x+0.2,y+0.76,cw-0.4,_rgb(T.muted),thickness=0.04)
        txt_body(slide,chap.title,x+0.2,y+0.88,cw-0.4,ch-1.3,
                 font=_FONT,size=14.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
        if chap.pages:
            txt_label(slide,f"ص {chap.pages}",x+0.18,y+ch-0.52,cw-0.36,0.44,
                      font="Calibri",size=11,color=_rgb(T.muted),
                      align=PP_ALIGN.LEFT,rtl=False,uppercase=False)
    return slide

def make_problem(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"إشكالية البحث","التساؤلات المحورية للدراسة")
    cy=3.08
    if req.main_problem or req.main_question:
        text=req.main_problem or req.main_question
        qh=min(3.1,H*0.31)
        _glass(slide,T,1.0,cy,_W(),qh,radius=12)
        hline(slide,1.0,cy+qh-0.14,_W(),_rgb(T.accent),thickness=0.14)
        txt(slide,"❝",1.45,cy+0.18,2.0,qh-0.36,
            font="Calibri",size=54,bold=True,
            color=_rgb(T.accent),align=PP_ALIGN.LEFT,txt_shadow=True)
        txt_quote(slide,text,1.52,cy+0.26,_W()-0.85,qh-0.52,
                  font=_FONT,size=15,color=_rgb(T.text_light),rtl=True)
        cy+=qh+0.26
    if req.sub_questions:
        sq=req.sub_questions[:6]; n2=len(sq)
        cols=2 if n2>3 else 1
        cw=(_W()-(cols-1)*0.28)/cols
        avail=H-cy-0.4
        sh=min(avail/((n2+cols-1)//cols)-0.1,1.22)
        for i,q in enumerate(sq):
            ci=i%cols; ri=i//cols
            x=1.0+ci*(cw+0.28); y=cy+ri*(sh+0.1)
            _glass(slide,T,x,y,cw,sh,radius=8)
            _pill(slide,T,x+cw-1.72,y+(sh-0.5)/2,1.58,0.5,str(i+1),11)
            vline(slide,x,y,sh,_rgb(T.accent),thickness=0.36)
            txt_body(slide,q,x+0.55,y+0.1,cw-2.38,sh-0.2,
                     font=_FONT,size=13.5,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_objectives(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"أهداف البحث وفرضياته","")
    cy=3.08; ch=H-cy-0.42; gap=0.4
    cols=[]
    if req.objectives: cols.append(("الأهداف",req.objectives))
    if req.hypotheses:  cols.append(("الفرضيات",req.hypotheses))
    if not cols: return slide
    widths=[_W()*0.56,_W()*0.42] if len(cols)==2 else [_W()]
    x=1.0
    for i,(lbl,items) in enumerate(cols[:2]):
        cw=widths[i]
        rect(slide,x,cy,cw,ch,_rgb(T.card))
        hd=rect(slide,x,cy,cw,0.66,_rgb(T.accent))
        if hd: gradient_fill(hd,T.accent_grad1,T.accent_grad2,0)
        txt_label(slide,lbl,x+0.22,cy+0.07,cw-0.44,0.56,
                  font=_FONT,size=15,color=_rgb(T.text_dark),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        ih=min((ch-0.75)/max(len(items),1),1.22)
        for j,item in enumerate(items[:8]):
            iy=cy+0.75+j*ih
            if iy+ih>cy+ch-0.08: break
            if j>0: hline(slide,x+0.22,iy,cw-0.44,_rgb(T.bg),thickness=0.03)
            _pill(slide,T,x+cw-1.5,iy+(ih-0.48)/2,1.36,0.48,str(j+1),10)
            txt_body(slide,item,x+0.22,iy+0.09,cw-1.92,ih-0.18,
                     font=_FONT,size=13,color=_rgb(T.text_light),
                     align=PP_ALIGN.RIGHT,rtl=True)
        x+=cw+gap
    return slide

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
    cy=3.08; avail=H-cy-0.42
    cw=(_W()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,3.85)
    accent_h=[0.65,0.82,0.56,0.74,0.62,0.78]
    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(cw+gap); y=cy+ri*(ch+gap)
        rect(slide,x,y,cw,ch,_rgb(T.card))
        ah=accent_h[i%len(accent_h)]
        tp=rect(slide,x,y,cw,ah,_rgb(T.accent))
        if tp: gradient_fill(tp,T.accent_grad1,T.accent_grad2,0)
        txt(slide,f"{i+1:02d}",x+0.16,y+0.04,1.1,ah-0.08,
            font="Calibri",size=22,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.LEFT,letter_spacing=-1.0)
        txt_body(slide,item,x+0.22,y+ah+0.18,cw-0.44,ch-ah-0.3,
                 font=_FONT,size=13,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

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
    n=len(fields); cy=3.08; ch=H-cy-0.46
    cw=(_W()-0.28*(n-1))/n
    if n>1: hline(slide,1.0+cw/2,cy+ch*0.5,_W()-cw,_rgb(T.muted),thickness=0.04)
    for i,(lbl,icon,val) in enumerate(fields[:4]):
        x=1.0+i*(cw+0.28)
        rect(slide,x,cy,cw,ch,_rgb(T.card))
        sz=1.72; ix=x+(cw-sz)/2
        ic=oval(slide,ix,cy+0.5,sz,sz,_rgb(T.accent))
        if ic:
            gradient_fill(ic,T.accent_grad1,T.accent_grad2,45)
            shadow_and_glow(ic,s_blur=12,s_dist=4,s_alpha=0.42,
                            g_color=T.accent_grad2,g_rad=6,g_alpha=0.25)
        txt(slide,icon,ix,cy+0.5,sz,sz,
            font=_FONT,size=26,bold=True,
            color=_rgb(T.text_dark),align=PP_ALIGN.CENTER)
        hline(slide,x+0.28,cy+2.42,cw-0.56,_rgb(T.accent),thickness=0.08)
        txt_label(slide,lbl,x+0.14,cy+2.54,cw-0.28,0.68,
                  font=_FONT,size=13.5,color=_rgb(T.accent),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        txt_body(slide,val,x+0.22,cy+3.3,cw-0.44,ch-3.46,
                 font=_FONT,size=13,color=_rgb(T.text_light),
                 align=PP_ALIGN.CENTER,rtl=True)
    return slide

def make_stats(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"الإحصاءات والأرقام الرئيسية","مؤشرات كمية للدراسة")
    stats=req.stats[:6]
    if not stats: return slide
    n=len(stats); cols=3 if n>=3 else n
    rows=(n+cols-1)//cols; gap=0.3
    cy=3.08; avail=H-cy-0.42
    cw=(_W()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,4.85)
    for i,st in enumerate(stats):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(cw+gap); y=cy+ri*(ch+gap)
        rect(slide,x,y,cw,ch,_rgb(T.card))
        top=rect(slide,x,y,cw,0.54,_rgb(T.accent))
        if top: gradient_fill(top,T.accent_grad1,T.accent_grad2,0)
        txt_stat(slide,st.value,x+0.2,y+0.62,cw-0.4,ch*0.46,
                 font="Calibri",color=_rgb(T.accent),align=PP_ALIGN.CENTER)
        if st.unit:
            txt_label(slide,st.unit,x+0.2,y+ch*0.57,cw-0.4,0.58,
                      font=_FONT,size=11.5,color=_rgb(T.muted),
                      align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
        hline(slide,x+0.35,y+ch-1.12,cw-0.7,_rgb(T.accent),thickness=0.07)
        txt_label(slide,st.label,x+0.15,y+ch-1.0,cw-0.3,0.9,
                  font=_FONT,size=12,color=_rgb(T.text_light),
                  align=PP_ALIGN.CENTER,rtl=True,uppercase=False)
    return slide

def make_results(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"نتائج البحث","أبرز ما توصلت إليه الدراسة")
    results=req.main_results[:8]
    if not results: return slide
    avail=H-3.12
    ih=min(avail/max(len(results),1)-0.12,1.65)
    for i,res in enumerate(results):
        y=3.12+i*(ih+0.12)
        row=rrect(slide,1.0,y,_W(),ih,
                  _rgb(T.bg2) if i%2==0 else _rgb(T.card),radius_pct=7)
        if row: shadow(row,blur=6,dist=1,alpha=0.18)
        _pill(slide,T,W-3.28,y+(ih-0.54)/2,2.1,0.54,str(i+1),12)
        vline(slide,1.0,y,ih,_rgb(T.accent),thickness=0.36)
        txt_body(slide,res,1.55,y+0.12,W-5.45,ih-0.24,
                 font=_FONT,size=13.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_conclusion(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"خاتمة البحث","الاستنتاج العام")
    cy=3.08; ch=H-cy-0.52
    _glass(slide,T,1.0,cy,_W(),ch,radius=14)
    lb=rrect(slide,1.0,cy,0.52,ch,_rgb(T.accent),radius_pct=14)
    if lb: gradient_fill(lb,T.accent_grad1,T.accent_grad2,90)
    top=rrect(slide,1.0,cy,_W(),0.44,_rgb(T.accent),radius_pct=14)
    if top: gradient_fill(top,T.accent_grad1,T.accent_grad2,0)
    txt(slide,"❝",1.62,cy+0.58,2.2,2.0,
        font="Calibri",size=62,bold=True,
        color=_rgb(T.accent),align=PP_ALIGN.LEFT,txt_shadow=True)
    txt_quote(slide,req.general_conclusion,
              4.55,cy+0.9,_W()-3.8,ch-1.3,
              font=_FONT,size=15,color=_rgb(T.text_light),rtl=True)
    return slide

def make_recommendations(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"توصيات البحث","المقترحات العملية")
    recs=req.recommendations[:8]
    if not recs: return slide
    avail=H-3.12
    ih=min(avail/max(len(recs),1)-0.12,1.56)
    for i,rec in enumerate(recs):
        y=3.12+i*(ih+0.12)
        row=rrect(slide,1.0,y,_W(),ih,
                  _rgb(T.bg2) if i%2==0 else _rgb(T.card),radius_pct=7)
        tg=rrect(slide,1.0,y,0.46,ih,_rgb(T.accent),radius_pct=7)
        if tg: gradient_fill(tg,T.accent_grad1,T.accent_grad2,90)
        txt_body(slide,rec,1.65,y+0.12,_W()-2.1,ih-0.24,
                 font=_FONT,size=13.5,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_future(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"آفاق البحث المستقبلية","مسارات الاستكشاف القادمة")
    items=req.future_work[:6]
    if not items: return slide
    cols=3 if len(items)>=4 else 2 if len(items)>=2 else 1
    rows=(len(items)+cols-1)//cols; gap=0.28
    cy=3.08; avail=H-cy-0.42
    cw=(_W()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap,3.38)
    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(cw+gap); y=cy+ri*(ch+gap)
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

def make_references(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,180)
    _hdr(slide,T,"المراجع والمصادر","قائمة المراجع المعتمدة")
    refs=req.references[:14]
    if not refs: return slide
    avail=H-3.12
    ih=max(min(avail/max(len(refs),1)-0.08,1.05),0.5)
    for i,ref in enumerate(refs):
        y=3.12+i*(ih+0.08)
        if y+ih>H-0.26: break
        if i%2==0: rrect(slide,1.0,y,_W(),ih,_rgb(T.bg2),radius_pct=5)
        _pill(slide,T,W-3.22,y+(ih-0.44)/2,1.98,0.44,f"[{i+1}]",9)
        txt_body(slide,ref,1.4,y+0.05,_W()-2.9,ih-0.1,
                 font=_FONT,size=11,color=_rgb(T.text_light),
                 align=PP_ALIGN.RIGHT,rtl=True)
    return slide

def make_final(prs, req, T):
    slide=blank_slide(prs)
    bg(slide,_rgb(T.bg))
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,150)
    oval(slide,W-4,-3,8,8,_rgb(T.accent),alpha=6)
    oval(slide,-1,H-4,7,7,_rgb(T.accent2),alpha=5)
    tb=rect(slide,0,0,W,0.55,_rgb(T.accent))
    if tb: gradient_fill(tb,T.accent_grad1,T.accent_grad2,0)
    bb=rect(slide,0,H-0.55,W,0.55,_rgb(T.accent))
    if bb: gradient_fill(bb,T.accent_grad2,T.accent_grad1,0)
    cw=28.5; ch=14.1; cx=(W-cw)/2; cy=(H-ch)/2
    _glass(slide,T,cx,cy,cw,ch,radius=16)
    ct=rrect(slide,cx,cy,cw,0.58,_rgb(T.accent),radius_pct=16)
    if ct: gradient_fill(ct,T.accent_grad1,T.accent_grad2,0)
    txt_hero(slide,"شكراً وتقديراً",cx+1,cy+0.72,cw-2,3.35,
             font=_FONT,size=48,color=_rgb(T.text_light),
             align=PP_ALIGN.CENTER,rtl=True,shadow_on=True)
    for j in range(5):
        dx=cx+cw/2-1.5+j*0.75; sz=0.46 if j==2 else 0.27
        dc=oval(slide,dx,cy+4.22,sz,sz,_rgb(T.accent))
        if dc:
            if j==2: glow(dc,T.accent,radius_pt=6,alpha=0.6)
            else: set_solid_alpha(dc,48)
    txt_hero(slide,req.student_name,cx+1,cy+4.88,cw-2,1.34,
             font=_FONT,size=22,color=_rgb(T.accent),
             align=PP_ALIGN.CENTER,rtl=True,shadow_on=False)
    short=req.title_ar[:88]+("..." if len(req.title_ar)>88 else "")
    txt_body(slide,short,cx+1.8,cy+6.42,cw-3.6,2.85,
             font=_FONT,size=13,color=_rgb(T.muted),
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
