"""
Slide Builder — CLASSIC Engine — مذكرتي Pro v18
Academic Structured Layout
"""
from __future__ import annotations
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, cm, pt, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, set_solid_alpha, txt, blank_slide,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"
def set_font(f): global _FONT; _FONT = f

def _rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[:2],16), int(h[2:4],16), int(h[4:],16))

_SIDE_W  = 7.2
_CONT_X  = _SIDE_W + 1.0
_CONT_W  = W - _CONT_X - 0.8

def _side_panel(slide, T):
    sp = rect(slide, 0, 0, _SIDE_W, H, _rgb(T.accent))
    if sp: gradient_fill(sp, T.accent_grad1, T.accent_grad2, 180)
    oval(slide, -2.5, H*0.55, 9, 9, _rgb(T.bg), alpha=10)
    oval(slide, 0.5, -2, 5.5, 5.5, _rgb(T.bg), alpha=7)
    oval(slide, _SIDE_W-2.5, H*0.2, 4, 4, _rgb(T.bg), alpha=8)

def _header_bar(slide, T, title, subtitle="", num=""):
    hb = rect(slide, _SIDE_W, 0, W-_SIDE_W, 3.6, _rgb(T.bg2))
    if hb: gradient_fill(hb, T.bg, T.bg2, 90)
    hline(slide, _SIDE_W, 3.45, W-_SIDE_W, _rgb(T.accent), thickness=0.15)
    if num:
        nb = rect(slide, _CONT_X, 0.55, 1.4, 1.4, _rgb(T.bg))
        if nb: shadow(nb, blur=8, dist=3, alpha=0.4)
        txt(slide, num, _CONT_X, 0.55, 1.4, 1.4,
            font="Calibri", size=22, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.CENTER)
    tx = _CONT_X + (1.6 if num else 0)
    tw = W - tx - 0.8
    sz = 22 if len(title)<22 else 18
    txt(slide, title, tx, 0.45, tw, 1.8,
        font=_FONT, size=sz, bold=True,
        color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    if subtitle:
        txt(slide, subtitle, tx, 2.1, tw, 1.1,
            font=_FONT, size=11.5, color=_rgb(T.muted),
            align=PP_ALIGN.RIGHT, rtl=True)

def make_cover(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    _side_panel(slide, T)
    if req.institution:
        parts = req.institution.split(" — ")
        for i,part in enumerate(parts[:3]):
            txt(slide, part, 0.3, 1.2+i*1.05, _SIDE_W-0.6, 0.95,
                font=_FONT, size=10, bold=(i==0),
                color=_rgb(T.text_dark), align=PP_ALIGN.CENTER, rtl=True)
    if req.year:
        yr = rect(slide, 0.6, H-1.5, _SIDE_W-1.2, 0.8, _rgb(T.bg))
        if yr: shadow(yr, blur=8, dist=2, alpha=0.3)
        txt(slide, req.year, 0.6, H-1.5, _SIDE_W-1.2, 0.8,
            font="Calibri", size=13, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.CENTER)
    cx = _SIDE_W + 0.8; cw = W - cx - 0.7
    hline(slide, cx, 2.8, cw, _rgb(T.accent), thickness=0.12)
    tsz = 28 if len(req.title_ar)<35 else 22 if len(req.title_ar)<55 else 17 if len(req.title_ar)<80 else 14
    txt(slide, req.title_ar, cx, 3.0, cw, H*0.38,
        font=_FONT, size=tsz, bold=True,
        color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    if req.title_en:
        txt(slide, req.title_en, cx, 3.0+H*0.38, cw, 0.85,
            font="Calibri", size=10.5, italic=True,
            color=_rgb(T.muted), align=PP_ALIGN.LEFT)
    info_y = H*0.62
    hline(slide, cx, info_y-0.2, cw, _rgb(T.accent), thickness=0.05)
    fields = [("الطالب :", req.student_name)]
    if req.supervisor:     fields.append(("المشرف :", req.supervisor))
    if req.co_supervisor:  fields.append(("م. مساعد :", req.co_supervisor))
    if req.specialization: fields.append(("التخصص :", req.specialization))
    row_h = min((H-info_y-0.5)/max(len(fields),1), 1.3)
    for i,(lbl,val) in enumerate(fields[:4]):
        y = info_y+i*row_h
        if i%2==0:
            rb = rect(slide, cx, y, cw, row_h, _rgb(T.bg2))
            if rb: set_solid_alpha(rb, 50)
        txt(slide, lbl, cx+0.2, y+0.05, 4.5, row_h-0.1,
            font=_FONT, size=11, bold=True, color=_rgb(T.accent), align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, val, cx+5.0, y+0.05, cw-5.2, row_h-0.1,
            font=_FONT, size=12, color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

def _content_row(slide, T, y, h, text, num=None, alt=False, font_size=12):
    rect(slide, _CONT_X, y, _CONT_W, h, _rgb(T.bg2) if alt else _rgb(T.card))
    vr = rect(slide, W-1.0, y, 1.0, h, _rgb(T.accent))
    if vr: gradient_fill(vr, T.accent_grad1, T.accent_grad2, 90)
    if num is not None:
        txt(slide, str(num), _CONT_X+0.2, y, 1.4, h,
            font="Calibri", size=18, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.CENTER)
        tx = _CONT_X+1.8; tw = _CONT_W-3.0
    else:
        tx = _CONT_X+0.4; tw = _CONT_W-1.6
    txt(slide, text, tx, y+0.07, tw, h-0.14,
        font=_FONT, size=font_size, color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)

def make_intro(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    _side_panel(slide, T)
    _header_bar(slide, T, "مقدمة البحث", "نظرة عامة وأسلوب المعالجة", "01")
    items=[]
    if req.intro_overview: items.append(("نظرة عامة", req.intro_overview))
    if req.intro_approach:  items.append(("المنهج المتبع", req.intro_approach))
    y=3.75
    for lbl,val in items:
        card_h=(H-y-0.5)/max(len(items),1)-0.25
        ch=rect(slide, _CONT_X, y, _CONT_W, 0.7, _rgb(T.accent))
        if ch: gradient_fill(ch, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, _CONT_X+0.3, y+0.05, _CONT_W-0.6, 0.65,
            font=_FONT, size=13, bold=True, color=_rgb(T.text_dark), align=PP_ALIGN.RIGHT, rtl=True)
        rect(slide, _CONT_X, y+0.7, _CONT_W, card_h, _rgb(T.card))
        vline(slide, _CONT_X, y+0.7, card_h, _rgb(T.accent), thickness=0.35)
        txt(slide, val, _CONT_X+0.7, y+0.85, _CONT_W-1.1, card_h-0.3,
            font=_FONT, size=12, spacing=19, color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
        y += card_h+0.95+0.25
    return slide

def make_plan(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    _side_panel(slide, T)
    _header_bar(slide, T, "خطة البحث", f"يتكون البحث من {len(req.chapters)} فصول", "02")
    chs=req.chapters[:9]
    if not chs: return slide
    avail=H-3.9; ih=min(avail/max(len(chs),1)-0.1, 1.7)
    for i,ch in enumerate(chs):
        y=3.8+i*(ih+0.1)
        _content_row(slide, T, y, ih, ch.title, num=i+1, alt=(i%2==0), font_size=12)
        if ch.pages:
            txt(slide, ch.pages, _CONT_X+1.8, y, 3.5, ih,
                font="Calibri", size=9.5, color=_rgb(T.muted), align=PP_ALIGN.LEFT)
    return slide

def make_problem(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    _side_panel(slide, T)
    _header_bar(slide, T, "إشكالية البحث", "التساؤلات الرئيسية والفرعية", "03")
    y=3.75
    if req.main_problem:
        h=2.5
        card=rect(slide, _CONT_X, y, _CONT_W, h, _rgb(T.card))
        if card: shadow(card, blur=12, dist=4, alpha=0.35)
        vline(slide, _CONT_X, y, h, _rgb(T.accent), thickness=0.4)
        txt(slide, "الإشكالية الرئيسية", W-4.5, y+0.1, 3.5, 0.6,
            font=_FONT, size=10, bold=True, color=_rgb(T.accent), align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, req.main_problem, _CONT_X+0.7, y+0.7, _CONT_W-1.1, h-0.85,
            font=_FONT, size=13, spacing=20, color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
        y += h+0.3
    if req.main_question:
        qh=1.35
        rect(slide, _CONT_X, y, _CONT_W, qh, _rgb(T.bg2))
        hline(slide, _CONT_X, y+qh-0.08, _CONT_W, _rgb(T.accent), thickness=0.08)
        txt(slide, "\u275d  "+req.main_question, _CONT_X+0.4, y+0.12, _CONT_W-0.8, qh-0.24,
            font=_FONT, size=13, bold=True, italic=True, color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
        y += qh+0.25
    if req.sub_questions:
        avail=H-y-0.4; sq_h=min(avail/max(len(req.sub_questions),1), 0.88)
        for i,q in enumerate(req.sub_questions[:6]):
            qy=y+i*sq_h
            nb=rect(slide, W-1.6, qy+sq_h*0.15, 0.58, 0.58, _rgb(T.accent))
            txt(slide, str(i+1), W-1.6, qy+sq_h*0.15, 0.58, 0.58,
                font="Calibri", size=9, bold=True, color=_rgb(T.text_dark), align=PP_ALIGN.CENTER)
            txt(slide, q, _CONT_X+0.3, qy+0.05, _CONT_W-2.0, sq_h-0.1,
                font=_FONT, size=11.5, color=_rgb(T.muted), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

def make_objectives(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    _side_panel(slide, T)
    _header_bar(slide, T, "أهداف البحث وفرضياته", "", "04")
    cols_data=[]
    if req.objectives: cols_data.append(("الأهداف", req.objectives))
    if req.hypotheses:  cols_data.append(("الفرضيات", req.hypotheses))
    if not cols_data: return slide
    n_cols=len(cols_data)
    cw=(_CONT_W-(n_cols-1)*0.35)/n_cols; ph=H-4.1
    for i,(lbl,items) in enumerate(cols_data[:2]):
        x=_CONT_X+i*(cw+0.35)
        ch=rect(slide, x, 3.75, cw, 0.75, _rgb(T.accent))
        if ch: gradient_fill(ch, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, x+0.2, 3.78, cw-0.4, 0.7,
            font=_FONT, size=14, bold=True, color=_rgb(T.text_dark), align=PP_ALIGN.CENTER, rtl=True)
        rect(slide, x, 4.5, cw, ph-0.75, _rgb(T.card))
        avail_h=ph-0.85; ih=min(avail_h/max(len(items),1), 1.1)
        for j,item in enumerate(items[:8]):
            iy=4.55+j*ih
            if j>0: hline(slide, x+0.2, iy, cw-0.4, _rgb(T.bg), thickness=0.04)
            nb=rect(slide, x+cw-0.75, iy+ih*0.22, 0.55, 0.55, _rgb(T.accent))
            txt(slide, str(j+1), x+cw-0.75, iy+ih*0.22, 0.55, 0.55,
                font="Calibri", size=9, bold=True, color=_rgb(T.text_dark), align=PP_ALIGN.CENTER)
            txt(slide, item, x+0.25, iy+0.07, cw-1.1, ih-0.14,
                font=_FONT, size=11, color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# IMPORTANCE
# ══════════════════════════════════════════════════════════════════════
def make_importance(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    _side_panel(slide, T)
    _header_bar(slide, T, "أهمية البحث", "مبررات اختيار الموضوع", "05")

    items = list(req.importance or [])
    if req.reasons and req.reasons not in items: items.append(req.reasons)
    items = items[:6]
    if not items: return slide

    avail = H - 3.9
    ih = min(avail/max(len(items),1) - 0.1, 1.8)

    for i,item in enumerate(items):
        y = 3.8 + i*(ih+0.1)
        # خلفية الصف مع تناوب
        rb = rect(slide, _CONT_X, y, _CONT_W, ih,
                  _rgb(T.card) if i%2==0 else _rgb(T.bg2))
        if rb and i%2==0: shadow(rb, blur=5, dist=1, alpha=0.2)
        # شريط رقم يميني
        nr = rect(slide, W-1.8, y, 1.8, ih, _rgb(T.accent))
        if nr: gradient_fill(nr, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, f"{i+1:02d}", W-1.8, y, 1.8, ih,
            font="Calibri", size=16, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.CENTER)
        # شريط لوني أيسر
        vline(slide, _CONT_X, y, ih, _rgb(T.accent), thickness=0.3)
        txt(slide, item, _CONT_X+0.6, y+0.1, _CONT_W-2.6, ih-0.2,
            font=_FONT, size=12, spacing=18,
            color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# METHODOLOGY
# ══════════════════════════════════════════════════════════════════════
def make_methodology(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    _side_panel(slide, T)
    _header_bar(slide, T, "منهجية البحث", "الإجراءات والأدوات", "06")

    fields = []
    if req.methodology:  fields.append(("المنهج",      req.methodology))
    if req.sample_type:  fields.append(("نوع العينة",  req.sample_type))
    if req.sample_size:  fields.append(("حجم العينة",  req.sample_size))
    if req.tool:         fields.append(("الأداة",       req.tool))
    if not fields: return slide

    cols = 2 if len(fields)>=3 else 1
    rows = (len(fields)+cols-1)//cols
    gap  = 0.3
    cw   = (_CONT_W - (cols-1)*gap) / cols
    avail= H - 3.9
    ch   = min(avail/rows - gap, 3.5)

    for i,(lbl,val) in enumerate(fields[:4]):
        ci = i%cols; ri = i//cols
        x  = _CONT_X + ci*(cw+gap)
        y  = 3.8 + ri*(ch+gap)

        card = rect(slide, x, y, cw, ch, _rgb(T.card))
        if card: shadow(card, blur=10, dist=3, alpha=0.3)
        # دائرة رقم كبيرة
        nc = oval(slide, x+0.3, y+(ch-1.2)/2, 1.2, 1.2, _rgb(T.accent))
        if nc: gradient_fill(nc, T.accent_grad1, T.accent_grad2, 45)
        txt(slide, str(i+1), x+0.3, y+(ch-1.2)/2, 1.2, 1.2,
            font="Calibri", size=18, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.CENTER)
        # عنوان + محتوى
        hline(slide, x+1.8, y+0.4, cw-2.1, _rgb(T.accent), thickness=0.06)
        txt(slide, lbl, x+1.8, y+0.12, cw-2.1, 0.65,
            font=_FONT, size=13, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, val, x+1.8, y+0.72, cw-2.1, ch-0.85,
            font=_FONT, size=11.5, spacing=18,
            color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# STATS
# ══════════════════════════════════════════════════════════════════════
def make_stats(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    _side_panel(slide, T)
    _header_bar(slide, T, "الإحصاءات والأرقام الرئيسية", "مؤشرات كمية للدراسة", "07")

    stats = req.stats[:6]
    if not stats: return slide
    n=len(stats)
    cols=3 if n>=3 else n
    rows=(n+cols-1)//cols
    gap=0.28
    cw=(_CONT_W-(cols-1)*gap)/cols
    avail=H-3.9
    ch=min(avail/rows-gap, 4.5)

    for i,st in enumerate(stats):
        ci=i%cols; ri=i//cols
        x=_CONT_X+ci*(cw+gap)
        y=3.8+ri*(ch+gap)
        card=rect(slide, x, y, cw, ch, _rgb(T.card))
        if card: shadow(card, blur=12, dist=4, alpha=0.35)
        # شريط سفلي
        bot=rect(slide, x, y+ch-0.4, cw, 0.4, _rgb(T.accent))
        if bot: gradient_fill(bot, T.accent_grad1, T.accent_grad2, 0)
        # القيمة الكبيرة
        vs=36 if len(st.value)<=4 else 26 if len(st.value)<=8 else 20
        txt(slide, st.value, x+0.2, y+0.35, cw-0.4, ch*0.5,
            font="Calibri", size=vs, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.CENTER)
        if st.unit:
            txt(slide, st.unit, x+0.2, y+ch*0.5+0.3, cw-0.4, 0.55,
                font=_FONT, size=10, color=_rgb(T.muted),
                align=PP_ALIGN.CENTER, rtl=True)
        txt(slide, st.label, x+0.2, y+ch-1.0, cw-0.4, 0.65,
            font=_FONT, size=11, color=_rgb(T.text_light),
            align=PP_ALIGN.CENTER, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# RESULTS
# ══════════════════════════════════════════════════════════════════════
def make_results(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    _side_panel(slide, T)
    _header_bar(slide, T, "نتائج البحث", "أبرز ما توصلت إليه الدراسة", "08")

    results=req.main_results[:8]
    if not results: return slide
    avail=H-3.9
    ih=min(avail/max(len(results),1)-0.1, 1.55)

    for i,res in enumerate(results):
        y=3.8+i*(ih+0.1)
        # صف بتناوب
        rb=rect(slide, _CONT_X, y, _CONT_W, ih,
                _rgb(T.bg2) if i%2==0 else _rgb(T.card))
        # شريط رقم يميني سميك
        nr=rect(slide, W-1.6, y, 1.6, ih, _rgb(T.accent))
        if nr: gradient_fill(nr, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, str(i+1), W-1.6, y, 1.6, ih,
            font="Calibri", size=16, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.CENTER)
        txt(slide, res, _CONT_X+0.4, y+0.1, _CONT_W-2.2, ih-0.2,
            font=_FONT, size=12, color=_rgb(T.text_light),
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# CONCLUSION
# ══════════════════════════════════════════════════════════════════════
def make_conclusion(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    _side_panel(slide, T)
    _header_bar(slide, T, "خاتمة البحث", "الاستنتاج العام", "09")

    cy=3.75; ch=H-cy-0.6
    card=rect(slide, _CONT_X, cy, _CONT_W, ch, _rgb(T.card))
    if card: shadow(card, blur=16, dist=5, alpha=0.4)
    # شريط أيسر سميك
    vline(slide, _CONT_X, cy, ch, _rgb(T.accent), thickness=0.5)
    # علامة اقتباس
    txt(slide, "❝", _CONT_X+0.7, cy+0.2, 2.5, 2.0,
        font="Calibri", size=56, color=_rgb(T.accent), align=PP_ALIGN.LEFT)
    txt(slide, req.general_conclusion,
        _CONT_X+0.9, cy+1.4, _CONT_W-1.6, ch-1.8,
        font=_FONT, size=14, spacing=22,
        color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# RECOMMENDATIONS
# ══════════════════════════════════════════════════════════════════════
def make_recommendations(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    _side_panel(slide, T)
    _header_bar(slide, T, "توصيات البحث", "المقترحات العملية", "10")

    recs=req.recommendations[:8]
    if not recs: return slide
    avail=H-3.9
    ih=min(avail/max(len(recs),1)-0.1, 1.5)

    for i,rec in enumerate(recs):
        y=3.8+i*(ih+0.1)
        rb=rect(slide, _CONT_X, y, _CONT_W, ih,
                _rgb(T.card) if i%2==0 else _rgb(T.bg2))
        # ◆ يميني
        txt(slide, "◆", W-1.5, y, 1.2, ih,
            font="Calibri", size=13, color=_rgb(T.accent),
            align=PP_ALIGN.CENTER)
        # شريط أيسر
        vline(slide, _CONT_X, y, ih, _rgb(T.accent), thickness=0.3)
        txt(slide, rec, _CONT_X+0.55, y+0.1, _CONT_W-1.9, ih-0.2,
            font=_FONT, size=12, color=_rgb(T.text_light),
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# FUTURE WORK
# ══════════════════════════════════════════════════════════════════════
def make_future(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    _side_panel(slide, T)
    _header_bar(slide, T, "آفاق البحث المستقبلية", "مسارات البحث القادمة", "11")

    items=req.future_work[:6]
    if not items: return slide
    cols=2 if len(items)>2 else 1
    rows=(len(items)+cols-1)//cols
    gap=0.3
    cw=(_CONT_W-(cols-1)*gap)/cols
    avail=H-3.9
    ch=min(avail/rows-gap, 3.0)

    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=_CONT_X+ci*(cw+gap)
        y=3.8+ri*(ch+gap)
        card=rect(slide, x, y, cw, ch, _rgb(T.card))
        if card: shadow(card, blur=8, dist=2, alpha=0.25)
        # شريط يميني
        vline(slide, x+cw-0.38, y, ch, _rgb(T.accent), thickness=0.38)
        # رقم
        txt(slide, f"{i+1:02d}", x+0.25, y+0.18, 1.2, 0.85,
            font="Calibri", size=20, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.LEFT)
        hline(slide, x+0.25, y+1.0, cw-0.8, _rgb(T.accent), thickness=0.06)
        txt(slide, item, x+0.3, y+1.1, cw-0.95, ch-1.25,
            font=_FONT, size=12, spacing=18,
            color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# REFERENCES
# ══════════════════════════════════════════════════════════════════════
def make_references(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    _side_panel(slide, T)
    _header_bar(slide, T, "المراجع والمصادر", "قائمة المصادر المعتمدة", "12")

    refs=req.references[:14]
    if not refs: return slide
    avail=H-3.9
    ih=max(min(avail/max(len(refs),1)-0.08, 1.0), 0.5)

    for i,ref in enumerate(refs):
        y=3.8+i*(ih+0.08)
        if y+ih>H-0.3: break
        if i%2==0:
            rect(slide, _CONT_X, y, _CONT_W, ih, _rgb(T.bg2))
        txt(slide, f"[{i+1}]", _CONT_X+0.2, y+0.04, 1.4, ih-0.08,
            font="Calibri", size=9, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.LEFT)
        txt(slide, ref, _CONT_X+1.8, y+0.04, _CONT_W-3.4, ih-0.08,
            font=_FONT, size=10, color=_rgb(T.text_light),
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# FINAL — يعكس غلاف الـ classic مع اللوح على اليمين
# ══════════════════════════════════════════════════════════════════════
def make_final(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    # لوح يميني (معكوس عن الغلاف)
    rp=rect(slide, W-_SIDE_W, 0, _SIDE_W, H, _rgb(T.accent))
    if rp: gradient_fill(rp, T.accent_grad2, T.accent_grad1, 180)
    oval(slide, W-_SIDE_W+1, H*0.5, 7, 7, _rgb(T.bg), alpha=10)
    oval(slide, W-3, -2, 5, 5, _rgb(T.bg), alpha=7)

    cx=0.8; cw=W-_SIDE_W-1.6
    cy=H/2

    hline(slide, cx, cy-3.8, cw, _rgb(T.accent), thickness=0.1)
    txt(slide, "شكراً وتقديراً",
        cx, cy-3.4, cw, 2.4,
        font=_FONT, size=38, bold=True,
        color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    hline(slide, cx, cy-0.8, cw*0.55, _rgb(T.accent), thickness=0.07)

    txt(slide, req.student_name,
        cx, cy-0.5, cw, 1.2,
        font=_FONT, size=20, bold=True,
        color=_rgb(T.accent), align=PP_ALIGN.RIGHT, rtl=True)

    short=req.title_ar[:80]+("..." if len(req.title_ar)>80 else "")
    txt(slide, short, cx, cy+0.9, cw, 2.0,
        font=_FONT, size=12, italic=True,
        color=_rgb(T.muted), align=PP_ALIGN.RIGHT, rtl=True)

    hline(slide, cx, cy+3.2, cw, _rgb(T.accent), thickness=0.05)
    footer=[]
    if req.institution: footer.append(req.institution.split(' — ')[0])
    if req.year: footer.append(req.year)
    if footer:
        txt(slide, " · ".join(footer), cx, cy+3.4, cw, 0.7,
            font=_FONT, size=11, color=_rgb(T.muted),
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide
