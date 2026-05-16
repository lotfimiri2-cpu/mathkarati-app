"""
Engine A — CANVA LEVEL
هوية بصرية: سينمائية عصرية
- بطاقات مستديرة مع ظلال
- Gradients متدرجة
- عناصر دائرية زخرفية
- typography حديث بأحجام متنوعة
- layout: بطاقات متعددة الأعمدة
"""
from __future__ import annotations
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from engine.primitives import (
    W, H, cm, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow,
    set_solid_alpha, txt, blank_slide,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"
def set_font(f): global _FONT; _FONT = f


def _hx(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))


def _header(slide, T, title, sub=""):
    """CANVA header: gradient band + right accent bar + circles"""
    gradient_rect(slide, 0, 0, W, 2.9, T.grad1, T.grad2, angle=145)
    # Accent bar right
    ab = rect(slide, 0, 0, 0.45, 2.9, T.accent_rgb)
    if ab: gradient_fill(ab, T.accent_grad1, T.accent_grad2, 90)
    # Decorative circles top-right
    oval(slide, W-5.5, -2.2, 6.5, 6.5, T.accent_rgb, alpha=7)
    oval(slide, W-3.2, -0.8, 3.0, 3.0, T.bg2_rgb, alpha=20)
    txt(slide, title, 0.7, 0.28, W-1.5, 1.5,
        font=_FONT, size=24, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    if sub:
        txt(slide, sub, 0.7, 1.7, W-1.5, 0.9,
            font=_FONT, size=12, color=T.muted_rgb, align=PP_ALIGN.RIGHT)


def make_cover(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=145)
    # Big decorative circle bottom-left
    oval(slide, -4, H-8, 14, 14, T.accent_rgb, alpha=6)
    # Top accent bar full width
    top = rect(slide, 0, 0, W, 0.55, T.accent_rgb)
    if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)
    # Institution chip
    if req.institution:
        chip = rrect(slide, 1.2, 0.85, 16, 0.68, T.card_rgb, radius_pct=50)
        if chip: shadow(chip, blur=8, dist=2, alpha=0.25)
        txt(slide, req.institution, 1.4, 0.85, 15.6, 0.68,
            font=_FONT, size=11, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    # Main title card
    ty = H*0.27
    tc = rrect(slide, 1.5, ty-0.3, W-3.0, H*0.32+0.6, T.card_rgb, radius_pct=12)
    if tc: shadow(tc, blur=22, dist=7, alpha=0.45)
    stripe = rrect(slide, 1.5, ty-0.3, W-3.0, 0.28, T.accent_rgb, radius_pct=0)
    if stripe: gradient_fill(stripe, T.accent_grad1, T.accent_grad2, 0)
    sz = 28 if len(req.title_ar)<50 else 22 if len(req.title_ar)<80 else 18
    txt(slide, req.title_ar, 2.0, ty, W-4.0, H*0.32,
        font=_FONT, size=sz, bold=True, color=T.text_light_rgb, align=PP_ALIGN.CENTER)
    # Divider
    div_y = ty + H*0.32 + 0.55
    hline(slide, W*0.25, div_y, W*0.5, T.accent_rgb, thickness=0.06)
    # Info rows
    iy = div_y + 0.35
    rh = 0.62
    for label, val in [("الطالب :", req.student_name),
                       ("المشرف :", req.supervisor),
                       ("التخصص :", req.specialization)]:
        if not val: continue
        rb = rrect(slide, 1.5, iy, W-3.0, rh-0.06, T.bg2_rgb, radius_pct=8)
        txt(slide, label, 1.7, iy, 5.0, rh,
            font=_FONT, size=11, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, val, 7.0, iy, W-8.7, rh,
            font=_FONT, size=12, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        iy += rh
    # Year pill
    if req.year:
        yp = rrect(slide, W/2-2.5, H-1.0, 5.0, 0.58, T.accent_rgb, radius_pct=50)
        if yp: gradient_fill(yp, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, req.year, W/2-2.5, H-1.0, 5.0, 0.58,
            font="Calibri", size=12, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER)
    # Bottom bar
    bot = rect(slide, 0, H-0.22, W, 0.22, T.accent_rgb)
    if bot: gradient_fill(bot, T.accent_grad1, T.accent_grad2, 0)
    return slide


def make_intro(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "مقدمة البحث", "نظرة عامة وأسلوب المعالجة")
    cy = 3.1; cw = (W-2.8)/2
    for i,(lbl,val) in enumerate([(x,y) for x,y in [("نظرة عامة",req.intro_overview),("المنهج",req.intro_approach)] if y][:2]):
        x = 1.2 + i*(cw+0.4)
        c = rrect(slide, x, cy, cw, H-cy-0.7, T.card_rgb, radius_pct=10)
        if c: shadow(c, blur=14, dist=4, alpha=0.35)
        s = rrect(slide, x, cy, cw, 0.22, T.accent_rgb, radius_pct=0)
        if s: gradient_fill(s, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, x+0.2, cy+0.28, cw-0.4, 0.7, font=_FONT, size=13, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, val, x+0.2, cy+1.05, cw-0.4, H-cy-2.1, font=_FONT, size=11.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_plan(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "خطة البحث", f"{len(req.chapters)} فصول")
    cy = 3.1; avail = H-cy-0.5
    chs = req.chapters[:8]; rh = min(avail/max(len(chs),1)-0.1, 1.55)
    for i,ch in enumerate(chs):
        y = cy + i*(rh+0.1)
        rb = rrect(slide, 1.2, y, W-2.4, rh, T.card_rgb, radius_pct=8)
        if rb: shadow(rb, blur=8, dist=2, alpha=0.22)
        nb = oval(slide, 1.55, y+(rh-0.72)/2, 0.72, 0.72, T.accent_rgb)
        if nb: gradient_fill(nb, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, str(i+1), 1.55, y+(rh-0.72)/2, 0.72, 0.72,
            font="Calibri", size=12, bold=True, color=T.text_dark_rgb, align=PP_ALIGN.CENTER)
        txt(slide, ch.title, 2.55, y, W-5.8, rh,
            font=_FONT, size=13, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        if ch.pages:
            txt(slide, ch.pages, W-3.8, y, 2.4, rh,
                font="Calibri", size=10, color=T.muted_rgb, align=PP_ALIGN.LEFT)
    return slide


def make_problem(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "إشكالية البحث", "التساؤلات الرئيسية والفرعية")
    cy = 3.1
    if req.main_problem:
        c = rrect(slide, 1.2, cy, W-2.4, 2.05, T.card_rgb, radius_pct=10)
        if c: shadow(c, blur=12, dist=4, alpha=0.35)
        lb = rrect(slide, 1.2, cy, 4.0, 0.55, T.accent_rgb, radius_pct=0)
        if lb: gradient_fill(lb, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, "الإشكالية الرئيسية", 1.3, cy, 3.8, 0.55,
            font=_FONT, size=11, bold=True, color=T.text_dark_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, req.main_problem, 1.5, cy+0.65, W-3.0, 1.3,
            font=_FONT, size=12, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        cy += 2.2
    if req.main_question:
        qc = rrect(slide, 1.2, cy, W-2.4, 1.55, T.bg2_rgb, radius_pct=8)
        txt(slide, "❓", 1.4, cy+0.1, 1.2, 1.2, font="Calibri", size=28, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, req.main_question, 2.8, cy, W-4.3, 1.55,
            font=_FONT, size=12.5, bold=True, italic=True, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        cy += 1.7
    for i,q in enumerate(req.sub_questions[:5]):
        y = cy + i*0.82
        oval(slide, W-2.7, y+0.28, 0.3, 0.3, T.accent_rgb)
        txt(slide, q, 1.2, y, W-3.3, 0.8, font=_FONT, size=11, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_objectives(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "أهداف البحث وفرضياته")
    cy = 3.1; cw = (W-2.8)/2
    for i,(lbl,items) in enumerate([(x,y) for x,y in [("الأهداف",req.objectives),("الفرضيات",req.hypotheses)] if y][:2]):
        x = 1.2 + i*(cw+0.4)
        c = rrect(slide, x, cy, cw, H-cy-0.55, T.card_rgb, radius_pct=10)
        if c: shadow(c, blur=14, dist=4, alpha=0.35)
        h = rrect(slide, x, cy, cw, 0.65, T.accent_rgb, radius_pct=0)
        if h: gradient_fill(h, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, x+0.15, cy, cw-0.3, 0.65,
            font=_FONT, size=14, bold=True, color=T.text_dark_rgb, align=PP_ALIGN.CENTER)
        ih = min((H-cy-1.45)/max(len(items),1), 1.1)
        for j,it in enumerate(items[:7]):
            iy = cy+0.75+j*ih
            nb = oval(slide, x+cw-0.95, iy+0.06, 0.52, 0.52, T.bg_rgb)
            txt(slide, str(j+1), x+cw-0.95, iy+0.06, 0.52, 0.52,
                font="Calibri", size=9, bold=True, color=T.accent_rgb, align=PP_ALIGN.CENTER)
            txt(slide, it, x+0.2, iy, cw-1.3, ih,
                font=_FONT, size=10.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_importance(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "أهمية البحث")
    items = (req.importance + ([req.reasons] if req.reasons else []))[:6]
    cy = 3.1; cols = 2 if len(items)>3 else 1
    cw = (W-2.4-(cols-1)*0.3)/cols
    rows = (len(items)+cols-1)//cols
    ch = min((H-cy-0.5)/rows-0.2, 2.1)
    for i,it in enumerate(items):
        x = 1.2+(i%cols)*(cw+0.3); y = cy+(i//cols)*(ch+0.2)
        c = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=10)
        if c: shadow(c, blur=10, dist=3, alpha=0.3)
        vline(slide, x, y, ch, T.accent_rgb, thickness=0.22)
        txt(slide, it, x+0.42, y+0.12, cw-0.62, ch-0.24,
            font=_FONT, size=11.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_methodology(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "منهجية البحث", "الإجراءات والأدوات")
    fields = [(l,v) for l,v in [("المنهج",req.methodology),("العينة",req.sample_type),
              ("حجم العينة",req.sample_size),("الأداة",req.tool)] if v]
    cols = 2 if len(fields)>2 else 1
    cw = (W-2.4-(cols-1)*0.3)/cols
    rows = (len(fields)+cols-1)//cols
    ch = min((H-3.6)/rows-0.2, 2.4); cy = 3.1
    for i,(l,v) in enumerate(fields[:4]):
        x = 1.2+(i%cols)*(cw+0.3); y = cy+(i//cols)*(ch+0.2)
        c = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=10)
        if c: shadow(c, blur=10, dist=3, alpha=0.3)
        lb = rrect(slide, x, y, cw, 0.55, T.accent_rgb, radius_pct=0)
        if lb: gradient_fill(lb, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, l, x+0.2, y, cw-0.4, 0.55,
            font=_FONT, size=12, bold=True, color=T.text_dark_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, v, x+0.2, y+0.62, cw-0.4, ch-0.75,
            font=_FONT, size=11.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_stats(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "الإحصاءات والأرقام الرئيسية")
    stats = req.stats[:6]; n = len(stats)
    if not n: return slide
    cols = 3 if n>=3 else n; gap = 0.3
    cw = (W-2.4-(cols-1)*gap)/cols
    rows = (n+cols-1)//cols; cy = 3.1
    ch = min((H-cy-0.5)/rows-gap, 3.6)
    for i,st in enumerate(stats):
        x = 1.2+(i%cols)*(cw+gap); y = cy+(i//cols)*(ch+gap)
        c = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=12)
        if c: shadow(c, blur=14, dist=5, alpha=0.4)
        band = rrect(slide, x, y, cw, 0.22, T.accent_rgb, radius_pct=0)
        if band: gradient_fill(band, T.accent_grad1, T.accent_grad2, 0)
        vsz = 32 if len(st.value)<=4 else 24 if len(st.value)<=8 else 18
        txt(slide, st.value, x+0.2, y+0.38, cw-0.4, ch*0.52,
            font="Calibri", size=vsz, bold=True, color=T.accent_rgb, align=PP_ALIGN.CENTER)
        if st.unit:
            txt(slide, st.unit, x+0.2, y+ch*0.52+0.25, cw-0.4, 0.55,
                font=_FONT, size=10, color=T.muted_rgb, align=PP_ALIGN.CENTER)
        txt(slide, st.label, x+0.2, y+ch-0.82, cw-0.4, 0.72,
            font=_FONT, size=11, color=T.text_light_rgb, align=PP_ALIGN.CENTER)
    return slide


def make_results(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "نتائج البحث", "أبرز ما توصلت إليه الدراسة")
    results = req.main_results[:8]; cy = 3.1
    avail = H-cy-0.5; ih = min(avail/max(len(results),1), 1.42)-0.1
    for i,r in enumerate(results):
        y = cy+i*(ih+0.1)
        rb = rrect(slide, 1.2, y, W-2.4, ih, T.card_rgb, radius_pct=8)
        if rb: shadow(rb, blur=6, dist=2, alpha=0.2)
        badge = rrect(slide, W-3.1, y+(ih-0.58)/2, 1.6, 0.58, T.accent_rgb, radius_pct=50)
        if badge: gradient_fill(badge, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, str(i+1), W-3.1, y+(ih-0.58)/2, 1.6, 0.58,
            font="Calibri", size=11, bold=True, color=T.text_dark_rgb, align=PP_ALIGN.CENTER)
        txt(slide, r, 1.5, y+0.1, W-5.2, ih-0.2,
            font=_FONT, size=11.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_conclusion(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "خاتمة البحث", "الاستنتاج العام")
    cy=3.1; ch=H-cy-0.8
    c = rrect(slide, 1.5, cy, W-3.0, ch, T.card_rgb, radius_pct=14)
    if c: shadow(c, blur=20, dist=6, alpha=0.4)
    ts = rrect(slide, 1.5, cy, W-3.0, 0.28, T.accent_rgb, radius_pct=0)
    if ts: gradient_fill(ts, T.accent_grad1, T.accent_grad2, 0)
    txt(slide, "❝", 2.5, cy+0.35, 1.5, 1.5,
        font="Calibri", size=42, color=T.accent_rgb, align=PP_ALIGN.LEFT)
    txt(slide, req.general_conclusion, 2.0, cy+1.05, W-4.5, ch-1.5,
        font=_FONT, size=14, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_recommendations(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "توصيات البحث")
    recs=req.recommendations[:8]; cy=3.1
    avail=H-cy-0.5; ih=min(avail/max(len(recs),1), 1.4)-0.1
    for i,r in enumerate(recs):
        y=cy+i*(ih+0.1)
        rb=rrect(slide, 1.2, y, W-2.4, ih, T.card_rgb, radius_pct=8)
        if rb: shadow(rb, blur=6, dist=2, alpha=0.2)
        oval(slide, W-2.5, y+(ih-0.34)/2, 0.34, 0.34, T.accent_rgb)
        txt(slide, r, 1.5, y+0.08, W-3.3, ih-0.16,
            font=_FONT, size=11.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_future(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "آفاق البحث المستقبلية")
    items=req.future_work[:6]; cols=2 if len(items)>3 else 1
    cw=(W-2.4-(cols-1)*0.3)/cols
    rows=(len(items)+cols-1)//cols; cy=3.1
    ch=min((H-cy-0.5)/rows-0.2, 2.2)
    for i,it in enumerate(items):
        x=1.2+(i%cols)*(cw+0.3); y=cy+(i//cols)*(ch+0.2)
        c=rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=10)
        if c: shadow(c, blur=10, dist=3, alpha=0.3)
        ic=rrect(slide, x, y, 0.6, ch, T.accent_rgb, radius_pct=0)
        if ic: gradient_fill(ic, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, it, x+0.75, y+0.15, cw-0.95, ch-0.3,
            font=_FONT, size=11.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_references(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "المراجع والمصادر")
    refs=req.references[:12]; cy=3.1
    avail=H-cy-0.5; ih=max(min(avail/max(len(refs),1)-0.1, 1.1), 0.5)
    for i,r in enumerate(refs):
        y=cy+i*(ih+0.1)
        if y+ih>H-0.3: break
        if i%2==0: rrect(slide, 1.2, y, W-2.4, ih, T.card_rgb, radius_pct=4)
        txt(slide, f"[{i+1}]", W-2.9, y, 1.4, ih,
            font="Calibri", size=9, bold=True, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, r, 1.5, y+0.04, W-4.6, ih-0.08,
            font=_FONT, size=10, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_final(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=145)
    oval(slide, -3, -3, 12, 12, T.accent_rgb, alpha=6)
    oval(slide, W-9, H-9, 14, 14, T.accent_rgb, alpha=5)
    cw,ch = 22, 10; cx=(W-cw)/2; cy=(H-ch)/2
    c=rrect(slide, cx, cy, cw, ch, T.card_rgb, radius_pct=14)
    if c: shadow(c, blur=24, dist=8, alpha=0.45)
    ts=rrect(slide, cx, cy, cw, 0.35, T.accent_rgb, radius_pct=0)
    if ts: gradient_fill(ts, T.accent_grad1, T.accent_grad2, 0)
    txt(slide, "شكراً وتقديراً", cx+1.0, cy+0.55, cw-2.0, 2.5,
        font=_FONT, size=36, bold=True, color=T.text_light_rgb, align=PP_ALIGN.CENTER)
    hline(slide, cx+cw*0.2, cy+3.2, cw*0.6, T.accent_rgb, thickness=0.05)
    txt(slide, req.student_name, cx+1.0, cy+3.45, cw-2.0, 1.2,
        font=_FONT, size=18, bold=True, color=T.accent_rgb, align=PP_ALIGN.CENTER)
    title_d = req.title_ar[:70]+("..." if len(req.title_ar)>70 else "")
    txt(slide, title_d, cx+1.0, cy+4.75, cw-2.0, 2.0,
        font=_FONT, size=12, italic=True, color=T.muted_rgb, align=PP_ALIGN.CENTER)
    fp = "  ·  ".join(filter(None,[req.institution, req.year]))
    if fp:
        txt(slide, fp, cx+1.0, cy+ch-1.0, cw-2.0, 0.8,
            font=_FONT, size=11, color=T.muted_rgb, align=PP_ALIGN.CENTER)
    bb=rect(slide, 0, H-0.25, W, 0.25, T.accent_rgb)
    if bb: gradient_fill(bb, T.accent_grad1, T.accent_grad2, 0)
    return slide
