"""
Engine C — BOLD MINIMAL  v17.2
هوية: جريء مينيمال — أرقام عملاقة، خطوط قوية، contrast عالي، flat design راقٍ
"""
from __future__ import annotations
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, cm, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, glow,
    set_solid_alpha, txt, blank_slide,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"
def set_font(f): global _FONT; _FONT = f


def _hx(h):
    h=h.lstrip('#')
    return RGBColor(int(h[0:2],16),int(h[2:4],16),int(h[4:6],16))


def _header(slide, T, title, sub=""):
    """Bold header: thick left bar + massive title + strong divider"""
    # Left accent block — thick
    ab = rect(slide, 0, 0, 1.0, 2.95, T.accent_rgb)
    if ab: gradient_fill(ab, T.accent_grad1, T.accent_grad2, 90)
    # Top accent thin strip
    rect(slide, 1.0, 0, W-1.0, 0.2, T.accent_rgb)
    # Title
    txt(slide, title, 1.2, 0.28, W-1.7, 1.65,
        font=_FONT, size=28, bold=True, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    if sub:
        txt(slide, sub, 1.2, 1.88, W-1.7, 0.8,
            font=_FONT, size=12, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    # Bold separator
    hline(slide, 1.0, 2.95, W-1.0, T.accent_rgb, thickness=0.12)
    hline(slide, 1.0, 3.08, W-1.0, T.bg2_rgb, thickness=0.04)


def make_cover(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    # Full background gradient
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=160)
    # Left thick accent column
    ab = rect(slide, 0, 0, 1.1, H, T.accent_rgb)
    if ab: gradient_fill(ab, T.accent_grad1, T.accent_grad2, 90)
    # Top + bottom thin lines
    rect(slide, 1.1, 0, W-1.1, 0.18, T.accent_rgb)
    rect(slide, 1.1, H-0.18, W-1.1, 0.18, T.accent_rgb)

    # Institution — small top
    if req.institution:
        txt(slide, req.institution, 1.4, 0.28, W-1.8, 0.75,
            font=_FONT, size=11, color=T.muted_rgb, align=PP_ALIGN.RIGHT)

    # HUGE title — central, minimal
    ty = H * 0.18
    sz = 36 if len(req.title_ar)<42 else 28 if len(req.title_ar)<65 else 22
    txt(slide, req.title_ar, 1.4, ty, W-1.9, H*0.44,
        font=_FONT, size=sz, bold=True, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)

    # Strong divider
    div_y = ty + H*0.44 + 0.25
    hline(slide, 1.4, div_y, W-2.0, T.accent_rgb, thickness=0.14)

    # Info — flat, no cards
    iy = div_y + 0.42
    gap = 0.72
    for val in filter(None,[req.student_name, req.supervisor, req.specialization, req.year]):
        txt(slide, val, 1.4, iy, W-1.9, gap-0.1,
            font=_FONT, size=15, bold=False, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        hline(slide, 1.4, iy+gap-0.16, W-2.0, T.bg2_rgb, thickness=0.04)
        iy += gap
    return slide


def make_intro(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=170)
    _header(slide, T, "مقدمة البحث")
    cy = 3.22
    for lbl, val in [(x,y) for x,y in [("نظرة عامة",req.intro_overview),("المنهج",req.intro_approach)] if y]:
        hline(slide, 1.0, cy, W-1.0, T.accent_rgb, thickness=0.1)
        cy += 0.25
        txt(slide, lbl, 1.2, cy, 10.0, 0.75,
            font=_FONT, size=15, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        cy += 0.82
        txt(slide, val, 1.2, cy, W-2.2, 3.2,
            font=_FONT, size=13, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        cy += 3.0
    return slide


def make_plan(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=170)
    _header(slide, T, "خطة البحث")
    chs = req.chapters[:8]; cy = 3.22
    avail = H-cy-0.25; rh = min(avail/max(len(chs),1)-0.08, 1.65)
    for i,ch in enumerate(chs):
        y = cy+i*(rh+0.08)
        hline(slide, 1.0, y, W-1.0, T.bg2_rgb, thickness=0.04)
        # Giant number
        nsz = min(int(rh*28), 42)
        txt(slide, f"{i+1:02d}", 1.05, y+0.08, 3.0, rh-0.1,
            font="Calibri", size=nsz, bold=True, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, ch.title, 4.3, y+0.15, W-5.8, rh-0.3,
            font=_FONT, size=14, bold=False, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        if ch.pages:
            txt(slide, f"ص {ch.pages}", W-1.5, y+0.15, 1.2, rh-0.3,
                font="Calibri", size=10, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_problem(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=170)
    _header(slide, T, "إشكالية البحث")
    cy = 3.22
    if req.main_problem:
        hline(slide, 1.0, cy, W-1.0, T.accent_rgb, thickness=0.12)
        cy += 0.28
        txt(slide, req.main_problem, 1.2, cy, W-2.2, 2.4,
            font=_FONT, size=16, bold=True, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        cy += 2.6
    if req.main_question:
        hline(slide, 1.0, cy, W-1.0, T.bg2_rgb, thickness=0.06)
        cy += 0.22
        txt(slide, req.main_question, 1.2, cy, W-2.2, 1.6,
            font=_FONT, size=15, italic=True, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        cy += 1.85
    for i,q in enumerate(req.sub_questions[:4]):
        hline(slide, 1.0, cy+i*0.82, W-1.0, T.bg2_rgb, thickness=0.03)
        txt(slide, f"{i+1}.  {q}", 1.2, cy+i*0.82+0.06, W-2.2, 0.72,
            font=_FONT, size=12, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_objectives(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=170)
    _header(slide, T, "أهداف البحث وفرضياته")
    cy = 3.22; items = req.objectives + req.hypotheses
    rh = min((H-cy-0.25)/max(len(items),1)-0.08, 1.35)
    for i,it in enumerate(items):
        y = cy+i*(rh+0.08)
        hline(slide, 1.0, y, W-1.0, T.bg2_rgb, thickness=0.04)
        txt(slide, str(i+1), 1.05, y+0.05, 2.4, rh-0.1,
            font="Calibri", size=int(rh*26), bold=True, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        # Tag for hypotheses
        if i >= len(req.objectives):
            tb = rrect(slide, W-4.2, y+(rh-0.42)/2, 2.8, 0.42, T.accent_rgb, radius_pct=50)
            if tb: gradient_fill(tb, T.accent_grad1, T.accent_grad2, 0)
            txt(slide, "فرضية", W-4.2, y+(rh-0.42)/2, 2.8, 0.42,
                font=_FONT, size=9, bold=True, color=T.text_dark_rgb, align=PP_ALIGN.CENTER)
        txt(slide, it, 3.6, y+0.1, W-8.2 if i>=len(req.objectives) else W-4.0, rh-0.2,
            font=_FONT, size=12.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_importance(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=170)
    _header(slide, T, "أهمية البحث")
    items = (req.importance+([req.reasons] if req.reasons else []))[:6]
    cy = 3.22; rh = min((H-cy-0.25)/max(len(items),1)-0.08, 1.42)
    for i,it in enumerate(items):
        y = cy+i*(rh+0.08)
        hline(slide, 1.0, y, W-1.0, T.bg2_rgb, thickness=0.04)
        # Accent dash pill
        dp = rrect(slide, 1.05, y+rh*0.38, 0.55, 0.14, T.accent_rgb, radius_pct=50)
        txt(slide, it, 1.75, y+0.12, W-2.75, rh-0.24,
            font=_FONT, size=13.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_methodology(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=170)
    _header(slide, T, "منهجية البحث")
    fields = [(l,v) for l,v in [("المنهج",req.methodology),("العينة",req.sample_type),
              ("الحجم",req.sample_size),("الأداة",req.tool)] if v]
    cy = 3.22; rh = min((H-cy-0.25)/max(len(fields),1)-0.1, 2.1)
    for i,(l,v) in enumerate(fields[:4]):
        y = cy+i*(rh+0.1)
        hline(slide, 1.0, y, W-1.0, T.bg2_rgb, thickness=0.06)
        txt(slide, l, 1.2, y+0.12, 8.0, 0.8,
            font=_FONT, size=14, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, v, 1.2, y+0.88, W-2.2, rh-1.0,
            font=_FONT, size=13.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_stats(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=170)
    _header(slide, T, "الإحصاءات والأرقام")
    stats = req.stats[:6]; n = len(stats)
    if not n: return slide
    cols = 3 if n>=3 else n; cy = 3.22; gap = 0.18
    cw = (W-1.0-(cols-1)*gap)/cols
    rows = (n+cols-1)//cols; ch = min((H-cy-0.25)/rows-gap, 4.0)
    for i,st in enumerate(stats):
        x = 1.0+(i%cols)*(cw+gap); y = cy+(i//cols)*(ch+gap)
        # Just left border line — pure minimal
        ab = rect(slide, x, y, 0.14, ch, T.accent_rgb)
        if ab: gradient_fill(ab, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, st.label, x+0.28, y+0.1, cw-0.38, 0.75,
            font=_FONT, size=11, bold=True, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
        vsz = 56 if len(st.value)<=4 else 40 if len(st.value)<=8 else 26
        txt(slide, st.value, x+0.28, y+0.65, cw-0.38, ch-1.15,
            font="Calibri", size=vsz, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        if st.unit:
            txt(slide, st.unit, x+0.28, y+ch-0.75, cw-0.38, 0.65,
                font=_FONT, size=11, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_results(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=170)
    _header(slide, T, "نتائج البحث")
    results = req.main_results[:8]; cy = 3.22
    rh = min((H-cy-0.25)/max(len(results),1)-0.08, 1.38)
    for i,r in enumerate(results):
        y = cy+i*(rh+0.08)
        hline(slide, 1.0, y, W-1.0, T.bg2_rgb, thickness=0.05)
        txt(slide, f"{i+1:02d}", 1.05, y+0.05, 2.5, rh-0.1,
            font="Calibri", size=int(rh*28), bold=True, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, r, 3.8, y+0.12, W-4.8, rh-0.24,
            font=_FONT, size=13, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_conclusion(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=170)
    _header(slide, T, "خاتمة البحث")
    # Huge opening quote
    txt(slide, "❝", 1.0, 2.9, 4.0, 4.0,
        font="Calibri", size=100, color=T.accent_rgb, align=PP_ALIGN.LEFT)
    txt(slide, req.general_conclusion, 1.2, 3.8, W-2.4, H-5.2,
        font=_FONT, size=17, bold=False, color=T.text_light_rgb,
        align=PP_ALIGN.RIGHT, spacing=26)
    hline(slide, 1.0, H-1.15, W-1.0, T.accent_rgb, thickness=0.1)
    txt(slide, req.student_name, 1.0, H-1.05, W-1.0, 0.85,
        font=_FONT, size=13, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_recommendations(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=170)
    _header(slide, T, "توصيات البحث")
    recs = req.recommendations[:8]; cy = 3.22
    rh = min((H-cy-0.25)/max(len(recs),1)-0.08, 1.38)
    for i,r in enumerate(recs):
        y = cy+i*(rh+0.08)
        hline(slide, 1.0, y, W-1.0, T.bg2_rgb, thickness=0.05)
        ab = rect(slide, 1.0, y, 0.14, rh, T.accent_rgb)
        if ab: gradient_fill(ab, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, r, 1.35, y+0.12, W-2.35, rh-0.24,
            font=_FONT, size=13, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_future(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=170)
    _header(slide, T, "آفاق البحث المستقبلية")
    items = req.future_work[:6]; cy = 3.22
    rh = min((H-cy-0.25)/max(len(items),1)-0.1, 1.65)
    for i,it in enumerate(items):
        y = cy+i*(rh+0.1)
        hline(slide, 1.0, y, W-1.0, T.bg2_rgb, thickness=0.05)
        txt(slide, "→", 1.05, y+0.12, 1.8, rh-0.24,
            font="Calibri", size=24, bold=True, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, it, 3.1, y+0.12, W-4.1, rh-0.24,
            font=_FONT, size=14, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_references(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=170)
    _header(slide, T, "المراجع والمصادر")
    refs = req.references[:12]; cy = 3.22
    ih = max(min((H-cy-0.25)/max(len(refs),1)-0.07, 1.05), 0.48)
    for i,r in enumerate(refs):
        y = cy+i*(ih+0.07)
        if y+ih > H-0.28: break
        hline(slide, 1.0, y, W-1.0, T.bg2_rgb, thickness=0.04)
        txt(slide, f"[{i+1}]", W-3.0, y+0.04, 1.8, ih-0.08,
            font="Calibri", size=11, bold=True, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, r, 1.2, y+0.04, W-4.4, ih-0.08,
            font=_FONT, size=11, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_final(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    # Pure accent color fill — maximum impact
    ab = rect(slide, 0, 0, W, H, T.accent_rgb)
    if ab: gradient_fill(ab, T.accent_grad1, T.accent_grad2, 135)
    # Dark strips top/bottom
    rect(slide, 0, 0, W, 1.8, T.bg_rgb)
    rect(slide, 0, H-1.8, W, 1.8, T.bg_rgb)
    # MASSIVE text
    txt(slide, "شكراً", 0, H*0.2, W, 3.8,
        font=_FONT, size=78, bold=True, color=T.bg_rgb, align=PP_ALIGN.CENTER)
    txt(slide, "وتقديراً", 0, H*0.2+3.6, W, 2.2,
        font=_FONT, size=44, bold=False, color=T.bg2_rgb, align=PP_ALIGN.CENTER)
    # Student name top strip
    txt(slide, req.student_name, 0, 0.25, W, 1.2,
        font=_FONT, size=18, bold=True, color=T.accent_rgb, align=PP_ALIGN.CENTER)
    # Bottom strip — institution + year
    fp = "  ·  ".join(filter(None,[req.institution, req.year]))
    txt(slide, fp or req.title_ar[:40], 0, H-1.55, W, 1.2,
        font=_FONT, size=12, color=T.muted_rgb, align=PP_ALIGN.CENTER)
    return slide
