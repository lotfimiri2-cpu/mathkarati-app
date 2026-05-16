"""
Engine C — BOLD MINIMAL
هوية بصرية: جريء مينيمال
- خلفية داكنة مع مساحات بيضاء كبيرة
- Typography ضخم جداً — الأرقام والعناوين عملاقة
- لا بطاقات ولا ظلال — مسطح تماماً (flat design)
- خطوط أفقية رفيعة فقط كفاصل
- لون accent واحد مسيطر بكثافة
- layout: قائمة مفتوحة مع أرقام عملاقة
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
    """BOLD header: huge number + accent line + massive title — NO card, NO shadow"""
    # Top thin accent strip
    rect(slide, 0, 0, W, 0.18, T.accent_rgb)
    # Left thick accent block
    rect(slide, 0, 0.18, 0.9, 2.4, T.accent_rgb)
    # Section number (huge, semi-transparent)
    txt(slide, title, 1.1, 0.25, W-1.6, 1.8,
        font=_FONT, size=30, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    if sub:
        txt(slide, sub, 1.1, 1.95, W-1.6, 0.75,
            font=_FONT, size=12, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    # Thin divider line below header
    hline(slide, 0, 2.75, W, T.accent_rgb, thickness=0.05)
    hline(slide, 0.9, 2.82, W-0.9, T.bg2_rgb, thickness=0.03)


def make_cover(prs, req, T):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    # Left accent half
    rect(slide, 0, 0, 0.9, H, T.accent_rgb)
    # Top thin line
    rect(slide, 0.9, 0, W-0.9, 0.12, T.accent_rgb)
    # Bottom thin line
    rect(slide, 0.9, H-0.12, W-0.9, 0.12, T.accent_rgb)
    # Institution — top right, small
    if req.institution:
        txt(slide, req.institution, 1.2, 0.3, W-1.7, 0.7,
            font=_FONT, size=11, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    # Big title — takes up most of slide
    ty = H*0.2
    sz = 34 if len(req.title_ar)<45 else 26 if len(req.title_ar)<70 else 20
    txt(slide, req.title_ar, 1.2, ty, W-1.7, H*0.42,
        font=_FONT, size=sz, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    # Accent divider
    hline(slide, 1.2, ty + H*0.42 + 0.2, W-1.7, T.accent_rgb, thickness=0.1)
    # Student info — flat list no cards
    iy = ty + H*0.42 + 0.55
    for val in filter(None, [req.student_name, req.supervisor, req.specialization, req.year]):
        txt(slide, val, 1.2, iy, W-1.7, 0.75,
            font=_FONT, size=14, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        iy += 0.78
    return slide


def make_intro(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "مقدمة البحث", "")
    cy = 3.0
    for lbl, val in [(x,y) for x,y in [("نظرة عامة",req.intro_overview),("المنهج",req.intro_approach)] if y]:
        hline(slide, 1.0, cy, W-2.0, T.accent_rgb, thickness=0.06)
        cy += 0.18
        txt(slide, lbl, 1.0, cy, 8.0, 0.7,
            font=_FONT, size=14, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        cy += 0.72
        txt(slide, val, 1.0, cy, W-2.0, 3.0,
            font=_FONT, size=12.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        cy += 2.8
    return slide


def make_plan(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "خطة البحث", "")
    chs = req.chapters[:8]; cy = 2.95
    avail = H-cy-0.3; rh = min(avail/max(len(chs),1)-0.06, 1.6)
    for i, ch in enumerate(chs):
        y = cy + i*(rh+0.06)
        hline(slide, 0.9, y, W-0.9, T.bg2_rgb, thickness=0.04)
        # Giant number
        txt(slide, f"{i+1:02d}", 0.95, y+0.1, 2.5, rh-0.1,
            font="Calibri", size=int(rh*26), bold=True,
            color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, ch.title, 3.7, y+0.1, W-5.3, rh-0.1,
            font=_FONT, size=14, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        if ch.pages:
            txt(slide, ch.pages, W-1.6, y+0.1, 1.3, rh-0.1,
                font="Calibri", size=10, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_problem(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "إشكالية البحث")
    cy = 2.95
    if req.main_problem:
        hline(slide, 0.9, cy, W-0.9, T.accent_rgb, thickness=0.08)
        cy += 0.25
        txt(slide, req.main_problem, 1.0, cy, W-2.0, 2.2,
            font=_FONT, size=15, bold=True, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        cy += 2.4
    if req.main_question:
        hline(slide, 0.9, cy, W-0.9, T.bg2_rgb, thickness=0.04)
        cy += 0.2
        txt(slide, req.main_question, 1.0, cy, W-2.0, 1.5,
            font=_FONT, size=14, italic=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        cy += 1.7
    for i, q in enumerate(req.sub_questions[:4]):
        txt(slide, f"{i+1}.  {q}", 1.0, cy+i*0.82, W-2.0, 0.78,
            font=_FONT, size=12, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_objectives(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "أهداف البحث وفرضياته")
    cy = 2.95; items = req.objectives[:5] + req.hypotheses[:3]
    rh = min((H-cy-0.3)/max(len(items),1)-0.06, 1.3)
    for i, it in enumerate(items):
        y = cy + i*(rh+0.06)
        hline(slide, 0.9, y, W-0.9, T.bg2_rgb, thickness=0.03)
        # Big number
        txt(slide, str(i+1), 0.95, y, 2.2, rh,
            font="Calibri", size=int(rh*24), bold=True,
            color=T.accent_rgb, align=PP_ALIGN.LEFT)
        # Label tag for hypotheses
        if i >= len(req.objectives):
            tag = rect(slide, W-3.5, y+(rh-0.4)/2, 2.2, 0.4, T.accent_rgb)
            txt(slide, "فرضية", W-3.5, y+(rh-0.4)/2, 2.2, 0.4,
                font=_FONT, size=9, bold=True, color=T.text_dark_rgb, align=PP_ALIGN.CENTER)
        txt(slide, it, 3.3, y+0.08, W-7.0 if i>=len(req.objectives) else W-3.8, rh-0.16,
            font=_FONT, size=12, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_importance(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "أهمية البحث")
    items = (req.importance + ([req.reasons] if req.reasons else []))[:6]
    cy = 2.95; rh = min((H-cy-0.3)/max(len(items),1)-0.06, 1.4)
    for i, it in enumerate(items):
        y = cy + i*(rh+0.06)
        hline(slide, 0.9, y, W-0.9, T.bg2_rgb, thickness=0.03)
        # Accent dash
        rect(slide, 0.9, y+rh*0.4, 0.45, 0.1, T.accent_rgb)
        txt(slide, it, 1.55, y+0.1, W-2.55, rh-0.2,
            font=_FONT, size=13, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_methodology(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "منهجية البحث")
    fields = [(l,v) for l,v in [("المنهج",req.methodology),("العينة",req.sample_type),
              ("الحجم",req.sample_size),("الأداة",req.tool)] if v]
    cy = 2.95; rh = min((H-cy-0.3)/max(len(fields),1)-0.08, 2.0)
    for i,(l,v) in enumerate(fields[:4]):
        y = cy + i*(rh+0.08)
        hline(slide, 0.9, y, W-0.9, T.bg2_rgb, thickness=0.04)
        # Field label — huge accent
        txt(slide, l, 0.95, y+0.1, 7.0, 0.75,
            font=_FONT, size=13, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, v, 0.95, y+0.8, W-2.0, rh-0.9,
            font=_FONT, size=13, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_stats(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "الإحصاءات والأرقام")
    stats = req.stats[:6]; n = len(stats)
    if not n: return slide
    cols = 3 if n>=3 else n; cy = 2.95; gap = 0.15
    cw = (W-0.9-(cols-1)*gap)/cols
    rows = (n+cols-1)//cols; ch = min((H-cy-0.3)/rows-gap, 3.8)
    for i, st in enumerate(stats):
        x = 0.9+(i%cols)*(cw+gap); y = cy+(i//cols)*(ch+gap)
        # Just left border — NO card
        rect(slide, x, y, 0.12, ch, T.accent_rgb)
        txt(slide, st.label, x+0.3, y+0.1, cw-0.4, 0.7,
            font=_FONT, size=11, bold=True, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
        vsz = 52 if len(st.value)<=4 else 36 if len(st.value)<=8 else 24
        txt(slide, st.value, x+0.3, y+0.6, cw-0.4, ch-1.1,
            font="Calibri", size=vsz, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        if st.unit:
            txt(slide, st.unit, x+0.3, y+ch-0.7, cw-0.4, 0.6,
                font=_FONT, size=11, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_results(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "نتائج البحث")
    results = req.main_results[:8]; cy = 2.95
    rh = min((H-cy-0.3)/max(len(results),1)-0.06, 1.35)
    for i, r in enumerate(results):
        y = cy + i*(rh+0.06)
        hline(slide, 0.9, y, W-0.9, T.bg2_rgb, thickness=0.04)
        txt(slide, f"{i+1:02d}", 0.95, y+0.05, 2.2, rh-0.1,
            font="Calibri", size=int(rh*26), bold=True,
            color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, r, 3.4, y+0.1, W-4.4, rh-0.2,
            font=_FONT, size=12.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_conclusion(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "خاتمة البحث")
    # Giant quote mark
    txt(slide, "❝", 0.9, 2.8, 3.0, 3.0,
        font="Calibri", size=80, color=T.accent_rgb, align=PP_ALIGN.LEFT)
    txt(slide, req.general_conclusion, 1.0, 3.5, W-2.0, H-4.5,
        font=_FONT, size=16, bold=False,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    hline(slide, 0.9, H-1.0, W-0.9, T.accent_rgb, thickness=0.08)
    txt(slide, req.student_name, 0.9, H-0.9, W-0.9, 0.8,
        font=_FONT, size=12, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_recommendations(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "توصيات البحث")
    recs = req.recommendations[:8]; cy = 2.95
    rh = min((H-cy-0.3)/max(len(recs),1)-0.06, 1.35)
    for i, r in enumerate(recs):
        y = cy + i*(rh+0.06)
        hline(slide, 0.9, y, W-0.9, T.bg2_rgb, thickness=0.04)
        rect(slide, 0.9, y, 0.12, rh, T.accent_rgb)
        txt(slide, r, 1.2, y+0.1, W-2.2, rh-0.2,
            font=_FONT, size=12.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_future(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "آفاق البحث المستقبلية")
    items = req.future_work[:6]; cy = 2.95
    rh = min((H-cy-0.3)/max(len(items),1)-0.08, 1.6)
    for i, it in enumerate(items):
        y = cy + i*(rh+0.08)
        hline(slide, 0.9, y, W-0.9, T.bg2_rgb, thickness=0.04)
        txt(slide, f"→", 0.95, y+0.1, 1.5, rh-0.2,
            font="Calibri", size=22, bold=True, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, it, 2.6, y+0.1, W-3.6, rh-0.2,
            font=_FONT, size=13, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_references(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    _header(slide, T, "المراجع والمصادر")
    refs = req.references[:12]; cy = 2.95
    ih = max(min((H-cy-0.3)/max(len(refs),1)-0.06, 1.0), 0.48)
    for i, r in enumerate(refs):
        y = cy + i*(ih+0.06)
        if y+ih > H-0.3: break
        hline(slide, 0.9, y, W-0.9, T.bg2_rgb, thickness=0.03)
        txt(slide, f"[{i+1}]", W-2.8, y+0.04, 1.6, ih-0.08,
            font="Calibri", size=10, bold=True, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, r, 0.95, y+0.04, W-4.2, ih-0.08,
            font=_FONT, size=10.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_final(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    # Full accent background — no image, pure color
    rect(slide, 0, 0, W, H, T.accent_rgb)
    # Dark overlay top and bottom strips
    rect(slide, 0, 0, W, 1.5, T.bg_rgb)
    rect(slide, 0, H-1.5, W, 1.5, T.bg_rgb)
    # Center text — black on accent
    txt(slide, "شكراً", 0, H*0.2, W, 3.5,
        font=_FONT, size=72, bold=True,
        color=T.bg_rgb, align=PP_ALIGN.CENTER)
    txt(slide, "وتقديراً", 0, H*0.2+3.2, W, 2.0,
        font=_FONT, size=40, bold=False,
        color=T.bg2_rgb, align=PP_ALIGN.CENTER)
    # Student name bottom
    txt(slide, req.student_name, 0, H-1.3, W, 0.9,
        font=_FONT, size=16, bold=True,
        color=T.bg_rgb, align=PP_ALIGN.CENTER)
    return slide
