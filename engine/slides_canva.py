"""
Engine A — CANVA LEVEL  v17.2 VISUAL UPGRADE
هوية: سينمائية عصرية — بطاقات عميقة، gradients ثلاثية، glow، ظلال ناعمة
"""
from __future__ import annotations
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, cm, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, glow, soft_shadow,
    set_solid_alpha, txt, blank_slide,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"
def set_font(f): global _FONT; _FONT = f

def _hx(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))


# ── Decorative elements ───────────────────────────────────────────────
def _decor_circles(slide, T, x_offset=0):
    """Floating decorative circles — depth & atmosphere"""
    oval(slide, W-7+x_offset, -3.5, 10, 10, T.accent_rgb, alpha=5)
    oval(slide, W-4+x_offset, -1.2, 5, 5, T.bg2_rgb, alpha=15)
    oval(slide, -2+x_offset, H-5, 8, 8, T.accent_rgb, alpha=4)

def _accent_bar(slide, T, y=0, h=0.55):
    b = rect(slide, 0, y, W, h, T.accent_rgb)
    if b: gradient_fill(b, T.accent_grad1, T.accent_grad2, 0)
    return b


def _header(slide, T, title, sub=""):
    # Multi-layer header: dark gradient + decorative circles
    gradient_rect(slide, 0, 0, W, 3.0, T.grad1, T.grad2, angle=145,
                  c3=T.bg2 if hasattr(T,'bg2') else None)
    # Left accent thick bar
    ab = rect(slide, 0, 0, 0.5, 3.0, T.accent_rgb)
    if ab: gradient_fill(ab, T.accent_grad1, T.accent_grad2, 90)
    # Decorative circles top-right
    oval(slide, W-6, -2.5, 7.5, 7.5, T.accent_rgb, alpha=6)
    oval(slide, W-3.5, -0.8, 3.5, 3.5, T.bg2_rgb, alpha=18)
    # Title
    txt(slide, title, 0.8, 0.2, W-1.6, 1.7,
        font=_FONT, size=25, bold=True, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    if sub:
        txt(slide, sub, 0.8, 1.8, W-1.6, 0.95,
            font=_FONT, size=12, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    # Bottom separator
    hline(slide, 0.5, 3.0, W-0.5, T.accent_rgb, thickness=0.05)


def _card(slide, T, x, y, w, h, radius=10, use_glow=False):
    c = rrect(slide, x, y, w, h, T.card_rgb, radius_pct=radius)
    if c:
        soft_shadow(c, alpha=0.38)
        if use_glow:
            glow(c, T.accent, radius=8, alpha=0.12)
        # Top accent stripe
        stripe = rrect(slide, x, y, w, 0.24, T.accent_rgb, radius_pct=0)
        if stripe: gradient_fill(stripe, T.accent_grad1, T.accent_grad2, 0)
    return c


def make_cover(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    # Full gradient background
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=145)
    _decor_circles(slide, T)
    # Extra subtle oval bottom-left
    oval(slide, -4, H-7, 11, 11, T.bg2_rgb, alpha=20)

    # Top accent bar
    _accent_bar(slide, T, y=0, h=0.6)

    # Institution badge
    if req.institution:
        chip = rrect(slide, 1.0, 0.82, 20, 0.72, T.card_rgb, radius_pct=50)
        if chip: shadow(chip, blur=8, dist=2, alpha=0.22)
        txt(slide, req.institution, 1.2, 0.82, 19.6, 0.72,
            font=_FONT, size=11, color=T.muted_rgb, align=PP_ALIGN.RIGHT)

    # Main title card — large, centered
    ty = H * 0.25
    th = H * 0.33
    tc = rrect(slide, 1.2, ty-0.35, W-2.4, th+0.7, T.card_rgb, radius_pct=14)
    if tc:
        soft_shadow(tc, alpha=0.48)
        glow(tc, T.accent, radius=12, alpha=0.08)
    # Card top stripe
    ts = rrect(slide, 1.2, ty-0.35, W-2.4, 0.3, T.accent_rgb, radius_pct=0)
    if ts: gradient_fill(ts, T.accent_grad1, T.accent_grad2, 0)

    # Title text
    sz = 30 if len(req.title_ar) < 45 else 24 if len(req.title_ar) < 70 else 19
    txt(slide, req.title_ar, 1.8, ty, W-3.6, th,
        font=_FONT, size=sz, bold=True, color=T.text_light_rgb, align=PP_ALIGN.CENTER)

    # English subtitle
    if req.title_en or (hasattr(req,'title_fr') and req.title_fr):
        sub = getattr(req,'title_en','') or getattr(req,'title_fr','')
        if sub:
            txt(slide, sub, 1.8, ty+th-0.85, W-3.6, 0.8,
                font="Calibri", size=12, italic=True,
                color=T.muted_rgb, align=PP_ALIGN.CENTER)

    # Gold divider
    div_y = ty + th + 0.55
    hline(slide, W*0.22, div_y, W*0.56, T.accent_rgb, thickness=0.07)
    # Small decorative dot in center of divider
    dot = oval(slide, W/2-0.22, div_y-0.14, 0.44, 0.44, T.accent_rgb)
    if dot: gradient_fill(dot, T.accent_grad1, T.accent_grad2, 0)

    # Info rows
    iy = div_y + 0.45
    rh = 0.65
    for label, val in [("الطالب", req.student_name),
                        ("المشرف", req.supervisor),
                        ("التخصص", req.specialization)]:
        if not val: continue
        rb = rrect(slide, 1.5, iy, W-3.0, rh-0.06, T.bg2_rgb, radius_pct=8)
        if rb: shadow(rb, blur=6, dist=2, alpha=0.18)
        # Label dot
        ldot = oval(slide, W-3.15, iy+(rh-0.28)/2, 0.28, 0.28, T.accent_rgb)
        txt(slide, label+" :", 1.8, iy, 5.5, rh,
            font=_FONT, size=11, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, val, 7.5, iy, W-9.2, rh,
            font=_FONT, size=12, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        iy += rh

    # Year pill
    if req.year:
        yp = rrect(slide, W/2-2.8, H-1.1, 5.6, 0.65, T.accent_rgb, radius_pct=50)
        if yp:
            gradient_fill(yp, T.accent_grad1, T.accent_grad2, 0)
            shadow(yp, blur=10, dist=3, alpha=0.3)
        txt(slide, req.year, W/2-2.8, H-1.1, 5.6, 0.65,
            font="Calibri", size=13, bold=True, color=T.text_dark_rgb, align=PP_ALIGN.CENTER)

    # Bottom accent bar
    _accent_bar(slide, T, y=H-0.28, h=0.28)
    return slide


def make_intro(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=160)
    _header(slide, T, "مقدمة البحث", "نظرة عامة وأسلوب المعالجة")
    cy = 3.15; cw = (W-2.8)/2; gap = 0.4

    pairs = [(x,y) for x,y in [("نظرة عامة",req.intro_overview),
                                 ("المنهج المتبع",req.intro_approach)] if y]
    for i,(lbl,val) in enumerate(pairs[:2]):
        x = 1.2 + i*(cw+gap)
        _card(slide, T, x, cy, cw, H-cy-0.8, radius=12, use_glow=(i==0))
        txt(slide, lbl, x+0.25, cy+0.35, cw-0.5, 0.75,
            font=_FONT, size=13, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, val, x+0.25, cy+1.15, cw-0.5, H-cy-2.3,
            font=_FONT, size=11.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_plan(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=160)
    _header(slide, T, "خطة البحث", f"يتكون البحث من {len(req.chapters)} فصول")
    cy = 3.15; avail = H-cy-0.5
    chs = req.chapters[:8]; rh = min(avail/max(len(chs),1)-0.12, 1.6)
    for i,ch in enumerate(chs):
        y = cy + i*(rh+0.12)
        rb = rrect(slide, 1.2, y, W-2.4, rh, T.card_rgb, radius_pct=10)
        if rb: shadow(rb, blur=10, dist=3, alpha=0.28)
        # Number badge with gradient
        nb = oval(slide, 1.55, y+(rh-0.75)/2, 0.75, 0.75, T.accent_rgb)
        if nb: gradient_fill(nb, T.accent_grad1, T.accent_grad2, 45)
        txt(slide, str(i+1), 1.55, y+(rh-0.75)/2, 0.75, 0.75,
            font="Calibri", size=13, bold=True, color=T.text_dark_rgb, align=PP_ALIGN.CENTER)
        # Left accent mini-bar
        vline(slide, 2.55, y+0.1, rh-0.2, T.accent_rgb, thickness=0.06)
        txt(slide, ch.title, 2.78, y, W-6.0, rh,
            font=_FONT, size=13, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        if ch.pages:
            pp = rrect(slide, W-3.9, y+(rh-0.38)/2, 2.5, 0.38, T.bg2_rgb, radius_pct=50)
            txt(slide, f"ص {ch.pages}", W-3.9, y+(rh-0.38)/2, 2.5, 0.38,
                font="Calibri", size=9, color=T.muted_rgb, align=PP_ALIGN.CENTER)
    return slide


def make_problem(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=160)
    _header(slide, T, "إشكالية البحث", "التساؤلات الرئيسية والفرعية")
    cy = 3.15
    if req.main_problem:
        c = rrect(slide, 1.2, cy, W-2.4, 2.1, T.card_rgb, radius_pct=12)
        if c: soft_shadow(c, alpha=0.38)
        lb = rrect(slide, 1.2, cy, 4.2, 0.58, T.accent_rgb, radius_pct=0)
        if lb: gradient_fill(lb, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, "الإشكالية الرئيسية", 1.35, cy, 4.0, 0.58,
            font=_FONT, size=11, bold=True, color=T.text_dark_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, req.main_problem, 1.45, cy+0.68, W-3.0, 1.3,
            font=_FONT, size=12, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        cy += 2.28
    if req.main_question:
        qc = rrect(slide, 1.2, cy, W-2.4, 1.6, T.bg2_rgb, radius_pct=10)
        if qc: shadow(qc, blur=8, dist=2, alpha=0.2)
        txt(slide, "❓", 1.4, cy+0.1, 1.3, 1.3,
            font="Calibri", size=30, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, req.main_question, 2.9, cy, W-4.4, 1.6,
            font=_FONT, size=12.5, bold=True, italic=True,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
        cy += 1.78
    for i,q in enumerate(req.sub_questions[:4]):
        y = cy + i*0.84
        dot = oval(slide, W-2.75, y+0.28, 0.3, 0.3, T.accent_rgb)
        txt(slide, q, 1.2, y, W-3.3, 0.8,
            font=_FONT, size=11, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_objectives(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=160)
    _header(slide, T, "أهداف البحث وفرضياته")
    cy = 3.15; cw = (W-2.8)/2; gap = 0.4
    pairs = [(x,y) for x,y in [("الأهداف",req.objectives),("الفرضيات",req.hypotheses)] if y]
    for i,(lbl,items) in enumerate(pairs[:2]):
        x = 1.2 + i*(cw+gap)
        c = rrect(slide, x, cy, cw, H-cy-0.6, T.card_rgb, radius_pct=12)
        if c: soft_shadow(c, alpha=0.35)
        hdr = rrect(slide, x, cy, cw, 0.68, T.accent_rgb, radius_pct=0)
        if hdr: gradient_fill(hdr, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, x+0.15, cy, cw-0.3, 0.68,
            font=_FONT, size=14, bold=True, color=T.text_dark_rgb, align=PP_ALIGN.CENTER)
        ih = min((H-cy-1.5)/max(len(items),1), 1.12)
        for j,it in enumerate(items[:7]):
            iy = cy+0.78+j*ih
            # Numbered oval
            nb = oval(slide, x+cw-0.98, iy+0.06, 0.54, 0.54, T.bg_rgb)
            txt(slide, str(j+1), x+cw-0.98, iy+0.06, 0.54, 0.54,
                font="Calibri", size=9, bold=True, color=T.accent_rgb, align=PP_ALIGN.CENTER)
            txt(slide, it, x+0.2, iy, cw-1.35, ih,
                font=_FONT, size=10.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_importance(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=160)
    _header(slide, T, "أهمية البحث")
    items = (req.importance + ([req.reasons] if req.reasons else []))[:6]
    cy = 3.15; cols = 2 if len(items)>3 else 1
    cw = (W-2.4-(cols-1)*0.35)/cols
    rows = (len(items)+cols-1)//cols
    ch = min((H-cy-0.5)/rows-0.22, 2.15)
    for i,it in enumerate(items):
        x = 1.2+(i%cols)*(cw+0.35); y = cy+(i//cols)*(ch+0.22)
        c = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=12)
        if c: shadow(c, blur=12, dist=3, alpha=0.3)
        # Left accent bar
        ab = rrect(slide, x, y, 0.28, ch, T.accent_rgb, radius_pct=0)
        if ab: gradient_fill(ab, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, it, x+0.45, y+0.14, cw-0.65, ch-0.28,
            font=_FONT, size=11.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_methodology(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=160)
    _header(slide, T, "منهجية البحث", "الإجراءات والأدوات")
    fields = [(l,v) for l,v in [("المنهج",req.methodology),("العينة",req.sample_type),
              ("حجم العينة",req.sample_size),("الأداة",req.tool)] if v]
    cy = 3.15; cols = 2 if len(fields)>2 else 1
    cw = (W-2.4-(cols-1)*0.35)/cols
    rows = (len(fields)+cols-1)//cols
    ch = min((H-cy-0.5)/rows-0.22, 2.4)
    for i,(l,v) in enumerate(fields[:4]):
        x = 1.2+(i%cols)*(cw+0.35); y = cy+(i//cols)*(ch+0.22)
        c = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=12)
        if c: shadow(c, blur=12, dist=3, alpha=0.3)
        lb = rrect(slide, x, y, cw, 0.58, T.accent_rgb, radius_pct=0)
        if lb: gradient_fill(lb, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, l, x+0.2, y, cw-0.4, 0.58,
            font=_FONT, size=12, bold=True, color=T.text_dark_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, v, x+0.2, y+0.65, cw-0.4, ch-0.78,
            font=_FONT, size=11.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_stats(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=160)
    _header(slide, T, "الإحصاءات والأرقام الرئيسية")
    stats = req.stats[:6]; n = len(stats)
    if not n: return slide
    cols = 3 if n>=3 else n; gap = 0.35; cy = 3.15
    cw = (W-2.4-(cols-1)*gap)/cols
    rows = (n+cols-1)//cols
    ch = min((H-cy-0.5)/rows-gap, 3.7)
    for i,st in enumerate(stats):
        x = 1.2+(i%cols)*(cw+gap); y = cy+(i//cols)*(ch+gap)
        c = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=14)
        if c:
            soft_shadow(c, alpha=0.42)
            glow(c, T.accent, radius=6, alpha=0.07)
        # Top accent band with gradient
        band = rrect(slide, x, y, cw, 0.26, T.accent_rgb, radius_pct=0)
        if band: gradient_fill(band, T.accent_grad1, T.accent_grad2, 0)
        # Big value
        vsz = 34 if len(st.value)<=4 else 26 if len(st.value)<=8 else 18
        txt(slide, st.value, x+0.18, y+0.4, cw-0.36, ch*0.5,
            font="Calibri", size=vsz, bold=True, color=T.accent_rgb, align=PP_ALIGN.CENTER)
        if st.unit:
            txt(slide, st.unit, x+0.18, y+ch*0.5+0.22, cw-0.36, 0.55,
                font=_FONT, size=10, color=T.muted_rgb, align=PP_ALIGN.CENTER)
        hline(slide, x+cw*0.2, y+ch-0.9, cw*0.6, T.accent_rgb, thickness=0.04)
        txt(slide, st.label, x+0.18, y+ch-0.82, cw-0.36, 0.72,
            font=_FONT, size=11, color=T.text_light_rgb, align=PP_ALIGN.CENTER)
    return slide


def make_results(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=160)
    _header(slide, T, "نتائج البحث", "أبرز ما توصلت إليه الدراسة")
    results = req.main_results[:8]; cy = 3.15
    avail = H-cy-0.5; ih = min(avail/max(len(results),1)-0.12, 1.45)
    for i,r in enumerate(results):
        y = cy+i*(ih+0.12)
        rb = rrect(slide, 1.2, y, W-2.4, ih, T.card_rgb, radius_pct=10)
        if rb: shadow(rb, blur=8, dist=2, alpha=0.25)
        # Badge
        badge = rrect(slide, W-3.2, y+(ih-0.6)/2, 1.7, 0.6, T.accent_rgb, radius_pct=50)
        if badge: gradient_fill(badge, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, str(i+1), W-3.2, y+(ih-0.6)/2, 1.7, 0.6,
            font="Calibri", size=12, bold=True, color=T.text_dark_rgb, align=PP_ALIGN.CENTER)
        txt(slide, r, 1.5, y+0.12, W-5.3, ih-0.24,
            font=_FONT, size=11.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_conclusion(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=145)
    _header(slide, T, "خاتمة البحث", "الاستنتاج العام")
    cy = 3.15; ch = H-cy-0.9
    c = rrect(slide, 1.5, cy, W-3.0, ch, T.card_rgb, radius_pct=16)
    if c: soft_shadow(c, alpha=0.45)
    ts = rrect(slide, 1.5, cy, W-3.0, 0.3, T.accent_rgb, radius_pct=0)
    if ts: gradient_fill(ts, T.accent_grad1, T.accent_grad2, 0)
    # Big quote marks
    txt(slide, "❝", 2.2, cy+0.3, 2.0, 2.0,
        font="Calibri", size=48, color=T.accent_rgb, align=PP_ALIGN.LEFT)
    txt(slide, "❞", W-4.2, cy+ch-2.2, 2.0, 2.0,
        font="Calibri", size=48, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
    txt(slide, req.general_conclusion, 2.2, cy+1.1, W-5.0, ch-1.8,
        font=_FONT, size=14, color=T.text_light_rgb, align=PP_ALIGN.RIGHT, spacing=22)
    return slide


def make_recommendations(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=160)
    _header(slide, T, "توصيات البحث")
    recs = req.recommendations[:8]; cy = 3.15
    avail = H-cy-0.5; ih = min(avail/max(len(recs),1)-0.12, 1.42)
    for i,r in enumerate(recs):
        y = cy+i*(ih+0.12)
        rb = rrect(slide, 1.2, y, W-2.4, ih, T.card_rgb, radius_pct=10)
        if rb: shadow(rb, blur=8, dist=2, alpha=0.22)
        dot = oval(slide, W-2.55, y+(ih-0.36)/2, 0.36, 0.36, T.accent_rgb)
        if dot: gradient_fill(dot, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, r, 1.5, y+0.1, W-3.4, ih-0.2,
            font=_FONT, size=11.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_future(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=160)
    _header(slide, T, "آفاق البحث المستقبلية")
    items = req.future_work[:6]; cols = 2 if len(items)>3 else 1
    cy = 3.15; cw = (W-2.4-(cols-1)*0.35)/cols
    rows = (len(items)+cols-1)//cols
    ch = min((H-cy-0.5)/rows-0.22, 2.3)
    for i,it in enumerate(items):
        x = 1.2+(i%cols)*(cw+0.35); y = cy+(i//cols)*(ch+0.22)
        c = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=12)
        if c: shadow(c, blur=10, dist=3, alpha=0.28)
        ic = rrect(slide, x, y, 0.65, ch, T.accent_rgb, radius_pct=0)
        if ic: gradient_fill(ic, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, it, x+0.82, y+0.18, cw-1.0, ch-0.36,
            font=_FONT, size=11.5, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_references(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg_rgb, angle=160)
    _header(slide, T, "المراجع والمصادر")
    refs = req.references[:12]; cy = 3.15
    avail = H-cy-0.5; ih = max(min(avail/max(len(refs),1)-0.1, 1.15), 0.5)
    for i,r in enumerate(refs):
        y = cy+i*(ih+0.1)
        if y+ih > H-0.35: break
        if i%2==0:
            rb = rrect(slide, 1.2, y, W-2.4, ih, T.card_rgb, radius_pct=6)
        txt(slide, f"[{i+1}]", W-2.95, y+0.04, 1.5, ih-0.08,
            font="Calibri", size=10, bold=True, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, r, 1.5, y+0.04, W-4.8, ih-0.08,
            font=_FONT, size=10, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    return slide


def make_final(prs, req, T):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    # Rich multi-layer background
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=145)
    # Decorative elements
    oval(slide, -4, -4, 14, 14, T.accent_rgb, alpha=5)
    oval(slide, W-10, H-10, 16, 16, T.accent_rgb, alpha=4)
    oval(slide, W-4, -2, 7, 7, T.bg2_rgb, alpha=15)
    # Center card
    cw, ch = 23, 11; cx = (W-cw)/2; cy = (H-ch)/2
    c = rrect(slide, cx, cy, cw, ch, T.card_rgb, radius_pct=16)
    if c:
        soft_shadow(c, alpha=0.5)
        glow(c, T.accent, radius=14, alpha=0.1)
    # Top stripe
    ts = rrect(slide, cx, cy, cw, 0.38, T.accent_rgb, radius_pct=0)
    if ts: gradient_fill(ts, T.accent_grad1, T.accent_grad2, 0)
    # Thank you text
    txt(slide, "شكراً وتقديراً", cx+1.0, cy+0.55, cw-2.0, 2.8,
        font=_FONT, size=38, bold=True, color=T.text_light_rgb, align=PP_ALIGN.CENTER)
    # Gold divider with dots
    hline(slide, cx+cw*0.2, cy+3.45, cw*0.6, T.accent_rgb, thickness=0.06)
    d = oval(slide, (W-0.45)/2, cy+3.38, 0.45, 0.45, T.accent_rgb)
    if d: gradient_fill(d, T.accent_grad1, T.accent_grad2, 0)
    # Name
    txt(slide, req.student_name, cx+1.0, cy+3.65, cw-2.0, 1.3,
        font=_FONT, size=20, bold=True, color=T.accent_rgb, align=PP_ALIGN.CENTER)
    # Title
    td = req.title_ar[:75]+("..." if len(req.title_ar)>75 else "")
    txt(slide, td, cx+1.0, cy+5.1, cw-2.0, 2.2,
        font=_FONT, size=12, italic=True, color=T.muted_rgb, align=PP_ALIGN.CENTER)
    # Footer
    fp = "  ·  ".join(filter(None,[req.institution, req.year]))
    if fp:
        txt(slide, fp, cx+1.0, cy+ch-1.05, cw-2.0, 0.85,
            font=_FONT, size=11, color=T.muted_rgb, align=PP_ALIGN.CENTER)
    _accent_bar(slide, T, y=H-0.28, h=0.28)
    return slide
