"""
Engine B — CLASSIC ACADEMIC  v17.2
هوية: أكاديمية رسمية راقية — خلفية فاتحة، ألوان محترمة، جداول هرمية نظيفة
"""
from __future__ import annotations
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, cm, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, soft_shadow,
    set_solid_alpha, txt, blank_slide,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"
def set_font(f): global _FONT; _FONT = f

# Classic fixed palette — light background
_IVORY  = RGBColor(0xF4, 0xF2, 0xED)
_IVORY2 = RGBColor(0xE8, 0xE5, 0xDE)
_DARK   = RGBColor(0x15, 0x20, 0x35)
_DARK2  = RGBColor(0x2A, 0x38, 0x5C)
_GREY   = RGBColor(0x6A, 0x72, 0x80)
_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)


def _header(slide, T, title, sub=""):
    """Classic: solid dark header + right border accent + elegant line"""
    rect(slide, 0, 0, W, 2.8, _DARK2)
    # Thick accent left border
    ab = rect(slide, 0, 0, 0.65, 2.8, T.accent_rgb)
    if ab: gradient_fill(ab, T.accent_grad1, T.accent_grad2, 90)
    # Right decorative element — thin accent column
    rect(slide, W-0.25, 0, 0.25, 2.8, T.accent_rgb)
    # Bottom separator
    hline(slide, 0.65, 2.8, W-0.9, T.accent_rgb, thickness=0.07)
    hline(slide, 0.65, 2.87, W-0.9, _IVORY2, thickness=0.04)
    txt(slide, title, 0.88, 0.22, W-1.5, 1.55,
        font=_FONT, size=27, bold=True, color=_WHITE, align=PP_ALIGN.RIGHT)
    if sub:
        txt(slide, sub, 0.88, 1.7, W-1.5, 0.88,
            font=_FONT, size=12, color=T.accent_rgb, align=PP_ALIGN.RIGHT)


def make_cover(prs, req, T):
    slide = blank_slide(prs)
    # Ivory body
    rect(slide, 0, 0, W, H, _IVORY)
    # Left full sidebar
    sb = rect(slide, 0, 0, 1.5, H, _DARK2)
    # Accent strip inside sidebar
    rect(slide, 0, 0, 0.35, H, T.accent_rgb)
    # Top dark header band
    rect(slide, 1.5, 0, W-1.5, 3.8, _DARK2)
    # Bottom dark footer
    rect(slide, 1.5, H-1.1, W-1.5, 1.1, _DARK2)

    # Institution
    if req.institution:
        txt(slide, req.institution, 1.8, 0.22, W-2.2, 0.75,
            font=_FONT, size=12, color=T.accent_rgb, align=PP_ALIGN.RIGHT)

    # Title
    sz = 28 if len(req.title_ar)<50 else 22 if len(req.title_ar)<75 else 17
    txt(slide, req.title_ar, 1.8, 1.0, W-2.2, 2.5,
        font=_FONT, size=sz, bold=True, color=_WHITE, align=PP_ALIGN.RIGHT)

    # Gold separator line below title
    hline(slide, 1.8, 3.95, W-3.4, T.accent_rgb, thickness=0.09)
    hline(slide, 1.8, 4.08, W-3.4, _IVORY2, thickness=0.04)

    # Info table
    iy = 4.3; rh = 0.95
    rows = [("اسم الطالب", req.student_name),
            ("المشرف", req.supervisor),
            ("التخصص", req.specialization),
            ("السنة الدراسية", req.year)]
    for i,(lbl,val) in enumerate([(l,v) for l,v in rows if v]):
        bg_c = _IVORY2 if i%2==0 else _IVORY
        rect(slide, 1.5, iy, W-1.5, rh, bg_c)
        # Left label cell
        rect(slide, 1.5, iy, 6.5, rh, _IVORY2 if i%2!=0 else _IVORY)
        hline(slide, 1.5, iy, W-1.5, _IVORY2, thickness=0.04)
        vline(slide, 8.0, iy, rh, _DARK2, thickness=0.05)
        txt(slide, lbl, 1.65, iy+0.1, 6.2, rh-0.2,
            font=_FONT, size=12, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, val, 8.2, iy+0.1, W-9.0, rh-0.2,
            font=_FONT, size=13, color=_DARK, align=PP_ALIGN.RIGHT)
        iy += rh

    # Footer info
    if req.year:
        txt(slide, req.year, 1.8, H-1.05, W-3.4, 0.9,
            font="Calibri", size=14, bold=True, color=T.accent_rgb, align=PP_ALIGN.CENTER)
    return slide


def _body(slide, T):
    rect(slide, 0, 0, W, H, _IVORY)
    rect(slide, 0, 0, 0.65, H, _DARK2)
    rect(slide, 0, 0, 0.35, H, T.accent_rgb)
    rect(slide, W-0.25, 0, 0.25, H, _IVORY2)


def make_intro(prs, req, T):
    slide = blank_slide(prs); _body(slide, T)
    _header(slide, T, "مقدمة البحث", "")
    cy = 3.05
    for lbl, val in [(x,y) for x,y in [("نظرة عامة",req.intro_overview),("المنهج المتبع",req.intro_approach)] if y]:
        # Section label bar
        rect(slide, 0.9, cy, W-1.4, 0.6, _DARK2)
        rect(slide, 0.9, cy, 0.45, 0.6, T.accent_rgb)
        txt(slide, lbl, 1.55, cy, W-2.2, 0.6,
            font=_FONT, size=13, bold=True, color=_WHITE, align=PP_ALIGN.RIGHT)
        cy += 0.65
        rect(slide, 0.9, cy, W-1.4, 3.1, _IVORY2)
        hline(slide, 0.9, cy+3.1, W-1.4, _DARK2, thickness=0.04)
        txt(slide, val, 1.05, cy+0.12, W-1.7, 2.85,
            font=_FONT, size=12, color=_DARK, align=PP_ALIGN.RIGHT)
        cy += 3.38
    return slide


def make_plan(prs, req, T):
    slide = blank_slide(prs); _body(slide, T)
    _header(slide, T, "خطة البحث", f"يتكون البحث من {len(req.chapters)} فصول")
    cy = 3.05; chs = req.chapters[:8]
    avail = H-cy-0.35; rh = min(avail/max(len(chs),1)-0.08, 1.65)
    for i,ch in enumerate(chs):
        y = cy+i*(rh+0.08)
        bg_c = _IVORY if i%2==0 else _IVORY2
        rect(slide, 0.9, y, W-1.4, rh, bg_c)
        # Chapter number cell — full height
        rect(slide, 0.9, y, 2.0, rh, T.accent_rgb if i%2==0 else _DARK2)
        txt(slide, str(i+1), 0.9, y, 2.0, rh,
            font="Calibri", size=20, bold=True, color=_WHITE, align=PP_ALIGN.CENTER)
        vline(slide, 2.9, y, rh, _DARK2, thickness=0.05)
        # Title
        txt(slide, ch.title, 3.1, y, W-6.2, rh,
            font=_FONT, size=13, color=_DARK, align=PP_ALIGN.RIGHT)
        # Pages badge
        if ch.pages:
            rect(slide, W-3.5, y+(rh-0.42)/2, 2.35, 0.42, _DARK2)
            txt(slide, f"ص {ch.pages}", W-3.5, y+(rh-0.42)/2, 2.35, 0.42,
                font="Calibri", size=10, bold=True, color=T.accent_rgb, align=PP_ALIGN.CENTER)
        hline(slide, 0.9, y+rh, W-1.4, _DARK2, thickness=0.03)
    return slide


def make_problem(prs, req, T):
    slide = blank_slide(prs); _body(slide, T)
    _header(slide, T, "إشكالية البحث")
    cy = 3.05
    if req.main_problem:
        rect(slide, 0.9, cy, W-1.4, 2.0, _IVORY2)
        rect(slide, 0.9, cy, 0.55, 2.0, T.accent_rgb)
        rect(slide, 0.9, cy, W-1.4, 0.52, _DARK2)
        txt(slide, "الإشكالية الرئيسية", 1.6, cy, W-2.2, 0.52,
            font=_FONT, size=12, bold=True, color=_WHITE, align=PP_ALIGN.RIGHT)
        txt(slide, req.main_problem, 1.65, cy+0.6, W-2.3, 1.28,
            font=_FONT, size=12.5, color=_DARK, align=PP_ALIGN.RIGHT)
        cy += 2.15
    if req.main_question:
        hline(slide, 0.9, cy, W-1.4, T.accent_rgb, thickness=0.07)
        cy += 0.25
        rect(slide, 0.9, cy, 0.55, 1.35, T.accent_rgb)
        txt(slide, "التساؤل الرئيسي", 1.6, cy, 8.0, 0.55,
            font=_FONT, size=11, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, req.main_question, 1.6, cy+0.58, W-2.3, 1.1,
            font=_FONT, size=13, bold=True, italic=True, color=_DARK2, align=PP_ALIGN.RIGHT)
        cy += 1.6
    for i,q in enumerate(req.sub_questions[:4]):
        y = cy+i*0.82
        rect(slide, 0.9, y, W-1.4, 0.78, _IVORY2 if i%2==0 else _IVORY)
        txt(slide, f"{i+1}.", 0.92, y+0.06, 1.2, 0.66,
            font="Calibri", size=13, bold=True, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, q, 2.3, y+0.06, W-3.1, 0.66,
            font=_FONT, size=11, color=_DARK, align=PP_ALIGN.RIGHT)
    return slide


def make_objectives(prs, req, T):
    slide = blank_slide(prs); _body(slide, T)
    _header(slide, T, "أهداف البحث وفرضياته")
    cy = 3.05; cw = (W-1.8)/2
    for ci,(lbl,items) in enumerate([(x,y) for x,y in [("الأهداف",req.objectives),("الفرضيات",req.hypotheses)] if y][:2]):
        x = 0.9+ci*(cw+0.0)
        rect(slide, x, cy, cw, 0.7, T.accent_rgb if ci==0 else _DARK2)
        txt(slide, lbl, x, cy, cw, 0.7,
            font=_FONT, size=15, bold=True, color=_WHITE, align=PP_ALIGN.CENTER)
        ih = min((H-cy-1.1)/max(len(items),1)-0.06, 1.0)
        for j,it in enumerate(items[:8]):
            iy = cy+0.76+j*(ih+0.06)
            bg_c = _IVORY2 if j%2==0 else _IVORY
            rect(slide, x, iy, cw, ih, bg_c)
            vline(slide, x, iy, ih, _DARK2, thickness=0.04)
            txt(slide, f"{j+1}.", x+0.12, iy+0.06, 1.4, ih-0.12,
                font="Calibri", size=12, bold=True, color=T.accent_rgb, align=PP_ALIGN.LEFT)
            txt(slide, it, x+1.6, iy+0.06, cw-1.75, ih-0.12,
                font=_FONT, size=10.5, color=_DARK, align=PP_ALIGN.RIGHT)
            hline(slide, x, iy+ih, cw, _DARK2, thickness=0.03)
        if ci==0: vline(slide, x+cw, cy, H-cy, _DARK2, thickness=0.06)
    return slide


def make_importance(prs, req, T):
    slide = blank_slide(prs); _body(slide, T)
    _header(slide, T, "أهمية البحث")
    items = (req.importance+([req.reasons] if req.reasons else []))[:8]
    cy = 3.05; ih = min((H-cy-0.35)/max(len(items),1)-0.07, 1.1)
    for i,it in enumerate(items):
        y = cy+i*(ih+0.07)
        rect(slide, 0.9, y, W-1.4, ih, _IVORY2 if i%2==0 else _IVORY)
        rect(slide, 0.9, y, 0.55, ih, T.accent_rgb)
        txt(slide, str(i+1), 0.9, y, 0.55, ih,
            font="Calibri", size=12, bold=True, color=_WHITE, align=PP_ALIGN.CENTER)
        txt(slide, it, 1.65, y+0.08, W-2.25, ih-0.16,
            font=_FONT, size=11.5, color=_DARK, align=PP_ALIGN.RIGHT)
        hline(slide, 0.9, y+ih, W-1.4, _IVORY2, thickness=0.04)
    return slide


def make_methodology(prs, req, T):
    slide = blank_slide(prs); _body(slide, T)
    _header(slide, T, "منهجية البحث")
    fields = [(l,v) for l,v in [("المنهج",req.methodology),("نوع العينة",req.sample_type),
              ("حجم العينة",req.sample_size),("أداة الدراسة",req.tool)] if v]
    cy = 3.05; rh = min((H-cy-0.35)/max(len(fields),1)-0.08, 1.75)
    for i,(l,v) in enumerate(fields[:4]):
        y = cy+i*(rh+0.08)
        rect(slide, 0.9, y, W-1.4, rh, _IVORY2 if i%2==0 else _IVORY)
        # Label column
        lw = 7.0
        rect(slide, 0.9, y, lw, rh, _IVORY2 if i%2!=0 else _IVORY)
        vline(slide, 0.9+lw, y, rh, T.accent_rgb, thickness=0.08)
        txt(slide, l, 1.0, y+0.12, lw-0.25, rh-0.24,
            font=_FONT, size=13, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, v, 0.9+lw+0.2, y+0.12, W-lw-1.7, rh-0.24,
            font=_FONT, size=13, color=_DARK, align=PP_ALIGN.RIGHT)
        hline(slide, 0.9, y+rh, W-1.4, _DARK2, thickness=0.04)
    return slide


def make_stats(prs, req, T):
    slide = blank_slide(prs); _body(slide, T)
    _header(slide, T, "الإحصاءات والأرقام")
    stats = req.stats[:6]; n = len(stats)
    if not n: return slide
    cols = 3 if n>=3 else n; cy = 3.05; gap = 0.4
    cw = (W-1.6-(cols-1)*gap)/cols
    rows = (n+cols-1)//cols; ch = min((H-cy-0.35)/rows-gap, 3.5)
    for i,st in enumerate(stats):
        x = 0.9+(i%cols)*(cw+gap); y = cy+(i//cols)*(ch+gap)
        rect(slide, x, y, cw, ch, _IVORY2)
        # Header band
        rect(slide, x, y, cw, 0.55, T.accent_rgb if (i%2==0) else _DARK2)
        txt(slide, st.label, x+0.1, y, cw-0.2, 0.55,
            font=_FONT, size=11, bold=True, color=_WHITE, align=PP_ALIGN.CENTER)
        # Big value
        vsz = 40 if len(st.value)<=4 else 28 if len(st.value)<=8 else 20
        txt(slide, st.value, x+0.1, y+0.6, cw-0.2, ch-1.1,
            font="Calibri", size=vsz, bold=True, color=T.accent_rgb, align=PP_ALIGN.CENTER)
        if st.unit:
            txt(slide, st.unit, x+0.1, y+ch-0.62, cw-0.2, 0.52,
                font=_FONT, size=10, color=_GREY, align=PP_ALIGN.CENTER)
        # Bottom border
        hline(slide, x, y+ch, cw, T.accent_rgb, thickness=0.06)
    return slide


def make_results(prs, req, T):
    slide = blank_slide(prs); _body(slide, T)
    _header(slide, T, "نتائج البحث")
    results = req.main_results[:8]; cy = 3.05
    ih = min((H-cy-0.35)/max(len(results),1)-0.08, 1.1)
    for i,r in enumerate(results):
        y = cy+i*(ih+0.08)
        rect(slide, 0.9, y, W-1.4, ih, _IVORY2 if i%2==0 else _IVORY)
        # Number cell
        rect(slide, W-2.8, y, 1.9, ih, T.accent_rgb if i%2==0 else _DARK2)
        txt(slide, str(i+1), W-2.8, y, 1.9, ih,
            font="Calibri", size=16, bold=True, color=_WHITE, align=PP_ALIGN.CENTER)
        txt(slide, r, 1.0, y+0.08, W-4.0, ih-0.16,
            font=_FONT, size=11.5, color=_DARK, align=PP_ALIGN.RIGHT)
        hline(slide, 0.9, y+ih, W-1.4, _DARK2, thickness=0.04)
    return slide


def make_conclusion(prs, req, T):
    slide = blank_slide(prs); _body(slide, T)
    _header(slide, T, "خاتمة البحث")
    cy = 3.05
    rect(slide, 0.9, cy, W-1.4, H-cy-0.4, _IVORY2)
    rect(slide, 0.9, cy, 0.55, H-cy-0.4, T.accent_rgb)
    hline(slide, 0.9, cy, W-1.4, T.accent_rgb, thickness=0.08)
    hline(slide, 0.9, H-0.4, W-1.4, T.accent_rgb, thickness=0.08)
    txt(slide, req.general_conclusion, 1.65, cy+0.3, W-2.25, H-cy-1.0,
        font=_FONT, size=14, color=_DARK2, align=PP_ALIGN.RIGHT, spacing=22)
    return slide


def make_recommendations(prs, req, T):
    slide = blank_slide(prs); _body(slide, T)
    _header(slide, T, "توصيات البحث")
    recs = req.recommendations[:8]; cy = 3.05
    ih = min((H-cy-0.35)/max(len(recs),1)-0.08, 1.1)
    for i,r in enumerate(recs):
        y = cy+i*(ih+0.08)
        rect(slide, 0.9, y, W-1.4, ih, _IVORY2 if i%2==0 else _IVORY)
        rect(slide, 0.9, y, 0.55, ih, T.accent_rgb)
        txt(slide, str(i+1), 0.9, y, 0.55, ih,
            font="Calibri", size=12, bold=True, color=_WHITE, align=PP_ALIGN.CENTER)
        txt(slide, r, 1.65, y+0.08, W-2.25, ih-0.16,
            font=_FONT, size=11.5, color=_DARK, align=PP_ALIGN.RIGHT)
        hline(slide, 0.9, y+ih, W-1.4, _IVORY2, thickness=0.04)
    return slide


def make_future(prs, req, T):
    slide = blank_slide(prs); _body(slide, T)
    _header(slide, T, "آفاق البحث المستقبلية")
    items = req.future_work[:6]; cy = 3.05
    ih = min((H-cy-0.35)/max(len(items),1)-0.08, 1.55)
    for i,it in enumerate(items):
        y = cy+i*(ih+0.08)
        rect(slide, 0.9, y, W-1.4, ih, _IVORY2 if i%2==0 else _IVORY)
        hline(slide, 0.9, y, W-1.4, T.accent_rgb, thickness=0.06)
        txt(slide, "◈", 0.92, y+0.1, 1.5, ih-0.2,
            font="Calibri", size=20, bold=True, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, it, 2.6, y+0.1, W-3.35, ih-0.2,
            font=_FONT, size=12.5, color=_DARK2, align=PP_ALIGN.RIGHT)
    return slide


def make_references(prs, req, T):
    slide = blank_slide(prs); _body(slide, T)
    _header(slide, T, "المراجع والمصادر")
    refs = req.references[:12]; cy = 3.05
    ih = max(min((H-cy-0.35)/max(len(refs),1)-0.06, 1.05), 0.48)
    for i,r in enumerate(refs):
        y = cy+i*(ih+0.06)
        if y+ih > H-0.32: break
        rect(slide, 0.9, y, W-1.4, ih, _IVORY2 if i%2==0 else _IVORY)
        txt(slide, f"[{i+1}]", W-2.8, y+0.04, 1.7, ih-0.08,
            font="Calibri", size=10, bold=True, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, r, 1.0, y+0.04, W-4.0, ih-0.08,
            font=_FONT, size=10, color=_DARK, align=PP_ALIGN.RIGHT)
        hline(slide, 0.9, y+ih, W-1.4, _IVORY2, thickness=0.04)
    return slide


def make_final(prs, req, T):
    slide = blank_slide(prs)
    rect(slide, 0, 0, W, H, _IVORY)
    # Left dark panel
    rect(slide, 0, 0, W*0.40, H, _DARK2)
    rect(slide, 0, 0, 0.55, H, T.accent_rgb)
    # Right ivory panel
    rect(slide, W*0.40, 0, W*0.60, H, _IVORY)
    # Accent top/bottom strips on right
    rect(slide, W*0.40, 0, W*0.60, 0.55, _DARK2)
    rect(slide, W*0.40, H-0.55, W*0.60, 0.55, _DARK2)

    # Left panel text
    txt(slide, "شكراً", 0.2, H*0.22, W*0.38, 3.5,
        font=_FONT, size=44, bold=True, color=_WHITE, align=PP_ALIGN.CENTER)
    txt(slide, "وتقديراً", 0.2, H*0.22+3.2, W*0.38, 2.0,
        font=_FONT, size=26, color=T.accent_rgb, align=PP_ALIGN.CENTER)

    # Gold divider line
    hline(slide, W*0.40, H/2-0.04, W*0.60, T.accent_rgb, thickness=0.08)

    # Right panel info
    rx = W*0.40+0.8; rw = W*0.55
    txt(slide, req.student_name, rx, H*0.28, rw, 1.5,
        font=_FONT, size=24, bold=True, color=_DARK2, align=PP_ALIGN.RIGHT)
    hline(slide, rx, H*0.28+1.6, rw, T.accent_rgb, thickness=0.07)
    td = req.title_ar[:85]+("..." if len(req.title_ar)>85 else "")
    txt(slide, td, rx, H*0.28+1.85, rw, 2.8,
        font=_FONT, size=13, color=_DARK, align=PP_ALIGN.RIGHT)
    fp = " | ".join(filter(None,[req.institution, req.year]))
    if fp:
        txt(slide, fp, rx, H-1.45, rw, 0.85,
            font=_FONT, size=11, color=_GREY, align=PP_ALIGN.RIGHT)
    return slide
