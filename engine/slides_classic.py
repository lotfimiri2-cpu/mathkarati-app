"""
Engine B — CLASSIC ACADEMIC
هوية بصرية: أكاديمية رسمية
- خلفية فاتحة (ivory/white) مع نص داكن
- خطوط أفقية فاصلة واضحة
- جداول وقوائم هرمية
- شريط جانبي ملون يسار
- typography رسمي، لا rounded corners
- layout: قائمة مرقمة كلاسيكية
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

# Classic uses light backgrounds — override theme colors for readability
_BG    = RGBColor(0xF5, 0xF5, 0xF0)  # ivory
_BG2   = RGBColor(0xE8, 0xE8, 0xE2)  # light grey
_DARK  = RGBColor(0x1A, 0x1A, 0x2E)  # near black
_DARK2 = RGBColor(0x2D, 0x3A, 0x5A)  # dark blue
_MID   = RGBColor(0x55, 0x60, 0x70)  # grey


def _hx(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))


def _header(slide, T, title, sub=""):
    """CLASSIC header: solid dark band, left sidebar accent, clean lines"""
    # Full top bar — solid dark
    rect(slide, 0, 0, W, 2.6, _DARK2)
    # Left accent sidebar
    rect(slide, 0, 0, 0.55, 2.6, T.accent_rgb)
    # Bottom border line on header
    hline(slide, 0, 2.6, W, T.accent_rgb, thickness=0.08)
    txt(slide, title, 0.75, 0.22, W-1.5, 1.4,
        font=_FONT, size=26, bold=True, color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    if sub:
        txt(slide, sub, 0.75, 1.55, W-1.5, 0.85,
            font=_FONT, size=12, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    # Body separator
    hline(slide, 1.0, 2.85, W-2.0, T.accent_rgb, thickness=0.05)


def make_cover(prs, req, T):
    slide = blank_slide(prs)
    # Light ivory background
    rect(slide, 0, 0, W, H, _BG)
    # Left full-height accent bar
    rect(slide, 0, 0, 1.2, H, T.accent_rgb)
    # Thin gold line right of sidebar
    hline(slide, 1.2, 0, 0.06, _DARK, thickness=H)
    # Top dark header band
    rect(slide, 1.2, 0, W-1.2, 3.5, _DARK2)
    # Institution line
    if req.institution:
        txt(slide, req.institution, 1.5, 0.25, W-2.0, 0.75,
            font=_FONT, size=12, color=T.muted_rgb, align=PP_ALIGN.RIGHT)
    # Title area
    txt(slide, req.title_ar, 1.5, 1.0, W-2.0, 2.2,
        font=_FONT, size=26 if len(req.title_ar)<60 else 20, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT)
    # Separator
    hline(slide, 1.5, 3.7, W-3.0, T.accent_rgb, thickness=0.08)
    hline(slide, 1.5, 3.85, W-3.0, _BG2, thickness=0.04)
    # Info table
    iy = 4.1; rh = 0.9
    rows = [("اسم الطالب", req.student_name),
            ("المشرف", req.supervisor),
            ("التخصص", req.specialization or req.institution),
            ("السنة الجامعية", req.year)]
    for lbl, val in [(l,v) for l,v in rows if v]:
        rect(slide, 1.5, iy, W-3.0, rh-0.08, _BG2)
        hline(slide, 1.5, iy, W-3.0, _BG2, thickness=0.02)
        txt(slide, lbl+" :", 1.6, iy+0.1, 6.0, rh-0.18,
            font=_FONT, size=11, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, val, 7.8, iy+0.1, W-9.5, rh-0.18,
            font=_FONT, size=12, color=_DARK, align=PP_ALIGN.RIGHT)
        iy += rh
    # Bottom bar
    rect(slide, 1.2, H-0.9, W-1.2, 0.9, _DARK2)
    if req.year:
        txt(slide, req.year, 1.5, H-0.9, W-3.0, 0.9,
            font="Calibri", size=13, bold=True, color=T.accent_rgb, align=PP_ALIGN.CENTER)
    return slide


def _body_bg(slide):
    rect(slide, 0, 0, W, H, _BG)
    rect(slide, 0, 0, 0.55, H, _DARK2)


def make_intro(prs, req, T):
    slide = blank_slide(prs); _body_bg(slide)
    _header(slide, T, "مقدمة البحث", "")
    cy = 3.1
    for lbl, val in [(x,y) for x,y in [("نظرة عامة:", req.intro_overview),("المنهج المتبع:", req.intro_approach)] if y]:
        txt(slide, lbl, 0.85, cy, W-1.7, 0.65,
            font=_FONT, size=13, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        cy += 0.65
        rect(slide, 0.75, cy, W-1.5, 0.04, T.accent_rgb)
        cy += 0.12
        txt(slide, val, 0.85, cy, W-1.7, 3.5,
            font=_FONT, size=12, color=_DARK, align=PP_ALIGN.RIGHT)
        cy += 2.8
    return slide


def make_plan(prs, req, T):
    slide = blank_slide(prs); _body_bg(slide)
    _header(slide, T, "خطة البحث", f"يتكون البحث من {len(req.chapters)} فصول")
    cy = 3.05; chs = req.chapters[:8]
    avail = H-cy-0.4; rh = min(avail/max(len(chs),1)-0.08, 1.6)
    for i, ch in enumerate(chs):
        y = cy + i*(rh+0.08)
        # Row bg alternating
        row_color = _BG2 if i%2==0 else _BG
        rect(slide, 0.75, y, W-1.5, rh, row_color)
        # Left number
        rect(slide, 0.75, y, 1.8, rh, T.accent_rgb)
        txt(slide, str(i+1), 0.75, y, 1.8, rh,
            font="Calibri", size=18, bold=True, color=_BG, align=PP_ALIGN.CENTER)
        # Chapter title
        txt(slide, ch.title, 2.75, y, W-5.8, rh,
            font=_FONT, size=13, color=_DARK, align=PP_ALIGN.RIGHT)
        if ch.pages:
            txt(slide, f"ص {ch.pages}", W-3.4, y, 2.5, rh,
                font="Calibri", size=11, color=_MID, align=PP_ALIGN.LEFT)
        hline(slide, 0.75, y+rh, W-1.5, _BG2, thickness=0.04)
    return slide


def make_problem(prs, req, T):
    slide = blank_slide(prs); _body_bg(slide)
    _header(slide, T, "إشكالية البحث")
    cy = 3.05
    if req.main_problem:
        rect(slide, 0.75, cy, W-1.5, 1.9, _BG2)
        rect(slide, 0.75, cy, 0.5, 1.9, T.accent_rgb)
        txt(slide, "الإشكالية:", 1.45, cy+0.1, 5.0, 0.6,
            font=_FONT, size=11, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, req.main_problem, 1.45, cy+0.65, W-2.7, 1.15,
            font=_FONT, size=12, color=_DARK, align=PP_ALIGN.RIGHT)
        cy += 2.05
    if req.main_question:
        hline(slide, 0.75, cy, W-1.5, T.accent_rgb, thickness=0.06)
        cy += 0.22
        txt(slide, "التساؤل الرئيسي:", 0.85, cy, 6.0, 0.6,
            font=_FONT, size=12, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        cy += 0.6
        txt(slide, req.main_question, 0.85, cy, W-1.7, 1.2,
            font=_FONT, size=13, bold=True, italic=True, color=_DARK2, align=PP_ALIGN.RIGHT)
        cy += 1.35
    for i,q in enumerate(req.sub_questions[:5]):
        txt(slide, f"{i+1}. {q}", 0.85, cy+i*0.78, W-1.7, 0.75,
            font=_FONT, size=11, color=_DARK, align=PP_ALIGN.RIGHT)
    return slide


def make_objectives(prs, req, T):
    slide = blank_slide(prs); _body_bg(slide)
    _header(slide, T, "أهداف البحث وفرضياته")
    cy = 3.05; cw = (W-2.0)/2
    for ci, (lbl, items) in enumerate([(x,y) for x,y in [("الأهداف",req.objectives),("الفرضيات",req.hypotheses)] if y][:2]):
        x = 0.75 + ci*(cw+0.5)
        rect(slide, x, cy, cw, 0.65, T.accent_rgb)
        txt(slide, lbl, x, cy, cw, 0.65,
            font=_FONT, size=14, bold=True, color=_BG, align=PP_ALIGN.CENTER)
        ih = min((H-cy-1.3)/max(len(items),1), 0.95)
        for j, it in enumerate(items[:8]):
            iy = cy+0.72+j*ih
            row_c = _BG2 if j%2==0 else _BG
            rect(slide, x, iy, cw, ih-0.04, row_c)
            txt(slide, f"{j+1}. {it}", x+0.15, iy, cw-0.3, ih-0.04,
                font=_FONT, size=10.5, color=_DARK, align=PP_ALIGN.RIGHT)
    return slide


def make_importance(prs, req, T):
    slide = blank_slide(prs); _body_bg(slide)
    _header(slide, T, "أهمية البحث")
    items = (req.importance + ([req.reasons] if req.reasons else []))[:8]
    cy = 3.05; ih = min((H-cy-0.4)/max(len(items),1)-0.08, 1.1)
    for i, it in enumerate(items):
        y = cy + i*(ih+0.08)
        rect(slide, 0.75, y, W-1.5, ih, _BG2 if i%2==0 else _BG)
        rect(slide, 0.75, y, 0.45, ih, T.accent_rgb)
        txt(slide, f"{i+1}", 0.75, y, 0.45, ih,
            font="Calibri", size=11, bold=True, color=_BG, align=PP_ALIGN.CENTER)
        txt(slide, it, 1.4, y+0.06, W-2.25, ih-0.12,
            font=_FONT, size=11.5, color=_DARK, align=PP_ALIGN.RIGHT)
    return slide


def make_methodology(prs, req, T):
    slide = blank_slide(prs); _body_bg(slide)
    _header(slide, T, "منهجية البحث")
    cy = 3.05
    fields = [(l,v) for l,v in [("المنهج المعتمد",req.methodology),("نوع العينة",req.sample_type),
              ("حجم العينة",req.sample_size),("أداة الدراسة",req.tool)] if v]
    for i,(l,v) in enumerate(fields[:4]):
        y = cy + i*1.5
        rect(slide, 0.75, y, W-1.5, 1.4, _BG2 if i%2==0 else _BG)
        rect(slide, 0.75, y, 5.5, 1.4, _BG2 if i%2!=0 else _BG)
        hline(slide, 0.75+5.5, y, 0.06, _DARK2, thickness=1.4)
        txt(slide, l, 0.85, y+0.1, 5.3, 1.2,
            font=_FONT, size=13, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
        txt(slide, v, 6.55, y+0.1, W-7.55, 1.2,
            font=_FONT, size=12, color=_DARK, align=PP_ALIGN.RIGHT)
    return slide


def make_stats(prs, req, T):
    slide = blank_slide(prs); _body_bg(slide)
    _header(slide, T, "الإحصاءات والأرقام الرئيسية")
    stats = req.stats[:6]; n = len(stats)
    if not n: return slide
    cols = 3 if n>=3 else n; cy = 3.05; gap = 0.4
    cw = (W-1.5-(cols-1)*gap)/cols
    rows = (n+cols-1)//cols; ch = min((H-cy-0.4)/rows-gap, 3.2)
    for i,st in enumerate(stats):
        x = 0.75+(i%cols)*(cw+gap); y = cy+(i//cols)*(ch+gap)
        rect(slide, x, y, cw, ch, _BG2)
        rect(slide, x, y, cw, 0.5, T.accent_rgb)
        txt(slide, st.label, x+0.1, y, cw-0.2, 0.5,
            font=_FONT, size=10, bold=True, color=_BG, align=PP_ALIGN.CENTER)
        vsz = 36 if len(st.value)<=4 else 26 if len(st.value)<=8 else 18
        txt(slide, st.value, x+0.1, y+0.55, cw-0.2, ch-0.85,
            font="Calibri", size=vsz, bold=True, color=T.accent_rgb, align=PP_ALIGN.CENTER)
        if st.unit:
            txt(slide, st.unit, x+0.1, y+ch-0.65, cw-0.2, 0.55,
                font=_FONT, size=10, color=_MID, align=PP_ALIGN.CENTER)
    return slide


def make_results(prs, req, T):
    slide = blank_slide(prs); _body_bg(slide)
    _header(slide, T, "نتائج البحث")
    results = req.main_results[:8]; cy = 3.05
    ih = min((H-cy-0.4)/max(len(results),1)-0.08, 1.1)
    for i,r in enumerate(results):
        y = cy+i*(ih+0.08)
        rect(slide, 0.75, y, W-1.5, ih, _BG2 if i%2==0 else _BG)
        rect(slide, W-2.55, y, 1.8, ih, T.accent_rgb)
        txt(slide, str(i+1), W-2.55, y, 1.8, ih,
            font="Calibri", size=14, bold=True, color=_BG, align=PP_ALIGN.CENTER)
        txt(slide, r, 0.85, y+0.06, W-3.7, ih-0.12,
            font=_FONT, size=11.5, color=_DARK, align=PP_ALIGN.RIGHT)
    return slide


def make_conclusion(prs, req, T):
    slide = blank_slide(prs); _body_bg(slide)
    _header(slide, T, "خاتمة البحث")
    cy = 3.1
    rect(slide, 0.75, cy, W-1.5, H-cy-0.5, _BG2)
    rect(slide, 0.75, cy, 0.5, H-cy-0.5, T.accent_rgb)
    hline(slide, 0.75, cy, W-1.5, T.accent_rgb, thickness=0.08)
    hline(slide, 0.75, H-0.5, W-1.5, T.accent_rgb, thickness=0.08)
    txt(slide, req.general_conclusion, 1.45, cy+0.3, W-2.7, H-cy-1.1,
        font=_FONT, size=14, color=_DARK2, align=PP_ALIGN.RIGHT)
    return slide


def make_recommendations(prs, req, T):
    slide = blank_slide(prs); _body_bg(slide)
    _header(slide, T, "توصيات البحث")
    recs = req.recommendations[:8]; cy = 3.05
    ih = min((H-cy-0.4)/max(len(recs),1)-0.08, 1.1)
    for i,r in enumerate(recs):
        y = cy+i*(ih+0.08)
        rect(slide, 0.75, y, W-1.5, ih, _BG2 if i%2==0 else _BG)
        rect(slide, 0.75, y, 0.45, ih, T.accent_rgb)
        txt(slide, str(i+1), 0.75, y, 0.45, ih,
            font="Calibri", size=11, bold=True, color=_BG, align=PP_ALIGN.CENTER)
        txt(slide, r, 1.4, y+0.06, W-2.25, ih-0.12,
            font=_FONT, size=11.5, color=_DARK, align=PP_ALIGN.RIGHT)
    return slide


def make_future(prs, req, T):
    slide = blank_slide(prs); _body_bg(slide)
    _header(slide, T, "آفاق البحث المستقبلية")
    items = req.future_work[:6]; cy = 3.05
    ih = min((H-cy-0.4)/max(len(items),1)-0.08, 1.5)
    for i,it in enumerate(items):
        y = cy+i*(ih+0.08)
        rect(slide, 0.75, y, W-1.5, ih, _BG2 if i%2==0 else _BG)
        hline(slide, 0.75, y, W-1.5, T.accent_rgb, thickness=0.04)
        txt(slide, f"▸  {it}", 0.85, y+0.1, W-1.7, ih-0.2,
            font=_FONT, size=12, color=_DARK2, align=PP_ALIGN.RIGHT)
    return slide


def make_references(prs, req, T):
    slide = blank_slide(prs); _body_bg(slide)
    _header(slide, T, "المراجع والمصادر")
    refs = req.references[:12]; cy = 3.05
    ih = max(min((H-cy-0.4)/max(len(refs),1)-0.06, 1.0), 0.48)
    for i,r in enumerate(refs):
        y = cy+i*(ih+0.06)
        if y+ih > H-0.3: break
        rect(slide, 0.75, y, W-1.5, ih, _BG2 if i%2==0 else _BG)
        txt(slide, f"[{i+1}]", W-2.7, y+0.04, 1.8, ih-0.08,
            font="Calibri", size=10, bold=True, color=T.accent_rgb, align=PP_ALIGN.LEFT)
        txt(slide, r, 0.85, y+0.04, W-3.9, ih-0.08,
            font=_FONT, size=10, color=_DARK, align=PP_ALIGN.RIGHT)
    return slide


def make_final(prs, req, T):
    slide = blank_slide(prs)
    rect(slide, 0, 0, W, H, _BG)
    # Left full accent panel
    rect(slide, 0, 0, W*0.38, H, T.accent_rgb)
    # Right dark panel
    rect(slide, W*0.38, 0, W*0.62, H, _DARK2)
    # Arabic text on accent panel
    txt(slide, "شكراً", 0.3, H*0.25, W*0.36, 3.0,
        font=_FONT, size=42, bold=True, color=_BG, align=PP_ALIGN.CENTER)
    txt(slide, "وتقديراً", 0.3, H*0.25+2.8, W*0.36, 1.8,
        font=_FONT, size=28, bold=False, color=_DARK2, align=PP_ALIGN.CENTER)
    # Info on dark panel
    ix = W*0.38+0.8
    txt(slide, req.student_name, ix, H*0.28, W*0.58, 1.4,
        font=_FONT, size=22, bold=True, color=T.accent_rgb, align=PP_ALIGN.RIGHT)
    hline(slide, ix, H*0.28+1.5, W*0.56, T.accent_rgb, thickness=0.06)
    title_d = req.title_ar[:80]+("..." if len(req.title_ar)>80 else "")
    txt(slide, title_d, ix, H*0.28+1.7, W*0.56, 2.5,
        font=_FONT, size=13, color=RGBColor(0xCC,0xCC,0xCC), align=PP_ALIGN.RIGHT)
    fp = "  ·  ".join(filter(None,[req.institution, req.year]))
    if fp:
        txt(slide, fp, ix, H-1.5, W*0.56, 0.9,
            font=_FONT, size=11, color=RGBColor(0x88,0x88,0x88), align=PP_ALIGN.RIGHT)
    return slide
