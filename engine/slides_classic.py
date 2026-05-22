"""
Slide Builder — Classic Engine — مذكرتي Pro v18
Layout philosophy: Bold, structured, academic. Left-aligned numbers,
horizontal dividers, wide content panels, minimal decoration.
Pure functions: (prs, request, theme) → slide
"""
from __future__ import annotations

from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from engine.primitives import (
    W, H, cm, pt,
    rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow,
    set_solid_alpha,
    txt, blank_slide,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"


def set_font(font_name: str):
    global _FONT
    _FONT = font_name


def _hx(h: str) -> RGBColor:
    h = h.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ══════════════════════════════════════════════════════════════════════
# CLASSIC COVER — Split layout: left accent panel + right content
# ══════════════════════════════════════════════════════════════════════
def make_cover(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)

    # Left thick accent panel
    panel_w = 7.5
    left_panel = rect(slide, 0, 0, panel_w, H, T.accent_rgb)
    if left_panel:
        gradient_fill(left_panel, T.accent_grad1, T.accent_grad2, angle=180)

    # Left panel decorative circles
    oval(slide, -2, H * 0.6, 8, 8, T.bg_rgb, alpha=12)
    oval(slide, 1, -2, 5, 5, T.bg_rgb, alpha=8)

    # Left panel: year text vertical
    if req.year:
        txt(slide, req.year, 0.4, H - 1.2, panel_w - 0.8, 0.9,
            font="Calibri", size=11, bold=False,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

    # Left panel: institution
    if req.institution:
        txt(slide, req.institution, 0.3, 0.6, panel_w - 0.6, 1.0,
            font=_FONT, size=10, bold=False,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # Right content area
    content_x = panel_w + 1.2
    content_w = W - content_x - 1.0

    # Top horizontal rule
    hline(slide, content_x, 2.5, content_w, T.accent_rgb, thickness=0.08)

    # Title — large, right-aligned
    title_size = 30 if len(req.title_ar) < 45 else 22 if len(req.title_ar) < 75 else 17
    txt(slide, req.title_ar,
        content_x, 2.8, content_w, H * 0.38,
        font=_FONT, size=title_size, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    if req.title_en:
        txt(slide, req.title_en,
            content_x, 2.8 + H * 0.38 - 0.5, content_w, 0.8,
            font="Calibri", size=11, bold=False, italic=True,
            color=T.muted_rgb, align=PP_ALIGN.LEFT, rtl=False)

    # Bottom horizontal rule
    info_y = H * 0.62
    hline(slide, content_x, info_y - 0.2, content_w, T.accent_rgb, thickness=0.04)

    # Info rows — clean label/value pairs
    row_h = 0.72
    fields = []
    fields.append(("الطالب", req.student_name))
    if req.supervisor:
        fields.append(("المشرف", req.supervisor))
    if req.co_supervisor:
        fields.append(("المشرف المساعد", req.co_supervisor))
    if req.specialization:
        fields.append(("التخصص", req.specialization))

    for i, (lbl, val) in enumerate(fields):
        y = info_y + i * row_h
        txt(slide, lbl + " :", content_x, y, 5.0, row_h,
            font=_FONT, size=11, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, val, content_x + 5.2, y, content_w - 5.2, row_h,
            font=_FONT, size=12, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# CLASSIC SECTION HEADER — Bottom bar + large number
# ══════════════════════════════════════════════════════════════════════
def _section_header(slide, T: Theme, title: str, subtitle: str = "", num: str = ""):
    # Full-width top bar (thin)
    top_bar = rect(slide, 0, 0, W, 3.0, T.bg2_rgb)
    if top_bar:
        gradient_fill(top_bar, T.grad1, T.grad2, angle=0)

    # Bold left number badge
    if num:
        num_bg = rrect(slide, 0.8, 0.55, 2.0, 2.0, T.accent_rgb, radius_pct=8)
        if num_bg:
            gradient_fill(num_bg, T.accent_grad1, T.accent_grad2, angle=135)
            shadow(num_bg, blur=12, dist=4, alpha=0.4)
        txt(slide, num, 0.8, 0.55, 2.0, 2.0,
            font="Calibri", size=28, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

    # Title
    title_x = 3.2 if num else 0.8
    txt(slide, title, title_x, 0.5, W - title_x - 0.8, 1.5,
        font=_FONT, size=24, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    if subtitle:
        txt(slide, subtitle, title_x, 1.9, W - title_x - 0.8, 0.8,
            font=_FONT, size=12, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    # Bottom divider of header area
    hline(slide, 0.8, 2.85, W - 1.6, T.accent_rgb, thickness=0.1)


# ══════════════════════════════════════════════════════════════════════
# INTRO — Two full-height panels
# ══════════════════════════════════════════════════════════════════════
def make_intro(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "مقدمة البحث", "نظرة عامة وأسلوب المعالجة")

    content_y = 3.2
    col_w = (W - 2.8) / 2
    gap = 0.4

    items = []
    if req.intro_overview:
        items.append(("نظرة عامة", req.intro_overview))
    if req.intro_approach:
        items.append(("المنهج المتبع", req.intro_approach))

    for i, (lbl, val) in enumerate(items[:2]):
        x = 1.2 + i * (col_w + gap)
        panel_h = H - content_y - 0.6
        # Full panel background
        panel = rect(slide, x, content_y, col_w, panel_h, T.card_rgb)
        if panel:
            shadow(panel, blur=10, dist=3, alpha=0.3)

        # Top accent (full width of panel)
        top = rect(slide, x, content_y, col_w, 0.6, T.accent_rgb)
        if top:
            gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)

        txt(slide, lbl, x + 0.2, content_y + 0.05, col_w - 0.4, 0.55,
            font=_FONT, size=13, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.RIGHT, rtl=True)

        txt(slide, val, x + 0.25, content_y + 0.75, col_w - 0.5, panel_h - 0.95,
            font=_FONT, size=11.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# PLAN — Numbered list with progress bar style
# ══════════════════════════════════════════════════════════════════════
def make_plan(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "خطة البحث", f"يتكون البحث من {len(req.chapters)} فصول")

    content_y = 3.2
    chapters = req.chapters[:9]
    avail_h = H - content_y - 0.5
    row_h = min(avail_h / max(len(chapters), 1) - 0.1, 1.6)

    for i, ch in enumerate(chapters):
        y = content_y + i * (row_h + 0.1)

        # Alternating row background
        row_bg = rect(slide, 1.0, y, W - 2.0, row_h,
                      T.bg2_rgb if i % 2 == 0 else T.card_rgb)
        if row_bg and i % 2 == 0:
            shadow(row_bg, blur=4, dist=1, alpha=0.15)

        # Left number box
        num_box = rect(slide, 1.0, y, 1.8, row_h, T.accent_rgb)
        if num_box:
            gradient_fill(num_box, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, str(i + 1), 1.0, y, 1.8, row_h,
            font="Calibri", size=18, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

        # Chapter title
        txt(slide, ch.title, 3.2, y, W - 6.0, row_h,
            font=_FONT, size=13, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

        # Pages right-aligned
        if ch.pages:
            txt(slide, ch.pages, W - 3.2, y, 2.0, row_h,
                font="Calibri", size=10, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=False)

        # Right vline accent
        vline(slide, W - 1.2, y, row_h, T.accent_rgb, thickness=0.18)

    return slide


# ══════════════════════════════════════════════════════════════════════
# PROBLEM — Stacked cards with left accent stripe
# ══════════════════════════════════════════════════════════════════════
def make_problem(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "إشكالية البحث", "التساؤلات الرئيسية والفرعية")

    cur_y = 3.2

    if req.main_problem:
        card_h = 2.2
        card = rrect(slide, 1.0, cur_y, W - 2.0, card_h, T.card_rgb, radius_pct=4)
        if card:
            shadow(card, blur=10, dist=3, alpha=0.3)
        vline(slide, 1.0, cur_y, card_h, T.accent_rgb, thickness=0.35)
        txt(slide, "الإشكالية الرئيسية", 1.6, cur_y + 0.12, W - 4.0, 0.55,
            font=_FONT, size=12, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, req.main_problem, 1.6, cur_y + 0.7, W - 3.2, card_h - 0.85,
            font=_FONT, size=12, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        cur_y += card_h + 0.25

    if req.main_question:
        q_h = 1.4
        q_card = rrect(slide, 1.0, cur_y, W - 2.0, q_h, T.bg2_rgb, radius_pct=4)
        vline(slide, 1.0, cur_y, q_h, T.accent_rgb, thickness=0.35)
        txt(slide, req.main_question, 1.6, cur_y + 0.1, W - 3.2, q_h - 0.2,
            font=_FONT, size=12.5, bold=True, italic=True,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        cur_y += q_h + 0.2

    if req.sub_questions:
        avail = H - cur_y - 0.4
        sub_h = min(avail / max(len(req.sub_questions), 1), 0.85)
        for i, q in enumerate(req.sub_questions[:6]):
            y = cur_y + i * sub_h
            txt(slide, f"◆  {q}", 1.5, y, W - 3.0, sub_h,
                font=_FONT, size=11, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# OBJECTIVES — Two-column checklist
# ══════════════════════════════════════════════════════════════════════
def make_objectives(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "أهداف البحث وفرضياته")

    content_y = 3.2
    col_w = (W - 2.8) / 2
    gap = 0.4

    cols = []
    if req.objectives:
        cols.append(("الأهداف", req.objectives))
    if req.hypotheses:
        cols.append(("الفرضيات", req.hypotheses))

    for i, (lbl, items) in enumerate(cols[:2]):
        x = 1.2 + i * (col_w + gap)
        panel_h = H - content_y - 0.5

        # Column header
        hdr = rect(slide, x, content_y, col_w, 0.7, T.accent_rgb)
        if hdr:
            gradient_fill(hdr, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, x, content_y, col_w, 0.7,
            font=_FONT, size=14, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=True)

        # Items — checklist style
        avail_h = panel_h - 0.85
        item_h = min(avail_h / max(len(items), 1), 1.0)
        for j, item in enumerate(items[:8]):
            iy = content_y + 0.85 + j * item_h
            # Check mark circle
            chk = oval(slide, x + col_w - 0.85, iy + item_h * 0.2, 0.5, 0.5, T.accent_rgb)
            txt(slide, "✓", x + col_w - 0.85, iy + item_h * 0.2, 0.5, 0.5,
                font="Calibri", size=9, bold=True,
                color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)
            # Separator line
            if j > 0:
                hline(slide, x + 0.2, iy, col_w - 0.4, T.muted_rgb, thickness=0.03)
            txt(slide, item, x + 0.2, iy + 0.05, col_w - 1.2, item_h - 0.1,
                font=_FONT, size=10.5, bold=False,
                color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# IMPORTANCE — Grid cards with bold icons
# ══════════════════════════════════════════════════════════════════════
def make_importance(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "أهمية البحث")

    content_y = 3.2
    items = req.importance[:6] if req.importance else []
    if req.reasons and len(items) < 6:
        items.append(req.reasons)

    if not items:
        return slide

    cols = 3 if len(items) >= 4 else 2 if len(items) >= 2 else 1
    rows = (len(items) + cols - 1) // cols
    gap = 0.25
    col_w = (W - 2.0 - (cols - 1) * gap) / cols
    avail_h = H - content_y - 0.4
    card_h = min(avail_h / rows - gap, 3.2)

    icons = ["01", "02", "03", "04", "05", "06"]

    for i, item in enumerate(items):
        col_idx = i % cols
        row_idx = i // cols
        x = 1.0 + col_idx * (col_w + gap)
        y = content_y + row_idx * (card_h + gap)

        card = rrect(slide, x, y, col_w, card_h, T.card_rgb, radius_pct=6)
        if card:
            shadow(card, blur=8, dist=2, alpha=0.25)

        # Top colored band
        top = rect(slide, x, y, col_w, 0.5, T.accent_rgb)
        if top:
            gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)

        # Number
        txt(slide, icons[i], x + 0.15, y + 0.05, 0.5, 0.42,
            font="Calibri", size=12, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.LEFT, rtl=False)

        txt(slide, item, x + 0.2, y + 0.6, col_w - 0.4, card_h - 0.75,
            font=_FONT, size=11, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# METHODOLOGY — Process steps (left to right numbered)
# ══════════════════════════════════════════════════════════════════════
def make_methodology(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "منهجية البحث", "الإجراءات والأدوات")

    content_y = 3.2
    fields = []
    if req.methodology:
        fields.append(("المنهج", req.methodology))
    if req.sample_type:
        fields.append(("العينة", req.sample_type))
    if req.sample_size:
        fields.append(("حجم العينة", req.sample_size))
    if req.tool:
        fields.append(("الأداة", req.tool))

    n = len(fields)
    if n == 0:
        return slide

    # If 1-2 fields: full-width stacked; if 3-4: 2 columns
    cols = 2 if n >= 3 else 1
    col_w = (W - 2.0 - (cols - 1) * 0.4) / cols
    rows = (n + cols - 1) // cols
    avail_h = H - content_y - 0.4
    card_h = min(avail_h / rows - 0.25, 3.0)

    for i, (lbl, val) in enumerate(fields[:4]):
        col_idx = i % cols
        row_idx = i // cols
        x = 1.0 + col_idx * (col_w + 0.4)
        y = content_y + row_idx * (card_h + 0.25)

        card = rrect(slide, x, y, col_w, card_h, T.card_rgb, radius_pct=5)
        if card:
            shadow(card, blur=8, dist=2, alpha=0.25)

        # Number circle left side
        step_circle = oval(slide, x + 0.25, y + card_h / 2 - 0.45, 0.9, 0.9, T.accent_rgb)
        if step_circle:
            gradient_fill(step_circle, T.accent_grad1, T.accent_grad2, 45)
        txt(slide, str(i + 1), x + 0.25, y + card_h / 2 - 0.45, 0.9, 0.9,
            font="Calibri", size=16, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

        # Label and value
        txt(slide, lbl, x + 1.4, y + 0.2, col_w - 1.6, 0.6,
            font=_FONT, size=13, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, val, x + 1.4, y + 0.85, col_w - 1.6, card_h - 1.0,
            font=_FONT, size=11.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# STATS — Large bold KPI boxes
# ══════════════════════════════════════════════════════════════════════
def make_stats(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "الإحصاءات والأرقام الرئيسية")

    content_y = 3.2
    stats = req.stats[:6]
    n = len(stats)
    if n == 0:
        return slide

    cols = 3 if n >= 3 else n
    rows = (n + cols - 1) // cols
    gap = 0.3
    card_w = (W - 2.0 - (cols - 1) * gap) / cols
    avail_h = H - content_y - 0.4
    card_h = min(avail_h / rows - gap, 4.0)

    for i, stat in enumerate(stats):
        col_idx = i % cols
        row_idx = i // cols
        x = 1.0 + col_idx * (card_w + gap)
        y = content_y + row_idx * (card_h + gap)

        # Full card with gradient
        card = rrect(slide, x, y, card_w, card_h, T.card_rgb, radius_pct=6)
        if card:
            shadow(card, blur=12, dist=4, alpha=0.35)

        # Bottom accent bar
        bot = rect(slide, x, y + card_h - 0.35, card_w, 0.35, T.accent_rgb)
        if bot:
            gradient_fill(bot, T.accent_grad1, T.accent_grad2, 0)

        # Big value
        val_size = 36 if len(stat.value) <= 4 else 26 if len(stat.value) <= 8 else 20
        txt(slide, stat.value, x + 0.2, y + 0.3, card_w - 0.4, card_h * 0.5,
            font="Calibri", size=val_size, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        if stat.unit:
            txt(slide, stat.unit, x + 0.2, y + card_h * 0.5 + 0.1, card_w - 0.4, 0.5,
                font=_FONT, size=10, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

        # Label
        txt(slide, stat.label, x + 0.2, y + card_h - 0.95, card_w - 0.4, 0.6,
            font=_FONT, size=11, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# RESULTS — Bold numbered rows with right accent
# ══════════════════════════════════════════════════════════════════════
def make_results(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "نتائج البحث", "أبرز ما توصلت إليه الدراسة")

    content_y = 3.2
    results = req.main_results[:8]
    avail_h = H - content_y - 0.4
    item_h = min(avail_h / max(len(results), 1) - 0.12, 1.5)

    for i, result in enumerate(results):
        y = content_y + i * (item_h + 0.12)
        # Row
        row = rect(slide, 1.0, y, W - 2.0, item_h,
                   T.bg2_rgb if i % 2 == 0 else T.card_rgb)

        # Left number stripe
        num_stripe = rect(slide, 1.0, y, 1.6, item_h, T.accent_rgb)
        if num_stripe:
            gradient_fill(num_stripe, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, str(i + 1), 1.0, y, 1.6, item_h,
            font="Calibri", size=16, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

        txt(slide, result, 3.0, y + 0.08, W - 4.2, item_h - 0.16,
            font=_FONT, size=11.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# CONCLUSION
# ══════════════════════════════════════════════════════════════════════
def make_conclusion(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "خاتمة البحث", "الاستنتاج العام")

    card_y = 3.2
    card_h = H - card_y - 0.8
    # Large centered card with thick left border
    card = rect(slide, 1.0, card_y, W - 2.0, card_h, T.card_rgb)
    if card:
        shadow(card, blur=16, dist=5, alpha=0.35)

    # Thick left accent
    vline(slide, 1.0, card_y, card_h, T.accent_rgb, thickness=0.5)

    txt(slide, "❝", 1.8, card_y + 0.3, 2.0, 1.5,
        font="Calibri", size=50, bold=False,
        color=T.accent_rgb, align=PP_ALIGN.LEFT, rtl=False)

    txt(slide, req.general_conclusion,
        2.0, card_y + 1.0, W - 3.5, card_h - 1.4,
        font=_FONT, size=14, bold=False,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# RECOMMENDATIONS
# ══════════════════════════════════════════════════════════════════════
def make_recommendations(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "توصيات البحث")

    content_y = 3.2
    recs = req.recommendations[:8]
    avail_h = H - content_y - 0.4
    item_h = min(avail_h / max(len(recs), 1) - 0.1, 1.4)

    for i, rec in enumerate(recs):
        y = content_y + i * (item_h + 0.1)
        row = rect(slide, 1.0, y, W - 2.0, item_h,
                   T.bg2_rgb if i % 2 == 0 else T.card_rgb)

        # Diamond bullet
        txt(slide, "◆", W - 2.5, y, 1.2, item_h,
            font="Calibri", size=12, bold=False,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        txt(slide, rec, 1.4, y + 0.08, W - 4.0, item_h - 0.16,
            font=_FONT, size=11.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# FUTURE WORK
# ══════════════════════════════════════════════════════════════════════
def make_future(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "آفاق البحث المستقبلية")

    content_y = 3.2
    items = req.future_work[:6]
    cols = 2 if len(items) > 3 else 1
    col_w = (W - 2.0 - (cols - 1) * 0.4) / cols
    rows = (len(items) + cols - 1) // cols
    avail_h = H - content_y - 0.4
    card_h = min(avail_h / rows - 0.2, 2.5)

    for i, item in enumerate(items):
        col_idx = i % cols
        row_idx = i // cols
        x = 1.0 + col_idx * (col_w + 0.4)
        y = content_y + row_idx * (card_h + 0.2)

        card = rrect(slide, x, y, col_w, card_h, T.card_rgb, radius_pct=5)
        if card:
            shadow(card, blur=8, dist=2, alpha=0.25)

        # Right accent stripe
        vline(slide, x + col_w - 0.35, y, card_h, T.accent_rgb, thickness=0.35)

        txt(slide, item, x + 0.3, y + 0.2, col_w - 0.9, card_h - 0.4,
            font=_FONT, size=11.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# REFERENCES
# ══════════════════════════════════════════════════════════════════════
def make_references(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "المراجع والمصادر")

    content_y = 3.2
    refs = req.references[:12]
    avail_h = H - content_y - 0.4
    item_h = max(min(avail_h / max(len(refs), 1) - 0.08, 1.0), 0.5)

    for i, ref in enumerate(refs):
        y = content_y + i * (item_h + 0.08)
        if y + item_h > H - 0.3:
            break

        if i % 2 == 0:
            rect(slide, 1.0, y, W - 2.0, item_h, T.bg2_rgb)

        txt(slide, f"[{i + 1}]", 1.2, y, 1.2, item_h,
            font="Calibri", size=9, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.LEFT, rtl=False)

        txt(slide, ref, 2.6, y + 0.04, W - 4.0, item_h - 0.08,
            font=_FONT, size=10, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# FINAL SLIDE — Split layout matching cover
# ══════════════════════════════════════════════════════════════════════
def make_final(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)

    # Right accent panel (mirror of cover)
    panel_w = 7.5
    right_panel = rect(slide, W - panel_w, 0, panel_w, H, T.accent_rgb)
    if right_panel:
        gradient_fill(right_panel, T.accent_grad2, T.accent_grad1, angle=180)

    oval(slide, W - panel_w - 2, H * 0.3, 8, 8, T.bg_rgb, alpha=8)
    oval(slide, W - 4, H - 6, 7, 7, T.bg_rgb, alpha=10)

    content_w = W - panel_w - 2.0
    center_y = H / 2

    hline(slide, 1.0, center_y - 3.5, content_w, T.accent_rgb, thickness=0.08)

    txt(slide, "شكراً وتقديراً",
        1.0, center_y - 3.0, content_w, 2.2,
        font=_FONT, size=38, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    hline(slide, 1.0, center_y - 0.6, content_w * 0.5, T.accent_rgb, thickness=0.06)

    txt(slide, req.student_name,
        1.0, center_y - 0.3, content_w, 1.1,
        font=_FONT, size=18, bold=True,
        color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    title_display = req.title_ar[:70] + ("..." if len(req.title_ar) > 70 else "")
    txt(slide, title_display,
        1.0, center_y + 1.0, content_w, 1.8,
        font=_FONT, size=12, bold=False, italic=True,
        color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    hline(slide, 1.0, center_y + 3.0, content_w, T.accent_rgb, thickness=0.04)

    footer = []
    if req.institution:
        footer.append(req.institution)
    if req.year:
        footer.append(req.year)
    if footer:
        txt(slide, "  ·  ".join(footer),
            1.0, center_y + 3.2, content_w, 0.7,
            font=_FONT, size=10, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide
