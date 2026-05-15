"""
Slide Builder — Premium Engine — مذكرتي Pro v18
Layout philosophy: Cinematic, editorial, magazine-style.
Full-bleed backgrounds, diagonal accents, large typography,
asymmetric compositions. Inspired by McKinsey/BCG decks.
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
# PREMIUM COVER — Full-bleed diagonal split
# ══════════════════════════════════════════════════════════════════════
def make_cover(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)

    # Full gradient background
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=150)

    # Large decorative accent circle (top-right bleed)
    oval(slide, W * 0.55, -H * 0.3, H * 1.3, H * 1.3, T.accent_rgb, alpha=7)

    # Bottom-left dark triangle area simulation
    dark_block = rect(slide, 0, H * 0.58, W * 0.48, H * 0.42, T.bg_rgb)
    if dark_block:
        set_solid_alpha(dark_block, 40)

    # Top micro-strip
    top_micro = rect(slide, 0, 0, W, 0.18, T.accent_rgb)
    if top_micro:
        gradient_fill(top_micro, T.accent_grad1, T.accent_grad2, 0)

    # Institution — top right
    if req.institution:
        txt(slide, req.institution, W * 0.35, 0.3, W * 0.6, 0.7,
            font=_FONT, size=10, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    # Giant title — dominates the slide
    title_y = H * 0.12
    title_size = 32 if len(req.title_ar) < 40 else 24 if len(req.title_ar) < 70 else 18
    txt(slide, req.title_ar,
        1.5, title_y, W - 3.0, H * 0.42,
        font=_FONT, size=title_size, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    if req.title_en:
        txt(slide, req.title_en,
            1.5, title_y + H * 0.42 - 0.2, W - 3.0, 0.8,
            font="Calibri", size=12, bold=False, italic=True,
            color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=False)

    # Diagonal accent line (simulated with a thin rotated-ish rect)
    hline(slide, W * 0.08, H * 0.595, W * 0.84, T.accent_rgb, thickness=0.12)

    # Bottom info block — pill badges
    info_y = H * 0.64
    badge_h = 0.62
    badge_gap = 0.3

    fields = []
    fields.append(req.student_name)
    if req.supervisor:
        fields.append(req.supervisor)
    if req.specialization:
        fields.append(req.specialization)
    if req.year:
        fields.append(req.year)

    badge_w = (W - 2.0) / max(len(fields), 1) - badge_gap

    for i, val in enumerate(fields[:4]):
        bx = W - 1.2 - (i + 1) * (badge_w + badge_gap)
        badge = rrect(slide, bx, info_y, badge_w, badge_h, T.card_rgb, radius_pct=50)
        if badge:
            shadow(badge, blur=8, dist=2, alpha=0.3)
        txt(slide, val, bx, info_y, badge_w, badge_h,
            font=_FONT, size=10.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # Label row below
    labels = ["الطالب", "المشرف", "التخصص", "السنة"]
    for i, lbl in enumerate(labels[:len(fields)]):
        bx = W - 1.2 - (i + 1) * (badge_w + badge_gap)
        txt(slide, lbl, bx, info_y + badge_h + 0.08, badge_w, 0.4,
            font=_FONT, size=9, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # Bottom bar
    bot = rect(slide, 0, H - 0.3, W, 0.3, T.accent_rgb)
    if bot:
        gradient_fill(bot, T.accent_grad1, T.accent_grad2, 0)

    return slide


# ══════════════════════════════════════════════════════════════════════
# PREMIUM SECTION HEADER — Minimal top strip + large title
# ══════════════════════════════════════════════════════════════════════
def _section_header(slide, T: Theme, title: str, subtitle: str = ""):
    # Thin top accent
    top = rect(slide, 0, 0, W, 0.22, T.accent_rgb)
    if top:
        gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)

    # Background gradient — subtle
    gradient_rect(slide, 0, 0.22, W, 2.9, T.grad1, T.grad2, angle=90)

    # Large title — magazine style, right-aligned
    txt(slide, title, 1.0, 0.35, W - 2.0, 1.8,
        font=_FONT, size=26, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    if subtitle:
        txt(slide, subtitle, 1.0, 1.95, W - 2.0, 0.85,
            font=_FONT, size=12, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    # Accent dot before title
    oval(slide, W - 1.3, 0.75, 0.5, 0.5, T.accent_rgb)

    # Bottom rule
    hline(slide, 0, 2.95, W, T.accent_rgb, thickness=0.05)


# ══════════════════════════════════════════════════════════════════════
# INTRO — Full-bleed panels with large typography
# ══════════════════════════════════════════════════════════════════════
def make_intro(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "مقدمة البحث", "نظرة عامة وأسلوب المعالجة")

    content_y = 3.1
    items = []
    if req.intro_overview:
        items.append(("نظرة عامة", req.intro_overview))
    if req.intro_approach:
        items.append(("المنهج المتبع", req.intro_approach))

    col_w = (W - 2.6) / 2 if len(items) == 2 else W - 2.4
    gap = 0.4

    for i, (lbl, val) in enumerate(items[:2]):
        x = 1.2 + i * (col_w + gap)
        panel_h = H - content_y - 0.5

        # Full gradient panel
        panel = rect(slide, x, content_y, col_w, panel_h, T.card_rgb)
        if panel:
            shadow(panel, blur=14, dist=4, alpha=0.4)

        # Accent diagonal strip (top-left corner)
        corner = rect(slide, x, content_y, col_w * 0.35, 0.55, T.accent_rgb)
        if corner:
            gradient_fill(corner, T.accent_grad1, T.accent_grad2, 0)

        txt(slide, lbl, x + col_w * 0.38, content_y + 0.05, col_w * 0.58, 0.55,
            font=_FONT, size=13, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)

        txt(slide, val, x + 0.3, content_y + 0.75, col_w - 0.6, panel_h - 1.0,
            font=_FONT, size=12, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# PLAN — Timeline-style horizontal flow
# ══════════════════════════════════════════════════════════════════════
def make_plan(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "خطة البحث", f"يتكون البحث من {len(req.chapters)} فصول")

    chapters = req.chapters[:8]
    n = len(chapters)
    if n == 0:
        return slide

    content_y = 3.2
    avail_h = H - content_y - 0.5

    # If ≤ 4: horizontal timeline; else: two columns
    if n <= 4:
        # Horizontal timeline
        card_w = (W - 2.0) / n - 0.3
        card_h = avail_h - 0.8
        timeline_y = content_y + 0.6

        # Connecting line
        hline(slide, 1.5, content_y + 0.35, W - 3.0, T.muted_rgb, thickness=0.04)

        for i, ch in enumerate(chapters):
            x = 1.0 + i * (card_w + 0.3)

            # Node dot
            node = oval(slide, x + card_w / 2 - 0.35, content_y + 0.08, 0.7, 0.7, T.accent_rgb)
            if node:
                gradient_fill(node, T.accent_grad1, T.accent_grad2, 45)

            txt(slide, str(i + 1), x + card_w / 2 - 0.35, content_y + 0.08, 0.7, 0.7,
                font="Calibri", size=12, bold=True,
                color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

            card = rrect(slide, x, timeline_y, card_w, card_h, T.card_rgb, radius_pct=8)
            if card:
                shadow(card, blur=10, dist=3, alpha=0.3)

            # Top accent
            top = rect(slide, x, timeline_y, card_w, 0.3, T.accent_rgb)
            if top:
                gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)

            txt(slide, ch.title, x + 0.2, timeline_y + 0.4, card_w - 0.4, card_h - 0.6,
                font=_FONT, size=11, bold=False,
                color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

            if ch.pages:
                txt(slide, ch.pages, x + 0.2, timeline_y + card_h - 0.55, card_w - 0.4, 0.45,
                    font="Calibri", size=9, bold=False,
                    color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=False)
    else:
        # Two columns
        col_w = (W - 2.6) / 2
        row_h = min(avail_h / ((n + 1) // 2) - 0.15, 1.5)

        for i, ch in enumerate(chapters):
            col_idx = i % 2
            row_idx = i // 2
            x = 1.2 + col_idx * (col_w + 0.4)
            y = content_y + row_idx * (row_h + 0.15)

            row_bg = rrect(slide, x, y, col_w, row_h, T.card_rgb, radius_pct=6)
            if row_bg:
                shadow(row_bg, blur=6, dist=2, alpha=0.2)

            # Accent left bar
            vline(slide, x, y, row_h, T.accent_rgb, thickness=0.3)

            txt(slide, str(i + 1), x + 0.4, y, 0.8, row_h,
                font="Calibri", size=14, bold=True,
                color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

            txt(slide, ch.title, x + 1.4, y, col_w - 1.6, row_h,
                font=_FONT, size=12, bold=False,
                color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# PROBLEM — Dramatic single focus + sub questions
# ══════════════════════════════════════════════════════════════════════
def make_problem(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "إشكالية البحث", "التساؤلات الرئيسية والفرعية")

    cur_y = 3.2

    if req.main_problem:
        # Full-width dramatic card
        card_h = 2.4
        grad_card = rect(slide, 0.8, cur_y, W - 1.6, card_h, T.card_rgb)
        if grad_card:
            shadow(grad_card, blur=16, dist=5, alpha=0.45)

        # Left bleed accent
        accent_bleed = rect(slide, 0.8, cur_y, 3.5, card_h, T.accent_rgb)
        if accent_bleed:
            gradient_fill(accent_bleed, T.accent_grad1, T.accent_grad2, 90)
            set_solid_alpha(accent_bleed, 30)

        # Bold question mark
        txt(slide, "؟", 1.0, cur_y + 0.2, 2.5, card_h - 0.4,
            font="Calibri", size=48, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        txt(slide, "الإشكالية الرئيسية", W - 4.5, cur_y + 0.15, 3.5, 0.55,
            font=_FONT, size=10, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)

        txt(slide, req.main_problem, 4.0, cur_y + 0.65, W - 5.2, card_h - 0.85,
            font=_FONT, size=13, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

        cur_y += card_h + 0.3

    if req.main_question:
        q_h = 1.2
        q_bg = rrect(slide, 0.8, cur_y, W - 1.6, q_h, T.bg2_rgb, radius_pct=4)
        hline(slide, 0.8, cur_y + q_h - 0.08, W - 1.6, T.accent_rgb, thickness=0.08)
        txt(slide, req.main_question, 1.2, cur_y + 0.1, W - 2.4, q_h - 0.2,
            font=_FONT, size=13, bold=True, italic=True,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        cur_y += q_h + 0.2

    if req.sub_questions:
        avail = H - cur_y - 0.4
        sub_h = min(avail / max(len(req.sub_questions), 1), 0.8)
        cols = 2 if len(req.sub_questions) > 3 else 1
        col_w = (W - 1.8) / cols - 0.2

        for i, q in enumerate(req.sub_questions[:6]):
            col_idx = i % cols
            row_idx = i // cols
            x = 0.9 + col_idx * (col_w + 0.2)
            y = cur_y + row_idx * sub_h

            # Accent dot
            oval(slide, x + col_w - 0.6, y + sub_h * 0.28, 0.32, 0.32, T.accent_rgb)
            txt(slide, q, x + 0.1, y, col_w - 0.8, sub_h,
                font=_FONT, size=11, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# OBJECTIVES — Asymmetric split
# ══════════════════════════════════════════════════════════════════════
def make_objectives(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "أهداف البحث وفرضياته")

    content_y = 3.1
    panel_h = H - content_y - 0.4

    cols = []
    if req.objectives:
        cols.append(("الأهداف", req.objectives))
    if req.hypotheses:
        cols.append(("الفرضيات", req.hypotheses))

    # Asymmetric widths if two columns: 55% / 45%
    if len(cols) == 2:
        widths = [W * 0.52 - 1.6, W * 0.44 - 0.4]
    elif len(cols) == 1:
        widths = [W - 2.4]
    else:
        return slide

    x_pos = [1.2]
    if len(cols) == 2:
        x_pos.append(1.2 + widths[0] + 0.4)

    for i, (lbl, items) in enumerate(cols[:2]):
        x = x_pos[i]
        cw = widths[i]

        panel = rrect(slide, x, content_y, cw, panel_h, T.card_rgb, radius_pct=6)
        if panel:
            shadow(panel, blur=12, dist=4, alpha=0.35)

        # Header gradient
        hdr = rect(slide, x, content_y, cw, 0.72, T.accent_rgb)
        if hdr:
            gradient_fill(hdr, T.accent_grad1, T.accent_grad2, 0)

        txt(slide, lbl, x + 0.2, content_y + 0.05, cw - 0.4, 0.65,
            font=_FONT, size=14, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=True)

        avail_h = panel_h - 0.88
        item_h = min(avail_h / max(len(items), 1), 1.05)
        for j, item in enumerate(items[:8]):
            iy = content_y + 0.82 + j * item_h

            # Numbered oval
            num_c = oval(slide, x + cw - 0.85, iy + 0.07, 0.55, 0.55, T.bg_rgb)
            txt(slide, str(j + 1), x + cw - 0.85, iy + 0.07, 0.55, 0.55,
                font="Calibri", size=10, bold=True,
                color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

            if j > 0:
                hline(slide, x + 0.2, iy, cw - 0.4, T.muted_rgb, thickness=0.025)

            txt(slide, item, x + 0.2, iy + 0.04, cw - 1.2, item_h - 0.08,
                font=_FONT, size=10.5, bold=False,
                color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# IMPORTANCE — Hexagon-inspired grid (rectangles with heavy top accent)
# ══════════════════════════════════════════════════════════════════════
def make_importance(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "أهمية البحث")

    content_y = 3.1
    items = req.importance[:6] if req.importance else []
    if req.reasons and len(items) < 6:
        items.append(req.reasons)

    if not items:
        return slide

    n = len(items)
    cols = 3 if n >= 4 else 2 if n >= 2 else 1
    rows = (n + cols - 1) // cols
    gap = 0.3
    col_w = (W - 2.0 - (cols - 1) * gap) / cols
    avail_h = H - content_y - 0.4
    card_h = min(avail_h / rows - gap, 3.5)

    accent_heights = [0.55, 0.7, 0.5, 0.65, 0.6, 0.75]  # varied tops

    for i, item in enumerate(items):
        col_idx = i % cols
        row_idx = i // cols
        x = 1.0 + col_idx * (col_w + gap)
        y = content_y + row_idx * (card_h + gap)

        card = rrect(slide, x, y, col_w, card_h, T.card_rgb, radius_pct=4)
        if card:
            shadow(card, blur=10, dist=3, alpha=0.3)

        ah = accent_heights[i % len(accent_heights)]
        top_acc = rect(slide, x, y, col_w, ah, T.accent_rgb)
        if top_acc:
            gradient_fill(top_acc, T.accent_grad1, T.accent_grad2, 0)
            set_solid_alpha(top_acc, 90)

        # Large number overlay on accent
        txt(slide, str(i + 1).zfill(2), x + 0.15, y + 0.02, 1.2, ah - 0.04,
            font="Calibri", size=18, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.LEFT, rtl=False)

        txt(slide, item, x + 0.2, y + ah + 0.12, col_w - 0.4, card_h - ah - 0.22,
            font=_FONT, size=11, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# METHODOLOGY — Icon-centric process flow
# ══════════════════════════════════════════════════════════════════════
def make_methodology(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "منهجية البحث", "الإجراءات والأدوات")

    content_y = 3.1
    fields = []
    if req.methodology:
        fields.append(("المنهج", req.methodology, "M"))
    if req.sample_type:
        fields.append(("العينة", req.sample_type, "S"))
    if req.sample_size:
        fields.append(("حجم العينة", req.sample_size, "N"))
    if req.tool:
        fields.append(("الأداة", req.tool, "T"))

    n = len(fields)
    if n == 0:
        return slide

    # Horizontal flow for ≤ 4 items
    card_w = (W - 2.0) / n - 0.3
    card_h = H - content_y - 0.5

    # Connector lines between cards
    for i in range(n - 1):
        x_line = 1.0 + (i + 1) * (card_w + 0.3) - 0.15
        hline(slide, x_line, content_y + card_h / 2, 0.3, T.accent_rgb, thickness=0.06)

    for i, (lbl, val, icon_char) in enumerate(fields[:4]):
        x = 1.0 + i * (card_w + 0.3)

        card = rrect(slide, x, content_y, card_w, card_h, T.card_rgb, radius_pct=8)
        if card:
            shadow(card, blur=12, dist=4, alpha=0.35)

        # Large icon circle
        icon_sz = 1.4
        icon_x = x + (card_w - icon_sz) / 2
        icon_circle = oval(slide, icon_x, content_y + 0.5, icon_sz, icon_sz, T.accent_rgb)
        if icon_circle:
            gradient_fill(icon_circle, T.accent_grad1, T.accent_grad2, 45)
            shadow(icon_circle, blur=8, dist=3, alpha=0.4)

        txt(slide, icon_char, icon_x, content_y + 0.5, icon_sz, icon_sz,
            font="Calibri", size=18, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

        # Label
        txt(slide, lbl, x + 0.15, content_y + 2.1, card_w - 0.3, 0.65,
            font=_FONT, size=13, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=True)

        hline(slide, x + 0.3, content_y + 2.8, card_w - 0.6, T.muted_rgb, thickness=0.04)

        txt(slide, val, x + 0.2, content_y + 2.95, card_w - 0.4, card_h - 3.1,
            font=_FONT, size=11, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# STATS — Bold KPI with accent background
# ══════════════════════════════════════════════════════════════════════
def make_stats(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "الإحصاءات والأرقام الرئيسية")

    content_y = 3.1
    stats = req.stats[:6]
    n = len(stats)
    if n == 0:
        return slide

    cols = 3 if n >= 3 else n
    rows = (n + cols - 1) // cols
    gap = 0.35
    card_w = (W - 2.0 - (cols - 1) * gap) / cols
    avail_h = H - content_y - 0.4
    card_h = min(avail_h / rows - gap, 4.2)

    for i, stat in enumerate(stats):
        col_idx = i % cols
        row_idx = i // cols
        x = 1.0 + col_idx * (card_w + gap)
        y = content_y + row_idx * (card_h + gap)

        # Card with gradient fill
        card = rrect(slide, x, y, card_w, card_h, T.card_rgb, radius_pct=10)
        if card:
            gradient_fill(card, T.bg2, T.card, angle=135)
            shadow(card, blur=16, dist=5, alpha=0.45)

        # Large accent circle background
        c_sz = min(card_w, card_h * 0.7)
        circ = oval(slide, x + (card_w - c_sz) / 2, y + 0.2, c_sz, c_sz, T.accent_rgb)
        if circ:
            set_solid_alpha(circ, 8)

        # Value
        val_size = 38 if len(stat.value) <= 4 else 28 if len(stat.value) <= 8 else 20
        txt(slide, stat.value, x + 0.2, y + 0.35, card_w - 0.4, card_h * 0.52,
            font="Calibri", size=val_size, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        if stat.unit:
            txt(slide, stat.unit, x + 0.2, y + card_h * 0.52 + 0.15, card_w - 0.4, 0.5,
                font=_FONT, size=10, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

        hline(slide, x + card_w * 0.2, y + card_h - 1.0, card_w * 0.6,
              T.accent_rgb, thickness=0.06)

        txt(slide, stat.label, x + 0.2, y + card_h - 0.9, card_w - 0.4, 0.8,
            font=_FONT, size=11, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# RESULTS — Magazine timeline rows
# ══════════════════════════════════════════════════════════════════════
def make_results(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "نتائج البحث", "أبرز ما توصلت إليه الدراسة")

    content_y = 3.1
    results = req.main_results[:8]
    avail_h = H - content_y - 0.4
    item_h = min(avail_h / max(len(results), 1) - 0.12, 1.55)

    for i, result in enumerate(results):
        y = content_y + i * (item_h + 0.12)

        # Alternating card
        row = rrect(slide, 0.8, y, W - 1.6, item_h,
                    T.bg2_rgb if i % 2 == 0 else T.card_rgb, radius_pct=4)
        if row and i % 2 == 0:
            shadow(row, blur=5, dist=1, alpha=0.15)

        # Right accent pill number
        pill = rrect(slide, W - 3.0, y + (item_h - 0.55) / 2, 1.8, 0.55,
                     T.accent_rgb, radius_pct=50)
        if pill:
            gradient_fill(pill, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, str(i + 1), W - 3.0, y + (item_h - 0.55) / 2, 1.8, 0.55,
            font="Calibri", size=12, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

        txt(slide, result, 1.2, y + 0.1, W - 5.0, item_h - 0.2,
            font=_FONT, size=11.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# CONCLUSION — Full-bleed quote card
# ══════════════════════════════════════════════════════════════════════
def make_conclusion(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _section_header(slide, T, "خاتمة البحث", "الاستنتاج العام")

    # Near full-slide card
    card_y = 3.1
    card_h = H - card_y - 0.6
    card = rect(slide, 0.8, card_y, W - 1.6, card_h, T.card_rgb)
    if card:
        shadow(card, blur=20, dist=6, alpha=0.45)

    # Gradient left bleed
    left_bleed = rect(slide, 0.8, card_y, 2.5, card_h, T.accent_rgb)
    if left_bleed:
        gradient_fill(left_bleed, T.accent_grad1, T.accent_grad2, 90)
        set_solid_alpha(left_bleed, 20)

    # Giant quote mark
    txt(slide, "❝", 1.0, card_y + 0.2, 2.2, 2.0,
        font="Calibri", size=60, bold=False,
        color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

    # Top accent stripe
    top_s = rect(slide, 0.8, card_y, W - 1.6, 0.3, T.accent_rgb)
    if top_s:
        gradient_fill(top_s, T.accent_grad1, T.accent_grad2, 0)

    txt(slide, req.general_conclusion,
        3.5, card_y + 0.8, W - 4.8, card_h - 1.2,
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

    content_y = 3.1
    recs = req.recommendations[:8]
    avail_h = H - content_y - 0.4
    item_h = min(avail_h / max(len(recs), 1) - 0.1, 1.45)

    for i, rec in enumerate(recs):
        y = content_y + i * (item_h + 0.1)

        row = rrect(slide, 0.8, y, W - 1.6, item_h,
                    T.bg2_rgb if i % 2 == 0 else T.card_rgb, radius_pct=4)

        # Accent left tag
        tag = rect(slide, 0.8, y, 0.4, item_h, T.accent_rgb)
        if tag:
            gradient_fill(tag, T.accent_grad1, T.accent_grad2, 90)

        txt(slide, rec, 1.5, y + 0.08, W - 3.0, item_h - 0.16,
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

    content_y = 3.1
    items = req.future_work[:6]
    n = len(items)
    if n == 0:
        return slide

    cols = 3 if n >= 4 else 2 if n >= 2 else 1
    rows = (n + cols - 1) // cols
    gap = 0.3
    col_w = (W - 2.0 - (cols - 1) * gap) / cols
    avail_h = H - content_y - 0.4
    card_h = min(avail_h / rows - gap, 2.8)

    for i, item in enumerate(items):
        col_idx = i % cols
        row_idx = i // cols
        x = 1.0 + col_idx * (col_w + gap)
        y = content_y + row_idx * (card_h + gap)

        card = rrect(slide, x, y, col_w, card_h, T.card_rgb, radius_pct=8)
        if card:
            shadow(card, blur=10, dist=3, alpha=0.3)

        # Bottom accent (different from other slides — variety)
        bot = rect(slide, x, y + card_h - 0.4, col_w, 0.4, T.accent_rgb)
        if bot:
            gradient_fill(bot, T.accent_grad1, T.accent_grad2, 0)

        # Number top-left
        txt(slide, f"{i + 1:02d}", x + 0.2, y + 0.15, 1.0, 0.7,
            font="Calibri", size=18, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.LEFT, rtl=False)

        txt(slide, item, x + 0.2, y + 0.85, col_w - 0.4, card_h - 1.4,
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

    content_y = 3.1
    refs = req.references[:12]
    avail_h = H - content_y - 0.4
    item_h = max(min(avail_h / max(len(refs), 1) - 0.08, 1.0), 0.5)

    for i, ref in enumerate(refs):
        y = content_y + i * (item_h + 0.08)
        if y + item_h > H - 0.3:
            break

        if i % 2 == 0:
            rrect(slide, 0.8, y, W - 1.6, item_h, T.bg2_rgb, radius_pct=3)

        # Accent number pill
        pill = rrect(slide, W - 2.8, y + (item_h - 0.38) / 2, 1.6, 0.38,
                     T.accent_rgb, radius_pct=50)
        if pill:
            set_solid_alpha(pill, 70)
        txt(slide, f"[{i + 1}]", W - 2.8, y + (item_h - 0.38) / 2, 1.6, 0.38,
            font="Calibri", size=9, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

        txt(slide, ref, 1.2, y + 0.04, W - 4.4, item_h - 0.08,
            font=_FONT, size=10, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# FINAL SLIDE — Cinematic full-bleed
# ══════════════════════════════════════════════════════════════════════
def make_final(prs: Presentation, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)

    # Full gradient
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=150)

    # Large accent circles — cinematic feel
    oval(slide, -4, -4, 18, 18, T.accent_rgb, alpha=6)
    oval(slide, W - 10, H - 10, 16, 16, T.accent_rgb, alpha=5)
    oval(slide, W * 0.3, -3, 12, 12, T.bg2_rgb, alpha=25)

    # Center card
    card_w, card_h = 24, 11
    card_x = (W - card_w) / 2
    card_y = (H - card_h) / 2

    card = rrect(slide, card_x, card_y, card_w, card_h, T.card_rgb, radius_pct=10)
    if card:
        shadow(card, blur=28, dist=10, alpha=0.5)

    # Top accent bleed
    top_s = rect(slide, card_x, card_y, card_w, 0.4, T.accent_rgb)
    if top_s:
        gradient_fill(top_s, T.accent_grad1, T.accent_grad2, 0)

    # Giant thank-you
    txt(slide, "شكراً وتقديراً",
        card_x + 1.0, card_y + 0.6, card_w - 2.0, 2.8,
        font=_FONT, size=40, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # Accent dot row
    for j in range(5):
        dot_x = card_x + card_w / 2 - 1.5 + j * 0.75
        is_center = j == 2
        c = oval(slide, dot_x, card_y + 3.6, 0.35 if is_center else 0.22,
                 0.35 if is_center else 0.22, T.accent_rgb)
        if c and not is_center:
            set_solid_alpha(c, 50)

    txt(slide, req.student_name,
        card_x + 1.0, card_y + 4.2, card_w - 2.0, 1.2,
        font=_FONT, size=20, bold=True,
        color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=True)

    title_display = req.title_ar[:75] + ("..." if len(req.title_ar) > 75 else "")
    txt(slide, title_display,
        card_x + 1.5, card_y + 5.5, card_w - 3.0, 2.2,
        font=_FONT, size=12, bold=False, italic=True,
        color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    hline(slide, card_x + card_w * 0.15, card_y + card_h - 1.3,
          card_w * 0.7, T.accent_rgb, thickness=0.04)

    footer = []
    if req.institution:
        footer.append(req.institution)
    if req.year:
        footer.append(req.year)
    if footer:
        txt(slide, "  ·  ".join(footer),
            card_x + 1.0, card_y + card_h - 1.1, card_w - 2.0, 0.7,
            font=_FONT, size=11, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    return slide
