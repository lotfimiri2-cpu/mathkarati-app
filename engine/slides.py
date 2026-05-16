"""
Slide Builder — CANVA Engine — مذكرتي Pro v18
فلسفة التصميم: Modern Card-Based Design
- بطاقات مستديرة بظلال عميقة
- تدرجات ناعمة متعددة الطبقات
- دوائر زخرفية كبيرة في الخلفية
- المحتوى يطفو فوق خلفية غنية
- مناسب: طلاب ماستر، علوم إنسانية
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
    h = h.lstrip('#')
    return RGBColor(int(h[:2],16), int(h[2:4],16), int(h[4:],16))

def _header(slide, T, title, subtitle=""):
    """شريط علوي موحد لجميع شرائح canva"""
    # خلفية الشريط بتدرج
    hdr = rect(slide, 0, 0, W, 3.4, _rgb(T.bg2))
    if hdr: gradient_fill(hdr, T.grad1, T.grad2, 180)
    # دائرة زخرفية شبه شفافة
    oval(slide, W-6, -3, 9, 9, _rgb(T.accent), alpha=6)
    # خط فاصل لامع
    bar = rect(slide, 0, 3.25, W, 0.15, _rgb(T.accent))
    if bar: gradient_fill(bar, T.accent_grad1, T.accent_grad2, 0)
    # العنوان
    size = 22 if len(title) < 20 else 18
    txt(slide, title, 1.2, 0.5, W-2.4, 1.8,
        font=_FONT, size=size, bold=True,
        color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    if subtitle:
        txt(slide, subtitle, 1.2, 2.1, W-2.4, 0.95,
            font=_FONT, size=11, bold=False,
            color=_rgb(T.muted), align=PP_ALIGN.RIGHT, rtl=True)

def _card(slide, x, y, w, h, T, radius=12, depth=True):
    c = rrect(slide, x, y, w, h, _rgb(T.card), radius_pct=radius)
    if c:
        gradient_fill(c, T.card, T.bg2, 135)
        if depth: shadow(c, blur=18, dist=6, alpha=0.45)
    return c

# ══════════════════════════════════════════════════════════════════════
# COVER — Floating card on rich gradient background
# ══════════════════════════════════════════════════════════════════════
def make_cover(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    # خلفية بتدرج عميق
    s = gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 150)
    # دوائر زخرفية خلفية
    oval(slide, -4, -4, 16, 16, _rgb(T.accent), alpha=5)
    oval(slide, W*0.4, -3, 18, 18, _rgb(T.bg2), alpha=60)
    oval(slide, W-8, H-6, 14, 14, _rgb(T.accent), alpha=7)
    oval(slide, 2, H-5, 10, 10, _rgb(T.accent2), alpha=5)

    # شريط علوي ذهبي
    top = rect(slide, 0, 0, W, 0.5, _rgb(T.accent))
    if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)

    # بطاقة المؤسسة
    if req.institution:
        inst_parts = req.institution.split(' — ')
        inst_line = inst_parts[0] if inst_parts else req.institution
        inst_card = rrect(slide, 1.5, 0.8, W-3, 0.75, _rgb(T.card), radius_pct=40)
        if inst_card:
            shadow(inst_card, blur=10, dist=3, alpha=0.3)
            set_solid_alpha(inst_card, 70)
        txt(slide, inst_line, 1.7, 0.8, W-3.4, 0.75,
            font=_FONT, size=11, bold=False,
            color=_rgb(T.muted), align=PP_ALIGN.CENTER, rtl=True)

    # البطاقة الرئيسية للعنوان — تطفو في المنتصف
    card_y = H*0.2
    card_h = H*0.38
    main_card = rrect(slide, 2.5, card_y, W-5, card_h, _rgb(T.card), radius_pct=15)
    if main_card:
        gradient_fill(main_card, T.card, T.bg2, 135)
        shadow(main_card, blur=25, dist=8, alpha=0.55)
    # شريط علوي للبطاقة
    card_top = rrect(slide, 2.5, card_y, W-5, 0.6, _rgb(T.accent), radius_pct=15)
    if card_top: gradient_fill(card_top, T.accent_grad1, T.accent_grad2, 0)

    title_size = 28 if len(req.title_ar)<35 else 22 if len(req.title_ar)<55 else 17 if len(req.title_ar)<80 else 14
    txt(slide, req.title_ar, 3.0, card_y+0.8, W-6, card_h-1.2,
        font=_FONT, size=title_size, bold=True,
        color=_rgb(T.text_light), align=PP_ALIGN.CENTER, rtl=True)

    if req.title_en:
        txt(slide, req.title_en, 3.0, card_y+card_h-0.9, W-6, 0.75,
            font="Calibri", size=10.5, bold=False, italic=True,
            color=_rgb(T.muted), align=PP_ALIGN.CENTER, rtl=False)

    # قسم المعلومات — بطاقات أفقية
    info_y = card_y + card_h + 0.5
    fields = [("الطالب", req.student_name)]
    if req.supervisor: fields.append(("المشرف", req.supervisor))
    if req.co_supervisor: fields.append(("م. مساعد", req.co_supervisor))
    if req.specialization: fields.append(("التخصص", req.specialization))

    n = len(fields)
    pill_w = min((W-3.0) / n - 0.3, 8.0)
    total_w = n*(pill_w+0.3)-0.3
    start_x = (W - total_w)/2

    for i,(lbl,val) in enumerate(fields[:4]):
        px = start_x + i*(pill_w+0.3)
        pill = rrect(slide, px, info_y, pill_w, 1.9, _rgb(T.card), radius_pct=12)
        if pill:
            shadow(pill, blur=12, dist=4, alpha=0.35)
            set_solid_alpha(pill, 85)
        # لون العنوان
        lbl_bar = rrect(slide, px, info_y, pill_w, 0.55, _rgb(T.accent), radius_pct=12)
        if lbl_bar: gradient_fill(lbl_bar, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, px, info_y+0.02, pill_w, 0.5,
            font=_FONT, size=10, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.CENTER, rtl=True)
        txt(slide, val, px+0.2, info_y+0.62, pill_w-0.4, 1.2,
            font=_FONT, size=11, bold=False,
            color=_rgb(T.text_light), align=PP_ALIGN.CENTER, rtl=True)

    # السنة الجامعية
    if req.year:
        yr_pill = rrect(slide, W/2-2.5, H-1.3, 5, 0.75, _rgb(T.accent), radius_pct=50)
        if yr_pill:
            gradient_fill(yr_pill, T.accent_grad1, T.accent_grad2, 0)
            shadow(yr_pill, blur=8, dist=2, alpha=0.3)
        txt(slide, req.year, W/2-2.5, H-1.3, 5, 0.75,
            font="Calibri", size=12, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.CENTER)
    return slide

# ══════════════════════════════════════════════════════════════════════
# INTRO — Two floating cards side by side
# ══════════════════════════════════════════════════════════════════════
def make_intro(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "مقدمة البحث", "نظرة عامة وأسلوب المعالجة")

    items = []
    if req.intro_overview: items.append(("🔍  نظرة عامة", req.intro_overview))
    if req.intro_approach:  items.append(("⚙️  المنهج المتبع", req.intro_approach))
    if not items: return slide

    cols = len(items)
    cw = (W-2.8)/cols - 0.2
    for i,(lbl,val) in enumerate(items):
        x = 1.2 + i*(cw+0.4)
        _card(slide, x, 3.6, cw, H-4.3, T)
        # رأس البطاقة
        ch = rrect(slide, x, 3.6, cw, 0.85, _rgb(T.accent), radius_pct=12)
        if ch: gradient_fill(ch, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, x+0.3, 3.63, cw-0.6, 0.82,
            font=_FONT, size=13, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, val, x+0.35, 4.6, cw-0.7, H-5.3,
            font=_FONT, size=12, bold=False, spacing=18,
            color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# PLAN — Cards with numbered chapters
# ══════════════════════════════════════════════════════════════════════
def make_plan(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "خطة البحث", f"يتكون البحث من {len(req.chapters)} فصول رئيسية")

    chs = req.chapters[:8]
    n = len(chs)
    if not n: return slide

    cols = 3 if n >= 5 else 2 if n >= 3 else 1
    rows = (n+cols-1)//cols
    gap = 0.3
    cw = (W-2.4-(cols-1)*gap)/cols
    avail_h = H-4.1
    ch_h = min(avail_h/rows-gap, 3.2)

    for i,ch in enumerate(chs):
        ci = i%cols; ri = i//cols
        x = 1.2 + ci*(cw+gap)
        y = 3.7 + ri*(ch_h+gap)
        _card(slide, x, y, cw, ch_h, T, radius=10)
        # رقم الفصل — دائرة
        num_c = oval(slide, x+0.3, y+(ch_h-1.1)/2, 1.1, 1.1, _rgb(T.accent))
        if num_c: gradient_fill(num_c, T.accent_grad1, T.accent_grad2, 45)
        txt(slide, str(i+1), x+0.3, y+(ch_h-1.1)/2, 1.1, 1.1,
            font="Calibri", size=16, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.CENTER)
        txt(slide, ch.title, x+1.7, y+0.2, cw-2.1, ch_h-0.5,
            font=_FONT, size=12, bold=False,
            color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
        if ch.pages:
            txt(slide, ch.pages, x+0.2, y+ch_h-0.55, cw-0.4, 0.45,
                font="Calibri", size=9, color=_rgb(T.muted), align=PP_ALIGN.CENTER)
    return slide

# ══════════════════════════════════════════════════════════════════════
# PROBLEM — Dramatic focus card + sub-questions
# ══════════════════════════════════════════════════════════════════════
def make_problem(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "إشكالية البحث", "التساؤلات الرئيسية والفرعية")
    y = 3.65

    if req.main_problem:
        _card(slide, 1.2, y, W-2.4, 2.6, T, radius=12)
        # أيقونة علامة الاستفهام
        qm = rrect(slide, 1.2, y, 2.2, 2.6, _rgb(T.accent), radius_pct=12)
        if qm: gradient_fill(qm, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, "؟", 1.2, y, 2.2, 2.6,
            font="Calibri", size=52, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.CENTER)
        txt(slide, "الإشكالية الرئيسية", W-5.0, y+0.15, 3.6, 0.55,
            font=_FONT, size=10, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, req.main_problem, 3.8, y+0.75, W-5.4, 1.7,
            font=_FONT, size=13, bold=False, spacing=20,
            color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
        y += 2.85

    if req.main_question:
        q_card = rrect(slide, 1.2, y, W-2.4, 1.3, _rgb(T.bg2), radius_pct=10)
        hline(slide, 1.2, y+1.22, W-2.4, _rgb(T.accent), thickness=0.08)
        txt(slide, "❝  "+req.main_question, 1.6, y+0.1, W-3.2, 1.1,
            font=_FONT, size=12.5, bold=True, italic=True,
            color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
        y += 1.5

    if req.sub_questions:
        avail = H-y-0.4
        sq_h = min(avail/max(len(req.sub_questions),1), 0.85)
        cols2 = 2 if len(req.sub_questions)>3 else 1
        cw2 = (W-2.8)/cols2 - 0.2
        for i,q in enumerate(req.sub_questions[:6]):
            ci = i%cols2; ri = i//cols2
            qx = 1.2+ci*(cw2+0.4)
            qy = y+ri*sq_h
            # نقطة ملونة
            oval(slide, qx+cw2-0.7, qy+sq_h*0.3, 0.38, 0.38, _rgb(T.accent))
            txt(slide, q, qx+0.15, qy+0.05, cw2-1.0, sq_h-0.1,
                font=_FONT, size=11, color=_rgb(T.muted),
                align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# OBJECTIVES — Two columns with check items
# ══════════════════════════════════════════════════════════════════════
def make_objectives(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "أهداف البحث وفرضياته")

    cols_data = []
    if req.objectives:  cols_data.append(("🎯  الأهداف", req.objectives))
    if req.hypotheses:  cols_data.append(("💡  الفرضيات", req.hypotheses))
    if not cols_data: return slide

    n_cols = len(cols_data)
    cw = (W-2.8)/n_cols - 0.2
    ph = H-4.3

    for i,(lbl,items) in enumerate(cols_data[:2]):
        x = 1.2+i*(cw+0.4)
        _card(slide, x, 3.65, cw, ph, T, radius=12)
        # رأس العمود
        ch2 = rrect(slide, x, 3.65, cw, 0.8, _rgb(T.accent), radius_pct=12)
        if ch2: gradient_fill(ch2, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, x+0.3, 3.68, cw-0.6, 0.75,
            font=_FONT, size=13, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.RIGHT, rtl=True)
        # العناصر
        avail_h = ph-0.98
        ih = min(avail_h/max(len(items),1), 1.15)
        for j,item in enumerate(items[:8]):
            iy = 4.6+j*ih
            if j>0: hline(slide, x+0.3, iy, cw-0.6, _rgb(T.bg), thickness=0.04)
            # دائرة صغيرة للرقم
            nc = oval(slide, x+cw-0.9, iy+ih*0.2, 0.52, 0.52, _rgb(T.accent))
            txt(slide, str(j+1), x+cw-0.9, iy+ih*0.2, 0.52, 0.52,
                font="Calibri", size=9, bold=True,
                color=_rgb(T.text_dark), align=PP_ALIGN.CENTER)
            txt(slide, item, x+0.3, iy+0.07, cw-1.5, ih-0.14,
                font=_FONT, size=11, color=_rgb(T.text_light),
                align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# IMPORTANCE — Grid cards
# ══════════════════════════════════════════════════════════════════════
def make_importance(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "أهمية البحث ومبررات اختياره")

    items = list(req.importance or [])
    if req.reasons and req.reasons not in items: items.append(req.reasons)
    items = items[:6]
    if not items: return slide

    icons = ["◆","★","✦","◉","▶","✿"]
    cols = 3 if len(items)>=4 else 2 if len(items)>=2 else 1
    rows = (len(items)+cols-1)//cols
    gap = 0.32
    cw = (W-2.4-(cols-1)*gap)/cols
    avail_h = H-4.2
    ch2 = min(avail_h/rows-gap, 4.0)

    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=1.2+ci*(cw+gap); y=3.7+ri*(ch2+gap)
        _card(slide, x, y, cw, ch2, T, radius=14)
        # شريط علوي ملون
        top2 = rrect(slide, x, y, cw, 0.7, _rgb(T.accent), radius_pct=14)
        if top2: gradient_fill(top2, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, icons[i%len(icons)], x+0.2, y+0.06, 1.0, 0.6,
            font="Calibri", size=14, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.LEFT)
        txt(slide, f"{i+1:02d}", x+cw-1.1, y+0.06, 0.9, 0.6,
            font="Calibri", size=13, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.RIGHT)
        txt(slide, item, x+0.3, y+0.85, cw-0.6, ch2-1.0,
            font=_FONT, size=11.5, spacing=17,
            color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# METHODOLOGY — Horizontal flow with icons
# ══════════════════════════════════════════════════════════════════════
def make_methodology(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "منهجية البحث", "الإجراءات والأدوات المستخدمة")

    fields = []
    if req.methodology:  fields.append(("المنهج","M", req.methodology))
    if req.sample_type:  fields.append(("العينة","S", req.sample_type))
    if req.sample_size:  fields.append(("الحجم","N", req.sample_size))
    if req.tool:         fields.append(("الأداة","T", req.tool))
    if not fields: return slide

    n = len(fields)
    cw = (W-2.4)/n - 0.28
    ch3 = H-4.3

    for i,(lbl,icon,val) in enumerate(fields[:4]):
        x = 1.2+i*(cw+0.28)
        _card(slide, x, 3.65, cw, ch3, T, radius=14)
        # دائرة الأيقونة الكبيرة
        ic_sz = 1.6
        ic_x = x+(cw-ic_sz)/2
        ic = oval(slide, ic_x, 3.95, ic_sz, ic_sz, _rgb(T.accent))
        if ic:
            gradient_fill(ic, T.accent_grad1, T.accent_grad2, 45)
            shadow(ic, blur=12, dist=4, alpha=0.4)
        txt(slide, icon, ic_x, 3.95, ic_sz, ic_sz,
            font="Calibri", size=22, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.CENTER)
        # خط فاصل
        hline(slide, x+0.4, 5.75, cw-0.8, _rgb(T.accent), thickness=0.06)
        txt(slide, lbl, x+0.2, 5.85, cw-0.4, 0.7,
            font=_FONT, size=13, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.CENTER, rtl=True)
        txt(slide, val, x+0.25, 6.65, cw-0.5, ch3-3.2,
            font=_FONT, size=11, spacing=17,
            color=_rgb(T.text_light), align=PP_ALIGN.CENTER, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# STATS — Big KPI cards
# ══════════════════════════════════════════════════════════════════════
def make_stats(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "الإحصاءات والأرقام الرئيسية", "مؤشرات كمية للدراسة")

    stats = req.stats[:6]
    if not stats: return slide
    n=len(stats)
    cols=3 if n>=3 else n
    rows=(n+cols-1)//cols
    gap=0.3
    cw=(W-2.4-(cols-1)*gap)/cols
    avail=H-4.2
    ch4=min(avail/rows-gap, 4.8)

    for i,st in enumerate(stats):
        ci=i%cols; ri=i//cols
        x=1.2+ci*(cw+gap); y=3.7+ri*(ch4+gap)
        _card(slide, x, y, cw, ch4, T, radius=16)
        # دائرة خلفية شبه شفافة
        c_sz=min(cw,ch4)*0.7
        bg_c=oval(slide, x+(cw-c_sz)/2, y+(ch4-c_sz)/2-0.3, c_sz, c_sz, _rgb(T.accent), alpha=6)
        # القيمة الكبيرة
        vs=36 if len(st.value)<=4 else 28 if len(st.value)<=8 else 20
        txt(slide, st.value, x+0.3, y+0.5, cw-0.6, ch4*0.52,
            font="Calibri", size=vs, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.CENTER)
        if st.unit:
            txt(slide, st.unit, x+0.3, y+ch4*0.52+0.35, cw-0.6, 0.6,
                font=_FONT, size=10, color=_rgb(T.muted), align=PP_ALIGN.CENTER, rtl=True)
        hline(slide, x+cw*0.2, y+ch4-1.1, cw*0.6, _rgb(T.accent), thickness=0.07)
        txt(slide, st.label, x+0.3, y+ch4-0.95, cw-0.6, 0.85,
            font=_FONT, size=11.5, color=_rgb(T.text_light),
            align=PP_ALIGN.CENTER, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# RESULTS — Numbered rows
# ══════════════════════════════════════════════════════════════════════
def make_results(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "نتائج البحث", "أبرز ما توصلت إليه الدراسة")

    results=req.main_results[:8]
    if not results: return slide
    avail=H-4.2
    ih=min(avail/max(len(results),1)-0.15, 1.6)

    for i,res in enumerate(results):
        y=3.7+i*(ih+0.15)
        row=rrect(slide, 1.2, y, W-2.4, ih, _rgb(T.card if i%2==0 else T.bg2), radius_pct=10)
        if row: shadow(row, blur=6, dist=2, alpha=0.2)
        # رقم في دائرة
        nc=oval(slide, 1.5, y+(ih-0.7)/2, 0.7, 0.7, _rgb(T.accent))
        if nc: gradient_fill(nc, T.accent_grad1, T.accent_grad2, 45)
        txt(slide, str(i+1), 1.5, y+(ih-0.7)/2, 0.7, 0.7,
            font="Calibri", size=11, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.CENTER)
        txt(slide, res, 2.6, y+0.1, W-4.2, ih-0.2,
            font=_FONT, size=12, color=_rgb(T.text_light),
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# CONCLUSION — Large quote card
# ══════════════════════════════════════════════════════════════════════
def make_conclusion(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "خاتمة البحث", "الاستنتاج العام")
    # دائرة زخرفية خلفية
    oval(slide, W*0.3, H*0.3, H*0.9, H*0.9, _rgb(T.accent), alpha=4)

    cy=3.65; ch5=H-cy-0.6
    _card(slide, 1.2, cy, W-2.4, ch5, T, radius=16)
    # شريط علوي
    ct=rrect(slide, 1.2, cy, W-2.4, 0.6, _rgb(T.accent), radius_pct=16)
    if ct: gradient_fill(ct, T.accent_grad1, T.accent_grad2, 0)
    # علامة اقتباس كبيرة
    txt(slide, "❝", 1.8, cy+0.7, 3, 2.2,
        font="Calibri", size=65, color=_rgb(T.accent), align=PP_ALIGN.LEFT)
    txt(slide, req.general_conclusion,
        4.2, cy+0.9, W-5.8, ch5-1.4,
        font=_FONT, size=14, spacing=22,
        color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# RECOMMENDATIONS
# ══════════════════════════════════════════════════════════════════════
def make_recommendations(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "توصيات البحث", "المقترحات والحلول العملية")

    recs=req.recommendations[:8]
    if not recs: return slide
    avail=H-4.2
    ih=min(avail/max(len(recs),1)-0.12, 1.5)

    for i,rec in enumerate(recs):
        y=3.7+i*(ih+0.12)
        row=rrect(slide, 1.2, y, W-2.4, ih, _rgb(T.card if i%2==0 else T.bg2), radius_pct=10)
        if row and i%2==0: shadow(row, blur=5, dist=1, alpha=0.15)
        # شريط لوني يسار
        bar=rrect(slide, 1.2, y, 0.45, ih, _rgb(T.accent), radius_pct=10)
        if bar: gradient_fill(bar, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, rec, 2.0, y+0.1, W-3.6, ih-0.2,
            font=_FONT, size=12, color=_rgb(T.text_light),
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# FUTURE WORK
# ══════════════════════════════════════════════════════════════════════
def make_future(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "آفاق البحث المستقبلية", "مسارات الاستكشاف القادمة")

    items=req.future_work[:6]
    if not items: return slide
    cols=2 if len(items)>2 else 1
    rows=(len(items)+cols-1)//cols
    gap=0.3
    cw=(W-2.4-(cols-1)*gap)/cols
    avail=H-4.2
    ch6=min(avail/rows-gap, 3.2)

    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=1.2+ci*(cw+gap); y=3.7+ri*(ch6+gap)
        _card(slide, x, y, cw, ch6, T, radius=12)
        # رقم كبير في الزاوية
        txt(slide, f"0{i+1}", x+0.3, y+0.2, 1.5, 0.9,
            font="Calibri", size=22, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.LEFT)
        hline(slide, x+0.3, y+ch6-0.4, cw-0.6, _rgb(T.accent), thickness=0.07)
        txt(slide, item, x+0.35, y+1.15, cw-0.7, ch6-1.6,
            font=_FONT, size=12, spacing=17,
            color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# REFERENCES
# ══════════════════════════════════════════════════════════════════════
def make_references(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "المراجع والمصادر", "قائمة المصادر والمراجع المعتمدة")

    refs=req.references[:14]
    if not refs: return slide
    avail=H-4.2
    ih=max(min(avail/max(len(refs),1)-0.08, 1.0), 0.48)

    for i,ref in enumerate(refs):
        y=3.7+i*(ih+0.08)
        if y+ih>H-0.3: break
        if i%2==0:
            rb=rrect(slide, 1.2, y, W-2.4, ih, _rgb(T.bg2), radius_pct=6)
        # رقم المرجع
        txt(slide, f"[{i+1}]", 1.4, y+0.04, 1.4, ih-0.08,
            font="Calibri", size=9, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.LEFT)
        txt(slide, ref, 3.1, y+0.04, W-4.5, ih-0.08,
            font=_FONT, size=10, color=_rgb(T.text_light),
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# FINAL — Thank you with floating elements
# ══════════════════════════════════════════════════════════════════════
def make_final(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 150)
    # دوائر زخرفية كبيرة
    oval(slide, -5, -5, 20, 20, _rgb(T.accent), alpha=5)
    oval(slide, W-12, H-10, 18, 18, _rgb(T.accent), alpha=6)
    oval(slide, W*0.25, -4, 14, 14, _rgb(T.bg2), alpha=70)
    # شريط علوي وسفلي
    top=rect(slide, 0, 0, W, 0.5, _rgb(T.accent))
    if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)
    bot=rect(slide, 0, H-0.5, W, 0.5, _rgb(T.accent))
    if bot: gradient_fill(bot, T.accent_grad2, T.accent_grad1, 0)

    # البطاقة المركزية
    cw=26; ch=12
    cx=(W-cw)/2; cy=(H-ch)/2
    mc=rrect(slide, cx, cy, cw, ch, _rgb(T.card), radius_pct=16)
    if mc:
        gradient_fill(mc, T.card, T.bg2, 135)
        shadow(mc, blur=30, dist=10, alpha=0.55)
    ct2=rrect(slide, cx, cy, cw, 0.55, _rgb(T.accent), radius_pct=16)
    if ct2: gradient_fill(ct2, T.accent_grad1, T.accent_grad2, 0)

    txt(slide, "شكراً وتقديراً",
        cx+1, cy+0.7, cw-2, 3.0,
        font=_FONT, size=42, bold=True,
        color=_rgb(T.text_light), align=PP_ALIGN.CENTER, rtl=True)

    # نقاط زخرفية
    for j in range(5):
        dx=cx+cw/2-1.5+j*0.75
        sz=0.42 if j==2 else 0.26
        dc=oval(slide, dx, cy+3.9, sz, sz, _rgb(T.accent))
        if dc and j!=2: set_solid_alpha(dc, 55)

    txt(slide, req.student_name, cx+1, cy+4.5, cw-2, 1.3,
        font=_FONT, size=22, bold=True,
        color=_rgb(T.accent), align=PP_ALIGN.CENTER, rtl=True)

    short=req.title_ar[:80]+("..." if len(req.title_ar)>80 else "")
    txt(slide, short, cx+1.5, cy+5.9, cw-3, 2.5,
        font=_FONT, size=12, italic=True,
        color=_rgb(T.muted), align=PP_ALIGN.CENTER, rtl=True)

    hline(slide, cx+cw*0.15, cy+ch-1.4, cw*0.7, _rgb(T.accent), thickness=0.05)
    footer=[]
    if req.institution: footer.append(req.institution.split(' — ')[0])
    if req.year: footer.append(req.year)
    if footer:
        txt(slide, " · ".join(footer), cx+1, cy+ch-1.25, cw-2, 0.75,
            font=_FONT, size=11, color=_rgb(T.muted),
            align=PP_ALIGN.CENTER, rtl=True)
    return slide
