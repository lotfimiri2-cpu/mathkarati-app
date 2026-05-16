"""
Slide Builder — PREMIUM Engine — مذكرتي Pro v18
فلسفة التصميم: Editorial Magazine — الصحافة الراقية
- تخطيط غير متماثل asymmetric بجرأة
- عناوين ضخمة تملأ المساحة
- شرائط لونية قطرية وبادجات pill
- طبقات متعددة: خلفية + overlay + محتوى
- مناسب: علوم اقتصادية، إدارة، هندسة
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

def _pill(slide, T, x, y, w, h_val, text, font_size=10, alpha=100):
    """بادج pill مستدير"""
    p = rrect(slide, x, y, w, h_val, _rgb(T.accent), radius_pct=50)
    if p:
        gradient_fill(p, T.accent_grad1, T.accent_grad2, 0)
        if alpha < 100: set_solid_alpha(p, alpha)
        shadow(p, blur=6, dist=2, alpha=0.25)
    txt(slide, text, x, y, w, h_val,
        font=_FONT, size=font_size, bold=True,
        color=_rgb(T.text_dark), align=PP_ALIGN.CENTER, rtl=True)

def _header(slide, T, title, subtitle=""):
    """رأس premium — شريط رفيع + عنوان ضخم يميني"""
    # شريط علوي رفيع جداً
    top = rect(slide, 0, 0, W, 0.22, _rgb(T.accent))
    if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)
    # خلفية رأس بتدرج عميق
    gradient_rect(slide, 0, 0.22, W, 3.25, T.grad1, T.grad2, 90)
    # نقطة لونية يمين
    oval(slide, W-1.6, 0.42, 0.55, 0.55, _rgb(T.accent))
    # العنوان — ضخم يميني
    sz = 26 if len(title)<18 else 20 if len(title)<28 else 17
    txt(slide, title, 1.0, 0.3, W-2.2, 2.1,
        font=_FONT, size=sz, bold=True,
        color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    if subtitle:
        txt(slide, subtitle, 1.0, 2.2, W-2.2, 0.9,
            font=_FONT, size=11.5, color=_rgb(T.muted),
            align=PP_ALIGN.RIGHT, rtl=True)
    # خط سفلي مع دائرة
    hline(slide, 0, 3.32, W, _rgb(T.accent), thickness=0.06)
    oval(slide, W*0.5-0.28, 3.19, 0.56, 0.56, _rgb(T.accent))

# ══════════════════════════════════════════════════════════════════════
# COVER — Full-bleed editorial
# ══════════════════════════════════════════════════════════════════════
def make_cover(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    # خلفية كاملة بتدرج قطري
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 145)
    # طبقة overlay للعمق
    ov = rect(slide, 0, 0, W, H, _rgb(T.bg))
    if ov: set_solid_alpha(ov, 25)
    # دوائر ضخمة خلفية
    oval(slide, W*0.5, -H*0.35, H*1.4, H*1.4, _rgb(T.accent), alpha=6)
    oval(slide, -H*0.3, H*0.4,  H*1.1, H*1.1, _rgb(T.bg2),   alpha=45)
    oval(slide, W*0.7, H*0.6,   H*0.7, H*0.7, _rgb(T.accent), alpha=4)

    # شريط علوي رفيع
    top = rect(slide, 0, 0, W, 0.22, _rgb(T.accent))
    if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)

    # المؤسسة — pill يمين
    if req.institution:
        inst = req.institution.split(' — ')[0]
        _pill(slide, T, W-len(inst)*0.18-2.5, 0.45, len(inst)*0.18+2.3, 0.62, inst, 10)

    # كتلة العنوان الكبرى — يملأ الثلث الأعلى
    tsz = 30 if len(req.title_ar)<32 else 24 if len(req.title_ar)<52 else 18 if len(req.title_ar)<80 else 14
    txt(slide, req.title_ar,
        1.2, H*0.1, W-2.4, H*0.45,
        font=_FONT, size=tsz, bold=True,
        color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)

    if req.title_en:
        txt(slide, req.title_en,
            1.2, H*0.55-0.3, W-2.4, 0.85,
            font="Calibri", size=11, italic=True,
            color=_rgb(T.muted), align=PP_ALIGN.RIGHT)

    # خط فاصل جريء
    hline(slide, 1.2, H*0.6, W-2.4, _rgb(T.accent), thickness=0.14)

    # بادجات المعلومات — صف أفقي
    fields = [req.student_name]
    if req.supervisor:    fields.append(req.supervisor)
    if req.specialization: fields.append(req.specialization)
    if req.year:           fields.append(req.year)

    badge_h = 0.65
    px = W - 1.2
    for val in fields[:4]:
        bw = min(len(val)*0.2+1.2, 8.5)
        px -= bw + 0.3
        _pill(slide, T, px, H*0.65, bw, badge_h, val, 11)
    # تسميات صغيرة فوق البادجات
    labels = ["الطالب","المشرف","التخصص","السنة"]
    px2 = W - 1.2
    for i,val in enumerate(fields[:4]):
        bw = min(len(val)*0.2+1.2, 8.5)
        px2 -= bw + 0.3
        txt(slide, labels[i], px2, H*0.65-0.45, bw, 0.4,
            font=_FONT, size=8.5, color=_rgb(T.muted),
            align=PP_ALIGN.CENTER, rtl=True)

    # شريط سفلي
    bot = rect(slide, 0, H-0.3, W, 0.3, _rgb(T.accent))
    if bot: gradient_fill(bot, T.accent_grad2, T.accent_grad1, 0)
    return slide

# ══════════════════════════════════════════════════════════════════════
# INTRO — Full panels asymmetric
# ══════════════════════════════════════════════════════════════════════
def make_intro(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "مقدمة البحث", "نظرة عامة وأسلوب المعالجة")

    items=[]
    if req.intro_overview: items.append(("نظرة عامة", req.intro_overview))
    if req.intro_approach:  items.append(("المنهج المتبع", req.intro_approach))
    if not items: return slide

    # لوحة كاملة بعرض غير متساوٍ (60/40)
    widths = [_CONT_W()*0.58, _CONT_W()*0.39] if len(items)==2 else [_CONT_W()]
    gap = 0.4; cy = 3.55; ch = H-cy-0.5

    x = 1.0
    for i,(lbl,val) in enumerate(items[:2]):
        cw = widths[i] if i<len(widths) else widths[-1]
        # بطاقة كاملة
        card = rrect(slide, x, cy, cw, ch, _rgb(T.card), radius_pct=8)
        if card:
            gradient_fill(card, T.card, T.bg2, 150)
            shadow(card, blur=20, dist=6, alpha=0.45)
        # شريط علوي قطري (مستطيل ملون يملأ الزاوية)
        corner_w = cw*0.42
        corner = rect(slide, x, cy, corner_w, 0.7, _rgb(T.accent))
        if corner: gradient_fill(corner, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, x+corner_w+0.2, cy+0.06, cw-corner_w-0.4, 0.62,
            font=_FONT, size=13, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, val, x+0.35, cy+0.88, cw-0.7, ch-1.1,
            font=_FONT, size=12, spacing=20,
            color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
        x += cw + gap
    return slide

def _CONT_W(): return W - 2.0

# ══════════════════════════════════════════════════════════════════════
# PLAN — Timeline horizontal أو grid
# ══════════════════════════════════════════════════════════════════════
def make_plan(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "خطة البحث", f"يتكون البحث من {len(req.chapters)} فصول رئيسية")

    chs=req.chapters[:8]; n=len(chs)
    if not n: return slide
    cy=3.55; avail_h=H-cy-0.4

    if n<=4:
        # تايم‌لاين أفقي
        cw=(W-2.4)/n-0.28
        ch=avail_h-0.85
        # خط تايم‌لاين
        hline(slide, 1.2+cw/2, cy+0.42, (W-2.4)-(cw/2)*2, _rgb(T.muted), thickness=0.04)
        for i,chapter in enumerate(chs):
            x=1.2+i*(cw+0.28)
            # نقطة Timeline
            node=oval(slide, x+cw/2-0.45, cy+0.14, 0.9, 0.9, _rgb(T.accent))
            if node: gradient_fill(node, T.accent_grad1, T.accent_grad2, 45)
            txt(slide, str(i+1), x+cw/2-0.45, cy+0.14, 0.9, 0.9,
                font="Calibri", size=14, bold=True,
                color=_rgb(T.text_dark), align=PP_ALIGN.CENTER)
            # البطاقة أسفل النقطة
            card=rrect(slide, x, cy+1.1, cw, ch, _rgb(T.card), radius_pct=10)
            if card: shadow(card, blur=10, dist=3, alpha=0.35)
            top=rect(slide, x, cy+1.1, cw, 0.32, _rgb(T.accent))
            if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)
            txt(slide, chapter.title, x+0.2, cy+1.5, cw-0.4, ch-0.65,
                font=_FONT, size=11, spacing=17,
                color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
            if chapter.pages:
                txt(slide, chapter.pages, x+0.2, cy+1.1+ch-0.55, cw-0.4, 0.45,
                    font="Calibri", size=9, color=_rgb(T.muted), align=PP_ALIGN.CENTER)
    else:
        # شبكة 2 × N
        cols=2; rows=(n+1)//2
        gap=0.3
        cw=(_CONT_W()-gap)/2
        ch=min(avail_h/rows-gap, 1.9)
        for i,chapter in enumerate(chs):
            ci=i%2; ri=i//2
            x=1.0+ci*(cw+gap); y=cy+ri*(ch+gap)
            card=rrect(slide, x, y, cw, ch, _rgb(T.card), radius_pct=8)
            if card: shadow(card, blur=8, dist=2, alpha=0.3)
            # pill رقم يمين
            _pill(slide, T, x+cw-1.5, y+(ch-0.5)/2, 1.4, 0.5, str(i+1), 11)
            vline(slide, x, y, ch, _rgb(T.accent), thickness=0.35)
            txt(slide, chapter.title, x+0.55, y+0.1, cw-2.3, ch-0.2,
                font=_FONT, size=12, spacing=17,
                color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# PROBLEM — Dramatic asymmetric
# ══════════════════════════════════════════════════════════════════════
def make_problem(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "إشكالية البحث", "التساؤلات الرئيسية والفرعية")

    y=3.55
    if req.main_problem:
        ch=2.55
        card=rrect(slide, 0.8, y, W-1.6, ch, _rgb(T.card), radius_pct=10)
        if card: shadow(card, blur=18, dist=6, alpha=0.5)
        # overlay يساري ملون
        left=rrect(slide, 0.8, y, 3.8, ch, _rgb(T.accent), radius_pct=10)
        if left:
            gradient_fill(left, T.accent_grad1, T.accent_grad2, 90)
            set_solid_alpha(left, 28)
        txt(slide, "؟", 1.0, y+0.1, 3.2, ch-0.2,
            font="Calibri", size=55, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.CENTER)
        txt(slide, "الإشكالية الرئيسية",
            W-5.2, y+0.12, 4.2, 0.58,
            font=_FONT, size=10, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, req.main_problem,
            4.8, y+0.72, W-5.9, ch-0.9,
            font=_FONT, size=13, spacing=20,
            color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
        y += ch+0.3

    if req.main_question:
        qh=1.3
        qc=rrect(slide, 0.8, y, W-1.6, qh, _rgb(T.bg2), radius_pct=8)
        hline(slide, 0.8, y+qh-0.1, W-1.6, _rgb(T.accent), thickness=0.1)
        txt(slide, "❝  "+req.main_question,
            1.3, y+0.12, W-2.4, qh-0.24,
            font=_FONT, size=13, bold=True, italic=True,
            color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
        y += qh+0.25

    if req.sub_questions:
        avail=H-y-0.4
        sq_h=min(avail/max(len(req.sub_questions),1), 0.85)
        cols2=2 if len(req.sub_questions)>3 else 1
        cw2=(_CONT_W()-(cols2-1)*0.3)/cols2
        for i,q in enumerate(req.sub_questions[:6]):
            ci=i%cols2; ri=i//cols2
            qx=1.0+ci*(cw2+0.3)
            qy=y+ri*sq_h
            # pill رقم صغير
            _pill(slide, T, qx+cw2-1.4, qy+sq_h*0.22, 1.2, 0.42, str(i+1), 9)
            txt(slide, q, qx+0.2, qy+0.06, cw2-1.8, sq_h-0.12,
                font=_FONT, size=11.5, color=_rgb(T.muted),
                align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# OBJECTIVES — Asymmetric 55/45
# ══════════════════════════════════════════════════════════════════════
def make_objectives(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "أهداف البحث وفرضياته")

    cols_data=[]
    if req.objectives: cols_data.append(("الأهداف", req.objectives))
    if req.hypotheses:  cols_data.append(("الفرضيات", req.hypotheses))
    if not cols_data: return slide

    cy=3.55; ch=H-cy-0.45; gap=0.4
    if len(cols_data)==2:
        widths=[_CONT_W()*0.56, _CONT_W()*0.4]
    else:
        widths=[_CONT_W()]

    x=1.0
    for i,(lbl,items) in enumerate(cols_data[:2]):
        cw=widths[i]
        card=rrect(slide, x, cy, cw, ch, _rgb(T.card), radius_pct=10)
        if card:
            gradient_fill(card, T.card, T.bg2, 150)
            shadow(card, blur=16, dist=5, alpha=0.4)
        # رأس بتدرج
        hdr=rect(slide, x, cy, cw, 0.78, _rgb(T.accent))
        if hdr: gradient_fill(hdr, T.accent_grad1, T.accent_grad2, 0)
        txt(slide, lbl, x+0.25, cy+0.06, cw-0.5, 0.7,
            font=_FONT, size=14, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.CENTER, rtl=True)
        # العناصر
        avail_h=ch-0.9
        ih=min(avail_h/max(len(items),1), 1.1)
        for j,item in enumerate(items[:8]):
            iy=cy+0.88+j*ih
            if j>0: hline(slide, x+0.25, iy, cw-0.5, _rgb(T.bg), thickness=0.035)
            # pill رقم يمين
            _pill(slide, T, x+cw-1.35, iy+ih*0.2, 1.15, 0.48, str(j+1), 9)
            txt(slide, item, x+0.25, iy+0.06, cw-1.7, ih-0.12,
                font=_FONT, size=11, color=_rgb(T.text_light),
                align=PP_ALIGN.RIGHT, rtl=True)
        x += cw+gap
    return slide

# ══════════════════════════════════════════════════════════════════════
# IMPORTANCE — Magazine grid
# ══════════════════════════════════════════════════════════════════════
def make_importance(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "أهمية البحث", "مبررات اختيار الموضوع")

    items=list(req.importance or [])
    if req.reasons and req.reasons not in items: items.append(req.reasons)
    items=items[:6]
    if not items: return slide

    # ارتفاعات علوية متنوعة لكل بطاقة
    accents=[0.62, 0.78, 0.55, 0.7, 0.65, 0.8]
    cols=3 if len(items)>=4 else 2 if len(items)>=2 else 1
    rows=(len(items)+cols-1)//cols
    gap=0.3; cy=3.55; avail_h=H-cy-0.4
    cw=(_CONT_W()-(cols-1)*gap)/cols
    ch=min(avail_h/rows-gap, 3.8)

    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(cw+gap); y=cy+ri*(ch+gap)
        card=rrect(slide, x, y, cw, ch, _rgb(T.card), radius_pct=6)
        if card: shadow(card, blur=12, dist=4, alpha=0.35)
        # شريط علوي بارتفاع متنوع
        ah=accents[i%len(accents)]
        top=rect(slide, x, y, cw, ah, _rgb(T.accent))
        if top:
            gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)
            set_solid_alpha(top, 92)
        # رقم يساري على الشريط
        txt(slide, f"{i+1:02d}", x+0.18, y+0.04, 1.1, ah-0.08,
            font="Calibri", size=18, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.LEFT)
        txt(slide, item, x+0.25, y+ah+0.15, cw-0.5, ch-ah-0.28,
            font=_FONT, size=11.5, spacing=18,
            color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# METHODOLOGY — Icon flow horizontal
# ══════════════════════════════════════════════════════════════════════
def make_methodology(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.bg2, 180)
    _header(slide, T, "منهجية البحث", "الإجراءات والأدوات المستخدمة")

    fields=[]
    if req.methodology:  fields.append(("المنهج",    "M", req.methodology))
    if req.sample_type:  fields.append(("العينة",    "S", req.sample_type))
    if req.sample_size:  fields.append(("الحجم",     "N", req.sample_size))
    if req.tool:         fields.append(("الأداة",    "T", req.tool))
    if not fields: return slide

    n=len(fields); cy=3.55; ch=H-cy-0.45
    cw=(_CONT_W()-0.28*(n-1))/n

    # خط رابط بين البطاقات
    if n>1:
        hline(slide, 1.0+cw/2, cy+1.05,
              (W-2.0)-cw/2*2, _rgb(T.muted), thickness=0.04)

    for i,(lbl,icon,val) in enumerate(fields[:4]):
        x=1.0+i*(cw+0.28)
        card=rrect(slide, x, cy, cw, ch, _rgb(T.card), radius_pct=12)
        if card:
            shadow(card, blur=14, dist=5, alpha=0.4)
        # دائرة icon ضخمة
        ic_sz=1.7; ic_x=x+(cw-ic_sz)/2
        ic=oval(slide, ic_x, cy+0.5, ic_sz, ic_sz, _rgb(T.accent))
        if ic:
            gradient_fill(ic, T.accent_grad1, T.accent_grad2, 45)
            shadow(ic, blur=10, dist=3, alpha=0.4)
        txt(slide, icon, ic_x, cy+0.5, ic_sz, ic_sz,
            font="Calibri", size=24, bold=True,
            color=_rgb(T.text_dark), align=PP_ALIGN.CENTER)
        # تسمية + خط + محتوى
        hline(slide, x+0.3, cy+2.45, cw-0.6, _rgb(T.accent), thickness=0.07)
        txt(slide, lbl, x+0.2, cy+2.55, cw-0.4, 0.7,
            font=_FONT, size=13, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.CENTER, rtl=True)
        txt(slide, val, x+0.22, cy+3.32, cw-0.44, ch-3.45,
            font=_FONT, size=11, spacing=17,
            color=_rgb(T.text_light), align=PP_ALIGN.CENTER, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# STATS — Bold KPI with glowing accent
# ══════════════════════════════════════════════════════════════════════
def make_stats(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "الإحصاءات والأرقام الرئيسية", "مؤشرات كمية للدراسة")

    stats=req.stats[:6]
    if not stats: return slide
    n=len(stats)
    cols=3 if n>=3 else n
    rows=(n+cols-1)//cols
    gap=0.32; cy=3.55; avail=H-cy-0.4
    cw=(_CONT_W()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap, 4.6)

    for i,st in enumerate(stats):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(cw+gap); y=cy+ri*(ch+gap)
        card=rrect(slide, x, y, cw, ch, _rgb(T.card), radius_pct=12)
        if card:
            shadow(card, blur=18, dist=6, alpha=0.5)
        # دائرة خلفية شبه شفافة ضخمة
        cs=min(cw,ch)*0.75
        bg_c=oval(slide, x+(cw-cs)/2, y+(ch-cs)/2-0.3, cs, cs, _rgb(T.accent), alpha=6)
        # القيمة الكبيرة
        vs=38 if len(st.value)<=4 else 28 if len(st.value)<=8 else 20
        txt(slide, st.value, x+0.2, y+0.45, cw-0.4, ch*0.5,
            font="Calibri", size=vs, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.CENTER)
        if st.unit:
            txt(slide, st.unit, x+0.2, y+ch*0.5+0.35, cw-0.4, 0.58,
                font=_FONT, size=10.5, color=_rgb(T.muted),
                align=PP_ALIGN.CENTER, rtl=True)
        # pill التسمية
        lbl_w=min(len(st.label)*0.18+1.0, cw-0.4)
        _pill(slide, T, x+(cw-lbl_w)/2, y+ch-1.08, lbl_w, 0.55, st.label, 10, alpha=85)
    return slide

# ══════════════════════════════════════════════════════════════════════
# RESULTS — pill-numbered rows
# ══════════════════════════════════════════════════════════════════════
def make_results(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "نتائج البحث", "أبرز ما توصلت إليه الدراسة")

    results=req.main_results[:8]
    if not results: return slide
    avail=H-3.58; ih=min(avail/max(len(results),1)-0.13, 1.58)

    for i,res in enumerate(results):
        y=3.55+i*(ih+0.13)
        row=rrect(slide, 0.8, y, W-1.6, ih,
                  _rgb(T.bg2) if i%2==0 else _rgb(T.card), radius_pct=6)
        if row and i%2==0: shadow(row, blur=5, dist=1, alpha=0.18)
        # pill رقم يمين
        _pill(slide, T, W-3.1, y+(ih-0.52)/2, 2.0, 0.52, str(i+1), 12)
        # شريط يساري رفيع
        vline(slide, 0.8, y, ih, _rgb(T.accent), thickness=0.3)
        txt(slide, res, 1.4, y+0.1, W-5.0, ih-0.2,
            font=_FONT, size=12, color=_rgb(T.text_light),
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# CONCLUSION — Full dramatic card
# ══════════════════════════════════════════════════════════════════════
def make_conclusion(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "خاتمة البحث", "الاستنتاج العام")

    cy=3.55; ch=H-cy-0.55
    card=rrect(slide, 0.8, cy, W-1.6, ch, _rgb(T.card), radius_pct=12)
    if card: shadow(card, blur=22, dist=7, alpha=0.5)
    # overlay يساري ملون
    left=rrect(slide, 0.8, cy, 3.2, ch, _rgb(T.accent), radius_pct=12)
    if left:
        gradient_fill(left, T.accent_grad1, T.accent_grad2, 90)
        set_solid_alpha(left, 22)
    # شريط علوي
    top=rrect(slide, 0.8, cy, W-1.6, 0.38, _rgb(T.accent), radius_pct=12)
    if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)
    # علامة اقتباس ضخمة
    txt(slide, "❝", 1.2, cy+0.5, 2.8, 2.4,
        font="Calibri", size=62, color=_rgb(T.accent), align=PP_ALIGN.CENTER)
    txt(slide, req.general_conclusion,
        4.3, cy+0.8, W-5.4, ch-1.3,
        font=_FONT, size=14, spacing=23,
        color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# RECOMMENDATIONS
# ══════════════════════════════════════════════════════════════════════
def make_recommendations(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "توصيات البحث", "المقترحات العملية")

    recs=req.recommendations[:8]
    if not recs: return slide
    avail=H-3.58; ih=min(avail/max(len(recs),1)-0.12, 1.5)

    for i,rec in enumerate(recs):
        y=3.55+i*(ih+0.12)
        row=rrect(slide, 0.8, y, W-1.6, ih,
                  _rgb(T.bg2) if i%2==0 else _rgb(T.card), radius_pct=6)
        # شريط لوني يساري
        tag=rrect(slide, 0.8, y, 0.42, ih, _rgb(T.accent), radius_pct=6)
        if tag: gradient_fill(tag, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, rec, 1.5, y+0.1, W-3.2, ih-0.2,
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
    cols=3 if len(items)>=4 else 2 if len(items)>=2 else 1
    rows=(len(items)+cols-1)//cols
    gap=0.3; cy=3.55; avail=H-cy-0.4
    cw=(_CONT_W()-(cols-1)*gap)/cols
    ch=min(avail/rows-gap, 3.2)

    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(cw+gap); y=cy+ri*(ch+gap)
        card=rrect(slide, x, y, cw, ch, _rgb(T.card), radius_pct=10)
        if card: shadow(card, blur=10, dist=3, alpha=0.3)
        # شريط سفلي
        bot=rect(slide, x, y+ch-0.4, cw, 0.4, _rgb(T.accent))
        if bot: gradient_fill(bot, T.accent_grad1, T.accent_grad2, 0)
        # رقم
        txt(slide, f"{i+1:02d}", x+0.2, y+0.15, 1.3, 0.88,
            font="Calibri", size=22, bold=True,
            color=_rgb(T.accent), align=PP_ALIGN.LEFT)
        hline(slide, x+0.2, y+1.05, cw-0.4, _rgb(T.accent), thickness=0.06)
        txt(slide, item, x+0.25, y+1.18, cw-0.5, ch-1.72,
            font=_FONT, size=12, spacing=18,
            color=_rgb(T.text_light), align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# REFERENCES
# ══════════════════════════════════════════════════════════════════════
def make_references(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 180)
    _header(slide, T, "المراجع والمصادر", "قائمة المصادر المعتمدة")

    refs=req.references[:14]
    if not refs: return slide
    avail=H-3.58; ih=max(min(avail/max(len(refs),1)-0.08,1.0),0.48)

    for i,ref in enumerate(refs):
        y=3.55+i*(ih+0.08)
        if y+ih>H-0.3: break
        if i%2==0: rrect(slide, 0.8, y, W-1.6, ih, _rgb(T.bg2), radius_pct=4)
        _pill(slide, T, W-2.9, y+(ih-0.4)/2, 1.8, 0.4, f"[{i+1}]", 9, alpha=75)
        txt(slide, ref, 1.2, y+0.04, W-4.6, ih-0.08,
            font=_FONT, size=10, color=_rgb(T.text_light),
            align=PP_ALIGN.RIGHT, rtl=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# FINAL — Cinematic full-bleed
# ══════════════════════════════════════════════════════════════════════
def make_final(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, _rgb(T.bg))
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, 150)
    # دوائر ضخمة
    oval(slide, -6, -6, 22, 22, _rgb(T.accent), alpha=5)
    oval(slide, W-14, H-12, 20, 20, _rgb(T.accent), alpha=5)
    oval(slide, W*0.28, -5, 16, 16, _rgb(T.bg2), alpha=55)
    # شريط علوي وسفلي
    top=rect(slide, 0, 0, W, 0.5, _rgb(T.accent))
    if top: gradient_fill(top, T.accent_grad1, T.accent_grad2, 0)
    bot=rect(slide, 0, H-0.5, W, 0.5, _rgb(T.accent))
    if bot: gradient_fill(bot, T.accent_grad2, T.accent_grad1, 0)

    # بطاقة مركزية كبيرة
    cw=27; ch=13; cx=(W-cw)/2; cy=(H-ch)/2
    mc=rrect(slide, cx, cy, cw, ch, _rgb(T.card), radius_pct=14)
    if mc:
        gradient_fill(mc, T.card, T.bg2, 145)
        shadow(mc, blur=32, dist=12, alpha=0.55)
    ct=rrect(slide, cx, cy, cw, 0.52, _rgb(T.accent), radius_pct=14)
    if ct: gradient_fill(ct, T.accent_grad1, T.accent_grad2, 0)

    txt(slide, "شكراً وتقديراً",
        cx+1, cy+0.65, cw-2, 3.1,
        font=_FONT, size=44, bold=True,
        color=_rgb(T.text_light), align=PP_ALIGN.CENTER, rtl=True)

    # نقاط زخرفية
    for j in range(5):
        dx=cx+cw/2-1.5+j*0.75
        sz=0.44 if j==2 else 0.26
        dc=oval(slide, dx, cy+3.95, sz, sz, _rgb(T.accent))
        if dc and j!=2: set_solid_alpha(dc, 50)

    txt(slide, req.student_name,
        cx+1, cy+4.55, cw-2, 1.3,
        font=_FONT, size=22, bold=True,
        color=_rgb(T.accent), align=PP_ALIGN.CENTER, rtl=True)

    short=req.title_ar[:82]+("..." if len(req.title_ar)>82 else "")
    txt(slide, short, cx+1.8, cy+6.0, cw-3.6, 2.6,
        font=_FONT, size=12.5, italic=True,
        color=_rgb(T.muted), align=PP_ALIGN.CENTER, rtl=True)

    hline(slide, cx+cw*0.15, cy+ch-1.45, cw*0.7, _rgb(T.accent), thickness=0.05)
    footer=[]
    if req.institution: footer.append(req.institution.split(' — ')[0])
    if req.year: footer.append(req.year)
    if footer:
        txt(slide, " · ".join(footer), cx+1, cy+ch-1.25, cw-2, 0.75,
            font=_FONT, size=11.5, color=_rgb(T.muted),
            align=PP_ALIGN.CENTER, rtl=True)
    return slide
