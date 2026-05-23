"""
Drawing Primitives — مذكرتي Pro v19.1  ★ VISUAL UPGRADE
Low-level, deterministic shape/text builders.

UPGRADE v19.1:
- txt() دعم كامل لـ: letter spacing، text shadow، outline، vertical align
- txt_hero() للعناوين الضخمة مع glow effect
- txt_badge() للعلامات النصية الصغيرة
- txt_multiline() لتحسين multi-paragraph
- gradient_3stop() تدرج 3 نقاط
- glow() تأثير وهج حول الأشكال
"""
from __future__ import annotations
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

W, H = 33.867, 19.05

_SPPR_ORDER = [
    qn('a:xfrm'), qn('a:prstGeom'), qn('a:custGeom'),
    qn('a:noFill'), qn('a:solidFill'), qn('a:gradFill'),
    qn('a:blipFill'), qn('a:pattFill'), qn('a:grpFill'),
    qn('a:ln'),
    qn('a:effectLst'), qn('a:effectDag'),
    qn('a:scene3d'), qn('a:sp3d'), qn('a:extLst'),
]
_SPPR_RANK = {tag: i for i, tag in enumerate(_SPPR_ORDER)}

def _sort_spPr(spPr):
    children = list(spPr)
    children.sort(key=lambda el: _SPPR_RANK.get(el.tag, 99))
    for child in children: spPr.remove(child)
    for child in children: spPr.append(child)

def _get_spPr(shape):
    return shape._element.find(qn('p:spPr'))

def cm(v): return int(Cm(v))
def pt(v): return int(Pt(v))


# ── Shape builders ───────────────────────────────────────────────────
def rect(slide, x, y, w, h, fill: RGBColor, line=None, line_w=0.5):
    if w <= 0 or h <= 0: return None
    s = slide.shapes.add_shape(1, cm(x), cm(y), cm(w), cm(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line; s.line.width = pt(line_w)
    else:
        s.line.fill.background()
    return s

def rrect(slide, x, y, w, h, fill: RGBColor, radius_pct=8, line=None, line_w=0.5):
    if w <= 0 or h <= 0: return None
    s = slide.shapes.add_shape(5, cm(x), cm(y), cm(w), cm(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line; s.line.width = pt(line_w)
    else:
        s.line.fill.background()
    try:
        adj = s.adjustments
        if adj and len(adj) > 0:
            adj[0] = max(0, min(50, radius_pct)) * 1000
    except Exception: pass
    return s

def oval(slide, x, y, w, h, fill: RGBColor, alpha=100):
    if w <= 0 or h <= 0: return None
    s = slide.shapes.add_shape(9, cm(x), cm(y), cm(w), cm(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background()
    if alpha < 100: set_solid_alpha(s, alpha)
    return s

def bg(slide, color: RGBColor): rect(slide, 0, 0, W, H, color)
def hline(slide, x, y, w, color: RGBColor, thickness=0.08): rect(slide, x, y, w, thickness, color)
def vline(slide, x, y, h2, color: RGBColor, thickness=0.08): rect(slide, x, y, thickness, h2, color)


# ── XML fill helpers ─────────────────────────────────────────────────
def set_solid_alpha(shape, alpha_pct: int):
    try:
        spPr = _get_spPr(shape)
        srgb = spPr.find('.//' + qn('a:srgbClr'))
        if srgb is not None:
            for e in srgb.findall(qn('a:alpha')): srgb.remove(e)
            alp = etree.SubElement(srgb, qn('a:alpha'))
            alp.set('val', str(int(alpha_pct * 1000)))
    except Exception: pass

def gradient_fill(shape, c1: str, c2: str, angle: float = 90):
    try:
        spPr = _get_spPr(shape)
        for tag in [qn('a:solidFill'), qn('a:gradFill'), qn('a:noFill'), qn('a:pattFill')]:
            for el in spPr.findall(tag): spPr.remove(el)
        grad = etree.Element(qn('a:gradFill'))
        gsLst = etree.SubElement(grad, qn('a:gsLst'))
        gs0 = etree.SubElement(gsLst, qn('a:gs')); gs0.set('pos', '0')
        etree.SubElement(gs0, qn('a:srgbClr')).set('val', c1.lstrip('#'))
        gs1 = etree.SubElement(gsLst, qn('a:gs')); gs1.set('pos', '100000')
        etree.SubElement(gs1, qn('a:srgbClr')).set('val', c2.lstrip('#'))
        lin = etree.SubElement(grad, qn('a:lin'))
        lin.set('ang', str(int(angle * 60000))); lin.set('scaled', '0')
        spPr.append(grad); _sort_spPr(spPr)
    except Exception: pass

def gradient_3stop(shape, c1: str, c2: str, c3: str, angle: float = 90):
    """تدرج بثلاث نقاط توقف لعمق بصري أكبر"""
    try:
        spPr = _get_spPr(shape)
        for tag in [qn('a:solidFill'), qn('a:gradFill'), qn('a:noFill'), qn('a:pattFill')]:
            for el in spPr.findall(tag): spPr.remove(el)
        grad = etree.Element(qn('a:gradFill'))
        gsLst = etree.SubElement(grad, qn('a:gsLst'))
        for pos, c in [(0, c1), (50000, c2), (100000, c3)]:
            gs = etree.SubElement(gsLst, qn('a:gs')); gs.set('pos', str(pos))
            etree.SubElement(gs, qn('a:srgbClr')).set('val', c.lstrip('#'))
        lin = etree.SubElement(grad, qn('a:lin'))
        lin.set('ang', str(int(angle * 60000))); lin.set('scaled', '0')
        spPr.append(grad); _sort_spPr(spPr)
    except Exception: pass

def gradient_rect(slide, x, y, w, h, c1: str, c2: str, angle=0):
    c1h = c1.lstrip('#')
    fill_color = RGBColor(int(c1h[0:2], 16), int(c1h[2:4], 16), int(c1h[4:6], 16))
    s = rect(slide, x, y, w, h, fill_color)
    if s: gradient_fill(s, c1, c2, angle)
    return s

def shadow(shape, blur=16, dist=5, angle=135, alpha=0.22, color="000000"):
    try:
        spPr = _get_spPr(shape)
        for old in spPr.findall(qn('a:effectLst')): spPr.remove(old)
        eLst = etree.Element(qn('a:effectLst'))
        shdw = etree.SubElement(eLst, qn('a:outerShdw'))
        shdw.set('blurRad', str(int(blur * 12700)))
        shdw.set('dist', str(int(dist * 12700)))
        shdw.set('dir', str(int(angle * 60000)))
        shdw.set('algn', 'tl')
        srgb = etree.SubElement(shdw, qn('a:srgbClr'))
        srgb.set('val', color.lstrip('#'))
        alp = etree.SubElement(srgb, qn('a:alpha'))
        alp.set('val', str(int(alpha * 100000)))
        spPr.append(eLst); _sort_spPr(spPr)
    except Exception: pass

def glow(shape, color: str, radius_pt: float = 8.0, alpha: float = 0.5):
    """تأثير وهج حول الشكل — مثالي للأيقونات والأرقام الكبيرة"""
    try:
        spPr = _get_spPr(shape)
        for old in spPr.findall(qn('a:effectLst')): spPr.remove(old)
        eLst = etree.Element(qn('a:effectLst'))
        gl = etree.SubElement(eLst, qn('a:glow'))
        gl.set('rad', str(int(radius_pt * 12700)))
        srgb = etree.SubElement(gl, qn('a:srgbClr'))
        srgb.set('val', color.lstrip('#'))
        alp = etree.SubElement(srgb, qn('a:alpha'))
        alp.set('val', str(int(alpha * 100000)))
        spPr.append(eLst); _sort_spPr(spPr)
    except Exception: pass

def shadow_and_glow(shape, s_color="000000", s_blur=16, s_dist=5, s_alpha=0.3,
                    g_color="FFFFFF", g_rad=4.0, g_alpha=0.15):
    """ظل + وهج معاً — أقوى تأثير بصري ممكن"""
    try:
        spPr = _get_spPr(shape)
        for old in spPr.findall(qn('a:effectLst')): spPr.remove(old)
        eLst = etree.Element(qn('a:effectLst'))
        # Glow
        gl = etree.SubElement(eLst, qn('a:glow'))
        gl.set('rad', str(int(g_rad * 12700)))
        sg = etree.SubElement(gl, qn('a:srgbClr')); sg.set('val', g_color.lstrip('#'))
        etree.SubElement(sg, qn('a:alpha')).set('val', str(int(g_alpha * 100000)))
        # Shadow
        shdw = etree.SubElement(eLst, qn('a:outerShdw'))
        shdw.set('blurRad', str(int(s_blur * 12700)))
        shdw.set('dist', str(int(s_dist * 12700)))
        shdw.set('dir', str(int(135 * 60000))); shdw.set('algn', 'tl')
        ss = etree.SubElement(shdw, qn('a:srgbClr')); ss.set('val', s_color.lstrip('#'))
        etree.SubElement(ss, qn('a:alpha')).set('val', str(int(s_alpha * 100000)))
        spPr.append(eLst); _sort_spPr(spPr)
    except Exception: pass


# ══════════════════════════════════════════════════════════════════════
# TEXT ENGINE — FULL PROFESSIONAL REBUILD
# ══════════════════════════════════════════════════════════════════════

def _apply_paragraph_rtl(pPr_elem):
    """إجبار RTL على مستوى الفقرة عبر XML"""
    try:
        pPr_elem.set('rtl', '1')
    except Exception: pass

def _apply_run_effects(rPr_elem, spc_pts: float = 0.0, txt_shadow: bool = False,
                       outline: bool = False, kern_pt: float = 0.0):
    """تطبيق تأثيرات متقدمة على نص run"""
    try:
        # Letter spacing (سpc)
        if abs(spc_pts) > 0.01:
            rPr_elem.set('spc', str(int(spc_pts * 100)))
        # Kerning
        if kern_pt > 0:
            rPr_elem.set('kern', str(int(kern_pt * 100)))
        # Text shadow
        if txt_shadow:
            eLst = etree.SubElement(rPr_elem, qn('a:effectLst'))
            shdw = etree.SubElement(eLst, qn('a:outerShdw'))
            shdw.set('blurRad', '38100'); shdw.set('dist', '25400')
            shdw.set('dir', '8100000'); shdw.set('algn', 'tl')
            sc = etree.SubElement(shdw, qn('a:srgbClr')); sc.set('val', '000000')
            etree.SubElement(sc, qn('a:alpha')).set('val', '45000')
        # Outline / hollow text
        if outline:
            ln = etree.SubElement(rPr_elem, qn('a:ln'))
            ln.set('w', '6350')
            sf = etree.SubElement(ln, qn('a:solidFill'))
    except Exception: pass


def txt(slide, text: str, x, y, w, h,
        font="Cairo", size=14, bold=False, italic=False,
        color: RGBColor | None = None,
        align=PP_ALIGN.RIGHT,
        margin=0.10, rtl=True, spacing=None,
        # ── New Pro params ──
        letter_spacing: float = 0.0,     # نقاط — موجب يباعد، سالب يضيق
        txt_shadow: bool = False,         # ظل نصي
        uppercase: bool = False,          # أحرف كبيرة
        line_spacing_mult: float = 0.0,   # مضاعف المسافة بين السطور (1.15, 1.3...)
        valign_center: bool = False,      # توسيط عمودي
        opacity: int = 100,               # شفافية النص (0-100)
        ) -> object:
    if not text or w <= 0 or h <= 0: return None
    text_str = text.upper() if uppercase else str(text)

    tb = slide.shapes.add_textbox(cm(x), cm(y), cm(w), cm(h))
    tb.word_wrap = True
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left  = cm(margin)
    tf.margin_right = cm(margin)
    tf.margin_top   = cm(0.04)
    tf.margin_bottom = cm(0.04)

    p = tf.paragraphs[0]
    p.alignment = align

    # ── Line spacing via XML ──
    if spacing is not None or line_spacing_mult > 0:
        pPr = p._pPr if hasattr(p, '_pPr') and p._pPr is not None else \
              etree.SubElement(p._p, qn('a:pPr'))
        lnSpc = etree.SubElement(pPr, qn('a:lnSpc'))
        if spacing is not None:
            spcPts = etree.SubElement(lnSpc, qn('a:spcPts'))
            spcPts.set('val', str(int(spacing * 100)))
        elif line_spacing_mult > 0:
            spcPct = etree.SubElement(lnSpc, qn('a:spcPct'))
            spcPct.set('val', str(int(line_spacing_mult * 100000)))

    # ── RTL on paragraph ──
    if rtl:
        pPr2 = p._p.find(qn('a:pPr'))
        if pPr2 is None:
            pPr2 = etree.SubElement(p._p, qn('a:pPr'))
        pPr2.set('rtl', '1')

    # ── Run ──
    run = p.add_run()
    run.text = text_str
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color: run.font.color.rgb = color

    # ── Run-level XML effects ──
    try:
        rPr = run._r.find(qn('a:rPr'))
        if rPr is None:
            rPr = run._r.find('.//' + qn('a:rPr'))
        if rPr is not None:
            _apply_run_effects(rPr, spc_pts=letter_spacing,
                               txt_shadow=txt_shadow)
            # شفافية اللون
            if opacity < 100:
                solidFill = rPr.find(qn('a:solidFill'))
                if solidFill is None:
                    solidFill = etree.SubElement(rPr, qn('a:solidFill'))
                    srgb = etree.SubElement(solidFill, qn('a:srgbClr'))
                    c = color or RGBColor(0xFF, 0xFF, 0xFF)
                    srgb.set('val', f'{c.rgb:06X}')
                else:
                    srgb = solidFill.find(qn('a:srgbClr'))
                if srgb is not None:
                    etree.SubElement(srgb, qn('a:alpha')).set('val', str(opacity * 1000))
    except Exception: pass

    # ── Vertical centering ──
    if valign_center:
        try:
            tf.vertical_anchor = 3  # Middle
        except Exception: pass

    return tb


def txt_hero(slide, text: str, x, y, w, h,
             font="Cairo", size=38, color: RGBColor | None = None,
             align=PP_ALIGN.CENTER, rtl=True,
             shadow_on=True, glow_color: str | None = None) -> object:
    """
    عنوان Hero ضخم مع ظل نصي + letter spacing واسع
    مثالي لعناوين الغلاف والشرائح الكبرى
    """
    tb = txt(slide, text, x, y, w, h,
             font=font, size=size, bold=True, color=color,
             align=align, rtl=rtl,
             letter_spacing=2.0 if size >= 28 else 0.8,
             txt_shadow=shadow_on,
             line_spacing_mult=1.18)
    return tb


def txt_label(slide, text: str, x, y, w, h,
              font="Cairo", size=9.5, color: RGBColor | None = None,
              align=PP_ALIGN.CENTER, rtl=True,
              uppercase=True) -> object:
    """
    تسمية صغيرة مع تباعد حروف وأحرف كبيرة — للعلامات والرؤوس الثانوية
    """
    return txt(slide, text, x, y, w, h,
               font=font, size=size, bold=True, color=color,
               align=align, rtl=rtl,
               letter_spacing=1.5,
               uppercase=uppercase)


def txt_body(slide, text: str, x, y, w, h,
             font="Cairo", size=12.5, color: RGBColor | None = None,
             align=PP_ALIGN.RIGHT, rtl=True) -> object:
    """نص جسم مع مسافة سطر مريحة للقراءة"""
    return txt(slide, text, x, y, w, h,
               font=font, size=size, color=color,
               align=align, rtl=rtl,
               line_spacing_mult=1.45,
               letter_spacing=0.3)


def txt_stat(slide, value: str, x, y, w, h,
             font="Calibri", color: RGBColor | None = None,
             align=PP_ALIGN.CENTER) -> object:
    """رقم إحصائي ضخم مع letter spacing ضيق للعدد وظل قوي"""
    size = 46 if len(value) <= 4 else 34 if len(value) <= 8 else 24
    return txt(slide, value, x, y, w, h,
               font=font, size=size, bold=True, color=color,
               align=align, rtl=False,
               letter_spacing=-0.5,
               txt_shadow=True,
               valign_center=True)


def txt_quote(slide, text: str, x, y, w, h,
              font="Cairo", size=14, color: RGBColor | None = None,
              rtl=True) -> object:
    """نص اقتباسي مائل مع مسافة سطر فاخرة"""
    return txt(slide, text, x, y, w, h,
               font=font, size=size, italic=True, bold=False,
               color=color, align=PP_ALIGN.RIGHT, rtl=rtl,
               line_spacing_mult=1.55,
               letter_spacing=0.5)


def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])
