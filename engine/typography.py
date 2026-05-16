"""
Typography System — مذكرتي Pro v17.2 (Enhanced)
نظام هرمية الخطوط — Type Scale للعروض الأكاديمية

DESIGN PHILOSOPHY (من منظور مشاهد):
─────────────────────────────────────
عندما ينظر الجمهور للشريحة يجب أن:
  1. يُدرك العنوان الرئيسي فوراً (أكبر + أثقل)
  2. يميّز عناوين الأقسام بسهولة (هوّة بصرية واضحة)
  3. يقرأ النص العادي بارتياح (لا صغير جداً)
  4. يلاحظ الأرقام الإحصائية بوضوح مختلف
  5. يرى الملاحظات كعنصر ثانوي مميّز

HIERARCHY RULES:
  • H1 (24pt) ≥ ضعف حجم BODY (12pt) دائماً
  • فارق أدنى 4pt بين كل مستوى
  • Bold للعناوين فقط (H1, H2, H3)
  • أرقام الإحصاء: Calibri لوضوح اللاتيني

SCALE (pt):
  DISPLAY     → 26-40pt  — عنوان الغلاف
  H1          → 24pt     — عناوين الشرائح
  H2          → 17pt     — عناوين الأقسام
  H3          → 13pt     — عناوين الفقرات
  BODY_LG     → 13pt     — نص رئيسي بارز
  BODY        → 12pt     — نص عادي
  BODY_SM     → 10pt     — نص مضغوط
  SMALL       → 9pt      — تسميات ثانوية
  STAT        → 22-48pt  — أرقام إحصائية
"""

# ── Font families ──────────────────────────────────────────────────────
FONT_TITLE   = "Cairo"
FONT_HEADING = "Cairo"
FONT_BODY    = "Cairo"
FONT_NUM     = "Calibri"
FONT_EN      = "Calibri"


# ── Type Scale (pt) ────────────────────────────────────────────────────
class TS:
    """
    Type Scale — الثوابت الوحيدة المسموح باستخدامها للأحجام.

    VISUAL HIERARCHY (من الأكبر للأصغر):
      DISPLAY_LG(40) → DISPLAY_MD(32) → DISPLAY_SM(26)
        → H1(24) → H2(17) → H3(13) → BODY_LG(13) → BODY(12) → BODY_SM(10) → LABEL(9) → CAPTION(8)
    """

    # ── Display — الغلاف فقط
    DISPLAY_LG  = 40   # عنوان قصير < 40 حرف
    DISPLAY_MD  = 32   # عنوان متوسط 40-70 حرف
    DISPLAY_SM  = 26   # عنوان طويل > 70 حرف

    # ── Headings — عناوين الشرائح
    H1          = 24   # عنوان الشريحة — واضح من مسافة بعيدة
    H2          = 17   # عنوان القسم / البطاقة
    H3          = 13   # عنوان الفقرة الفرعية

    # ── Body — نص شارح
    BODY_LG     = 13   # نص رئيسي (إشكالية، ملخص)
    BODY        = 12   # نص عادي (قوائم، شروح)
    BODY_SM     = 10   # نص مضغوط (جداول كثيفة)

    # ── Supporting
    LABEL       = 9    # تسميات الصفوف / الأعمدة
    CAPTION     = 8    # ملاحظات صغيرة، مصادر

    # ── Statistics — أرقام إحصائية بارزة
    STAT_XL     = 48
    STAT_LG     = 36
    STAT_MD     = 28
    STAT_SM     = 22

    # ── Sidebar
    SIDEBAR_LABEL = 14  # أوضح: 13→14
    SIDEBAR_ICON  = 38

    # ── Cover info rows
    COVER_LABEL   = 10  # تسمية (الطالب:، المشرف:)
    COVER_VALUE   = 14  # قيمة الاسم — أوضح: 13→14
    COVER_INST    = 11  # اسم المؤسسة
    COVER_YEAR    = 15  # السنة في badge — أوضح: 14→15

    # ── Slide index
    SLIDE_NUM     = 9
    WATERMARK     = 8


# ── Dynamic title size ────────────────────────────────────────────────
def display_size(text: str) -> int:
    """حجم عنوان الغلاف حسب طول النص"""
    n = len(text)
    if n < 40:
        return TS.DISPLAY_LG
    elif n < 70:
        return TS.DISPLAY_MD
    else:
        return TS.DISPLAY_SM


def h1_size(text: str) -> int:
    """
    حجم عنوان الشريحة حسب طول النص.
    الحد الأدنى 18pt (مقروء من مسافة 3 أمتار).
    """
    n = len(text)
    if n < 30:
        return TS.H1   # 24pt
    elif n < 55:
        return 21
    else:
        return 18      # أدنى حد مسموح


def stat_size(value: str) -> int:
    """حجم الرقم الإحصائي حسب طوله"""
    n = len(str(value))
    if n <= 4:
        return TS.STAT_LG
    elif n <= 6:
        return TS.STAT_MD
    else:
        return TS.STAT_SM


# ── Typography helpers ────────────────────────────────────────────────
def get_title_font(arabic_font: str = "Cairo") -> str:
    return "Cairo"


def get_body_font(arabic_font: str = "Cairo") -> str:
    return arabic_font


# ── Spacing ratios ────────────────────────────────────────────────────
class Heights:
    SLIDE_TITLE_ROW = 1.1
    SECTION_HEADER  = 0.80
    BODY_LINE       = 0.55
    STAT_ROW        = 2.0
    LABEL_ROW       = 0.45


# ── Line spacing (pt) ─────────────────────────────────────────────────
class LineSpacing:
    """تباعد الأسطر الموصى به — مرر إلى spacing= في txt()"""
    DISPLAY  = 52    # عنوان الغلاف
    H1       = 36    # عنوان الشريحة
    H2       = 26    # عنوان القسم
    BODY     = 20    # نص عادي
    BODY_SM  = 17    # نص مضغوط
    CAPTION  = 14    # تسميات
