# مذكرتي Pro v20 — Adaptive Layout Engine

## Architecture

```
mathkarati-v20/
├── app.py                 # Flask API
├── core/
│   ├── models.py          # Domain models + validation
│   └── themes.py          # 12 color themes
├── engine/
│   ├── primitives.py      # Drawing primitives (v20: valign_center on txt_body)
│   ├── slides.py          # CANVA engine (v20: adaptive layout)
│   ├── slides_classic.py  # CLASSIC engine (v20: full-fill adaptive)
│   ├── slides_premium.py  # PREMIUM engine (v20: adaptive layout)
│   └── pipeline.py        # Export orchestrator
├── public/
│   └── index.html
└── requirements.txt
```

## What changed from v19 → v20

| Issue | v19 | v20 |
|-------|-----|-----|
| 3 نتائج/توصيات فقط | فراغ 40% في الأسفل | الصفوف تملأ الشاشة كاملاً |
| قلة العناصر (أهمية، مستقبل) | بطاقات صغيرة + فراغ | بطاقات تملأ كل المساحة |
| النصوص في الأعلى | padding ثابت | توسيط ذكي |
| أرقام "01" تنقسم "0 1" | textbox ضيق | textbox واسع |
| الخاتمة: نص صغير | حجم ثابت 14.5pt | تكيفي حسب طول النص |
| الأهمية: cols ثابتة | cols=3 دائماً | cols تكيفية (1-3) |

## Adaptive Layout System

كل شريحة تحسب الحجم ديناميكياً:
- **≤ 3 عناصر**: بطاقات/صفوف تملأ الشاشة كاملاً، خط كبير
- **4-5 عناصر**: حجم متوسط
- **6+ عناصر**: التصميم الكثيف الأصلي

## Deploy

```bash
pip install -r requirements.txt
gunicorn app:app --bind 0.0.0.0:5000 --workers 2 --timeout 120
```
