"""
Domain Models — مذكرتي Pro v17.1
from_dict() maps ALL field names the frontend actually sends.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class SlideConfig:
    cover: bool = True
    intro: bool = True
    plan: bool = True
    problem: bool = True
    objectives: bool = True
    importance: bool = True
    methodology: bool = True
    kpi: bool = True
    results: bool = True
    conclusion: bool = True
    recommendations: bool = True
    future: bool = True
    references: bool = True
    thankyou: bool = True

    @classmethod
    def from_dict(cls, d: dict) -> "SlideConfig":
        if not d:
            return cls()
        fields = cls.__dataclass_fields__
        return cls(**{k: bool(v) for k, v in d.items() if k in fields})


@dataclass
class StatCard:
    label: str
    value: str
    unit: str = ""

    @classmethod
    def from_dict(cls, d: dict) -> Optional["StatCard"]:
        if not d or not d.get("label") or not d.get("value"):
            return None
        return cls(label=str(d["label"]), value=str(d["value"]), unit=str(d.get("unit", "")))


@dataclass
class Chapter:
    title: str
    pages: str = ""

    @classmethod
    def from_dict(cls, d: dict) -> Optional["Chapter"]:
        if not d or not d.get("title"):
            return None
        return cls(title=str(d["title"]), pages=str(d.get("pages", "")))


@dataclass
class PresentationRequest:
    # Required
    student_name: str
    title_ar: str

    # Metadata
    title_en: str = ""
    supervisor: str = ""
    co_supervisor: str = ""
    institution: str = ""
    year: str = ""
    specialization: str = ""
    lang: str = "ar"
    engine: str = "canva"
    theme: str = "navy_gold"

    # Content
    intro_overview: str = ""
    intro_approach: str = ""
    main_problem: str = ""
    main_question: str = ""
    sub_questions: list = field(default_factory=list)
    objectives: list = field(default_factory=list)
    hypotheses: list = field(default_factory=list)
    importance: list = field(default_factory=list)
    reasons: str = ""
    methodology: str = ""
    sample_type: str = ""
    sample_size: str = ""
    tool: str = ""
    stats: list = field(default_factory=list)
    main_results: list = field(default_factory=list)
    general_conclusion: str = ""
    recommendations: list = field(default_factory=list)
    future_work: list = field(default_factory=list)
    references: list = field(default_factory=list)
    chapters: list = field(default_factory=list)

    slides: SlideConfig = field(default_factory=SlideConfig)

    VALID_THEMES = {
        'navy_gold', 'dark_teal', 'burgundy', 'forest',
        'midnight_purple', 'charcoal_orange', 'ice_blue',
        'sand_gold', 'slate_crimson', 'noir', 'atlas', 'sakura'
    }

    @classmethod
    def from_dict(cls, raw: dict) -> "PresentationRequest":
        def s(key, *aliases):
            """Get string from raw, trying key then aliases."""
            for k in (key,) + aliases:
                v = raw.get(k)
                if v is not None:
                    return str(v).strip()
            return ""

        def lst(key, *aliases):
            """Get list from raw, trying key then aliases."""
            for k in (key,) + aliases:
                v = raw.get(k)
                if v is not None and isinstance(v, list):
                    return [str(x).strip() for x in v if str(x).strip()]
            return []

        # Theme: check 'theme' field, also peek inside palette radio if needed
        theme = s("theme")
        if theme not in cls.VALID_THEMES:
            theme = "navy_gold"

        stats = [c for c in (StatCard.from_dict(x) for x in (raw.get("stats") or [])) if c]
        chapters = [c for c in (Chapter.from_dict(x) for x in (raw.get("chapters") or [])) if c]
        slides = SlideConfig.from_dict(raw.get("slides") or {})

        return cls(
            # frontend sends: studentName
            student_name=s("studentName"),
            # frontend sends: titleAr
            title_ar=s("titleAr"),
            # frontend sends: fieldEn or titleEn
            title_en=s("titleEn", "fieldEn", "titleFr"),
            supervisor=s("supervisor"),
            co_supervisor=s("coSupervisor"),
            # frontend sends: university (not institution)
            institution=s("institution", "university", "faculty"),
            year=s("year"),
            # frontend sends: major or department
            specialization=s("specialization", "major", "department"),
            lang=s("lang") or "ar",
            engine=s("engine") or "canva",
            theme=theme,
            intro_overview=s("introOverview"),
            intro_approach=s("introApproach"),
            main_problem=s("mainProblem"),
            main_question=s("mainQuestion"),
            sub_questions=lst("subQuestions"),
            objectives=lst("objectives"),
            hypotheses=lst("hypotheses"),
            importance=lst("importance"),
            reasons=s("reasons", "dataSource"),
            methodology=s("methodology"),
            sample_type=s("sampleType"),
            sample_size=s("sampleSize"),
            tool=s("tool"),
            stats=stats,
            # frontend sends: mainResults
            main_results=lst("mainResults"),
            # frontend sends: generalConclusion
            general_conclusion=s("generalConclusion", "conclusion"),
            recommendations=lst("recommendations"),
            future_work=lst("futureWork"),
            references=lst("references"),
            chapters=chapters,
            slides=slides,
        )

    def validate(self) -> list[str]:
        errors = []
        if not self.student_name:
            errors.append("اسم الطالب مطلوب")
        if not self.title_ar:
            errors.append("عنوان المذكرة مطلوب")
        return errors
