"""
Export Pipeline — مذكرتي Pro v17.2
Three engines: canva | classic | bold
Each produces a completely different visual identity.
"""
from __future__ import annotations
import io, logging, os, shutil, subprocess, time, zipfile
from dataclasses import dataclass, field
from pptx import Presentation
from pptx.util import Cm
from pptx.oxml.ns import qn
from core.models import PresentationRequest
from core.themes import get_theme

log = logging.getLogger(__name__)
W_CM, H_CM = 33.867, 19.05
MIN_BYTES = 8_000

# ── Slide function maps per engine ────────────────────────────────────
def _load_engine(name: str) -> dict:
    if name == "classic":
        import engine.slides_classic as m
    elif name == "bold":
        import engine.slides_bold as m
    else:
        import engine.slides_canva as m   # default
    return {
        "set_font":        m.set_font,
        "cover":           m.make_cover,
        "intro":           m.make_intro,
        "plan":            m.make_plan,
        "problem":         m.make_problem,
        "objectives":      m.make_objectives,
        "importance":      m.make_importance,
        "methodology":     m.make_methodology,
        "kpi":             m.make_stats,
        "results":         m.make_results,
        "conclusion":      m.make_conclusion,
        "recommendations": m.make_recommendations,
        "future":          m.make_future,
        "references":      m.make_references,
        "thankyou":        m.make_final,
    }


def _detect_font() -> str:
    candidates = ["Cairo", "Amiri", "Tahoma", "Arial Unicode MS", "Calibri"]
    if shutil.which("fc-list"):
        try:
            out = subprocess.run(["fc-list","--format=%{family}\n"],
                                 capture_output=True, text=True, timeout=5).stdout.lower()
            for f in candidates:
                if f.lower() in out:
                    return f
        except Exception:
            pass
    for font in candidates[:3]:
        for d in ["/usr/share/fonts","/usr/local/share/fonts",
                  os.path.expanduser("~/.fonts"),"/tmp/fonts"]:
            if not os.path.isdir(d): continue
            for root,_,files in os.walk(d):
                for f in files:
                    if font.lower() in f.lower() and f.lower().endswith((".ttf",".otf")):
                        return font
    return "Calibri"


def _fix_slide_type(prs):
    try:
        sldSz = prs.element.find(qn('p:sldSz'))
        if sldSz is not None and 'type' in sldSz.attrib:
            del sldSz.attrib['type']
    except Exception:
        pass


@dataclass
class ExportResult:
    success: bool
    data: bytes = b""
    slide_count: int = 0
    font_used: str = ""
    error: str = ""
    stages: list = field(default_factory=list)
    elapsed: float = 0.0


class PPTXExportPipeline:
    def __init__(self):
        self._font = _detect_font()
        log.info(f"Pipeline ready | font={self._font}")

    def build(self, req: PresentationRequest) -> ExportResult:
        t0 = time.monotonic(); stages = []
        try:
            stages.append("validate")
            errors = req.validate()
            if errors:
                return ExportResult(success=False, error=" | ".join(errors), stages=stages)

            stages.append("init_prs")
            prs = Presentation()
            prs.slide_width = Cm(W_CM)
            prs.slide_height = Cm(H_CM)
            _fix_slide_type(prs)

            stages.append("load_theme")
            theme = get_theme(req.theme)

            stages.append(f"load_engine:{req.engine}")
            eng = _load_engine(req.engine)
            eng["set_font"](self._font)

            stages.append("build_slides")
            n = self._build_slides(prs, req, theme, eng, stages)

            stages.append("serialize")
            data = self._serialize(prs)

            stages.append("validate_output")
            self._validate(data, n)

            elapsed = time.monotonic() - t0
            log.info(f"OK engine={req.engine} slides={n} theme={req.theme} {len(data):,}B {elapsed:.2f}s")
            return ExportResult(success=True, data=data, slide_count=n,
                                font_used=self._font, stages=stages, elapsed=elapsed)
        except Exception as exc:
            stage = stages[-1] if stages else "unknown"
            log.error(f"FAIL [{stage}]: {exc}", exc_info=True)
            return ExportResult(success=False, error=f"[{stage}] {exc}",
                                stages=stages, elapsed=time.monotonic() - t0)

    def _build_slides(self, prs, req, theme, eng, stages) -> int:
        cfg = req.slides; count = 0
        def run(key, condition):
            nonlocal count
            if not condition: return
            stages.append(f"slide:{key}")
            eng[key](prs, req, theme)
            count += 1
        run("cover",           True)
        run("intro",           cfg.intro and bool(req.intro_overview or req.intro_approach))
        run("plan",            cfg.plan and bool(req.chapters))
        run("problem",         cfg.problem and bool(req.main_problem or req.main_question or req.sub_questions))
        run("objectives",      cfg.objectives and bool(req.objectives or req.hypotheses))
        run("importance",      cfg.importance and bool(req.importance or req.reasons))
        run("methodology",     cfg.methodology and bool(req.methodology or req.sample_type or req.tool))
        run("kpi",             cfg.kpi and bool(req.stats))
        run("results",         cfg.results and bool(req.main_results))
        run("conclusion",      cfg.conclusion and bool(req.general_conclusion))
        run("recommendations", cfg.recommendations and bool(req.recommendations))
        run("future",          cfg.future and bool(req.future_work))
        run("references",      cfg.references and bool(req.references))
        run("thankyou",        cfg.thankyou)
        return count

    def _serialize(self, prs):
        buf = io.BytesIO(); prs.save(buf); buf.seek(0); return buf.read()

    def _validate(self, data, n):
        if len(data) < MIN_BYTES: raise ValueError(f"Too small: {len(data)}B")
        if not data.startswith(b'PK'): raise ValueError("Not valid PPTX")
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            if 'ppt/presentation.xml' not in z.namelist():
                raise ValueError("Missing presentation.xml")
            slides = [x for x in z.namelist() if x.startswith('ppt/slides/slide') and x.endswith('.xml')]
            if len(slides) < n: raise ValueError(f"Expected {n} slides, got {len(slides)}")
            for s in slides:
                if len(z.read(s)) < 200: raise ValueError(f"Slide too small: {s}")


_pipeline: PPTXExportPipeline | None = None
def get_pipeline():
    global _pipeline
    if _pipeline is None: _pipeline = PPTXExportPipeline()
    return _pipeline
