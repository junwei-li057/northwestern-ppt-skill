"""Generate a sample deck exercising every component, then render previews.
Use to eyeball the visual system without a real paper."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import build_pptx
import render_preview

SAMPLE = {
    "meta": {"use_eyebrow": True},
    "slides": [
        {"type": "title", "title": "Northwestern Style Sample",
         "subtitle": "Every component on one deck", "source": "make_sample.py"},
        {"type": "bullets", "eyebrow": "Introduction", "title": "Why this deck",
         "bullets": ["Shows each body type", "Confirms palette and fonts",
                     "Used for visual regression by eye"]},
        {"type": "card_row", "eyebrow": "Metrics", "title": "Three process scores",
         "cards": [{"title": "Milestone KPI", "desc": "Tracks intermediate progress."},
                   {"title": "Communication", "desc": "Rates clarity and alignment."},
                   {"title": "Planning", "desc": "Rates strategy and adaptation."}]},
        {"type": "pill_flow", "eyebrow": "Pipeline", "title": "Practical pipeline",
         "pills": ["Ingest", "Bin", "Sample", "Score", "Report"],
         "note": "Each stage feeds the next, left to right."},
        {"type": "callout", "eyebrow": "Claim", "title": "The core point",
         "label": "Core claim",
         "text": "If interaction changes behavior, evaluation must observe the process."},
        {"type": "mapping", "eyebrow": "Diagnosis", "title": "Behavior to failure mode",
         "left_header": "Observed behavior", "right_header": "Attribution",
         "rows": [["Seer's claim not discussed", "Ignored other agent's input"],
                  ["No one asks for evidence", "Fail to ask for clarification"],
                  ["Vote without cross-checking", "No / incomplete verification"]]},
        {"type": "thanks"},
    ],
}


def main():
    out_pptx = ROOT / "tests" / "_out" / "sample.pptx"
    out_pptx.parent.mkdir(parents=True, exist_ok=True)
    template = ROOT / "assets" / "northwestern_template.pptx"
    build_pptx.build(SAMPLE, str(template), str(out_pptx))
    pngs = render_preview.render(str(out_pptx), str(out_pptx.parent / "sample_png"))
    print(f"sample deck: {out_pptx}\nslides rendered: {len(pngs)}")


if __name__ == "__main__":
    main()
