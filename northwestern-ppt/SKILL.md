---
name: northwestern-ppt
description: Use when turning a research paper / PDF into a Northwestern-styled PowerPoint presentation. Reads the paper, agrees an outline, then builds a branded .pptx with reference-derived layouts.
---

# Northwestern PPT

Turn a paper (PDF) into a Northwestern-branded `.pptx`. Workflow is **outline-first**:
extract → confirm duration → propose outline → user approves → build → visual self-check.

## Workflow

1. **Extract.** Run `python scripts/extract_pdf.py <paper.pdf> work/` →
   `work/text.json` (page text) and `work/figures/` + `work/figures.json`.
   Read the text to understand the paper. Note which figures map to which points.

2. **Confirm duration.** Ask the user the target talk length in minutes.
   Page count ≈ `round(duration_min * 0.7)` (≈1.2–1.4 min/slide; 10 min→~8, 15→~11, 20→~14).

3. **Plan the outline** following `STYLE.md` (skeleton, one-point-per-slide,
   body-type selection rules). Build the outline JSON (schema below). Present it to
   the user as a readable markdown outline and get approval. Revise as needed.

4. **Build.** Write the approved outline to `work/outline.json`, then
   `python scripts/build_pptx.py work/outline.json <paper>_northwestern.pptx`.

5. **Visual self-check.** `python scripts/render_preview.py <paper>_northwestern.pptx work/preview/`
   then open the PNGs (Read tool). Check overflow, overlap, color, balance.
   Fix the outline or report figure issues, rebuild, re-check.

## Outline JSON schema

`{"meta": {"title","subtitle","source","duration_min","use_eyebrow"}, "slides": [ ... ]}`

Slide types and fields:
- `title`: title, subtitle, source
- `bullets`: eyebrow, title, bullets[]   (≤5 bullets, ≤2 lines each)
- `card_row`: eyebrow, title, cards[{title,desc}]   (2–3 cards)
- `pill_flow`: eyebrow, title, pills[], note?
- `callout`: eyebrow, title, label, text
- `mapping`: eyebrow, title, left_header, right_header, rows[[l,r]]
- `figure`: eyebrow, title, image (path from figures.json), caption?, cards?[{title,desc}]
- `thanks`: title (default "Thanks!")

## Notes
- The deck's visual identity (white bg, purple footer, "Northwestern" wordmark, fonts,
  palette) comes entirely from `assets/northwestern_template.pptx` and `scripts/nwstyle.py`.
  Never invent colors — choose content and body types only.
- `use_eyebrow` toggles the eyebrow+underline header style for all slides at once
  (default true for academic talks).
- Run `python scripts/make_sample.py` anytime to regenerate the component gallery.
