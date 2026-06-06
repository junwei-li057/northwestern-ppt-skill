# Northwestern PPT Style Guide

The agent reads this when planning an outline. The build scripts own the exact
colors/sizes (see `scripts/nwstyle.py`); this file owns *content* decisions.

## Palette (reference only — do not hand-set colors)
- Purple `#542A84`, dark purple `#3A1C5C`, ink `#262626`.
- Card accent rotation: blue `#4574B8`, green `#3D916F`, gold `#CE973E`,
  coral `#CF4F61`, purple `#542A84` (paired with light tints).
- Font: Times New Roman throughout.

## Two layers
- **Stable frame** (every slide): title; slide number; branded canvas; optional
  eyebrow + purple underline (on/off as a set via `use_eyebrow`).
- **Body menu** (pick per slide to create variety): bullets, card_row, pill_flow,
  callout, mapping, figure.

## Length = duration-driven
- Ask target duration first. Pages ≈ `round(minutes * 0.7)`.
- Skeleton, stretched/compressed to fit: title → background/problem →
  method (2–4) → key results (2–3) → discussion/limits → takeaways → thanks.

## Pagination rules
- One point per slide.
- Titles: concise and apt — statement, phrase, or question, whatever fits.
  Do NOT mechanically copy the paper's section names.
- Figures/tables → their own `figure` slide + one-line read, not stuffed into bullets.

## Body-type selection (to avoid sameness)
- Default `bullets`, but **never more than 2 bullets slides in a row** — switch to a
  structured type by the 3rd.
- 2–3 parallel concepts/metrics → `card_row`.
- Ordered steps/pipeline → `pill_flow`.
- Core claim / one-line takeaway → `callout`.
- Left-right correspondence / number comparison → `mapping`.
- A paper figure to show → `figure`.
