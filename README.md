# Northwestern PPT Skill

Turn a research paper (PDF) into a clean, **Northwestern-branded** PowerPoint deck.

This is a [Claude Code](https://claude.com/claude-code) skill. You point it at a paper,
agree on an outline, and it builds a `.pptx` whose visual identity (white background,
purple footer bar, *Northwestern* wordmark, Times New Roman, purple palette) comes from
an official Northwestern template — and whose page layouts are a small set of components
distilled from real, well-designed academic decks. The goal is a deck that looks
hand-built, not "AI-template."

## What it produces

A deck assembled from a **stable frame** (title, slide number, branded canvas, optional
eyebrow + purple underline) plus a **body menu** you pick from per slide to keep pages
varied and full:

| Body type    | Use it for                                            |
|--------------|-------------------------------------------------------|
| `bullets`    | ordinary points / lists                               |
| `card_row`   | 2–3 parallel concepts or metrics (rotating accents)   |
| `pill_flow`  | an ordered pipeline / sequence of steps               |
| `callout`    | a single core claim or one-line takeaway              |
| `mapping`    | left→right correspondence (e.g. behavior → cause)     |
| `table`      | tabular / numeric data (a real PowerPoint table)      |
| `figure`     | a paper figure, auto-extracted and white-trimmed      |
| `title` / `thanks` | cover and closing slides                        |

## How it works

```
paper.pdf ──► extract_pdf.py ──► text.json + figures/ + figures.json
                                      │
              (read the paper, confirm talk length, plan an outline)
                                      │
work/outline.json ──► build_pptx.py ──► <paper>_northwestern.pptx
                                      │
                          render_preview.py ──► per-slide PNGs (visual self-check)
```

- **Outline-first.** The skill extracts the paper, asks for the target talk length
  (pages ≈ `round(minutes × 0.7)`), proposes an outline, and only builds after you approve.
- **Style lives in one place.** All colors/sizes are in [`scripts/nwstyle.py`](northwestern-ppt/scripts/nwstyle.py)
  and the official template; the outline only chooses content and body types.
- **Visual self-check.** Every build is rasterized to PNGs so layout problems
  (overflow, overlap, balance) are caught by eye before you ship.

## Requirements

- Python 3 with `python-pptx`, `PyMuPDF`, `Pillow`, `pytest` (`pip install -r northwestern-ppt/requirements.txt`)
- **LibreOffice** (`soffice`) — only for the PNG preview step

## Usage

```bash
cd northwestern-ppt

# 1. extract text + figures from the paper
python scripts/extract_pdf.py path/to/paper.pdf work/

# 2. write work/outline.json following SKILL.md / STYLE.md, then build
python scripts/build_pptx.py work/outline.json work/paper_northwestern.pptx

# 3. render previews and eyeball them
python scripts/render_preview.py work/paper_northwestern.pptx work/preview/

# see every component on one deck
python scripts/make_sample.py
```

The agent-facing instructions live in [`northwestern-ppt/SKILL.md`](northwestern-ppt/SKILL.md)
(workflow + outline schema) and [`northwestern-ppt/STYLE.md`](northwestern-ppt/STYLE.md)
(pagination and body-type selection rules).

## Project structure

```
northwestern-ppt/
├── SKILL.md                # workflow + outline JSON schema
├── STYLE.md                # palette, pagination & component-selection rules
├── assets/
│   └── northwestern_template.pptx   # official branded canvas (white + purple)
├── scripts/
│   ├── nwstyle.py          # central palette/fonts/geometry + shape helpers
│   ├── components.py        # one renderer per body type
│   ├── build_pptx.py        # outline JSON + template → .pptx (+ CLI)
│   ├── extract_pdf.py       # PDF → text + figures + figures.json
│   ├── render_preview.py    # .pptx → per-slide PNGs (LibreOffice)
│   └── make_sample.py       # component gallery
└── tests/                  # pytest suite (21 tests)
```

`docs/superpowers/` contains the design spec and the implementation plan the skill was
built from.

## Notes

- The deck's look comes entirely from the official template + `nwstyle.py`; the outline
  never sets colors.
- The reference decks used to derive the component designs are **not** included in this
  repository.
