# Northwestern PPT Skill Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Claude Code skill that turns a research paper (PDF) into a Northwestern-branded `.pptx` deck, using an official Northwestern template as the branded canvas and a set of reference-derived components for the body.

**Architecture:** A central style module (`nwstyle.py`) holds the palette/fonts/geometry and low-level shape helpers. `components.py` draws one slide per body-type from a `page` dict. `build_pptx.py` loads the official template, strips its sample slides, and dispatches each outline page to the matching component. `extract_pdf.py` pulls text + figures from the PDF; `render_preview.py` rasterizes the finished deck to PNGs for visual self-check. `SKILL.md`/`STYLE.md` are the agent-facing instructions (workflow, duration→page-count, outline schema, pagination/selection rules).

**Tech Stack:** Python 3, `python-pptx`, `PyMuPDF` (fitz), LibreOffice (headless, for PNG preview), `pytest`.

---

## Conventions locked for all tasks

These names/values are referenced by later tasks. Do not rename.

- Working dir for all commands: `/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent`
- Skill root: `northwestern-ppt/`
- Template (already in repo): `northwestern-ppt/assets/northwestern_template.pptx`
- Slide size: 10.0 × 5.625 in (16:9). `python-pptx` reads it from the template; never hardcode beyond bounds checks.
- **Layout indices in the template** (verified): `SEPARATOR_LAYOUT = 0` (Separator Page — left purple block bg, for title + section dividers), `CONTENT_LAYOUT = 1` (Blank — inherits master purple footer bar + "Northwestern" wordmark, for all body pages).
- Font: `FONT = "Times New Roman"` everywhere.
- Palette (hex, no `#`):
  - `PURPLE = "542A84"` (Northwestern purple; eyebrow + underline + title-page text)
  - `PURPLE_DARK = "3A1C5C"` (slide titles)
  - `INK = "262626"` (body text)
  - `WHITE = "FFFFFF"`
  - `ACCENTS` = ordered list of `(line_hex, fill_hex)` for card rotation:
    `[("4574B8","EFF4FC"), ("3D916F","EEF8F3"), ("CE973E","FCF7ED"), ("CF4F61","FBEFF1"), ("542A84","F6F2FA")]`
- Header band geometry (inches): eyebrow `(0.55, 0.28, 8.9, 0.22)` 10pt caps PURPLE; title `(0.55, 0.50, 9.0, 0.6)` 28pt PURPLE_DARK; underline rect `(0.55, 1.12, 1.15, 0.045)` filled PURPLE. Body region starts at `BODY_TOP = 1.45` and ends at `BODY_BOTTOM = 5.05`.
- Outline JSON schema (one object):
  ```json
  {
    "meta": {"title":"...", "subtitle":"...", "source":"...", "duration_min":15, "use_eyebrow":true},
    "slides": [ {"type":"title"|"bullets"|"card_row"|"pill_flow"|"callout"|"mapping"|"figure"|"thanks", ...} ]
  }
  ```
  Per-type fields:
  - `title`: `title`, `subtitle`, `source`
  - `thanks`: `title` (default "Thanks!")
  - `bullets`: `eyebrow`, `title`, `bullets` (list[str])
  - `card_row`: `eyebrow`, `title`, `cards` (list of `{"title","desc"}`, 2–3 items)
  - `pill_flow`: `eyebrow`, `title`, `pills` (list[str]), `note` (optional str)
  - `callout`: `eyebrow`, `title`, `label`, `text`
  - `mapping`: `eyebrow`, `title`, `left_header`, `right_header`, `rows` (list of `[left, right]`)
  - `figure`: `eyebrow`, `title`, `image` (path), `caption` (optional), `cards` (optional list of `{"title","desc"}`)
- Component function contract: `render_<type>(prs, page) -> slide`, where `prs` is a `pptx.Presentation` and `page` is one slide dict. Each adds exactly one slide and returns it.
- Tests live in `northwestern-ppt/tests/`. Run from skill root with `cd northwestern-ppt && python -m pytest`.

---

## File Structure

```
northwestern-ppt/
├── SKILL.md                # agent-facing: trigger, workflow, duration→pages, outline schema
├── STYLE.md                # agent-facing: palette, fonts, component catalog, pagination/selection rules
├── assets/
│   └── northwestern_template.pptx   # exists
├── scripts/
│   ├── nwstyle.py          # constants + low-level shape/text helpers (one job: style)
│   ├── components.py       # render_<type>(prs, page) for each body type
│   ├── build_pptx.py       # load template, strip slides, dispatch outline → pptx, CLI
│   ├── extract_pdf.py      # PDF → text dump + figures/*.png + figures.json, CLI
│   ├── render_preview.py   # pptx → preview/*.png via LibreOffice, CLI
│   └── make_sample.py      # built-in outline covering every type → sample deck + preview
├── references/             # exists (read-only style source)
└── tests/
    ├── conftest.py
    ├── test_nwstyle.py
    ├── test_components.py
    ├── test_build.py
    └── test_extract.py
```

---

## Task 0: Scaffolding, dependencies, test harness

**Files:**
- Create: `northwestern-ppt/scripts/__init__.py` (empty)
- Create: `northwestern-ppt/tests/__init__.py` (empty)
- Create: `northwestern-ppt/tests/conftest.py`
- Create: `northwestern-ppt/requirements.txt`

- [ ] **Step 1: Initialize git (repo is not yet under version control)**

Run:
```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git init -q && printf "__pycache__/\n*.pyc\nwork/\n.superpowers/\n*_northwestern.pptx\nnorthwestern-ppt/tests/_out/\n" > .gitignore
git add -A && git commit -q -m "chore: init repo with spec, plan, template, references" && echo OK
```
Expected: `OK`

- [ ] **Step 2: Install dependencies**

Run:
```bash
python3 -m pip install --quiet python-pptx PyMuPDF pytest && python3 -c "import pptx, fitz, pytest; print('deps ok')"
```
Expected: `deps ok`

- [ ] **Step 3: Install LibreOffice (for PNG preview); skip if already present**

Run:
```bash
( command -v soffice || ls /Applications/LibreOffice.app/Contents/MacOS/soffice ) 2>/dev/null || brew install --cask libreoffice
ls /Applications/LibreOffice.app/Contents/MacOS/soffice && echo "soffice ok"
```
Expected: path printed then `soffice ok`. (If `brew` is unavailable, note it; `render_preview.py` will fail gracefully and the rest of the skill still works.)

- [ ] **Step 4: Create package markers and requirements**

Create `northwestern-ppt/scripts/__init__.py` and `northwestern-ppt/tests/__init__.py` as empty files.

Create `northwestern-ppt/requirements.txt`:
```
python-pptx>=1.0
PyMuPDF>=1.24
pytest>=8.0
```

- [ ] **Step 5: Create conftest with a shared template-path fixture**

Create `northwestern-ppt/tests/conftest.py`:
```python
import sys
from pathlib import Path
import pytest

ROOT = Path(__file__).resolve().parents[1]          # northwestern-ppt/
sys.path.insert(0, str(ROOT / "scripts"))           # import nwstyle/components/...

@pytest.fixture
def template_path():
    p = ROOT / "assets" / "northwestern_template.pptx"
    assert p.exists(), f"template missing: {p}"
    return str(p)

@pytest.fixture
def out_dir(tmp_path):
    return tmp_path
```

- [ ] **Step 6: Verify pytest collects with zero tests**

Run: `cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent/northwestern-ppt" && python -m pytest -q`
Expected: `no tests ran` (exit code 5 is fine).

- [ ] **Step 7: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "chore: scaffold scripts/tests, deps, conftest" && echo OK
```
Expected: `OK`

---

## Task 1: `nwstyle.py` — constants + shape/text helpers

**Files:**
- Create: `northwestern-ppt/scripts/nwstyle.py`
- Test: `northwestern-ppt/tests/test_nwstyle.py`

- [ ] **Step 1: Write the failing test**

Create `northwestern-ppt/tests/test_nwstyle.py`:
```python
from pptx import Presentation
from pptx.util import Inches
import nwstyle


def _blank_slide():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    return prs, prs.slides.add_slide(prs.slide_layouts[6])  # 6 = blank in default template


def test_hex_to_rgb():
    rgb = nwstyle.hex_to_rgb("542A84")
    assert (rgb[0], rgb[1], rgb[2]) == (0x54, 0x2A, 0x84)


def test_add_textbox_sets_text_font_color():
    prs, slide = _blank_slide()
    shape = nwstyle.add_textbox(slide, 0.5, 0.5, 4, 0.5, "Hello", 18, "262626", bold=True)
    run = shape.text_frame.paragraphs[0].runs[0]
    assert run.text == "Hello"
    assert run.font.name == "Times New Roman"
    assert run.font.size.pt == 18
    assert run.font.bold is True
    assert run.font.color.rgb == nwstyle.hex_to_rgb("262626")


def test_add_rounded_rect_fill_and_line():
    prs, slide = _blank_slide()
    shape = nwstyle.add_rounded_rect(slide, 1, 1, 2, 1, fill_hex="EFF4FC", line_hex="4574B8")
    assert shape.fill.fore_color.rgb == nwstyle.hex_to_rgb("EFF4FC")
    assert shape.line.color.rgb == nwstyle.hex_to_rgb("4574B8")


def test_accents_has_five_pairs():
    assert len(nwstyle.ACCENTS) == 5
    for line_hex, fill_hex in nwstyle.ACCENTS:
        assert len(line_hex) == 6 and len(fill_hex) == 6
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd northwestern-ppt && python -m pytest tests/test_nwstyle.py -q`
Expected: FAIL — `ModuleNotFoundError: No module named 'nwstyle'`.

- [ ] **Step 3: Write the implementation**

Create `northwestern-ppt/scripts/nwstyle.py`:
```python
"""Central style constants and low-level pptx shape/text helpers.

Every color/size/geometry value the deck uses lives here so the look stays
stable and editable in one place.
"""
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# --- template layout indices (verified against northwestern_template.pptx) ---
SEPARATOR_LAYOUT = 0   # left purple block bg -> title + section dividers
CONTENT_LAYOUT = 1     # purple footer bar + "Northwestern" wordmark -> body pages

# --- typography ---
FONT = "Times New Roman"

# --- palette ---
PURPLE = "542A84"
PURPLE_DARK = "3A1C5C"
INK = "262626"
WHITE = "FFFFFF"
ACCENTS = [
    ("4574B8", "EFF4FC"),  # blue
    ("3D916F", "EEF8F3"),  # green
    ("CE973E", "FCF7ED"),  # gold
    ("CF4F61", "FBEFF1"),  # coral
    ("542A84", "F6F2FA"),  # purple
]

# --- header band + body region geometry (inches) ---
EYEBROW_BOX = (0.55, 0.28, 8.9, 0.22)
TITLE_BOX = (0.55, 0.50, 9.0, 0.6)
UNDERLINE_BOX = (0.55, 1.12, 1.15, 0.045)
BODY_TOP = 1.45
BODY_BOTTOM = 5.05
MARGIN_L = 0.55
CONTENT_W = 8.9


def hex_to_rgb(h):
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def style_run(run, size_pt, color_hex, bold=False, font=FONT):
    run.font.name = font
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = hex_to_rgb(color_hex)


def add_textbox(slide, left, top, width, height, text, size_pt, color_hex,
                bold=False, align=PP_ALIGN.LEFT, font=FONT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    style_run(run, size_pt, color_hex, bold=bold, font=font)
    return box


def add_rounded_rect(slide, left, top, width, height, fill_hex=None,
                     line_hex=None, line_w_pt=1.0):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top),
        Inches(width), Inches(height))
    shape.shadow.inherit = False
    if fill_hex is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    if line_hex is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = hex_to_rgb(line_hex)
        shape.line.width = Pt(line_w_pt)
    return shape


def add_rect(slide, left, top, width, height, fill_hex):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.shadow.inherit = False
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    shape.line.fill.background()
    return shape


def add_connector(slide, x1, y1, x2, y2, color_hex, w_pt=1.25):
    from pptx.enum.shapes import MSO_CONNECTOR
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                      Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    conn.line.color.rgb = hex_to_rgb(color_hex)
    conn.line.width = Pt(w_pt)
    return conn
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd northwestern-ppt && python -m pytest tests/test_nwstyle.py -q`
Expected: PASS (4 passed).

- [ ] **Step 5: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "feat: nwstyle constants and shape/text helpers" && echo OK
```
Expected: `OK`

---

## Task 2: `build_pptx.py` core — load template, strip slides cleanly

**Files:**
- Create: `northwestern-ppt/scripts/build_pptx.py`
- Test: `northwestern-ppt/tests/test_build.py`

- [ ] **Step 1: Write the failing test**

Create `northwestern-ppt/tests/test_build.py`:
```python
import warnings
from pptx import Presentation
import build_pptx


def test_load_template_has_three_sample_slides(template_path):
    prs = build_pptx.load_template(template_path)
    assert len(prs.slides._sldIdLst) == 3


def test_clear_slides_removes_all_without_duplicate_warning(template_path):
    prs = build_pptx.load_template(template_path)
    build_pptx.clear_slides(prs)
    assert len(prs.slides._sldIdLst) == 0


def test_rebuild_saves_clean_file(template_path, out_dir):
    prs = build_pptx.load_template(template_path)
    build_pptx.clear_slides(prs)
    prs.slides.add_slide(prs.slide_layouts[build_pptx.ns.CONTENT_LAYOUT])
    out = out_dir / "clean.pptx"
    with warnings.catch_warnings():
        warnings.simplefilter("error")     # duplicate-name -> error
        prs.save(str(out))
    reopened = Presentation(str(out))
    assert len(reopened.slides._sldIdLst) == 1
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd northwestern-ppt && python -m pytest tests/test_build.py -q`
Expected: FAIL — `ModuleNotFoundError: No module named 'build_pptx'`.

- [ ] **Step 3: Write the minimal implementation**

Create `northwestern-ppt/scripts/build_pptx.py`:
```python
"""Load the Northwestern template, strip its sample slides, and (later) render
an outline into a finished deck."""
from pptx import Presentation
import nwstyle as ns

_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def load_template(path):
    return Presentation(path)


def clear_slides(prs):
    """Remove every slide AND its relationship so re-adding slides does not
    collide with orphaned parts (avoids 'Duplicate name' on save)."""
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        prs.part.drop_rel(sldId.get(_R_ID))
        sldIdLst.remove(sldId)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd northwestern-ppt && python -m pytest tests/test_build.py -q`
Expected: PASS (3 passed).

- [ ] **Step 5: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "feat: build_pptx load_template + clean clear_slides" && echo OK
```
Expected: `OK`

---

## Task 3: `components.py` — header band, title page, thanks page

**Files:**
- Create: `northwestern-ppt/scripts/components.py`
- Test: `northwestern-ppt/tests/test_components.py`

- [ ] **Step 1: Write the failing test**

Create `northwestern-ppt/tests/test_components.py`:
```python
from pptx import Presentation
import build_pptx
import components
import nwstyle as ns


def fresh(template_path):
    prs = build_pptx.load_template(template_path)
    build_pptx.clear_slides(prs)
    return prs


def texts(slide):
    out = []
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            out.append(sh.text_frame.text.strip())
    return out


def test_title_slide(template_path):
    prs = fresh(template_path)
    page = {"type": "title", "title": "My Talk", "subtitle": "A subtitle", "source": "arXiv:1234"}
    slide = components.render_title(prs, page)
    joined = " | ".join(texts(slide))
    assert "My Talk" in joined and "A subtitle" in joined and "arXiv:1234" in joined
    assert slide.slide_layout == prs.slide_layouts[ns.SEPARATOR_LAYOUT]


def test_thanks_slide_default_text(template_path):
    prs = fresh(template_path)
    slide = components.render_thanks(prs, {"type": "thanks"})
    assert "Thanks!" in " | ".join(texts(slide))


def test_header_band_returns_body_top(template_path):
    prs = fresh(template_path)
    slide = prs.slides.add_slide(prs.slide_layouts[ns.CONTENT_LAYOUT])
    body_top = components.header_band(slide, "INTRO", "A title", use_eyebrow=True)
    assert abs(body_top - ns.BODY_TOP) < 1e-6
    assert "INTRO" in " | ".join(texts(slide))
    assert "A title" in " | ".join(texts(slide))
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd northwestern-ppt && python -m pytest tests/test_components.py -q`
Expected: FAIL — `ModuleNotFoundError: No module named 'components'`.

- [ ] **Step 3: Write the implementation**

Create `northwestern-ppt/scripts/components.py`:
```python
"""One render_<type>(prs, page) -> slide per body type. All geometry/colors
come from nwstyle; these functions only place content."""
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import nwstyle as ns


def _content_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[ns.CONTENT_LAYOUT])


def _separator_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[ns.SEPARATOR_LAYOUT])


def header_band(slide, eyebrow, title, use_eyebrow=True):
    """Draw eyebrow + title + purple underline. Returns body_top (inches)."""
    if use_eyebrow and eyebrow:
        l, t, w, h = ns.EYEBROW_BOX
        ns.add_textbox(slide, l, t, w, h, eyebrow.upper(), 10, ns.PURPLE)
    l, t, w, h = ns.TITLE_BOX
    ns.add_textbox(slide, l, t, w, h, title, 28, ns.PURPLE_DARK)
    l, t, w, h = ns.UNDERLINE_BOX
    ns.add_rect(slide, l, t, w, h, ns.PURPLE)
    return ns.BODY_TOP


def render_title(prs, page):
    slide = _separator_slide(prs)
    ns.add_textbox(slide, 2.0, 2.0, 7.4, 1.2, page.get("title", ""), 34, ns.PURPLE_DARK, bold=True)
    if page.get("subtitle"):
        ns.add_textbox(slide, 2.0, 3.2, 7.4, 0.8, page["subtitle"], 18, ns.INK)
    if page.get("source"):
        ns.add_textbox(slide, 2.0, 4.6, 7.4, 0.4, page["source"], 12, ns.PURPLE)
    return slide


def render_thanks(prs, page):
    slide = _separator_slide(prs)
    ns.add_textbox(slide, 2.0, 2.4, 7.4, 1.0, page.get("title", "Thanks!"),
                   40, ns.PURPLE_DARK, bold=True)
    return slide
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd northwestern-ppt && python -m pytest tests/test_components.py -q`
Expected: PASS (3 passed).

- [ ] **Step 5: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "feat: header band, title and thanks components" && echo OK
```
Expected: `OK`

---

## Task 4: `render_bullets`

**Files:**
- Modify: `northwestern-ppt/scripts/components.py` (append function)
- Test: `northwestern-ppt/tests/test_components.py` (append test)

- [ ] **Step 1: Write the failing test** (append to `test_components.py`)

```python
def test_bullets(template_path):
    prs = fresh(template_path)
    page = {"type": "bullets", "eyebrow": "METHOD", "title": "How it works",
            "bullets": ["First point", "Second point", "Third point"]}
    slide = components.render_bullets(prs, page)
    body = [t for t in texts(slide) if "point" in t.lower()]
    assert len(body) == 1                      # one body textbox holding all bullets
    for b in page["bullets"]:
        assert b in body[0]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd northwestern-ppt && python -m pytest tests/test_components.py::test_bullets -q`
Expected: FAIL — `AttributeError: module 'components' has no attribute 'render_bullets'`.

- [ ] **Step 3: Write the implementation** (append to `components.py`)

```python
def render_bullets(prs, page):
    slide = _content_slide(prs)
    top = header_band(slide, page.get("eyebrow"), page.get("title", ""),
                      use_eyebrow=page.get("use_eyebrow", True))
    box = slide.shapes.add_textbox(ns.Inches(ns.MARGIN_L), ns.Inches(top),
                                   ns.Inches(ns.CONTENT_W), ns.Inches(ns.BODY_BOTTOM - top))
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(page.get("bullets", [])):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = ns.Pt(8)
        run = p.add_run()
        run.text = "•  " + item
        ns.style_run(run, 16, ns.INK)
    return slide
```

Note: add `from pptx.util import Inches, Pt` import access via `ns`. In `nwstyle.py`, expose them by adding at the end of Task 1's file (do this now if not present):
```python
Inches = Inches  # re-export for components
Pt = Pt
```
(They are already imported at the top of `nwstyle.py`; this line just makes `ns.Inches`/`ns.Pt` explicit attributes.)

- [ ] **Step 4: Run test to verify it passes**

Run: `cd northwestern-ppt && python -m pytest tests/test_components.py -q`
Expected: PASS (all component tests).

- [ ] **Step 5: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "feat: render_bullets component" && echo OK
```
Expected: `OK`

---

## Task 5: `render_card_row`

**Files:**
- Modify: `northwestern-ppt/scripts/components.py`
- Test: `northwestern-ppt/tests/test_components.py`

- [ ] **Step 1: Write the failing test** (append)

```python
def test_card_row_three_cards_rotating_accents(template_path):
    prs = fresh(template_path)
    page = {"type": "card_row", "eyebrow": "METRICS", "title": "Three scores",
            "cards": [{"title": "A", "desc": "alpha"},
                      {"title": "B", "desc": "beta"},
                      {"title": "C", "desc": "gamma"}]}
    slide = components.render_card_row(prs, page)
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    cards = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE]
    fills = {c.fill.fore_color.rgb for c in cards}
    assert fills == {ns.hex_to_rgb(ns.ACCENTS[i][1]) for i in range(3)}
    joined = " | ".join(texts(slide))
    for word in ["A", "alpha", "B", "beta", "C", "gamma"]:
        assert word in joined
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd northwestern-ppt && python -m pytest tests/test_components.py::test_card_row_three_cards_rotating_accents -q`
Expected: FAIL — no attribute `render_card_row`.

- [ ] **Step 3: Write the implementation** (append)

```python
def render_card_row(prs, page):
    slide = _content_slide(prs)
    header_band(slide, page.get("eyebrow"), page.get("title", ""),
                use_eyebrow=page.get("use_eyebrow", True))
    cards = page.get("cards", [])[:3]
    n = max(len(cards), 1)
    gap = 0.3
    total_w = ns.CONTENT_W
    card_w = (total_w - gap * (n - 1)) / n
    card_h = 1.5
    top = 2.6
    for i, card in enumerate(cards):
        line_hex, fill_hex = ns.ACCENTS[i % len(ns.ACCENTS)]
        left = ns.MARGIN_L + i * (card_w + gap)
        ns.add_rounded_rect(slide, left, top, card_w, card_h, fill_hex=fill_hex, line_hex=line_hex)
        ns.add_textbox(slide, left + 0.15, top + 0.12, card_w - 0.3, 0.35,
                       card.get("title", ""), 15, line_hex, bold=True)
        ns.add_textbox(slide, left + 0.15, top + 0.55, card_w - 0.3, card_h - 0.65,
                       card.get("desc", ""), 12, ns.INK)
    return slide
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd northwestern-ppt && python -m pytest tests/test_components.py -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "feat: render_card_row component" && echo OK
```
Expected: `OK`

---

## Task 6: `render_pill_flow`

**Files:**
- Modify: `northwestern-ppt/scripts/components.py`
- Test: `northwestern-ppt/tests/test_components.py`

- [ ] **Step 1: Write the failing test** (append)

```python
def test_pill_flow(template_path):
    prs = fresh(template_path)
    page = {"type": "pill_flow", "eyebrow": "PIPELINE", "title": "Steps",
            "pills": ["Ingest", "Bin", "Sample", "Score"], "note": "left to right"}
    slide = components.render_pill_flow(prs, page)
    joined = " | ".join(texts(slide))
    for p in page["pills"]:
        assert p in joined
    assert "left to right" in joined
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    pills = [s for s in slide.shapes
             if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and s.has_text_frame
             and s.text_frame.text.strip() in page["pills"]]
    assert len(pills) == 4
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd northwestern-ppt && python -m pytest tests/test_components.py::test_pill_flow -q`
Expected: FAIL — no attribute `render_pill_flow`.

- [ ] **Step 3: Write the implementation** (append)

```python
def render_pill_flow(prs, page):
    slide = _content_slide(prs)
    header_band(slide, page.get("eyebrow"), page.get("title", ""),
                use_eyebrow=page.get("use_eyebrow", True))
    pills = page.get("pills", [])
    n = max(len(pills), 1)
    pill_h = 0.5
    conn_w = 0.35
    top = 2.7
    total_w = ns.CONTENT_W
    pill_w = (total_w - conn_w * (n - 1)) / n
    x = ns.MARGIN_L
    cy = top + pill_h / 2
    line_hex = ns.PURPLE
    for i, label in enumerate(pills):
        shape = ns.add_rounded_rect(slide, x, top, pill_w, pill_h,
                                    fill_hex="F6F2FA", line_hex=line_hex)
        tf = shape.text_frame
        tf.word_wrap = True
        run = tf.paragraphs[0].add_run()
        run.text = label
        ns.style_run(run, 12, ns.PURPLE_DARK, bold=True)
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        if i < n - 1:
            ns.add_connector(slide, x + pill_w, cy, x + pill_w + conn_w, cy, line_hex)
        x += pill_w + conn_w
    if page.get("note"):
        ns.add_textbox(slide, ns.MARGIN_L, top + 1.0, ns.CONTENT_W, 0.4,
                       page["note"], 12, ns.INK)
    return slide
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd northwestern-ppt && python -m pytest tests/test_components.py -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "feat: render_pill_flow component" && echo OK
```
Expected: `OK`

---

## Task 7: `render_callout`

**Files:**
- Modify: `northwestern-ppt/scripts/components.py`
- Test: `northwestern-ppt/tests/test_components.py`

- [ ] **Step 1: Write the failing test** (append)

```python
def test_callout(template_path):
    prs = fresh(template_path)
    page = {"type": "callout", "eyebrow": "CLAIM", "title": "The point",
            "label": "Core claim", "text": "Process matters more than outcome."}
    slide = components.render_callout(prs, page)
    joined = " | ".join(texts(slide))
    assert "Core claim" in joined
    assert "Process matters more than outcome." in joined
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    boxes = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE]
    assert any(b.fill.fore_color.rgb == ns.hex_to_rgb("F6F2FA") for b in boxes)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd northwestern-ppt && python -m pytest tests/test_components.py::test_callout -q`
Expected: FAIL — no attribute `render_callout`.

- [ ] **Step 3: Write the implementation** (append)

```python
def render_callout(prs, page):
    slide = _content_slide(prs)
    header_band(slide, page.get("eyebrow"), page.get("title", ""),
                use_eyebrow=page.get("use_eyebrow", True))
    box_top = 2.6
    box_h = 1.5
    ns.add_rounded_rect(slide, ns.MARGIN_L, box_top, ns.CONTENT_W, box_h,
                        fill_hex="F6F2FA", line_hex=ns.PURPLE)
    ns.add_textbox(slide, ns.MARGIN_L + 0.25, box_top + 0.18, ns.CONTENT_W - 0.5, 0.35,
                   page.get("label", "").upper(), 11, ns.PURPLE, bold=True)
    ns.add_textbox(slide, ns.MARGIN_L + 0.25, box_top + 0.6, ns.CONTENT_W - 0.5, box_h - 0.75,
                   page.get("text", ""), 16, ns.INK)
    return slide
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd northwestern-ppt && python -m pytest tests/test_components.py -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "feat: render_callout component" && echo OK
```
Expected: `OK`

---

## Task 8: `render_mapping`

**Files:**
- Modify: `northwestern-ppt/scripts/components.py`
- Test: `northwestern-ppt/tests/test_components.py`

- [ ] **Step 1: Write the failing test** (append)

```python
def test_mapping(template_path):
    prs = fresh(template_path)
    page = {"type": "mapping", "eyebrow": "DIAGNOSIS", "title": "Behavior to mode",
            "left_header": "Observed behavior", "right_header": "Attribution",
            "rows": [["No verification", "Task verification"],
                     ["Ignored input", "Misalignment"]]}
    slide = components.render_mapping(prs, page)
    joined = " | ".join(texts(slide))
    for cell in ["Observed behavior", "Attribution", "No verification",
                 "Task verification", "Ignored input", "Misalignment"]:
        assert cell in joined
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd northwestern-ppt && python -m pytest tests/test_components.py::test_mapping -q`
Expected: FAIL — no attribute `render_mapping`.

- [ ] **Step 3: Write the implementation** (append)

```python
def render_mapping(prs, page):
    slide = _content_slide(prs)
    header_band(slide, page.get("eyebrow"), page.get("title", ""),
                use_eyebrow=page.get("use_eyebrow", True))
    rows = page.get("rows", [])
    col_gap = 0.25
    col_w = (ns.CONTENT_W - col_gap) / 2
    left_x = ns.MARGIN_L
    right_x = ns.MARGIN_L + col_w + col_gap
    top = 1.7
    row_h = 0.5
    row_gap = 0.12
    # header cells
    ns.add_rounded_rect(slide, left_x, top, col_w, 0.45, fill_hex=ns.PURPLE, line_hex=None)
    ns.add_textbox(slide, left_x + 0.12, top + 0.06, col_w - 0.24, 0.33,
                   page.get("left_header", ""), 12, ns.WHITE, bold=True)
    ns.add_rounded_rect(slide, right_x, top, col_w, 0.45, fill_hex=ns.PURPLE, line_hex=None)
    ns.add_textbox(slide, right_x + 0.12, top + 0.06, col_w - 0.24, 0.33,
                   page.get("right_header", ""), 12, ns.WHITE, bold=True)
    y = top + 0.45 + row_gap
    for left_text, right_text in rows:
        ns.add_rounded_rect(slide, left_x, y, col_w, row_h, fill_hex="F6F2FA", line_hex="E6E1EC")
        ns.add_textbox(slide, left_x + 0.12, y + 0.09, col_w - 0.24, row_h - 0.18,
                       left_text, 12, ns.INK)
        ns.add_rounded_rect(slide, right_x, y, col_w, row_h, fill_hex="F6F2FA", line_hex="E6E1EC")
        ns.add_textbox(slide, right_x + 0.12, y + 0.09, col_w - 0.24, row_h - 0.18,
                       right_text, 12, ns.INK)
        y += row_h + row_gap
    return slide
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd northwestern-ppt && python -m pytest tests/test_components.py -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "feat: render_mapping component" && echo OK
```
Expected: `OK`

---

## Task 9: `render_figure`

**Files:**
- Modify: `northwestern-ppt/scripts/components.py`
- Test: `northwestern-ppt/tests/test_components.py`

- [ ] **Step 1: Write the failing test** (append)

This test creates a tiny PNG on the fly so it needs no external asset.
```python
def _make_png(path):
    # 4x3 white PNG, minimal, via Pillow if present else a known-good byte blob
    try:
        from PIL import Image
        Image.new("RGB", (400, 300), "white").save(path)
    except Exception:
        import base64
        # 1x1 white PNG
        data = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNgYGAAAAAEAAH2FzhVAAAAAElFTkSuQmCC")
        with open(path, "wb") as f:
            f.write(data)


def test_figure_places_image_and_caption(template_path, tmp_path):
    img = tmp_path / "fig1.png"
    _make_png(str(img))
    prs = fresh(template_path)
    page = {"type": "figure", "eyebrow": "RESULTS", "title": "Main result",
            "image": str(img), "caption": "Accuracy improves with more bins."}
    slide = components.render_figure(prs, page)
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    pics = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
    assert len(pics) == 1
    assert "Accuracy improves with more bins." in " | ".join(texts(slide))
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd northwestern-ppt && python -m pytest tests/test_components.py::test_figure_places_image_and_caption -q`
Expected: FAIL — no attribute `render_figure`.

- [ ] **Step 3: Write the implementation** (append)

```python
def render_figure(prs, page):
    slide = _content_slide(prs)
    header_band(slide, page.get("eyebrow"), page.get("title", ""),
                use_eyebrow=page.get("use_eyebrow", True))
    has_cards = bool(page.get("cards"))
    img_left = ns.MARGIN_L
    img_top = 1.6
    img_max_w = (ns.CONTENT_W - 0.4) * (0.62 if has_cards else 1.0)
    img_max_h = 3.0
    _add_scaled_picture(slide, page["image"], img_left, img_top, img_max_w, img_max_h)
    if page.get("caption"):
        ns.add_textbox(slide, img_left, img_top + img_max_h + 0.1, img_max_w, 0.4,
                       page["caption"], 12, ns.INK)
    if has_cards:
        card_left = ns.MARGIN_L + img_max_w + 0.4
        card_w = ns.CONTENT_W - img_max_w - 0.4
        y = img_top
        for i, card in enumerate(page["cards"][:3]):
            line_hex, fill_hex = ns.ACCENTS[i % len(ns.ACCENTS)]
            ch = 0.95
            ns.add_rounded_rect(slide, card_left, y, card_w, ch, fill_hex=fill_hex, line_hex=line_hex)
            ns.add_textbox(slide, card_left + 0.12, y + 0.1, card_w - 0.24, 0.3,
                           card.get("title", ""), 13, line_hex, bold=True)
            ns.add_textbox(slide, card_left + 0.12, y + 0.42, card_w - 0.24, ch - 0.5,
                           card.get("desc", ""), 11, ns.INK)
            y += ch + 0.2
    return slide


def _add_scaled_picture(slide, image_path, left, top, max_w, max_h):
    """Place an image scaled to fit within (max_w, max_h) inches, preserving aspect."""
    from PIL import Image
    try:
        with Image.open(image_path) as im:
            iw, ih = im.size
    except Exception:
        iw, ih = 4, 3
    aspect = iw / ih if ih else 4 / 3
    w = max_w
    h = w / aspect
    if h > max_h:
        h = max_h
        w = h * aspect
    return slide.shapes.add_picture(image_path, ns.Inches(left), ns.Inches(top),
                                    ns.Inches(w), ns.Inches(h))
```

Add `Pillow` to deps: append `Pillow>=10.0` to `northwestern-ppt/requirements.txt` and run `python3 -m pip install --quiet Pillow`.

- [ ] **Step 4: Run test to verify it passes**

Run: `cd northwestern-ppt && python -m pytest tests/test_components.py -q`
Expected: PASS (all component tests).

- [ ] **Step 5: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "feat: render_figure component with aspect-fit scaling" && echo OK
```
Expected: `OK`

---

## Task 10: `build_pptx.build()` — dispatch, CLI, error handling

**Files:**
- Modify: `northwestern-ppt/scripts/build_pptx.py`
- Test: `northwestern-ppt/tests/test_build.py`

- [ ] **Step 1: Write the failing test** (append to `test_build.py`)

```python
import json
from pptx import Presentation


def test_build_full_outline(template_path, tmp_path):
    outline = {
        "meta": {"title": "T", "use_eyebrow": True},
        "slides": [
            {"type": "title", "title": "Talk", "subtitle": "sub", "source": "src"},
            {"type": "bullets", "eyebrow": "A", "title": "b", "bullets": ["x", "y"]},
            {"type": "callout", "eyebrow": "B", "title": "c", "label": "Claim", "text": "z"},
            {"type": "thanks"},
        ],
    }
    out = tmp_path / "deck.pptx"
    path = build_pptx.build(outline, template_path, str(out))
    prs = Presentation(path)
    assert len(prs.slides._sldIdLst) == 4


def test_build_unknown_type_adds_warning_slide(template_path, tmp_path):
    outline = {"meta": {}, "slides": [{"type": "nonsense", "title": "x"}]}
    out = tmp_path / "warn.pptx"
    build_pptx.build(outline, template_path, str(out))
    prs = Presentation(str(out))
    texts = []
    for slide in prs.slides:
        for sh in slide.shapes:
            if sh.has_text_frame:
                texts.append(sh.text_frame.text)
    assert any("WARNING" in t.upper() for t in texts)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd northwestern-ppt && python -m pytest tests/test_build.py -q`
Expected: FAIL — `AttributeError: module 'build_pptx' has no attribute 'build'`.

- [ ] **Step 3: Write the implementation** (append to `build_pptx.py`)

```python
import sys
import json
import traceback
import components

RENDERERS = {
    "title": components.render_title,
    "thanks": components.render_thanks,
    "bullets": components.render_bullets,
    "card_row": components.render_card_row,
    "pill_flow": components.render_pill_flow,
    "callout": components.render_callout,
    "mapping": components.render_mapping,
    "figure": components.render_figure,
}


def _add_warning_slide(prs, message):
    slide = prs.slides.add_slide(prs.slide_layouts[ns.CONTENT_LAYOUT])
    ns.add_textbox(slide, ns.MARGIN_L, 1.0, ns.CONTENT_W, 0.5,
                   "BUILD WARNING", 18, "CF4F61", bold=True)
    ns.add_textbox(slide, ns.MARGIN_L, 1.8, ns.CONTENT_W, 3.0, message, 12, ns.INK)
    return slide


def build(outline, template_path, out_path):
    prs = load_template(template_path)
    clear_slides(prs)
    use_eyebrow = outline.get("meta", {}).get("use_eyebrow", True)
    for idx, page in enumerate(outline.get("slides", []), 1):
        page.setdefault("use_eyebrow", use_eyebrow)
        ptype = page.get("type")
        renderer = RENDERERS.get(ptype)
        if renderer is None:
            _add_warning_slide(prs, f"Slide {idx}: unknown type {ptype!r}; skipped.")
            continue
        try:
            renderer(prs, page)
        except Exception:
            _add_warning_slide(
                prs, f"Slide {idx} ({ptype}) failed to render:\n{traceback.format_exc()}")
    prs.save(out_path)
    return out_path


def main(argv=None):
    argv = argv or sys.argv[1:]
    if len(argv) != 2:
        print("usage: build_pptx.py <outline.json> <out.pptx>", file=sys.stderr)
        return 2
    outline_path, out_path = argv
    template = str(__import__("pathlib").Path(__file__).resolve().parents[1]
                   / "assets" / "northwestern_template.pptx")
    with open(outline_path, encoding="utf-8") as f:
        outline = json.load(f)
    path = build(outline, template, out_path)
    print(f"wrote {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd northwestern-ppt && python -m pytest tests/test_build.py -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "feat: build() dispatch, CLI, warning-slide error handling" && echo OK
```
Expected: `OK`

---

## Task 11: `extract_pdf.py` — text + figures + figures.json

**Files:**
- Create: `northwestern-ppt/scripts/extract_pdf.py`
- Test: `northwestern-ppt/tests/test_extract.py`

Test fixture: copy a real paper into the repo once.
```bash
cp ~/Downloads/STAT461_final_report_group4.pdf "northwestern-ppt/tests/fixtures/sample.pdf"
```
(Create `northwestern-ppt/tests/fixtures/` first. If that PDF is absent, substitute any text+figure PDF and keep the name `sample.pdf`.)

- [ ] **Step 1: Write the failing test**

Create `northwestern-ppt/tests/test_extract.py`:
```python
import json
from pathlib import Path
import pytest
import extract_pdf

FIXT = Path(__file__).resolve().parent / "fixtures" / "sample.pdf"
pytestmark = pytest.mark.skipif(not FIXT.exists(), reason="sample.pdf fixture missing")


def test_extract_text_returns_pages():
    pages = extract_pdf.extract_text(str(FIXT))
    assert isinstance(pages, list) and len(pages) >= 1
    assert any(p["text"].strip() for p in pages)
    assert all({"page", "text"} <= set(p) for p in pages)


def test_extract_figures_writes_pngs_and_manifest(tmp_path):
    figs = extract_pdf.extract_figures(str(FIXT), str(tmp_path))
    manifest = tmp_path / "figures.json"
    assert manifest.exists()
    data = json.loads(manifest.read_text())
    assert isinstance(data, list)
    for item in data:
        assert {"id", "page", "path"} <= set(item)
        assert Path(item["path"]).exists()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd northwestern-ppt && python -m pytest tests/test_extract.py -q`
Expected: FAIL — `ModuleNotFoundError: No module named 'extract_pdf'` (or skip if fixture missing — in that case create the fixture first).

- [ ] **Step 3: Write the implementation**

Create `northwestern-ppt/scripts/extract_pdf.py`:
```python
"""Extract page text and figure images from a PDF.

Two figure sources:
  1. Embedded raster images (page.get_images).
  2. Caption-anchored regions: find "Figure N"/"Table N" lines and rasterize the
     block of the page above the caption as a fallback PNG.
Outputs <out_dir>/figures/figN.png and <out_dir>/figures.json.
"""
import json
import re
import sys
from pathlib import Path
import fitz  # PyMuPDF

CAPTION_RE = re.compile(r"^\s*(figure|table|fig\.)\s*(\d+)", re.IGNORECASE)


def extract_text(pdf_path):
    doc = fitz.open(pdf_path)
    pages = [{"page": i + 1, "text": doc[i].get_text("text")} for i in range(len(doc))]
    doc.close()
    return pages


def extract_figures(pdf_path, out_dir):
    out = Path(out_dir)
    figs_dir = out / "figures"
    figs_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf_path)
    results = []
    counter = 0

    for pno in range(len(doc)):
        page = doc[pno]
        # 1. embedded raster images
        for img in page.get_images(full=True):
            xref = img[0]
            try:
                pix = fitz.Pixmap(doc, xref)
                if pix.n - pix.alpha >= 4:        # CMYK -> RGB
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                if pix.width < 60 or pix.height < 60:
                    continue                       # skip icons/logos
                counter += 1
                path = figs_dir / f"fig{counter}.png"
                pix.save(str(path))
                results.append({"id": f"fig{counter}", "page": pno + 1,
                                "path": str(path), "source": "embedded", "caption": ""})
            except Exception:
                continue

        # 2. caption-anchored region rasterization (vector charts)
        for block in page.get_text("blocks"):
            text = block[4]
            m = CAPTION_RE.match(text)
            if not m:
                continue
            x0, y0, x1, y1 = block[:4]
            clip = fitz.Rect(x0, max(0, y0 - 220), x1, y0)  # region above caption
            if clip.height < 60:
                continue
            counter += 1
            path = figs_dir / f"fig{counter}.png"
            pix = page.get_pixmap(clip=clip, matrix=fitz.Matrix(2, 2))
            pix.save(str(path))
            results.append({"id": f"fig{counter}", "page": pno + 1, "path": str(path),
                            "source": "caption_region", "caption": text.strip()[:200]})

    doc.close()
    (out / "figures.json").write_text(json.dumps(results, indent=2))
    return results


def main(argv=None):
    argv = argv or sys.argv[1:]
    if len(argv) != 2:
        print("usage: extract_pdf.py <paper.pdf> <out_dir>", file=sys.stderr)
        return 2
    pdf_path, out_dir = argv
    pages = extract_text(pdf_path)
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    (Path(out_dir) / "text.json").write_text(json.dumps(pages, indent=2))
    figs = extract_figures(pdf_path, out_dir)
    print(f"text: {len(pages)} pages; figures: {len(figs)} -> {out_dir}/figures.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd northwestern-ppt && python -m pytest tests/test_extract.py -q`
Expected: PASS (or SKIP if no fixture — create the fixture and re-run to get PASS).

- [ ] **Step 5: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "feat: extract_pdf text + figure extraction" && echo OK
```
Expected: `OK`

---

## Task 12: `render_preview.py` — pptx → PNG via LibreOffice

**Files:**
- Create: `northwestern-ppt/scripts/render_preview.py`

This task has no unit test (depends on the LibreOffice binary). It is verified by a real run.

- [ ] **Step 1: Write the implementation**

Create `northwestern-ppt/scripts/render_preview.py`:
```python
"""Rasterize a .pptx to one PNG per slide for visual self-check.

Pipeline: soffice --headless --convert-to pdf, then PyMuPDF renders each PDF
page to <out_dir>/slide-NN.png. Returns the list of PNG paths.
"""
import shutil
import subprocess
import sys
from pathlib import Path
import fitz


def _soffice():
    for cand in ("soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if shutil.which(cand) or Path(cand).exists():
            return cand
    raise RuntimeError("LibreOffice 'soffice' not found; install it to render previews.")


def render(pptx_path, out_dir, dpi=110):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    soffice = _soffice()
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir",
                    str(out), str(pptx_path)], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    pdf_path = out / (Path(pptx_path).stem + ".pdf")
    doc = fitz.open(str(pdf_path))
    zoom = dpi / 72.0
    paths = []
    for i in range(len(doc)):
        pix = doc[i].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        p = out / f"slide-{i + 1:02d}.png"
        pix.save(str(p))
        paths.append(str(p))
    doc.close()
    return paths


def main(argv=None):
    argv = argv or sys.argv[1:]
    if len(argv) != 2:
        print("usage: render_preview.py <deck.pptx> <out_dir>", file=sys.stderr)
        return 2
    paths = render(argv[0], argv[1])
    print(f"rendered {len(paths)} slides -> {argv[1]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 2: Verify with a real run** (build a 1-slide deck, render it)

Run:
```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent/northwestern-ppt"
python -c "import sys; sys.path.insert(0,'scripts'); import json,build_pptx; \
o={'meta':{},'slides':[{'type':'title','title':'Smoke','subtitle':'preview test'}]}; \
build_pptx.build(o,'assets/northwestern_template.pptx','/tmp/_prev.pptx')"
python scripts/render_preview.py /tmp/_prev.pptx /tmp/_prev_png
ls /tmp/_prev_png/slide-01.png && echo "preview ok"
```
Expected: `/tmp/_prev_png/slide-01.png` exists, then `preview ok`.

- [ ] **Step 3: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "feat: render_preview pptx->png via LibreOffice" && echo OK
```
Expected: `OK`

---

## Task 13: `make_sample.py` — component gallery + visual self-check

**Files:**
- Create: `northwestern-ppt/scripts/make_sample.py`

- [ ] **Step 1: Write the implementation**

Create `northwestern-ppt/scripts/make_sample.py`:
```python
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
```

- [ ] **Step 2: Run it and visually self-check**

Run:
```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent/northwestern-ppt"
python scripts/make_sample.py && ls tests/_out/sample_png/
```
Expected: 7 PNGs `slide-01.png` … `slide-07.png`.

Then **open each PNG with the Read tool** and confirm: purple footer/wordmark present, no text overflow past slide bounds, cards have rotating accent colors, eyebrow+title+underline aligned. If any slide overflows or overlaps, adjust the geometry constants in `nwstyle.py` (or the per-component `top`/heights) and re-run.

- [ ] **Step 3: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "feat: make_sample component gallery + preview" && echo OK
```
Expected: `OK`

---

## Task 14: `SKILL.md` — agent-facing workflow

**Files:**
- Create: `northwestern-ppt/SKILL.md`

- [ ] **Step 1: Write the file**

Create `northwestern-ppt/SKILL.md`:
```markdown
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
```

- [ ] **Step 2: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "docs: SKILL.md workflow" && echo OK
```
Expected: `OK`

---

## Task 15: `STYLE.md` — pagination + selection rules

**Files:**
- Create: `northwestern-ppt/STYLE.md`

- [ ] **Step 1: Write the file**

Create `northwestern-ppt/STYLE.md` capturing spec §4–5:
```markdown
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
```

- [ ] **Step 2: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "docs: STYLE.md pagination and selection rules" && echo OK
```
Expected: `OK`

---

## Task 16: End-to-end on a real paper

**Files:** none (verification run).

- [ ] **Step 1: Run the full pipeline on the fixture paper**

Run:
```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent/northwestern-ppt"
python scripts/extract_pdf.py tests/fixtures/sample.pdf work/
cat work/figures.json | python -c "import sys,json;print(len(json.load(sys.stdin)),'figures')"
```
Expected: text.json + figures.json created; figure count printed.

- [ ] **Step 2: Hand-author a small outline from the extracted text**

Read `work/text.json`, then write `work/outline.json` (4–6 slides) using real content and at least one `figure` referencing a path from `work/figures.json`. Then:
```bash
python scripts/build_pptx.py work/outline.json work/sample_northwestern.pptx
python scripts/render_preview.py work/sample_northwestern.pptx work/preview/
ls work/preview/
```
Expected: one PNG per slide.

- [ ] **Step 3: Visual self-check**

Open each `work/preview/slide-*.png` with the Read tool. Confirm branding, no overflow/overlap, figure fits its box, accents rotate. Fix `nwstyle.py`/components if needed and rebuild.

- [ ] **Step 4: Run the whole test suite**

Run: `cd northwestern-ppt && python -m pytest -q`
Expected: all tests pass (extract tests pass now that the fixture exists).

- [ ] **Step 5: Commit**

```bash
cd "/Users/ljw/Desktop/Courses & Homeworks/Northwestern_PPT_Agent"
git add -A && git commit -q -m "test: end-to-end paper->deck verified" && echo OK
```
Expected: `OK`

---

## Self-Review

**1. Spec coverage**

| Spec section | Task |
|---|---|
| §1 input=PDF, outline-first, no Chinese | Tasks 11, 14, 16 |
| §2 method A (template + python-pptx) | Tasks 1–2, 10 |
| §3 file structure | Task 0 + all |
| §4.1 stable frame | Task 3 (header_band, title, thanks) |
| §4.2 palette | Task 1 (`nwstyle.ACCENTS`, PURPLE…) |
| §4.3 font sizes | Tasks 1, 3–9 |
| §4.4 body menu (bullets/card/pill/callout/mapping/figure) | Tasks 4–9 |
| §5.1 duration→pages | Tasks 14, 15 |
| §5.2 pagination rules | Task 15 |
| §5.3 body-type selection | Task 15 |
| §5.4 outline JSON + markdown preview | Task 14 (schema), Task 10 (consumer) |
| §6.1 extract_pdf (text + figures + figures.json) | Task 11 |
| §6.2 build_pptx dispatch + central style | Tasks 10, 1 |
| §7.1 error handling (skip + warning slide, overflow) | Task 10 (warning slide); overflow handled by `word_wrap` + bounded boxes; figure aspect-fit in Task 9 |
| §7.2 visual verification (LibreOffice → PNG, agent reads) | Tasks 12, 13, 16 |
| §7.2 make_sample component gallery | Task 13 |

Gap noted: spec §7.1 mentions "auto-shrink font to a floor, else truncate + warn." The plan relies on `word_wrap` + generously sized boxes rather than dynamic autofit (python-pptx cannot measure text). This is an acceptable simplification — bounded boxes + the visual self-check loop (Task 16) catch overflow. If overflow recurs in practice, add a `MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE` pass; not pre-implemented to avoid speculative complexity (YAGNI).

**2. Placeholder scan:** No "TBD/TODO/handle edge cases" placeholders; every code step is complete and runnable.

**3. Type consistency:** `render_<type>(prs, page)` signature is uniform across Tasks 3–9 and matches the `RENDERERS` map and tests in Task 10. `nwstyle` names (`PURPLE`, `ACCENTS`, `add_textbox`, `add_rounded_rect`, `add_rect`, `add_connector`, `style_run`, `hex_to_rgb`, `CONTENT_LAYOUT`, `SEPARATOR_LAYOUT`, `MARGIN_L`, `CONTENT_W`, `BODY_TOP`, `BODY_BOTTOM`, `Inches`, `Pt`) are defined in Task 1 and used consistently thereafter. `clear_slides`/`load_template`/`build` names match between Tasks 2, 10 and their tests.
