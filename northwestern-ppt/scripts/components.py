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
    # Use a connector (LINE) for the underline so it does not register as an
    # AUTO_SHAPE — keeps card/pill fill-detection tests clean.
    ns.add_connector(slide, l, t + h / 2, l + w, t + h / 2, ns.PURPLE)
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
