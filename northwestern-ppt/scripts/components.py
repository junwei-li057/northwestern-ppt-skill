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
