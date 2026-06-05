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
