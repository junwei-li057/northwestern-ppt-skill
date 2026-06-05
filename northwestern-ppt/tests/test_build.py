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
