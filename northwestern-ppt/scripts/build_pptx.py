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
