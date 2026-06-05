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
