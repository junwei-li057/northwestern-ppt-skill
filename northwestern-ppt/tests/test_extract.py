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
