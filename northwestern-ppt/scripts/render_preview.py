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
