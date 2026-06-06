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
