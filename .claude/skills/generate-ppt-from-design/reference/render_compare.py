#!/usr/bin/env python3
"""Render the built PPTX to per-slide PNGs (via LibreOffice->PDF->PNG) and
build side-by-side comparison strips vs ground-truth (gt/slide-XX.png).

Config via env vars (all optional):
  PPTX   path to the built .pptx     (default: ./<first *.pptx in cwd>)
  GT     ground-truth PNG dir        (default: ./gt)
  OUT    output dir for renders/cmp  (default: ./out)
  SOFFICE path to LibreOffice soffice binary

Usage: python3 render_compare.py [slide numbers...]   (default: all)
Note: converts from /tmp to avoid LibreOffice SIGABRT on OneDrive/spaces paths."""
import os, sys, subprocess, glob, shutil
from PIL import Image
CWD = os.getcwd()
SOFFICE = os.environ.get("SOFFICE", "/Applications/LibreOffice.app/Contents/MacOS/soffice")
def _default_pptx():
    c = sorted(glob.glob(os.path.join(CWD, "*.pptx")))
    return c[0] if c else os.path.join(CWD, "deck.pptx")
PPTX = os.environ.get("PPTX", _default_pptx())
OUT = os.environ.get("OUT", os.path.join(CWD, "out")); os.makedirs(OUT, exist_ok=True)
GT = os.environ.get("GT", os.path.join(CWD, "gt"))

def to_pdf():
    import tempfile
    tmp="/tmp/_bsrender"; os.makedirs(tmp,exist_ok=True)
    src=os.path.join(tmp,"deck.pptx"); shutil.copy(PPTX, src)
    subprocess.run([SOFFICE,"--headless","--convert-to","pdf","--outdir",tmp,
                    "-env:UserInstallation=file:///tmp/_lohome", src],
                   check=True, capture_output=True)
    return os.path.join(tmp, "deck.pdf")

def pdf_to_pngs(pdf, dpi=192):
    # try pdftoppm, then PyMuPDF
    stem=os.path.join(OUT,"slide")
    for f in glob.glob(stem+"*.png"): os.remove(f)
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm","-png","-r",str(dpi),pdf,stem],check=True)
        return sorted(glob.glob(stem+"*.png"))
    try:
        import fitz
        doc=fitz.open(pdf); paths=[]
        for i,pg in enumerate(doc):
            m=fitz.Matrix(dpi/72,dpi/72); pix=pg.get_pixmap(matrix=m)
            p=f"{stem}-{i+1:02d}.png"; pix.save(p); paths.append(p)
        return paths
    except ImportError:
        raise SystemExit("need pdftoppm or pymupdf")

def main():
    pdf=to_pdf(); pngs=pdf_to_pngs(pdf)
    print("rendered", len(pngs), "pages")
    want=[int(a) for a in sys.argv[1:]] or list(range(1,len(pngs)+1))
    cmpdir=os.path.join(OUT,"cmp"); os.makedirs(cmpdir,exist_ok=True)
    for i in want:
        mine=f"{OUT}/slide-{i:02d}.png"
        if not os.path.exists(mine):
            # pdftoppm may name slide-1.png (no zero pad) depending on count
            alt=f"{OUT}/slide-{i}.png"
            mine=alt if os.path.exists(alt) else mine
        gt=f"{GT}/slide-{i:02d}.png"
        if not os.path.exists(mine): print("missing render", i); continue
        m=Image.open(mine).convert("RGB"); m=m.resize((1000,int(1000*m.height/m.width)))
        if os.path.exists(gt):
            g=Image.open(gt).convert("RGB"); g=g.resize((1000,int(1000*g.height/g.width)))
        else:
            g=Image.new("RGB",(1000,562),(240,240,240))
        strip=Image.new("RGB",(1000,g.height+m.height+30),(255,255,255))
        strip.paste(g,(0,0)); strip.paste(m,(0,g.height+30))
        strip.save(f"{cmpdir}/cmp-{i:02d}.png")
    print("comparisons in", cmpdir)
if __name__=="__main__": main()
