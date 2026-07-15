#!/usr/bin/env python3
"""Validate a .pptx for the corruption traps PowerPoint (not LibreOffice) rejects.
Run BEFORE delivering. Usage: python3 validate_pptx.py <deck.pptx>"""
import sys, re, glob, os, zipfile, tempfile, xml.etree.ElementTree as ET

def main():
    path = sys.argv[1]
    tmp = tempfile.mkdtemp()
    with zipfile.ZipFile(path) as z: z.extractall(tmp)
    slides = sorted(glob.glob(os.path.join(tmp, "ppt/slides/slide*.xml")))
    ok = True
    for f in slides:
        x = open(f).read()
        # 1) well-formed
        try: ET.parse(f)
        except Exception as e:
            print(f"[FAIL] {os.path.basename(f)} malformed XML: {e}"); ok = False; continue
        # 2) custGeom must wrap path in pathLst
        if re.search(r'<a:custGeom>\s*(<a:(avLst|gdLst|ahLst|cxnLst|rect)[^>]*/?>)*\s*<a:path[ >]', x):
            print(f"[FAIL] {os.path.basename(f)}: <a:path> not inside <a:pathLst> (PowerPoint will call the file defective)"); ok = False
        if x.count('<a:custGeom>') > x.count('<a:pathLst>'):
            print(f"[FAIL] {os.path.basename(f)}: custGeom without pathLst"); ok = False
        # 3) gradient fills must carry srgbClr, not fall back to schemeClr (blue)
        for g in re.findall(r'<a:gradFill.*?</a:gradFill>', x, re.S):
            if 'schemeClr' in g and 'srgbClr' not in g:
                print(f"[WARN] {os.path.basename(f)}: gradFill uses schemeClr only (may render as theme blue)"); ok = False
    # 4) opens with python-pptx
    try:
        from pptx import Presentation
        n = len(Presentation(path).slides._sldIdLst)
        print(f"[OK] opens with python-pptx: {n} slides")
    except Exception as e:
        print(f"[FAIL] python-pptx cannot open: {e}"); ok = False
    print("RESULT:", "PASS ✅" if ok else "ISSUES FOUND ❌")
    sys.exit(0 if ok else 1)

if __name__ == "__main__": main()
