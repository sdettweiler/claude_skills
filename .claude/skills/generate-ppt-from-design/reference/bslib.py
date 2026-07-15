"""Design -> native PPTX helper library (generate-ppt-from-design skill).
Assumes a deck authored at 1280x720 px. Slide = 13.333in x 7.5in. 1 px = 9525 EMU. 1 px = 0.75 pt.
Configure per design via env vars: DECK_FONT (font family name written into text runs),
DECK_FONTS_DIR (folder of .ttf files used ONLY for text-width measurement), RSVG_BIN.
"""
import os, hashlib, subprocess
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image

EMU_PX = 9525
BUILD = os.path.dirname(os.path.abspath(__file__))
DECK = os.path.join(BUILD, "deck")
CACHE = os.path.join(BUILD, "_raster")
os.makedirs(CACHE, exist_ok=True)
RSVG = os.environ.get("RSVG_BIN", "/opt/homebrew/bin/rsvg-convert")
FONT = os.environ.get("DECK_FONT", "Avenir Next LT Pro")   # font family name written into text runs; set per design

def E(px):  return Emu(int(round(px * EMU_PX)))
def P(px):  return Pt(px * 0.75)          # css px -> pt

# ---- text measurement using the design's real TTFs (for pixel-accurate pre-wrapping) ----
# Point DECK_FONTS_DIR at a folder containing the design's .ttf files. The filename
# pattern below is for Avenir Next LT Pro — adjust _font() to match your font's filenames.
from PIL import ImageFont
_FDIR = os.environ.get("DECK_FONTS_DIR", os.path.join(BUILD, "fonts"))
_FCACHE = {}
def _font(weight, italic):
    name = {300:"Light",400:"Regular",500:"Demi",600:"Demi",700:"Bold",800:"Bold"}.get(weight,"Regular")
    if italic: name = ("Bold" if weight>=700 else "Demi" if weight>=500 else "") + "Italic" if name!="Light" else "Italic"
    fn = f"AvenirNextLTPro-{name}.ttf"
    key=(fn,)
    if key not in _FCACHE:
        _FCACHE[key]=fn
    return os.path.join(_FDIR, fn)

def measure(txt, size_px, weight=400, ls=0.0, italic=False):
    try:
        f = ImageFont.truetype(_font(weight, italic), int(round(size_px*4)))
        w = f.getlength(txt)/4.0
    except Exception:
        w = len(txt)*size_px*0.5
    if ls: w += ls*size_px*max(0,len(txt)-1)
    return w

import re as _re2
def _run_w(t, r):
    return measure(t, r.get('size',17), 700 if r.get('bold') else 400, r.get('ls',0), r.get('italic',False))

def wrap_runs(runs, width):
    """Greedy word-wrap a list of styled runs to `width` px using the real font
    metrics (matches Chrome/the original). Returns a list of paragraphs (lines),
    each a list of run dicts."""
    toks=[]
    for r in runs:
        for p in _re2.split(r'(\s+)', r['t']):
            if p=='': continue
            toks.append((p, r, p.isspace()))
    lines=[]; cur=[]; curw=0.0
    for t,r,sp in toks:
        w=_run_w(t,r)
        if not sp and cur and curw+w > width+0.5:
            lines.append(cur); cur=[]; curw=0.0
        if sp and not cur:      # drop leading spaces on a line
            continue
        cur.append((t,r)); curw+=w
    if cur: lines.append(cur)
    out=[]
    for ln in lines:
        while ln and ln[-1][0].isspace(): ln.pop()      # strip trailing space
        merged=[]
        for t,r in ln:
            if merged and merged[-1][1] is r: merged[-1]=(merged[-1][0]+t, r)
            else: merged.append((t,r))
        out.append([{**r,'t':t} for t,r in merged])
    return out or [[]]

def logo_natural(path):
    im = Image.open(path); return im.size

def place_logo_row(slide, logos, x_start, y_center, maxh, maxw, gap):
    """logos: list of file paths. returns end_x"""
    x = x_start
    for p in logos:
        nw, nh = logo_natural(p)
        s = min(maxh/nh, maxw/nw)
        w, h = nw*s, nh*s
        picture(slide, p, x, y_center - h/2, w, h)
        x += w + gap
    return x - gap

# ---- palette (from colors_and_type.css + deck inline vars) ----
C = dict(
    ink="04093A", orange="FF7700", orange400="FF9333", orange50="FFF5EC",
    create="FF7700", govern="5C6DFF", activate="00098C", activate_txt="8FA0FF",
    deepblue="00098C", deepblue100="EFF4FF", deepblue300="5675C3",
    white="FFFFFF", n50="FAFAFB", n100="F4F5F8", n150="EFF4FF", n200="E5E7EE",
    n300="CDD1DC", n400="9AA0B2", n500="6B7184", n600="4A5066",
    fg2="4A5066", fg3="6B7184", fg4="9AA0B2",
    kpi_attention="283373", kpi_persuasion="099376", kpi_strategic="890088",
    kpi_processing="C87700", kpi_emotional="500003", kpi_branding="014201",
    green="2BB673", amber="F2B53C", red="E63946",
    score_low="DE0C00", score_high="80B646",
    lav="EFF4FF",
)

def col(x):
    """accept 'ink' key, hex string, or #hex"""
    if x is None: return None
    if x in C: x = C[x]
    x = x.lstrip('#')
    return RGBColor.from_string(x.upper())

# ---------------------------------------------------------------- presentation
def new_prs():
    prs = Presentation()
    prs.slide_width  = E(1280)
    prs.slide_height = E(720)
    return prs

def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])

def bg(slide, color):
    fe = slide.background.fill
    fe.solid(); fe.fore_color.rgb = col(color)

# ---------------------------------------------------------------- shapes
def _noline(sp):
    sp.line.fill.background()

def rect(slide, x, y, w, h, fill=None, line=None, line_w=1.0, radius=None,
         grad=None, grad_angle=90, shadow=None, dash=None):
    if radius is not None:
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, E(x), E(y), E(w), E(h))
        try:
            adj = max(0.0, min(0.5, radius / min(w, h)))
            shp.adjustments[0] = adj
        except Exception: pass
    else:
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(x), E(y), E(w), E(h))
    shp.shadow.inherit = False
    if grad:
        set_gradient(shp, grad, grad_angle)
    elif fill is not None:
        shp.fill.solid(); shp.fill.fore_color.rgb = col(fill)
    else:
        shp.fill.background()
    if line is not None:
        shp.line.color.rgb = col(line); shp.line.width = P(line_w)
        if dash: _dash(shp, dash)
    else:
        _noline(shp)
    if shadow: set_shadow(shp, **shadow)
    return shp

def _dash(shp, kind):
    ln = shp.line._get_or_add_ln()
    d = ln.find(qn('a:prstDash'))
    if d is None:
        d = etree.SubElement(ln, qn('a:prstDash'))
    d.set('val', kind)  # 'dash','sysDash'

def oval(slide, x, y, w, h, fill=None, line=None, line_w=1.0, grad=None, grad_angle=90, shadow=None, dash=None, glow=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, E(x), E(y), E(w), E(h))
    shp.shadow.inherit = False
    if grad: set_gradient(shp, grad, grad_angle)
    elif fill is not None: shp.fill.solid(); shp.fill.fore_color.rgb = col(fill)
    else: shp.fill.background()
    if line is not None:
        shp.line.color.rgb = col(line); shp.line.width = P(line_w)
        if dash: _dash(shp, dash)
    else: _noline(shp)
    if glow: set_glow(shp, **glow)
    if shadow: set_shadow(shp, **shadow)
    return shp

def accent_card(slide, x, y, w, h, radius, fill, accent, side='top', bar=4,
                shadow=None, accent_grad=None, accent_grad_angle=0):
    """rounded card with a colored bar on one side that follows the corner radius."""
    ac = rect(slide, x, y, w, h, fill=(None if accent_grad else accent),
              grad=accent_grad, grad_angle=accent_grad_angle, radius=radius, shadow=shadow)
    if side=='top':
        body = rect(slide, x, y+bar, w, h-bar, fill=fill, radius=radius)
    elif side=='left':
        body = rect(slide, x+bar, y, w-bar, h, fill=fill, radius=radius)
    elif side=='bottom':
        body = rect(slide, x, y, w, h-bar, fill=fill, radius=radius)
    else:  # right
        body = rect(slide, x, y, w-bar, h, fill=fill, radius=radius)
    return ac, body

def polygon(slide, pts, fill=None, grad=None, grad_angle=0, line=None, line_w=1.0):
    fb = slide.shapes.build_freeform(float(pts[0][0]), float(pts[0][1]), scale=EMU_PX)
    fb.add_line_segments([(float(x), float(y)) for x, y in pts[1:]], close=True)
    shp = fb.convert_to_shape()
    shp.shadow.inherit = False
    if grad: set_gradient(shp, grad, grad_angle)
    elif fill is not None: shp.fill.solid(); shp.fill.fore_color.rgb = col(fill)
    else: shp.fill.background()
    if line is not None: shp.line.color.rgb = col(line); shp.line.width = P(line_w)
    else: _noline(shp)
    return shp

def line(slide, x1, y1, x2, y2, color, w=1.0, dash=None, cap='rnd'):
    cn = slide.shapes.add_connector(2, E(x1), E(y1), E(x2), E(y2))  # straight
    cn.line.color.rgb = col(color); cn.line.width = P(w)
    cn.shadow.inherit = False
    ln = cn.line._get_or_add_ln(); ln.set('cap', cap)
    if dash: _dash(cn, dash)
    return cn

# ---------------------------------------------------------------- gradient / shadow xml
def set_gradient(shp, stops, angle_deg=90):
    """stops = [(pos0..1, 'hex', alpha0..1|None), ...]; angle_deg css-like.
    Seeds a schema-correct gradFill via python-pptx's native API (guarantees
    proper element placement in spPr + rotWithShape so strict PowerPoint accepts
    it), then swaps in our own colour stops. Hand-appended gradFill was being
    silently dropped by PowerPoint (falling back to theme accent1 = blue)."""
    try:
        shp.fill.gradient()  # correct placement, rotWithShape="1"
        grad = shp.fill._xPr.find(qn('a:gradFill'))
    except Exception:
        # fallback: insert manually right after geometry, before a:ln
        spPr = shp.fill._xPr
        for tag in ('a:noFill','a:solidFill','a:gradFill','a:blipFill','a:pattFill','a:grpFill'):
            el = spPr.find(qn(tag))
            if el is not None: spPr.remove(el)
        grad = etree.Element(qn('a:gradFill')); grad.set('rotWithShape','1')
        geom = spPr.find(qn('a:prstGeom'))
        if geom is None: geom = spPr.find(qn('a:custGeom'))
        if geom is not None: geom.addnext(grad)
        else: spPr.append(grad)
    # rebuild gsLst with our stops
    lst = grad.find(qn('a:gsLst'))
    if lst is None: lst = etree.SubElement(grad, qn('a:gsLst'))
    for gs in list(lst): lst.remove(gs)
    for pos, hexc, alpha in stops:
        gs = etree.SubElement(lst, qn('a:gs')); gs.set('pos', str(int(round(pos*100000))))
        c = etree.SubElement(gs, qn('a:srgbClr')); c.set('val', hexc.lstrip('#').upper())
        if alpha is not None:
            a = etree.SubElement(c, qn('a:alpha')); a.set('val', str(int(round(alpha*100000))))
    # angle on a:lin (0 = left->right, clockwise, 60000ths of a degree)
    lin = grad.find(qn('a:lin'))
    if lin is None: lin = etree.SubElement(grad, qn('a:lin'))
    ang = int(round(angle_deg * 60000)) % (360*60000)
    lin.set('ang', str(ang)); lin.set('scaled', '1')
    # a:path (radial leftover) must not coexist with a:lin
    p = grad.find(qn('a:path'))
    if p is not None: grad.remove(p)

def _effectlst(shp):
    spPr = shp._element.spPr
    el = spPr.find(qn('a:effectLst'))
    if el is None: el = etree.SubElement(spPr, qn('a:effectLst'))
    return el

def set_shadow(shp, blur=18, dist=8, dir=90, color='14182D', alpha=0.85, sx=None, sy=None, algn=None):
    el = _effectlst(shp)
    sh = etree.SubElement(el, qn('a:outerShdw'))
    sh.set('blurRad', str(int(blur*EMU_PX)))
    sh.set('dist', str(int(dist*EMU_PX)))
    sh.set('dir', str(int(dir*60000)))
    if sx is not None: sh.set('sx', str(int(sx*1000)))
    if sy is not None: sh.set('sy', str(int(sy*1000)))
    if algn: sh.set('algn', algn)
    sh.set('rotWithShape','0')
    c = etree.SubElement(sh, qn('a:srgbClr')); c.set('val', color.lstrip('#').upper())
    a = etree.SubElement(c, qn('a:alpha')); a.set('val', str(int((1-alpha)*100000)))

def set_glow(shp, color='FFFFFF', size_pct=120, alpha=0.05):
    """white halo glow via a scaled 0-blur/0-dist outer shadow (user's spec)."""
    set_shadow(shp, blur=0, dist=0, dir=0, color=color, alpha=alpha, sx=size_pct, sy=size_pct, algn='ctr')

def set_softglow(shp, color='FF7700', rad_px=9, opacity=0.35):
    """true soft halo via a:glow (blurred, edgeless) — not a hard scaled shadow ring."""
    el = _effectlst(shp)
    g = etree.SubElement(el, qn('a:glow')); g.set('rad', str(int(rad_px*EMU_PX)))
    c = etree.SubElement(g, qn('a:srgbClr')); c.set('val', color.lstrip('#').upper())
    a = etree.SubElement(c, qn('a:alpha')); a.set('val', str(int(opacity*100000)))

# ---------------------------------------------------------------- text
def _set_spacing(rpr, em, font_px):
    if em:
        spc = int(round(em * font_px * 0.75 * 100))
        rpr.set('spc', str(spc))

def text(slide, x, y, w, h, spans, align='l', valign='t', line_h=None,
         wrap=True, autosize=False, wrap_px=None):
    """spans: list of paragraphs; each paragraph = list of run dicts
       run = {t, size(px), color, bold, italic, font, ls(em), color..}
       OR spans can be a single list of runs (one paragraph).
       wrap_px: if set, pre-wrap runs to this width using real font metrics and
                emit explicit line breaks (renderer-independent, matches source)."""
    if wrap_px is not None:
        wrap = False
        if spans and isinstance(spans[0], dict):
            spans = wrap_runs(spans, wrap_px)
        else:
            ns=[]
            for para in spans:
                runs = para['runs'] if isinstance(para, dict) else para
                ns.extend(wrap_runs(runs, wrap_px))
            spans = ns
    tb = slide.shapes.add_textbox(E(x), E(y), E(w), E(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    from pptx.enum.text import MSO_AUTO_SIZE
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = {'t':MSO_ANCHOR.TOP,'m':MSO_ANCHOR.MIDDLE,'b':MSO_ANCHOR.BOTTOM}[valign]
    # normalize: if spans[0] is a dict -> single paragraph
    if spans and isinstance(spans[0], dict):
        spans = [spans]
    first = True
    for para in spans:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = {'l':PP_ALIGN.LEFT,'c':PP_ALIGN.CENTER,'r':PP_ALIGN.RIGHT,'j':PP_ALIGN.JUSTIFY}[align]
        lh = line_h
        # per-paragraph line height override
        if isinstance(para, dict):
            lh = para.get('line_h', line_h); align_p = para.get('align'); runs = para['runs']
            if align_p: p.alignment = {'l':PP_ALIGN.LEFT,'c':PP_ALIGN.CENTER,'r':PP_ALIGN.RIGHT,'j':PP_ALIGN.JUSTIFY}[align_p]
        else:
            runs = para
        # dominant font size in this paragraph (for exact leading)
        _sz = max([rn.get('size',17) for rn in runs], default=17)
        if lh:
            pPr = p._p.get_or_add_pPr()
            ls = pPr.find(qn('a:lnSpc'))
            if ls is None: ls = etree.SubElement(pPr, qn('a:lnSpc')); pPr.insert(0, ls)
            # EXACT leading = line_h * font_px * 0.75 pt  (CSS line-height semantics)
            pts = etree.SubElement(ls, qn('a:spcPts')); pts.set('val', str(int(round(lh*_sz*0.75*100))))
        p.space_before = Pt(0); p.space_after = Pt(0)
        for rn in runs:
            r = p.add_run(); r.text = rn['t']
            f = r.font
            f.size = P(rn.get('size',17))
            f.name = rn.get('font', FONT)
            f.bold = rn.get('bold', False)
            f.italic = rn.get('italic', False)
            f.color.rgb = col(rn.get('color','ink'))
            try: f.language_id = MSO_LANGUAGE_ID.ENGLISH_US
            except Exception: pass
            _set_spacing(r._r.get_or_add_rPr(), rn.get('ls',0), rn.get('size',17))
    return tb

# ---------------------------------------------------------------- images
def picture(slide, path, x, y, w, h):
    return slide.shapes.add_picture(path, E(x), E(y), E(w), E(h))

def set_round_top(shp, radius, w, h):
    """crop a picture/shape to a rectangle with only the top two corners rounded."""
    sp = shp._element.spPr
    g = sp.find(qn('a:prstGeom'))
    if g is not None: sp.remove(g)
    xfrm = sp.find(qn('a:xfrm'))
    g = etree.Element(qn('a:prstGeom')); g.set('prst','round2SameRect')
    av = etree.SubElement(g, qn('a:avLst'))
    r = max(0,min(50000,int(radius/min(w,h)*100000)))
    for nm,val in (('adj1',r),('adj2',0)):
        gd=etree.SubElement(av,qn('a:gd')); gd.set('name',nm); gd.set('fmla','val %d'%val)
    xfrm.addnext(g)
    return shp

def set_round_all(shp, radius, w, h):
    """crop a picture/shape to a rectangle with all four corners rounded."""
    sp = shp._element.spPr
    g = sp.find(qn('a:prstGeom'))
    if g is not None: sp.remove(g)
    xfrm = sp.find(qn('a:xfrm'))
    g = etree.Element(qn('a:prstGeom')); g.set('prst','roundRect')
    av = etree.SubElement(g, qn('a:avLst'))
    r = max(0,min(50000,int(radius/min(w,h)*100000)))
    gd=etree.SubElement(av,qn('a:gd')); gd.set('name','adj'); gd.set('fmla','val %d'%r)
    if xfrm is not None: xfrm.addnext(g)
    else: sp.append(g)
    return shp

def set_round_bottom(shp, radius, w, h):
    """clip a picture/shape to a rectangle whose BOTTOM two corners are rounded
    (top corners stay square — they tuck against a chrome bar)."""
    sp = shp._element.spPr
    g = sp.find(qn('a:prstGeom'))
    if g is not None: sp.remove(g)
    xfrm = sp.find(qn('a:xfrm'))
    W=int(round(w*EMU_PX)); H=int(round(h*EMU_PX)); R=int(round(min(radius,min(w,h)/2)*EMU_PX))
    cg = etree.Element(qn('a:custGeom'))
    for t in ('a:avLst','a:gdLst','a:ahLst','a:cxnLst'): etree.SubElement(cg,qn(t))
    rc=etree.SubElement(cg,qn('a:rect')); rc.set('l','0'); rc.set('t','0'); rc.set('r',str(W)); rc.set('b',str(H))
    plst=etree.SubElement(cg,qn('a:pathLst'))               # <a:path> MUST be wrapped in <a:pathLst>
    path=etree.SubElement(plst,qn('a:path')); path.set('w',str(W)); path.set('h',str(H))
    def pt(parent,X,Y):
        e=etree.SubElement(parent,qn('a:pt')); e.set('x',str(int(X))); e.set('y',str(int(Y)))
    m=etree.SubElement(path,qn('a:moveTo')); pt(m,0,0)
    l=etree.SubElement(path,qn('a:lnTo')); pt(l,W,0)
    l=etree.SubElement(path,qn('a:lnTo')); pt(l,W,H-R)
    a=etree.SubElement(path,qn('a:arcTo')); a.set('wR',str(R)); a.set('hR',str(R)); a.set('stAng','0'); a.set('swAng','5400000')
    l=etree.SubElement(path,qn('a:lnTo')); pt(l,R,H)
    a=etree.SubElement(path,qn('a:arcTo')); a.set('wR',str(R)); a.set('hR',str(R)); a.set('stAng','5400000'); a.set('swAng','5400000')
    etree.SubElement(path,qn('a:close'))
    if xfrm is not None: xfrm.addnext(cg)
    else: sp.append(cg)
    return shp

def pic_cover(slide, path, x, y, w, h, pos='center', radius=None):
    """object-fit:cover crop into (w,h) box; pos in {'center','top'}"""
    im = Image.open(path); iw, ih = im.size
    tr = w/h; ir = iw/ih
    if ir > tr:  # image wider -> crop sides
        cw = int(ih*tr); cx = (iw-cw)//2; crop=(cx,0,cx+cw,ih)
    else:        # image taller -> crop top/bottom
        ch = int(iw/tr)
        cy = 0 if pos=='top' else (ih-ch)//2
        crop=(0,cy,iw,cy+ch)
    tmp = os.path.join(CACHE, "cov_"+hashlib.md5((path+str(crop)).encode()).hexdigest()[:12]+".png")
    if not os.path.exists(tmp):
        im.crop(crop).save(tmp)
    p = slide.shapes.add_picture(tmp, E(x), E(y), E(w), E(h))
    return p

# ---------------------------------------------------------------- svg -> png raster
def raster_svg_file(path, out_w_px, scale=4):
    key = hashlib.md5((path+str(out_w_px)+str(scale)).encode()).hexdigest()[:14]
    out = os.path.join(CACHE, "svg_"+key+".png")
    if not os.path.exists(out):
        subprocess.run([RSVG, "-w", str(int(out_w_px*scale)), path, "-o", out], check=True)
    return out

def raster_svg_str(svg, out_w_px, scale=4, tag=""):
    key = hashlib.md5((svg+str(out_w_px)+str(scale)+tag).encode()).hexdigest()[:14]
    out = os.path.join(CACHE, "svgs_"+key+".png")
    if not os.path.exists(out):
        tmp = os.path.join(CACHE, "tmp_"+key+".svg")
        open(tmp,"w").write(svg)
        subprocess.run([RSVG, "-w", str(int(out_w_px*scale)), tmp, "-o", out], check=True)
    return out

def icon(slide, inner, x, y, size, stroke='ink', sw=1.9, fill='none', viewbox="0 0 24 24"):
    """lucide-style stroke icon. inner = the svg inner markup (paths)."""
    stroke_hex = C.get(stroke, stroke).lstrip('#')
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="{viewbox}" fill="{fill}" '
           f'stroke="#{stroke_hex}" stroke-width="{sw}" stroke-linecap="round" '
           f'stroke-linejoin="round">{inner}</svg>')
    png = raster_svg_str(svg, size)
    return slide.shapes.add_picture(png, E(x), E(y), E(size), E(size))

def svg_img(slide, svg, x, y, w, h, tag=""):
    png = raster_svg_str(svg, w, tag=tag)
    return slide.shapes.add_picture(png, E(x), E(y), E(w), E(h))

# ---------------------------------------------------------------- native SVG embed (vector + png fallback)
SVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
def _svg_embed_common(slide, svg_path, png_path, x, y, w, h):
    pic = slide.shapes.add_picture(png_path, E(x), E(y), E(w), E(h))
    # add svg as related part
    from pptx.parts.image import Image as PptxImage
    part = slide.part
    svg_bytes = open(svg_path, 'rb').read()
    # register image part
    from pptx.opc.package import Part
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    partname = part.package.next_partname('/ppt/media/svgimage%d.svg')
    svg_part = Part(partname, 'image/svg+xml', part.package, svg_bytes)
    rId = part.relate_to(svg_part, RT.IMAGE)
    blip = pic._element.blipFill.find(qn('a:blip'))
    ext_lst = blip.find(qn('a:extLst'))
    if ext_lst is None:
        ext_lst = etree.SubElement(blip, qn('a:extLst'))
    ext = etree.SubElement(ext_lst, qn('a:ext'))
    ext.set('uri', '{96DAC541-7B7A-43D3-8B79-37D633B846F1}')
    svgblip = etree.SubElement(ext, '{%s}svgBlip' % SVG_NS)
    svgblip.set(qn('r:embed'), rId)
    return pic

def svg_embed_file(slide, svg_path, x, y, w, h, scale=4):
    png = raster_svg_file(svg_path, w, scale=scale)
    return _svg_embed_common(slide, svg_path, png, x, y, w, h)

def svg_embed_str(slide, svg, x, y, w, h, scale=4, tag=""):
    png = raster_svg_str(svg, w, scale=scale, tag=tag)
    key = hashlib.md5((svg+tag).encode()).hexdigest()[:14]
    tmp = os.path.join(CACHE, "emb_"+key+".svg")
    if not os.path.exists(tmp): open(tmp,'w').write(svg)
    return _svg_embed_common(slide, tmp, png, x, y, w, h)

import re as _re
def svg_cover_embed(slide, svg_path, x, y, w, h, scale=4):
    """object-fit:cover for an SVG file, kept vector by cropping the viewBox."""
    svg = open(svg_path, encoding='utf-8').read()
    m = _re.search(r'viewBox="([-\d.eE ]+)"', svg)
    vx, vy, vw, vh = [float(t) for t in m.group(1).split()]
    ta = w/h; sa = vw/vh
    if ta > sa:                    # target wider -> fill width, crop top/bottom
        nvh = vw/ta; nvy = vy+(vh-nvh)/2; nvb = (vx, nvy, vw, nvh)
    else:                          # target taller -> fill height, crop sides
        nvw = vh*ta; nvx = vx+(vw-nvw)/2; nvb = (nvx, vy, nvw, vh)
    nsvg = _re.sub(r'viewBox="[-\d.eE ]+"',
                   'viewBox="%.3f %.3f %.3f %.3f"' % nvb, svg, count=1)
    nsvg = nsvg.replace('slice', 'meet')
    return svg_embed_str(slide, nsvg, x, y, w, h, scale=scale,
                         tag=os.path.basename(svg_path)+str(round(w))+str(round(h)))

def svg_cover_raster(slide, svg_path, x, y, w, h, scale=5):
    """object-fit:cover for a filtered SVG, rasterised to PNG (renders filters/glows
    identically in LibreOffice + PowerPoint). Returns the picture shape."""
    svg = open(svg_path, encoding='utf-8').read()
    m = _re.search(r'viewBox="([-\d.eE ]+)"', svg)
    vx, vy, vw, vh = [float(t) for t in m.group(1).split()]
    ta = w/h; sa = vw/vh
    if ta > sa: nvh=vw/ta; nvy=vy+(vh-nvh)/2; nvb=(vx,nvy,vw,nvh)
    else:       nvw=vh*ta; nvx=vx+(vw-nvw)/2; nvb=(nvx,vy,nvw,vh)
    nsvg = _re.sub(r'viewBox="[-\d.eE ]+"','viewBox="%.3f %.3f %.3f %.3f"'%nvb, svg, count=1).replace('slice','meet')
    png = raster_svg_str(nsvg, w, scale=scale, tag='cov'+os.path.basename(svg_path)+str(round(w))+str(round(h)))
    return slide.shapes.add_picture(png, E(x), E(y), E(w), E(h))
