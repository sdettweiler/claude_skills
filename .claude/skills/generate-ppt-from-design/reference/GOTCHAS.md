# GOTCHAS — read before writing PPTX code

Every item below is a real bug or trap hit during a full pixel-perfect build. Most stem from **one root truth: LibreOffice (your render/preview) is lenient and substitutes fonts; PowerPoint (the deliverable target) is strict and uses the real font.** A thing can look perfect in your LibreOffice render and still be wrong — or corrupt — in PowerPoint.

---

## 1. Gradients silently turn BLUE in PowerPoint (most costly bug)
**Symptom:** gradient-filled shapes look right in LibreOffice but render as flat theme-blue in PowerPoint.
**Cause:** a hand-rolled `<a:gradFill>` appended to `spPr` in the wrong position is **rejected by strict PowerPoint**, which then falls back to the shape's style default `<a:fillRef idx="…"><a:schemeClr val="accent1"/>` = blue. LibreOffice renders the misplaced gradient anyway.
**Fix:** seed the gradient with python-pptx's native API so element order + `rotWithShape` are correct, then swap in your own stops:
```python
shp.fill.gradient()                       # correct placement in spPr, rotWithShape="1"
grad = shp.fill._xPr.find(qn('a:gradFill'))
lst = grad.find(qn('a:gsLst'))
for gs in list(lst): lst.remove(gs)
for pos, hexc, alpha in stops:            # your srgbClr stops
    gs = SubElement(lst, qn('a:gs')); gs.set('pos', str(int(pos*100000)))
    c = SubElement(gs, qn('a:srgbClr')); c.set('val', hexc.lstrip('#').upper())
    if alpha is not None: SubElement(c, qn('a:alpha')).set('val', str(int(alpha*100000)))
lin = grad.find(qn('a:lin')) or SubElement(grad, qn('a:lin'))
lin.set('ang', str(int(angle_deg*60000)%21600000)); lin.set('scaled','1')
```
(This is exactly what `bslib.set_gradient` does.) **Validate:** the gradFill must contain `srgbClr`, not `schemeClr`. `accent1` may appear only inside the harmless `<p:style>` block.
**CSS→OOXML angle:** OOXML `ang` = CSS angle − 90 (both clockwise-ish); e.g. CSS `135deg` → OOXML `45°` = 2700000.

## 2. custGeom without `<a:pathLst>` = "file is defective" in PowerPoint
**Symptom:** PowerPoint says the file is corrupt / offers to repair; LibreOffice opened it fine.
**Cause:** `<a:path>` placed directly under `<a:custGeom>`. The schema requires it inside `<a:pathLst>`.
**Correct order inside custGeom:** `avLst, gdLst, ahLst, cxnLst, rect, pathLst > path`.
```xml
<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>
  <a:rect l="0" t="0" r="W" b="H"/>
  <a:pathLst><a:path w="W" h="H"> …moveTo/lnTo/arcTo/close… </a:path></a:pathLst>
</a:custGeom>
```
`arcTo` needs all four attrs: `wR hR stAng swAng` (angles in 60000ths; 90° = 5400000). See `bslib.set_round_bottom`.

## 3. Text line breaks won't match unless you PRE-WRAP at the true container width
**Symptom:** headline/paragraph breaks in the wrong place vs the original.
**Cause:** letting PowerPoint auto-wrap uses the fallback font's metrics and the box width, not the design's.
**Fix:** pre-compute the wrap yourself and emit explicit line paragraphs. `bslib.text(..., wrap_px=W)` does this using PIL `ImageFont.getlength` on the real TTF.
- **The container width is NOT the CSS `max-width`** — it's the *inner* content width after padding. E.g. a cover `.inner { width:720; padding:0 70 0 90 }` gives 560, but an `h1{max-width:660}` inside a 720 content-box stays 660. Measure the actual constraint.
- `measure()` (PIL) slightly **overestimates** vs Chrome and differs from LibreOffice. So the wrap width that reproduces the *original's* breaks is often a bit **larger** than the literal box. Tune `wrap_px` by testing which value reproduces the ground-truth breaks, then set the text-box frame wide enough that the pre-wrapped lines don't get re-wrapped or clipped.
- Set `auto_size = MSO_AUTO_SIZE.NONE` and `word_wrap` off for pre-wrapped text so PowerPoint doesn't re-flow it.
- If a line is 640px-of-measure but the box is 545px, LibreOffice may clip the last word. Widen the frame (text overflows invisibly in PPT) or drop the font size a hair.

## 4. Exact text spacing
- **Line height:** use `<a:spcPts val="int(line_h * size_px * 0.75 * 100)"/>` (absolute points), NOT `a:spcPct` — percentage spacing drifts and looks too loose on headlines.
- **Letter-spacing:** run property `spc` attr = `em × size_px × 0.75 × 100` (can be negative for tight display type). `measure()` doesn't account for it, so a line with negative tracking renders narrower than measured — leave headroom.

## 5. Glow / halo (two different techniques)
- **User's "PowerPoint glow" spec** = white outer shadow, *95% transparency*, 120% size, 0 blur, 0 distance, 0 angle. In OOXML `srgbClr` alpha is **opacity** (100000 = opaque). 95% transparency = 5% opacity = `alpha val="5000"`. If your helper does `val=(1-alpha)*100000`, pass `alpha=0.95` to get 5% opacity. **Getting this inverted gives a harsh ~95%-opaque ring** (user will say "you set it to 6%").
- That scaled-0-blur shadow makes a *hard-edged* ring. For a soft halo use the real `<a:glow rad="…"><a:srgbClr><a:alpha/></a:srgbClr></a:glow>` element (`bslib.set_softglow`). Use the hard one only when the user explicitly specced it.

## 6. Rounded corners — three tools
- All four corners: `prstGeom` **roundRect** (`bslib.set_round_all`).
- Top two only (e.g. a browser chrome bar): `prstGeom` **round2SameRect** (`bslib.set_round_top`).
- Bottom two only (or arbitrary): **custGeom** (`bslib.set_round_bottom`) — mind gotcha #2.
- Setting a geometry on a **picture** clips the image (blipFill) to that shape — this is how you mask screenshots.

## 7. Screenshots inside a browser "frame"
Reproduce the CSS `.shot { width:100% }` behavior: the frame height follows the image aspect; the image fills the width. Then:
- Cover-crop or size-to-aspect so there are **no white letterbox bands**.
- **Mask the image's bottom corners to the frame radius** and keep it just inside the frame so it doesn't clip past the rounded bottom (this was a specific user complaint on 3 slides). Draw the chrome bar (rounded top) on top of the image's straight top edge.
- All mockups here were aspect 1.551 (1912×1233).

## 8. SVGs: embed vs rasterize
- Native SVG in PPTX = `svgBlip` extension `{96DAC541-…}` + a PNG fallback (both needed).
- **But LibreOffice AND PowerPoint drop SVG filters** (`feDropShadow`, glow). If an SVG relies on filters, **rasterize it to PNG with `rsvg-convert`** at 2–4× and embed that instead (`bslib.svg_cover_raster`). Rasterize anything with shadows/glows; embed clean vector SVGs natively.

## 9. LibreOffice crashes (SIGABRT) converting a file on a OneDrive/spaces path
**Fix:** copy the `.pptx` to `/tmp` first, convert there, then read the PDF/PNGs back. `render_compare.py` already does this. Also use a throwaway `-env:UserInstallation=file:///tmp/…` profile.

## 10. Aspect ratio / bottom whitespace
The GT screenshots (2560×1440) and a 1280×720 deck are both 16:9 — so there's **no** aspect mismatch. When the user reports "too much whitespace at the bottom," it's almost always **content positioned too high**, not an aspect problem. For each flagged slide, measure the lowest non-white content row in GT vs your render; if yours ends higher, move the block down / increase inter-element spacing / enlarge proportionally to fill (per user: "space out/resize objects proportionally").

## 11. Asset extraction
Never hand-retype long base64 — you'll get "Incorrect padding" / corrupt PNGs. Decode byte-exact from the design payload or the tool-results transcript and validate with PIL. If DLP blocks browser extraction, ask the user for the files; don't circumvent.

## 12. Preserving manual edits — how to diff a PowerPoint-resaved file
PowerPoint rewrites/reorders XML on save, so raw XML diff is noise. Diff *semantically* with python-pptx: per slide, compare the set of text-run strings, and shape geometry rounded to px, and fills/font-size/weight. Watch for:
- PowerPoint merging your pre-wrapped multi-paragraph text into one auto-wrapping paragraph (cosmetically identical — usually ignore).
- Genuine nudges (uniform +N px shift of a group), width/font changes, wording edits.
Bake confirmed changes into `build.py` with a `# user manual edit` comment. `auto_size=NONE` prevents PowerPoint from silently resizing your boxes on open.

## 13. Units & geometry cheatsheet
- `1 px = 9525 EMU`; `1 px = 0.75 pt`. Deck 1280×720 px = 13.333"×7.5".
- Slide from a blank layout (index 6). Set `slide_width/height` explicitly.
- python-pptx autoshapes carry a `<p:style>` with `fillRef/lnRef/effectRef → accent1`; harmless as long as your explicit `spPr` fill/line is valid (see #1).

## 14. Review loop discipline
- Render is a preview; **truth is the GT PNG**. Read both images directly and compare.
- Tell the reviewer to **ignore font glyph/kerning** diffs and only flag structural ones — otherwise the fallback font generates endless false positives.
- Re-run the reviewer after fixes; fixing one slide can shift another. Stop when no High/Med findings remain.
- Keep a list of **intentional deviations** (things the user asked to differ from the original) and pass it to the reviewer so it doesn't re-flag them.
