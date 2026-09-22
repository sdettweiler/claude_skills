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

## 15. Font WEIGHT is not a bold flag — it's a family (the single biggest fidelity trap)
Designs use numeric weights (300/400/500/600/700). A run's `font.bold=True/False` only gives you **700 or 400** — so every `font-weight: 600` (semibold/Demi) element comes out either too heavy (if you used bold) or too light (if you didn't). This was the #1 correction on a real deck ("you use bold where the original uses semi-bold/demi").
**Why:** commercial families like Avenir Next LT Pro ship each weight as a **separate installed family name**, not as a weight axis of one family. Verify with fontTools before building:
```python
from fontTools.ttLib import TTFont
for f in ttfs:  # nameID 1 = family, 2 = subfamily
    n=TTFont(f)['name']; print(n.getName(1,3,1,0x409), '/', n.getName(2,3,1,0x409))
```
Typical Avenir Next LT Pro result:
- 400 → family `Avenir Next LT Pro`, style Regular  → run: `font="Avenir Next LT Pro", bold=False`
- 700 → family `Avenir Next LT Pro`, style Bold     → run: `font="Avenir Next LT Pro", bold=True`
- **600 → family `Avenir Next LT Pro Demi`, style Regular → run: `font="Avenir Next LT Pro Demi", bold=False`**
- 300 → family `Avenir Next LT Pro Light`
**Fix:** build a `run(text, size, color, w=<weight>)` helper that maps weight→(family_name, bold) and use it for EVERY text run. Read the exact `font-weight` of each element straight from the design HTML/CSS; don't eyeball it. (Eyebrows/headlines/card-titles are usually 700; taglines, kicker labels, "Q3 2026"-style pills, uppercase meta labels are usually 600.)

## 16. Ground-truth must load the design's real @font-face (or GT is wrong for 600/300)
If your GT harness only sets `font-family: "Avenir Next LT Pro"` and relies on the installed font, the browser **cannot find weight 600 in that family** (600 is a different installed family) and **fakes/synthesises** it — so your ground truth itself is wrong for every semibold element, and you'll "match" the wrong thing. **Fix:** in the harness, link the design's own `colors_and_type.css` (its `@font-face` rules map each weight to the correct `.ttf`) and copy the `.ttf`s to where that CSS expects them (`./fonts/…`). Then Chrome renders 600→Demi exactly like the live design.

## 17. LibreOffice renders the Demi/semibold family WIDER and HEAVIER than PowerPoint/Chrome
Your LO preview is unreliable specifically for 600-weight text: LO makes it noticeably heavier and ~6–12% wider than the real font. Consequences and fixes:
- A snug single-line Demi label/pill that is correct in PowerPoint will **overflow or wrap in the LO preview**. Don't chase it by fattening the box (that makes PowerPoint too loose). Instead **verify the run's `typeface=` in the slide XML** (`unzip -p deck.pptx ppt/slides/slideN.xml`) to confirm the correct family is set — that's the deliverable's truth, not the LO raster.
- The reviewer (which reads LO output) will flag Demi text as "too bold / wrong width." Pre-empt it: tell the reviewer the deck targets PowerPoint and to ignore weight/width differences on semibold text, OR sanity-check its weight findings against the XML before acting.
- Give measured-width pills a small (~6%) safety margin so the LO preview doesn't overflow, but keep it small so PowerPoint stays snug.

## 18. Single-line labels: turn word-wrap OFF
Any label meant to be one line (kickers, card titles, taglines, pills, step text) must have `text_frame.word_wrap = False`. Otherwise a renderer that measures the string a few px wider (see #17) silently wraps it to two lines and blows up the layout. Only multi-line text (headlines, body paragraphs, card body) should wrap — and those you pre-wrap yourself (#3). A good default: wrap OFF unless you passed an explicit `wrap_px`.

## 19. Measure geometry and colours from the GT pixels — don't trust CSS math alone
Deriving box heights/positions purely from CSS flexbox + line-height estimates drifts (real leading ≠ your guess). For anything the user calls "too short/too tall/too low," **sample the GT PNG** to get the truth, then divide by the GT scale (2×):
```python
# bounding box of a colour region, and a point colour
def bbox(px,W,H,pred,region): ...   # min/max x,y where pred(r,g,b)
```
Use it to read card heights, bar heights, pill heights, panel top/left, and CTA-bar top/bottom. On one deck this corrected: card height 112→137, orange info-box 90→108, CTA bar 112→150 sitting 40px too low.
For **semi-transparent overlays** (`rgba(...,a)` over a background) don't hand-compute the blend — **sample the rendered pixel** in the GT (e.g. slide-4 chip fill `#171B44`, chip border `#8B4A22`, an outline pill's border). Note a translucent border over a photo has no single colour (it varies with the image underneath) — pick a representative mid value.

## 20. Match line breaks exactly for `text-wrap: pretty`
Designs often set `text-wrap: pretty`/`balance`, which pulls a word down to avoid a short last line — your greedy word-wrap won't reproduce it and the user WILL notice ("check the line breaks"). For visible paragraphs/headlines, read the exact breaks off the GT and emit them as **explicit one-line paragraphs** rather than relying on `wrap_px`. (To extract breaks programmatically, render the harness and wrap each word in a `<span>`, then group by `getBoundingClientRect().top`.)

## 22. Stacked multi-line text: use valign='m' centered on GT-measured line centers
Placing each line in its own box with `valign='t'` (glyph pinned to box top) is renderer-dependent: PowerPoint pulls the glyph up relative to LibreOffice, so line 2 ends up hugging line 1 (user: "the second line sits right under the headline"). **Fix:** give each line its own box, `valign='m'` (middle), with the box's vertical CENTER placed exactly on the line's GT-measured center. That pins the glyph center to a known y in every renderer. Measure the centers off the GT with a **tight** per-line color band — a loose band catches the previous line's descenders/anti-alias and reports the center too high (this bit once: "sub" read 589 with a wide band, 607 with a tight one). Don't assume flex `gap`/`margin-top:auto` translate to even spacing — measure what the design actually renders.

## 23. OOXML shadows render more concentrated than CSS box-shadow
A literal port of `box-shadow: 0 10px 30px rgba(...,0.07)` (blur=30, dist=10, alpha=0.07) comes out noticeably darker/tighter in LibreOffice/PowerPoint than in the browser — enough to show as a grey band in a tight gap (e.g. between cards and a bar below them). To match the browser's soft, near-invisible shadow, **increase the blur and drop the opacity** (e.g. blur≈45–50, alpha≈0.04–0.05). Verify by sampling the GT pixel in the gap (should be within ~1–2% of the base background) and matching it.

## 24. Draw masked/"knockout" overlays NATIVELY, never rasterize them
Decorative SVGs that use a background-colored rect to knock a gap into another line (e.g. an orange connector "jumping" a border, dashed-gap masks) must be drawn with native shapes (`line`/`oval`/`rect`), NOT rasterized. Rasterizing+downscaling anti-aliases the knockout rect's edges into a grey fringe, and a "same as background" mask rect stops being invisible the moment it sits over a shadow/gradient (it erases the shading and reads as a bright box with grey borders — a real user complaint). Native vector shapes have crisp edges and no fringe. Cleanest of all: if the only goal is a line crossing another line, just draw the crossing line ON TOP (correct z-order) and skip the mask entirely. Keep rasterizing SVGs that are pure vector-over-flat-color with no knockouts (logos, clean paths).

## 25. Preserve user manual edits: diff the delivered PPTX, bake exact values
Users nudge shapes in PowerPoint between rounds ("I moved the screenshot, check the new position"). Read the delivered file with python-pptx, dump every shape's `left/top/width/height` in px (`/9525`), diff against your build's computed values, and bake the changed ones back into `build.py` as literals with a `# user-set` comment. Don't re-derive them from CSS — the user's value wins.

## 21. Byte-exact asset extraction when the share link needs auth
`get_file` (DesignSync MCP) auto-persists LARGE binaries to a tool-results `.txt` (byte-exact, decode with `base64` — reliable). SMALL binaries come back inline; do NOT hand-copy that base64 into a file — the model re-emitting a long base64 string corrupts it ("broken PNG / bad IDAT"). Instead: open the design in the **user's authenticated Chrome** (claude-in-chrome), read an `<img>.src` — assets are served from `https://<projectId>.claudeusercontent.com/v1/design/projects/<id>/serve/<path>?t=<signed-token>`. That signed URL is fetchable with **plain `curl` (no cookies)** → pipe straight to disk, byte-exact, no context cost. `curl "<BASE>/serve/<path>?t=<TOK>" -o deck/<path>` for every asset (grab the token once; it's shared across assets). SVGs return as text and can be written directly. Validate every file with `PIL.Image.open().load()` (use `.load()`, not just `.verify()` — verify passes on some streams that fail to decode).
