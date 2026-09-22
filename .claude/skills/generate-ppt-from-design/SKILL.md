---
name: generate-ppt-from-design
description: Turn a Claude design (claude.ai/design) into a pixel-perfect, fully-editable native PowerPoint deck. Imports the design, extracts exact assets, renders ground-truth screenshots, builds native PPTX shapes with python-pptx, then auto-reviews and iterates until it matches. Trigger: /generate-ppt-from-design
trigger: /generate-ppt-from-design
---

# /generate-ppt-from-design

Rebuild a Claude design as a **native, editable PowerPoint deck** (real shapes, text, gradients, images — not screenshots-on-slides) that matches the original pixel-for-pixel, then verify it with automated review loops until it's perfect.

This skill is the distilled playbook from a full end-to-end build. **Read `reference/GOTCHAS.md` before writing any PPTX code** — it lists the non-obvious bugs that will otherwise cost you hours (PowerPoint-vs-LibreOffice divergence, gradient/geometry corruption, font/line-break mismatch, etc.). The reusable engine is `reference/bslib.py`.

## Step 0 — Ask for the design link and confirm scope

Ask the user (use the project's preferred question mechanism):
1. **The Claude design link** (e.g. `https://claude.ai/design/…` or a design project ID + file path).
2. Output location for the `.pptx` (default: current working dir).
3. Fidelity bar: "pixel-perfect" (default — full review loop) vs "close enough" (single pass).

Then confirm the plan with a short task list and proceed. Do **not** wait for further approval between phases — the user asked for autonomous iteration until perfect.

## Workflow (7 phases)

```
1 Import design  →  2 Extract assets  →  3 Ground-truth screenshots
      →  4 Build native PPTX  →  5 Render + compare  →  6 Auto-review loop  →  7 Validate + deliver
```

### Phase 1 — Import the design
- Use the **claude_design / DesignSync MCP** (`https://api.anthropic.com/v1/design/mcp`, auth via `/design-login`). Call `get_file` to pull the design's HTML/CSS/JS and asset manifest.
- `get_file` caps at ~192KB of source per call and auto-saves large results to a tool-results `.txt`. For multi-file designs, pull each file.
- Save everything under a working dir, e.g. `build/deck/` (HTML, CSS, JS, `assets/`, fonts).
- Determine the slide size from the design (a web deck is usually **1280×720 px**). Record it — all geometry math depends on it.

### Phase 2 — Extract exact assets (byte-exact, never re-encode)
- Pull every image/SVG/font referenced by the design. **Do not** hand-reproduce base64 — it corrupts (padding errors, broken PNGs). Decode byte-exact from the design payload / tool-results and validate each with PIL (`Image.open().verify()`).
- **Never hand-copy inline base64** into a file — the model re-emitting a long base64 string corrupts it (broken PNG). `get_file` auto-persists LARGE binaries to a tool-results `.txt` (decode that, byte-exact). For SMALL binaries and when the share link needs auth, open the design in the user's authenticated Chrome, read an `<img>.src` to get the signed `…claudeusercontent.com/…/serve/<path>?t=<token>` URL, and `curl` each asset straight to disk (the signed token works without cookies). Validate each with `PIL.Image.open().load()`. See GOTCHAS #21.
- If an org DLP policy blocks byte extraction from the browser, ask the user to drop the assets in a folder — respect the policy, don't circumvent it.
- Note the fonts. If the design uses a **paid font** (e.g. Avenir Next LT Pro), you cannot embed it (licensing). The deck will render correctly only where that font is installed (it usually is on the author's Mac). Keep the TTFs locally **only for measuring text width** during build (see line-breaks in GOTCHAS).

### Phase 3 — Render ground-truth (GT) screenshots
- Build a self-contained HTML harness that renders each slide/section at the deck size, and screenshot each at **2× device scale** (`--force-device-scale-factor=2`) with headless Chrome → one PNG per slide (`gt/slide-01.png` …). These are your source of truth for the review loop.
- **Load the design's real fonts in the harness** — `<link>` its `colors_and_type.css` (or replicate its `@font-face`) and put the `.ttf`s where that CSS points. Without this the browser cannot render font-weight **600/300** (they're separate installed families) and silently synthesises them, so your GT is wrong for every semibold element. See GOTCHAS #16.
- Alternatively screenshot the live design if reachable. GT PNGs should be 2× the slide px (e.g. 2560×1440).

### Phase 4 — Build the native PPTX
- Copy `reference/bslib.py` into your build dir — it's the reusable engine (units, colors, rect/oval/line/polygon, text with exact spacing + pre-wrapping, gradients, shadows, glows, rounded-corner masks, SVG embed/rasterize, picture cover-crop, logo rows). **Read its docstrings.**
- Write a `build.py` with one function per slide (`s01(prs)…sNN(prs)`), using `bslib`. Match every element: position, size, color, font size/weight, letter-spacing, line-height, gradients, shadows, rounded corners, images.
- **Units**: `1 px = 9525 EMU`, `1 px = 0.75 pt`. Deck 1280×720 px → 13.333"×7.5".
- **Font weight ≠ bold flag.** Read each element's `font-weight` from the design HTML and use `bslib.run(text, size, color, w=<weight>)` for EVERY run so 600 → the Demi family (not synthetic bold), 700 → bold, 400/300 → their families. Verify installed family names with fontTools first and set `DECK_FONT_DEMI`/`DECK_FONT_LIGHT`. This is the single most common fidelity miss — see GOTCHAS #15.
- **Text is the hard part** — see GOTCHAS "Line breaks" and "Text spacing". Pre-wrap every multi-line text box at its *true container width* so breaks match the original; for `text-wrap: pretty` emit **explicit per-line paragraphs** read off the GT (#20). Keep single-line labels `word_wrap=False` (#18).
- **Don't derive box heights/positions from CSS math alone** — sample the GT PNG for anything sized/positioned (card/bar/pill heights, panel top/left) and for semi-transparent overlay colors (#19).
- Attach speaker notes from the design if present.

### Phase 5 — Render the PPTX and compare
- Use `reference/render_compare.py` (LibreOffice headless → PDF → PNG via `pdftoppm`) to render your PPTX to per-slide PNGs and build side-by-side comparison strips (`out/cmp/cmp-XX.png` = GT over yours).
- **LibreOffice is only a preview.** It is lenient and renders things PowerPoint rejects, and it substitutes fonts. Use it for layout/geometry/color checks, not as proof of PowerPoint validity (Phase 7 handles that).
- Read the comparison strips and the GT PNGs directly (Read tool on the images) and fix discrepancies in `build.py`. Iterate.

### Phase 6 — Automated review loop (the core ask)
Loop until clean:
1. Rebuild + render all slides to fresh PNGs.
2. Spawn a **reviewer subagent** (general-purpose) with the prompt in `reference/review-agent-prompt.md`. Give it the GT PNGs and your rendered PNGs and have it return a structured, severity-ranked, per-slide deviation report with concrete fix instructions.
3. Apply the fixes. For anything ambiguous, prefer matching the GT exactly.
4. Repeat until the reviewer returns no High/Med findings (or the user's bar is met). For a thorough bar, run the reviewer 2–3× across iterations — later passes catch what earlier fixes shifted.
- The reviewer MUST be told to ignore font-glyph/kerning/sub-pixel differences (fallback font in LibreOffice) and only flag **structural** text differences (wrong wording, clearly wrong size/weight/color, different line-break count). Otherwise it drowns you in false positives.
- Explicitly instruct it to check **bottom-whitespace**: content that ends too high vs the original (a recurring issue), and to say whether to move the block down, add spacing, or enlarge proportionally.

### Phase 7 — Validate the file, then deliver
Before handing over, **validate the OOXML** (PowerPoint is far stricter than LibreOffice — a file that renders in LibreOffice can still be "defective" in PowerPoint):
- Every `ppt/slides/slideN.xml` is well-formed.
- Every `<a:custGeom>` wraps its `<a:path>` in `<a:pathLst>` (a stray `<a:path>` corrupts the file — see GOTCHAS).
- No gradient shape falls back to theme `accent1` (blue) — colored `srgbClr` stops present, not `schemeClr` in the gradFill (see GOTCHAS).
- Re-open with python-pptx as a smoke test.
Then copy to the delivery path and report. Link the file with a clickable markdown path.

## Preserving the user's manual edits (recurring)

Users will hand-edit the delivered `.pptx` in PowerPoint between rounds and expect their edits kept. Your build **regenerates from scratch**, so you must detect and bake them in first:
1. If a `~$<name>.pptx` lock file exists, the deck is **open in PowerPoint** — writing will conflict. Ask them to close it (without saving) first.
2. Diff the delivered file against your last build: text runs, shape geometry (left/top/width/height), fills, and font size/weight (`reference/diff-pptx.py` snippet in GOTCHAS). Ask which detected changes to keep if unsure.
3. Bake the confirmed changes into `build.py` (with a comment noting they're user edits), then rebuild.
4. When you ask a genuinely blocking question (e.g. how to preserve manual edits), do so — but otherwise iterate autonomously.

## Environment / dependencies
`python-pptx`, `Pillow`, LibreOffice (`soffice`), `pdftoppm` (poppler), `rsvg-convert` (librsvg) for SVG rasterization, headless Chrome for GT. Do the LibreOffice conversion from `/tmp` (see GOTCHAS: OneDrive path crash).

## Success criteria
- Native, editable shapes/text throughout (no flattened screenshots of the whole slide).
- Opens in PowerPoint with **no repair prompt**.
- Reviewer returns no High/Med deviations against the GT.
- User's manual edits preserved across rebuilds.
