# claude_skills

A collection of [Claude Code](https://claude.com/claude-code) skills.

## Skills

### `generate-ppt-from-design`
Turn a [Claude design](https://claude.ai/design) into a **pixel-perfect, fully-editable native PowerPoint deck** — real shapes, text, gradients and images, not screenshots on slides. It imports the design, extracts exact assets, renders ground-truth screenshots, builds the PPTX with `python-pptx`, then runs automated review loops and iterates until it matches the original.

Invoke with `/generate-ppt-from-design`.

Bundled with the skill: the reusable `bslib.py` build engine, a `GOTCHAS.md` of hard-won PowerPoint/OOXML traps, a reviewer-agent prompt template, and helper scripts for rendering/diffing/validating.

## Installing a skill

**Per-user (all your projects):**
```bash
git clone https://github.com/sdettweiler/claude_skills.git
cp -R claude_skills/.claude/skills/<skill-name> ~/.claude/skills/
```
The skill is available next time you start Claude Code.

**Per-project (share with a team via your repo):**
Copy the skill folder into your project's `.claude/skills/<skill-name>/` and commit it. Anyone who clones the project gets it automatically.

## Prerequisites for `generate-ppt-from-design`
The skill orchestrates external tools, so a copy alone isn't self-contained. You also need:
- The **claude_design / DesignSync MCP** configured, and to run `/design-login` (to import designs).
- System tools: `python-pptx`, `Pillow`, **LibreOffice** (`soffice`), `pdftoppm` (poppler), `rsvg-convert` (librsvg), and headless Chrome.
- The design's font installed locally if you want PowerPoint to render type exactly (paid fonts are intentionally **not** embedded). Point `DECK_FONT` / `DECK_FONTS_DIR` at it for accurate line-break measurement.
