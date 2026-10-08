# claude_skills

A collection of [Claude Code](https://claude.com/claude-code) skills.

## Skills

### `generate-ppt-from-design`
Turn a [Claude design](https://claude.ai/design) into a **pixel-perfect, fully-editable native PowerPoint deck** — real shapes, text, gradients and images, not screenshots on slides. It imports the design, extracts exact assets, renders ground-truth screenshots, builds the PPTX with `python-pptx`, then runs automated review loops and iterates until it matches the original.

Invoke with `/generate-ppt-from-design`.

Bundled with the skill: the reusable `bslib.py` build engine, a `GOTCHAS.md` of hard-won PowerPoint/OOXML traps, a reviewer-agent prompt template, and helper scripts for rendering/diffing/validating.

### `remove-ai-slop`
Scan a piece of text for AI writing tells (word choice, structure, punctuation, tone), report what it found with quoted examples, then rewrite it to read like a person wrote it. Facts stay as they were. Based on Wikipedia's "Signs of AI writing" guide.

Invoke with `/remove-AI-slop`, or ask Claude to "humanize this" / "check if this sounds AI-written".

Bundled with the skill: `references/tells.md` (the full taxonomy the scan checks against) and `references/system-prompt.md`, a drop-in system prompt / CLAUDE.md block that stops the tells at the source.

## Quick install (easiest — no technical skills needed)

Open **Claude Code** and paste:

> Install the `generate-ppt-from-design` skill from https://github.com/sdettweiler/claude_skills — clone it, copy the skill into my `~/.claude/skills/`, then run its `setup.sh` to install any missing dependencies.

Claude clones the repo, installs the skill, and runs `setup.sh` to install the Python + system tools that are missing. Then start a **new** session and run `/generate-ppt-from-design`.

Two things can't be automated (Claude will flag them): running **`/design-login`** once, and installing the deck's font. If Homebrew isn't already on the machine, its installer needs your Mac password once — Claude will hand you that one line to paste.

### Or install the deps yourself
```bash
git clone https://github.com/sdettweiler/claude_skills.git
cp -R claude_skills/.claude/skills/generate-ppt-from-design ~/.claude/skills/
bash claude_skills/setup.sh      # installs missing deps (Homebrew asks for your password once)
```

## Installing a skill (manually)

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
