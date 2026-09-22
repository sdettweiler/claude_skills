#!/usr/bin/env bash
# setup.sh — one-shot dependency setup for the generate-ppt-from-design skill (macOS).
# Safe to re-run: installs ONLY what's missing.
#
# Note: if Homebrew isn't installed yet, its official installer asks for your Mac
# password once (it needs admin rights). If this script is run in a NON-interactive
# shell (e.g. by Claude Code), it can't answer that prompt, so it skips the Homebrew
# bootstrap and prints the exact command for you to paste, then everything else works.
set -u
BOLD=$(tput bold 2>/dev/null || true); RST=$(tput sgr0 2>/dev/null || true)
ok(){   echo "  ✓ $1"; }
add(){  echo "  → installing $1"; }
warn(){ echo "  ! $1"; }
have(){ command -v "$1" >/dev/null 2>&1; }
interactive(){ [ -t 0 ]; }
MANUAL=()

echo "${BOLD}generate-ppt-from-design — dependency setup${RST}"
UNAME=$(uname -s)
if [ "$UNAME" != "Darwin" ]; then
  warn "This installer targets macOS. On Debian/Ubuntu run:"
  echo "     sudo apt-get install -y python3-pip libreoffice poppler-utils librsvg2-bin"
  echo "     python3 -m pip install --user python-pptx Pillow   # + install Google Chrome"
fi

# ---- Python + pip packages ----
if have python3; then ok "python3 ($(python3 -V 2>&1))"; else
  warn "python3 not found"; MANUAL+=("Install Python 3 (macOS: 'brew install python' after Homebrew)")
fi
if have python3; then
  if python3 -c 'import pptx, PIL' >/dev/null 2>&1; then ok "python-pptx, Pillow (already installed)"
  else
    python3 -m pip --version >/dev/null 2>&1 || python3 -m ensurepip --upgrade >/dev/null 2>&1 || true
    add "python-pptx, Pillow"
    # try plain --user, then --break-system-packages (needed on PEP-668 'externally managed' Python)
    if python3 -m pip install --user --upgrade python-pptx Pillow >/dev/null 2>&1 \
       || python3 -m pip install --user --break-system-packages --upgrade python-pptx Pillow >/dev/null 2>&1; then
      ok "python-pptx, Pillow"
    else
      warn "pip couldn't install into this Python. Use a venv:"
      echo "     python3 -m venv ~/.venvs/ppt && ~/.venvs/ppt/bin/pip install python-pptx Pillow"
      echo "     (then run the skill's build scripts with ~/.venvs/ppt/bin/python)"
    fi
  fi
fi

# ---- Homebrew ----
if have brew; then ok "Homebrew"; else
  if [ "$UNAME" = "Darwin" ]; then
    if interactive; then
      add "Homebrew (you'll be asked for your Mac password once)"
      /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    else
      warn "Homebrew missing and this is a non-interactive shell — skipping its bootstrap."
      MANUAL+=('Install Homebrew (needs your password): /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"')
    fi
  fi
fi
# make brew available for the rest of this run if it was just installed
have brew || for p in /opt/homebrew/bin/brew /usr/local/bin/brew; do [ -x "$p" ] && eval "$("$p" shellenv)"; done

# ---- system tools via Homebrew ----
if have brew; then
  have pdftoppm      && ok "poppler (pdftoppm)" || { add "poppler";  brew install poppler        >/dev/null 2>&1 && ok "poppler"  || warn "poppler install failed"; }
  have rsvg-convert  && ok "librsvg (rsvg-convert)" || { add "librsvg"; brew install librsvg      >/dev/null 2>&1 && ok "librsvg"  || warn "librsvg install failed"; }
  if have soffice || [ -d "/Applications/LibreOffice.app" ]; then ok "LibreOffice"
  else add "LibreOffice"; brew install --cask libreoffice >/dev/null 2>&1 && ok "LibreOffice" || warn "LibreOffice install failed"; fi
  if [ -d "/Applications/Google Chrome.app" ] || have google-chrome; then ok "Google Chrome"
  else add "Google Chrome"; brew install --cask google-chrome >/dev/null 2>&1 && ok "Google Chrome" || warn "Chrome install failed"; fi
elif [ "$UNAME" = "Darwin" ]; then
  warn "Skipping LibreOffice / poppler / librsvg / Chrome — Homebrew needed first (see below)."
fi

# ---- manual steps ----
echo
echo "${BOLD}Steps that can't be automated:${RST}"
echo "  • Run  /design-login  once in Claude Code (to import designs)."
echo "  • Install the deck's font locally if you need PowerPoint to render the type exactly"
echo "    (paid fonts aren't embedded); then point DECK_FONT / DECK_FONTS_DIR at it."
for m in "${MANUAL[@]:-}"; do [ -n "${m:-}" ] && echo "  • $m"; done
echo
echo "Then start a NEW Claude Code session and run:  ${BOLD}/generate-ppt-from-design${RST}"
