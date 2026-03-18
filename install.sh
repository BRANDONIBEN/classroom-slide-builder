#!/bin/bash
# Classroom Slide Builder — One-command install
# Run: curl -sL https://raw.githubusercontent.com/BRANDONIBEN/classroom-slide-builder/main/install.sh | bash

INSTALL_DIR="$HOME/Documents/classroom-slide-builder"

echo ""
echo "  ╔════════════════════════════════════════╗"
echo "  ║   Classroom Slide Builder — Install    ║"
echo "  ╚════════════════════════════════════════╝"
echo ""

# Check for git
if ! command -v git &> /dev/null; then
  echo "  ❌  Git is not installed."
  echo "  Install it from: https://git-scm.com/downloads"
  echo ""
  exit 1
fi

# Clone or update
if [ -d "$INSTALL_DIR" ]; then
  echo "  📂  Plugin already installed at $INSTALL_DIR"
  echo "  🔄  Pulling latest updates..."
  cd "$INSTALL_DIR" && git pull origin main
else
  echo "  📥  Cloning plugin to $INSTALL_DIR..."
  git clone https://github.com/BRANDONIBEN/classroom-slide-builder.git "$INSTALL_DIR"
fi

echo ""
echo "  ✅  Plugin installed!"
echo ""
echo "  Next steps:"
echo "  1. Open Figma"
echo "  2. Go to Plugins → Development → Import plugin from manifest..."
echo "  3. Select: $INSTALL_DIR/manifest.json"
echo "  4. The plugin will appear under Plugins → Development"
echo ""
echo "  To update later, double-click: $INSTALL_DIR/update.command"
echo ""
