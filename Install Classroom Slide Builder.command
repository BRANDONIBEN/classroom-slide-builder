#!/bin/bash
# ============================================================
# Classroom Slide Builder — Easy Installer
# Double-click this file to install the Figma plugin
# ============================================================

INSTALL_DIR="$HOME/Documents/Classroom Slide Builder"
REPO_ZIP="https://github.com/BRANDONIBEN/classroom-slide-builder/archive/refs/heads/main.zip"
TMP_ZIP="/tmp/classroom-slide-builder.zip"

clear
echo ""
echo "  ┌─────────────────────────────────────────────┐"
echo "  │                                             │"
echo "  │     Classroom Slide Builder — Installer     │"
echo "  │     Passion Equip                           │"
echo "  │                                             │"
echo "  └─────────────────────────────────────────────┘"
echo ""

# Check if already installed
if [ -d "$INSTALL_DIR" ]; then
  echo "  📂  Plugin already installed."
  echo "  🔄  Updating to latest version..."
  echo ""
  rm -rf "$INSTALL_DIR"
fi

# Download ZIP (no git required)
echo "  📥  Downloading plugin..."
curl -sL "$REPO_ZIP" -o "$TMP_ZIP"

if [ ! -f "$TMP_ZIP" ]; then
  echo ""
  echo "  ❌  Download failed. Check your internet connection."
  echo ""
  read -p "  Press Enter to close..."
  exit 1
fi

# Extract
echo "  📦  Installing..."
mkdir -p "$INSTALL_DIR"
cd /tmp
unzip -qo "$TMP_ZIP"
cp -R /tmp/classroom-slide-builder-main/* "$INSTALL_DIR/"
rm -rf /tmp/classroom-slide-builder-main "$TMP_ZIP"

# Create update script
cat > "$INSTALL_DIR/Update Plugin.command" << 'UPDATER'
#!/bin/bash
INSTALL_DIR="$HOME/Documents/Classroom Slide Builder"
REPO_ZIP="https://github.com/BRANDONIBEN/classroom-slide-builder/archive/refs/heads/main.zip"
TMP_ZIP="/tmp/classroom-slide-builder.zip"
clear
echo ""
echo "  🔄  Updating Classroom Slide Builder..."
echo ""
curl -sL "$REPO_ZIP" -o "$TMP_ZIP"
cd /tmp && unzip -qo "$TMP_ZIP"
cp -R /tmp/classroom-slide-builder-main/* "$INSTALL_DIR/"
rm -rf /tmp/classroom-slide-builder-main "$TMP_ZIP"
echo "  ✅  Updated to latest version!"
echo ""
echo "  Close and reopen the plugin in Figma to use the new version."
echo ""
read -p "  Press Enter to close..."
UPDATER
chmod +x "$INSTALL_DIR/Update Plugin.command"

echo ""
echo "  ✅  Plugin installed to:"
echo "      ~/Documents/Classroom Slide Builder/"
echo ""
echo "  ─────────────────────────────────────────────"
echo ""
echo "  📋  NEXT STEP — Import into Figma:"
echo ""
echo "  1. Open the Figma desktop app"
echo "  2. Click: Plugins → Development → Import plugin from manifest..."
echo "  3. Navigate to: Documents → Classroom Slide Builder"
echo "  4. Select the file: manifest.json"
echo "  5. Done! Find it under Plugins → Development"
echo ""
echo "  🔑  Password: Ask Brandon for the team password"
echo ""
echo "  🔄  To update later: double-click 'Update Plugin.command'"
echo "      in the Classroom Slide Builder folder"
echo ""
echo "  ─────────────────────────────────────────────"
echo ""

# Open the plugin folder
open "$INSTALL_DIR"

# Try to open Figma
if [ -d "/Applications/Figma.app" ]; then
  echo "  🚀  Opening Figma..."
  echo "      In Figma: Plugins → Development → Import plugin from manifest..."
  echo "      Navigate to: ~/Documents/Classroom Slide Builder/"
  echo "      Select: manifest.json"
  echo ""
  open -a "Figma"
else
  echo "  💡  Open Figma manually, then:"
  echo "      Plugins → Development → Import plugin from manifest..."
  echo "      Navigate to: ~/Documents/Classroom Slide Builder/"
  echo ""
fi

read -p "  Press Enter to close this window..."
