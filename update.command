#!/bin/bash
# Double-click this file to update the Classroom Slide Builder plugin
cd "$(dirname "$0")"
echo ""
echo "  🔄  Updating Classroom Slide Builder..."
echo ""
git pull origin main
echo ""
echo "  ✅  Updated! Reload the plugin in Figma."
echo "  (Close and reopen the plugin, or restart Figma)"
echo ""
read -p "  Press Enter to close..."
