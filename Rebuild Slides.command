#!/bin/bash
cd "$(dirname "$0")"
echo "Rebuilding VOD Slide Builder data from docs/ PDFs..."
echo ""
python3 rebuild.py
echo ""
echo "Press any key to close."
read -n 1
