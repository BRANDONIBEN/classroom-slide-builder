# ET Slide Builder — Figma Plugin

Builds Essential Theology slide frames directly in a Figma Design file
from pasted Google Doc text. No PPTX, no import, no middleman.

---

## Setup (one time)

1. Open Figma Desktop (plugin dev requires the desktop app)
2. Menu → Plugins → Development → New Plugin
3. Choose "Link existing plugin"
4. Select the `manifest.json` file from this folder
5. Done — plugin is now available under Plugins → Development → ET Slide Builder

---

## Usage (per session)

1. Open your Figma Design file
2. Navigate to the page where you want the session slides
3. Run: Plugins → Development → ET Slide Builder
4. Paste your full Google Doc text into the text area
   (you can paste the entire doc — the plugin filters by session)
5. Select the session number from the dropdown
6. Check the preview shows the right slide count
7. Click Build Slides
8. Frames appear on the canvas, ready to style

---

## Output

Each frame is named:
  [TYPE] S1 · 6 — What Is Theology?

Types: STATEMENT, QUOTE, SCRIPTURE, LIST, BODY

Frame size: 1920 × 1080px
Background: #1E1E1A (dark)

Text layers inside each frame are named and editable.
Apply your component styles / fonts in the polish pass.

---

## Fonts

The plugin uses Georgia as a placeholder font.
After building, use Edit → Find/Replace or a bulk font plugin
(e.g. "Font Replacer") to swap Georgia → your actual typeface across
all frames at once.

---

## Files

  manifest.json   Plugin config
  code.js         Main plugin logic (runs in Figma sandbox)
  ui.html         Plugin UI (paste text, select session, build)
