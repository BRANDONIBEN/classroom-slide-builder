#!/usr/bin/env python3
"""
Rebuild VOD Slide Builder plugin data from PDFs in docs/ folder.

Usage:
  python3 rebuild.py

Scans docs/ for PDF files, extracts text, and updates the embedded
COURSE_TEXT data in ui.html. Also detects new courses and adds them
to the COURSES metadata.
"""

import json
import os
import re
import sys

try:
    import pdfplumber
except ImportError:
    print("Installing pdfplumber...")
    os.system(f"{sys.executable} -m pip install pdfplumber -q")
    import pdfplumber

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS_DIR = os.path.join(SCRIPT_DIR, "docs")
UI_HTML = os.path.join(SCRIPT_DIR, "ui.html")

# Known course mappings: PDF name pattern → course key
# Add new courses here or they'll be auto-detected
KNOWN_COURSES = {
    "Essential Theology": "et",
    "Scripture Narrative": "sn",
    "Biblical Finances": "bf",
}


def extract_pdf_text(pdf_path):
    """Extract all text from a PDF file."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text


def clean_extracted_text(text):
    """Remove page numbers, footnotes, and other artifacts."""
    lines = text.split("\n")
    cleaned = []
    skip_next = False
    for i, line in enumerate(lines):
        stripped = line.strip()
        # Skip standalone page numbers
        if re.match(r"^\d{1,3}$", stripped):
            continue
        # Skip footnote lines ("1 We want...", "2 We would...")
        if re.match(r"^\d+\s+(We want|We would)", stripped):
            skip_next = True
            continue
        if skip_next and stripped and not re.match(r"^(Slide|Session|\d+[:.]).", stripped):
            skip_next = False
            continue
        skip_next = False
        cleaned.append(line)
    return "\n".join(cleaned)


def guess_course_key(filename):
    """Generate a short course key from PDF filename."""
    name = os.path.splitext(os.path.basename(filename))[0]
    # Remove common suffixes
    name = re.sub(r"\s*(VOD\s*)?Slides?(\s*\[.*\])?$", "", name, flags=re.I)
    name = name.strip()
    # Check known mappings
    for known_name, key in KNOWN_COURSES.items():
        if known_name.lower() in name.lower():
            return key, known_name
    # Auto-generate: take first letter of each word
    words = name.split()
    key = "".join(w[0].lower() for w in words if w[0].isalpha())[:3]
    return key, name


def detect_sessions(text):
    """Parse session names from extracted text."""
    sessions = []
    seen = set()
    for m in re.finditer(r"Session\s+(\d+)\s*[:—\-–]\s*(.+)", text):
        num = int(m.group(1))
        label = m.group(2).split("(")[0].strip()
        # Clean version markers
        label = re.sub(r"\s*[-–—]\s*V\d+\s*$", "", label)
        label = re.sub(r"\s+V\d+\s*$", "", label)
        label = re.sub(r"\bPt\s*(\d)", r"Part \1", label)
        if num not in seen:
            sessions.append({"num": num, "label": label})
            seen.add(num)
    sessions.sort(key=lambda s: s["num"])
    return sessions


def rebuild():
    """Main rebuild function."""
    if not os.path.isdir(DOCS_DIR):
        print(f"Error: docs/ folder not found at {DOCS_DIR}")
        sys.exit(1)

    # Find all PDFs
    pdfs = sorted(
        f for f in os.listdir(DOCS_DIR) if f.lower().endswith(".pdf")
    )
    if not pdfs:
        print("No PDF files found in docs/")
        sys.exit(1)

    print(f"Found {len(pdfs)} PDF(s):")

    course_texts = {}
    course_meta = {}

    for pdf_file in pdfs:
        pdf_path = os.path.join(DOCS_DIR, pdf_file)
        key, name = guess_course_key(pdf_file)
        print(f"  [{key}] {pdf_file} → {name}")

        # Extract text
        raw_text = extract_pdf_text(pdf_path)
        cleaned = clean_extracted_text(raw_text)

        # Save intermediate JSON
        json_path = os.path.join(DOCS_DIR, f"{key}_text.json")
        with open(json_path, "w") as f:
            json.dump(cleaned, f)

        course_texts[key] = cleaned

        # Detect sessions
        sessions = detect_sessions(cleaned)
        course_meta[key] = {"name": name, "sessions": sessions}
        print(f"    → {len(sessions)} sessions detected")

    # Read current ui.html
    with open(UI_HTML, "r") as f:
        html = f.read()

    # Replace COURSE_TEXT
    course_text_json = json.dumps(course_texts, ensure_ascii=False)
    html = re.sub(
        r"var COURSE_TEXT\s*=\s*\{.*?\};",
        f"var COURSE_TEXT = {course_text_json};",
        html,
        count=1,
        flags=re.DOTALL,
    )

    # Replace COURSES metadata
    meta_lines = ["var COURSES = {"]
    for key, meta in course_meta.items():
        meta_lines.append(f"  {key}: {{")
        meta_lines.append(f"    name: '{meta['name']}',")
        meta_lines.append(f"    sessions: [")
        for s in meta["sessions"]:
            label = s["label"].replace("'", "\\'")
            meta_lines.append(
                f"      {{ num: {s['num']},  label: '{label}' }},"
            )
        meta_lines.append(f"    ]")
        meta_lines.append(f"  }},")
    meta_lines.append("};")
    meta_js = "\n".join(meta_lines)

    html = re.sub(
        r"var COURSES\s*=\s*\{.*?\n\};",
        meta_js,
        html,
        count=1,
        flags=re.DOTALL,
    )

    # Write updated ui.html
    with open(UI_HTML, "w") as f:
        f.write(html)

    print(f"\nDone! Updated ui.html with {len(course_texts)} course(s).")
    print("Reload the plugin in Figma to see changes.")


if __name__ == "__main__":
    rebuild()
