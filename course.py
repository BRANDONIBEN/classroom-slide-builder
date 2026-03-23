#!/usr/bin/env python3
"""
Course management CLI for Classroom Slide Builder.

Usage:
  python course.py new "Course Name" --pdf "docs/Course.pdf" --id cn
  python course.py update bf --pdf "docs/Biblical Finances v2.pdf"
  python course.py images et                    # regenerate page images only
  python course.py push                         # push all changes to GitHub
  python course.py list                         # list all courses
"""

import sys
import os
import json
import subprocess
import argparse

try:
    import fitz  # PyMuPDF
except ImportError:
    print("ERROR: PyMuPDF not installed. Run: pip install pymupdf")
    sys.exit(1)

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(PROJECT_DIR, "data")
DOCS_DIR = os.path.join(PROJECT_DIR, "docs")
PAGES_DIR = os.path.join(PROJECT_DIR, "pages")


def extract_text(pdf_path):
    """Extract raw text from PDF, preserving page breaks."""
    doc = fitz.open(pdf_path)
    full_text = ""
    for i, page in enumerate(doc):
        text = page.get_text()
        if i > 0:
            full_text += f"\n{i + 1}\n"
        full_text += text
    doc.close()
    return full_text


def generate_images(pdf_path, course_id, dpi=150):
    """Export each PDF page as a PNG image."""
    out_dir = os.path.join(PAGES_DIR, course_id)
    os.makedirs(out_dir, exist_ok=True)

    doc = fitz.open(pdf_path)
    count = len(doc)

    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        out_path = os.path.join(out_dir, f"page_{i + 1}.png")
        pix.save(out_path)
        print(f"  [{i + 1}/{count}] {out_path}")

    doc.close()
    return count


def detect_sessions(text):
    """Parse session headers from extracted text."""
    import re
    sessions = []
    # Match patterns like "Session 1: Topic Name" or "Session 1 - Topic"
    pattern = r'Session\s+(\d+)[:\s]+([^\n]+)'
    for match in re.finditer(pattern, text):
        num = int(match.group(1))
        label = match.group(2).strip()
        # Clean up label — remove "Teacher:" lines
        label = re.split(r'\s*Teacher:', label)[0].strip()
        if label and not any(s['num'] == num for s in sessions):
            sessions.append({"num": num, "label": label})
    return sorted(sessions, key=lambda s: s['num'])


def update_index(course_id, course_name, sessions):
    """Update data/index.json with course info."""
    index_path = os.path.join(DATA_DIR, "index.json")

    if os.path.exists(index_path):
        with open(index_path) as f:
            index = json.load(f)
    else:
        index = {"courses": []}

    # Find or create course entry
    existing = next((c for c in index["courses"] if c["id"] == course_id), None)

    if existing:
        existing["name"] = course_name
        existing["sessions"] = sessions
        print(f"  Updated existing course '{course_id}' in index.json")
    else:
        index["courses"].append({
            "id": course_id,
            "name": course_name,
            "file": f"{course_id}_text.json",
            "sessions": sessions
        })
        print(f"  Added new course '{course_id}' to index.json")

    with open(index_path, "w") as f:
        json.dump(index, f, indent=2)


def update_page_map(course_id, page_count, sessions):
    """Create/update data/page_map.json."""
    pm_path = os.path.join(DATA_DIR, "page_map.json")

    if os.path.exists(pm_path):
        with open(pm_path) as f:
            page_map = json.load(f)
    else:
        page_map = {}

    # Estimate session start pages (evenly distributed if we can't detect)
    session_start = {}
    if sessions:
        pages_per_session = max(1, page_count // len(sessions))
        for i, s in enumerate(sessions):
            session_start[str(s["num"])] = 1 + (i * pages_per_session)

    page_map[course_id] = {
        "pageCount": page_count,
        "sessionStart": session_start
    }

    with open(pm_path, "w") as f:
        json.dump(page_map, f, indent=2)

    print(f"  Updated page_map.json ({page_count} pages)")


def cmd_new(args):
    """Create a new course from PDF."""
    course_id = args.id
    course_name = args.name
    pdf_path = os.path.abspath(args.pdf)

    if not os.path.exists(pdf_path):
        print(f"ERROR: PDF not found: {pdf_path}")
        sys.exit(1)

    text_path = os.path.join(DATA_DIR, f"{course_id}_text.json")
    if os.path.exists(text_path):
        print(f"ERROR: Course '{course_id}' already exists. Use 'update' instead.")
        sys.exit(1)

    print(f"\n=== Creating new course: {course_name} ({course_id}) ===\n")

    # 1. Extract text
    print("Extracting text from PDF...")
    text = extract_text(pdf_path)
    with open(text_path, "w") as f:
        json.dump(text, f)
    print(f"  Saved to {text_path} ({len(text)} chars)")

    # 2. Detect sessions
    print("Detecting sessions...")
    sessions = detect_sessions(text)
    print(f"  Found {len(sessions)} sessions")
    for s in sessions:
        print(f"    S{s['num']}: {s['label']}")

    # 3. Generate page images
    print("Generating page images...")
    page_count = generate_images(pdf_path, course_id)

    # 4. Update index.json
    print("Updating index.json...")
    update_index(course_id, course_name, sessions)

    # 5. Update page_map.json
    print("Updating page_map.json...")
    update_page_map(course_id, page_count, sessions)

    # 6. Copy PDF to docs/
    dest_pdf = os.path.join(DOCS_DIR, os.path.basename(pdf_path))
    if os.path.abspath(pdf_path) != os.path.abspath(dest_pdf):
        os.makedirs(DOCS_DIR, exist_ok=True)
        import shutil
        shutil.copy2(pdf_path, dest_pdf)
        print(f"  Copied PDF to {dest_pdf}")

    print(f"\n✓ Course '{course_name}' created successfully.")
    print(f"  Run 'python course.py push' to push to GitHub.\n")


def cmd_update(args):
    """Update an existing course with a new PDF."""
    course_id = args.id
    pdf_path = os.path.abspath(args.pdf)

    if not os.path.exists(pdf_path):
        print(f"ERROR: PDF not found: {pdf_path}")
        sys.exit(1)

    text_path = os.path.join(DATA_DIR, f"{course_id}_text.json")

    print(f"\n=== Updating course: {course_id} ===\n")

    # 1. Extract text
    print("Extracting text from PDF...")
    text = extract_text(pdf_path)

    # Show diff if existing
    if os.path.exists(text_path):
        with open(text_path) as f:
            old_text = json.load(f)
        old_lines = old_text.split('\n')
        new_lines = text.split('\n')
        print(f"  Old: {len(old_lines)} lines, {len(old_text)} chars")
        print(f"  New: {len(new_lines)} lines, {len(text)} chars")

    with open(text_path, "w") as f:
        json.dump(text, f)
    print(f"  Saved to {text_path}")

    # 2. Detect sessions
    print("Detecting sessions...")
    sessions = detect_sessions(text)
    print(f"  Found {len(sessions)} sessions")

    # 3. Regenerate page images
    print("Regenerating page images...")
    page_count = generate_images(pdf_path, course_id)

    # 4. Update index + page_map
    # Get course name from index
    index_path = os.path.join(DATA_DIR, "index.json")
    course_name = course_id.upper()
    if os.path.exists(index_path):
        with open(index_path) as f:
            index = json.load(f)
        existing = next((c for c in index["courses"] if c["id"] == course_id), None)
        if existing:
            course_name = existing["name"]

    print("Updating index.json...")
    update_index(course_id, course_name, sessions)

    print("Updating page_map.json...")
    update_page_map(course_id, page_count, sessions)

    # 5. Copy PDF to docs/
    dest_pdf = os.path.join(DOCS_DIR, os.path.basename(pdf_path))
    if os.path.abspath(pdf_path) != os.path.abspath(dest_pdf):
        import shutil
        shutil.copy2(pdf_path, dest_pdf)
        print(f"  Copied PDF to {dest_pdf}")

    print(f"\n✓ Course '{course_id}' updated successfully.")
    print(f"  Run 'python course.py push' to push to GitHub.\n")


def cmd_images(args):
    """Regenerate page images only."""
    course_id = args.id

    # Find PDF
    pdf_path = None
    for f in os.listdir(DOCS_DIR):
        if f.lower().endswith('.pdf') and course_id in f.lower():
            pdf_path = os.path.join(DOCS_DIR, f)
            break

    if not pdf_path:
        print(f"ERROR: No PDF found for course '{course_id}' in {DOCS_DIR}")
        sys.exit(1)

    print(f"\n=== Regenerating images for {course_id} from {os.path.basename(pdf_path)} ===\n")
    page_count = generate_images(pdf_path, course_id, dpi=args.dpi)
    print(f"\n✓ Generated {page_count} images.\n")


def cmd_push(args):
    """Push all changes to GitHub."""
    os.chdir(PROJECT_DIR)

    print("\n=== Pushing to GitHub ===\n")

    # Check for changes
    result = subprocess.run(["git", "status", "--porcelain"], capture_output=True, text=True)
    if not result.stdout.strip():
        print("No changes to push.")
        return

    print("Changes detected:")
    print(result.stdout)

    # Stage data + pages
    subprocess.run(["git", "add", "data/", "pages/"], check=True)

    # Commit
    msg = args.message or "Update course data and page images"
    subprocess.run(["git", "commit", "-m", msg], check=True)

    # Push
    subprocess.run(["git", "push"], check=True)

    print("\n✓ Pushed to GitHub.\n")


def cmd_list(args):
    """List all courses."""
    index_path = os.path.join(DATA_DIR, "index.json")
    if not os.path.exists(index_path):
        print("No courses found.")
        return

    with open(index_path) as f:
        index = json.load(f)

    print("\n=== Courses ===\n")
    for c in index["courses"]:
        text_path = os.path.join(DATA_DIR, c.get("file", ""))
        has_text = "✓" if os.path.exists(text_path) else "✗"
        pages_dir = os.path.join(PAGES_DIR, c["id"])
        page_count = len(os.listdir(pages_dir)) if os.path.isdir(pages_dir) else 0
        print(f"  {c['id']}  {c['name']}")
        print(f"       Sessions: {len(c.get('sessions', []))}")
        print(f"       Text: {has_text}  Pages: {page_count}")
        print()


def main():
    parser = argparse.ArgumentParser(description="Course management for Classroom Slide Builder")
    sub = parser.add_subparsers(dest="command")

    # new
    p_new = sub.add_parser("new", help="Create a new course from PDF")
    p_new.add_argument("name", help="Course name (e.g., 'Church History')")
    p_new.add_argument("--pdf", required=True, help="Path to PDF file")
    p_new.add_argument("--id", required=True, help="Short course ID (e.g., 'ch')")

    # update
    p_update = sub.add_parser("update", help="Update existing course with new PDF")
    p_update.add_argument("id", help="Course ID (e.g., 'bf')")
    p_update.add_argument("--pdf", required=True, help="Path to new PDF file")

    # images
    p_images = sub.add_parser("images", help="Regenerate page images only")
    p_images.add_argument("id", help="Course ID")
    p_images.add_argument("--dpi", type=int, default=150, help="Image DPI (default: 150)")

    # push
    p_push = sub.add_parser("push", help="Push changes to GitHub")
    p_push.add_argument("-m", "--message", help="Commit message")

    # list
    sub.add_parser("list", help="List all courses")

    args = parser.parse_args()

    if args.command == "new":
        cmd_new(args)
    elif args.command == "update":
        cmd_update(args)
    elif args.command == "images":
        cmd_images(args)
    elif args.command == "push":
        cmd_push(args)
    elif args.command == "list":
        cmd_list(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
