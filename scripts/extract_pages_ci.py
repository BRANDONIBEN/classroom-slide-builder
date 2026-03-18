#!/usr/bin/env python3
"""
extract_pages_ci.py — CI-safe version of extract_pages.py.

Reads credentials from environment variables, checks sync-manifest.json
to determine which courses changed, downloads their PDFs from Dropbox,
renders each page as a 150 DPI PNG, and uploads page images to Dropbox.

Environment variables required:
    DROPBOX_REFRESH_TOKEN
    DROPBOX_APP_KEY
    DROPBOX_APP_SECRET
"""

import os
import sys
import json
import tempfile
import requests

try:
    import fitz  # pymupdf
except ImportError:
    print("ERROR: PyMuPDF (fitz) not installed. Run: pip install pymupdf")
    sys.exit(2)

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, 'data')
MANIFEST_PATH = os.path.join(DATA_DIR, 'sync-manifest.json')

DROPBOX_APP_PATH = '/Apps/Figma Classroom Slide Builder/pages'
DPI = 150  # Good balance of quality vs size


# ---------------------------------------------------------------------------
# Dropbox helpers
# ---------------------------------------------------------------------------

def get_token():
    """Get a fresh Dropbox access token using refresh token from env vars."""
    refresh_token = os.environ.get('DROPBOX_REFRESH_TOKEN')
    app_key = os.environ.get('DROPBOX_APP_KEY')
    app_secret = os.environ.get('DROPBOX_APP_SECRET')

    if not refresh_token or not app_key or not app_secret:
        print("ERROR: Missing Dropbox credentials.")
        print("Set DROPBOX_REFRESH_TOKEN, DROPBOX_APP_KEY, DROPBOX_APP_SECRET.")
        sys.exit(2)

    resp = requests.post('https://api.dropboxapi.com/oauth2/token', data={
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'client_id': app_key,
        'client_secret': app_secret,
    })
    resp.raise_for_status()
    return resp.json()['access_token']


def download_from_dropbox(token, dropbox_path):
    """Download a file from Dropbox. Returns bytes."""
    resp = requests.post(
        'https://content.dropboxapi.com/2/files/download',
        headers={
            'Authorization': f'Bearer {token}',
            'Dropbox-API-Arg': json.dumps({'path': dropbox_path}),
        },
    )
    if not resp.ok:
        print(f"  ERROR downloading {dropbox_path}: {resp.status_code} {resp.text[:200]}")
        return None
    return resp.content


def upload_to_dropbox(token, data, path):
    """Upload binary data to Dropbox."""
    resp = requests.post(
        'https://content.dropboxapi.com/2/files/upload',
        headers={
            'Authorization': f'Bearer {token}',
            'Dropbox-API-Arg': json.dumps({
                'path': path,
                'mode': 'overwrite',
                'autorename': False,
            }),
            'Content-Type': 'application/octet-stream',
        },
        data=data,
    )
    if not resp.ok:
        print(f"  ERROR uploading {path}: {resp.status_code} {resp.text[:200]}")
    return resp.ok


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("=== Extract Page Images (CI) ===")
    print()

    # 1. Load manifest to find changed courses
    if not os.path.exists(MANIFEST_PATH):
        print("No sync-manifest.json found. Run sync-dropbox.js first.")
        sys.exit(1)

    with open(MANIFEST_PATH, 'r') as f:
        manifest = json.load(f)

    if not manifest.get('files'):
        print("No files in manifest. Nothing to process.")
        sys.exit(0)

    # Collect changed courses — all entries in the manifest that have
    # a synced_at matching the lastSync time are considered "just changed"
    last_sync = manifest.get('lastSync')
    changed = []
    for dropbox_path, info in manifest['files'].items():
        if info.get('synced_at') == last_sync:
            changed.append({
                'dropbox_path': dropbox_path,
                'course_id': info['course_id'],
                'name': info['name'],
            })

    if not changed:
        print("No recently changed courses. Nothing to process.")
        sys.exit(0)

    print(f"Found {len(changed)} changed course(s) to process.")

    # 2. Get Dropbox token
    print("Refreshing Dropbox access token...")
    token = get_token()
    print("Token obtained.")

    # 3. Download page_map.json if it exists
    page_map = {}
    try:
        map_data = download_from_dropbox(token, '/Apps/Figma Classroom Slide Builder/page_map.json')
        if map_data:
            page_map = json.loads(map_data.decode('utf-8'))
            print("Loaded existing page_map.json from Dropbox.")
    except Exception as e:
        print(f"No existing page_map.json (starting fresh): {e}")

    # 4. Process each changed course
    with tempfile.TemporaryDirectory() as tmp_dir:
        for course_info in changed:
            course_id = course_info['course_id']
            dropbox_path = course_info['dropbox_path']
            pdf_name = course_info['name']

            print(f"\n=== {course_id}: {pdf_name} ===")

            # Download PDF
            print("  Downloading PDF...")
            pdf_data = download_from_dropbox(token, dropbox_path)
            if pdf_data is None:
                print(f"  SKIPPING {course_id}: download failed.")
                continue

            # Save to temp file
            tmp_path = os.path.join(tmp_dir, pdf_name)
            with open(tmp_path, 'wb') as f:
                f.write(pdf_data)

            print(f"  Downloaded ({len(pdf_data) // 1024} KB)")

            # Open with PyMuPDF
            doc = fitz.open(tmp_path)
            page_count = len(doc)
            print(f"  {page_count} pages")

            course_pages = {}

            for page_num in range(page_count):
                page = doc[page_num]

                # Render at DPI
                mat = fitz.Matrix(DPI / 72, DPI / 72)
                pix = page.get_pixmap(matrix=mat)
                png_data = pix.tobytes('png')

                # Upload path
                upload_path = f"{DROPBOX_APP_PATH}/{course_id}/page_{page_num + 1}.png"

                print(f"  Page {page_num + 1}/{page_count} ({len(png_data) // 1024}KB)...", end=' ')
                ok = upload_to_dropbox(token, png_data, upload_path)
                print('OK' if ok else 'FAILED')

                course_pages[str(page_num + 1)] = upload_path

            doc.close()

            page_map[course_id] = {
                'pageCount': page_count,
                'pages': course_pages,
            }

    # 5. Upload updated page_map.json
    map_json = json.dumps(page_map, indent=2)
    map_dropbox_path = '/Apps/Figma Classroom Slide Builder/page_map.json'
    print(f"\nUploading page_map.json ({len(page_map)} courses)...")
    upload_to_dropbox(token, map_json.encode(), map_dropbox_path)

    # 6. Also save a local copy in data/
    local_map_path = os.path.join(DATA_DIR, 'page_map.json')
    with open(local_map_path, 'w') as f:
        f.write(map_json)
    print(f"Saved local copy: {local_map_path}")

    print("\nDone!")


if __name__ == '__main__':
    main()
