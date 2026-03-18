#!/usr/bin/env python3
"""
Extract PDF pages as PNG images and upload to Dropbox.
Also generates a page_map.json mapping course → page number → image path.

Images go to /Apps/Figma Classroom Slide Builder/pages/{course_id}/page_{N}.png
"""

import os
import sys
import json
import requests
import fitz  # pymupdf

# Config
DOCS_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'docs')
DROPBOX_APP_PATH = '/Apps/Figma Classroom Slide Builder/pages'

COURSES = {
    'et': 'Essential Theology VOD Slides [Brandon Copy].pdf',
    'sn': 'Scripture Narrative VOD Slides.pdf',
    'bf': 'Biblical Finances VOD Slides.pdf',
}

DPI = 150  # Good balance of quality vs size

def get_token():
    """Get a fresh Dropbox access token using refresh token."""
    resp = requests.post('https://api.dropboxapi.com/oauth2/token', data={
        'grant_type': 'refresh_token',
        'refresh_token': '0chv99flxfIAAAAAAAAAAQAuBggMH_hxGj1Gj9Vh_N_ZN26a554Ef2s7WDirjRa9',
        'client_id': 'xni3vzybb0pnpea',
        'client_secret': 'jim2za1dphztq8n',
    })
    resp.raise_for_status()
    return resp.json()['access_token']

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
    if resp.status_code == 409:
        print(f'  Conflict uploading {path}, retrying with overwrite...')
    elif not resp.ok:
        print(f'  Error uploading {path}: {resp.status_code} {resp.text[:200]}')
    return resp.ok

def main():
    token = get_token()
    print(f'Got Dropbox token')

    page_map = {}

    for course_id, pdf_name in COURSES.items():
        pdf_path = os.path.join(DOCS_DIR, pdf_name)
        if not os.path.exists(pdf_path):
            print(f'WARNING: {pdf_path} not found, skipping')
            continue

        print(f'\n=== {course_id}: {pdf_name} ===')
        doc = fitz.open(pdf_path)
        page_count = len(doc)
        print(f'  {page_count} pages')

        course_pages = {}

        for page_num in range(page_count):
            page = doc[page_num]
            # Render at DPI
            mat = fitz.Matrix(DPI / 72, DPI / 72)
            pix = page.get_pixmap(matrix=mat)
            png_data = pix.tobytes('png')

            # Upload path
            dropbox_path = f'{DROPBOX_APP_PATH}/{course_id}/page_{page_num + 1}.png'

            print(f'  Page {page_num + 1}/{page_count} ({len(png_data) // 1024}KB)...', end=' ')
            ok = upload_to_dropbox(token, png_data, dropbox_path)
            print('OK' if ok else 'FAILED')

            course_pages[str(page_num + 1)] = dropbox_path

        doc.close()
        page_map[course_id] = {
            'pageCount': page_count,
            'pages': course_pages,
        }

    # Upload page_map.json
    map_json = json.dumps(page_map, indent=2)
    map_path = '/Apps/Figma Classroom Slide Builder/page_map.json'
    print(f'\nUploading page_map.json...')
    upload_to_dropbox(token, map_json.encode(), map_path)

    # Also save locally
    local_map = os.path.join(os.path.dirname(__file__), 'page_map.json')
    with open(local_map, 'w') as f:
        f.write(map_json)
    print(f'Saved local copy: {local_map}')

    print('\nDone!')

if __name__ == '__main__':
    main()
