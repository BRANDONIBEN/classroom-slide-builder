#!/usr/bin/env node
/**
 * sync-dropbox.js — Automated PDF-to-JSON pipeline.
 *
 * Watches the Dropbox Doc Watch folder for new/changed PDFs,
 * converts them to JSON text, detects sessions, and updates
 * data/index.json + data/sync-manifest.json.
 *
 * Environment variables required:
 *   DROPBOX_REFRESH_TOKEN
 *   DROPBOX_APP_KEY
 *   DROPBOX_APP_SECRET
 *
 * Exit codes:
 *   0 — changes were made
 *   1 — no changes (nothing new to sync)
 */

const fs = require('fs');
const path = require('path');
const https = require('https');
const { URL } = require('url');

// pdf-parse is loaded lazily after we confirm there's work to do
let pdfParse;

// ---------------------------------------------------------------------------
// Paths
// ---------------------------------------------------------------------------
const ROOT = path.resolve(__dirname, '..');
const DATA_DIR = path.join(ROOT, 'data');
const MANIFEST_PATH = path.join(DATA_DIR, 'sync-manifest.json');
const INDEX_PATH = path.join(DATA_DIR, 'index.json');
const TMP_DIR = path.join(ROOT, '.tmp-sync');

const DROPBOX_WATCH_FOLDER = '/Passion Equip Shared/Content/07 - Classroom/00 Doc Watch';

// ---------------------------------------------------------------------------
// Known course name → id mappings
// ---------------------------------------------------------------------------
const KNOWN_COURSES = {
  'essential theology': { id: 'et', name: 'Essential Theology' },
  'scripture narrative': { id: 'sn', name: 'Scripture Narrative' },
  'biblical finances':   { id: 'bf', name: 'Biblical Finances' },
};

// ---------------------------------------------------------------------------
// Helpers — HTTP
// ---------------------------------------------------------------------------

/**
 * Make an HTTPS request and return { statusCode, headers, body }.
 */
function httpRequest(url, options, body) {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const reqOpts = {
      hostname: parsed.hostname,
      path: parsed.pathname + parsed.search,
      method: options.method || 'POST',
      headers: options.headers || {},
    };

    const req = https.request(reqOpts, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        resolve({
          statusCode: res.statusCode,
          headers: res.headers,
          body: Buffer.concat(chunks),
        });
      });
    });

    req.on('error', reject);

    if (body) {
      req.write(body);
    }
    req.end();
  });
}

// ---------------------------------------------------------------------------
// Dropbox API helpers
// ---------------------------------------------------------------------------

async function refreshAccessToken() {
  const refreshToken = process.env.DROPBOX_REFRESH_TOKEN;
  const appKey = process.env.DROPBOX_APP_KEY;
  const appSecret = process.env.DROPBOX_APP_SECRET;

  if (!refreshToken || !appKey || !appSecret) {
    throw new Error(
      'Missing Dropbox credentials. Set DROPBOX_REFRESH_TOKEN, DROPBOX_APP_KEY, DROPBOX_APP_SECRET.'
    );
  }

  const params = new URLSearchParams({
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    client_id: appKey,
    client_secret: appSecret,
  });

  const res = await httpRequest('https://api.dropboxapi.com/oauth2/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  }, params.toString());

  if (res.statusCode !== 200) {
    throw new Error(`Failed to refresh Dropbox token: ${res.statusCode} ${res.body.toString()}`);
  }

  const data = JSON.parse(res.body.toString());
  return data.access_token;
}

/**
 * List all files in a Dropbox folder. Handles pagination via cursor.
 */
async function listFolder(token, folderPath) {
  const entries = [];

  // Initial request
  let res = await httpRequest('https://api.dropboxapi.com/2/files/list_folder', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
  }, JSON.stringify({
    path: folderPath,
    recursive: false,
    include_deleted: false,
  }));

  if (res.statusCode !== 200) {
    throw new Error(`list_folder failed: ${res.statusCode} ${res.body.toString()}`);
  }

  let data = JSON.parse(res.body.toString());
  entries.push(...data.entries);

  // Paginate
  while (data.has_more) {
    res = await httpRequest('https://api.dropboxapi.com/2/files/list_folder/continue', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
    }, JSON.stringify({ cursor: data.cursor }));

    if (res.statusCode !== 200) {
      throw new Error(`list_folder/continue failed: ${res.statusCode} ${res.body.toString()}`);
    }

    data = JSON.parse(res.body.toString());
    entries.push(...data.entries);
  }

  return entries;
}

/**
 * Download a file from Dropbox. Returns a Buffer.
 */
async function downloadFile(token, dropboxPath) {
  const res = await httpRequest('https://content.dropboxapi.com/2/files/download', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Dropbox-API-Arg': JSON.stringify({ path: dropboxPath }),
    },
  }, null);

  if (res.statusCode !== 200) {
    throw new Error(`download failed for ${dropboxPath}: ${res.statusCode} ${res.body.toString().slice(0, 300)}`);
  }

  return res.body;
}

// ---------------------------------------------------------------------------
// Course detection
// ---------------------------------------------------------------------------

/**
 * Detect course id and display name from a PDF filename.
 */
function detectCourse(filename) {
  const lower = filename.toLowerCase();

  for (const [keyword, info] of Object.entries(KNOWN_COURSES)) {
    if (lower.includes(keyword)) {
      return { id: info.id, name: info.name };
    }
  }

  // Fallback: first two words of the filename (minus extension)
  const base = path.basename(filename, path.extname(filename));
  const words = base.replace(/[^a-zA-Z0-9\s]/g, '').trim().split(/\s+/);
  const id = words.slice(0, 2).join('').toLowerCase().slice(0, 8);
  const name = words.slice(0, 4).join(' ');
  return { id, name };
}

// ---------------------------------------------------------------------------
// Session detection (matches add-course.js patterns)
// ---------------------------------------------------------------------------

/**
 * Detect sessions from extracted PDF text.
 * Supports "Session N: Label" and "Class N: Label" patterns.
 */
function detectSessions(text) {
  const sessions = [];
  const seen = {};

  // Pattern 1: "Session N: Label" (from add-course.js)
  const sessionPattern = /Session\s+(\d+):\s*(.+?)(?:\s*\(|$)/gm;
  let match;
  while ((match = sessionPattern.exec(text)) !== null) {
    const num = parseInt(match[1]);
    const label = match[2].trim();
    if (!seen[num]) {
      seen[num] = true;
      sessions.push({ num, label });
    }
  }

  // Pattern 2: "Class N: Label"
  if (sessions.length === 0) {
    const classPattern = /Class\s+(\d+):\s*(.+?)(?:\s*\(|$)/gm;
    while ((match = classPattern.exec(text)) !== null) {
      const num = parseInt(match[1]);
      const label = match[2].trim();
      if (!seen[num]) {
        seen[num] = true;
        sessions.push({ num, label });
      }
    }
  }

  if (sessions.length === 0) {
    console.warn('    Warning: Could not auto-detect sessions.');
    sessions.push({ num: 1, label: 'Session 1' });
  } else {
    sessions.sort((a, b) => a.num - b.num);
  }

  return sessions;
}

// ---------------------------------------------------------------------------
// Manifest helpers
// ---------------------------------------------------------------------------

function loadManifest() {
  if (fs.existsSync(MANIFEST_PATH)) {
    return JSON.parse(fs.readFileSync(MANIFEST_PATH, 'utf8'));
  }
  return { lastSync: null, files: {} };
}

function saveManifest(manifest) {
  manifest.lastSync = new Date().toISOString();
  fs.writeFileSync(MANIFEST_PATH, JSON.stringify(manifest, null, 2) + '\n');
}

function loadIndex() {
  if (fs.existsSync(INDEX_PATH)) {
    return JSON.parse(fs.readFileSync(INDEX_PATH, 'utf8'));
  }
  return { courses: [] };
}

function saveIndex(index) {
  fs.writeFileSync(INDEX_PATH, JSON.stringify(index, null, 2) + '\n');
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  console.log('=== Classroom Slide Builder: Dropbox Sync ===');
  console.log(`Watch folder: ${DROPBOX_WATCH_FOLDER}`);
  console.log();

  // 1. Refresh Dropbox token
  console.log('Refreshing Dropbox access token...');
  const token = await refreshAccessToken();
  console.log('Token obtained.');

  // 2. List PDFs in the watch folder
  console.log(`Listing files in watch folder...`);
  const entries = await listFolder(token, DROPBOX_WATCH_FOLDER);
  const pdfEntries = entries.filter(
    (e) => e['.tag'] === 'file' && e.name.toLowerCase().endsWith('.pdf')
  );
  console.log(`Found ${pdfEntries.length} PDF file(s).`);

  if (pdfEntries.length === 0) {
    console.log('No PDFs found. Nothing to do.');
    process.exit(1);
  }

  // 3. Load manifest and check for changes
  const manifest = loadManifest();
  const changedFiles = [];

  for (const entry of pdfEntries) {
    const prevHash = manifest.files[entry.path_lower]
      ? manifest.files[entry.path_lower].content_hash
      : null;
    if (prevHash !== entry.content_hash) {
      changedFiles.push(entry);
    }
  }

  if (changedFiles.length === 0) {
    console.log('All files up to date. No changes needed.');
    process.exit(1);
  }

  console.log(`${changedFiles.length} file(s) need processing.`);

  // 4. Load pdf-parse now that we need it
  pdfParse = require('pdf-parse');

  // 5. Ensure temp dir
  if (!fs.existsSync(TMP_DIR)) {
    fs.mkdirSync(TMP_DIR, { recursive: true });
  }

  // 6. Ensure data dir
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }

  // 7. Load index
  const index = loadIndex();

  // 8. Process each changed file
  for (const entry of changedFiles) {
    console.log(`\n--- Processing: ${entry.name} ---`);

    // Detect course
    const course = detectCourse(entry.name);
    console.log(`  Course: ${course.name} (${course.id})`);

    // Download PDF
    console.log('  Downloading...');
    const pdfBuffer = await downloadFile(token, entry.path_lower);
    const tmpPath = path.join(TMP_DIR, entry.name);
    fs.writeFileSync(tmpPath, pdfBuffer);
    console.log(`  Downloaded (${(pdfBuffer.length / 1024).toFixed(0)} KB)`);

    // Extract text
    console.log('  Extracting text...');
    const data = await pdfParse(pdfBuffer);
    const text = data.text;
    console.log(`  Extracted ${text.length} characters from ${data.numpages} pages`);

    // Write JSON (the plugin expects a single JSON string)
    const jsonFile = `${course.id}_text.json`;
    const jsonPath = path.join(DATA_DIR, jsonFile);
    fs.writeFileSync(jsonPath, JSON.stringify(text));
    console.log(`  Wrote ${jsonPath}`);

    // Detect sessions
    const sessions = detectSessions(text);
    console.log(`  Detected ${sessions.length} session(s)`);

    // Update index
    index.courses = index.courses.filter((c) => c.id !== course.id);
    index.courses.push({
      id: course.id,
      name: course.name,
      file: jsonFile,
      sessions: sessions,
    });

    // Update manifest entry
    manifest.files[entry.path_lower] = {
      content_hash: entry.content_hash,
      name: entry.name,
      course_id: course.id,
      size: entry.size,
      last_modified: entry.server_modified,
      synced_at: new Date().toISOString(),
    };
  }

  // 9. Sort courses in index by id for consistency
  index.courses.sort((a, b) => a.id.localeCompare(b.id));

  // 10. Save everything
  saveIndex(index);
  console.log(`\nUpdated index.json — ${index.courses.length} course(s) total`);

  saveManifest(manifest);
  console.log('Updated sync-manifest.json');

  // 11. Clean up temp dir
  try {
    const tmpFiles = fs.readdirSync(TMP_DIR);
    for (const f of tmpFiles) {
      fs.unlinkSync(path.join(TMP_DIR, f));
    }
    fs.rmdirSync(TMP_DIR);
  } catch (e) {
    // non-fatal
  }

  console.log('\nSync complete. Changes were made.');
  process.exit(0);
}

main().catch((err) => {
  console.error('FATAL:', err.message);
  console.error(err.stack);
  process.exit(2);
});
