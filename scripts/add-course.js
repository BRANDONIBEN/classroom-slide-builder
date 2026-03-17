#!/usr/bin/env node
/**
 * add-course.js — Add a new course to the remote index after converting its PDF.
 *
 * Usage:
 *   node scripts/add-course.js <course-id> <course-name> <input.pdf>
 *
 * Example:
 *   node scripts/add-course.js nt "New Testament Survey" ~/Dropbox/watch/nt_slides.pdf
 *
 * This will:
 *   1. Convert the PDF to JSON text
 *   2. Save it to remote/courses/<id>_text.json
 *   3. Auto-detect sessions from the text
 *   4. Add the course entry to remote/index.json
 */

const fs = require('fs');
const path = require('path');
const pdfParse = require('pdf-parse');

const courseId = process.argv[2];
const courseName = process.argv[3];
const pdfPath = process.argv[4];

if (!courseId || !courseName || !pdfPath) {
  console.error('Usage: node add-course.js <course-id> <course-name> <input.pdf>');
  console.error('  e.g. node add-course.js nt "New Testament Survey" ~/Dropbox/watch/nt_slides.pdf');
  process.exit(1);
}

const REMOTE_DIR = path.resolve(__dirname, '..', 'remote');
const COURSES_DIR = path.join(REMOTE_DIR, 'courses');
const INDEX_PATH = path.join(REMOTE_DIR, 'index.json');

async function run() {
  // 1. Convert PDF
  const dataBuffer = fs.readFileSync(path.resolve(pdfPath));
  const data = await pdfParse(dataBuffer);
  const text = data.text;

  // 2. Save JSON
  const jsonFile = courseId + '_text.json';
  const jsonPath = path.join(COURSES_DIR, jsonFile);
  fs.writeFileSync(jsonPath, JSON.stringify(text));
  console.log('Saved: ' + jsonPath);

  // 3. Auto-detect sessions from text
  var sessions = [];
  var sessionPattern = /Session\s+(\d+):\s*(.+?)(?:\s*\(|$)/gm;
  var match;
  var seen = {};
  while ((match = sessionPattern.exec(text)) !== null) {
    var num = parseInt(match[1]);
    var label = match[2].trim();
    if (!seen[num]) {
      seen[num] = true;
      sessions.push({ num: num, label: label });
    }
  }

  if (sessions.length === 0) {
    console.warn('Warning: Could not auto-detect sessions. You may need to edit remote/index.json manually.');
    sessions.push({ num: 1, label: 'Session 1' });
  } else {
    sessions.sort(function (a, b) { return a.num - b.num; });
    console.log('Detected ' + sessions.length + ' sessions');
  }

  // 4. Update index.json
  var index = JSON.parse(fs.readFileSync(INDEX_PATH, 'utf8'));

  // Remove existing entry for this ID if present
  index.courses = index.courses.filter(function (c) { return c.id !== courseId; });

  index.courses.push({
    id: courseId,
    name: courseName,
    file: 'courses/' + jsonFile,
    sessions: sessions
  });

  fs.writeFileSync(INDEX_PATH, JSON.stringify(index, null, 2) + '\n');
  console.log('Updated index.json — ' + index.courses.length + ' courses total');
  console.log('\nNext: copy the remote/ folder to your Dropbox shared folder.');
}

run().catch(function (err) {
  console.error('Error:', err.message);
  process.exit(1);
});
