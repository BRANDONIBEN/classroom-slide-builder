#!/usr/bin/env node
/**
 * pdf-to-json.js — Convert a course PDF to the JSON text format used by the plugin.
 *
 * Usage:
 *   node scripts/pdf-to-json.js <input.pdf> <output.json>
 *   node scripts/pdf-to-json.js "docs/Essential Theology VOD Slides.pdf" remote/courses/et_text.json
 *
 * Requires: npm install pdf-parse
 */

const fs = require('fs');
const path = require('path');
const pdfParse = require('pdf-parse');

const inputPath = process.argv[2];
const outputPath = process.argv[3];

if (!inputPath || !outputPath) {
  console.error('Usage: node pdf-to-json.js <input.pdf> <output.json>');
  process.exit(1);
}

async function convert() {
  const dataBuffer = fs.readFileSync(path.resolve(inputPath));
  const data = await pdfParse(dataBuffer);

  // The plugin expects a single JSON string of the full text with \n line breaks
  const text = data.text;
  fs.writeFileSync(path.resolve(outputPath), JSON.stringify(text));

  console.log('Converted: ' + inputPath + ' → ' + outputPath);
  console.log('  Pages: ' + data.numpages);
  console.log('  Characters: ' + text.length);
}

convert().catch(function (err) {
  console.error('Error:', err.message);
  process.exit(1);
});
