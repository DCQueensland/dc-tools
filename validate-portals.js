#!/usr/bin/env node
// DC Tools Portal Validator
// Catches the silent-failure bugs that have hit customer-intelligence portals
// before (data stripped, render function missing, syntax errors, etc.)
// Run manually: node validate-portals.js
// Auto-run: invoked by .git/hooks/pre-commit before any commit
//
// Exit code 0 = all portals valid, 1 = at least one portal broken.

const fs = require('fs');
const path = require('path');

const REQUIRED_FUNCTIONS = ['function render()', 'function tierClass', 'function badgeClass', 'function escapeHtml'];
const MIN_RECORDS = 5; // catches the "DATA failed to parse" case; early years (2016) legitimately have few records

function validatePortal(file) {
  const issues = [];
  const html = fs.readFileSync(file, 'utf8');

  // 1. Required functions present
  for (const fn of REQUIRED_FUNCTIONS) {
    if (!html.includes(fn)) issues.push(`missing ${fn}`);
  }

  // 2. DATA array parses + has records
  const m = html.match(/var DATA = \[[\s\S]*?\n\];/);
  if (!m) {
    issues.push('no DATA array found');
  } else {
    try {
      const fn = new Function(m[0] + '; return DATA');
      const data = fn();
      if (!Array.isArray(data)) issues.push('DATA is not an array');
      else if (data.length < MIN_RECORDS) issues.push(`only ${data.length} records (expected >= ${MIN_RECORDS})`);
    } catch (e) {
      issues.push(`DATA parse error: ${e.message}`);
    }
  }

  // 3. Tbody element present
  if (!html.includes('id="tbody"')) issues.push('missing <tbody id="tbody">');

  return issues;
}

const dir = path.join(__dirname);
const portals = fs.readdirSync(dir).filter(f => /^customer-intelligence-\d{4}\.html$/.test(f)).sort();

let failed = 0;
console.log('DC Tools Portal Validator\n=========================');
for (const f of portals) {
  const issues = validatePortal(path.join(dir, f));
  if (issues.length === 0) {
    console.log(`  ✓ ${f}`);
  } else {
    failed++;
    console.log(`  ✗ ${f}`);
    for (const i of issues) console.log(`      - ${i}`);
  }
}
console.log('=========================');
if (failed === 0) {
  console.log(`All ${portals.length} portals valid.`);
  process.exit(0);
} else {
  console.log(`${failed} of ${portals.length} portals BROKEN. Fix before committing.`);
  process.exit(1);
}
