/**
 * Build script: produces an obfuscated version of the dashboard HTML.
 * Usage: node build_obfuscated.js
 * Input:  output/performance_review_dashboard.html
 * Output: output/performance_review_dashboard_share.html
 */
const fs = require('fs');
const path = require('path');
const JavaScriptObfuscator = require('javascript-obfuscator');

const srcPath = path.join(__dirname, 'output', 'performance_review_dashboard.html');
const outPath = path.join(__dirname, 'output', 'performance_review_dashboard_share.html');

let html = fs.readFileSync(srcPath, 'utf-8');

// Extract all inline <script> blocks (skip CDN src= tags)
const scriptRegex = /<script>([\s\S]*?)<\/script>/gi;
let match;
const replacements = [];

while ((match = scriptRegex.exec(html)) !== null) {
  const originalJS = match[1];
  if (originalJS.trim().length < 50) continue; // skip tiny scripts

  console.log(`Obfuscating script block (${originalJS.length} chars)...`);

  const obfuscated = JavaScriptObfuscator.obfuscate(originalJS, {
    compact: true,
    controlFlowFlattening: true,
    controlFlowFlatteningThreshold: 0.5,
    deadCodeInjection: true,
    deadCodeInjectionThreshold: 0.2,
    debugProtection: false,
    disableConsoleOutput: false,
    identifierNamesGenerator: 'hexadecimal',
    renameGlobals: false,
    selfDefending: false,
    stringArray: true,
    stringArrayEncoding: ['base64'],
    stringArrayThreshold: 0.75,
    transformObjectKeys: true,
    unicodeEscapeSequence: true,
    target: 'browser',
  });

  replacements.push({
    original: match[0],
    replacement: '<script>' + obfuscated.getObfuscatedCode() + '</script>'
  });
}

// Apply replacements
for (const r of replacements) {
  html = html.replace(r.original, r.replacement);
}

// Add a comment at the top
html = html.replace('<!DOCTYPE html>',
  '<!DOCTYPE html>\n<!-- CAT Comp Review Dashboard - CONFIDENTIAL - Obfuscated build -->\n');

fs.writeFileSync(outPath, html, 'utf-8');

const srcSize = fs.statSync(srcPath).size;
const outSize = fs.statSync(outPath).size;
console.log(`\nDone!`);
console.log(`  Source: ${(srcSize/1024).toFixed(0)} KB`);
console.log(`  Output: ${(outSize/1024).toFixed(0)} KB`);
console.log(`  Saved to: ${outPath}`);
