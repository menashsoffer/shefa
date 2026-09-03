/*
 * assemble.js — מרכיב את index.html הסופי (קובץ יחיד, self-contained).
 * מטמיע: CSS ו-JS של Leaflet / markercluster / leaflet.heat / Chart.js  +  dist/site_data.json
 * מקורות הספריות (הורדו מראש ל-vendor/): cdnjs.cloudflare.com , cdn.jsdelivr.net (אפיון §12).
 */
'use strict';
const fs = require('fs');
const p = require('path');
const R = __dirname;
const rd = f => fs.readFileSync(p.join(R, f), 'utf8');

const css = [
  rd('vendor/leaflet.css'),
  rd('vendor/markercluster.css'),
  rd('vendor/markercluster.default.css'),
].join('\n');

const js = [
  rd('vendor/leaflet.js'),
  rd('vendor/markercluster.js'),
  rd('vendor/leaflet.heat.js'),
  rd('vendor/chart.umd.js'),
].join('\n;\n');

const data = rd('dist/site_data.json');

let html = rd('index.template.html');
html = html.replace('<!--VENDOR_CSS-->', '<style>\n' + css + '\n</style>');
html = html.replace('<!--VENDOR_JS-->', '<script>\n' + js + '\n</script>');
html = html.replace('<!--SITE_DATA-->', '<script>window.SITE_DATA=' + data + ';</script>');

fs.writeFileSync(p.join(R, 'index.html'), html);
console.log('wrote index.html  (%d KB)', Math.round(Buffer.byteLength(html) / 1024));
