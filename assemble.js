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

// ויקי — בסיס ידע עברי קטן ל"עוזר הווירטואלי". כל קובץ md = מקטע אחד,
// עם front-matter של title ו-tags. נטמע כמו הנתונים, כדי לשמור על self-contained.
const WIKI_DIR = p.join(R, 'data', 'wiki');
const wiki = !fs.existsSync(WIKI_DIR) ? [] : fs.readdirSync(WIKI_DIR)
  .filter(f => f.endsWith('.md')).sort()
  .map(f => {
    const raw = rd(p.join('data', 'wiki', f));
    const m = /^---\r?\n([\s\S]*?)\r?\n---\r?\n?/.exec(raw);
    const fm = {};
    if (m) m[1].split(/\r?\n/).forEach(line => {
      const i = line.indexOf(':');
      if (i > 0) fm[line.slice(0, i).trim()] = line.slice(i + 1).trim();
    });
    return {
      id: f.replace(/\.md$/, ''),
      title: fm.title || f.replace(/\.md$/, ''),
      tags: (fm.tags || '').split(',').map(s => s.trim()).filter(Boolean),
      text: raw.slice(m ? m[0].length : 0).trim(),
    };
  });

let html = rd('index.template.html');
html = html.replace('<!--VENDOR_CSS-->', '<style>\n' + css + '\n</style>');
html = html.replace('<!--VENDOR_JS-->', '<script>\n' + js + '\n</script>');
html = html.replace('<!--SITE_DATA-->', '<script>window.SITE_DATA=' + data + ';</script>');
html = html.replace('<!--SITE_WIKI-->', '<script>window.SITE_WIKI=' + JSON.stringify(wiki) + ';</script>');

fs.writeFileSync(p.join(R, 'index.html'), html);
console.log('wrote index.html  (%d KB)  ·  wiki: %d מקטעים', Math.round(Buffer.byteLength(html) / 1024), wiki.length);
