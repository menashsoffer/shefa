/*
 * build.js — מחולל הנתונים המוטמעים לאתר "מוקד שפ״ע אריאל: תמונת מצב".
 *
 * זהו המחולל שרץ בפועל בסביבה הזו (אין בה Python runtime).
 * build.py הוא מימוש-מראה זהה לוגית ב-pandas, ומתועד ב-ACCEPTANCE.md.
 *
 * קלט (כפי שסופק בפועל — קבצי XLSX, לא CSV):
 *   Downloads/tickets (99).xlsx        — 12,309 פניות
 *   Downloads/subjects (21).xlsx       — 165 שורות תצורת SLA
 *   Downloads/רחובות אריאל.xlsx        — טבלת רחוב→שכונה (קלט בונוס, לא באפיון המקורי)
 *   data/street_geo_final.json         — 107 רחובות ממופים (הודבק ע״י המשתמש)
 *   data/ariel_boundary.geojson        — גבול שיפוט
 *
 * פלט:
 *   dist/site_data.json                — כל הנתונים, מוטמע ב-index.html ע״י assemble.js
 *   ACCEPTANCE.md                      — 18 בדיקות קבלה, מספר מול מספר
 */
'use strict';
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const DL = '/Users/moranrachamim/Downloads';
const ROOT = __dirname;
const F_TICKETS = path.join(DL, 'tickets (99).xlsx');
const F_SUBJECTS = path.join(DL, 'subjects (21).xlsx');
const F_STREETS = path.join(DL, 'רחובות אריאל.xlsx');
const F_GEO = path.join(ROOT, 'data', 'street_geo_final.json');
const F_BOUNDARY = path.join(ROOT, 'data', 'ariel_boundary.geojson');

// ---------- helpers ----------
const TRIM = v => (v == null ? null : (typeof v === 'string' ? v.trim() : v));
// "-" בודד ומחרוזת ריקה => null (אפיון §3)
const NN = v => {
  const t = TRIM(v);
  if (t == null) return null;
  if (typeof t === 'string' && (t === '' || t === '-')) return null;
  return t;
};
// Excel serial (ימים מ-1899-12-30) => Date ב-UTC. מחרוזת dd/mm/yyyy hh:mm נתמכת גם.
function xlDate(v) {
  const t = NN(v);
  if (t == null) return null;
  if (typeof t === 'number') return new Date(Math.round((t - 25569) * 86400000));
  const m = String(t).match(/^(\d{2})\/(\d{2})\/(\d{4})[ T]+(\d{1,2}):(\d{2})/);
  if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]));
  const d = new Date(t);
  return isNaN(d) ? null : d;
}
const sheet1 = f => {
  const wb = XLSX.readFile(f, { raw: true, cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
};
const quantile = (sorted, p) => {
  if (!sorted.length) return null;
  const i = (sorted.length - 1) * p, lo = Math.floor(i), hi = Math.ceil(i);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (i - lo);
};

// ---------- status super-groups (אפיון §3.2) ----------
const STATUS_GROUPS = ['טופל', 'נסגר ללא טיפול', 'פתוח', 'הועבר לגורם חיצוני'];
const STATUS_TO_GROUP = new Map();
[
  ['טופל', ['טופל', 'תהליך הסתיים', 'טופל ללא סקר']],
  ['נסגר ללא טיפול', ['נסגר ללא ביצוע', 'לא נמצא מפגע', 'לא שייך לעירייה', 'כפל פנייה', 'שטח פרטי']],
  ['פתוח', ['פנייה חדשה', 'בטיפול', 'נא לחזור לתושב', 'פנייה שנפתחה מחדש', 'ממתין להפניה מחדש', 'תוכנית עבודה']],
  ['הועבר לגורם חיצוני', ['הועבר לגורם חיצוני', 'הועבר לנת"י', 'הועבר לטיפול חברת החשמל',
    'אין תמרור - יש להעביר לועדת תחבורה', 'אין תמרור – יש להעביר לועדת תחבורה']],
].forEach(([g, arr]) => arr.forEach(s => STATUS_TO_GROUP.set(s, g)));

// ---------- 1. subjects.csv → SLA lookup (אפיון §4) ----------
function buildSla() {
  const S = sheet1(F_SUBJECTS);
  const H = S[0].map(TRIM);
  const c = n => H.indexOf(n);
  const iDep = c('מחלקות'), iSub = c('נושאי הפנייה'), iSla = c('SLA'),
    iUrg = c('דחיפות'), iAct = c('פעיל?');
  const map = {};            // "dep|sub" -> {...}   keep='first'  (אפיון §4.1)
  let slaBlank = 0, urgBlank = 0, active = { 'כן': 0, 'לא': 0 };
  for (const r of S.slice(1)) {
    const dep = TRIM(r[iDep]), sub = TRIM(r[iSub]);
    const slaRaw = r[iSla];                                   // מספר = שבר של יום; ×24 = "שעות" = ימי עבודה
    if (slaRaw == null) slaBlank++;
    if (NN(r[iUrg]) == null) urgBlank++;
    if (r[iAct] in active) active[r[iAct]]++;
    if (dep == null) continue;
    const key = dep + '|' + (sub == null ? '' : sub);
    if (key in map) continue;                                 // keep first
    const workdays = slaRaw == null ? null : Math.round(slaRaw * 24);
    map[key] = {
      workdays,
      broken: slaRaw === 0,                                   // "תמרורים ושלטים" = 00:00:00 (אפיון §4.2)
      calendarDays: workdays == null ? null : Math.round(workdays * 1.4 * 10) / 10,
      urgency: NN(r[iUrg]),
      active: r[iAct] == null ? null : String(r[iAct]),
    };
  }
  return { map, stats: { rows: S.length - 1, slaBlank, urgBlank, active } };
}

// ---------- 2. רחובות אריאל.xlsx → רחוב→שכונה (בונוס) ----------
function buildHoods() {
  const N = sheet1(F_STREETS);
  const H = N[0].map(TRIM);
  const iName = H.indexOf('רחוב'), iHood = H.indexOf('שכונה');
  const hood = {};
  for (const r of N.slice(1)) {
    const nm = TRIM(r[iName]);
    if (nm) hood[nm] = NN(r[iHood]) ? String(r[iHood]).trim() : null;
  }
  return hood;
}

// ---------- 3. street_geo_final.json ----------
function buildGeo(hoods) {
  const raw = JSON.parse(fs.readFileSync(F_GEO, 'utf8'));
  const streets = {};
  for (const [name, o] of Object.entries(raw.streets)) {
    streets[name] = {
      lat: o.lat, lon: o.lon, conf: o.conf, src: o.src,
      approx: o.conf === 'בינוני' || o.conf === 'נמוך',       // קו מתאר מקווקו (אפיון §5.1.1)
      hood: hoods[name] || null,
    };
  }
  return {
    streets,
    unlocated_streets: raw.unlocated_streets,                 // 4 רחובות, 60 פניות — לא ב-OSM
    no_location: raw.no_location,                             // ריק / לא ידוע / לא באריאל — 43
  };
}

// ---------- 4. tickets.csv → per-ticket records + aggregates ----------
function buildTickets(slaMap, geo) {
  const T = sheet1(F_TICKETS);
  const H = T[0].map(TRIM);
  const c = n => H.indexOf(n);
  const cId = c("מס' פניה"), cStat = c('סטטוס פנייה'), cMetric = c('מדד SLA לפני חריגה'),
    cOpen = c('תאריך ושעת פתיחה'), cDue = c('תאריך יעד לסגירה'), cDep = c('מחלקה'),
    cSub = c('נושא משנה'), cStreet = c('רחוב'), cClose = c('תאריך סגירה');
  const rows = T.slice(1);

  // dimensions
  const depList = [], subList = [], streetList = [];
  const depIx = new Map(), subIx = new Map(), streetIx = new Map();
  const di = v => { if (!depIx.has(v)) { depIx.set(v, depList.length); depList.push(v); } return depIx.get(v); };
  const si = v => { if (!subIx.has(v)) { subIx.set(v, subList.length); subList.push(v); } return subIx.get(v); };
  const sti = v => { if (v == null) return -1; if (!streetIx.has(v)) { streetIx.set(v, streetList.length); streetList.push(v); } return streetIx.get(v); };

  const statusList = [], statusIx = new Map();
  const stiRaw = v => { if (!statusIx.has(v)) { statusIx.set(v, statusList.length); statusList.push(v); } return statusIx.get(v); };
  const ids = [];

  const EPOCH = Date.UTC(2025, 0, 1);
  const MONTHS = [];                                          // "YYYY-MM" בסדר עולה
  const monthIx = new Map();
  const mi = key => { if (!monthIx.has(key)) monthIx.set(key, -1); return key; };

  const FLAG = { METRIC: 1, BREACH: 2, CLOSE: 4, DUE: 8, LATE: 16, FOLLOWUP: 32 };
  const tickets = [];
  const treat = [];
  const monthCount = {};
  const slaUnmatched = new Map();

  let idBlank = 0, idSet = new Set(), badOpen = 0, closeBeforeOpen = 0;

  for (const r of rows) {
    const idFmt = r[cId] == null ? '' : String(r[cId]).trim();
    if (idFmt === '') idBlank++;
    idSet.add(idFmt);

    const dep = TRIM(r[cDep]);
    const sub = TRIM(r[cSub]);
    const statRaw = TRIM(r[cStat]);
    const grp = STATUS_TO_GROUP.get(statRaw);
    if (grp === undefined) throw new Error('סטטוס לא ממופה: ' + JSON.stringify(statRaw));

    const street = NN(r[cStreet]);
    const oDate = xlDate(r[cOpen]);
    const dDate = xlDate(r[cDue]);
    const cDate = xlDate(r[cClose]);
    if (!oDate) { badOpen++; }

    const metricRaw = NN(r[cMetric]);                         // שבר: 0.0058 => 0.58% ; 1 => 100% = חריגה
    const hasMetric = metricRaw != null;
    const breach = metricRaw === 1;

    let flags = 0;
    if (hasMetric) flags |= FLAG.METRIC;
    if (breach) flags |= FLAG.BREACH;
    if (cDate) flags |= FLAG.CLOSE;
    if (dDate) flags |= FLAG.DUE;
    if (dDate && cDate && cDate > dDate) flags |= FLAG.LATE;
    if (/-[֐-׿]+$/.test(idFmt)) flags |= FLAG.FOLLOWUP;

    let treatH = -1;
    if (oDate && cDate) {
      if (cDate < oDate) closeBeforeOpen++;
      else { treatH = Math.round((cDate - oDate) / 3600000 * 100) / 100; treat.push(treatH); }
    }

    let monthIdx = -1, openDay = -1;
    if (oDate) {
      const key = oDate.getUTCFullYear() + '-' + String(oDate.getUTCMonth() + 1).padStart(2, '0');
      monthCount[key] = (monthCount[key] || 0) + 1;
      mi(key);
      openDay = Math.floor((oDate.getTime() - EPOCH) / 86400000);
    }

    // SLA join (אפיון §4.1)
    const jkey = (dep || '') + '|' + (sub || '');
    if (!(jkey in slaMap)) slaUnmatched.set(jkey, (slaUnmatched.get(jkey) || 0) + 1);

    ids.push(idFmt);
    tickets.push([sti(street), di(dep), si(sub), STATUS_GROUPS.indexOf(grp), key_or_neg(oDate), openDay, treatH, flags, stiRaw(statRaw)]);
  }
  function key_or_neg(oDate) {
    if (!oDate) return -1;
    return oDate.getUTCFullYear() + '-' + String(oDate.getUTCMonth() + 1).padStart(2, '0');
  }

  // month index list, sorted ascending
  MONTHS.push(...[...monthIx.keys()].sort());
  const monthIdxOf = new Map(MONTHS.map((k, i) => [k, i]));
  // replace month key (position 4) with numeric index
  for (const t of tickets) t[4] = t[4] === -1 ? -1 : monthIdxOf.get(t[4]);

  const months = MONTHS.map(k => {
    const [y, m] = k.split('-');
    return { key: k, label: m + '/' + y, count: monthCount[k] };
  });

  // ---------- aggregates on the FULL dataset (for story mode + acceptance) ----------
  const IX = { STREET: 0, DEP: 1, SUB: 2, GRP: 3, MIDX: 4, DAY: 5, TREAT: 6, FLAGS: 7 };
  const has = (t, f) => (t[IX.FLAGS] & f) !== 0;

  const byDep = {}, byStatusGrp = [0, 0, 0, 0], bySubject = {}, byStreet = {};
  let breachCount = 0, breachOpen = 0, openNoClose = 0, lateClose = 0, haveDueClose = 0;
  let noDueNoSla = 0, noDueNoSlaVet = 0, followups = 0, y2025 = 0, y2026 = 0;
  const breachByDepAll = {}, ticketsByDepAll = {};
  let oldestOpenDay = Infinity, oldestOpenIdx = -1;

  tickets.forEach((t, i) => {
    const dep = depList[t[IX.DEP]];
    byDep[dep] = (byDep[dep] || 0) + 1;
    ticketsByDepAll[dep] = (ticketsByDepAll[dep] || 0) + 1;
    byStatusGrp[t[IX.GRP]]++;
    const sub = subList[t[IX.SUB]];
    bySubject[sub] = (bySubject[sub] || 0) + 1;
    const stName = t[IX.STREET] === -1 ? '(ללא רחוב)' : streetList[t[IX.STREET]];
    byStreet[stName] = (byStreet[stName] || 0) + 1;

    if (has(t, FLAG.BREACH)) { breachCount++; breachByDepAll[dep] = (breachByDepAll[dep] || 0) + 1; if (!has(t, FLAG.CLOSE)) breachOpen++; }
    if (!has(t, FLAG.CLOSE)) {
      openNoClose++;
      if (t[IX.DAY] >= 0 && t[IX.DAY] < oldestOpenDay) { oldestOpenDay = t[IX.DAY]; oldestOpenIdx = i; }
    }
    if (has(t, FLAG.DUE) && has(t, FLAG.CLOSE)) { haveDueClose++; if (has(t, FLAG.LATE)) lateClose++; }
    if (!has(t, FLAG.DUE) && !has(t, FLAG.METRIC)) { noDueNoSla++; if (dep === 'וטרינר') noDueNoSlaVet++; }
    if (has(t, FLAG.FOLLOWUP)) followups++;
    if (t[IX.MIDX] >= 0) {
      const y = +months[t[IX.MIDX]].key.slice(0, 4);
      if (y === 2025) y2025++; else if (y === 2026) y2026++;
    }
  });

  const treatSorted = treat.slice().sort((a, b) => a - b);
  const medianTreat = quantile(treatSorted, 0.5);
  const p90Treat = quantile(treatSorted, 0.9);
  const maxTreat = treatSorted[treatSorted.length - 1];

  // location buckets (אפיון §5)
  let located = 0, unlocatedStreet = 0, noStreet = 0;
  const unlocBreak = {}, noStreetBreak = {};
  tickets.forEach(t => {
    const idx = t[IX.STREET];
    if (idx === -1) { noStreet++; noStreetBreak['(ריק)'] = (noStreetBreak['(ריק)'] || 0) + 1; return; }
    const name = streetList[idx];
    if (geo.streets[name]) { located++; return; }
    if (name === 'לא ידוע' || name === 'לא באריאל') { noStreet++; noStreetBreak[name] = (noStreetBreak[name] || 0) + 1; return; }
    unlocatedStreet++; unlocBreak[name] = (unlocBreak[name] || 0) + 1;
  });

  const breachRateByDep = {};
  Object.keys(ticketsByDepAll).forEach(d => {
    breachRateByDep[d] = { breaches: breachByDepAll[d] || 0, total: ticketsByDepAll[d], pct: Math.round((breachByDepAll[d] || 0) / ticketsByDepAll[d] * 1000) / 10 };
  });
  const openByDep = {};
  tickets.forEach(t => { if (!has(t, FLAG.CLOSE)) { const d = depList[t[IX.DEP]]; openByDep[d] = (openByDep[d] || 0) + 1; } });

  const oldestOpen = oldestOpenIdx >= 0
    ? { date: new Date(EPOCH + oldestOpenDay * 86400000).toISOString().slice(0, 10), dep: depList[tickets[oldestOpenIdx][IX.DEP]], sub: subList[tickets[oldestOpenIdx][IX.SUB]] }
    : null;

  return {
    dims: { deps: depList, subjects: subList, streets: streetList, statusGroups: STATUS_GROUPS, statusesRaw: statusList },
    months,
    flags: FLAG,
    tickets,
    ids,
    slaUnmatched: [...slaUnmatched.entries()].map(([k, n]) => ({ key: k, count: n })),
    aggregates: {
      rows: rows.length,
      idBlank, idUnique: idSet.size, badOpen, closeBeforeOpen,
      byDep, byStatusGroup: Object.fromEntries(STATUS_GROUPS.map((g, i) => [g, byStatusGrp[i]])),
      bySubjectTop: Object.entries(bySubject).sort((a, b) => b[1] - a[1]).slice(0, 15),
      subjectCount: Object.keys(bySubject).length,
      located, unlocatedStreet, noStreet, offMapTotal: unlocatedStreet + noStreet,
      unlocBreak, noStreetBreak,
      breachCount, breachOpen, openNoClose, lateClose, haveDueClose,
      noDueNoSla, noDueNoSlaVet, followups, y2025, y2026,
      medianTreat, p90Treat, maxTreat, treatCount: treat.length,
      breachRateByDep, openByDep, oldestOpen,
      dateMin: (() => { let m = Infinity; tickets.forEach(t => { if (t[IX.DAY] >= 0) m = Math.min(m, t[IX.DAY]); }); return new Date(EPOCH + m * 86400000).toISOString().slice(0, 10); })(),
      dateMax: (() => { let m = -Infinity; tickets.forEach(t => { if (t[IX.DAY] >= 0) m = Math.max(m, t[IX.DAY]); }); return new Date(EPOCH + m * 86400000).toISOString().slice(0, 10); })(),
    },
  };
}

// ---------- run ----------
const sla = buildSla();
const hoods = buildHoods();
const geo = buildGeo(hoods);
const T = buildTickets(sla.map, geo);
const boundary = JSON.parse(fs.readFileSync(F_BOUNDARY, 'utf8'));

const site = {
  meta: {
    generated: new Date().toISOString(),
    source: 'tickets (99).xlsx · subjects (21).xlsx · רחובות אריאל.xlsx · street_geo_final.json · ariel_boundary.geojson',
    view: { center: [32.1050, 35.1900], zoom: 14, bbox: { lonMin: 35.166, lonMax: 35.211, latMin: 32.099, latMax: 32.111 } },
    attribution: 'OpenStreetMap contributors, ODbL — relation 10011903',
  },
  dims: T.dims,
  ids: T.ids,
  months: T.months,
  flagBits: T.flags,
  geo: { streets: geo.streets, unlocated_streets: geo.unlocated_streets, no_location: geo.no_location, boundary },
  sla: { map: sla.map, stats: sla.stats, unmatched: T.slaUnmatched },
  tickets: T.tickets,
  aggregates: T.aggregates,
};

fs.mkdirSync(path.join(ROOT, 'dist'), { recursive: true });
fs.writeFileSync(path.join(ROOT, 'dist', 'site_data.json'), JSON.stringify(site));
console.log('wrote dist/site_data.json  (%d KB)', Math.round(fs.statSync(path.join(ROOT, 'dist', 'site_data.json')).size / 1024));

// ---------- acceptance table ----------
const a = T.aggregates;
const depSum = Object.values(a.byDep).reduce((x, y) => x + y, 0);
const grpSum = Object.values(a.byStatusGroup).reduce((x, y) => x + y, 0);
const subjSum = a.bySubjectTop.reduce((x, y) => x + y[1], 0); // partial (top15) — full check via subjectCount
const rows = [
  ['1', 'סה״כ שורות', 12309, a.rows],
  ['2', 'סכום פילוח מחלקות = סה״כ', 12309, depSum],
  ['2b', 'סכום קבוצות-סטטוס = סה״כ', 12309, grpSum],
  ['3', 'פניות עם lat/lon', 12206, a.located],
  ['4', 'פניות ללא מיקום מפה (60+43)', 103, a.offMapTotal],
  ['5', 'חריגות SLA (מדד=100%)', 2685, a.breachCount],
  ['6', 'חריגות שעדיין פתוחות', 384, a.breachOpen],
  ['7', 'פניות פתוחות (ללא ת. סגירה)', 704, a.openNoClose],
  ['8', 'נסגרו אחרי ת. יעד', 2300, a.lateClose],
  ['8b', 'בסיס: יש יעד + סגירה', 10402, a.haveDueClose],
  ['9', 'ללא יעד וללא SLA', 1251, a.noDueNoSla],
  ['9b', '— מזה בוטרינר', 627, a.noDueNoSlaVet],
  ['10', 'פניות המשך', 881, a.followups],
  ['11', 'חציון זמן טיפול (שעות)', 34.5, Math.round(a.medianTreat * 10) / 10],
  ['12', 'אחוזון 90 זמן טיפול (שעות)', 856.8, Math.round(a.p90Treat * 10) / 10],
  ['13a', 'פניות 2025', 6617, a.y2025],
  ['13b', 'פניות 2026', 5692, a.y2026],
  ['14', 'צירופים ללא התאמה ל-subjects', 2, T.slaUnmatched.length],
];
let md = '# בדיקות קבלה — build.js\n\n';
md += 'נוצר: ' + site.meta.generated + '\n\n';
md += '| # | בדיקה | צפוי | בפועל | תוצאה |\n|---|---|---|---|---|\n';
let allPass = true;
for (const [n, name, exp, got] of rows) {
  const ok = Math.abs((+got) - (+exp)) < 0.06 || got === exp;
  if (!ok) allPass = false;
  md += `| ${n} | ${name} | ${exp} | ${got} | ${ok ? '✅' : '❌'} |\n`;
  console.log((ok ? 'PASS ' : 'FAIL ') + n.padEnd(3) + ' ' + name + '  expected=' + exp + ' got=' + got);
}
md += `\nמספר נושאי משנה ייחודיים: ${a.subjectCount} (צפוי 81)\n`;
md += `טווח תאריכי פתיחה: ${a.dateMin} → ${a.dateMax}\n`;
md += `close<open: ${a.closeBeforeOpen} · badOpen: ${a.badOpen} · id blank: ${a.idBlank} · id unique: ${a.idUnique}\n`;
md += `max זמן טיפול: ${Math.round(a.maxTreat)} שעות ≈ ${(a.maxTreat / 24).toFixed(1)} ימים\n`;
md += '\n**' + (allPass ? 'כל 18 הבדיקות עברו.' : 'יש בדיקה שנכשלה — ראה למעלה.') + '**\n';
fs.writeFileSync(path.join(ROOT, 'ACCEPTANCE_build.md'), md);
console.log('\n' + (allPass ? 'ALL PASS' : 'SOME FAIL') + '  ·  wrote ACCEPTANCE_build.md');
