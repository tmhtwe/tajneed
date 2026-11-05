#!/usr/bin/env node
// Rebuilds data/members.enc.json from a "MembersGroupedByJamaat" Excel export.
//
// The encrypted file is the only place DOB and Musi are kept: each run decrypts
// the current file, carries DOB/Musi over by member code, and stops if a member
// in the Excel has no DOB/Musi yet (pass them with --extra).
//
// Usage (key comes from the environment, never the command line):
//   TAJNEED_KEY=... node tools/update.js --excel "MembersGroupedByJamaat.xlsx" [--extra extra.json] [--tobe tobe.xlsx] [--dry-run]
//
// extra.json: { "2603": { "birthday": "1983-02-10", "musi": false } }
//   Also used to change an existing member's DOB or Musi.
// tobe.xlsx / tobe.json: next year's Ansars-to-be (Code, Name, Majlis, Birthday, Email, Mobile).
//   Without it the current list is kept.
// --seed old.json: first run only, plain {members, ansarsToBe} to start from instead of the encrypted file.

const fs = require('fs');
const path = require('path');
const { webcrypto } = require('crypto');
const XLSX = require('xlsx');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(ROOT, 'data', 'members.enc.json');
const ITER = 310000;
const NOW = new Date();

// ---------- args ----------
function parseArgs(argv) {
  const a = {};
  for (let i = 0; i < argv.length; i++) {
    const k = argv[i];
    if (!k.startsWith('--')) continue;
    const next = argv[i + 1];
    if (next === undefined || next.startsWith('--')) a[k.slice(2)] = true;
    else { a[k.slice(2)] = next; i++; }
  }
  return a;
}

// ---------- crypto (same scheme as index.html) ----------
const b64 = buf => Buffer.from(buf).toString('base64');
const unb64 = s => new Uint8Array(Buffer.from(s, 'base64'));

async function deriveKey(pass, salt, iter) {
  const base = await webcrypto.subtle.importKey('raw', new TextEncoder().encode(pass), 'PBKDF2', false, ['deriveKey']);
  return webcrypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: iter, hash: 'SHA-256' }, base,
    { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
}
async function encrypt(obj, pass) {
  const salt = webcrypto.getRandomValues(new Uint8Array(16));
  const iv = webcrypto.getRandomValues(new Uint8Array(12));
  const key = await deriveKey(pass, salt, ITER);
  const ct = await webcrypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, new TextEncoder().encode(JSON.stringify(obj)));
  return { v: 1, alg: 'PBKDF2-SHA256/AES-256-GCM', iter: ITER, salt: b64(salt), iv: b64(iv), ct: b64(ct) };
}
async function decrypt(env, pass) {
  const key = await deriveKey(pass, unb64(env.salt), env.iter);
  const pt = await webcrypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(env.iv) }, key, unb64(env.ct));
  return JSON.parse(new TextDecoder().decode(pt));
}

// ---------- helpers ----------
const MON = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
const pad = n => String(n).padStart(2, '0');
const clean = v => (v === undefined || v === null) ? '' : String(v).trim().replace(/\s{2,}/g, ' ');

// Accepts 1985-03-29, 29 Mar 1985, 29/03/1985 (day first, NZ style) or an Excel date.
function toIso(v) {
  if (v === undefined || v === null || v === '') return '';
  if (v instanceof Date) return v.getFullYear() + '-' + pad(v.getMonth() + 1) + '-' + pad(v.getDate());
  const s = String(v).trim();
  let m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return m[1] + '-' + pad(m[2]) + '-' + pad(m[3]);
  if ((m = s.match(/^(\d{1,2})[\s-]+([A-Za-z]{3})[A-Za-z]*[\s-]+(\d{4})$/)) && MON[m[2].toLowerCase()])
    return m[3] + '-' + pad(MON[m[2].toLowerCase()]) + '-' + pad(m[1]);
  if ((m = s.match(/^(\d{1,2})[/.](\d{1,2})[/.](\d{4})$/))) return m[3] + '-' + pad(m[2]) + '-' + pad(m[1]);
  throw new Error('Unrecognised date: ' + s);
}
function toBool(v) {
  if (typeof v === 'boolean') return v;
  return /^(y|yes|true|1|musi)$/i.test(clean(v));
}

function readRows(file) {
  if (/\.json$/i.test(file)) return JSON.parse(fs.readFileSync(file, 'utf8'));
  const wb = XLSX.readFile(file, { cellDates: true });
  const rows = [];
  for (const name of wb.SheetNames) {
    if (/^statistics$/i.test(name)) continue;
    for (const r of XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '', raw: true })) rows.push(r);
  }
  return rows;
}
const pick = (r, ...keys) => { for (const k of keys) for (const rk of Object.keys(r)) if (rk.trim().toLowerCase() === k) return r[rk]; return ''; };

// ---------- main ----------
(async () => {
  const args = parseArgs(process.argv.slice(2));
  const pass = process.env.TAJNEED_KEY;
  if (!args.excel) { console.error('Missing --excel <MembersGroupedByJamaat.xlsx>'); process.exit(1); }
  if (!pass) { console.error('Set TAJNEED_KEY to the dashboard access key.'); process.exit(1); }

  // Previous data: DOB, Musi and the Ansars-to-be list live here.
  let prev;
  if (args.seed) prev = JSON.parse(fs.readFileSync(args.seed, 'utf8'));
  else if (fs.existsSync(OUT)) {
    try { prev = await decrypt(JSON.parse(fs.readFileSync(OUT, 'utf8')), pass); }
    catch (e) { console.error('Could not decrypt data/members.enc.json with this key.'); process.exit(1); }
  } else { console.error('No data/members.enc.json yet: pass --seed for the first run.'); process.exit(1); }

  const known = new Map();   // code -> {birthday, musi}
  for (const m of [...Object.values(prev.archive || {}), ...(prev.members || [])])
    known.set(String(m.code), { name: m.name, birthday: toIso(m.birthday), musi: !!m.musi });
  const extra = args.extra ? JSON.parse(fs.readFileSync(args.extra, 'utf8')) : {};
  for (const [code, e] of Object.entries(extra)) {
    const k = known.get(String(code)) || {};
    known.set(String(code), Object.assign({}, k,
      e.birthday !== undefined ? { birthday: toIso(e.birthday) } : {},
      e.musi !== undefined ? { musi: toBool(e.musi) } : {}));
  }

  // Members from the Excel
  const members = [], missing = [], seen = new Set();
  for (const r of readRows(args.excel)) {
    const code = clean(pick(r, 'code', 'tajneed', 'tcode'));
    if (!code || seen.has(code)) continue;
    seen.add(code);
    const k = known.get(code);
    const m = {
      code: Number(code) || code,
      name: clean(pick(r, 'name')),
      majlis: clean(pick(r, 'jamaat', 'majlis')),
      birthday: k ? k.birthday : '',
      musi: k ? !!k.musi : false,
      email: clean(pick(r, 'email')),
      phone: clean(pick(r, 'phone')),
      mobile: clean(pick(r, 'mobile')),
      address: clean(pick(r, 'address')),
    };
    if (!k || !k.birthday) missing.push(m);
    members.push(m);
  }
  members.sort((a, b) => Number(a.code) - Number(b.code));

  // Report
  const prevByCode = new Map((prev.members || []).map(m => [String(m.code), m]));
  const added = members.filter(m => !prevByCode.has(String(m.code)));
  const removed = (prev.members || []).filter(m => !seen.has(String(m.code)));
  const moved = members.filter(m => prevByCode.has(String(m.code)) && prevByCode.get(String(m.code)).majlis !== m.majlis);
  console.log('Members in Excel: ' + members.length + ' (before: ' + prevByCode.size + ')');
  console.log('New: ' + added.length);            added.forEach(m => console.log('  ' + m.code + '  ' + m.name + '  (' + m.majlis + ')'));
  console.log('No longer listed: ' + removed.length); removed.forEach(m => console.log('  ' + m.code + '  ' + m.name + '  (' + m.majlis + ')'));
  console.log('Changed Majlis: ' + moved.length);  moved.forEach(m => console.log('  ' + m.code + '  ' + m.name + ': ' + prevByCode.get(String(m.code)).majlis + ' -> ' + m.majlis));

  if (missing.length) {
    console.log('\nNEED DOB AND MUSI for ' + missing.length + ' member(s), add them to --extra and run again:');
    missing.forEach(m => console.log('  ' + m.code + '  ' + m.name + '  (' + m.majlis + ')'));
    process.exit(2);
  }

  // Ansars-to-be
  let tobe = prev.ansarsToBe || [];
  if (args.tobe) tobe = readRows(args.tobe).map(r => ({
    code: Number(clean(pick(r, 'code', 'tajneed'))) || clean(pick(r, 'code', 'tajneed')),
    name: clean(pick(r, 'name')),
    majlis: clean(pick(r, 'majlis', 'jamaat')),
    birthday: toIso(pick(r, 'birthday', 'dob')),
    email: clean(pick(r, 'email')),
    mobile: clean(pick(r, 'mobile')),
  })).filter(t => t.code);
  else tobe = tobe.map(t => Object.assign({}, t, { birthday: toIso(t.birthday) }));

  // Removed members keep their DOB/Musi in case they come back.
  const archive = Object.assign({}, prev.archive || {});
  for (const m of removed) archive[m.code] = { code: m.code, name: m.name, birthday: toIso(m.birthday), musi: !!m.musi };
  for (const m of members) delete archive[m.code];

  const out = {
    source: path.basename(args.excel),
    updated: NOW.toISOString(),
    members,
    ansarsToBe: tobe,
    archive,
  };
  console.log('\nMusi: ' + members.filter(m => m.musi).length + ' | Ansars-to-be: ' + tobe.length);
  if (args['dry-run']) { console.log('Dry run: nothing written.'); return; }
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, JSON.stringify(await encrypt(out, pass)) + '\n');
  console.log('Wrote ' + path.relative(ROOT, OUT));
})().catch(e => { console.error(e.message || e); process.exit(1); });
