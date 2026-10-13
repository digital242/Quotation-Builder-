// Randomised consistency checks. Run: node tests/fuzz-tests.mjs
// Builds many random quotations (every product, frame, chamber, discounts, all three
// tax/currency modes) and checks that:
//  - every list price matches the price table directly,
//  - the model's totals match an independent calculation,
//  - each Excel formula's stored result matches the model,
//  - LibreOffice recalculates every H-column formula to the same value (sampled),
//  - each file prints on one A4 page (sampled).
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { loadCore, ROOT } from './core.mjs';

const require = createRequire(import.meta.url);
const ExcelJS = require('exceljs');
const TD = loadCore();
const OUT = path.join(ROOT, 'tests', 'out', 'fuzz'); fs.mkdirSync(OUT, { recursive: true });
const SO = ['/usr/bin/soffice', '/usr/bin/libreoffice'].find(p => fs.existsSync(p));

let seed = 20261013;
const rnd = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
const pick = a => a[Math.floor(rnd() * a.length)];
const r2 = n => Math.round((n + Number.EPSILON) * 100) / 100;
let checks = 0, failures = [];
const fail = (q, msg) => failures.push(`quote ${q}: ${msg}`);
const near = (a, b) => Math.abs(a - b) <= 0.011;

function randomItem() {
  if (rnd() < 0.12) {
    return { prod: 'custom', name: pick(['FRP Chamber', 'FRP Inspection Chamber', 'GRP Valve Chamber']),
      l: 300 + Math.floor(rnd() * 1500), w: 300 + Math.floor(rnd() * 1500), h: 200 + Math.floor(rnd() * 1200),
      finish: pick(['single', 'double']), frame: pick(['ms', 'frp', 'none']), load: pick([null, 0, 1, 2, 3, 4, 5]),
      price: Math.round(5000 + rnd() * 120000), qty: 1 + Math.floor(rnd() * 20), disc: pick([0, 0, 5, 7.5, 10, 12, 20]) };
  }
  for (;;) {
    const prod = pick(TD.PRODUCTS);
    const shapes = TD.SHAPES.filter(s => (TD.DATA[prod].shapes[s] || []).length);
    const row = pick(TD.DATA[prod].shapes[pick(shapes)]);
    const load = prod === 'tank' ? 0 : Math.floor(rnd() * 6);
    const frame = prod === 'gully' ? pick(['with', 'without']) : pick(['frp', 'frp', 'frp', 'ms', 'none']);
    const it = { prod, id: row[0], load, frame, qty: 1 + Math.floor(rnd() * 40), disc: pick([0, 0, 0, 5, 10, 15, 48]),
      depth: prod === 'recess' && rnd() < 0.5 ? pick([50, 75, 100]) : null };
    if (TD.isManual(it)) it.price = Math.round(2000 + rnd() * 90000);
    if (TD.unitPrice(it) === null) continue; // on request: the app never lets these in
    // independent list-price check straight from the table
    if (!TD.isManual(it)) {
      const arr = prod === 'gully' ? (frame === 'with' ? row[3] : row[4]) : row[3];
      const expect = prod === 'tank' ? row[3][0] : arr[load];
      checks++; if (TD.unitPrice(it) !== expect) failures.push(`price ${prod} ${row[0]} ${load} ${frame}: ${TD.unitPrice(it)} != ${expect}`);
    }
    return it;
  }
}

const FIELDS = { company: 'Fuzz Test Co', city: 'Pune', cityCode: 'PNQ', contact: 'Ms Shah', mobile: '9000000000', email: 'a@b.example',
  site: 'Plant 4', enqBy: 'Email', enqDate: '2026-10-01', serial: 7, qdate: '2026-10-13', usdRate: '88.37',
  paySlabs: [{ pct: 30, stage: 'ADVANCE' }, { pct: 70, stage: 'BEFORE DISPATCH' }] };

const N = 60, SAMPLE = 8;
for (let q = 0; q < N; q++) {
  const n = 1 + Math.floor(rnd() * 9);
  const items = Array.from({ length: n }, randomItem);
  const mode = pick(['gst', 'gst', 'nogst', 'usd']);
  const M = TD.buildModel(items, FIELDS, mode);
  const rate = 88.37;
  // independent totals
  let sub = 0, disc = 0;
  for (const it of items) {
    const p = mode === 'usd' ? r2(TD.unitPrice(it) / rate) : TD.unitPrice(it);
    sub += p * it.qty; disc += p * it.qty * (it.disc || 0) / 100;
  }
  const base = sub - disc, gst = mode === 'gst' ? base * 0.18 : 0, grand = base + gst;
  checks++; if (!near(M.totals.sub, sub)) fail(q, `sub ${M.totals.sub} != ${sub}`);
  checks++; if (!near(M.totals.grand, grand)) fail(q, `grand ${M.totals.grand} != ${grand}`);
  checks++; if (mode !== 'gst' && M.totals.gst !== 0) fail(q, 'GST charged outside GST mode');
  checks++; if (mode === 'gst' && M.totalRows.some(r => /US\$|USD/.test(r.label))) fail(q, 'USD shown in GST mode');
  checks++; if (mode === 'usd' && M.totalRows.some(r => /GST/.test(r.label))) fail(q, 'GST row in USD mode');
  checks++; if (M.groups.reduce((a, g) => a + g.items.length, 0) !== items.length) fail(q, 'item count');
  checks++; if (M.groups.flatMap(g => g.items).some((x, i) => x.sr !== i + 1)) fail(q, 'serial numbers not running 1..n');
  for (const g of M.groups) for (const x of g.items) { checks++; if (!x.wording || /undefined|NaN/.test(x.wording + x.load + g.title)) fail(q, `bad text: ${g.title} / ${x.wording} / ${x.load}`); }

  const { wb, totalRowsAt, firstItemRow, lastItemRow } = TD.buildWorkbook(ExcelJS, M, TD.logos);
  // stored formula results agree with the model
  const ws = wb.getWorksheet('Quotation');
  let hsum = 0;
  for (let r = firstItemRow; r <= lastItemRow; r++) { const v = ws.getCell('H' + r).value; if (v && v.formula) { checks++; const expect = ws.getCell('G' + r).value * ws.getCell('F' + r).value; if (!near(v.result, expect)) fail(q, `row ${r} amount`); hsum += v.result; } }
  checks++; if (!near(ws.getCell('H' + totalRowsAt.total).value.result, hsum)) fail(q, 'TOTAL row');
  const lastKey = Object.keys(totalRowsAt).pop();
  checks++; if (!near(ws.getCell('H' + totalRowsAt[lastKey]).value.result, M.totals.grand)) fail(q, 'last total row != grand');
  const file = path.join(OUT, `fuzz_${q}.xlsx`);
  await wb.xlsx.writeFile(file);

  if (SO && q < SAMPLE) {
    // drop cached results so LibreOffice must recalculate, then compare every H value
    const wb2 = new ExcelJS.Workbook(); await wb2.xlsx.readFile(file);
    const ws2 = wb2.getWorksheet('Quotation');
    const expected = {};
    ws2.eachRow((row, rn) => { const c = row.getCell(8); if (c.value && c.value.formula) { expected[rn] = c.value.result; c.value = { formula: c.value.formula }; } });
    const nc = path.join(OUT, `fuzz_${q}_nc.xlsx`); await wb2.xlsx.writeFile(nc);
    const prof = 'file://' + path.join(os.tmpdir(), 'td-lo-profile');
    execFileSync(SO, ['-env:UserInstallation=' + prof, '--headless', '--convert-to', 'csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false', '--outdir', OUT, nc], { stdio: 'pipe', timeout: 120000 });
    const lines = fs.readFileSync(nc.replace('.xlsx', '.csv'), 'utf8').split('\n');
    for (const [rn, exp] of Object.entries(expected)) {
      const cells = (lines[rn - 1] || '').match(/("([^"]|"")*"|[^,]*)(,|$)/g) || [];
      const raw = (cells[7] || '').replace(/[",]/g, '').replace(/[^\d.\-]/g, '');
      checks++; if (!near(+raw, exp)) fail(q, `LibreOffice row ${rn}: ${raw} != ${exp}`);
    }
    execFileSync(SO, ['-env:UserInstallation=' + prof, '--headless', '--convert-to', 'pdf', '--outdir', OUT, file], { stdio: 'pipe', timeout: 120000 });
    const pages = execFileSync('pdfinfo', [file.replace('.xlsx', '.pdf')]).toString().match(/Pages:\s+(\d+)/)[1];
    checks++; if (pages !== '1') fail(q, `${pages} pages (${n} items)`);
  }
}
console.log(`${checks} checks on ${N} random quotations, ${failures.length} failures`);
failures.slice(0, 30).forEach(f => console.log('  FAIL ' + f));
process.exit(failures.length ? 1 : 0);
