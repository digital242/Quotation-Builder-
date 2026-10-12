// Parser, pricing, reference and Excel tests. Run: node tests/run-tests.mjs
// Excel layout is checked by re-reading the file with ExcelJS, then printed to
// PDF with LibreOffice (if installed) to prove it fits one A4 page and that the
// H-column formulas recalculate.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { loadCore, ROOT } from './core.mjs';

const require = createRequire(import.meta.url);
const ExcelJS = require('exceljs');
const TD = loadCore();
const OUT = path.join(ROOT, 'tests', 'out');
fs.mkdirSync(OUT, { recursive: true });

let pass = 0, fail = 0;
const results = [];
async function test(name, fn) {
  try { await fn(); pass++; results.push('PASS  ' + name); }
  catch (e) { fail++; results.push('FAIL  ' + name + '\n      ' + (e && e.message || e)); }
}
function eq(a, b, msg) { if (a !== b) throw new Error((msg ? msg + ': ' : '') + 'expected ' + JSON.stringify(b) + ', got ' + JSON.stringify(a)); }
function ok(c, msg) { if (!c) throw new Error(msg || 'assertion failed'); }
const near = (a, b, msg) => { if (Math.abs(a - b) > 0.005) throw new Error((msg || '') + ` expected ${b}, got ${a}`); };

/* Mirror of the UI's merge logic: same product+frame+size+load adds qty. */
function addAll(items, lines) {
  const bad = [];
  for (const l of lines) {
    const r = TD.parseLine(l);
    if (!r.ok) { bad.push(r); continue; }
    const ex = items.find(x => TD.itemKey(x) === TD.itemKey(r.item));
    if (ex) ex.qty += r.item.qty; else items.push({ ...r.item });
  }
  return bad;
}
const FIELDS = { company: 'Acme Infra Pvt Ltd', city: 'Bengaluru', cityCode: 'BGLR', contact: 'Mr. Rao', mobile: '9876543210',
  email: 'rao@acme.example', site: 'Whitefield metro depot', enqBy: 'Email', enqDate: '2026-09-20', serial: 25, qdate: '2026-09-24',
  delivery: '3-4 weeks from PO', validity: '30 days' };

/* ---------------- 1. The four-line paste ---------------- */
const FOUR = [
  'solid top cover 600x600 B125 Qnty 4',
  'Solid top cover 1000x1000 D400 quantity 2',
  'Water gully cover 395x990 5 tons quantity 2',
  'recess cover 900x600 5-ton quantity 4'
];
await test('1. four pasted lines add 4 items at the right prices', () => {
  const rs = FOUR.map(l => TD.parseLine(l));
  rs.forEach((r, i) => ok(r.ok, 'line ' + (i + 1) + ' failed: ' + r.error));
  const exp = [['solid', 's600', 2, null, 4, 12345, 49380], ['solid', 's1000', 4, null, 2, 69140, 138280],
    ['gully', 'r335x930', 1, 'without', 2, 4200, 8400], ['recess', 'r900x600', 1, null, 4, 17740, 70960]];
  rs.forEach((r, i) => {
    const [prod, id, load, frame, qty, price, amt] = exp[i];
    eq(r.item.prod, prod); eq(r.item.id, id); eq(r.item.load, load); eq(r.item.frame, frame); eq(r.item.qty, qty);
    eq(r.price, price); eq(r.amount, amt);
  });
  eq(TD.describeParsed(rs[0]), 'Solid Top MHC, 600 × 600 mm, B125 — 4 nos × ₹12,345 = ₹49,380');
  const t = TD.totals(rs.map(r => r.item), 0, true);
  eq(TD.rupees2(t.sub), '₹2,67,020.00'); eq(TD.rupees2(t.gst), '₹48,063.60'); eq(TD.rupees2(t.grand), '₹3,15,083.60');
  eq(TD.rupees(rs[1].amount), '₹1,38,280');
});
await test('2. air water tight 750x750 c250 3 nos → 3 × ₹27,695 = ₹83,085', () => {
  const r = TD.parseLine('air water tight 750x750 c250 3 nos');
  ok(r.ok, r.error); eq(r.item.prod, 'aw'); eq(r.item.id, 's750'); eq(r.item.load, 3); eq(r.item.qty, 3); eq(r.price, 27695); eq(r.amount, 83085);
});
await test('3. solid 777x777 b125 → not-in-price-list error', () => {
  const r = TD.parseLine('solid 777x777 b125 qty 1');
  ok(!r.ok); eq(r.error, '777 × 777 is not in the Solid Top MHC price list.');
});
await test('4. recess 600x600 e600 → on-request error', () => {
  const r = TD.parseLine('recess 600x600 e600 qty 1');
  ok(!r.ok); eq(r.error, "E600 is on request for this size, so it can't be priced here.");
});
await test('5. gully with frame → ₹7,300 in its own WITH FRAME group', () => {
  const r = TD.parseLine('water gully 395x990 5 ton with frame qty 2');
  ok(r.ok, r.error); eq(r.item.frame, 'with'); eq(r.price, 7300); eq(r.item.qty, 2);
  const items = [];
  addAll(items, ['Water gully cover 395x990 5 tons quantity 2', 'water gully 395x990 5 ton with frame qty 2']);
  const M = TD.buildModel(items, FIELDS, 0, true);
  eq(M.groups.length, 2);
  eq(M.groups[0].title, '‘THERMODRAIN’ FRP WATER GULLY COVER WITHOUT FRAME (with 2 side support & inbuilt beam design)');
  eq(M.groups[1].title, '‘THERMODRAIN’ FRP WATER GULLY COVER WITH FRAME');
  eq(M.groups[1].items[0].price, 7300);
  eq(M.sizeHeader, 'OUT TO OUT SIZE');
  eq(M.groups[0].items[0].wording, '395mm W. X 990mm L. (out to out)');
});
await test('6. tank 600 qty 3 → ₹5,420 at 500 KG', () => {
  const r = TD.parseLine('tank 600 qty 3');
  ok(r.ok, r.error); eq(r.item.prod, 'tank'); eq(r.item.id, 'c600'); eq(r.price, 5420); eq(r.item.qty, 3);
  eq(TD.loadPrint('tank', r.item.load), '500 KG');
});
await test('7. same item twice merges into one line', () => {
  const items = [];
  addAll(items, ['solid top 600x600 B125 qty 4', 'manhole cover 600 x 600 mm b 125 2 nos']);
  eq(items.length, 1); eq(items[0].qty, 6);
});
await test('10. reference number + Indian financial year', () => {
  eq(TD.refNo({ series: 'TP1', cityCode: 'BGLR', serial: 25, qdate: '2026-09-24' }), 'THERMODRAIN/QTN/TP1-BGLR/025/2026-27');
  eq(TD.refNo({ series: 'TP1', cityCode: '', serial: 7, qdate: '2027-02-10' }), 'THERMODRAIN/QTN/TP1/007/2026-27');
  eq(TD.fyOf('2027-04-01'), '2027-28'); eq(TD.fyOf('2027-03-31'), '2026-27');
});

/* ---------------- Extra parser coverage ---------------- */
await test('parser: load words, boundaries and 2.5 before 5 T', () => {
  const L = l => TD.parseLine(l);
  eq(L('solid 600x600 2.5 t qty 2').item.load, 0);
  eq(L('solid 600x600 pedestrian').item.load, 0);
  eq(L('solid 600x600 5 mt').item.load, 1);
  eq(L('solid 600x600 light vehicle').item.load, 1);
  eq(L('solid 600x600 2 wheeler').item.load, 1);
  eq(L('solid 600x600 4 wheeler').item.load, 2);
  eq(L('solid 600x600 12.5 t').item.load, 2);
  eq(L('solid 600x600 truck').item.load, 3);
  eq(L('solid 600x600 fire engine').item.load, 4);
  eq(L('solid 600x600 trailor').item.load, 4);
  eq(L('solid 600x600 60 t').item.load, 5);
  eq(L('solid 600x600 for port').item.load, 5);
  const r = L('solid 1500x1500 qty 2');
  ok(!r.ok); eq(r.error, 'Add the load class: 2.5 T, 5 T, B125, C250, D400 or E600.');
  const r2 = L('solid 1500x1500 d400 2'); eq(r2.item.id, 's1500'); eq(r2.item.qty, 2);
});
await test('parser: errors are plain sentences', () => {
  eq(TD.parseLine('600x600 b125').error, 'Name the product: solid top, recess, gully, air & water tight, hinged or tank.');
  eq(TD.parseLine('solid top b125 qty 3').error, 'Add the size, for example 600x600 or 900 dia.');
});
await test('parser: circular, × and *, twin preference, gully out-to-out map', () => {
  const r = TD.parseLine('solid 900 dia D400 qty 1'); eq(r.item.id, 'c900'); eq(r.price, 31895);
  eq(TD.parseLine('solid 600×600 b125').item.id, 's600');
  eq(TD.parseLine('solid 600*450 b125').item.id, 'r600x450');
  eq(TD.parseLine('solid 450x600 b125').item.id, 'r600x450');
  eq(TD.parseLine('solid 900x900 2.5t').item.id, 's900');
  eq(TD.parseLine('gully 342x342 lv').item.id, 's300');
  eq(TD.parseLine('gully 360x360 b125').item.id, 's300');
  eq(TD.parseLine('gully 660 dia b125').item.id, 'c600');
  eq(TD.parseLine('grating cover 515x965 c250').item.id, 'r900x450');
  eq(TD.parseLine('water tight 600 b125').item.prod, 'aw');
  eq(TD.parseLine('hinged 600x600 b125').item.prod, 'hinged');
  eq(TD.parseLine('solid 600 square b125').item.id, 'c600'); // a single number means circular first
  eq(TD.parseLine('solid 675 b125').item.id, 's675');        // falls back to square N×N
});
await test('price-review flags', () => {
  eq(TD.flagFor({ prod: 'gully', id: 's900', load: 4, frame: 'without' }), TD.FLAGS['gully|s900|4|without']);
  eq(TD.flagFor({ prod: 'gully', id: 's900', load: 4, frame: 'with' }), null);
  ok(TD.flagFor({ prod: 'solid', id: 'r900x600', load: 3 }));
  ok(TD.flagFor({ prod: 'recess', id: 'c1000', load: 3 }));
});
await test('price data is untouched (spot checks + null = on request)', () => {
  eq(TD.priceOf('solid', 's1500', 5), 222770);
  eq(TD.priceOf('aw', 's1500', 5), 289600);
  eq(TD.priceOf('recess', 's1500', 5), null);
  eq(TD.priceOf('solid', 's900t', 1), null);
  eq(TD.priceOf('gully', 'c600', 0, 'with'), 8295);
  eq(TD.priceOf('gully', 'c600', 0, 'without'), 7200);
  eq(TD.priceOf('tank', 's750', 3), 7525);
});
await test('wording, subject and intro', () => {
  eq(TD.defaultWording({ prod: 'solid', id: 's600', load: 2 }), '600 mm x 600 mm (24" x 24") clear opening');
  eq(TD.defaultWording({ prod: 'solid', id: 'c600', load: 2 }), '600 mm dia (24") clear opening');
  eq(TD.defaultWording({ prod: 'solid', id: 's900t', load: 0 }), '900 mm x 900 mm (36" x 36") clear opening, 2 covers with single frame');
  eq(TD.defaultWording({ prod: 'gully', id: 's600', load: 2, frame: 'with' }), '660mm X 660mm (out to out)');
  eq(TD.defaultWording({ prod: 'gully', id: 'c600', load: 2, frame: 'with' }), '660mm DIA (out to out)');
  eq(TD.defaultWording({ prod: 'gully', id: 's300', load: 1 }), '342mm X 342mm (out to out)');
  eq(TD.defaultWording({ prod: 'gully', id: 's300', load: 2 }), '360mm X 360mm (out to out)');
  const items = []; addAll(items, FOUR); addAll(items, ['air water tight 750x750 c250 3 nos']);
  eq(TD.autoSubject(items, 'Whitefield metro depot'),
    'OFFER FOR THERMODRAIN FRP SOLID TOP MANHOLE COVER, WATER GULLY COVER, RECESS COVER & AIR & WATER TIGHT COVER FOR WHITEFIELD METRO DEPOT');
  eq(TD.introText({ enqBy: 'Phone call', enqDate: '2026-09-20' }),
    'This is in reference to our telephonic discussion dt.20/09/2026 regarding the above mentioned subject. Kindly find our offer (for supply only) as given below:');
  eq(TD.fmtIN(41312390, 2), '4,13,12,390.00'); eq(TD.rupees2(413123.9), '₹4,13,123.90');
});

/* ---------------- Excel ---------------- */
async function makeXlsx(name, lines, discount, gst, fields = FIELDS) {
  const items = []; addAll(items, lines);
  const M = TD.buildModel(items, fields, discount, gst);
  const { wb, totalRowsAt, firstItemRow, lastItemRow, lastRow } = TD.buildWorkbook(ExcelJS, M, {});
  const file = path.join(OUT, name);
  await wb.xlsx.writeFile(file);
  const rb = new ExcelJS.Workbook(); await rb.xlsx.readFile(file);
  return { M, file, ws: rb.getWorksheet('Quotation'), totalRowsAt, firstItemRow, lastItemRow, lastRow };
}
const ALL = [...FOUR, 'air water tight 750x750 c250 3 nos', 'water gully 395x990 5 ton with frame qty 2', 'tank 600 qty 3',
  'hinged 600 b125 qty 1', 'solid 900x900 2.5 t twin qty 1'];

let x48;
await test('8. 48% discount: total rows, labels, formulas and computed results', async () => {
  x48 = await makeXlsx('QTN_discount48.xlsx', FOUR, 48, true);
  const { ws, totalRowsAt: at, firstItemRow: f, lastItemRow: l } = x48;
  const rows = ['total', 'disc', 'after', 'gst', 'grand'].map(k => at[k]);
  rows.forEach((r, i) => i && eq(r, rows[i - 1] + 1, 'consecutive total rows'));
  const lab = r => ws.getCell('E' + r).value;
  eq(lab(at.total), 'TOTAL'); eq(lab(at.disc), 'Less Discount @48%'); eq(lab(at.after), 'TOTAL');
  eq(lab(at.gst), 'ADD : GST @ 18%'); eq(lab(at.grand), 'GRAND TOTAL (INCL. 18% GST)');
  ok(ws.getCell('E' + at.after).font.bold, 'TOTAL after discount bold');
  ok(!ws.getCell('E' + at.total).font.bold, 'first TOTAL not bold');
  const H = r => ws.getCell('H' + r).value;
  eq(H(at.total).formula, `SUM(H${f}:H${l})`); near(H(at.total).result, 267020);
  eq(H(at.disc).formula, `H${at.total}*48%`); near(H(at.disc).result, 128169.6);
  eq(H(at.after).formula, `H${at.total}-H${at.disc}`); near(H(at.after).result, 138850.4);
  eq(H(at.gst).formula, `H${at.after}*18%`); near(H(at.gst).result, 24993.07);
  eq(H(at.grand).formula, `H${at.after}+H${at.gst}`); near(H(at.grand).result, 163843.47);
  for (let c = 1; c <= 8; c++) eq(ws.getCell(at.grand, c).fill?.fgColor?.argb, 'FFFFF2CC', 'yellow fill col ' + c);
  ok(!ws.getCell('H' + at.gst).fill || ws.getCell('H' + at.gst).fill.pattern === 'none', 'GST row not yellow');
  eq(ws.getRow(at.grand).height, 20);
  for (let r = f; r <= l; r++) { const v = H(r); if (v && v.formula) { eq(v.formula, `G${r}*F${r}`); near(v.result, ws.getCell('G' + r).value * ws.getCell('F' + r).value); } }
});
await test('Excel: sheet setup (name, A4, fit 1×1, margins, gridlines, widths, print area, fonts)', async () => {
  const { ws, lastRow } = x48;
  eq(ws.name, 'Quotation');
  eq(ws.pageSetup.paperSize, 9); eq(ws.pageSetup.orientation, 'portrait');
  eq(ws.pageSetup.fitToPage, true); eq(ws.pageSetup.fitToWidth, 1); eq(ws.pageSetup.fitToHeight, 1);
  eq(ws.pageSetup.horizontalCentered, true);
  eq(ws.pageSetup.margins.left, 0.45); eq(ws.pageSetup.margins.right, 0.45); eq(ws.pageSetup.margins.top, 0.5); eq(ws.pageSetup.margins.bottom, 0.5);
  eq(ws.views[0].showGridLines, false);
  [10.57, 1.57, 18.71, 29.57, 11.71, 10.29, 12.29, 14].forEach((w, i) => eq(ws.getColumn(i + 1).width, w));
  eq(ws.pageSetup.printArea, 'A1:H' + lastRow);
  ws.eachRow(row => row.eachCell(c => { if (c.value !== null && c.value !== undefined && c.font) { eq(c.font.name, 'Calibri'); eq(c.font.size, 11); } }));
  eq(ws.getCell('A4').value, 'Office address: Raheja Arcade, Off.No.613, Sec – 11, Plot no. 61, CBD Belapur, Navi Mumbai – 400 614.');
  eq(ws.getCell('A6').font.color.argb, 'FF0070C0'); eq(ws.getCell('F6').alignment.horizontal, 'right');
  eq(ws.getRow(7).height, 4); eq(ws.getCell('D7').border.bottom.style, 'medium'); eq(ws.getCell('D7').border.bottom.color.argb, 'FF1F3A68');
  const find = txt => { let hit = null; ws.eachRow((row, rn) => row.eachCell(c => { if (!hit && c.value === txt) hit = c; })); return hit; };
  const date = find('DATE'); ok(date, 'DATE row');
  const dv = ws.getCell(date.row, 3);
  ok(dv.value instanceof Date, 'date is a real date'); eq(dv.numFmt, 'dd/mm/yyyy'); eq(dv.value.toISOString().slice(0, 10), '2026-09-24');
  eq(ws.getCell(find('OUR REF').row, 3).value, 'THERMODRAIN/QTN/TP1-BGLR/025/2026-27');
  const em = ws.getCell(find('EMAIL ON').row, 3); eq(em.font.underline, true);
  ok(find('ACME INFRA PVT LTD').font.bold); ok(find('BENGALURU.')); ok(find('MOB.NO.: 9876543210'));
  eq(ws.getCell(find('KIND ATTN').row, 3).value, 'MR. RAO');
  const sub = find('SUB'); eq(sub.alignment.vertical, 'top'); eq(ws.getCell(sub.row, 3).alignment.vertical, 'top'); eq(ws.getCell(sub.row, 2).alignment.vertical, 'top');
  ok(ws.getCell(sub.row, 3).font.bold);
  ok(find('FRAME OPENING SIZE') === null, 'mixed quote uses SIZE header'); ok(find('SIZE'));
  ok(find('TERMS & CONDITIONS').font.bold);
  const gst = find('GST'); eq(gst.col, 4); eq(ws.getCell(gst.row, 5).value, ': 18% (included in grand total above)');
  ok(find('DELIVERY') && find('VALIDITY'));
  ok(find('PRAVIN SAKPAL').font.bold); ok(find('MARKETING DIRECTOR')); ok(find('M.NO. +91 9619908011'));
  const styles = execFileSync('unzip', ['-p', x48.file, 'xl/styles.xml']).toString().replace(/&gt;/g, '>');
  ok(styles.includes('formatCode="' + TD.NUMFMT + '"'), 'Indian-grouping number format in styles.xml');
  eq(ws.getCell('G' + (x48.firstItemRow + 2)).numFmt, TD.NUMFMT.replace(/\\/g, ''));
});
await test('Excel: GST off → last TOTAL is bold+yellow and term says 18% Extra', async () => {
  const x = await makeXlsx('QTN_gst_off.xlsx', FOUR, 0, false);
  const at = x.totalRowsAt; eq(Object.keys(at).join(','), 'total');
  ok(x.ws.getCell('E' + at.total).font.bold); eq(x.ws.getCell('A' + at.total).fill.fgColor.argb, 'FFFFF2CC');
  let term = null; x.ws.eachRow(row => row.eachCell(c => { if (c.value === ': 18% Extra') term = c; })); ok(term, '18% Extra');
});
await test('Excel: wrapped rows are tall enough (row-height rule)', async () => {
  const x = await makeXlsx('QTN_all.xlsx', ALL, 5, true, { ...FIELDS, site: 'Very long project name for the new integrated multi-modal logistics park at JNPT, Navi Mumbai, Phase II' });
  const ws = x.ws; const W = (a, b) => TD.COLW.slice(a - 1, b).reduce((s, v) => s + v, 0);
  ws.eachRow((row, rn) => {
    row.eachCell(c => {
      if (!c.alignment || !c.alignment.wrapText || typeof c.value !== 'string' || c.isMerged && c.master !== c) return;
      let c2 = c.col; while (c2 < 8 && ws.getCell(rn, c2 + 1).isMerged && ws.getCell(rn, c2 + 1).master.address === c.address) c2++;
      const need = TD.rowHeightFor(c.value, W(c.col, c2), !!(c.font && c.font.bold), 15);
      const hdr = rn === 0; if (hdr) return;
      ok((row.height || 15) * (ws.getCell(rn + 1, c.col).isMerged && ws.getCell(rn + 1, c.col).master.address === c.address ? 2 : 1) >= need - 0.01 || (c.value.length < 20),
        `row ${rn} "${c.value.slice(0, 40)}" height ${row.height} < ${need}`);
    });
  });
});

/* ---------------- LibreOffice: one page + live formulas ---------------- */
function soffice() { for (const p of ['/usr/bin/soffice', '/usr/bin/libreoffice']) if (fs.existsSync(p)) return p; return null; }
const SO = soffice();
function convert(file, fmt, outdir) {
  const profile = 'file://' + path.join(os.tmpdir(), 'td-lo-profile');
  execFileSync(SO, ['-env:UserInstallation=' + profile, '--headless', '--convert-to', fmt, '--outdir', outdir, file], { stdio: 'pipe', timeout: 120000 });
}
await test('9a. LibreOffice prints each generated file on exactly one A4 page', async () => {
  if (!SO) throw new Error('LibreOffice not installed: skipped');
  for (const f of ['QTN_discount48.xlsx', 'QTN_all.xlsx', 'QTN_gst_off.xlsx']) {
    convert(path.join(OUT, f), 'pdf', OUT);
    const info = execFileSync('pdfinfo', [path.join(OUT, f.replace('.xlsx', '.pdf'))]).toString();
    eq(info.match(/Pages:\s+(\d+)/)[1], '1', f + ' pages');
    ok(/Page size:\s+595\.\d* x 841\.89/.test(info), f + ' is A4: ' + info.match(/Page size:.*/)[0]);
  }
});
await test('9b. H-column formulas recalculate after a qty is edited', async () => {
  if (!SO) throw new Error('LibreOffice not installed: skipped');
  const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(x48.file);
  const ws = wb.getWorksheet('Quotation');
  const r = x48.firstItemRow + 2; // first item row (after group title + note)
  eq(ws.getCell('G' + r).value, 12345);
  ws.getCell('F' + r).value = 10; // was 4
  ws.eachRow(row => row.eachCell(c => { if (c.value && c.value.formula) c.value = { formula: c.value.formula }; })); // drop cached results
  const edited = path.join(OUT, 'edited.xlsx'); await wb.xlsx.writeFile(edited);
  convert(edited, 'csv', OUT);
  const csv = fs.readFileSync(path.join(OUT, 'edited.csv'), 'utf8').split('\n').map(l => l.split(','));
  const num = s => +String(s || '').replace(/[^\d.]/g, '');
  const line = csv.find(c => c.includes('600 mm x 600 mm (24" x 24") clear opening') || c.some(v => v.includes('600 mm x 600 mm')));
  ok(line, 'item line in csv');
  const grand = fs.readFileSync(path.join(OUT, 'edited.csv'), 'utf8').split('\n').find(l => l.includes('GRAND TOTAL'));
  // new subtotal = 267020 + 6*12345 = 341090; after 48% = 177366.8; +18% = 209292.82
  ok(/209[,]?292\.82/.test(grand.replace(/"/g, '')), 'grand total recalculated: ' + grand);
});

console.log(results.join('\n'));
console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
