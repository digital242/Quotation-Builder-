// Import-enquiry tests (PDF, Excel, CSV). Run: node tests/import-tests.mjs
// The pdf.js CDN files are served from node_modules/pdfjs-dist (same version, 3.11.174).
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { INDEX, ROOT, loadCore } from './core.mjs';

const require = createRequire(import.meta.url);
const { chromium } = require('playwright');
const EXCELJS_JS = fs.readFileSync(require.resolve('exceljs/dist/exceljs.min.js'));
const PDFJS_DIR = path.dirname(require.resolve('pdfjs-dist/build/pdf.min.js'));
const FX = f => path.join(ROOT, 'tests', 'fixtures', f);
const TD = loadCore();

let pass = 0, fail = 0; const results = [];
async function test(name, fn) { try { await fn(); pass++; results.push('PASS  ' + name); } catch (e) { fail++; results.push('FAIL  ' + name + '\n      ' + String(e && e.stack || e).split('\n').slice(0, 2).join(' ')); } }
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ', got ' + JSON.stringify(a)); };
const ok = (c, m) => { if (!c) throw new Error(m || 'assertion failed'); };

await test('core: product rows picked out of a BOQ; item numbers stripped; wrapped sizes joined', () => {
  const rows = ['TENDER ENQUIRY', 'S.No Description Qty Unit', '1. Supply of FRP manhole cover 600 x 600 mm clear opening, D400, 10 Nos',
    '2) Water gully cover 395x990 C250 6 nos', '3 Recess cover for paving', '900x600 mm, 5 ton 4 nos', 'a) Civil works, excavation 1000 cum',
    '600 dia manhole cover B125 2 nos', 'Delivery within 4 weeks', '1.2 Supply of FRP manhole cover 600 x 600 mm clear opening, D400, 10 Nos'];
  const out = TD.enquiryLines(rows);
  eq(out.length, 5);
  eq(out[0], 'Supply of FRP manhole cover 600 x 600 mm clear opening, D400, 10 Nos');
  eq(out[2], 'Recess cover for paving 900x600 mm, 5 ton 4 nos', 'size on the next row is joined');
  eq(out[3], '600 dia manhole cover B125 2 nos', 'a leading 600 is a size, not an item number');
  eq(out[4], out[0], 'the same item twice is kept (quantities add up)');
  ok(out.every(l => TD.checkLine(l).state === 'ready'), out.map(l => TD.checkLine(l).text).join(' / '));
});

const browser = await chromium.launch();
async function newPage() {
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  await ctx.route('https://cdn.jsdelivr.net/**', r => r.fulfill({ body: EXCELJS_JS, contentType: 'application/javascript' }));
  await ctx.route('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/*', r => r.fulfill({ body: fs.readFileSync(path.join(PDFJS_DIR, path.basename(new URL(r.request().url()).pathname))), contentType: 'application/javascript' }));
  await ctx.route(/fonts\.(googleapis|gstatic)\.com/, r => r.abort());
  const page = await ctx.newPage(); page.errors = []; page.on('pageerror', e => page.errors.push(e.message));
  await page.goto('file://' + INDEX); await page.waitForFunction(() => !!window.__tdApp && !!window.ExcelJS);
  await page.evaluate(() => localStorage.clear()); await page.reload(); await page.waitForFunction(() => !!window.__tdApp && !!window.ExcelJS);
  return { ctx, page };
}
const lastToast = p => p.locator('#toasts .toast').last().textContent();
async function importFile(page, file) {
  const before = await page.locator('#toasts .toast').count();
  await page.setInputFiles('#impFile', FX(file));
  await page.waitForFunction(n => document.querySelectorAll('#toasts .toast').length > n || !/^Reading/.test(document.querySelector('#impNote').textContent), before, { timeout: 30000 });
  await page.waitForFunction(() => !/^Reading/.test(document.querySelector('#impNote').textContent), null, { timeout: 30000 });
}

await test('Excel BOQ: products found with the right quantities, then added at list prices', async () => {
  const { ctx, page } = await newPage();
  await importFile(page, 'boq.xlsx');
  const lines = (await page.inputValue('#quick')).split('\n');
  eq(lines.length, 5, lines.join(' / '));
  ok(!lines.some(l => /Excavation/.test(l)), 'civil works are skipped');
  ok(lines[0].startsWith('Supply of FRP manhole cover 600 x 600 mm') && lines[0].endsWith('10 Nos'), lines[0]);
  eq(await page.locator('#lineCheck li.ready').count(), 5);
  ok((await page.textContent('#impNote')).includes('5 product lines found, 5 ready'));
  await page.press('#quick', 'Enter');
  const M = await page.evaluate(() => window.__tdApp.model());
  const items = M.groups.flatMap(g => g.items.map(i => [g.key, i.src.id, i.qty, i.price]));
  eq(JSON.stringify(items), JSON.stringify([['solid|frp', 's600', 14, 21365], ['gully|without', 'r335x930', 6, 12265], ['aw|frp', 's750', 3, 30245], ['fuel|frp', 's900cic', 2, 81095]]),
    'the repeated 600x600 row adds up to 14; every price is from the price list');
  eq(page.errors.length, 0, page.errors.join());
  await ctx.close();
});

await test('PDF BOQ (text PDF): rows read in order, item numbers and non-product rows dropped', async () => {
  const { ctx, page } = await newPage();
  await importFile(page, 'boqtable.pdf');
  const lines = (await page.inputValue('#quick')).split('\n');
  eq(lines.length, 3, lines.join(' / '));
  ok(/^Solid top manhole cover 600x600 B125\s+4 Nos$/.test(lines[0]), lines[0]);
  ok(lines[2].includes('Chamber 1000x1000x600 two sided B125 price 45000'), lines[2]);
  eq(await page.locator('#lineCheck li.ready').count(), 3);
  await page.press('#quick', 'Enter');
  eq(await page.locator('#sumItems .item').count(), 3);
  ok((await page.textContent('#sumItems')).includes('FRP Chamber'), 'chamber from the PDF added');
  eq(page.errors.length, 0, page.errors.join());
  await ctx.close();
});

await test('CSV adds to what is already typed; scanned PDF and wrong files give clear messages', async () => {
  const { ctx, page } = await newPage();
  await page.fill('#quick', 'tank 600 qty 1');
  await importFile(page, 'boq.csv');
  const lines = (await page.inputValue('#quick')).split('\n');
  eq(lines[0], 'tank 600 qty 1', 'existing text kept'); eq(lines.length, 4);
  await importFile(page, 'scanned.pdf');
  ok((await lastToast(page)).includes('looks like a scanned image with no text inside'), await lastToast(page));
  eq((await page.inputValue('#quick')).split('\n').length, 4, 'box unchanged after a failed import');
  // a file that is not a quotation format
  await page.setInputFiles('#impFile', { name: 'photo.png', mimeType: 'image/png', buffer: Buffer.from([137, 80, 78, 71]) });
  await page.waitForFunction(() => /Could not read/.test(document.querySelector('#impNote').textContent));
  ok((await lastToast(page)).includes('is not a PDF, Excel (.xlsx) or CSV file'));
  await page.setInputFiles('#impFile', { name: 'old.xls', mimeType: 'application/vnd.ms-excel', buffer: Buffer.from('x') });
  await page.waitForFunction(() => /old\.xls/.test(document.querySelector('#impNote').textContent));
  ok((await lastToast(page)).includes('save it as .xlsx'));
  eq(page.errors.length, 0, page.errors.join());
  await ctx.close();
});

await browser.close();
console.log(results.join('\n'));
console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
