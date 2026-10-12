// Browser tests for index.html. Run: node tests/ui-tests.mjs
// The ExcelJS CDN request is served from node_modules so the tests run offline;
// Google Fonts requests are aborted (the page must still work without them).
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { ROOT, INDEX } from './core.mjs';

const require = createRequire(import.meta.url);
const { chromium } = require('playwright');
const ExcelJS = require('exceljs');
const EXCELJS_JS = fs.readFileSync(require.resolve('exceljs/dist/exceljs.min.js'));
const OUT = path.join(ROOT, 'tests', 'out'); fs.mkdirSync(OUT, { recursive: true });
const URL_ = 'file://' + INDEX;

let pass = 0, fail = 0; const results = [];
async function test(name, fn) {
  try { await fn(); pass++; results.push('PASS  ' + name); }
  catch (e) { fail++; results.push('FAIL  ' + name + '\n      ' + String(e && e.stack || e).split('\n').slice(0,3).join(' ')); }
}
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ', got ' + JSON.stringify(a)); };
const ok = (c, m) => { if (!c) throw new Error(m || 'assertion failed'); };

const browser = await chromium.launch({ executablePath: fs.existsSync('/opt/pw-browsers/chromium') ? undefined : undefined });
async function newPage(opts = {}) {
  const ctx = await browser.newContext({ viewport: opts.viewport || { width: 1440, height: 900 }, colorScheme: opts.colorScheme || 'light',
    acceptDownloads: true, reducedMotion: 'reduce' });
  await ctx.route('https://cdn.jsdelivr.net/**', r => r.fulfill({ body: EXCELJS_JS, contentType: 'application/javascript' }));
  await ctx.route(/fonts\.(googleapis|gstatic)\.com/, r => r.abort());
  if (opts.blockStorage) await ctx.addInitScript(() => {
    Object.defineProperty(window, 'localStorage', { configurable: true, get() { throw new DOMException('blocked', 'SecurityError'); } });
  });
  const page = await ctx.newPage();
  page.errors = [];
  page.on('pageerror', e => page.errors.push(e.message));
  page.on('dialog', d => d.accept());
  await page.goto(URL_);
  await page.waitForFunction(() => !!window.ExcelJS && !!window.__tdApp);
  if (!opts.keepStorage) { await page.evaluate(() => { try { localStorage.clear(); } catch (e) {} }); await page.reload(); await page.waitForFunction(() => !!window.__tdApp); }
  return { ctx, page };
}
const FOUR = ['solid top cover 600x600 B125 Qnty 4', 'Solid top cover 1000x1000 D400 quantity 2',
  'Water gully cover 395x990 5 tons quantity 2', 'recess cover 900x600 5-ton quantity 4'];
async function typeLines(page, lines) {
  await page.fill('#quick', lines.join('\n'));
  await page.focus('#quick');
}
const lastToast = page => page.locator('#toasts .toast').last().textContent();

/* 1 */
await test('UI 1: paste four lines + Enter → 4 items and correct totals', async () => {
  const { ctx, page } = await newPage();
  await typeLines(page, FOUR);
  eq((await page.textContent('#quickAdd')).trim(), 'Add 4 items');
  eq(await page.locator('#parseList li.ok').count(), 4);
  ok((await page.textContent('#parseList')).includes('Solid Top MHC, 600 × 600 mm, B125 — 4 nos × ₹12,345 = ₹49,380'));
  await page.keyboard.press('Enter');
  eq(await page.locator('#sumItems .item').count(), 4);
  eq(await page.inputValue('#quick'), '');
  const tot = await page.textContent('#totals');
  ok(tot.includes('₹2,67,020.00'), 'subtotal ' + tot); ok(tot.includes('₹48,063.60'), 'gst');
  eq(await page.textContent('#grandVal'), '₹3,15,083.60');
  eq(await page.textContent('#sumCount'), '4 items');
  eq(await lastToast(page), '4 items added.');
  const amts = await page.locator('#sumItems .amt').allTextContents();
  eq(amts.join('|'), '₹49,380|₹1,38,280|₹8,400|₹70,960');
  eq(page.errors.length, 0, page.errors.join());
  await page.screenshot({ path: path.join(OUT, 'ui-desktop.png'), fullPage: true });
  await ctx.close();
});

/* 2 + 3 + 4 */
await test('UI 2-4: valid lines added, invalid lines stay with plain errors', async () => {
  const { ctx, page } = await newPage();
  await typeLines(page, ['air water tight 750x750 c250 3 nos', 'solid 777x777 b125 qty 1', 'recess 600x600 e600 qty 1']);
  await page.keyboard.press('Enter');
  eq(await page.locator('#sumItems .item').count(), 1);
  ok((await page.textContent('#sumItems')).includes('₹83,085'));
  eq(await page.inputValue('#quick'), 'solid 777x777 b125 qty 1\nrecess 600x600 e600 qty 1');
  const errs = await page.locator('#parseList li.err .msg').allTextContents();
  ok(errs[0].includes('777 × 777 is not in the Solid Top MHC price list.'), errs[0]);
  ok(errs[1].includes("E600 is on request for this size, so it can't be priced here."), errs[1]);
  eq(await lastToast(page), '1 item added. 2 lines still need fixing.');
  await ctx.close();
});

/* Shift+Enter, merge, tank */
await test('UI 5-7: Shift+Enter newline, gully with frame group, tank 500 kg, merge duplicates', async () => {
  const { ctx, page } = await newPage();
  await page.focus('#quick');
  await page.keyboard.type('water gully 395x990 5 ton with frame qty 2');
  await page.keyboard.press('Shift+Enter');
  await page.keyboard.type('tank 600 qty 3');
  eq((await page.inputValue('#quick')).split('\n').length, 2);
  await page.keyboard.press('Enter');
  await typeLines(page, ['tank 600 qty 3']); await page.click('#quickAdd');
  eq(await page.locator('#sumItems .item').count(), 2);
  const txt = await page.textContent('#sumItems');
  ok(txt.includes('Water Gully Cover with frame') && txt.includes('× ₹7,300'), 'gully with frame');
  ok(txt.includes('Ø 600 mm, 500 kg') && txt.includes('× ₹5,420'), 'tank');
  eq(await page.inputValue('#sumItems .item:nth-child(2) input[data-act=qty]'), '6');
  const M = await page.evaluate(() => window.__tdApp.model());
  eq(M.groups[0].title, '‘THERMODRAIN’ FRP WATER GULLY COVER WITH FRAME');
  eq(M.groups[1].items[0].load, '500 KG');
  await ctx.close();
});

await test('UI picker: gully flag, frame toggle, on-request loads disabled, step-down, add & merge', async () => {
  const { ctx, page } = await newPage();
  await page.click('[data-prod="gully"]');
  eq(await page.textContent('#sizeHint'), 'Sizes are cover out-to-out, as printed on the quotation');
  ok(await page.isVisible('#frameWrap'));
  await page.click('[data-shape="Square"]');
  await page.click('[data-size="s900"]');
  ok((await page.textContent('[data-size="s900"]')).includes('966 × 966') && (await page.textContent('[data-size="s900"]')).includes('opening 900 × 900'));
  ok(await page.isDisabled('[data-load="5"]')); ok((await page.textContent('[data-load="5"]')).includes('On request'));
  await page.click('[data-load="4"]');
  ok(await page.isVisible('#abFlag'), 'flag visible for gully|s900|D400|without');
  eq(await page.getAttribute('[data-load="4"]', 'aria-pressed'), 'true');
  await page.click('[data-frame="with"]');
  ok(await page.isHidden('#abFlag'), 'flag hidden with frame');
  eq(await page.getAttribute('[data-frame="with"]', 'aria-pressed'), 'true');
  // step-down: solid s900 E600 → twin row (only 2.5 T priced) → 2.5 T, never up
  await page.click('[data-prod="solid"]'); await page.click('[data-shape="Square"]'); await page.click('[data-size="s900"]');
  await page.click('[data-load="5"]');
  await page.click('[data-size="s900t"]');
  eq(await page.getAttribute('[data-load="0"]', 'aria-pressed'), 'true');
  ok((await page.textContent('[data-size="s900t"]')).includes('2 covers, 1 frame'));
  eq((await page.textContent('#abPrice')).startsWith('₹20,725'), true);
  await page.click('#abPlus'); await page.click('#abPlus');
  eq(await page.textContent('#abAdd'), 'Add 3 to quote, ₹62,175');
  await page.click('#abAdd'); await page.click('#abAdd');
  eq(await page.locator('#sumItems .item').count(), 1);
  eq(await page.inputValue('#sumItems input[data-act=qty]'), '6');
  // tank shows single 500 kg button
  await page.click('[data-prod="tank"]');
  eq(await page.locator('#loadGrid button').count(), 1);
  eq((await page.textContent('#loadGrid')).includes('500 kg'), true);
  await ctx.close();
});

await test('UI: download blocked without company name (focuses the field)', async () => {
  const { ctx, page } = await newPage();
  await typeLines(page, ['solid 600x600 b125 qty 1']); await page.keyboard.press('Enter');
  ok((await page.textContent('#ready')).startsWith('To download'));
  await page.click('#btnDownload');
  eq(await page.evaluate(() => document.activeElement.id), 'f_company');
  ok((await lastToast(page)).includes('company name'));
  await ctx.close();
});

await test('UI 8: download .xlsx (48% discount), filename, formulas, serial +1', async () => {
  const { ctx, page } = await newPage();
  await typeLines(page, FOUR); await page.keyboard.press('Enter');
  await page.fill('#f_company', 'Acme Infra Pvt. Ltd.');
  await page.fill('#f_cityCode', 'bglr');
  await page.click('#moreBox summary');
  await page.fill('#f_serial', '25');
  await page.fill('#f_qdate', '2026-09-24');
  eq(await page.textContent('#hdrRef'), 'THERMODRAIN/QTN/TP1-BGLR/025/2026-27');
  await page.fill('#discount', '48');
  ok((await page.textContent('#totals')).includes('Less discount @48%'));
  ok((await page.textContent('#ready')).startsWith('Ready.'));
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#btnDownload')]);
  eq(dl.suggestedFilename(), 'QTN_025_Acme_Infra_Pvt_Ltd.xlsx');
  const file = path.join(OUT, 'ui-download.xlsx'); await dl.saveAs(file);
  const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(file);
  const ws = wb.getWorksheet('Quotation');
  const labels = []; ws.eachRow((row, rn) => { const v = row.getCell(5).value; if (typeof v === 'string' && /TOTAL|Discount|GST @/.test(v)) labels.push([v, row.getCell(8).value, rn]); });
  eq(labels.map(l => l[0]).join(' | '), 'TOTAL | Less Discount @48% | TOTAL | ADD : GST @ 18% | GRAND TOTAL (INCL. 18% GST)');
  eq(labels[4][1].result, 163843.47);
  eq(await page.inputValue('#f_serial'), '26');
  eq(await page.textContent('#hdrRef'), 'THERMODRAIN/QTN/TP1-BGLR/026/2026-27');
  ok((await page.locator('#toasts .toast').allTextContents()).some(t => t === 'Saved. Next quotation will be no. 026.'));
  await ctx.close();
});

await test('UI: claude downloads API is used when present', async () => {
  const { ctx, page } = await newPage();
  await page.evaluate(() => { window.__saved = null; window.claude = { use: async (k) => k === 'downloads' ? { save: async (o) => { window.__saved = { name: o.filename, size: o.data.size }; } } : null }; });
  await typeLines(page, ['tank 600 qty 3']); await page.keyboard.press('Enter');
  await page.fill('#f_company', 'Tank Co');
  await page.click('#btnDownload');
  await page.waitForFunction(() => window.__saved);
  const s = await page.evaluate(() => window.__saved);
  eq(s.name, 'QTN_001_Tank_Co.xlsx'); ok(s.size > 4000);
  await ctx.close();
});

await test('UI: preview dialog, WhatsApp text, persistence, new quotation keeps terms', async () => {
  const { ctx, page } = await newPage();
  await typeLines(page, FOUR); await page.keyboard.press('Enter');
  await page.fill('#f_company', 'Acme'); await page.fill('#f_contact', 'Mr Rao');
  await page.click('#btnPreview');
  ok(await page.isVisible('#pvDlg'));
  const paper = await page.textContent('#paper');
  ok(paper.includes('GRAND TOTAL (INCL. 18% GST)') && paper.includes('3,15,083.60') && paper.includes('KIND ATTN'));
  await page.click('#pvClose');
  const wa = await page.evaluate(() => TD.whatsappText(window.__tdApp.model()));
  ok(wa.includes('*Grand total (incl. 18% GST): ₹3,15,083.60*'), wa);
  ok(wa.includes('*‘THERMODRAIN’ FRP/GRP SOLID TOP MANHOLE COVER WITH FRAME*'));
  await page.click('#btnCopy');
  await page.waitForTimeout(150);
  ok((await lastToast(page)).startsWith('Copied') || (await lastToast(page)).startsWith('Couldn'));
  // persistence
  await page.reload(); await page.waitForFunction(() => !!window.__tdApp);
  eq(await page.locator('#sumItems .item').count(), 4);
  eq(await page.inputValue('#f_company'), 'Acme');
  // new quotation (confirm auto-accepted)
  await page.click('#moreBox summary'); await page.fill('#f_signedBy', 'A N OTHER');
  await page.click('#btnNew');
  eq(await page.locator('#sumItems .item').count(), 0);
  eq(await page.inputValue('#f_company'), ''); eq(await page.inputValue('#f_contact'), '');
  eq(await page.inputValue('#f_signedBy'), 'A N OTHER');
  ok((await page.textContent('#sumItems')).includes('No items yet.'));
  await ctx.close();
});

await test('UI: transport choice and payment slabs (presets, add, remove, 100% check blocks download)', async () => {
  const { ctx, page } = await newPage();
  await typeLines(page, ['solid 600x600 b125 qty 1']); await page.keyboard.press('Enter');
  await page.fill('#f_company', 'Slab Co');
  await page.click('#moreBox summary');
  eq(await page.inputValue('#f_transportMode'), 'Ex-works');
  ok((await page.textContent('#transportHint')).endsWith('EX-WORKS PANVEL'));
  await page.selectOption('#f_transportMode', 'CIF');
  await page.fill('#f_transportPlace', 'Jebel Ali');
  ok((await page.textContent('#transportHint')).endsWith('CIF JEBEL ALI'));
  await page.selectOption('#f_transportMode', 'Extra');
  ok(await page.isHidden('#placeFld'));
  await page.selectOption('#f_transportMode', 'FOB');
  await page.fill('#f_transportPlace', 'Nhava Sheva');
  await page.click('#payPresets button:nth-child(3)'); // 30 / 70
  eq(await page.locator('#slabList .slab').count(), 2);
  await page.fill('#sp1', '60');
  ok((await page.textContent('#slabTotal')).includes('add up to 90%'));
  ok((await page.textContent('#ready')).includes('fix the payment slabs'));
  await page.click('#btnDownload');
  ok((await lastToast(page)).includes('add up to 90%'));
  await page.click('#slabAdd');
  eq(await page.inputValue('#sp2'), '10');
  eq(await page.evaluate(() => document.activeElement.id), 'ss2');
  await page.keyboard.type('WITHIN 30 DAYS OF DELIVERY');
  ok((await page.textContent('#slabTotal')).startsWith('Total 100%'));
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#btnDownload')]);
  const file = path.join(OUT, 'ui-slabs.xlsx'); await dl.saveAs(file);
  const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(file); const ws = wb.getWorksheet('Quotation');
  const terms = {}; ws.eachRow(row => { const k = row.getCell(4).value; if (typeof k === 'string') terms[k] = row.getCell(5).value; });
  eq(terms.TRANSPORTATION, ': FOB NHAVA SHEVA');
  eq(terms.PAYMENTS, ': 30 % ADVANCE WITH PO\n  60 % BEFORE DISPATCH\n  10 % WITHIN 30 DAYS OF DELIVERY');
  // remove a slab, reload: settings persist
  await page.click('#slabList .slab:nth-child(3) .rm');
  eq(await page.locator('#slabList .slab').count(), 2);
  await page.reload(); await page.waitForFunction(() => !!window.__tdApp);
  eq(await page.inputValue('#f_transportMode'), 'FOB'); eq(await page.inputValue('#f_transportPlace'), 'Nhava Sheva');
  eq(await page.locator('#slabList .slab').count(), 2);
  eq(page.errors.length, 0, page.errors.join());
  await ctx.close();
});

await test('UI: old saved terms (free-text transport/payments) are migrated', async () => {
  const { ctx, page } = await newPage();
  await page.evaluate(() => localStorage.setItem('td_qtn_fields', JSON.stringify({ company: 'Old', transport: 'Extra', payments: '100 % ADVANCE', serial: 7 })));
  await page.reload(); await page.waitForFunction(() => !!window.__tdApp);
  eq(await page.inputValue('#f_transportMode'), 'Extra');
  eq(await page.inputValue('#sp0'), '100'); eq(await page.inputValue('#ss0'), 'ADVANCE');
  eq(await page.inputValue('#f_serial'), '7');
  await ctx.close();
});

await test('UI 11a: 360 px wide, no horizontal scroll (empty, with items, picker on gully)', async () => {
  const { ctx, page } = await newPage({ viewport: { width: 360, height: 780 } });
  const sw = () => page.evaluate(() => Math.max(document.documentElement.scrollWidth, document.body.scrollWidth));
  eq(await sw(), 360, 'empty');
  await typeLines(page, [...FOUR, 'solid 777x777 b125 a very long line that goes on and on without stopping qty 1']); await page.keyboard.press('Enter');
  await page.click('[data-prod="gully"]');
  await page.click('#sumItems details summary');
  eq(await sw(), 360, 'with items');
  ok(await page.isVisible('#mbar'), 'mobile bar visible');
  eq(await page.textContent('#mTotal'), '₹3,15,083.60');
  await page.screenshot({ path: path.join(OUT, 'ui-360.png'), fullPage: true });
  await page.click('#btnPreview'); eq(await sw(), 360, 'preview open');
  await page.screenshot({ path: path.join(OUT, 'ui-360-preview.png') });
  await ctx.close();
});

await test('UI 11b: works with localStorage blocked', async () => {
  const { ctx, page } = await newPage({ blockStorage: true, keepStorage: true });
  await typeLines(page, FOUR); await page.keyboard.press('Enter');
  eq(await page.locator('#sumItems .item').count(), 4);
  await page.fill('#f_company', 'No Storage Ltd');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#btnDownload')]);
  eq(dl.suggestedFilename(), 'QTN_001_No_Storage_Ltd.xlsx');
  eq(page.errors.length, 0, page.errors.join());
  await ctx.close();
});

await test('UI 11c: stays white and blue in dark mode', async () => {
  const { ctx, page } = await newPage({ colorScheme: 'dark' });
  const c = await page.evaluate(() => ({
    body: getComputedStyle(document.body).backgroundColor, card: getComputedStyle(document.querySelector('.card')).backgroundColor,
    header: getComputedStyle(document.querySelector('.top')).backgroundColor, text: getComputedStyle(document.body).color,
    input: getComputedStyle(document.querySelector('#f_company')).backgroundColor }));
  eq(c.body, 'rgb(244, 247, 252)'); eq(c.card, 'rgb(255, 255, 255)'); eq(c.header, 'rgb(44, 68, 116)');
  eq(c.text, 'rgb(27, 42, 74)'); eq(c.input, 'rgb(255, 255, 255)');
  await ctx.close();
});

await test('UI: accessibility basics (labels, aria-pressed, focus ring)', async () => {
  const { ctx, page } = await newPage();
  const unlabeled = await page.evaluate(() => [...document.querySelectorAll('input,select,textarea')].filter(el => el.type !== 'hidden' && !el.labels?.length && !el.getAttribute('aria-label')).map(el => el.id || el.outerHTML.slice(0, 60)));
  eq(unlabeled.length, 0, unlabeled.join());
  eq(await page.locator('#prodGrid button[aria-pressed]').count(), 6);
  await page.focus('[data-prod="aw"]'); await page.keyboard.press('Enter');
  eq(await page.getAttribute('[data-prod="aw"]', 'aria-pressed'), 'true');
  eq(await page.evaluate(() => document.activeElement.dataset.prod), 'aw', 'focus kept after re-render');
  await page.keyboard.press('Tab');
  const outline = await page.evaluate(() => getComputedStyle(document.activeElement).outlineStyle);
  eq(outline, 'solid');
  await ctx.close();
});

await browser.close();
console.log(results.join('\n'));
console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
