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
  if (process.env.UI_TRACE) console.error('… ' + name);
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
  eq(await page.locator('#lineCheck li.ready').count(), 4, 'all four lines show green before adding');
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
  const errs = await page.locator('#lineCheck li.blocked .lc-text').allTextContents();
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
  ok((await page.textContent('[data-size="s900t"]')).includes('2 covers, 1 MS frame'));
  eq((await page.textContent('#abPrice')).startsWith('₹20,725'), true);
  await page.click('#abPlus'); await page.click('#abPlus');
  eq(await page.textContent('#abAdd'), 'Add 3 to quote, ₹62,175');
  await page.click('#abAdd');
  eq(await page.inputValue('#abQty'), '1', 'quantity resets to 1 after adding');
  eq(await page.textContent('#abAdd'), 'Add 1 to quote, ₹20,725');
  await page.click('#abAdd');
  eq(await page.locator('#sumItems .item').count(), 1);
  eq(await page.inputValue('#sumItems input[data-act=qty]'), '4', 'same item merges: 3 + 1');
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
  await page.fill('#discount', '48'); await page.click('#discApply');
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

await test('UI: live checklist (grey while missing, green when every detail is filled)', async () => {
  const { ctx, page } = await newPage();
  await page.focus('#quick'); await page.keyboard.type('solid');
  const chips = async () => (await page.locator('#lineCheck .chip').evaluateAll(els => els.map(e => (e.classList.contains('done') ? '+' : '-') + e.textContent.replace(/, (filled|missing)$/, ''))));
  eq((await chips()).join(' '), '+✓Product -Size -Load class -Qty');
  eq(await page.locator('#lineCheck li.pending').count(), 1, 'grey, not red, while details are missing');
  eq(await page.locator('#lineCheck li.blocked, #lineCheck li.ready').count(), 0);
  ok((await page.textContent('#lineCheck .lc-text')).includes('Still needed: size, load class, qty.'));
  await page.keyboard.type(' top 600x600 b125');
  eq((await chips()).join(' '), '+✓Product +✓Size +✓Load class -Qty');
  eq(await page.locator('#lineCheck li.ready').count(), 0, 'not green until qty is given');
  await page.keyboard.type(' qty 4');
  eq(await page.locator('#lineCheck li.ready').count(), 1, 'green once every detail is filled');
  ok((await page.textContent('#lineCheck .lc-text')).includes('4 nos × ₹12,345 = ₹49,380'));
  // chamber needs a price; second line shows its own checklist
  await page.keyboard.press('Shift+Enter'); await page.keyboard.type('chamber 1000x1000x600 two sided B125 qty 2');
  const rows = page.locator('#lineCheck li');
  eq(await rows.count(), 2);
  ok((await rows.nth(1).textContent()).includes('Still needed: price.'));
  await page.keyboard.type(' price 45000');
  eq(await page.locator('#lineCheck li.ready').count(), 2);
  eq((await page.textContent('#quickAdd')).trim(), 'Add 2');
  await page.keyboard.press('Enter');
  eq(await page.locator('#sumItems .item').count(), 2);
  eq(await page.locator('#lineCheck li').count(), 0, 'checklist clears once lines are added');
  // the Add button sits inside the text box
  const tb = await page.locator('#quick').boundingBox(), bb = await page.locator('#quickAdd').boundingBox();
  ok(bb.x >= tb.x && bb.x + bb.width <= tb.x + tb.width && bb.y >= tb.y && bb.y + bb.height <= tb.y + tb.height, 'Add inside the box');
  await ctx.close();
});

await test('UI: custom chamber, MS frame with typed price, recess depth', async () => {
  const { ctx, page } = await newPage();
  await page.click('[data-prod="custom"]');
  ok(await page.isVisible('#customWrap')); ok(await page.isHidden('#stdWrap'));
  await page.fill('#cName', 'FRP Inspection Chamber');
  await page.fill('#cL', '1200'); await page.fill('#cW', '900'); await page.fill('#cH', '750');
  await page.click('[data-finish="double"]'); await page.click('[data-load="4"]');
  eq(await page.getAttribute('[data-frame="ms"]', 'aria-pressed'), 'true', 'chambers default to MS frame');
  ok(await page.isDisabled('#abAdd'), 'no price yet');
  await page.fill('#abPriceIn', '52000');
  eq(await page.textContent('#abAdd'), 'Add 1 to quote, ₹52,000');
  await page.click('#abAdd');
  const sum = await page.textContent('#sumItems');
  ok(sum.includes('FRP Inspection Chamber, with MS frame') && sum.includes('1200 × 900 × 750 mm, two-sided finish, D400'), sum);
  // single solid cover: no MS frame option; without frame needs a price
  await page.click('[data-prod="solid"]'); await page.click('[data-shape="Square"]'); await page.click('[data-size="s600"]'); await page.click('[data-load="2"]');
  eq((await page.textContent('#abPrice')).startsWith('₹12,345'), true);
  eq(await page.locator('#frameSeg [data-frame="ms"]').count(), 0, 'no MS frame on a single cover');
  await page.click('[data-frame="none"]');
  ok(await page.isVisible('#abManual')); ok(await page.isDisabled('#abAdd'));
  await page.fill('#abPriceIn', '15000'); await page.click('#abAdd');
  // recess depth
  await page.click('[data-prod="recess"]'); ok(await page.isVisible('#depthWrap'));
  await page.click('#depthWrap [data-depth="75"]');
  ok((await page.textContent('#abSpec')).includes('recess depth 75 mm'));
  await page.click('#abAdd');
  const M = await page.evaluate(() => window.__tdApp.model());
  eq(M.groups.map(g => g.title).join(' | '), '‘THERMODRAIN’ FRP INSPECTION CHAMBER WITH MS FRAME | ‘THERMODRAIN’ FRP/GRP SOLID TOP MANHOLE COVER WITHOUT FRAME | ‘THERMODRAIN’ FRP/GRP RECESS COVER WITH FRAME');
  ok(M.groups[2].items[0].wording.endsWith('recess depth 75 mm'));
  // manual price can be changed in the summary
  await page.fill('#sumItems .item:nth-child(2) input[data-act=price]', '16000'); await page.press('#sumItems .item:nth-child(2) input[data-act=price]', 'Tab');
  ok((await page.textContent('#sumItems .item:nth-child(2) .amt')).includes('₹16,000'));
  eq(page.errors.length, 0, page.errors.join());
  await ctx.close();
});

await test('UI: per-item discounts, average discount, apply to all; US$ export mode has no GST', async () => {
  const { ctx, page } = await newPage();
  await typeLines(page, ['solid 600x600 b125 qty 4', 'recess 900x600 5 ton qty 4']); await page.keyboard.press('Enter');
  await page.fill('#sumItems .item:nth-child(1) input[data-act=disc]', '10');
  await page.fill('#sumItems .item:nth-child(2) input[data-act=disc]', '20');
  const tot = await page.textContent('#totals');
  ok(tot.includes('Less discount, average 15.90%') && tot.includes('−₹19,130.00'), tot);
  await page.click('#f_company');
  eq(await page.inputValue('#discount'), '15.9', 'discount box shows the average of all items');
  await page.fill('#discount', '12'); await page.click('#discApply');
  eq(await page.inputValue('#sumItems .item:nth-child(2) input[data-act=disc]'), '12');
  ok((await page.textContent('#totals')).includes('Less discount @12%'));
  // GST mode: no USD anywhere, rate box hidden
  ok(await page.isHidden('#rateFld'));
  for (const sel of ['#sumItems', '#totals', '#grandVal']) ok(!(await page.textContent(sel)).includes('US$'), 'no US$ amounts in GST mode: ' + sel);
  // US$ export: no GST, rate required
  eq(await page.getAttribute('#modeToggle', 'aria-checked'), 'false');
  await page.click('#modeToggle');
  eq(await page.getAttribute('#modeToggle', 'aria-checked'), 'true');
  ok(await page.isVisible('#rateFld'));
  eq(await page.textContent('#grandVal'), 'US$ —');
  await page.fill('#f_company', 'USD Co');
  ok((await page.textContent('#ready')).includes('exchange rate'));
  await page.click('#btnDownload');
  eq(await page.evaluate(() => document.activeElement.id), 'f_usdRate');
  await page.fill('#f_usdRate', '88.5');
  const tot2 = await page.textContent('#totals');
  ok(!tot2.includes('GST') && tot2.includes('US$'), tot2);
  eq(await page.textContent('#grandLabel'), 'Total in US$, export, no GST');
  ok((await page.textContent('#sumItems')).includes('= US$ 139.49'));
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#btnDownload')]);
  const file = path.join(OUT, 'ui-usd.xlsx'); await dl.saveAs(file);
  const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(file); const ws = wb.getWorksheet('Quotation');
  const all = []; ws.eachRow(row => row.eachCell(c => all.push(String(c.value && c.value.formula ? '' : c.value))));
  ok(all.includes('PRICE /PC (US$)') && all.includes('TOTAL (US$)'), 'US$ quotation');
  ok(!all.some(v => /ADD : GST|INCL\. 18% GST/.test(v)), 'no GST rows');
  // back to GST: USD disappears again
  await page.click('#modeToggle');
  ok((await page.textContent('#totals')).includes('Add GST @18%')); ok(!(await page.textContent('#totals')).includes('US$'));
  ok(await page.isHidden('#rateFld'));
  await ctx.close();
});

await test('UI: saved customers (save, pick, delete, auto-save on download, persists)', async () => {
  const { ctx, page } = await newPage();
  await page.evaluate(() => localStorage.removeItem('td_customers'));
  await page.reload(); await page.waitForFunction(() => !!window.__tdApp);
  await page.fill('#f_company', 'Brigade Infra'); await page.fill('#f_address', 'No. 29, Brigade Road');
  await page.fill('#f_city', 'Bengaluru'); await page.fill('#f_cityCode', 'BGLR'); await page.fill('#f_contact', 'Mr Rao');
  await page.click('#custSave');
  ok((await lastToast(page)).startsWith('Customer saved'));
  await page.click('#btnNew');
  eq(await page.inputValue('#f_company'), ''); eq(await page.inputValue('#f_address'), '');
  await page.reload(); await page.waitForFunction(() => !!window.__tdApp);
  await page.selectOption('#custPick', { label: 'Brigade Infra, Bengaluru' });
  eq(await page.inputValue('#f_address'), 'No. 29, Brigade Road'); eq(await page.inputValue('#f_cityCode'), 'BGLR'); eq(await page.inputValue('#f_contact'), 'Mr Rao');
  ok((await page.textContent('#hdrRef')).includes('TP1-BGLR'));
  // auto-save on download
  await typeLines(page, ['tank 600 qty 1']); await page.keyboard.press('Enter');
  await page.fill('#f_company', 'Prestige Estates'); await page.fill('#f_city', 'Mysuru');
  await Promise.all([page.waitForEvent('download'), page.click('#btnDownload')]);
  const opts = await page.locator('#custPick option').allTextContents();
  ok(opts.includes('Prestige Estates, Mysuru') && opts.includes('Brigade Infra, Bengaluru'), opts.join());
  await page.selectOption('#custPick', { label: 'Brigade Infra, Bengaluru' });
  await page.click('#custDel');
  ok(!(await page.locator('#custPick option').allTextContents()).includes('Brigade Infra, Bengaluru'));
  await ctx.close();
});

await test('UI: no stale typed price, serial stays consistent, twin frame wording', async () => {
  const { ctx, page } = await newPage();
  await page.click('[data-prod="solid"]'); await page.click('[data-shape="Square"]'); await page.click('[data-size="s600"]'); await page.click('[data-load="2"]');
  await page.click('[data-frame="none"]'); await page.fill('#abPriceIn', '15000');
  eq(await page.textContent('#abAdd'), 'Add 1 to quote, ₹15,000');
  await page.click('[data-size="s900"]');
  eq(await page.inputValue('#abPriceIn'), '', 'price cleared after size change');
  ok(await page.isDisabled('#abAdd'), 'cannot add without a fresh price');
  await page.fill('#abPriceIn', '21000'); await page.click('[data-load="3"]');
  eq(await page.inputValue('#abPriceIn'), '', 'price cleared after load change');
  await page.fill('#abPriceIn', '22000'); await page.press('#abPriceIn', 'Enter');
  eq(await page.locator('#sumItems .item').count(), 1, 'Enter in the price box adds the item');
  // twin row with FRP frame: wording follows the frame
  await page.click('[data-size="s900t"]');
  eq(await page.getAttribute('[data-frame="ms"]', 'aria-pressed'), 'true', 'twin defaults to MS frame (list price)');
  ok((await page.textContent('#abPrice')).startsWith('₹20,725'));
  await page.click('[data-frame="frp"]'); await page.fill('#abPriceIn', '24000'); await page.click('#abAdd');
  const M = await page.evaluate(() => window.__tdApp.model());
  ok(M.groups.flatMap(g => g.items).some(x => x.wording.endsWith('2 covers with single FRP frame')), 'FRP twin wording');
  // blank serial: reference and file name agree on 001
  await page.fill('#f_company', 'Serial Co');
  await page.click('#moreBox summary'); await page.fill('#f_serial', '');
  ok((await page.textContent('#hdrRef')).includes('/001/'));
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#btnDownload')]);
  eq(dl.suggestedFilename(), 'QTN_001_Serial_Co.xlsx');
  eq(await page.inputValue('#f_serial'), '2');
  eq(page.errors.length, 0, page.errors.join());
  await ctx.close();
});

await test('UI: custom multi-cover manhole cover from the picker', async () => {
  const { ctx, page } = await newPage();
  await page.click('[data-prod="custom"]');
  await page.click('[data-kind="multi"]');
  ok(await page.isVisible('#multiWrap')); ok(await page.isHidden('#chamberWrap'));
  eq(await page.getAttribute('[data-frame="ms"]', 'aria-pressed'), 'true', 'MS frame standard');
  eq(await page.locator('#loadGrid [data-load="-1"]').count(), 0, 'load class required for a manhole cover');
  await page.fill('#mL', '2400'); await page.fill('#mW', '900'); await page.fill('#mN', '25');
  eq(await page.textContent('#abAdd'), 'Covers: 2 to 20');
  await page.fill('#mN', '4'); await page.click('[data-load="3"]');
  ok(await page.isDisabled('#abAdd'), 'needs a price');
  await page.fill('#abPriceIn', '70000');
  eq(await page.textContent('#abAdd'), 'Add 1 to quote, ₹70,000');
  ok((await page.textContent('#abSpec')).includes('2400 × 900 mm, 4 covers, C250'));
  await page.click('#abAdd');
  const M = await page.evaluate(() => window.__tdApp.model());
  eq(M.groups[0].title, '‘THERMODRAIN’ FRP/GRP MULTIPLE MANHOLE COVER WITH MS FRAME');
  eq(M.groups[0].items[0].wording, '2400 mm x 900 mm clear opening, 4 covers with single MS frame');
  // switching back to chamber keeps the chamber fields
  await page.click('[data-kind="chamber"]'); ok(await page.isVisible('#chamberWrap'));
  await page.reload(); await page.waitForFunction(() => !!window.__tdApp);
  eq(await page.locator('#sumItems .item').count(), 1, 'multi-cover item survives a reload');
  eq(page.errors.length, 0, page.errors.join());
  await ctx.close();
});

await test('UI: fuel station covers in the picker (C250 / D400 only, CIC sizes)', async () => {
  const { ctx, page } = await newPage();
  await page.click('[data-prod="fuel"]');
  eq(await page.getAttribute('[data-load="3"]', 'aria-pressed'), 'true', 'starts on C250');
  for (const l of [0, 1, 2, 5]) ok(await page.isDisabled(`[data-load="${l}"]`), 'load ' + l + ' not offered');
  await page.click('[data-shape="Square"]'); await page.click('[data-size="s900cic"]');
  ok((await page.textContent('[data-size="s900cic"]')).includes('200 inner dia (CIC)'));
  await page.click('[data-load="4"]');
  ok((await page.textContent('#abPrice')).startsWith('₹76,505'));
  ok((await page.textContent('#abSpec')).includes('model 900*900 (A&W.C) CIC, 90 kg'));
  await page.click('#abAdd');
  const M = await page.evaluate(() => window.__tdApp.model());
  eq(M.groups[0].items[0].price, 76505);
  eq(M.groups[0].items[0].wording, '900 mm x 900 mm (36" x 36") clear opening with 200 mm inner dia (cover in cover), Model 900*900 (A&W.C) CIC, cover weight 90 kg');
  eq(page.errors.length, 0, page.errors.join());
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
  eq(await page.locator('#prodGrid button[aria-pressed]').count(), 8);
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
