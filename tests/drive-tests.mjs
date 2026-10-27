// Google Drive client-database tests. Run: node tests/drive-tests.mjs
// Google sign-in and the Drive REST API are replaced by in-memory stand-ins that
// accept the same requests (files.list with q, folder create, multipart upload,
// PATCH update, alt=media download), so the folder logic and register are checked
// without a real Google account.
import fs from 'node:fs';
import { createRequire } from 'node:module';
import { INDEX } from './core.mjs';

const require = createRequire(import.meta.url);
const { chromium } = require('playwright');
const EXCELJS_JS = fs.readFileSync(require.resolve('exceljs/dist/exceljs.min.js'));

let pass = 0, fail = 0; const results = [];
async function test(name, fn) { try { await fn(); pass++; results.push('PASS  ' + name); } catch (e) { fail++; results.push('FAIL  ' + name + '\n      ' + String(e && e.stack || e).split('\n').slice(0, 2).join(' ')); } }
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ', got ' + JSON.stringify(a)); };
const ok = (c, m) => { if (!c) throw new Error(m || 'assertion failed'); };

/* ---- fake Drive ---- */
function fakeDrive() {
  const files = new Map(); let n = 0; const calls = []; let failNext = 0;
  const add = f => { const id = 'id' + (++n); const rec = { id, parents: ['root'], webViewLink: 'https://drive.example/' + id, ...f }; files.set(id, rec); return rec; };
  const unq = s => s.replace(/\\'/g, "'").replace(/\\\\/g, '\\');
  function list(q) {
    const name = (q.match(/name = '((?:\\.|[^'])*)'/) || [])[1];
    const folder = /mimeType = 'application\/vnd\.google-apps\.folder'/.test(q);
    const parent = (q.match(/'((?:\\.|[^'])*)' in parents/) || [])[1];
    return [...files.values()].filter(f => (name === undefined || f.name === unq(name)) && (!folder || f.mimeType === 'application/vnd.google-apps.folder') && (!parent || f.parents.includes(unq(parent))));
  }
  function parseMultipart(buf, ctype) {
    const b = '--' + ctype.match(/boundary=(.+)$/)[1];
    const parts = buf.toString('latin1').split(b).slice(1, -1).map(p => { const i = p.indexOf('\r\n\r\n'); return { head: p.slice(0, i), body: Buffer.from(p.slice(i + 4, -2), 'latin1') }; });
    return { meta: JSON.parse(parts[0].body.toString('utf8')), data: parts[1].body };
  }
  async function handle(route) {
    const req = route.request(), url = new URL(req.url()), method = req.method();
    calls.push(method + ' ' + url.pathname);
    if (req.headers()['authorization'] !== 'Bearer tok123') return route.fulfill({ status: 401, body: '{}' });
    if (failNext > 0) { failNext--; return route.fulfill({ status: 500, contentType: 'application/json', body: JSON.stringify({ error: { message: 'Backend error' } }) }); }
    const json = o => route.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(o) });
    if (url.pathname === '/drive/v3/files' && method === 'GET') return json({ files: list(url.searchParams.get('q')).map(f => ({ id: f.id, name: f.name, webViewLink: f.webViewLink })) });
    if (url.pathname === '/drive/v3/files' && method === 'POST') { const m = JSON.parse(req.postData()); const f = add({ name: m.name, mimeType: m.mimeType, parents: m.parents || ['root'] }); return json({ id: f.id, name: f.name, webViewLink: f.webViewLink }); }
    let m;
    // Google Sheets: export as CSV, and a plain media upload replaces the sheet's content
    if ((m = url.pathname.match(/^\/drive\/v3\/files\/([^/]+)\/export$/))) { const f = files.get(m[1]); if (f.mimeType !== 'application/vnd.google-apps.spreadsheet' || url.searchParams.get('mimeType') !== 'text/csv') return route.fulfill({ status: 400, body: '{}' }); f.exports = (f.exports || 0) + 1; return route.fulfill({ status: 200, contentType: 'text/csv', body: f.data }); }
    if ((m = url.pathname.match(/^\/upload\/drive\/v3\/files\/([^/]+)$/)) && method === 'PATCH' && url.searchParams.get('uploadType') === 'media') { const f = files.get(m[1]); f.data = req.postDataBuffer(); f.updates = (f.updates || 0) + 1; return json({ id: f.id, name: f.name, webViewLink: f.webViewLink }); }
    if ((m = url.pathname.match(/^\/drive\/v3\/files\/([^/]+)$/)) && url.searchParams.get('alt') === 'media') return route.fulfill({ status: 200, body: files.get(m[1]).data });
    if (url.pathname === '/upload/drive/v3/files' && method === 'POST') { const p = parseMultipart(req.postDataBuffer(), req.headers()['content-type']); const f = add({ name: p.meta.name, mimeType: p.meta.mimeType, parents: p.meta.parents || ['root'], data: p.data }); return json({ id: f.id, name: f.name, webViewLink: f.webViewLink }); }
    if ((m = url.pathname.match(/^\/upload\/drive\/v3\/files\/([^/]+)$/)) && method === 'PATCH') { const p = parseMultipart(req.postDataBuffer(), req.headers()['content-type']); const f = files.get(m[1]); f.data = p.data; return json({ id: f.id, name: f.name, webViewLink: f.webViewLink }); }
    return route.fulfill({ status: 404, body: '{}' });
  }
  const path = id => { const f = files.get(id); if (!f) return ''; const p = f.parents[0]; return (p === 'root' ? '' : path(p) + ' / ') + f.name; };
  return { files, calls, handle, path, failOnce: (k = 1) => { failNext = k; } };
}

const browser = await chromium.launch();
async function newPage(drive) {
  const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1280, height: 900 } });
  await ctx.route('https://cdn.jsdelivr.net/**', r => r.fulfill({ body: EXCELJS_JS, contentType: 'application/javascript' }));
  await ctx.route(/fonts\.(googleapis|gstatic)\.com/, r => r.abort());
  await ctx.route('https://www.googleapis.com/**', r => drive.handle(r));
  await ctx.addInitScript(() => {
    window.__tokReq = 0;
    window.google = { accounts: { oauth2: { initTokenClient(cfg) { return { requestAccessToken() { window.__tokReq++; setTimeout(() => cfg.callback({ access_token: 'tok123', expires_in: 3600 }), 5); } }; } } } };
  });
  const page = await ctx.newPage(); page.errors = []; page.on('pageerror', e => page.errors.push(e.message)); page.on('dialog', d => d.accept());
  await page.goto('file://' + INDEX); await page.waitForFunction(() => !!window.__tdApp && !!window.ExcelJS);
  await page.evaluate(() => localStorage.clear()); await page.reload(); await page.waitForFunction(() => !!window.__tdApp && !!window.ExcelJS);
  return { ctx, page };
}
const lastToast = p => p.locator('#toasts .toast').last().textContent();
async function quote(page, company, city, line, waitDrive = true) {
  await page.fill('#quick', line); await page.press('#quick', 'Enter');
  await page.fill('#f_company', company); await page.keyboard.press('Escape'); await page.fill('#f_city', city);
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#btnDownload')]);
  if (waitDrive) await page.waitForFunction(() => { const t = document.querySelector('#driveSaved').textContent; return t && !/Saving to Google Drive…/.test(t); });
  return dl.suggestedFilename();
}

await test('Drive: folders by city and client, every quotation saved, register row per quotation', async () => {
  const drive = fakeDrive(); const { ctx, page } = await newPage(drive);
  await page.click('#driveBox summary');
  eq((await page.textContent('#driveBadge')).trim(), 'Off');
  await page.fill('#f_driveClientId', '123-abc.apps.googleusercontent.com');
  await page.check('#driveOn');
  await page.click('#driveConnect');
  await page.waitForFunction(() => /Connected/.test(document.querySelector('#driveStatus').textContent));
  eq((await page.textContent('#driveBadge')).trim(), 'On, connected');

  const f1 = await quote(page, 'Brigade Infra', 'Bengaluru', 'solid top 600x600 b125 qty 4');
  ok((await page.textContent('#driveSaved')).includes('Saved to Google Drive: THERMODRAIN Quotations / BENGALURU / BRIGADE INFRA'), await page.textContent('#driveSaved'));
  const f2 = await quote(page, 'Brigade Infra', 'Bengaluru', 'recess cover 900x600 5 ton qty 2');
  const f3 = await quote(page, 'Indian Oil Corporation', 'Mumbai', 'fuel station 900 dia d400 qty 3');
  const f4 = await quote(page, 'Walk-in Customer', '', 'water tank cover 600 qty 1');

  const all = [...drive.files.values()];
  const folders = all.filter(f => f.mimeType === 'application/vnd.google-apps.folder').map(f => drive.path(f.id)).sort();
  eq(folders.join(' | '), ['THERMODRAIN Quotations', 'THERMODRAIN Quotations / BENGALURU', 'THERMODRAIN Quotations / BENGALURU / BRIGADE INFRA',
    'THERMODRAIN Quotations / CITY NOT GIVEN', 'THERMODRAIN Quotations / CITY NOT GIVEN / WALK-IN CUSTOMER',
    'THERMODRAIN Quotations / MUMBAI', 'THERMODRAIN Quotations / MUMBAI / INDIAN OIL CORPORATION'].join(' | '), 'one folder per city and client, no duplicates');
  const xl = all.filter(f => f.name.endsWith('.xlsx')).map(f => drive.path(f.id)).sort();
  eq(xl.join(' | '), [
    'THERMODRAIN Quotations / BENGALURU / BRIGADE INFRA / ' + f1, 'THERMODRAIN Quotations / BENGALURU / BRIGADE INFRA / ' + f2,
    'THERMODRAIN Quotations / CITY NOT GIVEN / WALK-IN CUSTOMER / ' + f4, 'THERMODRAIN Quotations / MUMBAI / INDIAN OIL CORPORATION / ' + f3].sort().join(' | '));
  ok(all.filter(f => f.name.endsWith('.xlsx')).every(f => f.data.length > 4000 && f.data.slice(0, 2).toString() === 'PK'), 'real xlsx files uploaded');
  const reg = all.find(f => f.name === 'Quotation register.csv');
  eq(drive.path(reg.id), 'THERMODRAIN Quotations / Quotation register.csv');
  const lines = reg.data.toString('utf8').replace(/^﻿/, '').trim().split(/\r\n/);
  eq(lines.length, 5, 'header + 4 quotations');
  ok(lines[0].startsWith('Date,Quotation ref,Client,City,'), lines[0]);
  ok(lines[1].includes('BRIGADE INFRA,BENGALURU') && lines[1].includes('/001/') && lines[1].includes('INR'), lines[1]);
  ok(lines[3].includes('INDIAN OIL CORPORATION,MUMBAI') && lines[3].includes('/003/') && lines[3].includes('3 x 900mm Dia'), lines[3]);
  ok(lines[4].includes('CITY NOT GIVEN'), lines[4]);
  ok(lines[1].includes('https://drive.example/'), 'file link in the register');
  // each client folder has a Google Sheet with one row per product; a new quotation adds rows
  const sheets = all.filter(f => f.mimeType === 'application/vnd.google-apps.spreadsheet');
  eq(sheets.map(f => drive.path(f.id)).sort().join(' | '), ['THERMODRAIN Quotations / BENGALURU / BRIGADE INFRA / BRIGADE INFRA - client data',
    'THERMODRAIN Quotations / CITY NOT GIVEN / WALK-IN CUSTOMER / WALK-IN CUSTOMER - client data',
    'THERMODRAIN Quotations / MUMBAI / INDIAN OIL CORPORATION / INDIAN OIL CORPORATION - client data'].join(' | '));
  const brig = sheets.find(f => f.name.startsWith('BRIGADE'));
  const rows = brig.data.toString('utf8').trim().split(/\r\n/);
  eq(rows.length, 4, 'header + one row per product line (the 2nd quotation has 2 products)');
  eq(rows[0], 'Date,Quotation ref,Client name,City,Product,Size,Load class,Qty,Discount %,Net rate /pc,Value (excl. GST),Quote value (incl. GST),Currency');
  ok(rows[1].includes('/001/') && rows[1].includes('BRIGADE INFRA,BENGALURU') && rows[1].includes('600mm X 600mm,B 125,4,0,12345,49380,58268.4,INR'), rows[1]);
  ok(rows[2].includes('/002/') && rows[2].includes('600mm X 600mm'), rows[2]);
  ok(rows[3].includes('/002/') && rows[3].includes('5 Ton'), rows[3]);
  eq(brig.updates, 1, 'second quotation updated the same sheet'); eq(brig.exports, 1);
  ok((await page.locator('#driveSaved a', { hasText: 'client data sheet' }).count()) === 1, 'link to the sheet');
  eq(await page.evaluate(() => window.__tokReq), 1, 'signed in once, reused for every download');
  eq(page.errors.length, 0, page.errors.join());
  await ctx.close();
});

await test('Drive: downloading the same quotation again replaces its rows in the client sheet', async () => {
  const drive = fakeDrive(); const { ctx, page } = await newPage(drive);
  await page.click('#driveBox summary'); await page.check('#driveOn'); await page.click('#driveConnect');
  await page.waitForFunction(() => /Connected/.test(document.querySelector('#driveStatus').textContent));
  await quote(page, 'Repeat Co', 'Pune', 'solid top 600x600 b125 qty 4');
  await page.evaluate(() => { const el = document.querySelector('#f_serial'); el.value = '1'; el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); }); // same quotation number again, e.g. after a correction
  await page.fill('#quick', 'recess cover 900x600 5 ton qty 2'); await page.press('#quick', 'Enter');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#btnDownload')]);
  await page.waitForFunction(() => /Saved to Google Drive/.test(document.querySelector('#driveSaved').textContent));
  const sh = [...drive.files.values()].find(f => f.name === 'REPEAT CO - client data');
  const rows = sh.data.toString('utf8').trim().split(/\r\n/);
  eq(rows.length, 3, 'header + the 2 products of the corrected quotation, old rows replaced');
  ok(rows.slice(1).every(r => r.includes('/001/')));
  eq(page.errors.length, 0, page.errors.join());
  await ctx.close();
});

await test('Drive: a failure never blocks the download; Try again uploads it', async () => {
  const drive = fakeDrive(); const { ctx, page } = await newPage(drive);
  await page.click('#driveBox summary');
  await page.fill('#f_driveClientId', '123-abc.apps.googleusercontent.com'); await page.check('#driveOn');
  drive.failOnce(1);
  const fn = await quote(page, 'Prestige Estates', 'Mysuru', 'solid top 600x600 b125 qty 1');
  eq(fn, 'QTN_001_Prestige_Estates.xlsx', 'Excel still downloaded');
  ok((await page.textContent('#driveSaved')).includes('Not saved to Google Drive'), await page.textContent('#driveSaved'));
  ok((await lastToast(page)).includes('downloaded, but not saved to Google Drive'));
  await page.click('#driveRetry');
  await page.waitForFunction(() => /Saved to Google Drive/.test(document.querySelector('#driveSaved').textContent));
  const xl = [...drive.files.values()].filter(f => f.name === fn);
  eq(xl.length, 1); eq(drive.path(xl[0].id), 'THERMODRAIN Quotations / MYSURU / PRESTIGE ESTATES / ' + fn);
  eq(page.errors.length, 0, page.errors.join());
  await ctx.close();
});

await test('Drive: off by default (company Client ID built in); on without a Client ID explains what to do', async () => {
  const drive = fakeDrive(); const { ctx, page } = await newPage(drive);
  eq(await page.inputValue('#f_driveClientId'), '294947591975-gu0rcpscqt0phbk54mbsbsm4slejpc82.apps.googleusercontent.com', 'company Client ID filled in');
  ok(!(await page.isChecked('#driveOn')), 'still needs the tick in each browser');
  await quote(page, 'No Drive Co', 'Pune', 'solid top 600x600 b125 qty 1', false);
  await page.waitForTimeout(300);
  eq(drive.calls.length, 0, 'nothing sent to Google when Drive is off');
  ok(await page.isHidden('#driveSaved'));
  await page.click('#driveBox summary'); await page.fill('#f_driveClientId', ''); await page.check('#driveOn');
  await quote(page, 'No Drive Co', 'Pune', 'solid top 600x600 b125 qty 1');
  ok((await page.textContent('#driveSaved')).includes('Add the Google OAuth Client ID'), await page.textContent('#driveSaved'));
  await ctx.close();
});

await browser.close();
console.log(results.join('\n'));
console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
