// Loads the pure core script (<script id="td-core">) out of index.html into a Node VM.
import fs from 'node:fs';
import vm from 'node:vm';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const INDEX = path.join(ROOT, 'index.html');

export function loadCore() {
  const html = fs.readFileSync(INDEX, 'utf8');
  const m = html.match(/<script id="td-core">([\s\S]*?)<\/script>/);
  if (!m) throw new Error('core script not found in index.html');
  const ctx = vm.createContext({ console, Date });
  vm.runInContext(m[1] + '\n;globalThis.TD = TD;', ctx, { filename: 'td-core.js' });
  return ctx.TD;
}
