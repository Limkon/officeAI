import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const dist = path.resolve(__dirname, '../addin-ui/dist');
const target = path.resolve(__dirname, 'ui');
if (!fs.existsSync(path.join(dist, 'index.html'))) {
  console.warn('[wps-sync] addin-ui/dist not found, skip');
  process.exit(0);
}
fs.rmSync(target, { recursive: true, force: true });
fs.mkdirSync(path.join(target, 'assets'), { recursive: true });
fs.copyFileSync(path.join(dist, 'index.html'), path.join(target, 'index.html'));
for (const f of fs.readdirSync(path.join(dist, 'assets'))) {
  fs.copyFileSync(path.join(dist, 'assets', f), path.join(target, 'assets', f));
}
let html = fs.readFileSync(path.join(target, 'index.html'), 'utf8');
html = html.replaceAll('src="/assets/', 'src="./assets/').replaceAll('href="/assets/', 'href="./assets/').replaceAll("src='/assets/", "src='./assets/");
fs.writeFileSync(path.join(target, 'index.html'), html, 'utf8');
console.log('[wps-sync] synced', fs.readdirSync(target), '->', path.relative(process.cwd(), target));
