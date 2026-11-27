import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const dist = path.resolve(__dirname, '../addin-ui/dist');
const target = path.resolve(__dirname, 'ui');

function injectFallback(html, isWps) {
  if (html.includes('oa-fallback')) return html;
  const fallback = `  <style>#oa-fallback{font-family:system-ui;padding:20px;font-size:13px;color:#334155}#oa-fallback h3{margin:0 0 8px;font-size:15px}</style>
</head>
<body class="bg-slate-50 text-slate-900 antialiased">
  <div id="root"><div id="oa-fallback"><h3>OfficeAI 正在加载…</h3><div id="oa-fallback-msg">若 3 秒后仍空白，${isWps ? '请点 WPS 功能区“调试日志”并检查网关 http://127.0.0.1:18080 是否存活' : '请检查网关 http://127.0.0.1:18080 是否存活'}</div><div id="oa-fallback-detail" style="margin-top:10px;font-size:11px;color:#64748b;white-space:pre-wrap;word-break:break-all"></div></div></div>
  <script>
    setTimeout(function(){var d=document.getElementById('oa-fallback-detail'); if(d){try{d.textContent='location='+location.href+'\\nprotocol='+location.protocol}catch(e){d.textContent=e.message}}},3500);
  </script>
</body>`;
  return html.replace('</head>\n<body class="bg-slate-50 text-slate-900 antialiased">\n  <div id="root"></div>\n</body>', fallback);
}

if (!fs.existsSync(path.join(dist, 'index.html'))) {
  console.warn('[wps-sync] addin-ui/dist not found, skip');
  process.exit(0);
}
// 1) 给 dist 本身注入白屏兜底（供 http 网关的 /dashboard 使用）
let distHtml = fs.readFileSync(path.join(dist, 'index.html'), 'utf8');
let distPatched = injectFallback(distHtml, false);
if (distPatched !== distHtml) {
  fs.writeFileSync(path.join(dist, 'index.html'), distPatched, 'utf8');
  console.log('[wps-sync] patched dist/index.html fallback');
}
// 2) 同步到 wps 离线包（file://）
fs.rmSync(target, { recursive: true, force: true });
fs.mkdirSync(path.join(target, 'assets'), { recursive: true });
fs.copyFileSync(path.join(dist, 'index.html'), path.join(target, 'index.html'));
for (const f of fs.readdirSync(path.join(dist, 'assets'))) {
  fs.copyFileSync(path.join(dist, 'assets', f), path.join(target, 'assets', f));
}
let html = fs.readFileSync(path.join(target, 'index.html'), 'utf8');
// file:// 需要相对路径
html = html.replaceAll('src="/assets/', 'src="./assets/').replaceAll('href="/assets/', 'href="./assets/');
html = injectFallback(html, true);
fs.writeFileSync(path.join(target, 'index.html'), html, 'utf8');
console.log('[wps-sync] synced', fs.readdirSync(target), '->', path.relative(process.cwd(), target));
