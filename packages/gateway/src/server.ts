import crypto from 'node:crypto';
import fs from 'node:fs';
import http, { type IncomingMessage, type ServerResponse } from 'node:http';
import https from 'node:https';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  buildGatewayLaunchUrl,
  type AppSettings,
  type ChatRequest,
  type PromptAction,
  type ProviderConfig,
  type StreamEvent
} from '@officeai/shared';
import { ensureLocalhostCert } from './certs.js';
import { ConfigStore } from './config-store.js';
import {
  createLauncherShortcut,
  registerProtocolHandler,
  removeLauncherShortcut,
  resolveLaunchTarget,
  unregisterProtocolHandler,
  type LaunchTarget
} from './desktop-tools.js';
import { LifecycleMonitor } from './lifecycle.js';
import { probeLocalRuntimes, resolveProviderAdapter } from './providers/index.js';
import { WindowsOfficeIntegrator } from './windows/office-integrator.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export interface GatewayPaths {
  uiDistDir: string;
  officeAddinDir: string;
  wpsAddinDir: string;
  assetsRoot: string;
}

export function resolveGatewayPaths(): GatewayPaths {
  if (process.env.OFFICEAI_ASSETS_DIR) {
    const assetsRoot = path.resolve(process.env.OFFICEAI_ASSETS_DIR);
    return {
      assetsRoot,
      uiDistDir: path.join(assetsRoot, 'addin-ui'),
      officeAddinDir: path.join(assetsRoot, 'office-addin'),
      wpsAddinDir: path.join(assetsRoot, 'wps-addin')
    };
  }

  // In monorepo: packages/gateway/dist -> packages/
  const monorepoPackages = path.resolve(__dirname, '..', '..');
  return {
    assetsRoot: monorepoPackages,
    uiDistDir: path.join(monorepoPackages, 'addin-ui', 'dist'),
    officeAddinDir: path.join(monorepoPackages, 'office-addin'),
    wpsAddinDir: path.join(monorepoPackages, 'wps-addin')
  };
}

const MIME_TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.mjs': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.xml': 'application/xml; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.woff2': 'font/woff2',
  '.woff': 'font/woff'
};

export interface RunningGateway {
  httpServer: http.Server;
  httpsServer?: https.Server;
  httpUrl: string;
  httpsUrl?: string;
  store: ConfigStore;
  integrator: WindowsOfficeIntegrator;
  lifecycle: LifecycleMonitor;
  launchTarget: LaunchTarget;
  close: () => Promise<void>;
}

interface GatewayOptions {
  dataDir?: string;
  httpPort?: number;
  httpsPort?: number;
  host?: string;
  enableHttps?: boolean;
  /** Who started this process (add-in / desktop / cli) */
  launcher?: string;
  /** Internal hook used by the CLI, Electron app and tests */
  onExitRequested?: (reason: string) => void;
}

function sendJson(res: ServerResponse, status: number, body: unknown): void {
  const json = JSON.stringify(body);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(json)
  });
  res.end(json);
}

async function readJsonBody<T = Record<string, unknown>>(req: IncomingMessage): Promise<T> {
  const chunks: Buffer[] = [];
  let size = 0;
  const maxBytes = 4 * 1024 * 1024;
  for await (const chunk of req) {
    const buf = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    size += buf.length;
    if (size > maxBytes) {
      throw new Error('Request payload too large');
    }
    chunks.push(buf);
  }
  const raw = Buffer.concat(chunks).toString('utf8').trim();
  if (!raw) return {} as T;
  return JSON.parse(raw) as T;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Minimal offline page served to WPS when the gateway is not running. It is
 * written next to the local add-in, so it must not reference any external file.
 */
function renderLauncherPage(options: {
  gatewayUrl: string;
  launchCommand: string;
  title?: string;
  hint?: string;
}): string {
  const { gatewayUrl, launchCommand } = options;
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>${escapeHtml(options.title ?? 'OfficeAI 服务未启动')}</title>
<style>
  body { margin:0; font-family:"Microsoft YaHei","PingFang SC",system-ui,sans-serif; background:#f8fafc; color:#0f172a; }
  .wrap { max-width:520px; margin:0 auto; padding:32px 24px; }
  h1 { font-size:17px; margin:0 0 8px; }
  p { font-size:13px; line-height:1.7; color:#475569; margin:0 0 14px; }
  .card { background:#fff; border:1px solid #e2e8f0; border-radius:14px; padding:18px 20px; }
  button { width:100%; border:0; border-radius:10px; padding:11px 14px; font-size:14px; font-weight:600;
    color:#fff; background:#2563eb; cursor:pointer; }
  button.secondary { background:#f1f5f9; color:#334155; margin-top:10px; }
  code { display:block; background:#0f172a; color:#e2e8f0; border-radius:8px; padding:9px 11px;
    font-size:11px; word-break:break-all; margin:10px 0 0; }
  .state { font-size:12px; color:#64748b; margin-top:14px; }
</style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <h1>OfficeAI 服务未启动</h1>
    <p>OfficeAI 助手按需启动：点击下面的按钮会自动拉起本地网关，随后即可正常使用。</p>
    <button id="start">一键启动 OfficeAI</button>
    <button class="secondary" id="retry">我已启动，重新连接</button>
    <p style="margin-top:14px">如果按钮没有反应，请手动双击开始菜单中的「OfficeAI 启动」，或在命令行执行：</p>
    <code>${escapeHtml(launchCommand || 'Start-OfficeAI.bat')}</code>
    <div class="state" id="state"></div>
  </div>
</div>
<script>
  var GATEWAY_URL = ${JSON.stringify(gatewayUrl)};
  var LAUNCH_URL = ${JSON.stringify(buildGatewayLaunchUrl(gatewayUrl, 'wps'))};
  var stateEl = document.getElementById('state');
  function setState(text) { stateEl.textContent = text; }
  function go() { window.location.href = GATEWAY_URL + '/?host=wps'; }
  function poll(remaining) {
    fetch(GATEWAY_URL + '/api/health', { cache: 'no-store' })
      .then(function (r) { if (r.ok) { go(); } else { throw new Error('bad'); } })
      .catch(function () {
        if (remaining <= 0) {
          setState('仍未检测到服务，请手动启动后再点「重新连接」。');
          return;
        }
        setState('正在等待 OfficeAI 启动…（' + remaining + '）');
        setTimeout(function () { poll(remaining - 1); }, 1000);
      });
  }
  document.getElementById('start').onclick = function () {
    setState('正在唤起 OfficeAI…');
    window.location.href = LAUNCH_URL;
    setTimeout(function () { poll(30); }, 1200);
  };
  document.getElementById('retry').onclick = function () {
    setState('正在检查服务…');
    poll(3);
  };
  poll(2);
</script>
</body>
</html>`;
}

function isWithinDir(targetPath: string, baseDir: string): boolean {
  const resolvedTarget = path.resolve(targetPath);
  const resolvedBase = path.resolve(baseDir);
  return resolvedTarget === resolvedBase || resolvedTarget.startsWith(resolvedBase + path.sep);
}

export function isAllowedOrigin(origin: string | undefined): boolean {
  if (!origin) return true; // same-origin / curl / native desktop
  if (origin === 'null') return true; // file:// or sandboxed Office/WPS webview
  try {
    const u = new URL(origin);
    if (u.hostname === '127.0.0.1' || u.hostname === 'localhost' || u.hostname === '[::1]') {
      return u.protocol === 'http:' || u.protocol === 'https:';
    }
    if (
      u.protocol === 'https:' &&
      (u.hostname.endsWith('.office.com') ||
        u.hostname.endsWith('.office365.com') ||
        u.hostname.endsWith('.sharepoint.com') ||
        u.hostname.endsWith('.live.com'))
    ) {
      return true;
    }
    return false;
  } catch {
    return false;
  }
}

function setCorsHeaders(req: IncomingMessage, res: ServerResponse): boolean {
  const origin = req.headers.origin;
  if (!isAllowedOrigin(origin)) {
    return false;
  }
  if (origin) {
    res.setHeader('Access-Control-Allow-Origin', origin);
    res.setHeader('Vary', 'Origin');
  }
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With');
  res.setHeader('Access-Control-Allow-Private-Network', 'true');
  return true;
}

function serveStaticFile(
  res: ServerResponse,
  filePath: string,
  replacements?: Record<string, string>
): boolean {
  if (!fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) {
    return false;
  }
  const ext = path.extname(filePath).toLowerCase();
  const mime = MIME_TYPES[ext] || 'application/octet-stream';

  if (replacements && ['.xml', '.js', '.html', '.json'].includes(ext)) {
    let text = fs.readFileSync(filePath, 'utf8');
    for (const [k, v] of Object.entries(replacements)) {
      text = text.split(k).join(v);
    }
    res.writeHead(200, { 'Content-Type': mime });
    res.end(text);
    return true;
  }

  res.writeHead(200, { 'Content-Type': mime });
  fs.createReadStream(filePath).pipe(res);
  return true;
}

export async function startGateway(options?: GatewayOptions): Promise<RunningGateway> {
  const store = new ConfigStore(options?.dataDir);
  const paths = resolveGatewayPaths();
  const cfg = store.getConfig();
  const requestedHttpPort = options?.httpPort ?? cfg.settings.httpPort ?? 18080;
  const requestedHttpsPort = options?.httpsPort ?? cfg.settings.httpsPort ?? 18443;
  const host = options?.host ?? '127.0.0.1';
  const enableHttps = options?.enableHttps ?? true;

  let actualHttpPort = requestedHttpPort;
  let actualHttpsPort = requestedHttpsPort;

  const getPortReplacements = () => ({
    'https://127.0.0.1:18443': `https://127.0.0.1:${actualHttpsPort}`,
    'http://127.0.0.1:18080': `http://127.0.0.1:${actualHttpPort}`
  });

  const mutableSettings = () => store.getConfig().settings;
  const launchTarget: LaunchTarget = resolveLaunchTarget({
    customCommand: mutableSettings().gatewayLaunchCommand,
    dataDir: store.dataDir
  });
  let lifecycle: LifecycleMonitor | null = null;
  let lifecycleStarter: (settings: AppSettings) => void = () => {};

  const handler = async (req: IncomingMessage, res: ServerResponse) => {
    const allowedOrigin = setCorsHeaders(req, res);

    if (req.method === 'OPTIONS') {
      res.writeHead(allowedOrigin ? 204 : 403);
      res.end();
      return;
    }

    if (!allowedOrigin) {
      sendJson(res, 403, { error: 'Origin not allowed' });
      return;
    }

    const urlObj = new URL(req.url || '/', `http://${req.headers.host || '127.0.0.1'}`);
    const pathname = decodeURIComponent(urlObj.pathname);
    lifecycle?.touch();

    try {
      // --- API routes ---
      if (pathname === '/api/health' && req.method === 'GET') {
        sendJson(res, 200, {
          ok: true,
          service: 'officeai-gateway',
          version: '1.0.0',
          platform: process.platform,
          httpPort: actualHttpPort,
          httpsPort: enableHttps ? actualHttpsPort : null,
          dataDir: store.dataDir
        });
        return;
      }

      // --- On-demand launch helpers (used by the WPS add-in & Office deep link) ---
      if (pathname === '/api/office/status' && req.method === 'GET') {
        sendJson(res, 200, {
          ok: true,
          gatewayHttpUrl: httpUrl,
          gatewayHttpsUrl: httpsUrl ?? null,
          running: true,
          startPageUrl: `${httpUrl}/start`,
          launchUrl: buildGatewayLaunchUrl(httpUrl, urlObj.searchParams.get('app') ?? undefined),
          launchTarget: launchTarget.command
            ? {
                command: launchTarget.command,
                kind: launchTarget.kind,
                source: launchTarget.source
              }
            : null,
          settings: {
            launchMode: mutableSettings().launchMode,
            exitWithOffice: mutableSettings().exitWithOffice,
            idleExitMinutes: mutableSettings().idleExitMinutes,
            sessionGraceMinutes: mutableSettings().sessionGraceMinutes
          },
          lifecycle: lifecycle?.getStatus() ?? null
        });
        return;
      }

      // Heartbeat from an open add-in pane (Word / Excel / PPT / WPS).
      // Keeps the gateway alive while the office window is in use and lets it
      // exit as soon as every pane is gone.
      if (pathname === '/api/office/session' && req.method === 'POST') {
        const body = await readJsonBody<{ app?: string; source?: string }>(req).catch(
          () => ({}) as { app?: string; source?: string }
        );
        const app = (body.app || urlObj.searchParams.get('app') || 'unknown').toString();
        const source = (body.source || 'addin').toString();
        lifecycle?.watch({ app, source });
        sendJson(res, 200, { ok: true, session: { app, source } });
        return;
      }

      if (pathname === '/start' && req.method === 'GET') {
        const html = renderLauncherPage({
          gatewayUrl: httpUrl,
          launchCommand: launchTarget.command
        });
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' });
        res.end(html);
        return;
      }

      if (pathname === '/api/config' && req.method === 'GET') {
        sendJson(res, 200, store.getPublicConfig());
        return;
      }

      if (pathname === '/api/config/settings' && req.method === 'PUT') {
        const body = await readJsonBody<Partial<AppSettings>>(req);
        const previousMode = store.getConfig().settings.launchMode;
        const updated = store.updateSettings(body);
        if (updated.launchMode !== previousMode) {
          lifecycleStarter(updated);
        }
        sendJson(res, 200, { ok: true, settings: updated });
        return;
      }

      if (pathname === '/api/providers' && req.method === 'GET') {
        sendJson(res, 200, store.getPublicConfig().providers);
        return;
      }

      if (pathname === '/api/providers' && req.method === 'POST') {
        const body = await readJsonBody<Partial<ProviderConfig> & { id: string; clearApiKey?: boolean }>(req);
        if (!body.id || !body.id.trim()) {
          sendJson(res, 400, { error: 'Provider id is required' });
          return;
        }
        if (body.protocol && !['openai', 'anthropic', 'gemini', 'ollama'].includes(body.protocol)) {
          sendJson(res, 400, { error: 'Unsupported provider protocol' });
          return;
        }
        if (body.baseUrl && !/^https?:\/\//i.test(body.baseUrl.trim())) {
          sendJson(res, 400, { error: 'Base URL must start with http:// or https://' });
          return;
        }
        store.upsertProvider({ ...body, id: body.id.trim() });
        sendJson(res, 200, { ok: true, config: store.getPublicConfig() });
        return;
      }

      const providerMatch = pathname.match(/^\/api\/providers\/([^/]+)$/);
      if (providerMatch && req.method === 'PUT') {
        const id = decodeURIComponent(providerMatch[1]);
        const body = await readJsonBody<Partial<ProviderConfig> & { clearApiKey?: boolean }>(req);
        if (body.protocol && !['openai', 'anthropic', 'gemini', 'ollama'].includes(body.protocol)) {
          sendJson(res, 400, { error: 'Unsupported provider protocol' });
          return;
        }
        if (body.baseUrl && !/^https?:\/\//i.test(body.baseUrl.trim())) {
          sendJson(res, 400, { error: 'Base URL must start with http:// or https://' });
          return;
        }
        store.upsertProvider({ ...body, id });
        sendJson(res, 200, { ok: true, config: store.getPublicConfig() });
        return;
      }

      if (providerMatch && req.method === 'DELETE') {
        const id = decodeURIComponent(providerMatch[1]);
        const deleted = store.deleteProvider(id);
        sendJson(res, deleted ? 200 : 404, { ok: deleted });
        return;
      }

      const testMatch = pathname.match(/^\/api\/providers\/([^/]+)\/test$/);
      if (testMatch && req.method === 'POST') {
        const id = decodeURIComponent(testMatch[1]);
        const existing = store.getProvider(id);
        const override = await readJsonBody<Partial<ProviderConfig>>(req).catch(() => ({}));
        const target: ProviderConfig | undefined = existing
          ? {
              ...existing,
              ...override,
              apiKey:
                override && 'apiKey' in override && typeof override.apiKey === 'string' && override.apiKey.length > 0
                  ? override.apiKey
                  : existing.apiKey
            }
          : undefined;

        if (!target) {
          sendJson(res, 404, { ok: false, message: 'Provider not found' });
          return;
        }

        const adapter = resolveProviderAdapter(target.protocol);
        const result = await adapter.testConnection(target);
        if (result.ok && result.models.length > 0) {
          store.upsertProvider({
            id: target.id,
            models: result.models,
            selectedModel: result.models.includes(target.selectedModel)
              ? target.selectedModel
              : result.models[0]
          });
        }
        sendJson(res, 200, result);
        return;
      }

      const modelsMatch = pathname.match(/^\/api\/providers\/([^/]+)\/models$/);
      if (modelsMatch && req.method === 'GET') {
        const id = decodeURIComponent(modelsMatch[1]);
        const provider = store.getProvider(id);
        if (!provider) {
          sendJson(res, 404, { error: 'Provider not found' });
          return;
        }
        const adapter = resolveProviderAdapter(provider.protocol);
        const models = await adapter.fetchModels(provider);
        if (models.length > 0) {
          store.upsertProvider({ id: provider.id, models });
        }
        sendJson(res, 200, { models });
        return;
      }

      if (pathname === '/api/local/status' && req.method === 'GET') {
        const statuses = await probeLocalRuntimes(store.getConfig().providers);
        sendJson(res, 200, { runtimes: statuses });
        return;
      }

      if (pathname === '/api/prompts' && req.method === 'GET') {
        sendJson(res, 200, { prompts: store.getPrompts() });
        return;
      }

      if (pathname === '/api/prompts' && req.method === 'POST') {
        const body = await readJsonBody<PromptAction>(req);
        if (!body.id || !body.title) {
          sendJson(res, 400, { error: 'Prompt id and title are required' });
          return;
        }
        const saved = store.upsertPrompt(body);
        sendJson(res, 200, { ok: true, prompt: saved });
        return;
      }

      const promptMatch = pathname.match(/^\/api\/prompts\/([^/]+)$/);
      if (promptMatch && req.method === 'DELETE') {
        const id = decodeURIComponent(promptMatch[1]);
        const ok = store.deletePrompt(id);
        sendJson(res, ok ? 200 : 404, { ok });
        return;
      }

      if (pathname === '/api/usage' && req.method === 'GET') {
        sendJson(res, 200, { records: store.getUsage() });
        return;
      }

      if (pathname === '/api/usage' && req.method === 'DELETE') {
        store.clearUsage();
        sendJson(res, 200, { ok: true });
        return;
      }

      if (pathname === '/api/chat' && req.method === 'POST') {
        const body = await readJsonBody<ChatRequest>(req);
        if (!Array.isArray(body.messages) || body.messages.length === 0) {
          sendJson(res, 400, { error: 'messages 不能为空' });
          return;
        }

        const currentCfg = store.getConfig();
        const providerId = body.providerId || currentCfg.settings.activeProviderId;
        const provider = store.getProvider(providerId);

        if (!provider) {
          sendJson(res, 404, { error: `未找到模型提供方: ${providerId}` });
          return;
        }

        const model = body.model || provider.selectedModel || provider.models[0];
        if (!model) {
          sendJson(res, 400, { error: `模型提供方 ${provider.name} 未选择可用模型` });
          return;
        }

        if (provider.category === 'cloud' && !provider.apiKey.trim()) {
          sendJson(res, 400, {
            error: `请先在配置中心填写 ${provider.name} 的 API Key`
          });
          return;
        }

        const temperature = body.temperature ?? currentCfg.settings.temperature;
        const maxTokens = body.maxTokens ?? currentCfg.settings.maxTokens;
        const stream = body.stream !== false;
        const adapter = resolveProviderAdapter(provider.protocol);

        const controller = new AbortController();
        req.on('close', () => controller.abort());

        const startedAt = Date.now();

        if (!stream) {
          let fullText = '';
          let reasoningText = '';
          let usage: { promptTokens?: number; completionTokens?: number } | undefined;

          try {
            await adapter.streamChat({
              provider,
              model,
              messages: body.messages,
              temperature,
              maxTokens,
              signal: controller.signal,
              onEvent: (ev) => {
                if (ev.type === 'delta' && ev.content) fullText += ev.content;
                if (ev.type === 'reasoning' && ev.content) reasoningText += ev.content;
                if (ev.type === 'done' && ev.usage) usage = ev.usage;
              }
            });

            store.recordUsage({
              id: crypto.randomUUID(),
              timestamp: new Date().toISOString(),
              providerId: provider.id,
              providerName: provider.name,
              model,
              hostApp: body.hostApp ?? 'standalone',
              promptTokens: usage?.promptTokens,
              completionTokens: usage?.completionTokens,
              durationMs: Date.now() - startedAt,
              success: true
            });

            sendJson(res, 200, {
              content: fullText,
              reasoning: reasoningText || undefined,
              model,
              providerId: provider.id,
              usage
            });
          } catch (err) {
            if (!controller.signal.aborted) {
              const message = err instanceof Error ? err.message : String(err);
              store.recordUsage({
                id: crypto.randomUUID(),
                timestamp: new Date().toISOString(),
                providerId: provider.id,
                providerName: provider.name,
                model,
                hostApp: body.hostApp ?? 'standalone',
                durationMs: Date.now() - startedAt,
                success: false,
                error: message
              });
              sendJson(res, 502, { error: message });
            }
          }
          return;
        }

        res.writeHead(200, {
          'Content-Type': 'text/event-stream; charset=utf-8',
          'Cache-Control': 'no-cache, no-transform',
          Connection: 'keep-alive',
          'X-Accel-Buffering': 'no'
        });

        const writeEvent = (ev: StreamEvent) => {
          if (!res.writableEnded) {
            res.write(`data: ${JSON.stringify(ev)}\n\n`);
          }
        };

        writeEvent({ type: 'meta', providerId: provider.id, model });

        let usage: { promptTokens?: number; completionTokens?: number } | undefined;
        try {
          await adapter.streamChat({
            provider,
            model,
            messages: body.messages,
            temperature,
            maxTokens,
            signal: controller.signal,
            onEvent: (ev) => {
              if (ev.type === 'done') {
                usage = ev.usage;
              }
              writeEvent(ev);
            }
          });
          store.recordUsage({
            id: crypto.randomUUID(),
            timestamp: new Date().toISOString(),
            providerId: provider.id,
            providerName: provider.name,
            model,
            hostApp: body.hostApp ?? 'standalone',
            promptTokens: usage?.promptTokens,
            completionTokens: usage?.completionTokens,
            durationMs: Date.now() - startedAt,
            success: true
          });
        } catch (err) {
          const isAborted =
            controller.signal.aborted ||
            (err instanceof Error && err.name === 'AbortError');
          if (!isAborted) {
            const message = err instanceof Error ? err.message : String(err);
            store.recordUsage({
              id: crypto.randomUUID(),
              timestamp: new Date().toISOString(),
              providerId: provider.id,
              providerName: provider.name,
              model,
              hostApp: body.hostApp ?? 'standalone',
              durationMs: Date.now() - startedAt,
              success: false,
              error: message
            });
            writeEvent({ type: 'error', error: message });
          }
        } finally {
          if (!res.writableEnded) {
            res.end();
          }
        }
        return;
      }

      // --- Windows integration APIs ---
      if (pathname === '/api/windows/status' && req.method === 'GET') {
        const status = await integrator.getStatus();
        sendJson(res, 200, status);
        return;
      }

      if (pathname === '/api/windows/trust-cert' && req.method === 'POST') {
        const result = await integrator.trustLocalCertificate();
        sendJson(res, 200, result);
        return;
      }

      if (pathname === '/api/windows/register-protocol' && req.method === 'POST') {
        const exe = launchTarget.kind === 'exe' ? launchTarget.command : '';
        if (!exe) {
          sendJson(res, 200, {
            ok: false,
            message:
              '未找到 OfficeAI.exe，无法注册 officeai:// 协议。请在「启动方式」中填写 OfficeAI 可执行文件路径。'
          });
          return;
        }
        const result = await registerProtocolHandler(exe);
        if (result.ok) {
          const shortcut = await createLauncherShortcut(buildGatewayLaunchUrl(httpUrl, 'office'));
          result.message = `${result.message}${shortcut.ok ? ` ${shortcut.message}` : ''}`;
        }
        sendJson(res, 200, result);
        return;
      }

      if (pathname === '/api/windows/unregister-protocol' && req.method === 'POST') {
        const result = await unregisterProtocolHandler();
        await removeLauncherShortcut();
        sendJson(res, 200, result);
        return;
      }

      if (pathname === '/api/windows/install-waker' && req.method === 'POST') {
        const result = await integrator.installOfficeWaker();
        sendJson(res, 200, result);
        return;
      }

      if (pathname === '/api/windows/uninstall-waker' && req.method === 'POST') {
        const result = await integrator.uninstallOfficeWaker();
        sendJson(res, 200, result);
        return;
      }

      if (pathname === '/api/windows/install-office' && req.method === 'POST') {
        const result = await integrator.installMsOffice();
        sendJson(res, 200, result);
        return;
      }

      if (pathname === '/api/windows/uninstall-office' && req.method === 'POST') {
        const result = await integrator.uninstallMsOffice();
        sendJson(res, 200, result);
        return;
      }

      if (pathname === '/api/windows/install-wps' && req.method === 'POST') {
        const result = await integrator.installWps();
        sendJson(res, 200, result);
        return;
      }

      if (pathname === '/api/windows/uninstall-wps' && req.method === 'POST') {
        const result = await integrator.uninstallWps();
        sendJson(res, 200, result);
        return;
      }

      // --- Static routes for MS Office Add-in & WPS JS Add-in ---
      if (pathname.startsWith('/office/')) {
        const rel = pathname.slice('/office/'.length) || 'manifest.xml';
        const target = path.resolve(paths.officeAddinDir, rel);
        if (isWithinDir(target, paths.officeAddinDir) && serveStaticFile(res, target, getPortReplacements())) {
          return;
        }
        sendJson(res, 404, { error: 'Not found', path: pathname });
        return;
      }

      // Launcher page opened inside WPS right after `OAAssist.ShellExecute`:
      // it polls /api/health and navigates to the assistant once ready.
      const wpsBootstrap = path.resolve(paths.wpsAddinDir, 'launcher.html');
      if ((pathname === '/wps/bootstrap.html' || pathname === '/bootstrap.html') && fs.existsSync(wpsBootstrap)) {
        if (serveStaticFile(res, wpsBootstrap)) {
          return;
        }
      }

      if (pathname === '/wps' || pathname.startsWith('/wps/')) {
        const rel = pathname === '/wps' ? 'index.html' : pathname.slice('/wps/'.length) || 'index.html';
        const target = path.resolve(paths.wpsAddinDir, rel);
        if (isWithinDir(target, paths.wpsAddinDir) && serveStaticFile(res, target, getPortReplacements())) {
          return;
        }
        sendJson(res, 404, { error: 'Not found', path: pathname });
        return;
      }

      // --- Static UI (Add-in TaskPane + Desktop Control Center) ---
      if (fs.existsSync(paths.uiDistDir)) {
        const cleanRel = pathname === '/' ? 'index.html' : pathname.replace(/^\/+/, '');
        const candidate = path.resolve(paths.uiDistDir, cleanRel);
        if (isWithinDir(candidate, paths.uiDistDir) && serveStaticFile(res, candidate)) {
          return;
        }
        // SPA fallback only for extensionless client-side routes (never for missing assets or /api/*)
        if (!pathname.startsWith('/api/') && !path.extname(pathname)) {
          const indexHtml = path.join(paths.uiDistDir, 'index.html');
          if (serveStaticFile(res, indexHtml)) {
            return;
          }
        }
      }

      sendJson(res, 404, { error: 'Not found', path: pathname });
    } catch (err) {
      if (!res.headersSent) {
        sendJson(res, 500, {
          error: err instanceof Error ? err.message : String(err)
        });
      } else {
        res.end();
      }
    }
  };

  const httpServer = http.createServer(handler);
  let integrator!: WindowsOfficeIntegrator;

  await new Promise<void>((resolve, reject) => {
    httpServer.once('error', reject);
    httpServer.listen(requestedHttpPort, host, () => resolve());
  });

  actualHttpPort = (httpServer.address() as { port: number }).port;
  const httpUrl = `http://${host}:${actualHttpPort}`;

  let httpsServer: https.Server | undefined;
  let httpsUrl: string | undefined;

  if (enableHttps) {
    const certs = ensureLocalhostCert(store.dataDir);
    httpsServer = https.createServer({ key: certs.key, cert: certs.cert }, handler);
    await new Promise<void>((resolve, reject) => {
      httpsServer!.once('error', reject);
      httpsServer!.listen(requestedHttpsPort, host, () => resolve());
    });
    actualHttpsPort = (httpsServer.address() as { port: number }).port;
    httpsUrl = `https://${host}:${actualHttpsPort}`;
  }

  integrator = new WindowsOfficeIntegrator(
    store.dataDir,
    actualHttpPort,
    actualHttpsPort,
    paths.assetsRoot,
    launchTarget.command,
    () => {
      if (!lifecycle) return null;
      const status = lifecycle.getStatus();
      return {
        mode: status.mode,
        launcher: status.launcher,
        upMinutes: status.upMinutes,
        lastActivitySeconds: status.lastActivitySeconds,
        sessions: status.sessions
      };
    }
  );

  const onExitRequested =
    options?.onExitRequested ??
    ((reason: string) => {
      console.log(`[OfficeAI] ${reason}，网关将自动退出。`);
      setTimeout(() => process.exit(0), 150);
    });

  lifecycle = new LifecycleMonitor(() => store.getConfig().settings, {
    launcher: options?.launcher,
    onExitRequested
  });

  function restartLifecycle(nextSettings: AppSettings): void {
    lifecycle?.stop();
    if (nextSettings.launchMode === 'on-demand') {
      lifecycle?.start();
    }
  }

  // Expose the lifecycle starter on the handler scope (used by PUT /api/config/settings)
  lifecycleStarter = restartLifecycle;

  if (store.getConfig().settings.launchMode === 'on-demand') {
    lifecycle.start();
  }

  const close = async () => {
    lifecycle?.stop();
    await Promise.all([
      new Promise<void>((resolve) => httpServer.close(() => resolve())),
      httpsServer
        ? new Promise<void>((resolve) => httpsServer!.close(() => resolve()))
        : Promise.resolve()
    ]);
  };

  return {
    httpServer,
    httpsServer,
    httpUrl,
    httpsUrl,
    store,
    integrator,
    lifecycle,
    launchTarget,
    close
  };
}
