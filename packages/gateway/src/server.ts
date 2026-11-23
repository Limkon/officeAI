import crypto from 'node:crypto';
import fs from 'node:fs';
import http, { type IncomingMessage, type ServerResponse } from 'node:http';
import https from 'node:https';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import type {
  AppSettings,
  ChatRequest,
  PromptAction,
  ProviderConfig,
  StreamEvent
} from '@officeai/shared';
import { ensureLocalhostCert } from './certs.js';
import { ConfigStore } from './config-store.js';
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
  close: () => Promise<void>;
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

function setCorsHeaders(req: IncomingMessage, res: ServerResponse): void {
  const origin = req.headers.origin || '*';
  res.setHeader('Access-Control-Allow-Origin', origin);
  res.setHeader('Vary', 'Origin');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With');
  res.setHeader('Access-Control-Allow-Private-Network', 'true');
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

export async function startGateway(options?: {
  dataDir?: string;
  httpPort?: number;
  httpsPort?: number;
  host?: string;
  enableHttps?: boolean;
}): Promise<RunningGateway> {
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

  const handler = async (req: IncomingMessage, res: ServerResponse) => {
    setCorsHeaders(req, res);

    if (req.method === 'OPTIONS') {
      res.writeHead(204);
      res.end();
      return;
    }

    const urlObj = new URL(req.url || '/', `http://${req.headers.host || '127.0.0.1'}`);
    const pathname = decodeURIComponent(urlObj.pathname);

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

      if (pathname === '/api/config' && req.method === 'GET') {
        sendJson(res, 200, store.getPublicConfig());
        return;
      }

      if (pathname === '/api/config/settings' && req.method === 'PUT') {
        const body = await readJsonBody<Partial<AppSettings>>(req);
        const updated = store.updateSettings(body);
        sendJson(res, 200, { ok: true, settings: updated });
        return;
      }

      if (pathname === '/api/providers' && req.method === 'GET') {
        sendJson(res, 200, store.getPublicConfig().providers);
        return;
      }

      if (pathname === '/api/providers' && req.method === 'POST') {
        const body = await readJsonBody<Partial<ProviderConfig> & { id: string }>(req);
        if (!body.id) {
          sendJson(res, 400, { error: 'Provider id is required' });
          return;
        }
        store.upsertProvider(body);
        sendJson(res, 200, { ok: true, config: store.getPublicConfig() });
        return;
      }

      const providerMatch = pathname.match(/^\/api\/providers\/([^/]+)$/);
      if (providerMatch && req.method === 'PUT') {
        const id = decodeURIComponent(providerMatch[1]);
        const body = await readJsonBody<Partial<ProviderConfig>>(req);
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
          return;
        }

        res.writeHead(200, {
          'Content-Type': 'text/event-stream; charset=utf-8',
          'Cache-Control': 'no-cache, no-transform',
          Connection: 'keep-alive',
          'X-Accel-Buffering': 'no'
        });

        const writeEvent = (ev: StreamEvent) => {
          res.write(`data: ${JSON.stringify(ev)}\n\n`);
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
        } finally {
          res.end();
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
        if (target.startsWith(paths.officeAddinDir) && serveStaticFile(res, target, getPortReplacements())) {
          return;
        }
      }

      if (pathname === '/wps' || pathname.startsWith('/wps/')) {
        const rel = pathname === '/wps' ? 'index.html' : pathname.slice('/wps/'.length) || 'index.html';
        const target = path.resolve(paths.wpsAddinDir, rel);
        if (target.startsWith(paths.wpsAddinDir) && serveStaticFile(res, target, getPortReplacements())) {
          return;
        }
      }

      // --- Static UI (Add-in TaskPane + Desktop Control Center) ---
      if (fs.existsSync(paths.uiDistDir)) {
        const cleanRel = pathname === '/' ? 'index.html' : pathname.replace(/^\/+/, '');
        const candidate = path.resolve(paths.uiDistDir, cleanRel);
        if (candidate.startsWith(paths.uiDistDir) && serveStaticFile(res, candidate)) {
          return;
        }
        // SPA fallback
        const indexHtml = path.join(paths.uiDistDir, 'index.html');
        if (serveStaticFile(res, indexHtml)) {
          return;
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
    paths.assetsRoot
  );

  return {
    httpServer,
    httpsServer,
    httpUrl,
    httpsUrl,
    store,
    integrator,
    close: async () => {
      await Promise.all([
        new Promise<void>((resolve) => httpServer.close(() => resolve())),
        httpsServer
          ? new Promise<void>((resolve) => httpsServer!.close(() => resolve()))
          : Promise.resolve()
      ]);
    }
  };
}
