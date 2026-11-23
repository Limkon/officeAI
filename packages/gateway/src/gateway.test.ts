import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import assert from 'node:assert/strict';
import { isAllowedOrigin, startGateway } from './server.js';
import { ConfigStore } from './config-store.js';

function createTempDir(prefix: string): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix));
}

test('isAllowedOrigin permits loopback, Office domains, and null, while blocking third-party sites', () => {
  assert.equal(isAllowedOrigin(undefined), true);
  assert.equal(isAllowedOrigin('null'), true);
  assert.equal(isAllowedOrigin('http://127.0.0.1:18080'), true);
  assert.equal(isAllowedOrigin('https://127.0.0.1:18443'), true);
  assert.equal(isAllowedOrigin('http://localhost:5173'), true);
  assert.equal(isAllowedOrigin('https://word-edit.officeapps.live.com'), true);
  assert.equal(isAllowedOrigin('https://evil.example.com'), false);
});

test('ConfigStore masks API keys, supports clearApiKey, and falls back activeProviderId on delete', () => {
  const dir = createTempDir('officeai-cfg-');
  const store = new ConfigStore(dir);

  store.upsertProvider({
    id: 'openai',
    apiKey: 'sk-secret-key-9999',
    selectedModel: 'gpt-4o'
  });

  const pub = store.getPublicConfig();
  const openai = pub.providers.find((p) => p.id === 'openai');
  assert.ok(openai);
  assert.equal(openai.hasApiKey, true);
  assert.equal(openai.apiKeyMasked, 'sk-s••••9999');
  assert.equal('apiKey' in openai, false);

  // Updating model without passing apiKey should keep existing secret
  store.upsertProvider({ id: 'openai', selectedModel: 'gpt-4.1' });
  assert.equal(store.getProvider('openai')?.apiKey, 'sk-secret-key-9999');

  // Explicit clearApiKey removes secret
  store.upsertProvider({ id: 'openai', clearApiKey: true });
  assert.equal(store.getProvider('openai')?.apiKey, '');

  // Custom provider deletion updates activeProviderId fallback
  store.upsertProvider({ id: 'custom-temp', name: 'Temp', enabled: true });
  store.updateSettings({ activeProviderId: 'custom-temp', activeModel: 'gpt-4o' });
  assert.equal(store.deleteProvider('custom-temp'), true);
  assert.notEqual(store.getConfig().settings.activeProviderId, 'custom-temp');

  fs.rmSync(dir, { recursive: true, force: true });
});

test('Gateway serves config, handles CORS/path-traversal, stream_options fallback, Gemini header, and manifests', async () => {
  const dataDir = createTempDir('officeai-gw-');
  let sawGeminiHeaderKey = '';
  let openaiAttemptCount = 0;

  const mockUpstream = http.createServer(async (req, res) => {
    const chunks: Buffer[] = [];
    for await (const c of req) chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c));
    const bodyText = Buffer.concat(chunks).toString('utf8');

    if (req.url === '/v1/models' && req.method === 'GET') {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ data: [{ id: 'mock-gpt-4o' }, { id: 'mock-mini' }] }));
      return;
    }

    if (req.url === '/v1/chat/completions' && req.method === 'POST') {
      openaiAttemptCount++;
      const parsed = JSON.parse(bodyText);
      // Simulate upstream that rejects stream_options on first attempt
      if (parsed.stream_options) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Unrecognized request argument supplied: stream_options' }));
        return;
      }
      res.writeHead(200, { 'Content-Type': 'text/event-stream' });
      res.write(
        'data: ' +
          JSON.stringify({ choices: [{ delta: { content: '<think>plan</think>Hello ' } }] }) +
          '\n\n'
      );
      res.write(
        'data: ' +
          JSON.stringify({
            choices: [{ delta: { content: 'OfficeAI!' } }],
            usage: { prompt_tokens: 10, completion_tokens: 5 }
          }) +
          '\n\n'
      );
      res.write('data: [DONE]\n\n');
      res.end();
      return;
    }

    if (req.url?.startsWith('/v1beta/models/gemini-2.5-flash:streamGenerateContent')) {
      sawGeminiHeaderKey = String(req.headers['x-goog-api-key'] || '');
      res.writeHead(200, { 'Content-Type': 'text/event-stream' });
      res.write(
        'data: ' +
          JSON.stringify({
            candidates: [{ content: { parts: [{ text: 'Gemini OK' }] } }],
            usageMetadata: { promptTokenCount: 4, candidatesTokenCount: 2 }
          }) +
          '\n\n'
      );
      res.end();
      return;
    }

    res.writeHead(404);
    res.end();
  });

  await new Promise<void>((resolve) => mockUpstream.listen(0, '127.0.0.1', () => resolve()));
  const mockPort = (mockUpstream.address() as { port: number }).port;

  const gw = await startGateway({
    dataDir,
    httpPort: 0,
    httpsPort: 0,
    enableHttps: true
  });

  try {
    // 1. CORS rejection for untrusted web origin
    const evilRes = await fetch(`${gw.httpUrl}/api/config`, {
      headers: { Origin: 'https://evil.example.com' }
    });
    assert.equal(evilRes.status, 403);

    // 2. Path traversal blocked (both encoded %2e%2e and missing asset extension)
    const travRes = await fetch(`${gw.httpUrl}/office/%2e%2e/package.json`);
    assert.equal(travRes.status, 404);
    const missingAssetRes = await fetch(`${gw.httpUrl}/package.json`);
    assert.equal(missingAssetRes.status, 404);

    // 3. Configure OpenAI & test stream_options retry fallback
    await fetch(`${gw.httpUrl}/api/providers/openai`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        baseUrl: `http://127.0.0.1:${mockPort}/v1`,
        apiKey: 'sk-test-123456',
        selectedModel: 'mock-gpt-4o'
      })
    });

    const chatRes = await fetch(`${gw.httpUrl}/api/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        providerId: 'openai',
        model: 'mock-gpt-4o',
        stream: false,
        hostApp: 'wps-writer',
        messages: [{ role: 'user', content: 'hi' }]
      })
    });
    assert.equal(chatRes.status, 200);
    const chatJson = (await chatRes.json()) as {
      content: string;
      reasoning?: string;
      usage?: { promptTokens?: number; completionTokens?: number };
    };
    assert.equal(openaiAttemptCount, 2); // first failed on stream_options, second succeeded
    assert.equal(chatJson.reasoning, 'plan');
    assert.equal(chatJson.content, 'Hello OfficeAI!');
    assert.equal(chatJson.usage?.promptTokens, 10);

    // 4. Configure Gemini & verify x-goog-api-key header (not URL query param)
    await fetch(`${gw.httpUrl}/api/providers/gemini`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        baseUrl: `http://127.0.0.1:${mockPort}`,
        apiKey: 'AIza-test-header-key',
        selectedModel: 'gemini-2.5-flash'
      })
    });

    const gemRes = await fetch(`${gw.httpUrl}/api/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        providerId: 'gemini',
        model: 'gemini-2.5-flash',
        stream: false,
        messages: [{ role: 'user', content: 'hi' }]
      })
    });
    assert.equal(gemRes.status, 200);
    const gemJson = (await gemRes.json()) as { content: string };
    assert.equal(gemJson.content, 'Gemini OK');
    assert.equal(sawGeminiHeaderKey, 'AIza-test-header-key');

    // 5. Verify manifest.xml and ribbon.xml static serving with dynamic port replacement
    const manifestRes = await fetch(`${gw.httpUrl}/office/manifest.xml`);
    assert.equal(manifestRes.status, 200);
    const manifestText = await manifestRes.text();
    assert.ok(manifestText.includes('<OfficeApp'));
    assert.ok(gw.httpsUrl && manifestText.includes(gw.httpsUrl));

    const ribbonRes = await fetch(`${gw.httpUrl}/wps/ribbon.xml`);
    assert.equal(ribbonRes.status, 200);
    const ribbonText = await ribbonRes.text();
    assert.ok(ribbonText.includes('OfficeAITab'));
  } finally {
    await gw.close();
    await new Promise<void>((resolve) => mockUpstream.close(() => resolve()));
    fs.rmSync(dataDir, { recursive: true, force: true });
  }
});

test('On-demand launch: settings normalize, start page & office status are served', async () => {
  const dataDir = createTempDir('officeai-launch-');
  const gw = await startGateway({
    dataDir,
    httpPort: 0,
    httpsPort: 0,
    enableHttps: false,
    launcher: 'test'
  });

  try {
    // Default is on-demand, never autostart
    const cfgRes = await fetch(`${gw.httpUrl}/api/config`);
    const cfg = (await cfgRes.json()) as { settings: { launchMode: string; autoStartGateway: boolean } };
    assert.equal(cfg.settings.launchMode, 'on-demand');
    assert.equal(cfg.settings.autoStartGateway, false);

    // Legacy payload `{ autoStartGateway: true }` maps to launchMode=always
    const putRes = await fetch(`${gw.httpUrl}/api/config/settings`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ autoStartGateway: true })
    });
    const putJson = (await putRes.json()) as { settings: { launchMode: string } };
    assert.equal(putJson.settings.launchMode, 'always');

    const idleRes = await fetch(`${gw.httpUrl}/api/config/settings`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ launchMode: 'on-demand', idleExitMinutes: 999 })
    });
    const idleJson = (await idleRes.json()) as { settings: { idleExitMinutes: number } };
    assert.equal(idleJson.settings.idleExitMinutes, 240);

    // launchTrigger must stay one of the two supported values
    const triggerRes = await fetch(`${gw.httpUrl}/api/config/settings`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ launchTrigger: 'waker' })
    });
    const triggerJson = (await triggerRes.json()) as { settings: { launchTrigger: string } };
    assert.equal(triggerJson.settings.launchTrigger, 'waker');

    const badTrigger = await fetch(`${gw.httpUrl}/api/config/settings`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ launchTrigger: 'nonsense' })
    });
    const badJson = (await badTrigger.json()) as { settings: { launchTrigger: string } };
    assert.equal(badJson.settings.launchTrigger, 'waker');

    // The launcher page is what WPS/Office users see when the gateway is down
    const startRes = await fetch(`${gw.httpUrl}/start`);
    assert.equal(startRes.status, 200);
    const startHtml = await startRes.text();
    assert.ok(startHtml.includes('OfficeAI 服务未启动'));
    assert.ok(startHtml.includes('officeai://activate/'));

    // /api/office/status is polled by the local WPS add-in
    const statusRes = await fetch(`${gw.httpUrl}/api/office/status?app=wps`);
    assert.equal(statusRes.status, 200);
    const status = (await statusRes.json()) as {
      running: boolean;
      gatewayHttpUrl: string;
      startPageUrl: string;
      launchUrl: string;
      launchTarget: unknown;
    };
    assert.equal(status.running, true);
    assert.equal(status.gatewayHttpUrl, gw.httpUrl);
    assert.equal(status.startPageUrl, `${gw.httpUrl}/start`);
    assert.ok(status.launchUrl.startsWith('officeai://activate/'));

    // Task pane heartbeat registers a live session that holds the gateway
    const beatRes = await fetch(`${gw.httpUrl}/api/office/session`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ app: 'word', source: 'addin' })
    });
    assert.equal(beatRes.status, 200);
    const officeStatus2 = (await (
      await fetch(`${gw.httpUrl}/api/office/status`)
    ).json()) as { lifecycle: { sessions: Array<{ app: string }> } };
    assert.equal(officeStatus2.lifecycle.sessions.length, 1);
    assert.equal(officeStatus2.lifecycle.sessions[0].app, 'word');

    // WPS bootstrap launcher page is bundled with the add-in
    const bootRes = await fetch(`${gw.httpUrl}/wps/bootstrap.html`);
    assert.equal(bootRes.status, 200);
  } finally {
    await gw.close();
    fs.rmSync(dataDir, { recursive: true, force: true });
  }
});
