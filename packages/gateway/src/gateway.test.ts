import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import assert from 'node:assert/strict';
import { startGateway } from './server.js';
import { ConfigStore } from './config-store.js';

function createTempDir(prefix: string): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix));
}

test('ConfigStore masks API keys and preserves secrets on partial update', () => {
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

  fs.rmSync(dir, { recursive: true, force: true });
});

test('Gateway serves config, mock OpenAI/Ollama streaming chat, and static manifests', async () => {
  const dataDir = createTempDir('officeai-gw-');

  // Start mock upstream server simulating OpenAI + Ollama
  const mockUpstream = http.createServer(async (req, res) => {
    if (req.url === '/v1/models' && req.method === 'GET') {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ data: [{ id: 'mock-gpt-4o' }, { id: 'mock-mini' }] }));
      return;
    }
    if (req.url === '/v1/chat/completions' && req.method === 'POST') {
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
    if (req.url === '/api/tags' && req.method === 'GET') {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ models: [{ name: 'qwen2.5:7b' }] }));
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
    // 1. Health check
    const healthRes = await fetch(`${gw.httpUrl}/api/health`);
    assert.equal(healthRes.status, 200);
    const health = (await healthRes.json()) as { ok: boolean };
    assert.equal(health.ok, true);

    // 2. Configure openai provider against mock upstream
    await fetch(`${gw.httpUrl}/api/providers/openai`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        baseUrl: `http://127.0.0.1:${mockPort}/v1`,
        apiKey: 'sk-test-123456',
        selectedModel: 'mock-gpt-4o'
      })
    });

    // 3. Test connection
    const testRes = await fetch(`${gw.httpUrl}/api/providers/openai/test`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: '{}'
    });
    const testJson = (await testRes.json()) as { ok: boolean; models: string[] };
    assert.equal(testJson.ok, true);
    assert.deepEqual(testJson.models, ['mock-gpt-4o', 'mock-mini']);

    // 4. Test non-stream chat endpoint (verifies ThinkTagSplitter + usage)
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
    const chatJson = (await chatRes.json()) as {
      content: string;
      reasoning?: string;
      usage?: { promptTokens?: number; completionTokens?: number };
    };
    assert.equal(chatJson.reasoning, 'plan');
    assert.equal(chatJson.content, 'Hello OfficeAI!');
    assert.equal(chatJson.usage?.promptTokens, 10);

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
