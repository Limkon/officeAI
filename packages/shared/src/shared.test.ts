import test from 'node:test';
import assert from 'node:assert/strict';
import {
  ThinkTagSplitter,
  maskApiKey,
  renderPromptTemplate,
  createDefaultConfig,
  PROVIDER_PRESETS
} from './index.js';

test('ThinkTagSplitter splits reasoning and answer across chunk boundaries', () => {
  const splitter = new ThinkTagSplitter();
  const out1 = splitter.push('<thi');
  const out2 = splitter.push('nk>step 1\nstep 2</thi');
  const out3 = splitter.push('nk>Final answer');
  const out4 = splitter.flush();

  const reasoning = out1.reasoning + out2.reasoning + out3.reasoning + out4.reasoning;
  const content = out1.content + out2.content + out3.content + out4.content;

  assert.equal(reasoning, 'step 1\nstep 2');
  assert.equal(content, 'Final answer');
});

test('maskApiKey masks secrets safely', () => {
  assert.equal(maskApiKey(''), '');
  assert.equal(maskApiKey('short'), '****');
  assert.equal(maskApiKey('sk-1234567890abcdef'), 'sk-1••••cdef');
});

test('renderPromptTemplate substitutes variables', () => {
  const rendered = renderPromptTemplate('Ctx:{{context}}\nSel:{{selection}}', {
    context: 'A1:B2',
    selection: 'Sum column B'
  });
  assert.equal(rendered, 'Ctx:A1:B2\nSel:Sum column B');
});

test('createDefaultConfig includes all cloud and local presets', () => {
  const cfg = createDefaultConfig();
  assert.equal(cfg.providers.length, PROVIDER_PRESETS.length);
  assert.ok(cfg.providers.some((p) => p.id === 'ollama' && p.category === 'local'));
  assert.ok(cfg.providers.some((p) => p.id === 'anthropic' && p.protocol === 'anthropic'));
  assert.ok(cfg.providers.some((p) => p.id === 'gemini' && p.protocol === 'gemini'));
});
