import test from 'node:test';
import assert from 'node:assert/strict';
import {
  ThinkTagSplitter,
  maskApiKey,
  renderPromptTemplate,
  createDefaultConfig,
  PROVIDER_PRESETS,
  parseOfficeActions,
  stripStreamingOfficeActionTags,
  buildOfficeAgentSystemPrompt,
  buildGatewayLaunchUrl,
  parseGatewayLaunchUrl
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

test('parseOfficeActions extracts structured Word/Excel commands and strips tags', () => {
  const raw = [
    '已为你准备好汇总公式和表头格式：',
    '<office_action>',
    '{"action":"set_cells","summary":"在 E2 写入求和公式","range":"E2","formula":"=SUM(B2:D2)"}',
    '</office_action>',
    '<office_action>```json\n{"action":"format_range","summary":"表头加粗","range":"A1:E1","bold":true}\n```</office_action>'
  ].join('\n');

  const parsed = parseOfficeActions(raw);
  assert.equal(parsed.displayText, '已为你准备好汇总公式和表头格式：');
  assert.equal(parsed.actions.length, 2);
  assert.equal(parsed.actions[0].action, 'set_cells');
  assert.equal(parsed.actions[0].formula, '=SUM(B2:D2)');
  assert.equal(parsed.actions[1].action, 'format_range');
  assert.equal(parsed.actions[1].bold, true);

  const sys = buildOfficeAgentSystemPrompt('Base prompt', 'excel');
  assert.ok(sys.includes('set_cells'));
  assert.ok(sys.includes('<office_action>'));
});

test('stripStreamingOfficeActionTags suppresses partial opening tags during stream', () => {
  assert.equal(stripStreamingOfficeActionTags('正在处理\n<office_ac'), '正在处理');
  assert.equal(stripStreamingOfficeActionTags('完成\n<office_action>{"action":"set_cells"'), '完成');
  assert.equal(stripStreamingOfficeActionTags('<office_action>{"action":'), '');
});

test('buildGatewayLaunchUrl / parseGatewayLaunchUrl round-trip loopback only', () => {
  const url = buildGatewayLaunchUrl('http://127.0.0.1:18080', 'wps');
  assert.ok(url.startsWith('officeai://activate/'));
  assert.equal(parseGatewayLaunchUrl(url), 'http://127.0.0.1:18080');

  assert.equal(parseGatewayLaunchUrl(buildGatewayLaunchUrl('http://localhost:19000')), 'http://localhost:19000');

  // Reject anything that is not a loopback http endpoint
  assert.equal(parseGatewayLaunchUrl('officeai://activate/https%3A%2F%2Fevil.example.com'), undefined);
  assert.equal(parseGatewayLaunchUrl('https://example.com'), undefined);
  assert.equal(parseGatewayLaunchUrl('officeai://activate/not-a-url'), undefined);
});

test('createDefaultConfig defaults to always launch (prevent blank panel after auto-exit)', () => {
  const cfg = createDefaultConfig();
  assert.equal(cfg.settings.launchMode, 'always');
  assert.equal(cfg.settings.autoStartGateway, true);
  assert.equal(cfg.settings.exitWithOffice, false);
  assert.equal(cfg.settings.idleExitMinutes, 0);
});
