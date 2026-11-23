import type { AppConfig, PromptAction, ProviderConfig, ProviderPreset } from './types.js';

export const PROVIDER_PRESETS: ProviderPreset[] = [
  {
    id: 'openai',
    name: 'OpenAI',
    nameEn: 'OpenAI',
    category: 'cloud',
    protocol: 'openai',
    defaultBaseUrl: 'https://api.openai.com/v1',
    defaultModels: ['gpt-4o', 'gpt-4o-mini', 'gpt-4.1', 'o3-mini'],
    requiresApiKey: true,
    docUrl: 'https://platform.openai.com/api-keys',
    description: 'GPT-4o / GPT-4.1 / o3 系列模型',
    badgeColor: '#10a37f'
  },
  {
    id: 'anthropic',
    name: 'Anthropic Claude',
    nameEn: 'Anthropic Claude',
    category: 'cloud',
    protocol: 'anthropic',
    defaultBaseUrl: 'https://api.anthropic.com',
    defaultModels: ['claude-sonnet-4-5', 'claude-3-7-sonnet-latest', 'claude-3-5-haiku-latest'],
    requiresApiKey: true,
    docUrl: 'https://console.anthropic.com/',
    description: 'Claude 4.5 / 3.7 Sonnet 长文写作与复杂推理',
    badgeColor: '#d97757'
  },
  {
    id: 'gemini',
    name: 'Google Gemini',
    nameEn: 'Google Gemini',
    category: 'cloud',
    protocol: 'gemini',
    defaultBaseUrl: 'https://generativelanguage.googleapis.com',
    defaultModels: ['gemini-2.5-pro', 'gemini-2.5-flash', 'gemini-2.0-flash'],
    requiresApiKey: true,
    docUrl: 'https://aistudio.google.com/apikey',
    description: 'Gemini 2.5 Pro / Flash 超长上下文',
    badgeColor: '#4285f4'
  },
  {
    id: 'deepseek',
    name: 'DeepSeek 深度求索',
    nameEn: 'DeepSeek',
    category: 'cloud',
    protocol: 'openai',
    defaultBaseUrl: 'https://api.deepseek.com/v1',
    defaultModels: ['deepseek-chat', 'deepseek-reasoner'],
    requiresApiKey: true,
    docUrl: 'https://platform.deepseek.com/',
    description: 'DeepSeek-V3 & R1 推理模型（高性价比中文优化）',
    badgeColor: '#4d6bfe'
  },
  {
    id: 'qwen',
    name: '通义千问 Qwen (阿里云百炼)',
    nameEn: 'Alibaba Qwen',
    category: 'cloud',
    protocol: 'openai',
    defaultBaseUrl: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
    defaultModels: ['qwen-max', 'qwen-plus', 'qwen-turbo', 'qwq-plus'],
    requiresApiKey: true,
    docUrl: 'https://bailian.console.aliyun.com/',
    description: '通义千问 Max / Plus / QwQ 推理模型',
    badgeColor: '#615ced'
  },
  {
    id: 'zhipu',
    name: '智谱 GLM',
    nameEn: 'Zhipu GLM',
    category: 'cloud',
    protocol: 'openai',
    defaultBaseUrl: 'https://open.bigmodel.cn/api/paas/v4',
    defaultModels: ['glm-4-plus', 'glm-4-air', 'glm-4-flash', 'glm-z1-flash'],
    requiresApiKey: true,
    docUrl: 'https://open.bigmodel.cn/',
    description: '智谱 GLM-4 系列（含免费 Flash 模型）',
    badgeColor: '#2563eb'
  },
  {
    id: 'moonshot',
    name: 'Moonshot Kimi',
    nameEn: 'Moonshot Kimi',
    category: 'cloud',
    protocol: 'openai',
    defaultBaseUrl: 'https://api.moonshot.cn/v1',
    defaultModels: ['kimi-k2-0711-preview', 'moonshot-v1-8k', 'moonshot-v1-32k', 'moonshot-v1-128k'],
    requiresApiKey: true,
    docUrl: 'https://platform.moonshot.cn/',
    description: 'Kimi 长文档总结与办公助手',
    badgeColor: '#111827'
  },
  {
    id: 'doubao',
    name: '火山方舟 Doubao (豆包)',
    nameEn: 'Volcengine Doubao',
    category: 'cloud',
    protocol: 'openai',
    defaultBaseUrl: 'https://ark.cn-beijing.volces.com/api/v3',
    defaultModels: ['doubao-1-5-pro-32k-250115', 'doubao-1-5-lite-32k-250115'],
    requiresApiKey: true,
    docUrl: 'https://console.volcengine.com/ark',
    description: '字节跳动豆包大模型（支持自定义接入点 ID）',
    badgeColor: '#0ea5e9'
  },
  {
    id: 'siliconflow',
    name: '硅基流动 SiliconFlow',
    nameEn: 'SiliconFlow',
    category: 'cloud',
    protocol: 'openai',
    defaultBaseUrl: 'https://api.siliconflow.cn/v1',
    defaultModels: ['Qwen/Qwen2.5-72B-Instruct', 'deepseek-ai/DeepSeek-V3', 'deepseek-ai/DeepSeek-R1'],
    requiresApiKey: true,
    docUrl: 'https://cloud.siliconflow.cn/',
    description: '一站式聚合国内主流开源大模型',
    badgeColor: '#7c3aed'
  },
  {
    id: 'openrouter',
    name: 'OpenRouter',
    nameEn: 'OpenRouter',
    category: 'cloud',
    protocol: 'openai',
    defaultBaseUrl: 'https://openrouter.ai/api/v1',
    defaultModels: ['openai/gpt-4o', 'anthropic/claude-3.7-sonnet', 'google/gemini-2.5-pro'],
    requiresApiKey: true,
    docUrl: 'https://openrouter.ai/',
    description: '统一接入全球主流闭源与开源模型',
    badgeColor: '#6366f1'
  },
  {
    id: 'groq',
    name: 'Groq',
    nameEn: 'Groq',
    category: 'cloud',
    protocol: 'openai',
    defaultBaseUrl: 'https://api.groq.com/openai/v1',
    defaultModels: ['llama-3.3-70b-versatile', 'qwen-qwq-32b'],
    requiresApiKey: true,
    docUrl: 'https://console.groq.com/',
    description: '超高速推理芯片云服务',
    badgeColor: '#f97316'
  },
  {
    id: 'ollama',
    name: 'Ollama 本地模型',
    nameEn: 'Ollama (Local)',
    category: 'local',
    protocol: 'ollama',
    defaultBaseUrl: 'http://127.0.0.1:11434',
    defaultModels: ['qwen2.5:7b', 'deepseek-r1:7b', 'llama3.2:latest'],
    requiresApiKey: false,
    docUrl: 'https://ollama.com/',
    description: '本机离线运行 Qwen / DeepSeek / Llama 等模型，数据不出本地',
    badgeColor: '#059669'
  },
  {
    id: 'lmstudio',
    name: 'LM Studio 本地模型',
    nameEn: 'LM Studio (Local)',
    category: 'local',
    protocol: 'openai',
    defaultBaseUrl: 'http://127.0.0.1:1234/v1',
    defaultModels: ['local-model'],
    requiresApiKey: false,
    docUrl: 'https://lmstudio.ai/',
    description: '通过 LM Studio Local Server 一键接入 GGUF 本地模型',
    badgeColor: '#16a34a'
  },
  {
    id: 'vllm',
    name: 'vLLM / llama.cpp / 自定义本地服务',
    nameEn: 'Local OpenAI-Compatible',
    category: 'local',
    protocol: 'openai',
    defaultBaseUrl: 'http://127.0.0.1:8000/v1',
    defaultModels: ['default'],
    requiresApiKey: false,
    description: '局域网/本机自建 vLLM、llama-server、LocalAI 服务',
    badgeColor: '#0d9488'
  },
  {
    id: 'custom-cloud',
    name: '自定义 OpenAI 兼容 API',
    nameEn: 'Custom OpenAI-Compatible',
    category: 'cloud',
    protocol: 'openai',
    defaultBaseUrl: 'https://api.example.com/v1',
    defaultModels: ['gpt-4o'],
    requiresApiKey: true,
    description: '接入企业私有网关、OneAPI / NewAPI 等兼容接口',
    badgeColor: '#64748b'
  }
];

export const DEFAULT_PROMPT_ACTIONS: PromptAction[] = [
  {
    id: 'polish',
    title: '润色优化',
    titleEn: 'Polish & Refine',
    category: 'write',
    apps: ['word', 'wps-writer', 'powerpoint', 'wps-presentation', 'standalone'],
    systemPrompt: '你是资深中文与英文文字编辑。请在不改变原意的前提下优化语序、用词与可读性，直接输出润色后的正文，不要添加多余解释。',
    userPromptTemplate: '请润色以下内容：\n\n{{selection}}',
    applyMode: 'replace',
    builtIn: true
  },
  {
    id: 'academic',
    title: '学术/正式改写',
    titleEn: 'Formal / Academic',
    category: 'write',
    apps: ['word', 'wps-writer', 'standalone'],
    systemPrompt: '你是严谨的学术论文与公文编辑。请将文本改写为更正式、客观、结构清晰的书面表达，直接输出改写后的正文。',
    userPromptTemplate: '请将以下文本改写为正式/学术风格：\n\n{{selection}}',
    applyMode: 'replace',
    builtIn: true
  },
  {
    id: 'summarize',
    title: '提炼摘要',
    titleEn: 'Summarize',
    category: 'analyze',
    apps: 'all',
    systemPrompt: '你是专业办公助手。请提炼核心要点与关键结论，用清晰的条目列出。',
    userPromptTemplate: '请总结以下内容的核心要点：\n\n{{selection}}',
    applyMode: 'panel',
    builtIn: true
  },
  {
    id: 'expand',
    title: '扩写充实',
    titleEn: 'Expand',
    category: 'write',
    apps: ['word', 'wps-writer', 'standalone'],
    systemPrompt: '你是内容创作专家。请保持原文逻辑与语气，补充合理细节、论据与过渡句，使内容更加充实完整，直接输出正文。',
    userPromptTemplate: '请扩写以下段落：\n\n{{selection}}',
    applyMode: 'replace',
    builtIn: true
  },
  {
    id: 'translate-en',
    title: '翻译为英文',
    titleEn: 'Translate to English',
    category: 'translate',
    apps: 'all',
    systemPrompt: 'You are a professional translator. Translate the user text into natural, idiomatic English while preserving formatting and terminology. Output only the translation.',
    userPromptTemplate: 'Translate to English:\n\n{{selection}}',
    applyMode: 'insert',
    builtIn: true
  },
  {
    id: 'translate-zh',
    title: '翻译为中文',
    titleEn: 'Translate to Chinese',
    category: 'translate',
    apps: 'all',
    systemPrompt: '你是专业翻译专家。请将用户输入准确流畅地翻译为简体中文，保留专有名词与格式，仅输出译文。',
    userPromptTemplate: '请将以下内容翻译为简体中文：\n\n{{selection}}',
    applyMode: 'insert',
    builtIn: true
  },
  {
    id: 'proofread',
    title: '纠错校对',
    titleEn: 'Proofread',
    category: 'write',
    apps: ['word', 'wps-writer', 'powerpoint', 'wps-presentation', 'standalone'],
    systemPrompt: '你是专业校对编辑。请修正错别字、语法错误、标点问题和明显语病，直接输出修正后的文本。',
    userPromptTemplate: '请校对并修正以下文本：\n\n{{selection}}',
    applyMode: 'replace',
    builtIn: true
  },
  {
    id: 'excel-formula',
    title: '生成 Excel/WPS 公式',
    titleEn: 'Generate Formula',
    category: 'excel',
    apps: ['excel', 'wps-spreadsheets', 'standalone'],
    systemPrompt:
      '你是精通 Microsoft Excel 与 WPS 表格的函数专家。根据用户描述和当前选区数据，给出最简洁准确的公式（以 = 开头）。请优先用代码块仅输出可直接复制使用的公式，并在下一行简要说明用法。',
    userPromptTemplate: '当前选区/表头参考：\n{{context}}\n\n需求：{{selection}}',
    applyMode: 'panel',
    builtIn: true
  },
  {
    id: 'excel-analyze',
    title: '数据洞察分析',
    titleEn: 'Analyze Table Data',
    category: 'excel',
    apps: ['excel', 'wps-spreadsheets', 'standalone'],
    systemPrompt:
      '你是资深数据分析师。请分析提供的表格数据，指出关键趋势、异常值、汇总指标及可执行的业务建议，结构清晰。',
    userPromptTemplate: '请分析以下表格数据：\n\n{{selection}}',
    applyMode: 'panel',
    builtIn: true
  },
  {
    id: 'ppt-outline',
    title: '生成 PPT 大纲',
    titleEn: 'Generate Slide Outline',
    category: 'ppt',
    apps: ['powerpoint', 'wps-presentation', 'word', 'wps-writer', 'standalone'],
    systemPrompt:
      '你是演示文稿架构师。请根据主题生成结构清晰的 PPT 逐页大纲（每页包含：页面标题、3-5条核心要点、建议配图/图表）。',
    userPromptTemplate: '请根据以下主题或内容生成 PPT 大纲：\n\n{{selection}}',
    applyMode: 'panel',
    builtIn: true
  },
  {
    id: 'ppt-notes',
    title: '生成演讲备注',
    titleEn: 'Speaker Notes',
    category: 'ppt',
    apps: ['powerpoint', 'wps-presentation', 'standalone'],
    systemPrompt:
      '你是演讲教练。请根据当前幻灯片内容编写自然流畅、有感染力的演讲者备注口播稿，约150-250字。',
    userPromptTemplate: '请为以下幻灯片内容生成演讲备注：\n\n{{selection}}',
    applyMode: 'insert',
    builtIn: true
  }
];

export function createDefaultProviders(): ProviderConfig[] {
  return PROVIDER_PRESETS.map((preset) => ({
    id: preset.id,
    presetId: preset.id,
    name: preset.name,
    category: preset.category,
    protocol: preset.protocol,
    baseUrl: preset.defaultBaseUrl,
    apiKey: '',
    models: [...preset.defaultModels],
    selectedModel: preset.defaultModels[0] ?? '',
    enabled: preset.id === 'ollama' || preset.id === 'deepseek' || preset.id === 'openai'
  }));
}

export function createDefaultConfig(): AppConfig {
  return {
    version: 1,
    settings: {
      httpPort: 18080,
      httpsPort: 18443,
      activeProviderId: 'ollama',
      activeModel: 'qwen2.5:7b',
      temperature: 0.7,
      maxTokens: 4096,
      systemPrompt:
        '你是 OfficeAI 智能办公助手，擅长在 Word、Excel、PowerPoint 和 WPS 中进行文本写作、润色、翻译、表格公式生成与数据分析。回答请准确、高效、结构清晰。',
      language: 'zh-CN',
      autoStartGateway: true,
      launchMode: 'always',
      gatewayLaunchCommand: '',
      exitWithOffice: false,
      idleExitMinutes: 0,
      sessionGraceMinutes: 2
    },
    providers: createDefaultProviders(),
    prompts: [...DEFAULT_PROMPT_ACTIONS]
  };
}
