export type ProviderProtocol = 'openai' | 'anthropic' | 'gemini' | 'ollama';
export type ProviderCategory = 'cloud' | 'local';

export type OfficeHostSuite = 'ms-office' | 'wps' | 'standalone';
export type OfficeHostApp =
  | 'word'
  | 'excel'
  | 'powerpoint'
  | 'wps-writer'
  | 'wps-spreadsheets'
  | 'wps-presentation'
  | 'standalone';

export interface ProviderPreset {
  id: string;
  name: string;
  nameEn: string;
  category: ProviderCategory;
  protocol: ProviderProtocol;
  defaultBaseUrl: string;
  defaultModels: string[];
  requiresApiKey: boolean;
  docUrl?: string;
  description: string;
  badgeColor?: string;
}

export interface ProviderConfig {
  id: string;
  presetId: string;
  name: string;
  category: ProviderCategory;
  protocol: ProviderProtocol;
  baseUrl: string;
  apiKey: string;
  models: string[];
  selectedModel: string;
  enabled: boolean;
  extraHeaders?: Record<string, string>;
}

export interface MaskedProviderConfig extends Omit<ProviderConfig, 'apiKey'> {
  apiKeyMasked: string;
  hasApiKey: boolean;
}

export type PromptCategory = 'write' | 'translate' | 'excel' | 'ppt' | 'analyze';
export type PromptApplyMode = 'replace' | 'insert' | 'panel';

export interface PromptAction {
  id: string;
  title: string;
  titleEn: string;
  category: PromptCategory;
  apps: OfficeHostApp[] | 'all';
  systemPrompt: string;
  userPromptTemplate: string;
  applyMode: PromptApplyMode;
  builtIn?: boolean;
}

export interface AppSettings {
  httpPort: number;
  httpsPort: number;
  activeProviderId: string;
  activeModel: string;
  temperature: number;
  maxTokens: number;
  systemPrompt: string;
  language: 'zh-CN' | 'en-US';
  autoStartGateway: boolean;
}

export interface UsageRecord {
  id: string;
  timestamp: string;
  providerId: string;
  providerName: string;
  model: string;
  hostApp: OfficeHostApp;
  promptTokens?: number;
  completionTokens?: number;
  durationMs: number;
  success: boolean;
  error?: string;
}

export interface AppConfig {
  version: number;
  settings: AppSettings;
  providers: ProviderConfig[];
  prompts: PromptAction[];
}

export interface ChatMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

export interface ChatRequest {
  providerId?: string;
  model?: string;
  messages: ChatMessage[];
  temperature?: number;
  maxTokens?: number;
  stream?: boolean;
  hostApp?: OfficeHostApp;
}

export interface StreamEvent {
  type: 'delta' | 'reasoning' | 'done' | 'error' | 'meta';
  content?: string;
  error?: string;
  usage?: {
    promptTokens?: number;
    completionTokens?: number;
  };
  model?: string;
  providerId?: string;
}

export interface ConnectionTestResult {
  ok: boolean;
  latencyMs: number;
  models: string[];
  message: string;
}

export interface LocalRuntimeStatus {
  id: 'ollama' | 'lmstudio' | 'vllm';
  name: string;
  baseUrl: string;
  running: boolean;
  models: string[];
  error?: string;
}

export interface WindowsOfficeStatus {
  platform: string;
  isWindows: boolean;
  gatewayHttpUrl: string;
  gatewayHttpsUrl: string;
  certExists: boolean;
  certTrusted: boolean;
  msOffice: {
    detected: boolean;
    version?: string;
    installed: boolean;
    manifestPath?: string;
    registryConfigured?: boolean;
  };
  wps: {
    detected: boolean;
    addonsDir?: string;
    installed: boolean;
    jspluginsXmlPath?: string;
  };
}
