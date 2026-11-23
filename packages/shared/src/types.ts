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

/**
 * 网关启动策略：
 * - manual    仅手动启动（桌面客户端 / Start-OfficeAI.bat）
 * - on-demand 由 Office / WPS 加载项在用到时自动拉起（推荐）
 * - always    桌面客户端启动即常驻（不做开机自启）
 */
export type GatewayLaunchMode = 'manual' | 'on-demand' | 'always';

export interface AppSettings {
  httpPort: number;
  httpsPort: number;
  activeProviderId: string;
  activeModel: string;
  temperature: number;
  maxTokens: number;
  systemPrompt: string;
  language: 'zh-CN' | 'en-US';
  /** 桌面客户端启动时是否立即拉起网关（= launchMode 为 always） */
  autoStartGateway: boolean;
  launchMode: GatewayLaunchMode;
  /** 自定义启动命令；为空时由加载项按安装位置自动探测 */
  gatewayLaunchCommand: string;
  /** Word/Excel/PPT/WPS 全部关闭后自动退出网关（仅 on-demand 生效） */
  exitWithOffice: boolean;
  /** 最后一段时间无请求后自动退出；0 表示不按空闲退出 */
  idleExitMinutes: number;
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

export interface LaunchTargetInfo {
  command: string;
  kind: 'bat' | 'exe' | 'command' | 'none';
  source: string;
}

export interface WindowsOfficeStatus {
  platform: string;
  isWindows: boolean;
  gatewayHttpUrl: string;
  gatewayHttpsUrl: string;
  certExists: boolean;
  certTrusted: boolean;
  /** How OfficeAI is started on demand, and whether the one-click link is ready */
  launch?: {
    target: LaunchTargetInfo;
    protocolRegistered: boolean;
    startPageUrl: string;
    launchUrl: string;
    officeRunning: boolean;
    runningProcesses: string[];
  };
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
    /** Local (offline) add-in folder, used for on-demand launch */
    localAddinPath?: string;
    mode?: 'local' | 'online' | 'none';
  };
}
