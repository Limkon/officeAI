import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  createDefaultConfig,
  createDefaultProviders,
  DEFAULT_PROMPT_ACTIONS,
  maskApiKey,
  type AppConfig,
  type AppSettings,
  type MaskedProviderConfig,
  type PromptAction,
  type ProviderConfig,
  type UsageRecord
} from '@officeai/shared';

export function resolveDataDir(): string {
  if (process.env.OFFICEAI_DATA_DIR) {
    return path.resolve(process.env.OFFICEAI_DATA_DIR);
  }
  if (process.platform === 'win32') {
    const appData = process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming');
    return path.join(appData, 'OfficeAI');
  }
  return path.join(os.homedir(), '.config', 'officeai');
}

export interface PublicAppConfig {
  version: number;
  settings: AppSettings;
  providers: MaskedProviderConfig[];
  prompts: PromptAction[];
}

export class ConfigStore {
  readonly dataDir: string;
  readonly configPath: string;
  readonly usagePath: string;
  private config: AppConfig;
  private usage: UsageRecord[];

  constructor(dataDir = resolveDataDir()) {
    this.dataDir = dataDir;
    this.configPath = path.join(dataDir, 'config.json');
    this.usagePath = path.join(dataDir, 'usage.json');
    fs.mkdirSync(this.dataDir, { recursive: true });
    this.config = this.loadConfig();
    this.usage = this.loadUsage();
  }

  private loadConfig(): AppConfig {
    const defaults = createDefaultConfig();
    if (!fs.existsSync(this.configPath)) {
      this.writeJsonAtomic(this.configPath, defaults);
      return defaults;
    }

    try {
      const raw = fs.readFileSync(this.configPath, 'utf8');
      const parsed = JSON.parse(raw) as Partial<AppConfig>;
      const defaultProviders = createDefaultProviders();
      const savedProviders = Array.isArray(parsed.providers) ? parsed.providers : [];

      // Keep saved providers and append any newly introduced presets
      const existingIds = new Set(savedProviders.map((p) => p.id));
      const mergedProviders: ProviderConfig[] = [
        ...savedProviders,
        ...defaultProviders.filter((dp) => !existingIds.has(dp.id))
      ];

      const savedPrompts = Array.isArray(parsed.prompts) ? parsed.prompts : [];
      const savedPromptIds = new Set(savedPrompts.map((p) => p.id));
      const mergedPrompts: PromptAction[] = [
        ...savedPrompts,
        ...DEFAULT_PROMPT_ACTIONS.filter((p) => !savedPromptIds.has(p.id))
      ];

      return {
        version: parsed.version ?? defaults.version,
        settings: { ...defaults.settings, ...(parsed.settings ?? {}) },
        providers: mergedProviders,
        prompts: mergedPrompts
      };
    } catch {
      return defaults;
    }
  }

  private loadUsage(): UsageRecord[] {
    if (!fs.existsSync(this.usagePath)) {
      return [];
    }
    try {
      const raw = fs.readFileSync(this.usagePath, 'utf8');
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  private writeJsonAtomic(filePath: string, data: unknown): void {
    fs.mkdirSync(path.dirname(filePath), { recursive: true });
    const tmpPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
    fs.writeFileSync(tmpPath, JSON.stringify(data, null, 2), 'utf8');
    fs.renameSync(tmpPath, filePath);
  }

  getConfig(): AppConfig {
    return this.config;
  }

  getPublicConfig(): PublicAppConfig {
    return {
      version: this.config.version,
      settings: { ...this.config.settings },
      providers: this.config.providers.map((p) => ({
        id: p.id,
        presetId: p.presetId,
        name: p.name,
        category: p.category,
        protocol: p.protocol,
        baseUrl: p.baseUrl,
        models: [...p.models],
        selectedModel: p.selectedModel,
        enabled: p.enabled,
        extraHeaders: p.extraHeaders,
        apiKeyMasked: maskApiKey(p.apiKey),
        hasApiKey: Boolean(p.apiKey && p.apiKey.trim().length > 0)
      })),
      prompts: [...this.config.prompts]
    };
  }

  updateSettings(patch: Partial<AppSettings>): AppSettings {
    this.config.settings = {
      ...this.config.settings,
      ...patch
    };
    this.writeJsonAtomic(this.configPath, this.config);
    return this.config.settings;
  }

  getProvider(id: string): ProviderConfig | undefined {
    return this.config.providers.find((p) => p.id === id);
  }

  upsertProvider(
    provider: Partial<ProviderConfig> & { id: string }
  ): ProviderConfig {
    const idx = this.config.providers.findIndex((p) => p.id === provider.id);
    if (idx === -1) {
      const created: ProviderConfig = {
        id: provider.id,
        presetId: provider.presetId ?? 'custom-cloud',
        name: provider.name ?? provider.id,
        category: provider.category ?? 'cloud',
        protocol: provider.protocol ?? 'openai',
        baseUrl: provider.baseUrl ?? 'https://api.openai.com/v1',
        apiKey: provider.apiKey ?? '',
        models: provider.models ?? ['gpt-4o'],
        selectedModel: provider.selectedModel ?? provider.models?.[0] ?? 'gpt-4o',
        enabled: provider.enabled ?? true,
        extraHeaders: provider.extraHeaders
      };
      this.config.providers.push(created);
      this.writeJsonAtomic(this.configPath, this.config);
      return created;
    }

    const current = this.config.providers[idx];
    const updated: ProviderConfig = {
      ...current,
      ...provider,
      // Preserve existing API key if empty/undefined unless explicitly cleared via clearApiKey
      apiKey:
        provider.apiKey !== undefined ? provider.apiKey : current.apiKey
    };
    this.config.providers[idx] = updated;
    this.writeJsonAtomic(this.configPath, this.config);
    return updated;
  }

  deleteProvider(id: string): boolean {
    const before = this.config.providers.length;
    this.config.providers = this.config.providers.filter((p) => p.id !== id);
    if (this.config.providers.length !== before) {
      this.writeJsonAtomic(this.configPath, this.config);
      return true;
    }
    return false;
  }

  getPrompts(): PromptAction[] {
    return this.config.prompts;
  }

  upsertPrompt(prompt: PromptAction): PromptAction {
    const idx = this.config.prompts.findIndex((p) => p.id === prompt.id);
    if (idx === -1) {
      this.config.prompts.push(prompt);
    } else {
      this.config.prompts[idx] = prompt;
    }
    this.writeJsonAtomic(this.configPath, this.config);
    return prompt;
  }

  deletePrompt(id: string): boolean {
    const before = this.config.prompts.length;
    this.config.prompts = this.config.prompts.filter((p) => p.id !== id);
    if (this.config.prompts.length !== before) {
      this.writeJsonAtomic(this.configPath, this.config);
      return true;
    }
    return false;
  }

  recordUsage(record: UsageRecord): void {
    this.usage.unshift(record);
    if (this.usage.length > 500) {
      this.usage = this.usage.slice(0, 500);
    }
    this.writeJsonAtomic(this.usagePath, this.usage);
  }

  getUsage(): UsageRecord[] {
    return this.usage;
  }

  clearUsage(): void {
    this.usage = [];
    this.writeJsonAtomic(this.usagePath, this.usage);
  }
}
