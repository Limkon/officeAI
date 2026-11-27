import { createDefaultConfig } from '@officeai/shared';
import type { AppConfig, ProviderConfig } from '@officeai/shared';

const STORAGE_KEY = 'officeai_direct_config_v1';

function load(): AppConfig {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as AppConfig;
      // 合并预置，防止新增 provider 丢失
      const defaults = createDefaultConfig();
      const merged = { ...defaults, ...parsed, settings: { ...defaults.settings, ...(parsed as any).settings } } as AppConfig;
      // 确保 providers 包含所有预置
      const ids = new Set(merged.providers.map(p => p.id));
      for (const p of defaults.providers) {
        if (!ids.has(p.id)) merged.providers.push(p);
      }
      return merged;
    }
  } catch {}
  return createDefaultConfig();
}

function save(cfg: AppConfig) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
  } catch {}
}

export const directConfigStore = {
  get(): AppConfig {
    return load();
  },
  getProvider(id: string): ProviderConfig | undefined {
    return load().providers.find(p => p.id === id);
  },
  upsertProvider(patch: Partial<ProviderConfig> & { id: string }) {
    const cfg = load();
    const idx = cfg.providers.findIndex(p => p.id === patch.id);
    if (idx >= 0) {
      cfg.providers[idx] = { ...cfg.providers[idx], ...patch } as ProviderConfig;
    } else {
      cfg.providers.push(patch as ProviderConfig);
    }
    save(cfg);
    return cfg;
  },
  updateSettings(patch: Partial<AppConfig['settings']>) {
    const cfg = load();
    cfg.settings = { ...cfg.settings, ...patch } as any;
    save(cfg);
    return cfg.settings;
  },
};
