import type {
  AppSettings,
  ChatRequest,
  ConnectionTestResult,
  LocalRuntimeStatus,
  MaskedProviderConfig,
  PromptAction,
  ProviderConfig,
  StreamEvent,
  UsageRecord,
  WindowsOfficeStatus
} from '@officeai/shared';

export interface PublicConfigResponse {
  version: number;
  settings: AppSettings;
  providers: MaskedProviderConfig[];
  prompts: PromptAction[];
}

export interface IntegrationResult {
  ok: boolean;
  message: string;
  details?: string[];
}

const API_BASE = '';

async function requestJson<T>(url: string, init?: RequestInit): Promise<T> {
  const res = await fetch(`${API_BASE}${url}`, {
    ...init,
    headers: {
      'Content-Type': 'application/json',
      ...(init?.headers ?? {})
    }
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error((data as { error?: string; message?: string }).error || `HTTP ${res.status}`);
  }
  return data as T;
}

export const apiClient = {
  getConfig(): Promise<PublicConfigResponse> {
    return requestJson<PublicConfigResponse>('/api/config');
  },

  updateSettings(patch: Partial<AppSettings>): Promise<{ ok: boolean; settings: AppSettings }> {
    return requestJson('/api/config/settings', {
      method: 'PUT',
      body: JSON.stringify(patch)
    });
  },

  saveProvider(
    id: string,
    patch: Partial<ProviderConfig> & { clearApiKey?: boolean }
  ): Promise<{ ok: boolean; config: PublicConfigResponse }> {
    return requestJson(`/api/providers/${encodeURIComponent(id)}`, {
      method: 'PUT',
      body: JSON.stringify(patch)
    });
  },

  createProvider(
    provider: Partial<ProviderConfig> & { id: string }
  ): Promise<{ ok: boolean; config: PublicConfigResponse }> {
    return requestJson('/api/providers', {
      method: 'POST',
      body: JSON.stringify(provider)
    });
  },

  deleteProvider(id: string): Promise<{ ok: boolean }> {
    return requestJson(`/api/providers/${encodeURIComponent(id)}`, {
      method: 'DELETE'
    });
  },

  testProvider(
    id: string,
    override?: Partial<ProviderConfig>
  ): Promise<ConnectionTestResult> {
    return requestJson(`/api/providers/${encodeURIComponent(id)}/test`, {
      method: 'POST',
      body: JSON.stringify(override ?? {})
    });
  },

  getLocalStatus(): Promise<{ runtimes: LocalRuntimeStatus[] }> {
    return requestJson('/api/local/status');
  },

  savePrompt(prompt: PromptAction): Promise<{ ok: boolean; prompt: PromptAction }> {
    return requestJson('/api/prompts', {
      method: 'POST',
      body: JSON.stringify(prompt)
    });
  },

  deletePrompt(id: string): Promise<{ ok: boolean }> {
    return requestJson(`/api/prompts/${encodeURIComponent(id)}`, {
      method: 'DELETE'
    });
  },

  getUsage(): Promise<{ records: UsageRecord[] }> {
    return requestJson('/api/usage');
  },

  clearUsage(): Promise<{ ok: boolean }> {
    return requestJson('/api/usage', { method: 'DELETE' });
  },

  /** Heartbeat: keeps the gateway alive while this task pane is open */
  pingSession(app: string, source = 'addin'): Promise<{ ok: boolean }> {
    return requestJson('/api/office/session', {
      method: 'POST',
      body: JSON.stringify({ app, source })
    });
  },

  getWindowsStatus(): Promise<WindowsOfficeStatus> {
    return requestJson('/api/windows/status');
  },

  installOffice(): Promise<IntegrationResult> {
    return requestJson('/api/windows/install-office', { method: 'POST' });
  },

  uninstallOffice(): Promise<IntegrationResult> {
    return requestJson('/api/windows/uninstall-office', { method: 'POST' });
  },

  installWps(): Promise<IntegrationResult> {
    return requestJson('/api/windows/install-wps', { method: 'POST' });
  },

  uninstallWps(): Promise<IntegrationResult> {
    return requestJson('/api/windows/uninstall-wps', { method: 'POST' });
  },

  trustCert(): Promise<IntegrationResult> {
    return requestJson('/api/windows/trust-cert', { method: 'POST' });
  },

  /** Registers officeai:// so Office users can start the gateway with one click */
  registerProtocol(): Promise<IntegrationResult> {
    return requestJson('/api/windows/register-protocol', { method: 'POST' });
  },

  /** Office 唤醒器：让 Word/Excel/PPT 启动时自动拉起网关 */
  installWaker(): Promise<IntegrationResult> {
    return requestJson('/api/windows/install-waker', { method: 'POST' });
  },

  uninstallWaker(): Promise<IntegrationResult> {
    return requestJson('/api/windows/uninstall-waker', { method: 'POST' });
  },

  unregisterProtocol(): Promise<IntegrationResult> {
    return requestJson('/api/windows/unregister-protocol', { method: 'POST' });
  },

  async streamChat(
    req: ChatRequest,
    onEvent: (event: StreamEvent) => void,
    signal?: AbortSignal
  ): Promise<void> {
    const res = await fetch(`${API_BASE}/api/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ...req, stream: true }),
      signal
    });

    if (!res.ok || !res.body) {
      const errBody = await res.json().catch(() => ({}));
      throw new Error(
        (errBody as { error?: string }).error || `请求失败 (HTTP ${res.status})`
      );
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder('utf-8');
    let buffer = '';

    try {
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        let idx = buffer.indexOf('\n');
        while (idx !== -1) {
          const line = buffer.slice(0, idx).replace(/\r$/, '');
          buffer = buffer.slice(idx + 1);
          if (line.startsWith('data:')) {
            const payload = line.slice(5).trim();
            if (payload) {
              try {
                const ev = JSON.parse(payload) as StreamEvent;
                onEvent(ev);
              } catch {
                // ignore malformed chunk
              }
            }
          }
          idx = buffer.indexOf('\n');
        }
      }
    } finally {
      reader.releaseLock();
    }
  }
};
