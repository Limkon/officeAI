import React, { useEffect, useState } from 'react';
import {
  Activity,
  ArrowLeft,
  CheckCircle2,
  Cloud,
  Cpu,
  Download,
  Eye,
  EyeOff,
  Loader2,
  MonitorCheck,
  Plus,
  RefreshCw,
  ShieldCheck,
  Sparkles,
  Trash2,
  Wand2,
  XCircle,
  Zap
} from 'lucide-react';
import {
  PROVIDER_PRESETS,
  type LocalRuntimeStatus,
  type MaskedProviderConfig,
  type PromptAction,
  type UsageRecord,
  type WindowsOfficeStatus
} from '@officeai/shared';
import { apiClient, type PublicConfigResponse } from '../api/client';

interface DashboardViewProps {
  config: PublicConfigResponse;
  onRefreshConfig: () => Promise<void>;
  onOpenTaskPane: () => void;
}

type TabKey = 'providers' | 'integration' | 'prompts' | 'usage';

export const DashboardView: React.FC<DashboardViewProps> = ({
  config,
  onRefreshConfig,
  onOpenTaskPane
}) => {
  const [tab, setTab] = useState<TabKey>('providers');
  const [selectedId, setSelectedId] = useState<string>(config.settings.activeProviderId || 'ollama');
  const [localRuntimes, setLocalRuntimes] = useState<LocalRuntimeStatus[]>([]);
  const [scanningLocal, setScanningLocal] = useState(false);

  // Provider edit form state
  const selectedProvider: MaskedProviderConfig | undefined =
    config.providers.find((p) => p.id === selectedId) || config.providers[0];

  const [formBaseUrl, setFormBaseUrl] = useState('');
  const [formApiKey, setFormApiKey] = useState('');
  const [showKey, setShowKey] = useState(false);
  const [formModel, setFormModel] = useState('');
  const [formCustomModels, setFormCustomModels] = useState('');
  const [formEnabled, setFormEnabled] = useState(true);
  const [saving, setSaving] = useState(false);
  const [testing, setTesting] = useState(false);
  const [testMsg, setTestMsg] = useState<{ ok: boolean; text: string } | null>(null);

  // Add custom provider modal/inline state
  const [addingProvider, setAddingProvider] = useState(false);
  const [newProvName, setNewProvName] = useState('');
  const [newProvCategory, setNewProvCategory] = useState<'cloud' | 'local'>('cloud');
  const [newProvProtocol, setNewProvProtocol] = useState<'openai' | 'anthropic' | 'gemini' | 'ollama'>('openai');
  const [newProvBaseUrl, setNewProvBaseUrl] = useState('https://api.example.com/v1');
  const [newProvModel, setNewProvModel] = useState('gpt-4o');

  // Global generation params
  const [globalTemp, setGlobalTemp] = useState(config.settings.temperature);
  const [globalMaxTokens, setGlobalMaxTokens] = useState(config.settings.maxTokens);
  const [globalSystemPrompt, setGlobalSystemPrompt] = useState(config.settings.systemPrompt);
  const [savingGlobal, setSavingGlobal] = useState(false);
  const [globalMsg, setGlobalMsg] = useState<string | null>(null);

  // Windows integration state
  const [winStatus, setWinStatus] = useState<WindowsOfficeStatus | null>(null);
  const [winActionMsg, setWinActionMsg] = useState<string | null>(null);
  const [winBusy, setWinBusy] = useState(false);

  // On-demand launch settings
  const [launchMode, setLaunchMode] = useState<'manual' | 'on-demand' | 'always'>(
    config.settings.launchMode ?? 'on-demand'
  );
  const [gatewayLaunchCommand, setGatewayLaunchCommand] = useState(
    config.settings.gatewayLaunchCommand ?? ''
  );
  const [exitWithOffice, setExitWithOffice] = useState(config.settings.exitWithOffice !== false);
  const [idleExitMinutes, setIdleExitMinutes] = useState(config.settings.idleExitMinutes ?? 10);
  const [sessionGraceMinutes, setSessionGraceMinutes] = useState(
    config.settings.sessionGraceMinutes ?? 2
  );
  const [launchTrigger, setLaunchTrigger] = useState<'watcher' | 'waker'>(
    config.settings.launchTrigger ?? 'watcher'
  );
  const [savingLaunch, setSavingLaunch] = useState(false);
  const [launchMsg, setLaunchMsg] = useState<string | null>(null);

  // Usage state
  const [usage, setUsage] = useState<UsageRecord[]>([]);

  // New prompt form
  const [newPromptTitle, setNewPromptTitle] = useState('');
  const [newPromptSystem, setNewPromptSystem] = useState('');
  const [newPromptTemplate, setNewPromptTemplate] = useState('请处理以下内容：\n\n{{selection}}');

  useEffect(() => {
    if (selectedProvider) {
      setFormBaseUrl(selectedProvider.baseUrl);
      setFormApiKey('');
      setFormModel(selectedProvider.selectedModel || selectedProvider.models[0] || '');
      setFormCustomModels(selectedProvider.models.join(', '));
      setFormEnabled(selectedProvider.enabled);
      setTestMsg(null);
    }
  }, [selectedProvider?.id, config.providers]);

  const scanLocal = async () => {
    setScanningLocal(true);
    try {
      const res = await apiClient.getLocalStatus();
      setLocalRuntimes(res.runtimes);
    } finally {
      setScanningLocal(false);
    }
  };

  const loadWinStatus = async () => {
    const status = await apiClient.getWindowsStatus();
    setWinStatus(status);
  };

  const loadUsage = async () => {
    const res = await apiClient.getUsage();
    setUsage(res.records);
  };

  useEffect(() => {
    void scanLocal();
    void loadWinStatus();
    void loadUsage();
  }, []);

  const handleSaveProvider = async (setAsActive = false) => {
    if (!selectedProvider) return;
    setSaving(true);
    try {
      const models = formCustomModels
        .split(/[,，\n]/)
        .map((s) => s.trim())
        .filter(Boolean);
      const chosenModel = formModel.trim() || models[0] || '';
      if (chosenModel && !models.includes(chosenModel)) {
        models.unshift(chosenModel);
      }

      await apiClient.saveProvider(selectedProvider.id, {
        baseUrl: formBaseUrl.trim(),
        ...(formApiKey.trim() ? { apiKey: formApiKey.trim() } : {}),
        models,
        selectedModel: chosenModel,
        enabled: setAsActive ? true : formEnabled
      });

      if (setAsActive) {
        await apiClient.updateSettings({
          activeProviderId: selectedProvider.id,
          activeModel: chosenModel
        });
      }

      await onRefreshConfig();
      setTestMsg({
        ok: true,
        text: setAsActive ? '已保存并设为当前默认模型' : '配置已保存'
      });
    } catch (err) {
      setTestMsg({ ok: false, text: err instanceof Error ? err.message : String(err) });
    } finally {
      setSaving(false);
    }
  };

  const handleTestProvider = async () => {
    if (!selectedProvider) return;
    setTesting(true);
    setTestMsg(null);
    try {
      const res = await apiClient.testProvider(selectedProvider.id, {
        baseUrl: formBaseUrl.trim(),
        ...(formApiKey.trim() ? { apiKey: formApiKey.trim() } : {}),
        selectedModel: formModel.trim()
      });
      setTestMsg({ ok: res.ok, text: res.message });
      if (res.ok && res.models.length > 0) {
        setFormCustomModels(res.models.join(', '));
        if (!res.models.includes(formModel)) {
          setFormModel(res.models[0]);
        }
        await onRefreshConfig();
      }
    } catch (err) {
      setTestMsg({ ok: false, text: err instanceof Error ? err.message : String(err) });
    } finally {
      setTesting(false);
    }
  };

  const handleWinAction = async (fn: () => Promise<{ ok: boolean; message: string }>) => {
    setWinBusy(true);
    setWinActionMsg(null);
    try {
      const res = await fn();
      setWinActionMsg(res.message);
      await loadWinStatus();
    } catch (err) {
      setWinActionMsg(err instanceof Error ? err.message : String(err));
    } finally {
      setWinBusy(false);
    }
  };

  useEffect(() => {
    setLaunchMode(config.settings.launchMode ?? 'on-demand');
    setGatewayLaunchCommand(config.settings.gatewayLaunchCommand ?? '');
    setExitWithOffice(config.settings.exitWithOffice !== false);
    setIdleExitMinutes(config.settings.idleExitMinutes ?? 10);
    setSessionGraceMinutes(config.settings.sessionGraceMinutes ?? 2);
    setLaunchTrigger(config.settings.launchTrigger ?? 'watcher');
  }, [
    config.settings.launchMode,
    config.settings.gatewayLaunchCommand,
    config.settings.exitWithOffice,
    config.settings.idleExitMinutes,
    config.settings.sessionGraceMinutes,
    config.settings.launchTrigger
  ]);

  const handleSaveLaunchSettings = async () => {
    setSavingLaunch(true);
    setLaunchMsg(null);
    try {
      await apiClient.updateSettings({
        launchMode,
        autoStartGateway: launchMode === 'always',
        gatewayLaunchCommand: gatewayLaunchCommand.trim(),
        exitWithOffice,
        idleExitMinutes,
        sessionGraceMinutes,
        launchTrigger
      });
      await onRefreshConfig();
      setLaunchMsg('启动方式已保存');
    } catch (err) {
      setLaunchMsg(err instanceof Error ? err.message : String(err));
    } finally {
      setSavingLaunch(false);
    }
  };

  useEffect(() => {
    setGlobalTemp(config.settings.temperature);
    setGlobalMaxTokens(config.settings.maxTokens);
    setGlobalSystemPrompt(config.settings.systemPrompt);
  }, [config.settings]);

  const handleClearApiKey = async () => {
    if (!selectedProvider) return;
    setSaving(true);
    try {
      await apiClient.saveProvider(selectedProvider.id, { clearApiKey: true });
      setFormApiKey('');
      await onRefreshConfig();
      setTestMsg({ ok: true, text: '已清除保存的 API Key' });
    } catch (err) {
      setTestMsg({ ok: false, text: err instanceof Error ? err.message : String(err) });
    } finally {
      setSaving(false);
    }
  };

  const handleCreateCustomProvider = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newProvName.trim() || !newProvBaseUrl.trim()) return;
    const id = `custom-${Date.now()}`;
    await apiClient.createProvider({
      id,
      presetId: newProvCategory === 'local' ? 'vllm' : 'custom-cloud',
      name: newProvName.trim(),
      category: newProvCategory,
      protocol: newProvProtocol,
      baseUrl: newProvBaseUrl.trim(),
      models: [newProvModel.trim() || 'default'],
      selectedModel: newProvModel.trim() || 'default',
      enabled: true
    });
    setAddingProvider(false);
    setNewProvName('');
    setSelectedId(id);
    await onRefreshConfig();
  };

  const handleSaveGlobalSettings = async (e: React.FormEvent) => {
    e.preventDefault();
    setSavingGlobal(true);
    setGlobalMsg(null);
    try {
      await apiClient.updateSettings({
        temperature: Number(globalTemp),
        maxTokens: Number(globalMaxTokens),
        systemPrompt: globalSystemPrompt.trim()
      });
      await onRefreshConfig();
      setGlobalMsg('全局参数已保存');
    } catch (err) {
      setGlobalMsg(err instanceof Error ? err.message : String(err));
    } finally {
      setSavingGlobal(false);
    }
  };

  const handleAddPrompt = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newPromptTitle.trim() || !newPromptSystem.trim()) return;
    const prompt: PromptAction = {
      id: `custom-${Date.now()}`,
      title: newPromptTitle.trim(),
      titleEn: newPromptTitle.trim(),
      category: 'write',
      apps: 'all',
      systemPrompt: newPromptSystem.trim(),
      userPromptTemplate: newPromptTemplate.trim() || '{{selection}}',
      applyMode: 'panel',
      builtIn: false
    };
    await apiClient.savePrompt(prompt);
    setNewPromptTitle('');
    setNewPromptSystem('');
    await onRefreshConfig();
  };

  return (
    <div className="min-h-full bg-slate-50 text-slate-900">
      {/* Top Bar */}
      <header className="border-b border-slate-200 bg-white">
        <div className="mx-auto flex max-w-6xl items-center justify-between px-6 py-4">
          <div className="flex items-center gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-gradient-to-br from-blue-600 to-indigo-600 text-white shadow-sm">
              <Sparkles className="h-5 w-5" />
            </div>
            <div>
              <h1 className="text-lg font-bold text-slate-900">OfficeAI 配置与集成中心</h1>
              <p className="text-xs text-slate-500">
                统一接入云端大模型 API 与本地模型（Ollama / LM Studio），支持 Microsoft Office & WPS Office
              </p>
            </div>
          </div>

          <button
            onClick={onOpenTaskPane}
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3.5 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50"
          >
            <ArrowLeft className="h-4 w-4" />
            <span>返回 AI 侧边栏</span>
          </button>
        </div>

        {/* Tabs */}
        <div className="mx-auto flex max-w-6xl gap-1 px-6">
          {[
            { key: 'providers', label: '模型接入配置', icon: Cloud },
            { key: 'integration', label: 'Office / WPS 安装集成', icon: MonitorCheck },
            { key: 'prompts', label: '快捷指令库', icon: Wand2 },
            { key: 'usage', label: '调用日志', icon: Activity }
          ].map((item) => {
            const Icon = item.icon;
            const active = tab === item.key;
            return (
              <button
                key={item.key}
                onClick={() => setTab(item.key as TabKey)}
                className={`inline-flex items-center gap-2 border-b-2 px-4 py-2.5 text-sm font-medium transition ${
                  active
                    ? 'border-blue-600 text-blue-600'
                    : 'border-transparent text-slate-600 hover:text-slate-900'
                }`}
              >
                <Icon className="h-4 w-4" />
                <span>{item.label}</span>
              </button>
            );
          })}
        </div>
      </header>

      <main className="mx-auto max-w-6xl px-6 py-6">
        {/* TAB 1: PROVIDERS */}
        {tab === 'providers' && (
          <div className="space-y-6">
            {/* Local Runtimes Scanner */}
            <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-2xs">
              <div className="flex items-center justify-between mb-4">
                <div className="flex items-center gap-2">
                  <Cpu className="h-5 w-5 text-emerald-600" />
                  <h2 className="text-base font-semibold">本机离线模型引擎检测</h2>
                </div>
                <button
                  onClick={() => void scanLocal()}
                  disabled={scanningLocal}
                  className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50"
                >
                  <RefreshCw className={`h-3.5 w-3.5 ${scanningLocal ? 'animate-spin' : ''}`} />
                  <span>重新检测</span>
                </button>
              </div>

              <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
                {localRuntimes.map((rt) => (
                  <div
                    key={rt.id}
                    onClick={() => setSelectedId(rt.id)}
                    className={`cursor-pointer rounded-xl border p-3.5 transition ${
                      rt.running
                        ? 'border-emerald-200 bg-emerald-50/40 hover:border-emerald-300'
                        : 'border-slate-200 bg-slate-50/60 hover:border-slate-300'
                    }`}
                  >
                    <div className="flex items-center justify-between">
                      <span className="font-semibold text-sm text-slate-900">{rt.name}</span>
                      <span
                        className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-[11px] font-medium ${
                          rt.running
                            ? 'bg-emerald-100 text-emerald-700'
                            : 'bg-slate-200 text-slate-600'
                        }`}
                      >
                        <span
                          className={`h-1.5 w-1.5 rounded-full ${
                            rt.running ? 'bg-emerald-500' : 'bg-slate-400'
                          }`}
                        />
                        {rt.running ? '运行中' : '未启动'}
                      </span>
                    </div>
                    <div className="mt-1 text-xs text-slate-500">{rt.baseUrl}</div>
                    <div className="mt-2 text-xs text-slate-700">
                      {rt.running
                        ? rt.models.length > 0
                          ? `已发现模型: ${rt.models.slice(0, 3).join(', ')}${rt.models.length > 3 ? ` 等${rt.models.length}个` : ''}`
                          : '服务已连接（暂无模型）'
                        : '点击配置地址或启动本地引擎后重新检测'}
                    </div>
                  </div>
                ))}
              </div>
            </section>

            {/* Provider List + Editor */}
            <div className="grid grid-cols-1 gap-6 lg:grid-cols-12">
              <aside className="lg:col-span-4 space-y-3">
                <div className="rounded-2xl border border-slate-200 bg-white p-3 shadow-2xs space-y-1 max-h-[540px] overflow-y-auto">
                  <div className="flex items-center justify-between px-2 py-1">
                    <span className="text-xs font-semibold uppercase tracking-wider text-slate-400">
                      模型提供方列表
                    </span>
                    <button
                      onClick={() => setAddingProvider((v) => !v)}
                      className="inline-flex items-center gap-1 rounded-md bg-blue-50 px-2 py-1 text-[11px] font-medium text-blue-700 hover:bg-blue-100"
                    >
                      <Plus className="h-3 w-3" />
                      <span>新增接入</span>
                    </button>
                  </div>

                  {addingProvider && (
                    <form
                      onSubmit={(e) => void handleCreateCustomProvider(e)}
                      className="mb-2 rounded-xl border border-blue-200 bg-blue-50/40 p-3 space-y-2 text-xs"
                    >
                      <input
                        required
                        value={newProvName}
                        onChange={(e) => setNewProvName(e.target.value)}
                        placeholder="名称（如：公司内网 vLLM）"
                        className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 outline-none"
                      />
                      <div className="grid grid-cols-2 gap-2">
                        <select
                          value={newProvCategory}
                          onChange={(e) => setNewProvCategory(e.target.value as 'cloud' | 'local')}
                          className="rounded-lg border border-slate-200 bg-white px-2 py-1.5"
                        >
                          <option value="cloud">云端 API</option>
                          <option value="local">本地/内网</option>
                        </select>
                        <select
                          value={newProvProtocol}
                          onChange={(e) =>
                            setNewProvProtocol(
                              e.target.value as 'openai' | 'anthropic' | 'gemini' | 'ollama'
                            )
                          }
                          className="rounded-lg border border-slate-200 bg-white px-2 py-1.5"
                        >
                          <option value="openai">OpenAI 兼容</option>
                          <option value="ollama">Ollama</option>
                          <option value="anthropic">Claude</option>
                          <option value="gemini">Gemini</option>
                        </select>
                      </div>
                      <input
                        required
                        value={newProvBaseUrl}
                        onChange={(e) => setNewProvBaseUrl(e.target.value)}
                        placeholder="Base URL (http/https)"
                        className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 outline-none"
                      />
                      <input
                        value={newProvModel}
                        onChange={(e) => setNewProvModel(e.target.value)}
                        placeholder="默认模型 ID"
                        className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 outline-none"
                      />
                      <div className="flex gap-2">
                        <button
                          type="submit"
                          className="flex-1 rounded-lg bg-blue-600 py-1.5 font-medium text-white hover:bg-blue-700"
                        >
                          添加
                        </button>
                        <button
                          type="button"
                          onClick={() => setAddingProvider(false)}
                          className="rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-slate-600"
                        >
                          取消
                        </button>
                      </div>
                    </form>
                  )}
                  {config.providers.map((p) => {
                    const isActive = p.id === config.settings.activeProviderId;
                    const isSelected = p.id === selectedProvider?.id;
                    return (
                      <button
                        key={p.id}
                        onClick={() => setSelectedId(p.id)}
                        className={`w-full flex items-center justify-between rounded-xl px-3 py-2.5 text-left transition ${
                          isSelected
                            ? 'bg-blue-50 text-blue-900 border border-blue-200'
                            : 'hover:bg-slate-50 text-slate-800 border border-transparent'
                        }`}
                      >
                        <div className="min-w-0">
                          <div className="flex items-center gap-1.5">
                            <span className="text-sm font-medium truncate">{p.name}</span>
                            {isActive && (
                              <span className="rounded bg-blue-600 px-1.5 py-0.5 text-[10px] font-medium text-white">
                                当前默认
                              </span>
                            )}
                          </div>
                          <div className="text-[11px] text-slate-500 truncate">
                            {p.category === 'local' ? '本地离线' : p.hasApiKey ? `Key: ${p.apiKeyMasked}` : '待填写 API Key'}
                          </div>
                        </div>
                        <span
                          className={`h-2 w-2 rounded-full shrink-0 ${
                            p.category === 'local' || p.hasApiKey ? 'bg-emerald-500' : 'bg-slate-300'
                          }`}
                        />
                      </button>
                    );
                  })}
                </div>
              </aside>

              {selectedProvider && (
                <section className="lg:col-span-8 rounded-2xl border border-slate-200 bg-white p-6 shadow-2xs space-y-5">
                  <div className="flex items-start justify-between border-b border-slate-100 pb-4">
                    <div>
                      <div className="flex items-center gap-2">
                        <h3 className="text-base font-bold text-slate-900">{selectedProvider.name}</h3>
                        <span className="rounded-md bg-slate-100 px-2 py-0.5 text-xs text-slate-600">
                          协议: {selectedProvider.protocol}
                        </span>
                      </div>
                      <p className="mt-1 text-xs text-slate-500">
                        {PROVIDER_PRESETS.find((pr) => pr.id === selectedProvider.presetId)?.description}
                      </p>
                    </div>
                    <label className="inline-flex items-center gap-2 text-xs font-medium text-slate-700">
                      <input
                        type="checkbox"
                        checked={formEnabled}
                        onChange={(e) => setFormEnabled(e.target.checked)}
                        className="h-4 w-4 rounded border-slate-300 text-blue-600"
                      />
                      <span>在侧边栏显示</span>
                    </label>
                  </div>

                  <div className="space-y-4">
                    <div>
                      <label className="block text-xs font-medium text-slate-700 mb-1">
                        API Base URL
                      </label>
                      <input
                        type="text"
                        value={formBaseUrl}
                        onChange={(e) => setFormBaseUrl(e.target.value)}
                        className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-500"
                      />
                    </div>

                    {selectedProvider.category === 'cloud' && (
                      <div>
                        <label className="block text-xs font-medium text-slate-700 mb-1">
                          API Key {selectedProvider.hasApiKey && `(当前已保存: ${selectedProvider.apiKeyMasked})`}
                        </label>
                        <div className="relative">
                          <input
                            type={showKey ? 'text' : 'password'}
                            value={formApiKey}
                            onChange={(e) => setFormApiKey(e.target.value)}
                            placeholder={
                              selectedProvider.hasApiKey
                                ? '留空表示保持当前已保存的 API Key 不变'
                                : '请输入你的 API Key...'
                            }
                            className="w-full rounded-xl border border-slate-200 px-3 py-2 pr-10 text-sm outline-none focus:border-blue-500"
                          />
                          <button
                            type="button"
                            onClick={() => setShowKey((s) => !s)}
                            className="absolute right-2.5 top-2.5 text-slate-400 hover:text-slate-600"
                          >
                            {showKey ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
                          </button>
                        </div>
                      </div>
                    )}

                    <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                      <div>
                        <label className="block text-xs font-medium text-slate-700 mb-1">
                          默认使用模型 (Model ID)
                        </label>
                        <input
                          type="text"
                          value={formModel}
                          onChange={(e) => setFormModel(e.target.value)}
                          placeholder="例如 gpt-4o / deepseek-chat / qwen2.5:7b"
                          className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-500"
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-medium text-slate-700 mb-1">
                          快速选择已发现模型
                        </label>
                        <select
                          value={formModel}
                          onChange={(e) => setFormModel(e.target.value)}
                          className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none focus:border-blue-500"
                        >
                          {selectedProvider.models.map((m) => (
                            <option key={m} value={m}>
                              {m}
                            </option>
                          ))}
                        </select>
                      </div>
                    </div>

                    <div>
                      <label className="block text-xs font-medium text-slate-700 mb-1">
                        候选模型列表（逗号分隔，可通过「测试并拉取模型」自动填充）
                      </label>
                      <textarea
                        rows={2}
                        value={formCustomModels}
                        onChange={(e) => setFormCustomModels(e.target.value)}
                        className="w-full rounded-xl border border-slate-200 px-3 py-2 text-xs outline-none focus:border-blue-500"
                      />
                    </div>
                  </div>

                  {testMsg && (
                    <div
                      className={`flex items-center gap-2 rounded-xl border px-3.5 py-2.5 text-xs ${
                        testMsg.ok
                          ? 'border-emerald-200 bg-emerald-50 text-emerald-800'
                          : 'border-rose-200 bg-rose-50 text-rose-800'
                      }`}
                    >
                      {testMsg.ok ? (
                        <CheckCircle2 className="h-4 w-4 shrink-0 text-emerald-600" />
                      ) : (
                        <XCircle className="h-4 w-4 shrink-0 text-rose-600" />
                      )}
                      <span>{testMsg.text}</span>
                    </div>
                  )}

                  <div className="flex flex-wrap items-center gap-3 pt-2">
                    <button
                      onClick={() => void handleTestProvider()}
                      disabled={testing}
                      className="inline-flex items-center gap-1.5 rounded-xl border border-slate-200 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50 disabled:opacity-50"
                    >
                      {testing ? (
                        <Loader2 className="h-4 w-4 animate-spin" />
                      ) : (
                        <Zap className="h-4 w-4 text-amber-500" />
                      )}
                      <span>测试连接并拉取模型</span>
                    </button>

                    <button
                      onClick={() => void handleSaveProvider(false)}
                      disabled={saving}
                      className="inline-flex items-center gap-1.5 rounded-xl border border-blue-200 bg-blue-50 px-4 py-2 text-sm font-medium text-blue-700 hover:bg-blue-100 disabled:opacity-50"
                    >
                      <span>保存配置</span>
                    </button>

                    <button
                      onClick={() => void handleSaveProvider(true)}
                      disabled={saving}
                      className="inline-flex items-center gap-1.5 rounded-xl bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700 disabled:opacity-50"
                    >
                      <span>保存并设为默认模型</span>
                    </button>

                    {selectedProvider.category === 'cloud' && selectedProvider.hasApiKey && (
                      <button
                        onClick={() => void handleClearApiKey()}
                        disabled={saving}
                        className="rounded-xl border border-slate-200 px-3 py-2 text-xs text-slate-600 hover:bg-slate-50"
                      >
                        清除 Key
                      </button>
                    )}

                    {selectedProvider.id.startsWith('custom-') && (
                      <button
                        onClick={async () => {
                          await apiClient.deleteProvider(selectedProvider.id);
                          await onRefreshConfig();
                          setSelectedId('ollama');
                        }}
                        className="rounded-xl border border-rose-200 px-3 py-2 text-xs text-rose-600 hover:bg-rose-50"
                      >
                        删除此配置
                      </button>
                    )}
                  </div>
                </section>
              )}
            </div>

            {/* Global Parameters */}
            <form
              onSubmit={(e) => void handleSaveGlobalSettings(e)}
              className="rounded-2xl border border-slate-200 bg-white p-5 shadow-2xs space-y-4"
            >
              <div className="flex items-center justify-between">
                <h3 className="text-sm font-bold text-slate-900">全局推理参数与默认系统提示词</h3>
                {globalMsg && <span className="text-xs text-emerald-600 font-medium">{globalMsg}</span>}
              </div>
              <div className="grid grid-cols-1 gap-4 md:grid-cols-4">
                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    Temperature (0 ~ 2)
                  </label>
                  <input
                    type="number"
                    min={0}
                    max={2}
                    step={0.1}
                    value={globalTemp}
                    onChange={(e) => setGlobalTemp(Number(e.target.value))}
                    className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-500"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    Max Tokens (0 = 不限)
                  </label>
                  <input
                    type="number"
                    min={0}
                    max={128000}
                    step={256}
                    value={globalMaxTokens}
                    onChange={(e) => setGlobalMaxTokens(Number(e.target.value))}
                    className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-500"
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    全局 System Prompt
                  </label>
                  <input
                    type="text"
                    value={globalSystemPrompt}
                    onChange={(e) => setGlobalSystemPrompt(e.target.value)}
                    className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-500"
                  />
                </div>
              </div>
              <div className="flex justify-end">
                <button
                  type="submit"
                  disabled={savingGlobal}
                  className="rounded-xl border border-slate-200 bg-slate-900 px-4 py-2 text-xs font-medium text-white hover:bg-slate-800 disabled:opacity-50"
                >
                  {savingGlobal ? '保存中...' : '保存全局参数'}
                </button>
              </div>
            </form>
          </div>
        )}

        {/* TAB 2: WINDOWS INTEGRATION */}
        {tab === 'integration' && (
          <div className="space-y-6">
            {winActionMsg && (
              <div className="rounded-xl border border-blue-200 bg-blue-50 px-4 py-3 text-sm text-blue-900">
                {winActionMsg}
              </div>
            )}

            <div className="grid grid-cols-1 gap-6 md:grid-cols-2">
              {/* WPS Card */}
              <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-2xs space-y-4">
                <div className="flex items-start justify-between">
                  <div>
                    <h3 className="text-base font-bold text-slate-900">WPS Office (文字 / 表格 / 演示)</h3>
                    <p className="mt-1 text-xs text-slate-500">
                      本地离线加载项：即使 OfficeAI 服务未启动，「OfficeAI 助手」选项卡也会显示，点击按钮时自动拉起服务。兼容个人免费版与专业版。
                    </p>
                  </div>
                  <span
                    className={`rounded-full px-2.5 py-1 text-xs font-medium ${
                      winStatus?.wps.installed
                        ? 'bg-emerald-100 text-emerald-700'
                        : 'bg-slate-100 text-slate-600'
                    }`}
                  >
                    {winStatus?.wps.installed ? '已安装' : '未安装'}
                  </span>
                </div>

                <div className="rounded-xl bg-slate-50 p-3 text-xs text-slate-600 space-y-1">
                  <div>环境检测: {winStatus?.wps.detected ? '已检测到 WPS 目录' : '未检测到或非 Windows 环境'}</div>
                  <div className="truncate">加载项目录: {winStatus?.wps.addonsDir || '%APPDATA%\\kingsoft\\wps\\jsaddons'}</div>
                  <div>
                    加载方式:{' '}
                    {winStatus?.wps.mode === 'local'
                      ? '本地离线加载项（点击按钮时可自动拉起服务）'
                      : winStatus?.wps.mode === 'online'
                        ? '在线加载项（需要先启动服务，重新「一键安装」可切换为本地模式）'
                        : '未安装'}
                  </div>
                  {winStatus?.wps.localAddinPath && (
                    <div className="truncate">本地插件目录: {winStatus.wps.localAddinPath}</div>
                  )}
                </div>

                <div className="flex flex-wrap gap-2">
                  <button
                    disabled={winBusy}
                    onClick={() => void handleWinAction(() => apiClient.installWps())}
                    className="rounded-xl bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700 disabled:opacity-50"
                  >
                    一键安装到 WPS Office
                  </button>
                  <button
                    disabled={winBusy}
                    onClick={() => void handleWinAction(() => apiClient.uninstallWps())}
                    className="rounded-xl border border-slate-200 px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50 disabled:opacity-50"
                  >
                    卸载 WPS 插件
                  </button>
                </div>
              </div>

              {/* Microsoft Office Card */}
              <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-2xs space-y-4">
                <div className="flex items-start justify-between">
                  <div>
                    <h3 className="text-base font-bold text-slate-900">Microsoft Office (Word / Excel / PPT)</h3>
                    <p className="mt-1 text-xs text-slate-500">
                      通过 Office Web Add-in 旁加载与本地 HTTPS 信任证书集成（无需管理员权限），兼容 Office 2016 / 2019 / 2021 / Microsoft 365。
                    </p>
                  </div>
                  <span
                    className={`rounded-full px-2.5 py-1 text-xs font-medium ${
                      winStatus?.msOffice.installed
                        ? 'bg-emerald-100 text-emerald-700'
                        : 'bg-slate-100 text-slate-600'
                    }`}
                  >
                    {winStatus?.msOffice.installed ? '已安装' : '未安装'}
                  </span>
                </div>

                <div className="rounded-xl bg-slate-50 p-3 text-xs text-slate-600 space-y-1">
                  <div>Office 检测: {winStatus?.msOffice.detected ? winStatus.msOffice.version : '未检测到或非 Windows 环境'}</div>
                  <div>本地 HTTPS 证书: {winStatus?.certTrusted ? '已受系统信任' : winStatus?.certExists ? '已生成（安装时将自动信任）' : '待生成'}</div>
                  <div className="truncate">Manifest: {winStatus?.msOffice.manifestPath}</div>
                </div>

                <div className="flex flex-wrap gap-2">
                  <button
                    disabled={winBusy}
                    onClick={() => void handleWinAction(() => apiClient.installOffice())}
                    className="rounded-xl bg-indigo-600 px-4 py-2 text-sm font-medium text-white hover:bg-indigo-700 disabled:opacity-50"
                  >
                    一键安装到 MS Office
                  </button>
                  <button
                    disabled={winBusy}
                    onClick={() => void handleWinAction(() => apiClient.trustCert())}
                    className="inline-flex items-center gap-1 rounded-xl border border-slate-200 px-3 py-2 text-xs font-medium text-slate-700 hover:bg-slate-50"
                  >
                    <ShieldCheck className="h-4 w-4 text-emerald-600" />
                    <span>信任本地证书</span>
                  </button>
                  <a
                    href="/office/manifest.xml"
                    download="officeai-manifest.xml"
                    className="inline-flex items-center gap-1 rounded-xl border border-slate-200 px-3 py-2 text-xs font-medium text-slate-700 hover:bg-slate-50"
                  >
                    <Download className="h-4 w-4" />
                    <span>下载 Manifest</span>
                  </a>
                  <button
                    disabled={winBusy}
                    onClick={() => void handleWinAction(() => apiClient.uninstallOffice())}
                    className="rounded-xl border border-slate-200 px-3 py-2 text-xs font-medium text-slate-700 hover:bg-slate-50"
                  >
                    卸载
                  </button>
                </div>
              </div>
            </div>

            {/* On-demand launch settings */}
            <form
              onSubmit={(e) => {
                e.preventDefault();
                void handleSaveLaunchSettings();
              }}
              className="rounded-2xl border border-slate-200 bg-white p-6 shadow-2xs space-y-4"
            >
              <div className="flex items-start justify-between gap-4">
                <div>
                  <h3 className="text-base font-bold text-slate-900">启动方式（按需启动，不做开机自启）</h3>
                  <p className="mt-1 text-xs text-slate-500">
                    助手生命周期与办公软件绑定：打开 Office / WPS 时自动拉起（安装唤醒器后 Word/Excel/PPT 也能自动拉起），
                    所有办公窗口关闭或侧边栏心跳停止后自动退出，不浪费系统资源。
                  </p>
                </div>
                {launchMsg && <span className="text-xs font-medium text-emerald-600">{launchMsg}</span>}
              </div>

              <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
                {[
                  {
                    id: 'on-demand' as const,
                    title: '按需启动（推荐）',
                    desc: '点击加载项按钮时自动启动，全部 Office/WPS 关闭后自动退出'
                  },
                  {
                    id: 'always' as const,
                    title: '常驻托盘',
                    desc: '启动 OfficeAI 后一直后台运行，直到手动退出'
                  },
                  {
                    id: 'manual' as const,
                    title: '仅手动',
                    desc: '只能通过桌面图标 / 批处理启动，加载项只做提示'
                  }
                ].map((option) => (
                  <label
                    key={option.id}
                    className={`cursor-pointer rounded-xl border p-3 text-xs transition ${
                      launchMode === option.id
                        ? 'border-blue-500 bg-blue-50/60'
                        : 'border-slate-200 hover:bg-slate-50'
                    }`}
                  >
                    <input
                      type="radio"
                      name="launchMode"
                      className="mr-2"
                      checked={launchMode === option.id}
                      onChange={() => setLaunchMode(option.id)}
                    />
                    <span className="font-semibold text-slate-900">{option.title}</span>
                    <p className="mt-1 leading-relaxed text-slate-500">{option.desc}</p>
                  </label>
                ))}
              </div>

              <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    启动命令（留空自动探测 OfficeAI.exe / Start-OfficeAI.bat）
                  </label>
                  <input
                    type="text"
                    value={gatewayLaunchCommand}
                    onChange={(e) => setGatewayLaunchCommand(e.target.value)}
                    placeholder="C:\Users\me\AppData\Local\Programs\OfficeAI\OfficeAI.exe"
                    className="w-full rounded-xl border border-slate-200 px-3 py-2 text-xs outline-none focus:border-blue-500"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    空闲自动退出（分钟，0 = 不按空闲退出）
                  </label>
                  <input
                    type="number"
                    min={0}
                    max={240}
                    value={idleExitMinutes}
                    onChange={(e) => setIdleExitMinutes(Number(e.target.value))}
                    className="w-full rounded-xl border border-slate-200 px-3 py-2 text-xs outline-none focus:border-blue-500"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    侧边栏心跳超时（分钟，页面关闭后多久释放网关）
                  </label>
                  <input
                    type="number"
                    min={1}
                    max={30}
                    value={sessionGraceMinutes}
                    onChange={(e) => setSessionGraceMinutes(Number(e.target.value))}
                    className="w-full rounded-xl border border-slate-200 px-3 py-2 text-xs outline-none focus:border-blue-500"
                  />
                </div>
              </div>

              <label className="flex items-center gap-2 text-xs text-slate-700">
                <input
                  type="checkbox"
                  checked={exitWithOffice}
                  onChange={(e) => setExitWithOffice(e.target.checked)}
                />
                <span>Word / Excel / PowerPoint / WPS 全部关闭后自动退出网关</span>
              </label>

              <div className="rounded-xl bg-slate-50 p-3 text-xs text-slate-600 space-y-1">
                <div>
                  当前启动程序：
                  {winStatus?.launch?.target.command ? (
                    <span className="font-mono">{winStatus.launch.target.command}</span>
                  ) : (
                    <span className="text-amber-600">未检测到，请在上方手动填写</span>
                  )}
                </div>
                <div>
                  一键启动协议 officeai://：
                  {winStatus?.launch?.protocolRegistered ? '已注册' : '未注册'}
                  {!winStatus?.launch?.protocolRegistered && (
                    <button
                      type="button"
                      disabled={winBusy}
                      onClick={() => void handleWinAction(() => apiClient.registerProtocol())}
                      className="ml-2 rounded-lg border border-slate-200 bg-white px-2 py-0.5 text-xs text-slate-700 hover:bg-slate-50 disabled:opacity-50"
                    >
                      立即注册并创建桌面快捷方式
                    </button>
                  )}
                </div>
                <div>Office / WPS 进程：{winStatus?.launch?.officeRunning ? '运行中' : '未检测到'}</div>
                {winStatus?.launch?.lifecycle?.sessions?.length ? (
                  <div>
                    活跃侧边栏：
                    {winStatus.launch.lifecycle.sessions
                      .map((s) => `${s.app}（${s.idleSeconds}s）`)
                      .join('、')}
                  </div>
                ) : (
                  <div>活跃侧边栏：无（网关将按设定自动退出）</div>
                )}
              </div>

              <div className="rounded-xl border border-indigo-100 bg-indigo-50/60 p-3 text-xs text-slate-700 space-y-3">
                <div className="font-semibold text-slate-900">由谁负责「打开办公软件就启动」</div>
                <div className="grid grid-cols-1 gap-2 md:grid-cols-2">
                  <label
                    className={`cursor-pointer rounded-lg border p-2.5 ${
                      launchTrigger === 'watcher'
                        ? 'border-indigo-500 bg-white'
                        : 'border-slate-200 bg-white/60'
                    }`}
                  >
                    <input
                      type="radio"
                      name="launchTrigger"
                      className="mr-2"
                      checked={launchTrigger === 'watcher'}
                      onChange={() => setLaunchTrigger('watcher')}
                    />
                    <span className="font-semibold">后台监听器（推荐）</span>
                    <p className="mt-1 leading-relaxed text-slate-600">
                      常驻约 10MB 的轻量进程，每 2 秒看一次进程列表；发现 Office/WPS 启动就拉起网关。
                      <strong>Office 启动零开销</strong>，不会被“慢加载项”禁用。需要登录自启。
                    </p>
                    <div className="mt-2 flex items-center gap-2">
                      <span
                        className={`rounded-full px-2 py-0.5 text-[10px] ${
                          winStatus?.watcher?.installed
                            ? 'bg-emerald-100 text-emerald-700'
                            : 'bg-slate-200 text-slate-600'
                        }`}
                      >
                        {winStatus?.watcher?.installed
                          ? winStatus.watcher?.running
                            ? `运行中 · ${winStatus.watcher.memoryMB ?? '?'}MB`
                            : '已注册未运行'
                          : '未安装'}
                      </span>
                      <button
                        type="button"
                        disabled={winBusy}
                        onClick={() =>
                          void handleWinAction(() =>
                            winStatus?.watcher?.installed
                              ? apiClient.uninstallWatcher()
                              : apiClient.installWatcher()
                          )
                        }
                        className="rounded-lg border border-slate-300 bg-white px-2 py-0.5 text-[11px] text-slate-700 hover:bg-slate-50 disabled:opacity-50"
                      >
                        {winStatus?.watcher?.installed ? '卸载监听器' : '安装监听器'}
                      </button>
                    </div>
                  </label>

                  <label
                    className={`cursor-pointer rounded-lg border p-2.5 ${
                      launchTrigger === 'waker' ? 'border-indigo-500 bg-white' : 'border-slate-200 bg-white/60'
                    }`}
                  >
                    <input
                      type="radio"
                      name="launchTrigger"
                      className="mr-2"
                      checked={launchTrigger === 'waker'}
                      onChange={() => setLaunchTrigger('waker')}
                    />
                    <span className="font-semibold">Office 唤醒器（无常驻）</span>
                    <p className="mt-1 leading-relaxed text-slate-600">
                      极小的 COM 加载项，由 Word/Excel/PPT 在启动时加载并探活（300ms 超时）。
                      不占常驻内存，但会给 Office 启动增加极少量耗时。
                    </p>
                    <div className="mt-2 flex items-center gap-2">
                      <span
                        className={`rounded-full px-2 py-0.5 text-[10px] ${
                          winStatus?.waker?.installed
                            ? 'bg-emerald-100 text-emerald-700'
                            : 'bg-slate-200 text-slate-600'
                        }`}
                      >
                        {winStatus?.waker?.installed ? '已安装' : '未安装'}
                      </span>
                      <button
                        type="button"
                        disabled={winBusy}
                        onClick={() =>
                          void handleWinAction(() =>
                            winStatus?.waker?.installed
                              ? apiClient.uninstallWaker()
                              : apiClient.installWaker()
                          )
                        }
                        className="rounded-lg border border-slate-300 bg-white px-2 py-0.5 text-[11px] text-slate-700 hover:bg-slate-50 disabled:opacity-50"
                      >
                        {winStatus?.waker?.installed ? '卸载唤醒器' : '安装唤醒器'}
                      </button>
                    </div>
                    {winStatus?.waker?.installed && (
                      <div className="mt-1 text-[10px] text-slate-500">
                        LoadBehavior:{' '}
                        {Object.entries(winStatus.waker.loadBehavior)
                          .map(([app, v]) => `${app}=${v ?? '未注册'}`)
                          .join(' · ')}
                      </div>
                    )}
                  </label>
                </div>
                <p className="text-[11px] text-slate-500">
                  只安装与所选方式对应的组件；切换后请点「保存启动方式」并重新安装一次，卸载掉另一种方式。
                </p>
              </div>

              <div className="flex justify-end">
                <button
                  type="submit"
                  disabled={savingLaunch}
                  className="rounded-xl bg-slate-900 px-4 py-2 text-xs font-medium text-white hover:bg-slate-800 disabled:opacity-50"
                >
                  {savingLaunch ? '保存中...' : '保存启动方式'}
                </button>
              </div>
            </form>

            {/* Instructions */}
            <div className="rounded-2xl border border-slate-200 bg-white p-6 text-xs text-slate-600 space-y-2">
              <h4 className="text-sm font-semibold text-slate-900">Windows 使用说明</h4>
              <p>1. 默认「按需启动」：无需开机自启，也不会常驻后台——关闭所有 Office / WPS 后网关自动退出。</p>
              <p>2. <strong>WPS 用户</strong>：安装后加载项是本地离线版，即使服务没启动也能出现在功能区；点击任意按钮会自动拉起网关（首次可能弹出一次 WPS 安全确认）。</p>
              <p>3. <strong>Microsoft Office 用户</strong>：Office 加载项必须由本地服务提供，无法自行启动进程。请先注册 <code>officeai://</code> 协议（上方按钮），之后点击桌面「OfficeAI 启动」快捷方式即可拉起网关，再打开 Word/Excel/PPT 使用。</p>
            </div>
          </div>
        )}

        {/* TAB 3: PROMPTS */}
        {tab === 'prompts' && (
          <div className="grid grid-cols-1 gap-6 lg:grid-cols-12">
            <div className="lg:col-span-7 space-y-3">
              {config.prompts.map((p) => (
                <div
                  key={p.id}
                  className="flex items-start justify-between rounded-2xl border border-slate-200 bg-white p-4 shadow-2xs"
                >
                  <div className="space-y-1 pr-4">
                    <div className="flex items-center gap-2">
                      <span className="font-semibold text-sm text-slate-900">{p.title}</span>
                      <span className="rounded bg-slate-100 px-2 py-0.5 text-[11px] text-slate-600">
                        {p.category}
                      </span>
                    </div>
                    <p className="text-xs text-slate-500">{p.systemPrompt}</p>
                  </div>
                  {!p.builtIn && (
                    <button
                      onClick={async () => {
                        await apiClient.deletePrompt(p.id);
                        await onRefreshConfig();
                      }}
                      className="rounded-lg p-1.5 text-slate-400 hover:bg-rose-50 hover:text-rose-600"
                    >
                      <Trash2 className="h-4 w-4" />
                    </button>
                  )}
                </div>
              ))}
            </div>

            <form
              onSubmit={(e) => void handleAddPrompt(e)}
              className="lg:col-span-5 h-fit rounded-2xl border border-slate-200 bg-white p-5 shadow-2xs space-y-4"
            >
              <h3 className="flex items-center gap-2 text-sm font-bold text-slate-900">
                <Plus className="h-4 w-4 text-blue-600" />
                <span>新增自定义快捷指令</span>
              </h3>

              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">指令名称</label>
                <input
                  required
                  value={newPromptTitle}
                  onChange={(e) => setNewPromptTitle(e.target.value)}
                  placeholder="例如：生成周报 / SQL 优化"
                  className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-500"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">系统提示词 (System Prompt)</label>
                <textarea
                  required
                  rows={3}
                  value={newPromptSystem}
                  onChange={(e) => setNewPromptSystem(e.target.value)}
                  placeholder="定义 AI 的角色与输出要求..."
                  className="w-full rounded-xl border border-slate-200 px-3 py-2 text-xs outline-none focus:border-blue-500"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  用户模板 (支持 {'{{selection}}'} 占位符)
                </label>
                <textarea
                  rows={2}
                  value={newPromptTemplate}
                  onChange={(e) => setNewPromptTemplate(e.target.value)}
                  className="w-full rounded-xl border border-slate-200 px-3 py-2 text-xs outline-none focus:border-blue-500"
                />
              </div>

              <button
                type="submit"
                className="w-full rounded-xl bg-blue-600 py-2 text-sm font-medium text-white hover:bg-blue-700"
              >
                添加快捷指令
              </button>
            </form>
          </div>
        )}

        {/* TAB 4: USAGE */}
        {tab === 'usage' && (
          <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-2xs space-y-4">
            <div className="flex items-center justify-between">
              <h3 className="text-sm font-bold text-slate-900">最近调用记录 ({usage.length})</h3>
              <div className="flex gap-2">
                <button
                  onClick={() => void loadUsage()}
                  className="rounded-lg border border-slate-200 px-3 py-1.5 text-xs text-slate-700 hover:bg-slate-50"
                >
                  刷新
                </button>
                <button
                  onClick={async () => {
                    await apiClient.clearUsage();
                    setUsage([]);
                  }}
                  className="rounded-lg border border-rose-200 px-3 py-1.5 text-xs text-rose-600 hover:bg-rose-50"
                >
                  清空记录
                </button>
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs">
                <thead>
                  <tr className="border-b border-slate-200 text-slate-500">
                    <th className="py-2 pr-3">时间</th>
                    <th className="py-2 pr-3">宿主应用</th>
                    <th className="py-2 pr-3">提供方 / 模型</th>
                    <th className="py-2 pr-3">耗时</th>
                    <th className="py-2 pr-3">Token (入/出)</th>
                    <th className="py-2">状态</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {usage.length === 0 && (
                    <tr>
                      <td colSpan={6} className="py-8 text-center text-slate-400">
                        暂无调用记录
                      </td>
                    </tr>
                  )}
                  {usage.map((r) => (
                    <tr key={r.id}>
                      <td className="py-2 pr-3 text-slate-500">
                        {new Date(r.timestamp).toLocaleString()}
                      </td>
                      <td className="py-2 pr-3">{r.hostApp}</td>
                      <td className="py-2 pr-3 font-medium">
                        {r.providerName} · {r.model}
                      </td>
                      <td className="py-2 pr-3">{r.durationMs} ms</td>
                      <td className="py-2 pr-3">
                        {r.promptTokens ?? '-'} / {r.completionTokens ?? '-'}
                      </td>
                      <td className="py-2">
                        {r.success ? (
                          <span className="text-emerald-600 font-medium">成功</span>
                        ) : (
                          <span className="text-rose-600" title={r.error}>
                            失败
                          </span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>
    </div>
  );
};
