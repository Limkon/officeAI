import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  Check,
  ChevronDown,
  ChevronRight,
  Copy,
  FileText,
  Loader2,
  Replace,
  Send,
  Settings,
  Sparkles,
  Square,
  Table2,
  Trash2,
  Wand2
} from 'lucide-react';
import {
  renderPromptTemplate,
  type ChatMessage,
  type MaskedProviderConfig,
  type PromptAction
} from '@officeai/shared';
import { apiClient, type PublicConfigResponse } from '../api/client';
import {
  readHostSelection,
  writeHostContent,
  type HostContextInfo
} from '../office/hostAdapter';

interface UiMessage {
  id: string;
  role: 'user' | 'assistant';
  content: string;
  reasoning?: string;
  model?: string;
  error?: string;
}

interface TaskPaneViewProps {
  config: PublicConfigResponse;
  host: HostContextInfo;
  onRefreshConfig: () => Promise<void>;
  onOpenDashboard: () => void;
}

export const TaskPaneView: React.FC<TaskPaneViewProps> = ({
  config,
  host,
  onRefreshConfig,
  onOpenDashboard
}) => {
  const [messages, setMessages] = useState<UiMessage[]>([]);
  const [input, setInput] = useState('');
  const [selectionText, setSelectionText] = useState('');
  const [selectionContext, setSelectionContext] = useState('');
  const [streaming, setStreaming] = useState(false);
  const [statusTip, setStatusTip] = useState<string | null>(null);
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [expandedThinking, setExpandedThinking] = useState<Record<string, boolean>>({});

  const abortRef = useRef<AbortController | null>(null);
  const chatEndRef = useRef<HTMLDivElement | null>(null);

  const enabledProviders = useMemo(() => {
    return config.providers.filter(
      (p) => p.enabled || p.id === config.settings.activeProviderId
    );
  }, [config]);

  const activeProvider: MaskedProviderConfig | undefined = useMemo(() => {
    return (
      config.providers.find((p) => p.id === config.settings.activeProviderId) ||
      enabledProviders[0] ||
      config.providers[0]
    );
  }, [config, enabledProviders]);

  const activeModel =
    config.settings.activeModel ||
    activeProvider?.selectedModel ||
    activeProvider?.models[0] ||
    '';

  const contextPrompts = useMemo(() => {
    return config.prompts.filter(
      (p) => p.apps === 'all' || p.apps.includes(host.app)
    );
  }, [config.prompts, host.app]);

  const isExcelLike = host.app === 'excel' || host.app === 'wps-spreadsheets';

  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, streaming]);

  useEffect(() => {
    // Handle WPS Ribbon direct action (?action=polish etc.)
    const params = new URLSearchParams(window.location.search);
    const actionId = params.get('action');
    if (actionId) {
      const prompt = config.prompts.find((p) => p.id === actionId);
      if (prompt) {
        void handleQuickAction(prompt);
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const showTip = (msg: string) => {
    setStatusTip(msg);
    setTimeout(() => setStatusTip((prev) => (prev === msg ? null : prev)), 3500);
  };

  const handleModelChange = async (providerId: string, model: string) => {
    await apiClient.updateSettings({
      activeProviderId: providerId,
      activeModel: model
    });
    await apiClient.saveProvider(providerId, { selectedModel: model, enabled: true });
    await onRefreshConfig();
  };

  const handlePullSelection = async () => {
    try {
      const snap = await readHostSelection(host.app);
      if (snap.text) {
        setSelectionText(snap.text);
        setSelectionContext(snap.context || '');
        showTip(`已读取选区 (${snap.text.length} 字)`);
      } else if (snap.context) {
        setSelectionContext(snap.context);
        showTip('未选中局部文本，已读取当前文档上下文');
      } else {
        showTip(
          host.suite === 'standalone'
            ? '当前为独立预览模式，可在下方直接输入或粘贴待处理文本'
            : '未检测到选中内容，请先在文档中选中文本或单元格'
        );
      }
    } catch (err) {
      showTip(`读取选区失败: ${err instanceof Error ? err.message : String(err)}`);
    }
  };

  const runChat = async (
    userContent: string,
    customSystemPrompt?: string,
    displayUserLabel?: string
  ) => {
    if (!userContent.trim() || streaming) return;
    if (!activeProvider) {
      showTip('请先在设置中心配置模型');
      return;
    }

    const userMsg: UiMessage = {
      id: `u-${Date.now()}`,
      role: 'user',
      content: displayUserLabel || userContent
    };
    const assistantId = `a-${Date.now() + 1}`;
    const assistantMsg: UiMessage = {
      id: assistantId,
      role: 'assistant',
      content: '',
      reasoning: '',
      model: `${activeProvider.name} · ${activeModel}`
    };

    const historyMessages: ChatMessage[] = [
      {
        role: 'system',
        content: customSystemPrompt || config.settings.systemPrompt
      },
      ...messages.slice(-8).map((m) => ({ role: m.role, content: m.content })),
      { role: 'user', content: userContent }
    ];

    setMessages((prev) => [...prev, userMsg, assistantMsg]);
    setStreaming(true);

    const controller = new AbortController();
    abortRef.current = controller;

    try {
      await apiClient.streamChat(
        {
          providerId: activeProvider.id,
          model: activeModel,
          messages: historyMessages,
          hostApp: host.app
        },
        (ev) => {
          if (ev.type === 'delta' && ev.content) {
            setMessages((prev) =>
              prev.map((m) =>
                m.id === assistantId ? { ...m, content: m.content + ev.content } : m
              )
            );
          } else if (ev.type === 'reasoning' && ev.content) {
            setMessages((prev) =>
              prev.map((m) =>
                m.id === assistantId
                  ? { ...m, reasoning: (m.reasoning || '') + ev.content }
                  : m
              )
            );
          } else if (ev.type === 'error') {
            setMessages((prev) =>
              prev.map((m) =>
                m.id === assistantId ? { ...m, error: ev.error || '生成出错' } : m
              )
            );
          }
        },
        controller.signal
      );
    } catch (err) {
      if ((err as Error).name !== 'AbortError') {
        const msg = err instanceof Error ? err.message : String(err);
        setMessages((prev) =>
          prev.map((m) => (m.id === assistantId ? { ...m, error: msg } : m))
        );
      }
    } finally {
      abortRef.current = null;
      setStreaming(false);
    }
  };

  const handleQuickAction = async (action: PromptAction) => {
    let text = selectionText.trim();
    let ctx = selectionContext.trim();

    if (!text && host.suite !== 'standalone') {
      const snap = await readHostSelection(host.app).catch(() => ({ text: '', context: '' }));
      text = snap.text.trim();
      ctx = (snap.context || '').trim();
      if (text) setSelectionText(text);
      if (ctx) setSelectionContext(ctx);
    }

    if (!text && input.trim()) {
      text = input.trim();
      setInput('');
    }

    if (!text) {
      showTip('请先在文档中选中内容，或在输入框/待处理区输入文本');
      return;
    }

    const promptText = renderPromptTemplate(action.userPromptTemplate, {
      selection: text,
      context: ctx || text
    });

    await runChat(
      promptText,
      action.systemPrompt,
      `[${action.title}] ${text.slice(0, 120)}${text.length > 120 ? '...' : ''}`
    );
  };

  const handleWriteBack = async (
    text: string,
    mode: 'replace' | 'insert' | 'formula'
  ) => {
    try {
      const res = await writeHostContent(host.app, text, mode);
      showTip(res.message);
    } catch (err) {
      showTip(`写入失败: ${err instanceof Error ? err.message : String(err)}`);
    }
  };

  const handleCopy = async (id: string, text: string) => {
    await navigator.clipboard.writeText(text);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 1500);
  };

  return (
    <div className="flex h-full flex-col bg-slate-50">
      {/* Header */}
      <header className="flex items-center justify-between border-b border-slate-200 bg-white px-3 py-2.5 shadow-2xs">
        <div className="flex items-center gap-2 min-w-0">
          <div className="flex h-7 w-7 shrink-0 items-center justify-center rounded-lg bg-gradient-to-br from-blue-600 to-indigo-600 text-white shadow-xs">
            <Sparkles className="h-4 w-4" />
          </div>
          <div className="min-w-0">
            <div className="flex items-center gap-1.5">
              <span className="text-sm font-semibold text-slate-900">OfficeAI</span>
              <span className="truncate rounded-full bg-blue-50 px-2 py-0.5 text-[11px] font-medium text-blue-700 border border-blue-100">
                {host.label}
              </span>
            </div>
          </div>
        </div>

        <button
          onClick={onOpenDashboard}
          title="打开模型配置与安装中心"
          className="inline-flex items-center gap-1 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 transition"
        >
          <Settings className="h-3.5 w-3.5" />
          <span>配置中心</span>
        </button>
      </header>

      {/* Model Switcher */}
      <div className="border-b border-slate-200 bg-white px-3 py-2">
        <div className="flex items-center gap-2">
          <select
            value={`${activeProvider?.id || ''}::${activeModel}`}
            onChange={(e) => {
              const [pid, ...rest] = e.target.value.split('::');
              const mdl = rest.join('::');
              if (pid && mdl) void handleModelChange(pid, mdl);
            }}
            className="w-full rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-1.5 text-xs font-medium text-slate-800 outline-none focus:border-blue-500 focus:bg-white"
          >
            {enabledProviders.map((provider) => (
              <optgroup
                key={provider.id}
                label={`${provider.category === 'local' ? '🖥️ 本地' : '☁️ 云端'} · ${provider.name}${
                  provider.category === 'cloud' && !provider.hasApiKey ? ' (未填Key)' : ''
                }`}
              >
                {provider.models.map((m) => (
                  <option key={`${provider.id}::${m}`} value={`${provider.id}::${m}`}>
                    {provider.name} — {m}
                  </option>
                ))}
              </optgroup>
            ))}
          </select>
        </div>
      </div>

      {/* Selection & Quick Actions */}
      <div className="border-b border-slate-200 bg-slate-50/80 px-3 py-2 space-y-2">
        <div className="flex items-center justify-between gap-2">
          <button
            onClick={() => void handlePullSelection()}
            className="inline-flex items-center gap-1.5 rounded-lg border border-blue-200 bg-blue-50 px-2.5 py-1 text-xs font-medium text-blue-700 hover:bg-blue-100 transition"
          >
            {isExcelLike ? <Table2 className="h-3.5 w-3.5" /> : <FileText className="h-3.5 w-3.5" />}
            <span>{isExcelLike ? '读取表格选区' : '读取文档选区'}</span>
          </button>

          {selectionText && (
            <button
              onClick={() => {
                setSelectionText('');
                setSelectionContext('');
              }}
              className="text-[11px] text-slate-500 hover:text-slate-700"
            >
              清空选区 ({selectionText.length}字)
            </button>
          )}
        </div>

        {(selectionText || host.suite === 'standalone') && (
          <textarea
            value={selectionText}
            onChange={(e) => setSelectionText(e.target.value)}
            rows={2}
            placeholder={
              host.suite === 'standalone'
                ? '可在此粘贴需要润色/翻译/分析的文档内容或表格数据，点击下方快捷按钮一键处理...'
                : '已读取的选区内容...'
            }
            className="w-full resize-none rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-700 outline-none focus:border-blue-400"
          />
        )}

        <div className="flex flex-wrap gap-1.5">
          {contextPrompts.slice(0, 8).map((action) => (
            <button
              key={action.id}
              disabled={streaming}
              onClick={() => void handleQuickAction(action)}
              className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2 py-1 text-[11px] font-medium text-slate-700 shadow-2xs hover:border-blue-300 hover:bg-blue-50/60 hover:text-blue-700 disabled:opacity-50 transition"
            >
              <Wand2 className="h-3 w-3 text-blue-500" />
              <span>{action.title}</span>
            </button>
          ))}
        </div>
      </div>

      {/* Status Toast */}
      {statusTip && (
        <div className="mx-3 mt-2 rounded-lg border border-blue-200 bg-blue-50 px-3 py-1.5 text-xs text-blue-800">
          {statusTip}
        </div>
      )}

      {/* Messages */}
      <div className="flex-1 overflow-y-auto px-3 py-3 space-y-3">
        {messages.length === 0 && (
          <div className="rounded-xl border border-dashed border-slate-300 bg-white p-4 text-center space-y-2 mt-2">
            <Sparkles className="mx-auto h-6 w-6 text-blue-500" />
            <div className="text-sm font-medium text-slate-800">
              AI 办公助手已就绪
            </div>
            <p className="text-xs text-slate-500 leading-relaxed">
              支持 OpenAI、Claude、Gemini、DeepSeek、通义千问、智谱 GLM、Kimi 与 Ollama / LM Studio 本地模型。
              选中内容后点击上方快捷按钮，或直接在下方提问。
            </p>
          </div>
        )}

        {messages.map((msg) => (
          <div
            key={msg.id}
            className={`flex flex-col ${
              msg.role === 'user' ? 'items-end' : 'items-start'
            }`}
          >
            <div
              className={`max-w-[94%] rounded-xl px-3 py-2 text-xs leading-relaxed shadow-2xs ${
                msg.role === 'user'
                  ? 'bg-blue-600 text-white'
                  : 'border border-slate-200 bg-white text-slate-800'
              }`}
            >
              {msg.role === 'assistant' && msg.model && (
                <div className="mb-1 text-[10px] font-medium text-slate-400">
                  {msg.model}
                </div>
              )}

              {msg.reasoning && (
                <div className="mb-2 rounded-lg border border-amber-200 bg-amber-50/70 p-2 text-[11px] text-amber-900">
                  <button
                    onClick={() =>
                      setExpandedThinking((p) => ({ ...p, [msg.id]: !p[msg.id] }))
                    }
                    className="flex items-center gap-1 font-medium text-amber-800"
                  >
                    {expandedThinking[msg.id] ? (
                      <ChevronDown className="h-3 w-3" />
                    ) : (
                      <ChevronRight className="h-3 w-3" />
                    )}
                    <span>深度思考过程</span>
                  </button>
                  {expandedThinking[msg.id] && (
                    <pre className="mt-1.5 whitespace-pre-wrap font-sans text-[11px] text-amber-900/90">
                      {msg.reasoning}
                    </pre>
                  )}
                </div>
              )}

              {msg.error ? (
                <div className="text-red-600 font-medium">错误: {msg.error}</div>
              ) : (
                <div className="whitespace-pre-wrap break-words">
                  {msg.content ||
                    (streaming && msg.role === 'assistant' ? (
                      <span className="inline-flex items-center gap-1 text-slate-400">
                        <Loader2 className="h-3.5 w-3.5 animate-spin" />
                        思考生成中...
                      </span>
                    ) : (
                      ''
                    ))}
                </div>
              )}

              {msg.role === 'assistant' && msg.content && !msg.error && (
                <div className="mt-2 flex flex-wrap items-center gap-1.5 border-t border-slate-100 pt-2">
                  <button
                    onClick={() => void handleCopy(msg.id, msg.content)}
                    className="inline-flex items-center gap-1 rounded bg-slate-100 px-2 py-1 text-[11px] text-slate-700 hover:bg-slate-200"
                  >
                    {copiedId === msg.id ? (
                      <Check className="h-3 w-3 text-emerald-600" />
                    ) : (
                      <Copy className="h-3 w-3" />
                    )}
                    <span>{copiedId === msg.id ? '已复制' : '复制'}</span>
                  </button>

                  <button
                    onClick={() => void handleWriteBack(msg.content, 'insert')}
                    className="inline-flex items-center gap-1 rounded bg-blue-50 px-2 py-1 text-[11px] text-blue-700 hover:bg-blue-100"
                  >
                    <span>插入文档</span>
                  </button>

                  {!isExcelLike && (
                    <button
                      onClick={() => void handleWriteBack(msg.content, 'replace')}
                      className="inline-flex items-center gap-1 rounded bg-indigo-50 px-2 py-1 text-[11px] text-indigo-700 hover:bg-indigo-100"
                    >
                      <Replace className="h-3 w-3" />
                      <span>替换选区</span>
                    </button>
                  )}

                  {isExcelLike && (
                    <button
                      onClick={() => void handleWriteBack(msg.content, 'formula')}
                      className="inline-flex items-center gap-1 rounded bg-emerald-50 px-2 py-1 text-[11px] text-emerald-700 hover:bg-emerald-100"
                    >
                      <Table2 className="h-3 w-3" />
                      <span>写入公式/单元格</span>
                    </button>
                  )}
                </div>
              )}
            </div>
          </div>
        ))}
        <div ref={chatEndRef} />
      </div>

      {/* Input Area */}
      <div className="border-t border-slate-200 bg-white p-2.5">
        <div className="flex items-end gap-2">
          <textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                const text = input.trim();
                if (text) {
                  setInput('');
                  const fullUser = selectionText.trim()
                    ? `参考选区内容：\n${selectionText.trim()}\n\n用户指令：${text}`
                    : text;
                  void runChat(fullUser, undefined, text);
                }
              }
            }}
            rows={2}
            placeholder="输入需求（Enter 发送，Shift+Enter 换行）..."
            className="flex-1 resize-none rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-800 outline-none focus:border-blue-500 focus:bg-white"
          />

          <div className="flex flex-col gap-1">
            {streaming ? (
              <button
                onClick={() => abortRef.current?.abort()}
                title="停止生成"
                className="flex h-8 w-8 items-center justify-center rounded-lg bg-rose-600 text-white hover:bg-rose-700"
              >
                <Square className="h-3.5 w-3.5" />
              </button>
            ) : (
              <button
                onClick={() => {
                  const text = input.trim();
                  if (!text) return;
                  setInput('');
                  const fullUser = selectionText.trim()
                    ? `参考选区内容：\n${selectionText.trim()}\n\n用户指令：${text}`
                    : text;
                  void runChat(fullUser, undefined, text);
                }}
                title="发送"
                className="flex h-8 w-8 items-center justify-center rounded-lg bg-blue-600 text-white hover:bg-blue-700"
              >
                <Send className="h-3.5 w-3.5" />
              </button>
            )}

            {messages.length > 0 && !streaming && (
              <button
                onClick={() => setMessages([])}
                title="清空对话"
                className="flex h-7 w-8 items-center justify-center rounded-lg border border-slate-200 text-slate-500 hover:bg-slate-100"
              >
                <Trash2 className="h-3.5 w-3.5" />
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};
