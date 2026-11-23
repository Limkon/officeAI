import React, { useCallback, useEffect, useState } from 'react';
import { Loader2 } from 'lucide-react';
import { apiClient, type PublicConfigResponse } from './api/client';
import { DashboardView } from './components/DashboardView';
import { TaskPaneView } from './components/TaskPaneView';
import { detectHostContext, type HostContextInfo } from './office/hostAdapter';

export const App: React.FC = () => {
  const [config, setConfig] = useState<PublicConfigResponse | null>(null);
  const [host, setHost] = useState<HostContextInfo>({
    suite: 'standalone',
    app: 'standalone',
    label: '独立桌面/预览模式'
  });
  const [error, setError] = useState<string | null>(null);
  const [mode, setMode] = useState<'taskpane' | 'dashboard'>(() => {
    const p = new URLSearchParams(window.location.search);
    if (window.location.pathname.startsWith('/dashboard') || p.get('mode') === 'dashboard') {
      return 'dashboard';
    }
    return 'taskpane';
  });
  // Dashboard is a management surface: always available from the tray,
  // even when the page itself is opened by Word / Excel / WPS.
  const isDashboard = mode === 'dashboard';

  const refreshConfig = useCallback(async () => {
    const cfg = await apiClient.getConfig();
    setConfig(cfg);
  }, []);

  // 心跳：页面存活期间保持网关运行；页面关闭后网关即可自行退出。
  // 只有真正嵌在 Word/WPS 里时才发送，避免独立窗口把网关一直吊着。
  useEffect(() => {
    if (!host.embedded || mode === 'dashboard') return;
    let cancelled = false;
    const beat = () => {
      if (cancelled) return;
      void apiClient.pingSession(host.app).catch(() => {});
    };
    beat();
    const timer = window.setInterval(beat, 30_000);
    return () => {
      cancelled = true;
      window.clearInterval(timer);
    };
  }, [host.embedded, host.app, mode]);

  useEffect(() => {
    Promise.all([detectHostContext(), apiClient.getConfig()])
      .then(([hostInfo, cfg]) => {
        setHost(hostInfo);
        setConfig(cfg);
      })
      .catch((err) => {
        setError(err instanceof Error ? err.message : String(err));
      });
  }, []);

  if (error) {
    return (
      <div className="flex h-full items-center justify-center p-6 text-center">
        <div className="max-w-md rounded-2xl border border-rose-200 bg-white p-6 shadow-sm space-y-2">
          <div className="text-sm font-bold text-rose-600">无法连接到 OfficeAI 本地网关</div>
          <p className="text-xs text-slate-600">{error}</p>
          <button
            onClick={() => window.location.reload()}
            className="mt-2 rounded-lg bg-blue-600 px-4 py-1.5 text-xs font-medium text-white"
          >
            重试连接
          </button>
        </div>
      </div>
    );
  }

  if (!config) {
    return (
      <div className="flex h-full items-center justify-center text-slate-500 gap-2 text-sm">
        <Loader2 className="h-4 w-4 animate-spin text-blue-600" />
        <span>正在加载 OfficeAI...</span>
      </div>
    );
  }

  if (isDashboard) {
    return (
      <DashboardView
        config={config}
        onRefreshConfig={refreshConfig}
        onOpenTaskPane={() => {
          setMode('taskpane');
          if (window.location.pathname.startsWith('/dashboard')) {
            window.history.replaceState({}, '', '/');
          }
        }}
      />
    );
  }

  // 嵌入在 Word / Excel / WPS 侧边栏时，直接展示 TaskPaneView。
  // 之前此处错误地对 embedded 返回了“服务未运行”占位，导致面板空白、助手看不到。
  // 网关未运行时，上面的 error 分支已会显示“无法连接”；此处只要 config 已就绪就展示助手。
  return (
    <TaskPaneView
      config={config}
      host={host}
      onRefreshConfig={refreshConfig}
      onOpenDashboard={() => {
        setMode('dashboard');
        if (!window.location.pathname.startsWith('/dashboard')) {
          window.history.replaceState({}, '', '/dashboard');
        }
      }}
    />
  );
};
