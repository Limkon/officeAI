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
    let cancelled = false;
    // 给 detectHostContext 与 /api/config 都加超时，避免无限转圈
    const withTimeout = <T,>(p: Promise<T>, ms: number, label: string): Promise<T> =>
      Promise.race([
        p,
        new Promise<never>((_, reject) =>
          setTimeout(() => reject(new Error(`${label} 超时（${ms}ms），请检查本地网关是否已启动（端口 18080/18443）`)), ms)
        )
      ]);

    withTimeout(detectHostContext(), 2500, 'Office 环境检测')
      .then((hostInfo) => {
        if (!cancelled) setHost(hostInfo);
        return withTimeout(apiClient.getConfig(), 6000, '网关连接');
      })
      .then((cfg) => {
        if (!cancelled) setConfig(cfg);
      })
      .catch((err) => {
        if (!cancelled) setError(err instanceof Error ? err.message : String(err));
      });

    // 兜底：8 秒仍未就绪则强制提示，避免用户一直看到转圈
    const fallback = setTimeout(() => {
      if (!cancelled && !config && !error) {
        setError('加载超时：未能在 8 秒内连接到 OfficeAI 本地网关。请确认 OfficeAI 已安装并已重启 Word/WPS，或点击“配置中心”查看运行状态。');
      }
    }, 8000);

    return () => {
      cancelled = true;
      clearTimeout(fallback);
    };
  }, [config, error]);

  if (error) {
    const gwHint = (() => {
      try {
        const u = new URL(window.location.href);
        return `${u.protocol}//${u.host}`;
      } catch {
        return 'https://127.0.0.1:18443 / http://127.0.0.1:18080';
      }
    })();
    const isHttps = window.location.protocol === 'https:';
    const httpFallback = (() => {
      try {
        const u = new URL(window.location.href);
        // 保留路径与查询，仅切换为 http 18080
        u.protocol = 'http:';
        u.port = '18080';
        return u.toString();
      } catch {
        return 'http://127.0.0.1:18080/?host=office';
      }
    })();
    return (
      <div className="flex h-full items-center justify-center p-6 text-center">
        <div className="max-w-md rounded-2xl border border-rose-200 bg-white p-6 shadow-sm space-y-3">
          <div className="text-sm font-bold text-rose-600">无法连接到 OfficeAI 本地网关</div>
          <p className="text-xs text-slate-600 break-words">{error}</p>
          <p className="text-[11px] text-slate-500 break-words">当前页面地址：{window.location.href}<br />网关地址：{gwHint}/api/health</p>
          {isHttps && (
            <p className="text-[11px] text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-left">
              提示：Office 任务窗格使用 https（需信任自签名证书）。若证书未信任会导致空白/超时，可点击“尝试 http”或在配置中心重新“一键安装”以信任证书。
            </p>
          )}
          <div className="flex flex-wrap gap-2 justify-center pt-1">
            <button
              onClick={() => window.location.reload()}
              className="rounded-lg bg-blue-600 px-4 py-1.5 text-xs font-medium text-white"
            >
              重试连接
            </button>
            {isHttps && (
              <button
                onClick={() => {
                  window.location.href = httpFallback;
                }}
                className="rounded-lg border border-amber-300 bg-amber-50 px-4 py-1.5 text-xs font-medium text-amber-800"
              >
                尝试 http
              </button>
            )}
            <button
              onClick={() => {
                // 尝试通过 officeai:// 协议唤起本地服务（若已注册）
                const u = new URL(window.location.href);
                const launchUrl = `officeai://activate/${encodeURIComponent(`${u.protocol}//${u.host}/start?app=office`)}`;
                window.location.href = launchUrl;
                setTimeout(() => window.location.reload(), 1500);
              }}
              className="rounded-lg border border-slate-200 bg-white px-4 py-1.5 text-xs font-medium text-slate-700"
            >
              一键唤起
            </button>
            <button
              onClick={() => {
                setError(null);
                setMode('dashboard');
                if (!window.location.pathname.startsWith('/dashboard')) {
                  window.history.replaceState({}, '', '/dashboard');
                }
                window.location.reload();
              }}
              className="rounded-lg border border-slate-200 bg-white px-4 py-1.5 text-xs font-medium text-slate-700"
            >
              打开配置中心
            </button>
          </div>
          <p className="text-[11px] text-slate-400">提示：首次安装后需重启 Word/Excel/WPS；若仍失败，请从开始菜单搜索“OfficeAI”手动启动，或在配置中心执行“一键安装”。</p>
        </div>
      </div>
    );
  }

  if (!config) {
    return (
      <div className="flex h-full flex-col items-center justify-center text-slate-500 gap-3 text-sm p-6 text-center">
        <div className="flex items-center gap-2">
          <Loader2 className="h-4 w-4 animate-spin text-blue-600" />
          <span>正在加载 OfficeAI...</span>
        </div>
        <span className="text-xs text-slate-400">若持续转圈，请稍后重试或点击配置中心查看网关状态</span>
        <button
          onClick={() => {
            setMode('dashboard');
            if (!window.location.pathname.startsWith('/dashboard')) {
              window.history.replaceState({}, '', '/dashboard');
            }
          }}
          className="text-xs text-blue-600 underline"
        >
          前往配置中心
        </button>
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
