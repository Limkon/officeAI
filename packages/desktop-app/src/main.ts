import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow, Menu, Tray, dialog, nativeImage } from 'electron';
import { parseGatewayLaunchUrl } from '@officeai/shared';
import {
  ConfigStore,
  WindowsOfficeIntegrator,
  ensureLocalhostCert,
  resolveGatewayPaths,
  resolveLaunchTarget,
  startGateway,
  type RunningGateway
} from '@officeai/gateway';

let mainWindow: BrowserWindow | null = null;
let tray: Tray | null = null;
let gateway: RunningGateway | null = null;
let isQuitting = false;
let gatewayExitReason: string | null = null;

/** `officeai://activate/<encoded url>` (or `--launch <url>`) opened by Office */
function resolveLaunchArguments(argv: string[]): { deepLink?: string; appHint?: string } {
  const launchIdx = argv.indexOf('--launch');
  const deepLink = launchIdx >= 0 ? argv[launchIdx + 1] : argv.find((a) => a.startsWith('officeai://'));
  const appIdx = argv.indexOf('--app');
  const appHint = appIdx >= 0 ? argv[appIdx + 1] : undefined;
  return { deepLink, appHint };
}

const launchArgs = resolveLaunchArguments(process.argv.slice(1));

/** 安装程序（NSIS）在安装完成后以 `OfficeAI.exe --integrate` 无头调用，自动完成全部集成 */
const integrateRequested = process.argv.slice(1).includes('--integrate');

/**
 * 无头全量集成：证书信任 + Microsoft Office + WPS + 唤醒器。
 * 由安装程序在安装完成后调用，用户无需手动点按钮、手动装证书。
 * 不弹窗口、不启动网关服务、不占用单例；全部步骤幂等，可重复执行。
 */
async function runHeadlessIntegration(): Promise<void> {
  const assetsDir = resolveAssetsDir();
  if (fs.existsSync(path.join(assetsDir, 'addin-ui'))) {
    process.env.OFFICEAI_ASSETS_DIR = assetsDir;
  }
  const store = new ConfigStore();
  const launchTarget = resolveLaunchTarget({
    customCommand: store.getConfig().settings.gatewayLaunchCommand,
    dataDir: store.dataDir
  });
  ensureLocalhostCert(store.dataDir);
  const paths = resolveGatewayPaths();
  const integrator = new WindowsOfficeIntegrator(
    store.dataDir,
    store.getConfig().settings.httpPort,
    store.getConfig().settings.httpsPort,
    paths.assetsRoot,
    launchTarget.command
  );
  const res = await integrator.installAll();
  console.log(`[OfficeAI:integrate] ok=${res.ok} ${res.message}`);
  res.details?.forEach((d: string) => console.log(`  - ${d}`));
  if (!res.ok) {
    // 安装器 ExecWait 会等待此进程退出；即使部分失败也返回 0 让安装继续，
    // 失败细节已通过日志与控制中心可见，避免安装被判失败。
    console.warn('[OfficeAI:integrate] 部分集成未完成，请在控制中心查看详情后重试一键安装。');
  }
}

function resolveAssetsDir(): string {
  // Packaged Electron puts extraResources under process.resourcesPath/assets
  const packagedAssets = path.join(process.resourcesPath || '', 'assets');
  if (fs.existsSync(path.join(packagedAssets, 'addin-ui'))) {
    return packagedAssets;
  }
  return path.resolve(__dirname, '..', '..');
}

function resolveTrayIcon(): Electron.NativeImage {
  const candidates = [
    path.join(__dirname, '..', 'assets', 'icon-32.png'),
    path.join(process.resourcesPath || '', 'assets', 'wps-addin', 'assets', 'icon-32.png')
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) {
      return nativeImage.createFromPath(p);
    }
  }
  return nativeImage.createEmpty();
}

async function createMainWindow(urlPath = '/dashboard'): Promise<void> {
  if (!gateway) return;
  const targetUrl = `${gateway.httpUrl}${urlPath}`;

  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.show();
    mainWindow.focus();
    await mainWindow.loadURL(targetUrl);
    return;
  }

  mainWindow = new BrowserWindow({
    width: 1120,
    height: 760,
    minWidth: 420,
    minHeight: 560,
    title: 'OfficeAI 智能办公助手',
    icon: resolveTrayIcon(),
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.on('close', (event) => {
    // On-demand mode must not leave a hidden process behind forever: closing
    // the window hands control to the lifecycle monitor instead of hiding.
    const mode = gateway?.store.getConfig().settings.launchMode;
    if (!isQuitting && tray && mode !== 'on-demand') {
      event.preventDefault();
      mainWindow?.hide();
      return;
    }
    if (!isQuitting && mode === 'on-demand') {
      isQuitting = true;
    }
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
    if (isQuitting) {
      app.quit();
    }
  });

  await mainWindow.loadURL(targetUrl);
}

/** Hide the taskbar/dock presence so OfficeAI behaves like a background service */
function hideToTray(): void {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.hide();
  }
  if (process.platform === 'darwin') {
    app.dock?.hide();
  }
}

function createTray(): void {
  if (!gateway) return;
  tray = new Tray(resolveTrayIcon());
  const mode = gateway.store.getConfig().settings.launchMode;
  tray.setToolTip(
    mode === 'on-demand'
      ? 'OfficeAI 按需运行中（关闭所有 Office / WPS 后会自动退出）'
      : 'OfficeAI 网关运行中 (Office / WPS AI 助手)'
  );

  const menu = Menu.buildFromTemplate([
    {
      label: '打开模型与集成配置中心',
      click: () => void createMainWindow('/dashboard')
    },
    {
      label: '打开独立 AI 助手窗口',
      click: () => void createMainWindow('/')
    },
    { type: 'separator' },
    {
      label: '一键安装到 WPS Office（本地离线加载项）',
      click: async () => {
        if (!gateway) return;
        const res = await gateway.integrator.installWps();
        dialog.showMessageBox({
          type: res.ok ? 'info' : 'warning',
          title: 'WPS Office 集成',
          message: res.message
        });
      }
    },
    {
      label: '一键安装到 Microsoft Office',
      click: async () => {
        if (!gateway) return;
        const res = await gateway.integrator.installMsOffice();
        dialog.showMessageBox({
          type: res.ok ? 'info' : 'warning',
          title: 'Microsoft Office 集成',
          message: res.message
        });
      }
    },
    { type: 'separator' },
    {
      label: '退出 OfficeAI',
      click: () => {
        isQuitting = true;
        app.quit();
      }
    }
  ]);

  tray.setContextMenu(menu);
  tray.on('double-click', () => void createMainWindow('/dashboard'));
}

// --integrate 为 NSIS 安装器的无头调用：必须在 singleInstance 锁之前处理，
// 且不与正常桌面进程互斥（幂等，可与已运行实例并行）。直接用 process 退出，
// 不走 Electron 窗口/托盘生命周期，避免“闪退”与单例抢占。
if (integrateRequested) {
  // 完全绕过 Electron 单例与窗口初始化
  runHeadlessIntegration()
    .then(() => {
      // 确保日志落盘后再退出；用 app.exit 而非 process.exit 以便 Electron 清理
      try {
        if (app.isReady()) app.quit();
      } catch {}
      setTimeout(() => process.exit(0), 300);
    })
    .catch((err) => {
      console.error('[OfficeAI:integrate] 集成失败:', err);
      try {
        if (app.isReady()) app.quit();
      } catch {}
      setTimeout(() => process.exit(1), 300);
    });
} else {
  const singleLock = app.requestSingleInstanceLock();
  if (!singleLock) {
    // A gateway is already running: let the existing instance handle the deep
    // link (e.g. WPS asked to start the assistant while it is already up).
    app.quit();
  } else {
  app.on('second-instance', (_event, argv) => {
    const { deepLink } = resolveLaunchArguments(argv);
    if (deepLink) {
      const target = parseGatewayLaunchUrl(deepLink);
      if (target) {
        void createMainWindow('/dashboard');
        return;
      }
    }
    void createMainWindow('/dashboard');
  });

  app.whenReady().then(async () => {
    try {
      const assetsDir = resolveAssetsDir();
      if (fs.existsSync(path.join(assetsDir, 'addin-ui'))) {
        process.env.OFFICEAI_ASSETS_DIR = assetsDir;
      }

      // Declare who started us so the lifecycle monitor and the dashboard can
      // show "started by WPS / by the start menu".
      const launcher = launchArgs.deepLink ? 'addin' : 'desktop';
      const target = launchArgs.deepLink ? parseGatewayLaunchUrl(launchArgs.deepLink) : undefined;

      gateway = await startGateway({
        launcher,
        onExitRequested: (reason) => {
          gatewayExitReason = reason;
          isQuitting = true;
          app.quit();
        }
      });

      // The gateway may have fallen back to another port; report it back to the
      // deep link target so we never land on a dead page.
      if (target && target !== gateway.httpUrl) {
        console.log(`[OfficeAI] 端口不同，期望 ${target}，实际 ${gateway.httpUrl}`);
      }

      createTray();

      const mode = gateway.store.getConfig().settings.launchMode;
      // 展示窗口：桌面启动应保持可见，避免“闪退”。只有由加载项唤起时才隐藏到托盘。
      if (launcher === 'addin') {
        // Started from WPS/Office: stay invisible, user is waiting inside Office
        hideToTray();
      } else if (mode === 'manual') {
        hideToTray();
      } else {
        await createMainWindow('/dashboard');
      }

      // 兜底集成改为后台异步，不阻塞窗口/托盘的展示，避免首次启动长时间白屏或“闪退”感。
      // 若唤醒器未安装或证书未信任（例如 NSIS 的 --integrate 被跳过/失败，或旧版升级），自动补做。
      void (async () => {
        try {
          const status = await gateway!.integrator.getStatus();
          if (!status.waker?.installed || !status.certTrusted) {
            const res = await gateway!.integrator.installAll();
            console.log(`[OfficeAI] 启动时后台自动集成：ok=${res.ok} ${res.message}`);
            res.details?.forEach((d: string) => console.log(`  - ${d}`));
          }
        } catch (err) {
          console.warn('[OfficeAI] 启动时后台自动集成跳过:', err);
        }
      })();
    } catch (err) {
      dialog.showErrorBox(
        'OfficeAI 启动失败',
        `本地网关启动失败，请检查端口 18080 / 18443 是否被占用：\n${err instanceof Error ? err.message : String(err)}`
      );
      isQuitting = true;
      app.quit();
    }
  });

  app.on('before-quit', async () => {
    isQuitting = true;
    if (gateway) {
      await gateway.close().catch(() => {});
    }
    if (gatewayExitReason) {
      console.log(`[OfficeAI] 自动退出：${gatewayExitReason}`);
    }
  });

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin' && !tray) {
      app.quit();
    }
  });
  }
}
