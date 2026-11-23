import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow, Menu, Tray, dialog, nativeImage } from 'electron';
import { parseGatewayLaunchUrl } from '@officeai/shared';
import { startGateway, type RunningGateway } from '@officeai/gateway';

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
      if (launcher === 'addin') {
        // Started from WPS: stay invisible, user is waiting inside WPS
        hideToTray();
      } else if (mode === 'on-demand') {
        await createMainWindow('/dashboard');
        hideToTray();
      } else if (mode === 'manual') {
        hideToTray();
      } else {
        await createMainWindow('/dashboard');
      }
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
