import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow, Menu, Tray, dialog, nativeImage } from 'electron';
import { startGateway, type RunningGateway } from '@officeai/gateway';

let mainWindow: BrowserWindow | null = null;
let tray: Tray | null = null;
let gateway: RunningGateway | null = null;
let isQuitting = false;

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
    if (!isQuitting && tray) {
      event.preventDefault();
      mainWindow?.hide();
    }
  });

  await mainWindow.loadURL(targetUrl);
}

function createTray(): void {
  if (!gateway) return;
  tray = new Tray(resolveTrayIcon());
  tray.setToolTip('OfficeAI 网关运行中 (Office / WPS AI 助手)');

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
      label: '一键安装到 WPS Office',
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
  app.quit();
} else {
  app.on('second-instance', () => {
    void createMainWindow('/dashboard');
  });

  app.whenReady().then(async () => {
    const assetsDir = resolveAssetsDir();
    if (fs.existsSync(path.join(assetsDir, 'addin-ui'))) {
      process.env.OFFICEAI_ASSETS_DIR = assetsDir;
    }

    gateway = await startGateway();
    createTray();
    await createMainWindow('/dashboard');
  });

  app.on('before-quit', async () => {
    isQuitting = true;
    if (gateway) {
      await gateway.close().catch(() => {});
    }
  });

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin' && !tray) {
      app.quit();
    }
  });
}
