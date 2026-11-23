import fs from 'node:fs';
import path from 'node:path';
import { app } from 'electron';
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
    console.warn('[OfficeAI:integrate] 部分集成未完成，请在Office面板的设置中查看详情后重试一键安装。');
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

// --integrate 为 NSIS 安装器的无头调用：必须在 singleInstance 锁之前处理，
// 且不与正常桌面进程互斥（幂等，可与已运行实例并行）。直接退出，
// 不走任何窗口/托盘生命周期（已移除托盘图标，按用户要求无单独界面）。
if (integrateRequested) {
  runHeadlessIntegration()
    .then(() => {
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
    // 已有网关在运行：让既有实例处理（WPS/Office 唤起时）
    app.quit();
  } else {
    // 隐藏 Dock/任务栏图标，完全作为后台服务，参考其他 Office 助手（无托盘、面板内完成全部设置）
    if (process.platform === 'darwin') {
      app.dock?.hide();
    }

    app.on('second-instance', (_event, argv) => {
      const { deepLink } = resolveLaunchArguments(argv);
      if (deepLink) {
        const target = parseGatewayLaunchUrl(deepLink);
        if (target) {
          console.log(`[OfficeAI] 收到二次唤起: ${target}`);
        }
      }
      // 无窗口/托盘，不再弹 Dashboard；Office 面板内已可完成全部设置
    });

    app.whenReady().then(async () => {
      try {
        const assetsDir = resolveAssetsDir();
        if (fs.existsSync(path.join(assetsDir, 'addin-ui'))) {
          process.env.OFFICEAI_ASSETS_DIR = assetsDir;
        }

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

        if (target && target !== gateway.httpUrl) {
          console.log(`[OfficeAI] 端口不同，期望 ${target}，实际 ${gateway.httpUrl}`);
        }

        // 无托盘、无独立窗口：所有设置已整合到 Office 侧边栏的「配置中心」
        // （TaskPane 右上角“配置中心”按钮 ↔ DashboardView），此处不再 createTray / createMainWindow
        console.log(`[OfficeAI] 网关已启动 ${gateway.httpUrl} (后台常驻，无托盘图标)`);

        // 后台兜底集成（证书/Office/WPS/唤醒器），失败不影响网关运行，详情在 Office 面板可见
        void (async () => {
          try {
            const status = await gateway!.integrator.getStatus();
            if (!status.waker?.installed || !status.certTrusted) {
              const res = await gateway!.integrator.installAll();
              console.log(`[OfficeAI] 后台自动集成：ok=${res.ok} ${res.message}`);
              res.details?.forEach((d: string) => console.log(`  - ${d}`));
            }
          } catch (err) {
            console.warn('[OfficeAI] 后台自动集成跳过:', err);
          }
        })();
      } catch (err) {
        console.error(
          `[OfficeAI] 本地网关启动失败，请检查端口 18080 / 18443 是否被占用：${
            err instanceof Error ? err.message : String(err)
          }`
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

    // 无窗口应用：关闭所有窗口也不退出，保持后台常驻，由 Office 唤醒器/生命周期控制
    app.on('window-all-closed', () => {
      // 刻意不 quit，保持后台服务
    });
  }
}
