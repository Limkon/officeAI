#!/usr/bin/env node
import {
  createLauncherShortcut,
  registerProtocolHandler,
  resolveLaunchTarget,
  unregisterProtocolHandler
} from './desktop-tools.js';
import { ensureLocalhostCert } from './certs.js';
import { ConfigStore } from './config-store.js';
import { resolveGatewayPaths, startGateway } from './server.js';
import { WindowsOfficeIntegrator } from './windows/office-integrator.js';
import { buildGatewayLaunchUrl, parseGatewayLaunchUrl } from '@officeai/shared';

interface CliFlags {
  positional: string[];
  flags: Map<string, string | true>;
}

function parseArgs(argv: string[]): CliFlags {
  const positional: string[] = [];
  const flags = new Map<string, string | true>();
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg.startsWith('--')) {
      positional.push(arg);
      continue;
    }
    const [rawKey, inlineValue] = arg.slice(2).split('=');
    if (inlineValue !== undefined) {
      flags.set(rawKey, inlineValue);
      continue;
    }
    const next = argv[i + 1];
    if (next && !next.startsWith('--')) {
      flags.set(rawKey, next);
      i += 1;
    } else {
      flags.set(rawKey, true);
    }
  }
  return { positional, flags };
}

function flagString(flags: CliFlags['flags'], key: string): string | undefined {
  const value = flags.get(key);
  return typeof value === 'string' ? value : undefined;
}

async function main() {
  const { positional, flags } = parseArgs(process.argv.slice(2));
  const cmd = positional[0] || 'serve';
  const launchedBy = flagString(flags, 'launched-by') ?? 'cli';
  // `--launch officeai://activate/<encoded url>` is what the officeai://
  // protocol handler passes in when an Office add-in asks for the gateway.
  const launchDeepLink = flagString(flags, 'launch');

  const store = new ConfigStore();
  const launchTarget = resolveLaunchTarget({
    customCommand: store.getConfig().settings.gatewayLaunchCommand,
    dataDir: store.dataDir
  });

  const integrationCommands = [
    'install-wps',
    'install-office',
    'install-waker',
    'install-all',
    'status',
    'install-protocol',
    'uninstall-protocol'
  ];
  if (integrationCommands.includes(cmd)) {
    ensureLocalhostCert(store.dataDir);
    const paths = resolveGatewayPaths();
    const integrator = new WindowsOfficeIntegrator(
      store.dataDir,
      store.getConfig().settings.httpPort,
      store.getConfig().settings.httpsPort,
      paths.assetsRoot,
      launchTarget.command
    );

    if (cmd === 'status') {
      const status = await integrator.getStatus();
      console.log(JSON.stringify(status, null, 2));
      return;
    }

    if (cmd === 'install-protocol' || cmd === 'uninstall-protocol') {
      const httpPort = store.getConfig().settings.httpPort;
      const gatewayUrl = `http://127.0.0.1:${httpPort}`;
      const exe =
        launchTarget.kind === 'exe'
          ? launchTarget.command
          : '';

      if (cmd === 'uninstall-protocol') {
        const res = await unregisterProtocolHandler();
        console.log('[Protocol]', res.message);
        return;
      }

      if (!exe) {
        console.log(
          '[Protocol] 未找到 OfficeAI.exe，只能注册 officeai:// 协议到 Start-OfficeAI.bat 所在目录。\n' +
            '  请在配置中心「启动方式」中填写 OfficeAI.exe 的完整路径后重试。'
        );
        return;
      }

      const res = await registerProtocolHandler(exe);
      console.log('[Protocol]', res.message);
      if (res.ok) {
        const shortcut = await createLauncherShortcut(buildGatewayLaunchUrl(gatewayUrl, 'office'));
        console.log('[Shortcut]', shortcut.message);
      }
      return;
    }

    if (cmd === 'install-wps' || cmd === 'install-all') {
      const res = await integrator.installWps();
      console.log('[WPS]', res.message);
      res.details?.forEach((d) => console.log('  -', d));
    }

    if (cmd === 'install-waker' || cmd === 'install-all') {
      const res = await integrator.installOfficeWaker();
      console.log('[Waker]', res.message);
      res.details?.forEach((d) => console.log('  -', d));
    }

    if (cmd === 'install-office' || cmd === 'install-all') {
      const res = await integrator.installMsOffice();
      console.log('[Office]', res.message);
      res.details?.forEach((d) => console.log('  -', d));
    }
    return;
  }

  const gw = await startGateway({
    launcher: launchedBy,
    onExitRequested: (reason) => {
      console.log(`[OfficeAI] ${reason}，网关自动退出。`);
      void gw.close().finally(() => process.exit(0));
    }
  });

  console.log(`OfficeAI Gateway 已启动:`);
  console.log(`  控制中心 (HTTP):  ${gw.httpUrl}/dashboard`);
  if (gw.httpsUrl) {
    console.log(`  Office加载项 (HTTPS): ${gw.httpsUrl}`);
  }
  console.log(`  数据目录: ${gw.store.dataDir}`);
  console.log(`  启动方式: ${gw.store.getConfig().settings.launchMode}`);
  if (launchTarget.command) {
    console.log(`  按需启动命令: ${launchTarget.command}`);
  }

  if (launchDeepLink) {
    const target = parseGatewayLaunchUrl(launchDeepLink);
    console.log(`  由 ${launchedBy} 唤起: ${target ?? launchDeepLink}`);
  }

  const shutdown = async () => {
    await gw.close();
    process.exit(0);
  };
  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
}

main().catch((err) => {
  console.error('OfficeAI Gateway 启动失败:', err);
  process.exit(1);
});
