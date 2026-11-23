#!/usr/bin/env node
import { ensureLocalhostCert } from './certs.js';
import { ConfigStore } from './config-store.js';
import { resolveGatewayPaths, startGateway } from './server.js';
import { WindowsOfficeIntegrator } from './windows/office-integrator.js';

async function main() {
  const args = process.argv.slice(2);
  const cmd = args[0] || 'serve';

  if (cmd === 'install-wps' || cmd === 'install-office' || cmd === 'install-all' || cmd === 'status') {
    const store = new ConfigStore();
    ensureLocalhostCert(store.dataDir);
    const paths = resolveGatewayPaths();
    const integrator = new WindowsOfficeIntegrator(
      store.dataDir,
      store.getConfig().settings.httpPort,
      store.getConfig().settings.httpsPort,
      paths.assetsRoot
    );

    if (cmd === 'status') {
      const status = await integrator.getStatus();
      console.log(JSON.stringify(status, null, 2));
      return;
    }

    if (cmd === 'install-wps' || cmd === 'install-all') {
      const res = await integrator.installWps();
      console.log('[WPS]', res.message);
      res.details?.forEach((d) => console.log('  -', d));
    }

    if (cmd === 'install-office' || cmd === 'install-all') {
      const res = await integrator.installMsOffice();
      console.log('[Office]', res.message);
      res.details?.forEach((d) => console.log('  -', d));
    }
    return;
  }

  const gw = await startGateway();
  console.log(`OfficeAI Gateway 已启动:`);
  console.log(`  控制中心 (HTTP):  ${gw.httpUrl}/dashboard`);
  if (gw.httpsUrl) {
    console.log(`  Office加载项 (HTTPS): ${gw.httpsUrl}`);
  }
  console.log(`  数据目录: ${gw.store.dataDir}`);

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
