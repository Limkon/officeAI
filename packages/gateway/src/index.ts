export {
  ConfigStore,
  resolveDataDir,
  normalizeSettings,
  type PublicAppConfig
} from './config-store.js';
export { ensureLocalhostCert, type CertPair } from './certs.js';
export {
  startGateway,
  resolveGatewayPaths,
  type RunningGateway,
  type GatewayPaths
} from './server.js';
export {
  WindowsOfficeIntegrator,
  type IntegrationActionResult
} from './windows/office-integrator.js';
export { probeLocalRuntimes, resolveProviderAdapter } from './providers/index.js';
export {
  resolveLaunchTarget,
  registerProtocolHandler,
  unregisterProtocolHandler,
  createLauncherShortcut,
  removeLauncherShortcut,
  type LaunchTarget,
  type LaunchTargetKind
} from './desktop-tools.js';
export { LifecycleMonitor, type LifecycleStatus, type LifecycleOptions } from './lifecycle.js';
export { detectOfficeProcesses, isOfficeProcessName, type OfficeProcessState } from './process-watch.js';
