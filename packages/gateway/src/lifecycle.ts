import type { AppSettings } from '@officeai/shared';
import { detectOfficeProcesses, type OfficeProcessState } from './process-watch.js';

export interface LifecycleStatus {
  mode: AppSettings['launchMode'];
  /** How this gateway process was started: add-in, desktop app, cli, unknown */
  launcher: string;
  upMinutes: number;
  lastActivitySeconds: number;
  exitWithOffice: boolean;
  idleExitMinutes: number;
  office: OfficeProcessState;
}

export interface LifecycleOptions {
  /** Declares who started this process, read from `--launched-by` */
  launcher?: string;
  /** Called when the gateway decides it is no longer needed */
  onExitRequested: (reason: string) => void;
  pollMs?: number;
  now?: () => number;
  /** Injectable process probe (tests / non-Windows) */
  detectOffice?: () => Promise<OfficeProcessState>;
}

/**
 * Implements the "only run while Office is open" part of on-demand startup:
 *  - starts a grace period so a freshly launched gateway survives the gap
 *    between the click in WPS and the office process actually starting
 *  - exits once every Office/WPS host has been closed (`exitWithOffice`)
 *  - exits after `idleExitMinutes` without any gateway request
 */
export class LifecycleMonitor {
  private timer: NodeJS.Timeout | null = null;
  private lastActivityAt: number;
  private readonly startedAt: number;
  private readonly pollMs: number;
  private readonly now: () => number;
  private office: OfficeProcessState = { checked: false, running: true, processes: [] };
  private officeSeen = false;
  private checking = false;
  private exitRequested = false;
  private readonly launcher: string;
  private readonly detectOffice: () => Promise<OfficeProcessState>;

  constructor(
    private readonly getSettings: () => AppSettings,
    private readonly options: LifecycleOptions
  ) {
    this.now = options.now ?? (() => Date.now());
    this.pollMs = options.pollMs ?? 20_000;
    this.launcher = options.launcher ?? 'unknown';
    this.detectOffice = options.detectOffice ?? detectOfficeProcesses;
    this.startedAt = this.now();
    this.lastActivityAt = this.startedAt;
  }

  touch(): void {
    this.lastActivityAt = this.now();
  }

  start(): void {
    if (this.timer) return;
    this.timer = setInterval(() => {
      void this.check();
    }, this.pollMs);
    this.timer.unref?.();
    void this.check();
  }

  stop(): void {
    if (this.timer) {
      clearInterval(this.timer);
      this.timer = null;
    }
  }

  getStatus(): LifecycleStatus {
    const settings = this.getSettings();
    return {
      mode: settings.launchMode,
      launcher: this.launcher,
      upMinutes: Math.round((this.now() - this.startedAt) / 60_000),
      lastActivitySeconds: Math.round((this.now() - this.lastActivityAt) / 1000),
      exitWithOffice: settings.exitWithOffice,
      idleExitMinutes: settings.idleExitMinutes,
      office: this.office
    };
  }

  private requestExit(reason: string): void {
    if (this.exitRequested) return;
    this.exitRequested = true;
    this.stop();
    this.options.onExitRequested(reason);
  }

  private async check(): Promise<void> {
    if (this.checking || this.exitRequested) return;
    const settings = this.getSettings();
    if (settings.launchMode !== 'on-demand') return;
    // Auto-exit relies on Windows process detection; on other platforms the
    // gateway would otherwise kill itself during development.
    if (process.platform !== 'win32' && !this.options.detectOffice) return;

    this.checking = true;
    try {
      this.office = await this.detectOffice();
      if (this.office.running) {
        this.officeSeen = true;
      } else if (this.office.checked) {
        // Give Office/WPS time to show up before treating "nothing running" as final
        const graceMs = this.officeSeen ? 90_000 : 240_000;
        const quietFor = this.now() - this.lastActivityAt;
        if (settings.exitWithOffice && quietFor > graceMs) {
          this.requestExit('所有 Office / WPS 窗口均已关闭');
          return;
        }
      }

      if (settings.idleExitMinutes > 0) {
        const idleMs = this.now() - this.lastActivityAt;
        if (idleMs > settings.idleExitMinutes * 60_000) {
          this.requestExit(`已 ${settings.idleExitMinutes} 分钟无任何请求`);
        }
      }
    } finally {
      this.checking = false;
    }
  }
}
