import type { AppSettings } from '@officeai/shared';
import { detectOfficeProcesses, type OfficeProcessState } from './process-watch.js';

/** A hosting app currently holding the gateway alive (leak-ref style) */
export interface OfficeSessionInfo {
  app: string;
  /** Agent that reported the session (ws / wps / office / manual) */
  source: string;
  lastSeen: number;
}

export interface LifecycleStatus {
  mode: AppSettings['launchMode'];
  /** Add-ins that currently hold a session (Word / WPS / …) */
  sessions: Array<{ app: string; source: string; idleSeconds: number }>;
  /** How this gateway process was started: add-in, desktop app, cli, unknown */
  launcher: string;
  upMinutes: number;
  lastActivitySeconds: number;
  exitWithOffice: boolean;
  idleExitMinutes: number;
  sessionGraceMinutes: number;
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
  /** key = host app name (winword / wps / …) */
  private readonly sessions = new Map<string, OfficeSessionInfo>();
  private exitRequested = false;
  private readonly launcher: string;
  private readonly detectOffice: () => Promise<OfficeProcessState>;

  constructor(
    private readonly getSettings: () => AppSettings,
    private readonly options: LifecycleOptions
  ) {
    this.now = options.now ?? (() => Date.now());
    this.pollMs = options.pollMs ?? 5_000;
    this.launcher = options.launcher ?? 'unknown';
    this.detectOffice = options.detectOffice ?? detectOfficeProcesses;
    this.startedAt = this.now();
    this.lastActivityAt = this.startedAt;
  }

  touch(): void {
    this.lastActivityAt = this.now();
  }

  /**
   * Called by an add-in page (Word / WPS / PPT) while it holds a task pane
   * open. Two independent ways to keep the gateway alive:
   *   1. the OS process is still running (tasklist based), and
   *   2. at least one session was reported recently (heartbeat based).
   * Using both makes "exit when the office app exits" robust: if a process
   * name is unexpected, the heartbeat still holds the gateway; if the pane
   * crashes without closing, the process check still releases it.
   */
  watch(session: { app: string; source?: string }): void {
    const app = (session.app || 'unknown').toLowerCase();
    this.sessions.set(app, {
      app,
      source: session.source ?? 'unknown',
      lastSeen: this.now()
    });
    this.touch();
  }

  private activeSessions(settings: AppSettings): OfficeSessionInfo[] {
    const now = this.now();
    const graceMs = Math.max(1, settings.sessionGraceMinutes || 2) * 60_000;
    const alive: OfficeSessionInfo[] = [];
    for (const session of this.sessions.values()) {
      if (now - session.lastSeen > graceMs) {
        this.sessions.delete(session.app);
        continue;
      }
      alive.push(session);
    }
    return alive;
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
      sessionGraceMinutes: settings.sessionGraceMinutes,
      sessions: this.activeSessions(settings).map((s) => ({
        app: s.app,
        source: s.source,
        idleSeconds: Math.round((this.now() - s.lastSeen) / 1000)
      })),
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
      const sessions = this.activeSessions(settings);

      if (this.office.running || sessions.length > 0) {
        this.officeSeen = true;
      } else if (settings.exitWithOffice) {
        // Nothing is holding us: no office process and no live task pane.
        // Give a freshly started gateway time for Office to show up.
        const graceMs = this.officeSeen ? 15_000 : 240_000;
        if (this.now() - this.lastActivityAt > graceMs) {
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
