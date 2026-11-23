import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);

/** File-name (without extension) of every app that hosts an OfficeAI add-in. */
const OFFICE_PROCESS_NAMES = new Set([
  'winword',
  'excel',
  'powerpnt',
  'wps',
  'et',
  'wpp',
  'wpsoffice'
]);

export interface OfficeProcessState {
  /** Anything was observed at least once since the last check */
  checked: boolean;
  /** Office / WPS is currently running */
  running: boolean;
  processes: string[];
}

export function isOfficeProcessName(rawName: string): boolean {
  const name = rawName
    .trim()
    .toLowerCase()
    .replace(/\.exe$/, '');
  if (OFFICE_PROCESS_NAMES.has(name)) return true;
  // WPS ships multi-process helpers such as `wpscloudsvr` / `wpsupdate` — only
  // the main binaries (wps/et/wpp) count, so we deliberately do not prefix-match.
  return false;
}

/**
 * Windows-only: list running Office/WPS main processes. On other platforms we
 * report "unknown" (running = true) so that idle auto-exit never fires on dev
 * machines.
 */
export async function detectOfficeProcesses(): Promise<OfficeProcessState> {
  if (process.platform !== 'win32') {
    return { checked: false, running: true, processes: [] };
  }

  try {
    const { stdout } = await execFileAsync(
      'tasklist.exe',
      ['/fo', 'csv', '/nh'],
      { timeout: 8000, maxBuffer: 8 * 1024 * 1024 }
    );
    const processes: string[] = [];
    for (const line of stdout.split(/\r?\n/)) {
      const match = /^"([^"]+)"/.exec(line.trim());
      if (!match) continue;
      if (isOfficeProcessName(match[1])) {
        processes.push(match[1].toLowerCase().replace(/\.exe$/, ''));
      }
    }
    return {
      checked: true,
      running: processes.length > 0,
      processes: [...new Set(processes)]
    };
  } catch {
    // Never let a failing tasklist kill the gateway; assume Office is running.
    return { checked: false, running: true, processes: [] };
  }
}
