import { execFile } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { OFFICEAI_PROTOCOL_SCHEME } from '@officeai/shared';

const execFileAsync = promisify(execFile);
const __dirname = path.dirname(fileURLToPath(import.meta.url));

export type LaunchTargetKind = 'bat' | 'exe' | 'command' | 'none';

export interface LaunchTarget {
  /** Absolute path or raw command that starts the OfficeAI desktop client / gateway */
  command: string;
  kind: LaunchTargetKind;
  /** Where the target came from, for diagnostics in the dashboard */
  source: 'custom' | 'env' | 'installed-exe' | 'none';
}

function firstExisting(candidates: Array<{ file: string; source: LaunchTarget['source'] }>): LaunchTarget | null {
  for (const candidate of candidates) {
    if (candidate.file && fs.existsSync(candidate.file)) {
      return {
        command: candidate.file,
        kind: candidate.file.toLowerCase().endsWith('.bat') ? 'bat' : 'exe',
        source: candidate.source
      };
    }
  }
  return null;
}

/**
 * Finds how to start OfficeAI on this machine. The result is handed to the
 * WPS add-in (via `/api/office/status`) and used to register the custom URL
 * protocol so Office users can start the gateway with one click.
 */
export function resolveLaunchTarget(options: {
  customCommand?: string;
  dataDir?: string;
} = {}): LaunchTarget {
  const custom = (options.customCommand ?? '').trim();
  if (custom) {
    if (custom.toLowerCase().endsWith('.bat')) {
      return { command: custom, kind: 'bat', source: 'custom' };
    }
    if (custom.toLowerCase().endsWith('.exe') && fs.existsSync(custom)) {
      return { command: custom, kind: 'exe', source: 'custom' };
    }
    return { command: custom, kind: 'command', source: 'custom' };
  }

  const envLauncher = (process.env.OFFICEAI_LAUNCHER ?? '').trim();
  if (envLauncher && fs.existsSync(envLauncher)) {
    return {
      command: envLauncher,
      kind: envLauncher.toLowerCase().endsWith('.bat') ? 'bat' : 'exe',
      source: 'env'
    };
  }

  // On-demand launch is a Windows-only feature (WPS OAAssist / officeai://).
  if (process.platform !== 'win32') {
    return { command: '', kind: 'none', source: 'none' };
  }

  const localAppData = process.env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Local');
  const programFiles = process.env.ProgramFiles || 'C:\\Program Files';
  const programFilesX86 = process.env['ProgramFiles(x86)'] || 'C:\\Program Files (x86)';
  const found = firstExisting([
    // Running inside the packaged desktop app: process.execPath is OfficeAI.exe
    {
      file: process.execPath.toLowerCase().endsWith('officeai.exe') ? process.execPath : '',
      source: 'installed-exe'
    },
    { file: path.join(localAppData, 'Programs', 'OfficeAI', 'OfficeAI.exe'), source: 'installed-exe' },
    { file: path.join(programFiles, 'OfficeAI', 'OfficeAI.exe'), source: 'installed-exe' },
    { file: path.join(programFilesX86, 'OfficeAI', 'OfficeAI.exe'), source: 'installed-exe' }
  ]);

  return found ?? { command: '', kind: 'none', source: 'none' };
}

/**
 * Office (MS Office) add-ins cannot start a local process from their sandbox,
 * so we register the `officeai://` URL protocol to the desktop client once.
 */
export async function registerProtocolHandler(exePath: string): Promise<{ ok: boolean; message: string }> {
  if (process.platform !== 'win32') {
    return { ok: false, message: '当前系统不是 Windows，已跳过 officeai:// 协议注册。' };
  }
  if (!fs.existsSync(exePath)) {
    return { ok: false, message: `未找到可执行文件：${exePath}` };
  }

  const root = `HKCU\\Software\\Classes\\${OFFICEAI_PROTOCOL_SCHEME}`;
  try {
    await execFileAsync('reg.exe', ['add', root, '/ve', '/t', 'REG_SZ', '/d', 'URL:OfficeAI Protocol', '/f']);
    await execFileAsync('reg.exe', ['add', root, '/v', 'URL Protocol', '/t', 'REG_SZ', '/d', '', '/f']);
    await execFileAsync('reg.exe', [
      'add',
      `${root}\\shell\\open\\command`,
      '/ve',
      '/t',
      'REG_SZ',
      '/d',
      `"${exePath}" --launch "%1"`,
      '/f'
    ]);
    return {
      ok: true,
      message: `已注册 officeai:// 协议，双击桌面「OfficeAI 启动」快捷方式即可拉起网关。`
    };
  } catch (err) {
    return {
      ok: false,
      message: `注册 officeai:// 协议失败：${err instanceof Error ? err.message : String(err)}`
    };
  }
}

export async function unregisterProtocolHandler(): Promise<{ ok: boolean; message: string }> {
  if (process.platform !== 'win32') {
    return { ok: false, message: '当前系统不是 Windows。' };
  }
  try {
    await execFileAsync('reg.exe', ['delete', `HKCU\\Software\\Classes\\${OFFICEAI_PROTOCOL_SCHEME}`, '/f']);
    return { ok: true, message: '已移除 officeai:// 协议注册。' };
  } catch {
    return { ok: true, message: '未检测到 officeai:// 协议注册。' };
  }
}

/** Drops a `.url` shortcut (desktop + start menu) that opens the deep link. */
export async function createLauncherShortcut(deepLink: string): Promise<{ ok: boolean; message: string }> {
  if (process.platform !== 'win32') {
    return { ok: false, message: '当前系统不是 Windows，已跳过快捷方式创建。' };
  }
  const appData = process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming');
  const targets = [
    path.join(path.join(os.homedir(), 'Desktop'), 'OfficeAI 启动.url'),
    path.join(appData, 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'OfficeAI 启动.url')
  ];
  const written: string[] = [];
  for (const target of targets) {
    try {
      fs.mkdirSync(path.dirname(target), { recursive: true });
      fs.writeFileSync(target, `[InternetShortcut]\r\nURL=${deepLink}\r\nIconIndex=0\r\n`, 'utf8');
      written.push(target);
    } catch {
      // ignore unwritable locations
    }
  }
  return written.length > 0
    ? { ok: true, message: `已创建一键启动快捷方式：${written.join('、')}` }
    : { ok: false, message: '创建一键启动快捷方式失败。' };
}

export async function removeLauncherShortcut(): Promise<void> {
  const appData = process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming');
  for (const target of [
    path.join(os.homedir(), 'Desktop', 'OfficeAI 启动.url'),
    path.join(appData, 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'OfficeAI 启动.url')
  ]) {
    try {
      fs.rmSync(target, { force: true });
    } catch {
      // ignore
    }
  }
}
