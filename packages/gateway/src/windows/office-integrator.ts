import { execFile } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { promisify } from 'node:util';
import { buildGatewayLaunchUrl, type WindowsOfficeStatus } from '@officeai/shared';
import { ensureLocalhostCert } from '../certs.js';
import { detectOfficeProcesses, type OfficeProcessState } from '../process-watch.js';
import {
  createLauncherShortcut,
  registerProtocolHandler
} from '../desktop-tools.js';

const execFileAsync = promisify(execFile);

const OFFICE_ADDIN_ID = 'b8d4c3e2-71a0-4f88-9c3d-0f1ce2026a00';
const OFFICE_DEV_REG_KEY = 'HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer';

export interface IntegrationActionResult {
  ok: boolean;
  message: string;
  details?: string[];
}

export class WindowsOfficeIntegrator {
  constructor(
    private readonly dataDir: string,
    private readonly httpPort: number,
    private readonly httpsPort: number,
    private readonly assetsRoot: string,
    /** How OfficeAI itself is started (used to wire on-demand launch) */
    private readonly launchCommand = '',
    /** Optional lifecycle snapshot provider (sessions / uptime) */
    private readonly getLifecycleStatus?: () => {
      mode: string;
      launcher: string;
      upMinutes: number;
      lastActivitySeconds: number;
      sessions: Array<{ app: string; source: string; idleSeconds: number }>;
    } | null
  ) {}

  /** Folder name of the locally installed (offline) WPS add-in */
  private get wpsAddonFolderName(): string {
    return 'officeai_1.0.0';
  }

  private get wpsLocalAddonPath(): string {
    return path.join(this.wpsAddonsRoot, this.wpsAddonFolderName);
  }

  /** Office 唤醒器（COM 加载项）DLL 的位置 */
  private get wakerDllCandidates(): string[] {
    return [
      path.join(this.assetsRoot, 'waker', 'OfficeAI.Waker.dll'),
      path.join(this.dataDir, 'waker', 'OfficeAI.Waker.dll'),
      path.join(this.assetsRoot, '..', 'windows', 'waker', 'build', 'OfficeAI.Waker.dll')
    ];
  }

  /** install-waker.ps1 的位置（安装包内 assets/waker；仓库内 windows/waker） */
  private get wakerScriptCandidates(): string[] {
    return [
      path.join(this.assetsRoot, 'waker', 'install-waker.ps1'),
      path.join(this.assetsRoot, '..', 'windows', 'waker', 'install-waker.ps1')
    ];
  }

  private async getWakerStatus(): Promise<{ installed: boolean; loadBehavior: Record<string, number | null> }> {
    const loadBehavior: Record<string, number | null> = { Word: null, Excel: null, PowerPoint: null };
    if (!this.isWindows) return { installed: false, loadBehavior };

    let installed = false;
    for (const app of Object.keys(loadBehavior)) {
      try {
        const { stdout } = await this.runReg([
          'query',
          `HKCU\\Software\\Microsoft\\Office\\${app}\\Addins\\OfficeAI.Waker`,
          '/v',
          'LoadBehavior'
        ]);
        const match = /LoadBehavior\s+REG_DWORD\s+0x([0-9a-f]+)/i.exec(stdout);
        loadBehavior[app] = match ? parseInt(match[1], 16) : 3;
        if (loadBehavior[app] === 3) installed = true;
      } catch {
        loadBehavior[app] = null;
      }
    }
    return { installed, loadBehavior };
  }

  private async isProtocolRegistered(): Promise<boolean> {
    if (!this.isWindows) return false;
    try {
      const { stdout } = await this.runReg([
        'query',
        'HKCU\\Software\\Classes\\officeai\\shell\\open\\command',
        '/ve'
      ]);
      return stdout.toLowerCase().includes('officeai');
    } catch {
      return false;
    }
  }

  private get isWindows(): boolean {
    return process.platform === 'win32';
  }

  private get certPath(): string {
    return path.join(this.dataDir, 'certs', 'localhost.crt');
  }

  private get manifestOutputPath(): string {
    return path.join(this.dataDir, 'office-manifest', 'manifest.xml');
  }

  private get wpsAddonsRoot(): string {
    const appData = process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming');
    return path.join(appData, 'kingsoft', 'wps', 'jsaddons');
  }

  private async runPowerShell(script: string): Promise<{ stdout: string; stderr: string }> {
    return execFileAsync(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', script],
      { timeout: 15000 }
    );
  }

  private async runReg(args: string[]): Promise<{ stdout: string; stderr: string }> {
    return execFileAsync('reg.exe', args, { timeout: 10000 });
  }

  async getStatus(): Promise<WindowsOfficeStatus> {
    const certExists = fs.existsSync(this.certPath);
    let certTrusted = false;
    let msDetected = false;
    let msVersion: string | undefined;
    let msInstalled = false;
    let wpsDetected = false;
    let wpsInstalled = false;
    let wpsLocalInstalled = false;
    let wpsMode: 'local' | 'online' | 'none' = 'none';

    if (this.isWindows) {
      // Check certificate trust
      if (certExists) {
        try {
          const { stdout } = await this.runPowerShell(
            `$c = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2('${this.certPath.replace(/'/g, "''")}'); ` +
              `$store = Get-ChildItem Cert:\\CurrentUser\\Root | Where-Object { $_.Thumbprint -eq $c.Thumbprint }; ` +
              `if ($store) { 'TRUSTED' } else { 'UNTRUSTED' }`
          );
          certTrusted = stdout.includes('TRUSTED');
        } catch {
          certTrusted = false;
        }
      }

      // Check MS Office
      for (const ver of ['16.0', '15.0']) {
        try {
          await this.runReg(['query', `HKCU\\Software\\Microsoft\\Office\\${ver}`]);
          msDetected = true;
          msVersion = ver === '16.0' ? 'Office 2016/2019/2021/365 (16.0)' : 'Office 2013 (15.0)';
          break;
        } catch {
          // try HKLM
          try {
            await this.runReg(['query', `HKLM\\SOFTWARE\\Microsoft\\Office\\${ver}`]);
            msDetected = true;
            msVersion = ver === '16.0' ? 'Office 2016/2019/2021/365 (16.0)' : 'Office 2013 (15.0)';
            break;
          } catch {
            // continue
          }
        }
      }

      try {
        const { stdout } = await this.runReg(['query', OFFICE_DEV_REG_KEY, '/v', OFFICE_ADDIN_ID]);
        msInstalled = stdout.includes(OFFICE_ADDIN_ID) && fs.existsSync(this.manifestOutputPath);
      } catch {
        msInstalled = false;
      }

      // Check WPS
      const wpsAppData = path.dirname(this.wpsAddonsRoot);
      if (fs.existsSync(wpsAppData)) {
        wpsDetected = true;
      } else {
        try {
          await this.runReg(['query', 'HKCU\\Software\\Kingsoft\\Office']);
          wpsDetected = true;
        } catch {
          wpsDetected = false;
        }
      }

      wpsLocalInstalled = fs.existsSync(path.join(this.wpsLocalAddonPath, 'ribbon.xml'));

      const publishXml = path.join(this.wpsAddonsRoot, 'publish.xml');
      const jspluginsXml = path.join(this.wpsAddonsRoot, 'jsplugins.xml');
      for (const xmlPath of [publishXml, jspluginsXml]) {
        if (!fs.existsSync(xmlPath)) continue;
        const txt = fs.readFileSync(xmlPath, 'utf8');
        if (!txt.includes('name="officeai"')) continue;
        wpsInstalled = true;
        if (txt.includes('file:///')) {
          wpsMode = 'local';
        } else if (wpsMode !== 'local') {
          wpsMode = 'online';
        }
      }
      if (wpsLocalInstalled) {
        wpsInstalled = true;
        wpsMode = 'local';
      }
    }

    const officeState: OfficeProcessState = await detectOfficeProcesses();
    const httpUrl = `http://127.0.0.1:${this.httpPort}`;

    return {
      platform: process.platform,
      isWindows: this.isWindows,
      gatewayHttpUrl: httpUrl,
      gatewayHttpsUrl: `https://127.0.0.1:${this.httpsPort}`,
      certExists,
      certTrusted,
      waker: await this.getWakerStatus(),
      launch: {
        target: {
          command: this.launchCommand,
          kind: this.launchCommand
            ? this.launchCommand.toLowerCase().endsWith('.bat')
              ? 'bat'
              : 'exe'
            : 'none',
          source: this.launchCommand ? 'detected' : 'none'
        },
        protocolRegistered: await this.isProtocolRegistered(),
        startPageUrl: `${httpUrl}/start`,
        launchUrl: buildGatewayLaunchUrl(httpUrl, 'office'),
        officeRunning: officeState.running,
        runningProcesses: officeState.processes,
        lifecycle: this.getLifecycleStatus?.() ?? undefined
      },
      msOffice: {
        detected: msDetected,
        version: msVersion,
        installed: msInstalled,
        manifestPath: this.manifestOutputPath,
        registryConfigured: msInstalled
      },
      wps: {
        detected: wpsDetected,
        addonsDir: this.wpsAddonsRoot,
        installed: wpsInstalled,
        localAddinPath: wpsLocalInstalled ? this.wpsLocalAddonPath : undefined,
        mode: wpsMode,
        jspluginsXmlPath: path.join(this.wpsAddonsRoot, 'publish.xml')
      }
    };
  }

  buildOfficeManifestXml(): string {
    const templatePath = path.join(this.assetsRoot, 'office-addin', 'manifest.xml');
    let content = fs.readFileSync(templatePath, 'utf8');
    content = content
      .replace(/https:\/\/127\.0\.0\.1:18443/g, `https://127.0.0.1:${this.httpsPort}`)
      .replace(/http:\/\/127\.0\.0\.1:18080/g, `http://127.0.0.1:${this.httpPort}`)
      // Office users may land on the launcher page so they can start the gateway
      .replace(/(<SupportUrl DefaultValue="[^"]*)\/dashboard"/, '$1/start"');
    return content;
  }

  async trustLocalCertificate(): Promise<IntegrationActionResult> {
    if (!fs.existsSync(this.certPath)) {
      return { ok: false, message: '未找到本地 HTTPS 证书文件，请先重启网关生成证书。' };
    }
    if (!this.isWindows) {
      return {
        ok: false,
        message: '当前非 Windows 系统，已跳过系统证书导入（请在 Windows 上执行）。'
      };
    }

    // 1) 首选 certutil：-user -addstore Root，无需管理员、无弹窗，最可靠
    try {
      await execFileAsync('certutil.exe', ['-user', '-addstore', 'Root', this.certPath], {
        timeout: 15000
      });
      return {
        ok: true,
        message: '已成功将本地 HTTPS 证书信任到当前用户根证书存储 (Cert:\\CurrentUser\\Root)。'
      };
    } catch (certUtilErr) {
      // 2) 回退尝试 PowerShell Import-Certificate（部分系统 certutil 被策略拦截）
      try {
        const psScript =
          `$c = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2('${this.certPath.replace(/'/g, "''")}');` +
          `Import-Certificate -FilePath '${this.certPath.replace(/'/g, "''")}' -CertStoreLocation Cert:\\CurrentUser\\Root | Out-Null;` +
          `'TRUSTED'`;
        const { stdout } = await this.runPowerShell(psScript);
        if (stdout.includes('TRUSTED')) {
          return {
            ok: true,
            message: '已通过 PowerShell 将本地 HTTPS 证书信任到当前用户根证书存储 (Cert:\\CurrentUser\\Root)。'
          };
        }
        throw new Error(stdout.trim() || 'Import-Certificate 未返回 TRUSTED');
      } catch (psErr) {
        const m1 = certUtilErr instanceof Error ? certUtilErr.message : String(certUtilErr);
        const m2 = psErr instanceof Error ? psErr.message : String(psErr);
        return {
          ok: false,
          message: `证书信任失败：certutil 失败(${m1}); PowerShell 备用也失败(${m2})。请手动以管理员打开 PowerShell 执行: certutil -user -addstore Root "${this.certPath}"`
        };
      }
    }
  }

  /**
   * 安装 Office 唤醒器：一个极小的 COM 加载项，Word/Excel/PPT 启动时由它
   * 拉起本地网关（不提供 UI，不常驻）。这是"打开 Office 才启动"的关键。
   */
  async installOfficeWaker(): Promise<IntegrationActionResult> {
    if (!this.isWindows) {
      return { ok: false, message: '唤醒器仅支持 Windows。' };
    }

    const script = this.wakerScriptCandidates.find((candidate) => fs.existsSync(candidate));
    if (!script) {
      return {
        ok: false,
        message:
          '未找到安装脚本 install-waker.ps1（安装包内应在 resources 的 assets\\waker 目录）。请重新安装 OfficeAI。'
      };
    }

    // 安装包里直接用随包分发的 DLL；源码仓库里则让脚本现场编译（Windows 自带 csc.exe）。
    const dll = this.wakerDllCandidates.find((candidate) => fs.existsSync(candidate));
    const canBuild = fs.existsSync(path.join(path.dirname(script), 'Waker.cs'));
    if (!dll && !canBuild) {
      return {
        ok: false,
        message:
          '未找到 OfficeAI.Waker.dll。请重新运行 OfficeAI 安装包，或执行 windows/waker/install-waker.ps1 -Action Install -Build。'
      };
    }
    const dllArgs = dll ? `-DllPath '${dll.replace(/'/g, "''")}'` : '-Build';

    try {
      const { stdout, stderr } = await this.runPowerShell(
        `& '${script.replace(/'/g, "''")}' -Action Install ${dllArgs}`
      );
      return {
        ok: true,
        message:
          'Office 唤醒器安装成功！下次打开 Word / Excel / PowerPoint 时会自动启动 OfficeAI，全部关闭后自动退出。',
        details: [stdout.trim(), stderr.trim()].filter(Boolean)
      };
    } catch (err) {
      return {
        ok: false,
        message: `唤醒器安装失败：${err instanceof Error ? err.message : String(err)}`
      };
    }
  }

  async uninstallOfficeWaker(): Promise<IntegrationActionResult> {
    if (!this.isWindows) {
      return { ok: false, message: '唤醒器仅支持 Windows。' };
    }
    const script = this.wakerScriptCandidates.find((candidate) => fs.existsSync(candidate));
    if (!script) {
      return { ok: false, message: '未找到安装脚本 install-waker.ps1，无需卸载。' };
    }
    try {
      await this.runPowerShell(`& '${script.replace(/'/g, "''")}' -Action Uninstall`);
      return { ok: true, message: '已卸载 Office 唤醒器，Office 不再自动启动 OfficeAI。' };
    } catch (err) {
      return {
        ok: false,
        message: `卸载唤醒器失败：${err instanceof Error ? err.message : String(err)}`
      };
    }
  }

  async installMsOffice(): Promise<IntegrationActionResult> {
    const details: string[] = [];
    const manifestDir = path.dirname(this.manifestOutputPath);
    fs.mkdirSync(manifestDir, { recursive: true });
    fs.writeFileSync(this.manifestOutputPath, this.buildOfficeManifestXml(), 'utf8');
    details.push(`已生成 Office Manifest: ${this.manifestOutputPath}`);

    if (!this.isWindows) {
      return {
        ok: false,
        message: '已生成 manifest.xml，但当前非 Windows 系统，无法写入 Office 注册表。',
        details
      };
    }

    const certRes = await this.trustLocalCertificate();
    details.push(certRes.message);
    if (!certRes.ok) {
      // Manifest 现已默认走 http://127.0.0.1:18080（无需证书），证书仅为 https 备用，
      // 失败不应阻塞 Office 按钮安装，否则会出现“Office里没有按钮”且 installAll 整体失败。
      details.push(`证书备用信任失败（已改用 http，不影响加载）：${certRes.message}`);
    }

    try {
      await this.runReg([
        'add',
        OFFICE_DEV_REG_KEY,
        '/v',
        OFFICE_ADDIN_ID,
        '/t',
        'REG_SZ',
        '/d',
        this.manifestOutputPath,
        '/f'
      ]);
      details.push(`已写入注册表加载项路径: ${OFFICE_DEV_REG_KEY}`);

      return {
        ok: true,
        message:
          'Microsoft Office 加载项安装成功！请重启 Word / Excel / PowerPoint，在「开始」或「插入 → 我的加载项 → 开发人员加载项」中即可看到 OfficeAI。',
        details
      };
    } catch (err) {
      return {
        ok: false,
        message: `写入 Office 注册表失败: ${err instanceof Error ? err.message : String(err)}`,
        details
      };
    }
  }

  async uninstallMsOffice(): Promise<IntegrationActionResult> {
    if (!this.isWindows) {
      return { ok: false, message: '当前非 Windows 系统。' };
    }
    try {
      await this.runReg(['delete', OFFICE_DEV_REG_KEY, '/v', OFFICE_ADDIN_ID, '/f']);
      return { ok: true, message: '已从 Microsoft Office 移除 OfficeAI 加载项。' };
    } catch {
      return { ok: true, message: 'Microsoft Office 中未检测到已安装的 OfficeAI 注册项。' };
    }
  }

  /** Public wrapper used by the launcher page / dashboard */
  async installWpsLocalAddon(): Promise<IntegrationActionResult> {
    return this.installWps();
  }

  get localWpsAddonPath(): string {
    return this.wpsLocalAddonPath;
  }

  private buildWpsPluginEntries(): string[] {
    // Local (offline) add-in: WPS loads these files straight from jsaddons with
    // the file:// protocol, so nothing depends on the gateway being up.
    const normalizedRoot = this.wpsAddonsRoot.replace(/\\/g, '/').replace(/^\/+/, '');
    const localUrl = `file:///${normalizedRoot}/${this.wpsAddonFolderName}/`;
    return ['wps', 'et', 'wpp'].map(
      (type) =>
        `<jsplugin name="officeai" type="${type}" url="${localUrl}" version="1.0.0" enable="enable_dev"/>`
    );
  }

  private mergeWpsPluginXml(existingXml: string | null): string {
    const entries = this.buildWpsPluginEntries();
    if (!existingXml || !existingXml.includes('<jsplugins')) {
      return `<?xml version="1.0" encoding="UTF-8"?>\n<jsplugins>\n  ${entries.join('\n  ')}\n</jsplugins>\n`;
    }

    // Remove previous officeai entries (online + offline) and keep other plugins intact
    let normalized = existingXml.replace(/<jsplugins\s*\/>/i, '<jsplugins></jsplugins>');
    normalized = normalized
      .replace(/<jsplugin\b[^>]*name="officeai"[^>]*\/>\s*/gi, '')
      .replace(/<\/jsplugins>/i, `  ${entries.join('\n  ')}\n</jsplugins>`);
    return normalized;
  }

  private removeWpsPluginXml(existingXml: string): string {
    return existingXml.replace(/\s*<jsplugin\b[^>]*name="officeai"[^>]*\/>/gi, '');
  }

  /**
   * Copies the WPS add-in into jsaddons as a self-contained folder, with the
   * launcher placeholders filled in so a ribbon click can start the gateway.
   */
  private installLocalWpsAddon(): string {
    const src = path.join(this.assetsRoot, 'wps-addin');
    const dest = path.join(this.wpsAddonsRoot, this.wpsAddonFolderName);
    const gatewayHttpUrl = `http://127.0.0.1:${this.httpPort}`;
    const replacements: Record<string, string> = {
      __OFFICEAI_GATEWAY_URL__: gatewayHttpUrl,
      __OFFICEAI_LAUNCH_COMMAND__: this.launchCommand,
      __OFFICEAI_LAUNCH_URL__: buildGatewayLaunchUrl(gatewayHttpUrl, 'wps'),
      __OFFICEAI_BOOTSTRAP_URL__: `${gatewayHttpUrl}/wps/bootstrap.html`
    };

    // Values are injected inside single-quoted JS strings: escape both
    // backslashes and quotes so Windows paths survive (C:\\Users\\...).
    const escapeJsString = (value: string) => value.replace(/\\/g, '\\\\').replace(/'/g, "\\'");

    fs.rmSync(dest, { recursive: true, force: true });
    const copy = (from: string, to: string) => {
      fs.mkdirSync(to, { recursive: true });
      for (const entry of fs.readdirSync(from, { withFileTypes: true })) {
        const srcPath = path.join(from, entry.name);
        const destPath = path.join(to, entry.name);
        if (entry.isDirectory()) {
          copy(srcPath, destPath);
        } else if (/[.](html?|js|json|xml)$/i.test(entry.name)) {
          let text = fs.readFileSync(srcPath, 'utf8');
          const isJs = /\.js$/i.test(entry.name);
          for (const [key, value] of Object.entries(replacements)) {
            text = text.split(key).join(isJs ? escapeJsString(value) : value);
          }
          fs.writeFileSync(destPath, text, 'utf8');
        } else {
          fs.copyFileSync(srcPath, destPath);
        }
      }
    };

    if (!fs.existsSync(src)) {
      // 源文件缺失时必须报错，不能静默跳过拷贝却返回成功（此前 jsaddons 会是空的）。
      throw new Error(`未找到 WPS 加载项源文件: ${src}（安装包可能损坏，请重新安装 OfficeAI）`);
    }
    copy(src, dest);
    return dest;
  }

  async installWps(): Promise<IntegrationActionResult> {
    const details: string[] = [];

    if (!this.isWindows) {
      const wpsTarget = path.join(this.dataDir, 'wps-addon');
      fs.rmSync(wpsTarget, { recursive: true, force: true });
      const src = path.join(this.assetsRoot, 'wps-addin');
      if (fs.existsSync(src)) {
        fs.mkdirSync(wpsTarget, { recursive: true });
        for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
          if (entry.isFile()) {
            fs.copyFileSync(path.join(src, entry.name), path.join(wpsTarget, entry.name));
          }
        }
      }
      return {
        ok: false,
        message: '已生成 WPS 加载项文件，但当前非 Windows 系统，未写入 %APPDATA%\\kingsoft\\wps\\jsaddons。',
        details
      };
    }

    try {
      fs.mkdirSync(this.wpsAddonsRoot, { recursive: true });
      const addonDir = this.installLocalWpsAddon();
      details.push(`已安装本地离线加载项: ${addonDir}`);

      for (const xmlName of ['publish.xml', 'jsplugins.xml']) {
        const xmlPath = path.join(this.wpsAddonsRoot, xmlName);
        const existing = fs.existsSync(xmlPath) ? fs.readFileSync(xmlPath, 'utf8') : null;
        fs.writeFileSync(xmlPath, this.mergeWpsPluginXml(existing), 'utf8');
        details.push(`已更新 WPS 配置: ${xmlPath}`);
      }
    } catch (err) {
      return {
        ok: false,
        message: `WPS 加载项安装失败: ${err instanceof Error ? err.message : String(err)}`,
        details
      };
    }

    const launchHint = this.launchCommand
      ? `点击「OfficeAI 助手」选项卡时若服务未运行，会自动启动：${this.launchCommand}`
      : '未找到 OfficeAI 启动程序，点击按钮时只会提示手动启动。请在配置中心填写「启动命令」。';

    return {
      ok: true,
      message:
        'WPS Office 加载项安装成功！（本地离线模式，不依赖服务先启动）请重启 WPS（文字 / 表格 / 演示），顶部将出现「OfficeAI 助手」选项卡。',
      details: [...details, launchHint]
    };
  }

  /**
   * 一键全量集成：证书信任 + Microsoft Office + WPS + 唤醒器。
   * 供安装程序（OfficeAI.exe --integrate）与首次启动兜底调用，全部幂等，
   * 用户无需手动点任何按钮、手动装证书。
   */
  async installAll(): Promise<IntegrationActionResult> {
    const details: string[] = [];
    // 确保证书存在（installMsOffice 内部会把它信任到系统根证书存储）
    ensureLocalhostCert(this.dataDir);

    const office = await this.installMsOffice();
    details.push(...(office.details ?? []));
    const wps = await this.installWps();
    details.push(...(wps.details ?? []));
    const waker = await this.installOfficeWaker();
    details.push(...(waker.details ?? []));

    // 注册 officeai:// 协议与一键启动快捷方式，保障 WPS 按钮的 LAUNCH_URL 兜底也能拉起
    // （之前 installAll 未包含此步，导致占位符回落到 officeai:// 时 ShellExecute 找不到处理器而静默失败）
    if (this.isWindows && this.launchCommand) {
      try {
        const reg = await registerProtocolHandler(this.launchCommand);
        details.push(reg.message);
        if (reg.ok) {
          const shortcut = await createLauncherShortcut(
            buildGatewayLaunchUrl(`http://127.0.0.1:${this.httpPort}`, 'wps')
          );
          details.push(shortcut.message);
        }
      } catch (err) {
        details.push(`协议/快捷方式注册跳过: ${err instanceof Error ? err.message : String(err)}`);
      }
    } else if (!this.launchCommand) {
      details.push('未检测到 OfficeAI 启动程序，已跳过 officeai:// 协议注册（WPS 将尝试直接用 exe 路径启动）');
    }

    const ok = office.ok && wps.ok && waker.ok;
    return {
      ok,
      message: ok
        ? 'Office / WPS / 唤醒器全部集成完成！重启 Word / Excel / PowerPoint / WPS 后即可使用「AI 助手」。'
        : '部分集成失败，请查看详情。',
      details
    };
  }

  async uninstallWps(): Promise<IntegrationActionResult> {
    if (!this.isWindows) {
      return { ok: false, message: '当前非 Windows 系统。' };
    }

    for (const xmlName of ['publish.xml', 'jsplugins.xml']) {
      const xmlPath = path.join(this.wpsAddonsRoot, xmlName);
      if (fs.existsSync(xmlPath)) {
        const current = fs.readFileSync(xmlPath, 'utf8');
        fs.writeFileSync(xmlPath, this.removeWpsPluginXml(current), 'utf8');
      }
    }
    fs.rmSync(path.join(this.wpsAddonsRoot, this.wpsAddonFolderName), { recursive: true, force: true });
    return { ok: true, message: '已从 WPS Office 移除 OfficeAI 加载项。' };
  }
}
