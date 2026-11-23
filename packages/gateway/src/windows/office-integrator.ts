import { execFile } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { promisify } from 'node:util';
import type { WindowsOfficeStatus } from '@officeai/shared';

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
    private readonly assetsRoot: string
  ) {}

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

      const publishXml = path.join(this.wpsAddonsRoot, 'publish.xml');
      const jspluginsXml = path.join(this.wpsAddonsRoot, 'jsplugins.xml');
      for (const xmlPath of [publishXml, jspluginsXml]) {
        if (fs.existsSync(xmlPath)) {
          const txt = fs.readFileSync(xmlPath, 'utf8');
          if (txt.includes('name="officeai"')) {
            wpsInstalled = true;
            break;
          }
        }
      }
    }

    return {
      platform: process.platform,
      isWindows: this.isWindows,
      gatewayHttpUrl: `http://127.0.0.1:${this.httpPort}`,
      gatewayHttpsUrl: `https://127.0.0.1:${this.httpsPort}`,
      certExists,
      certTrusted,
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
        jspluginsXmlPath: path.join(this.wpsAddonsRoot, 'publish.xml')
      }
    };
  }

  buildOfficeManifestXml(): string {
    const templatePath = path.join(this.assetsRoot, 'office-addin', 'manifest.xml');
    let content = fs.readFileSync(templatePath, 'utf8');
    content = content
      .replace(/https:\/\/127\.0\.0\.1:18443/g, `https://127.0.0.1:${this.httpsPort}`)
      .replace(/http:\/\/127\.0\.0\.1:18080/g, `http://127.0.0.1:${this.httpPort}`);
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

    try {
      await this.runPowerShell(
        `Import-Certificate -FilePath '${this.certPath.replace(/'/g, "''")}' -CertStoreLocation Cert:\\CurrentUser\\Root | Out-Null`
      );
      return {
        ok: true,
        message: '已成功将本地 HTTPS 证书信任到当前用户根证书存储 (Cert:\\CurrentUser\\Root)。'
      };
    } catch (err) {
      return {
        ok: false,
        message: `导入证书失败: ${err instanceof Error ? err.message : String(err)}`
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

  private mergeWpsPluginXml(existingXml: string | null, gatewayHttpUrl: string): string {
    const cleanUrl = `${gatewayHttpUrl.replace(/\/+$/, '')}/wps/`;
    const entries = [
      `<jsplugin name="officeai" type="wps" url="${cleanUrl}" version="1.0.0" enable="enable_dev"/>`,
      `<jsplugin name="officeai" type="et" url="${cleanUrl}" version="1.0.0" enable="enable_dev"/>`,
      `<jsplugin name="officeai" type="wpp" url="${cleanUrl}" version="1.0.0" enable="enable_dev"/>`
    ];

    if (!existingXml || !existingXml.includes('<jsplugins')) {
      return `<?xml version="1.0" encoding="UTF-8"?>\n<jsplugins>\n  ${entries.join('\n  ')}\n</jsplugins>\n`;
    }

    // Remove previous officeai entries and keep other user plugins intact
    let normalized = existingXml.replace(/<jsplugins\s*\/>/i, '<jsplugins></jsplugins>');
    normalized = normalized
      .replace(/<jsplugin\b[^>]*name="officeai"[^>]*\/>\s*/g, '')
      .replace(/<\/jsplugins>/i, `  ${entries.join('\n  ')}\n</jsplugins>`);
    return normalized;
  }

  private removeWpsPluginXml(existingXml: string): string {
    return existingXml.replace(/\s*<jsplugin\b[^>]*name="officeai"[^>]*\/>/g, '');
  }

  async installWps(): Promise<IntegrationActionResult> {
    const details: string[] = [];
    const gatewayHttpUrl = `http://127.0.0.1:${this.httpPort}`;

    // Copy static WPS addon files into dataDir/wps-addon for reference / offline inspection
    const wpsSrc = path.join(this.assetsRoot, 'wps-addin');
    const wpsTarget = path.join(this.dataDir, 'wps-addon');
    if (fs.existsSync(wpsSrc)) {
      fs.mkdirSync(wpsTarget, { recursive: true });
      for (const file of fs.readdirSync(wpsSrc)) {
        const srcFile = path.join(wpsSrc, file);
        if (fs.statSync(srcFile).isFile()) {
          let content = fs.readFileSync(srcFile, 'utf8');
          content = content.replace(/http:\/\/127\.0\.0\.1:18080/g, gatewayHttpUrl);
          fs.writeFileSync(path.join(wpsTarget, file), content, 'utf8');
        }
      }
      details.push(`已同步 WPS 加载项资源到: ${wpsTarget}`);
    }

    if (!this.isWindows) {
      return {
        ok: false,
        message: '已生成 WPS 加载项文件，但当前非 Windows 系统，未写入 %APPDATA%\\kingsoft\\wps\\jsaddons。',
        details
      };
    }

    fs.mkdirSync(this.wpsAddonsRoot, { recursive: true });

    for (const xmlName of ['publish.xml', 'jsplugins.xml']) {
      const xmlPath = path.join(this.wpsAddonsRoot, xmlName);
      const existing = fs.existsSync(xmlPath) ? fs.readFileSync(xmlPath, 'utf8') : null;
      const merged = this.mergeWpsPluginXml(existing, gatewayHttpUrl);
      fs.writeFileSync(xmlPath, merged, 'utf8');
      details.push(`已更新 WPS 配置: ${xmlPath}`);
    }

    return {
      ok: true,
      message:
        'WPS Office 加载项安装成功！请重启 WPS（文字 / 表格 / 演示），顶部菜单栏将出现「OfficeAI」选项卡。',
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
    return { ok: true, message: '已从 WPS Office 移除 OfficeAI 加载项。' };
  }
}
