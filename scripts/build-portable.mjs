import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';

const root = process.cwd();
const outDir = path.join(root, 'release', 'OfficeAI-Portable');

fs.rmSync(outDir, { recursive: true, force: true });
fs.mkdirSync(outDir, { recursive: true });

function copyDir(src, dst) {
  fs.mkdirSync(dst, { recursive: true });
  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    const s = path.join(src, entry.name);
    const d = path.join(dst, entry.name);
    if (entry.isDirectory()) {
      copyDir(s, d);
    } else {
      fs.copyFileSync(s, d);
    }
  }
}

// Copy compiled assets
copyDir(path.join(root, 'packages', 'addin-ui', 'dist'), path.join(outDir, 'assets', 'addin-ui'));
copyDir(path.join(root, 'packages', 'office-addin'), path.join(outDir, 'assets', 'office-addin'));
copyDir(path.join(root, 'packages', 'wps-addin'), path.join(outDir, 'assets', 'wps-addin'));

// Copy compiled gateway + shared
copyDir(path.join(root, 'packages', 'gateway', 'dist'), path.join(outDir, 'gateway'));
copyDir(path.join(root, 'packages', 'shared', 'dist'), path.join(outDir, 'node_modules', '@officeai', 'shared', 'dist'));
fs.copyFileSync(
  path.join(root, 'packages', 'shared', 'package.json'),
  path.join(outDir, 'node_modules', '@officeai', 'shared', 'package.json')
);
copyDir(path.join(root, 'node_modules', 'selfsigned'), path.join(outDir, 'node_modules', 'selfsigned'));
if (fs.existsSync(path.join(root, 'node_modules', 'node-forge'))) {
  copyDir(path.join(root, 'node_modules', 'node-forge'), path.join(outDir, 'node_modules', 'node-forge'));
}

// Copy Windows helper scripts
fs.copyFileSync(path.join(root, 'scripts', 'windows', 'Start-OfficeAI.bat'), path.join(outDir, 'Start-OfficeAI.bat'));
fs.copyFileSync(path.join(root, 'scripts', 'windows', 'Install-OfficeAI.bat'), path.join(outDir, 'Install-OfficeAI.bat'));
fs.copyFileSync(path.join(root, 'scripts', 'windows', 'Install-OfficeAI.ps1'), path.join(outDir, 'Install-OfficeAI.ps1'));
fs.copyFileSync(path.join(root, 'scripts', 'windows', 'Uninstall-OfficeAI.ps1'), path.join(outDir, 'Uninstall-OfficeAI.ps1'));
fs.writeFileSync(
  path.join(outDir, 'package.json'),
  JSON.stringify({ name: 'officeai-portable', version: '1.0.0', type: 'module' }, null, 2)
);

const zipPath = path.join(root, 'release', 'OfficeAI-Windows-Portable.zip');
try {
  execFileSync('zip', ['-rq', zipPath, 'OfficeAI-Portable'], {
    cwd: path.join(root, 'release')
  });
  console.log(`Built portable package: ${zipPath}`);
} catch {
  console.log(`Built portable directory: ${outDir}`);
}
