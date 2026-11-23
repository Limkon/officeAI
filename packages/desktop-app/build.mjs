import { build } from 'esbuild';
import fs from 'node:fs';
import path from 'node:path';

await build({
  entryPoints: ['src/main.ts'],
  bundle: true,
  platform: 'node',
  target: 'node20',
  format: 'cjs',
  outfile: 'dist/main.cjs',
  external: ['electron'],
  sourcemap: false,
  banner: {
    js: "const __officeaiImportMetaUrl = require('node:url').pathToFileURL(__filename).href;"
  },
  define: {
    'import.meta.url': '__officeaiImportMetaUrl'
  }
});

// Ensure assets exist in dist parent (already in packages/desktop-app/assets)
const iconPath = path.resolve('assets', 'icon-32.png');
if (!fs.existsSync(iconPath)) {
  throw new Error(`Missing desktop icon at ${iconPath}`);
}

console.log('Bundled desktop-app main.cjs successfully.');
