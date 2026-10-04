const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const rootDir = path.resolve(__dirname, '..');
const dartEntry = path.join(rootDir, 'bin', 'excel_community_npm.dart');
const outDir = path.join(__dirname, 'dist');
const rawOutput = path.join(outDir, 'excel_community_core.raw.js');
const finalOutput = path.join(outDir, 'excel_community_core.js');

if (!fs.existsSync(outDir)) {
  fs.mkdirSync(outDir, { recursive: true });
}

console.log('Compiling excel_community to JavaScript...');
execSync(`dart compile js -O2 -o "${rawOutput}" "${dartEntry}"`, {
  cwd: rootDir,
  stdio: 'inherit',
});

const polyfill = `if (typeof globalThis.self === 'undefined') globalThis.self = globalThis;\n`;
const code = fs.readFileSync(rawOutput, 'utf8');
fs.writeFileSync(finalOutput, polyfill + code, 'utf8');

if (fs.existsSync(rawOutput)) fs.unlinkSync(rawOutput);
const deps = `${rawOutput}.deps`;
if (fs.existsSync(deps)) fs.unlinkSync(deps);
const map = `${rawOutput}.map`;
if (fs.existsSync(map)) fs.unlinkSync(map);

console.log(`Build complete: ${finalOutput}`);
