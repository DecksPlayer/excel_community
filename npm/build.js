const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const rootDir = path.resolve(__dirname, '..');
// The JS bindings live outside lib/ so the Dart package stays free of
// web-only imports (dart:js_interop) on Android, iOS and desktop.
const dartEntry = path.join(__dirname, 'dart', 'main.dart');
const outDir = path.join(__dirname, 'dist');
const rawOutput = path.join(outDir, 'excel_community_core.raw.js');
const finalOutput = path.join(outDir, 'excel_community_core.js');
const browserOutput = path.join(outDir, 'excel_community.browser.js');
const esmOutput = path.join(outDir, 'excel_community.mjs');

if (!fs.existsSync(outDir)) {
  fs.mkdirSync(outDir, { recursive: true });
}

console.log('Compiling excel_community to JavaScript...');
// -O2 on purpose: -O3/-O4 save only 2.5%/6% gzipped but drop the type checks
// on JS arguments, so e.g. `setColumnWidth('x', 5)` would be accepted and
// write a broken file instead of throwing.
execSync(`dart compile js -O2 -o "${rawOutput}" "${dartEntry}"`, {
  cwd: rootDir,
  stdio: 'inherit',
});

let code = fs.readFileSync(rawOutput, 'utf8');

// dart2js resolves types through `constructor.name`, but bundlers and minifiers
// (esbuild, Terser, Vite...) rename functions, breaking every type check.
// Right after the holders are declared, restore each function's name from its
// property key, which bundlers never rename.
const holders = /^var w=\[[\w$,]+\]$/m;
if (!holders.test(code)) {
  throw new Error('dart2js holder list not found; cannot make the output bundler-safe.');
}
code = code.replace(
  holders,
  (m) =>
    `${m}\n;(function(h){for(var i=0;i<h.length;i++){var o=h[i];for(var k in o){var f=o[k];` +
    `if(typeof f=="function"&&f.name!==k)Object.defineProperty(f,"name",{value:k,configurable:true})}}})(w);`
);

// dart2js expects a browser-like `self`; provide it locally instead of
// defining a global that would make other libraries think they run in a browser.
const wrapped = `(function (self) {\n${code}\n})(typeof self !== 'undefined' ? self : globalThis);\n`;
fs.writeFileSync(finalOutput, wrapped, 'utf8');

// Browser builds bundle the core with wrapper.js, without `fs` or CommonJS.
const wrapper = fs.readFileSync(path.join(__dirname, 'wrapper.js'), 'utf8');
const createBrowserExcel = `(function () {
var module = { exports: {} };
${wrapper}
return module.exports(globalThis.__excelCommunityCore, null);
})()`;

const errorNames = 'ExcelError, ExcelArgumentError, ExcelStateError, ExcelFormatError';

// Plain <script> tag: exposes `ExcelCommunity.Excel` and the error classes.
fs.writeFileSync(
  browserOutput,
  `${wrapped}(function (Excel) {\nconst { ${errorNames} } = Excel;\n` +
    `globalThis.ExcelCommunity = { Excel, ${errorNames} };\n})(${createBrowserExcel});\n`,
  'utf8'
);

// Pure ES module for browsers and bundlers ("browser" condition in package.json),
// so it also works where CommonJS is not converted, e.g. linked packages in Vite.
fs.writeFileSync(
  esmOutput,
  `${wrapped}const Excel = ${createBrowserExcel};\nconst { ${errorNames} } = Excel;\n` +
    `export { Excel, ${errorNames} };\nexport default Excel;\n`,
  'utf8'
);

if (fs.existsSync(rawOutput)) fs.unlinkSync(rawOutput);
const deps = `${rawOutput}.deps`;
if (fs.existsSync(deps)) fs.unlinkSync(deps);
const map = `${rawOutput}.map`;
if (fs.existsSync(map)) fs.unlinkSync(map);

console.log(`Build complete: ${finalOutput}, ${browserOutput}, ${esmOutput}`);
