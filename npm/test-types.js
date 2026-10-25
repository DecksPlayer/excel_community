// Checks that index.d.ts declares exactly the members the runtime objects
// have, so the types do not fall behind when the core gains features.
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { Excel } = require('./index.js');

const dts = fs.readFileSync(path.join(__dirname, 'index.d.ts'), 'utf8');

/** Member names of `export interface <name> { ... }` (two-space indented). */
function declaredMembers(name) {
  const block = dts.match(new RegExp(`^export interface ${name} \\{\\n([\\s\\S]*?)^\\}`, 'm'));
  assert.ok(block, `interface ${name} not found in index.d.ts`);
  const names = new Set();
  for (const line of block[1].split('\n')) {
    const m = line.match(/^ {2}(?:readonly |get |set )?(\w+)\??[(:]/);
    if (m) names.add(m[1]);
  }
  return names;
}

/** Public members of a runtime object: own properties and class members. */
function runtimeMembers(obj) {
  const names = new Set();
  for (let o = obj; o && o !== Object.prototype; o = Object.getPrototypeOf(o)) {
    for (const key of Object.getOwnPropertyNames(o)) {
      if (key !== 'constructor' && !key.startsWith('_')) names.add(key);
    }
  }
  return names;
}

// Core objects are proxies over plain objects, so read their own properties
// through `Object.getOwnPropertyNames` of the proxy (forwarded to the target).
const wb = Excel.create();
const sheet = wb.sheet('Sheet1');
const internal = new Set(['cellOps']);
const objects = {
  Workbook: wb,
  Sheet: sheet,
  Cell: sheet.cell('A1'),
};

let failed = 0;
for (const [name, obj] of Object.entries(objects)) {
  const declared = declaredMembers(name);
  const actual = new Set([...runtimeMembers(obj)].filter((k) => !internal.has(k)));
  const missingTypes = [...actual].filter((k) => !declared.has(k)).sort();
  const missingRuntime = [...declared].filter((k) => !actual.has(k)).sort();
  if (missingTypes.length || missingRuntime.length) {
    failed++;
    console.log(`  ✗ ${name}`);
    if (missingTypes.length) console.log(`      not in index.d.ts: ${missingTypes.join(', ')}`);
    if (missingRuntime.length) console.log(`      not in the runtime: ${missingRuntime.join(', ')}`);
  } else {
    console.log(`  ✓ ${name}: ${declared.size} members`);
  }
}

if (failed) process.exit(1);
console.log('\n🎉 index.d.ts matches the runtime API');
