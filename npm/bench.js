// Write and read benchmark against exceljs (a dev dependency).
//
//   node bench.js [rows...]      default: 10000 50000 100000
//
// Write: build a workbook with `rows` × 10 mixed columns and encode it.
// Read: decode the same file and walk every value. Each measurement runs in
// a fresh Node process (so no library pays for another's garbage), the
// libraries alternate, and the time excludes loading the module. Median of
// 5 runs (long runs on a laptop vary a lot).
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const COLUMNS = 10;
const RUNS = 5;
const LIBRARIES = ['excel-community', 'exceljs'];

function makeRows(count) {
  const rows = [['ID', 'Name', 'Region', 'Units', 'Price', 'Total', 'Paid', 'Score', 'Code', 'Note']];
  const regions = ['North', 'South', 'East', 'West'];
  for (let i = 1; i <= count; i++) {
    rows.push([
      i, `Customer ${i}`, regions[i % 4], i % 97, (i % 50) + 0.99, (i % 97) * ((i % 50) + 0.99),
      i % 3 === 0, (i * 7) % 101, `C-${i % 1000}`, i % 10 === 0 ? 'priority' : '',
    ]);
  }
  return rows;
}

// ==================== CHILD: one measurement ====================

async function child(library, operation, arg, out) {
  const lib = library === 'excel-community' ? require('./index.js').Excel : require('exceljs');
  let ms;
  let size = 0;
  let cells = 0;
  if (operation === 'write') {
    const rows = makeRows(Number(arg));
    const start = process.hrtime.bigint();
    let bytes;
    if (library === 'excel-community') {
      const wb = lib.create();
      wb.sheet('Data').appendRows(rows);
      bytes = wb.encode();
    } else {
      const wb = new lib.Workbook();
      wb.addWorksheet('Data').addRows(rows);
      bytes = new Uint8Array(await wb.xlsx.writeBuffer());
    }
    ms = Number(process.hrtime.bigint() - start) / 1e6;
    size = bytes.length;
    if (out) fs.writeFileSync(out, bytes);
  } else {
    const bytes = fs.readFileSync(arg);
    const start = process.hrtime.bigint();
    if (library === 'excel-community') {
      for (const row of lib.read(bytes).sheet('Data').rowsAsValues()) cells += row.length;
    } else {
      const wb = new lib.Workbook();
      await wb.xlsx.load(bytes);
      wb.getWorksheet('Data').eachRow((row) => {
        cells += row.values.length - 1; // row.values is 1-based
      });
    }
    ms = Number(process.hrtime.bigint() - start) / 1e6;
  }
  console.log(JSON.stringify({ ms, size, cells }));
}

// ==================== PARENT ====================

function measure(library, operation, arg, out) {
  const args = [__filename, '--child', library, operation, String(arg)];
  if (out) args.push(out);
  return JSON.parse(execFileSync(process.execPath, args, { encoding: 'utf8', maxBuffer: 1 << 20 }));
}

const median = (values) => [...values].sort((a, b) => a - b)[Math.floor(values.length / 2)];
const seconds = (ms) => `${(ms / 1000).toFixed(2)} s`;

function main(counts) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-bench-'));
  console.log(`Node ${process.version}, ${COLUMNS} columns, fresh process per run, median of ${RUNS} runs\n`);
  console.log('| Rows | Library | Write | Read | File size |');
  console.log('| ---: | :--- | ---: | ---: | ---: |');
  for (const count of counts) {
    // Both libraries read the same file, written by excel-community.
    const file = path.join(dir, `data_${count}.xlsx`);
    measure('excel-community', 'write', count, file);
    const results = Object.fromEntries(LIBRARIES.map((l) => [l, { write: [], read: [], size: 0 }]));
    for (let run = 0; run < RUNS; run++) {
      for (const library of LIBRARIES) {
        const w = measure(library, 'write', count);
        results[library].write.push(w.ms);
        results[library].size = w.size;
      }
      for (const library of LIBRARIES) {
        const r = measure(library, 'read', file);
        if (r.cells < (count + 1) * (COLUMNS - 1)) throw new Error(`${library} read only ${r.cells} cells`);
        results[library].read.push(r.ms);
      }
    }
    for (const library of LIBRARIES) {
      const r = results[library];
      console.log(
        `| ${count.toLocaleString('en')} | ${library} | ${seconds(median(r.write))} | ` +
          `${seconds(median(r.read))} | ${(r.size / 1024).toFixed(0)} KB |`
      );
    }
  }
  fs.rmSync(dir, { recursive: true, force: true });
}

if (process.argv[2] === '--child') {
  child(...process.argv.slice(3)).catch((e) => {
    console.error(e);
    process.exit(1);
  });
} else {
  const sizes = process.argv.slice(2).map(Number);
  main(sizes.length ? sizes : [10000, 50000, 100000]);
}
