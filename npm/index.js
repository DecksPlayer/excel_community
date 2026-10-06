require('./dist/excel_community_core.js');
const createExcel = require('./wrapper.js');

// Bundlers replace `fs` with an empty module for browsers (see "browser" in package.json).
let fs = null;
try {
  fs = require('fs');
} catch (e) {}

const Excel = createExcel(globalThis.__excelCommunityCore, fs);

Object.defineProperty(exports, '__esModule', { value: true });
exports.Excel = Excel;
exports.default = Excel;
