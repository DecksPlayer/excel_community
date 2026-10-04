if (typeof globalThis.self === 'undefined') globalThis.self = globalThis;
if (typeof window === 'undefined') globalThis.window = globalThis;

require('./dist/excel_community_core.js');

const ec = globalThis.ExcelCommunity;
let fs = null;
try {
  fs = require('fs');
} catch (e) {}

function wrapWorkbook(wb) {
  wb.toBuffer = function () {
    const bytes = wb.encode();
    if (!bytes) return null;
    return typeof Buffer !== 'undefined'
      ? Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
      : bytes;
  };

  wb.save = function (fileName = 'Workbook.xlsx') {
    const bytes = wb.encode();
    if (!bytes) throw new Error('Failed to encode workbook.');

    if (fs && fs.writeFileSync) {
      fs.writeFileSync(fileName, bytes);
      return fileName;
    }

    if (typeof window !== 'undefined' && window.document && window.Blob) {
      const blob = new window.Blob([bytes], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = window.URL.createObjectURL(blob);
      const a = window.document.createElement('a');
      a.href = url;
      a.download = fileName;
      window.document.body.appendChild(a);
      a.click();
      window.document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
      return fileName;
    }

    return bytes;
  };

  return wb;
}

class Excel {
  static create() {
    return wrapWorkbook(ec.create());
  }

  static read(data) {
    if (typeof data === 'string') {
      return wrapWorkbook(ec.readBase64(data));
    }
    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data.buffer || data);
    return wrapWorkbook(ec.read(bytes));
  }

  static fromFile(filePath) {
    if (!fs || !fs.readFileSync) {
      throw new Error('Excel.fromFile is only supported in a Node.js environment.');
    }
    return Excel.read(fs.readFileSync(filePath));
  }
}

module.exports = { Excel };
