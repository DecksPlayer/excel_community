// Builds the public `Excel` API on top of the compiled Dart core.
// Shared by index.js (Node, bundlers) and dist/excel_community.browser.js
// (plain <script> tag), so it must not use `require`.
//
// `ec` is the compiled core and `fs` is Node's fs module, or null in browsers.
module.exports = function createExcel(ec, fs) {
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
        // Revoking synchronously can cancel the download in some browsers.
        setTimeout(() => window.URL.revokeObjectURL(url), 0);
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
      const bytes = ArrayBuffer.isView(data)
        ? new Uint8Array(data.buffer, data.byteOffset, data.byteLength)
        : new Uint8Array(data);
      return wrapWorkbook(ec.read(bytes));
    }

    static fromFile(filePath) {
      if (!fs || !fs.readFileSync) {
        throw new Error('Excel.fromFile is only supported in a Node.js environment.');
      }
      return Excel.read(fs.readFileSync(filePath));
    }
  }

  return Excel;
};
