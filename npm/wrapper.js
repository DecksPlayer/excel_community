// Builds the public `Excel` API on top of the compiled Dart core.
// Shared by index.js (Node, bundlers) and dist/excel_community.browser.js
// (plain <script> tag), so it must not use `require`.
//
// `ec` is the compiled core and `fs` is Node's fs module, or null in browsers.
module.exports = function createExcel(ec, fs) {
  // ==================== ERRORS ====================

  // Dart errors reach JS as an `Error` with an empty `name`. They are
  // rethrown as these classes, so callers can tell them apart with
  // `instanceof` or `name`; the original error is kept as `cause`.
  class ExcelError extends Error {}
  /** An invalid argument or option: bad cell reference, range, value... */
  class ExcelArgumentError extends ExcelError {}
  /** The operation is not possible in the current state. */
  class ExcelStateError extends ExcelError {}
  /** The bytes given to `Excel.read` are not a readable workbook. */
  class ExcelFormatError extends ExcelError {}
  // Set explicitly: minifiers rename classes.
  ExcelError.prototype.name = 'ExcelError';
  ExcelArgumentError.prototype.name = 'ExcelArgumentError';
  ExcelStateError.prototype.name = 'ExcelStateError';
  ExcelFormatError.prototype.name = 'ExcelFormatError';

  // dart2js keeps the Dart exception in `dartException`, and Dart error
  // messages start with the error type.
  function convertError(e, reading) {
    if (!e || e.dartException === undefined) return e;
    const message = e.message;
    let Type = ExcelError;
    if (reading) {
      Type = ExcelFormatError;
    } else if (/^(Invalid argument|RangeError|FormatException|TypeError)/.test(message)) {
      Type = ExcelArgumentError;
    } else if (/^(Bad state|Unsupported operation)/.test(message)) {
      Type = ExcelStateError;
    }
    return new Type(message, { cause: e });
  }

  // Wraps a core object so its methods, getters and setters throw the error
  // classes above. `children` maps member names to wrappers for the API
  // objects they return.
  function guard(target, children) {
    return new Proxy(target, {
      get(t, key) {
        let value;
        try {
          value = t[key];
        } catch (e) {
          throw convertError(e);
        }
        const child = children[key];
        if (typeof value !== 'function') return child ? child(value) : value;
        return function (...args) {
          try {
            const result = value.apply(t, args);
            return child ? child(result) : result;
          } catch (e) {
            throw convertError(e);
          }
        };
      },
      set(t, key, value) {
        try {
          t[key] = value;
        } catch (e) {
          throw convertError(e);
        }
        return true;
      },
    });
  }

  // ==================== CELLS ====================

  function columnName(col) {
    let name = '';
    for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) {
      name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
    }
    return name;
  }

  // A cell is a position on its sheet; the core keeps one object with the
  // cell operations per sheet, because building a core object for every
  // cell is slow.
  class Cell {
    #ops;
    #row;
    #col;

    constructor(ops, row, col) {
      this.#ops = ops;
      this.#row = row;
      this.#col = col;
    }

    #call(name, ...args) {
      try {
        return this.#ops[name](this.#row, this.#col, ...args);
      } catch (e) {
        throw convertError(e);
      }
    }

    get row() { return this.#row; }
    get col() { return this.#col; }
    get cellId() { return columnName(this.#col) + (this.#row + 1); }
    get displayText() { return this.#call('displayText'); }
    get type() { return this.#call('type'); }
    get value() { return this.#call('getValue'); }
    set value(value) { this.#call('setValue', value); }
    get dateValue() { return this.#call('dateValue'); }
    get comment() { return this.#call('getComment'); }
    set comment(text) { this.#call('setComment', text); }
    get formula() { return this.#call('getFormula'); }
    set formula(formula) { this.#call('setFormula', formula); }
    get cachedValue() { return this.#call('cachedValue'); }
    get style() { return this.#call('style'); }
    get dataValidation() { return this.#call('dataValidation'); }

    setFormula(formula) { this.#call('setFormula', formula); }
    setTime(time, minute, second) { this.#call('setTime', time, minute, second); }
    setStyle(options) { this.#call('setStyle', options); }
    resetStyle() { this.#call('resetStyle'); }
    setHyperlink(target, tooltipOrOptions, display) {
      this.#call('setHyperlink', target, tooltipOrOptions, display);
    }
    getHyperlink() { return this.#call('getHyperlink'); }
    removeHyperlink() { this.#call('removeHyperlink'); }
    validates(value) { return this.#call('validates', value); }
  }

  // ==================== SHEETS ====================

  // The core returns the same object for a sheet every time; keep its proxy
  // too, so `wb.sheet(name) === wb.sheet(name)`.
  const sheetProxies = new WeakMap();

  function wrapSheet(sheet) {
    let proxy = sheetProxies.get(sheet);
    if (proxy) return proxy;
    let ops;
    // Positions are `row * 16384 + column`, `null` where there is no cell.
    const cellAt = (position) =>
      position == null
        ? null
        : new Cell((ops ??= sheet.cellOps), Math.floor(position / 16384), position % 16384);
    proxy = guard(sheet, {
      cell: cellAt,
      rows: (rows) => rows.map((row) => row.map(cellAt)),
    });
    sheetProxies.set(sheet, proxy);
    return proxy;
  }

  // ==================== WORKBOOK ====================

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
      if (!bytes) throw new ExcelStateError('Failed to encode workbook.');

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

    return guard(wb, { sheet: wrapSheet, createSheet: wrapSheet });
  }

  function readCore(decode) {
    let wb;
    try {
      wb = decode();
    } catch (e) {
      throw convertError(e, true);
    }
    return wrapWorkbook(wb);
  }

  class Excel {
    static create() {
      return wrapWorkbook(ec.create());
    }

    static read(data) {
      if (typeof data === 'string') {
        return readCore(() => ec.readBase64(data));
      }
      const bytes = ArrayBuffer.isView(data)
        ? new Uint8Array(data.buffer, data.byteOffset, data.byteLength)
        : new Uint8Array(data);
      return readCore(() => ec.read(bytes));
    }

    static fromFile(filePath) {
      if (!fs || !fs.readFileSync) {
        throw new ExcelStateError('Excel.fromFile is only supported in a Node.js environment.');
      }
      return Excel.read(fs.readFileSync(filePath));
    }
  }

  Excel.ExcelError = ExcelError;
  Excel.ExcelArgumentError = ExcelArgumentError;
  Excel.ExcelStateError = ExcelStateError;
  Excel.ExcelFormatError = ExcelFormatError;

  return Excel;
};
