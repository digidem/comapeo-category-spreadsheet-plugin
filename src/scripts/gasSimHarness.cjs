#!/usr/bin/env node

// Shared harness for simulating the Apps Script global-scope boot: mocks the GAS
// APIs, transpiles + concatenates project .ts files in the same alphabetical
// order GAS uses, and runs them in a single vm context. Used by both
// check-gas-boot.cjs (menu/load-order probe) and check-lint-parity.cjs
// (runLintParityTests()) so the two mock environments can't drift apart.

const fs = require("fs");
const path = require("path");
const vm = require("vm");
const ts = require("typescript");

const projectRoot = path.resolve(__dirname, "../..");

function collectTypeScriptFiles(ignoredDirs) {
  const files = ["index.ts"];

  function walk(directory) {
    for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
      const fullPath = path.join(directory, entry.name);
      const relativePath = path.relative(projectRoot, fullPath);

      if (entry.isDirectory()) {
        if (!ignoredDirs.has(entry.name)) {
          walk(fullPath);
        }
        continue;
      }

      if (entry.isFile() && entry.name.endsWith(".ts") && !entry.name.endsWith(".d.ts")) {
        files.push(relativePath);
      }
    }
  }

  walk(path.join(projectRoot, "src"));
  return [...new Set(files)].sort();
}

function createGasContext() {
  let uuidCounter = 0;
  const context = {
    console,
    exports: {},
    globalThis: null,
    Session: {
      getActiveUserLocale() {
        return "en_US";
      },
    },
  };
  context.globalThis = context;

  const menu = {
    addItem() {
      return this;
    },
    addSeparator() {
      return this;
    },
    addSubMenu() {
      return this;
    },
    addToUi() {
      context.__addedMenus.push(this.__name);
      return this;
    },
  };

  const ui = {
    Button: { NO: "NO", YES: "YES" },
    ButtonSet: { OK: "OK", YES_NO: "YES_NO" },
    alert() {
      return "OK";
    },
    createMenu(name) {
      return Object.assign(Object.create(menu), { __name: name });
    },
  };

  context.__addedMenus = [];
  context.SpreadsheetApp = {
    getActiveSpreadsheet() {
      return {
        getSheetByName() {
          return null;
        },
        getSpreadsheetLocale() {
          return "en";
        },
      };
    },
    getUi() {
      return ui;
    },
  };
  context.PropertiesService = {
    getScriptProperties() {
      return {
        deleteProperty() {},
        getProperty() {
          return null;
        },
        setProperty() {},
      };
    },
    getUserProperties() {
      return {
        deleteProperty() {},
        getProperty() {
          return null;
        },
        setProperty() {},
      };
    },
  };
  context.CacheService = {
    getScriptCache() {
      return {
        get() {
          return null;
        },
        put() {},
        remove() {},
      };
    },
  };
  context.Utilities = {
    base64Decode(value) {
      return Buffer.from(String(value), "base64");
    },
    base64Encode(value) {
      return Buffer.from(String(value)).toString("base64");
    },
    formatDate() {
      return "2026-06-28";
    },
    getUuid() {
      uuidCounter += 1;
      return `uuid-${uuidCounter}`;
    },
    newBlob(value) {
      return {
        getBytes() {
          return Array.from(Buffer.from(String(value)));
        },
        getDataAsString() {
          return String(value);
        },
      };
    },
    sleep() {},
  };
  context.DriveApp = {};
  context.LanguageApp = {};
  context.UrlFetchApp = {};

  return vm.createContext(context);
}

function transpile(filePath) {
  const source = fs.readFileSync(path.join(projectRoot, filePath), "utf8");
  return ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.None,
      target: ts.ScriptTarget.ES2020,
    },
    fileName: filePath,
  }).outputText;
}

function loadFiles(context, files) {
  for (const file of files) {
    try {
      new vm.Script(transpile(file), { filename: file }).runInContext(context);
    } catch (error) {
      throw new Error(`GAS boot failed while loading ${file}: ${error.stack || error.message}`);
    }
  }
}

// Shared fake spreadsheet used by offline checks (check-custom-lang.cjs,
// check-translation-cancel.cjs). Supports numeric and A1-notation ranges,
// value reads/writes, formulas (recorded only) and toasts.
class FakeSheet {
  constructor(name, rows = [[]]) {
    this.name = name;
    this.values = rows.map((row) => row.slice());
    this.formulas = [];
    this.backgrounds = [];
  }

  getName() {
    return this.name;
  }

  getLastColumn() {
    return this.values.reduce((max, row) => Math.max(max, row.length), 0);
  }

  getLastRow() {
    return this.values.length;
  }

  getRange(rowOrA1, col, numRows = 1, numCols = 1) {
    const sheet = this;
    let row = rowOrA1;
    if (typeof rowOrA1 === "string") {
      const match = /^([A-Z]+)(\d+)$/.exec(rowOrA1.trim());
      if (!match) {
        throw new Error(`FakeSheet: unsupported A1 range "${rowOrA1}"`);
      }
      row = Number(match[2]);
      col = match[1].split("").reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0);
      numRows = 1;
      numCols = 1;
    }
    return {
      getValues() {
        const out = [];
        for (let r = 0; r < numRows; r++) {
          const source = sheet.values[row - 1 + r] || [];
          const line = [];
          for (let c = 0; c < numCols; c++) {
            line.push(source[col - 1 + c] !== undefined ? source[col - 1 + c] : "");
          }
          out.push(line);
        }
        return out;
      },
      getValue() {
        return this.getValues()[0][0];
      },
      setValues(values) {
        for (let r = 0; r < values.length; r++) {
          while (sheet.values.length <= row - 1 + r) {
            sheet.values.push([]);
          }
          const line = values[r];
          for (let c = 0; c < line.length; c++) {
            const targetRow = sheet.values[row - 1 + r];
            while (targetRow.length <= col - 1 + c) {
              targetRow.push("");
            }
            targetRow[col - 1 + c] = line[c];
          }
        }
        return this;
      },
      setFormula(formula) {
        sheet.formulas.push(formula);
        return this;
      },
      setFontWeight() {
        return this;
      },
      setFontColors() {
        return this;
      },
      setBackground(color) {
        sheet.backgrounds[row - 1] = sheet.backgrounds[row - 1] || [];
        sheet.backgrounds[row - 1][col - 1] = color;
        return this;
      },
      getBackground() {
        const rowBgs = sheet.backgrounds[row - 1] || [];
        return rowBgs[col - 1] !== undefined ? rowBgs[col - 1] : null;
      },
      setBackgrounds(values) {
        for (let r = 0; r < values.length; r++) {
          while (sheet.backgrounds.length <= row - 1 + r) {
            sheet.backgrounds.push([]);
          }
          const line = values[r];
          for (let c = 0; c < line.length; c++) {
            const targetRow = sheet.backgrounds[row - 1 + r];
            while (targetRow.length <= col - 1 + c) {
              targetRow.push(null);
            }
            targetRow[col - 1 + c] = line[c];
          }
        }
        return this;
      },
      getBackgrounds() {
        const out = [];
        for (let r = 0; r < numRows; r++) {
          const source = sheet.backgrounds[row - 1 + r] || [];
          const line = [];
          for (let c = 0; c < numCols; c++) {
            line.push(source[col - 1 + c] !== undefined ? source[col - 1 + c] : null);
          }
          out.push(line);
        }
        return out;
      },
    };
  }

  getDataRange() {
    return this.getRange(1, 1, this.getLastRow(), this.getLastColumn());
  }

  autoResizeColumns() {}

  insertColumnsAfter() {}

  insertColumns() {}
}

function createFakeSpreadsheet(initialSheets = {}) {
  const sheetsByName = new Map(Object.entries(initialSheets));
  const toasts = [];
  return {
    spreadsheet: {
      getName() {
        return "Test Config";
      },
      getSheetByName(name) {
        return sheetsByName.get(name) || null;
      },
      insertSheet(name) {
        const sheet = new FakeSheet(name);
        sheetsByName.set(name, sheet);
        return sheet;
      },
      getSpreadsheetLocale() {
        return "en";
      },
      toast(message, title) {
        toasts.push({ message, title });
      },
    },
    sheetsByName,
    toasts,
  };
}

module.exports = {
  projectRoot,
  collectTypeScriptFiles,
  createGasContext,
  transpile,
  loadFiles,
  FakeSheet,
  createFakeSpreadsheet,
};
