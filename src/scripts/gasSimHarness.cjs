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

module.exports = {
  projectRoot,
  collectTypeScriptFiles,
  createGasContext,
  transpile,
  loadFiles,
};
