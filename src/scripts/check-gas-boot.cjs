#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const vm = require("vm");
const ts = require("typescript");

const projectRoot = path.resolve(__dirname, "../..");
const ignoredDirs = new Set(["scripts", "test"]);

function collectTypeScriptFiles() {
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

function assertExpression(context, expression, description) {
  const result = new vm.Script(expression).runInContext(context);
  if (!result) {
    throw new Error(description);
  }
}

function main() {
  const context = createGasContext();
  const files = collectTypeScriptFiles();

  for (const file of files) {
    try {
      new vm.Script(transpile(file), { filename: file }).runInContext(context);
    } catch (error) {
      throw new Error(`GAS boot failed while loading ${file}: ${error.stack || error.message}`);
    }
  }

  if (typeof context.onOpen !== "function") {
    throw new Error("GAS boot failed: onOpen is not globally available.");
  }

  try {
    context.onOpen();
  } catch (error) {
    throw new Error(`GAS boot failed while running onOpen(): ${error.stack || error.message}`);
  }

  if (!context.__addedMenus.includes("CoMapeo Tools")) {
    throw new Error("GAS boot failed: onOpen() did not add the CoMapeo Tools menu.");
  }

  assertExpression(
    context,
    'LANGUAGES_FALLBACK.pt === "Portuguese"',
    "GAS boot failed: fallback language map did not initialize.",
  );
  assertExpression(
    context,
    'VALID_IMPORT_FIELD_TYPES.includes("selectMultiple")',
    "GAS boot failed: import field type validation did not initialize.",
  );

  console.log(`GAS boot probe passed: ${files.length} TypeScript files loaded; onOpen() added CoMapeo Tools.`);
}

main();
