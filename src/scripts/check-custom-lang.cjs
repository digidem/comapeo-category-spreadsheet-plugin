#!/usr/bin/env node

// Offline regression suite for custom-language translation targets.
// Covers the fix that lets Indigenous/custom languages missing from the
// remote languages.json be used as translation targets when registered as
// "Name - code" sheet headers (see src/translation.ts:
// resolveTranslationColumnHeader, autoTranslateSheetsBidirectional, and the
// translator echo-guard).

const { collectTypeScriptFiles, createGasContext, loadFiles } = require("./gasSimHarness.cjs");

const ignoredDirs = new Set(["scripts"]);

class FakeSheet {
  constructor(name, rows = [[]]) {
    this.name = name;
    this.values = rows.map((row) => row.slice());
    this.formulas = [];
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
    };
  }

  getDataRange() {
    return this.getRange(1, 1, this.getLastRow(), this.getLastColumn());
  }

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

function assert(condition, message) {
  if (!condition) {
    throw new Error(`check-custom-lang FAILED: ${message}`);
  }
  console.log(`  ok: ${message}`);
}

function pickCustomCode(allLanguages) {
  const candidates = ["xav", "qaa", "zzq"];
  const code = candidates.find((candidate) => !allLanguages[candidate] && !allLanguages[candidate.toLowerCase()]);
  if (!code) {
    throw new Error("No candidate custom language code available; update candidates list.");
  }
  return code;
}

function main() {
  const context = createGasContext();
  const { spreadsheet, sheetsByName, toasts } = createFakeSpreadsheet({
    Categories: new FakeSheet("Categories", [["English"]]),
    Details: new FakeSheet("Details", [["Name", "Label"], ["cat1", "Water"]]),
    "Category Translations": new FakeSheet("Category Translations"),
  });
  const echoTargets = new Set();
  context.SpreadsheetApp = {
    getActiveSpreadsheet() {
      return spreadsheet;
    },
    getUi() {
      return {
        Button: { NO: "NO", YES: "YES" },
        ButtonSet: { OK: "OK", YES_NO: "YES_NO" },
        alert() {
          return "OK";
        },
      };
    },
  };
  context.LanguageApp = {
    translate(text, sourceLang, targetLang) {
      if (echoTargets.has(String(targetLang).toLowerCase())) {
        return text;
      }
      return `[${targetLang}] ${text}`;
    },
  };

  loadFiles(context, collectTypeScriptFiles(ignoredDirs));

  const customCode = pickCustomCode(context.getAllLanguages());
  const customHeader = `Xavante - ${customCode}`;

  // 1) ensureLanguageColumnsExist: existing custom header is not duplicated,
  //    missing standard language is added with its display name.
  //    Custom headers are only recognized past column index 2, matching real
  //    sheets where custom languages are appended after standard columns.
  console.log("ensureLanguageColumnsExist:");
  const categorySheet = sheetsByName.get("Category Translations");
  categorySheet.values = [["English", "French", "Italian", customHeader]];
  context.ensureLanguageColumnsExist(categorySheet, ["es", customCode]);
  let headers = categorySheet.values[0];
  assert(headers.length === 5, `adds exactly one missing column (got ${headers.length}: ${headers.join(", ")})`);
  assert(headers.includes("Spanish"), "standard target uses display name header");
  assert(headers.filter((header) => header === customHeader).length === 1, "existing custom header not duplicated");
  context.ensureLanguageColumnsExist(categorySheet, ["es", customCode]);
  assert(categorySheet.values[0].length === 5, "re-run adds nothing");

  // 2) createTranslationSheet resolves the custom name from the other
  //    translation sheet's header ("Name - code" format, not bare code).
  console.log("createTranslationSheet:");
  const detailSheet = context.createTranslationSheet("Detail Label Translations", ["es", customCode], "Details", "B");
  const detailHeaders = detailSheet.values[0];
  assert(
    detailHeaders.includes(customHeader),
    `custom target header is "${customHeader}" (got: ${detailHeaders.join(", ")})`,
  );

  // 3) Bare-code fallback when the custom language is registered nowhere.
  console.log("createCategoryTranslationsSheet (no custom headers yet):");
  const fresh = createFakeSpreadsheet({
    Categories: new FakeSheet("Categories", [["English"]]),
    Details: new FakeSheet("Details", [["Name", "Label"], ["cat1", "Water"]]),
  });
  context.SpreadsheetApp.getActiveSpreadsheet = () => fresh.spreadsheet;
  const freshCategorySheet = context.createCategoryTranslationsSheet(["es", customCode]);
  const freshHeaders = freshCategorySheet.values[0];
  assert(
    freshHeaders[freshHeaders.length - 1] === customCode,
    `unknown custom target falls back to bare code header (got: ${freshHeaders.join(", ")})`,
  );

  // Back to the main spreadsheet for the translation run.
  context.SpreadsheetApp.getActiveSpreadsheet = () => spreadsheet;

  // 4) autoTranslateSheetsBidirectional accepts the custom target, warns via
  //    toast, rejects unknown codes that are not registered as headers.
  console.log("autoTranslateSheetsBidirectional:");
  categorySheet.values = [
    ["English", "French", "Italian", customHeader, "Spanish"],
    ["Hello", "", "", "", ""],
    ["Water", "", "", "", ""],
  ];

  let threw = null;
  try {
    context.autoTranslateSheetsBidirectional(["zzzz"]);
  } catch (error) {
    threw = error;
  }
  assert(threw && /Target languages not available/.test(threw.message), "unknown non-registered code still throws");

  echoTargets.add(customCode); // translator echoes source for the custom code
  context.autoTranslateSheetsBidirectional(["es", customCode]);
  assert(
    categorySheet.values[1][4] === "[es] Hello" && categorySheet.values[2][4] === "[es] Water",
    `standard target cells translated (got: ${categorySheet.values[1][4]}, ${categorySheet.values[2][4]})`,
  );
  assert(
    categorySheet.values[1][3] === "" && categorySheet.values[2][3] === "",
    "echoed source text is NOT written as a fake translation for the custom target",
  );
  assert(
    toasts.some((toast) => toast.message.includes(customCode)),
    "toast warns that the translator may not support the custom language",
  );

  // Echo-guard off for known languages: with the stub echoing for "es", the
  // returned text is cached and written as-is — exactly the prior behavior.
  echoTargets.add("es");
  categorySheet.values = [
    ["English", "French", "Italian", customHeader, "Spanish"],
    ["GPS", "", "", "", ""],
  ];
  context.autoTranslateSheetsBidirectional(["es"]);
  assert(
    categorySheet.values[1][4] === "GPS",
    `known target language keeps prior behavior (no echo-guard); got: ${categorySheet.values[1][4]}`,
  );

  console.log("check-custom-lang PASSED");
}

main();
