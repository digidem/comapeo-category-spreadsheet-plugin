#!/usr/bin/env node

// Offline regression suite for the translation cancel flow
// (src/translation.ts: requestTranslationCancel / isTranslationCancelRequested
// / autoTranslateSheetsBidirectional boolean contract).
//
// Verifies: cancel mid-run stops at the next checkpoint, already-paid
// translations are kept (Phase 3 runs), remaining sheets are skipped, the
// cancel property is cleaned up, stale flags from earlier runs are ignored,
// and cancelled runs return false (so callers suppress success toasts).

const {
  collectTypeScriptFiles,
  createGasContext,
  loadFiles,
  FakeSheet,
  createFakeSpreadsheet,
} = require("./gasSimHarness.cjs");

const ignoredDirs = new Set(["scripts"]);

const SOURCE_ROWS = 60;

function assert(condition, message) {
  if (!condition) {
    throw new Error(`check-translation-cancel FAILED: ${message}`);
  }
  console.log(`  ok: ${message}`);
}

function createTest() {
  const categoryRows = [["English", "Spanish"]];
  for (let i = 1; i <= SOURCE_ROWS; i++) {
    categoryRows.push([`source text ${i}`, ""]);
  }
  const detailRows = [["Name", "Label"]];
  for (let i = 1; i <= 10; i++) {
    detailRows.push([`cat${i}`, ""]);
  }
  const detailTranslationRows = [["English", "Spanish"]];
  for (let i = 1; i <= 10; i++) {
    detailTranslationRows.push([`detail source ${i}`, ""]);
  }
  const { spreadsheet, sheetsByName, toasts } = createFakeSpreadsheet({
    Categories: new FakeSheet("Categories", [["English"]]),
    Details: new FakeSheet("Details", detailRows),
    "Category Translations": new FakeSheet("Category Translations", categoryRows),
    "Detail Label Translations": new FakeSheet("Detail Label Translations", detailTranslationRows),
  });

  const userProperties = new Map();
  const state = { translateCalls: 0, cancelAtCall: 0 };

  const context = createGasContext();
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
  context.PropertiesService = {
    getScriptProperties() {
      return {
        getProperty: () => null,
        setProperty: () => {},
        deleteProperty: () => {},
      };
    },
    getUserProperties() {
      return {
        getProperty: (key) => (userProperties.has(key) ? userProperties.get(key) : null),
        setProperty: (key, value) => {
          userProperties.set(key, String(value));
        },
        deleteProperty: (key) => {
          userProperties.delete(key);
        },
      };
    },
  };
  context.Utilities.sleep = () => {};
  context.LanguageApp = {
    translate(text, sourceLang, targetLang) {
      state.translateCalls++;
      if (state.cancelAtCall > 0 && state.translateCalls === state.cancelAtCall) {
        // The client-side cancel arrives while the run is in flight.
        context.requestTranslationCancel();
      }
      return `[${targetLang}] ${text}`;
    },
  };

  loadFiles(context, collectTypeScriptFiles(ignoredDirs));

  return { context, sheetsByName, toasts, userProperties, state, spreadsheet };
}

function main() {
  // 1) Cancel mid-run: stops at the checkpoint, keeps partial results.
  console.log("cancel mid-run:");
  const t1 = createTest();
  t1.state.cancelAtCall = 10;
  const completed = t1.context.autoTranslateSheetsBidirectional(["es"]);
  assert(completed === false, "autoTranslateSheetsBidirectional returns false when cancelled");
  assert(t1.state.translateCalls === 25, `cancel noticed at the first checkpoint, 25 pairs (got ${t1.state.translateCalls})`);

  const category = t1.sheetsByName.get("Category Translations");
  let filled = 0;
  for (let row = 1; row <= SOURCE_ROWS; row++) {
    if (String(category.values[row][1]).startsWith("[es]")) {
      filled++;
    }
  }
  assert(filled === 25, `exactly the 25 translated pairs were written (got ${filled})`);
  assert(String(category.values[25][1]).startsWith("[es]"), "rows up to the checkpoint translated");
  assert(category.values[26][1] === "", "rows after the checkpoint left empty for re-run");

  const detailSheet = t1.sheetsByName.get("Detail Label Translations");
  assert(
    !detailSheet || detailSheet.values.filter((row) => String(row[1]).startsWith("[es]")).length === 0,
    "remaining sheets skipped after cancel",
  );
  // The Detail Label sheet now has 10 real source pairs. If the sheet-loop
  // checkpoint were removed, those 10 pairs would translate and the total
  // would be 35. 25 proves the run stopped before the second sheet.
  assert(t1.state.translateCalls === 25, `sheet-loop checkpoint stops before later sheets (got ${t1.state.translateCalls} calls)`);

  assert(
    !t1.userProperties.has("TRANSLATION_CANCEL_REQUESTED"),
    "cancel property deleted after the run",
  );
  assert(
    t1.toasts.some((toast) => /cancelled/i.test(toast.message)),
    "cancel toast shown",
  );

  // 2) Stale flag from a previous crashed run does not cancel a new run.
  console.log("stale cancel flag:");
  const t2 = createTest();
  t2.userProperties.set("TRANSLATION_CANCEL_REQUESTED", String(Date.now() - 60000));
  const completed2 = t2.context.autoTranslateSheetsBidirectional(["es"]);
  assert(completed2 === true, "run with a stale (older) flag completes normally");
  assert(t2.state.translateCalls === SOURCE_ROWS + 10, `all pairs translated (got ${t2.state.translateCalls})`);
  const staleFlags = t2.toasts.filter((toast) => /cancelled/i.test(toast.message)).length;
  assert(staleFlags === 0, "no cancel toast for a stale flag");

  // 3) Cancelled run via manageLanguagesAndTranslate shows no success toast.
  console.log("cancelled run via menu handler:");
  const t3 = createTest();
  t3.state.cancelAtCall = 10;
  t3.context.manageLanguagesAndTranslate({ autoTranslateLanguages: ["es"], customLanguages: [] });
  assert(
    !t3.toasts.some((toast) => toast.message.includes("applied successfully")),
    "success toast suppressed when the run was cancelled",
  );
  assert(
    t3.toasts.some((toast) => /cancelled/i.test(toast.message)),
    "cancel toast shown by the flow itself",
  );

  console.log("check-translation-cancel PASSED");
}

main();
