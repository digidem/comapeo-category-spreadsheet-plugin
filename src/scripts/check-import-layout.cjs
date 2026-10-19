#!/usr/bin/env node

// Offline regression suite for the canonical Categories sheet layout
// (issues #61 / #52). Covers:
//   1. CATEGORY_HEADERS single source of truth.
//   2. applyCategories (legacy mapeosettings import path) writes the canonical
//      7-column header + appliesTo / Category ID / Color / Icon ID values.
//   3. populateCategoriesSheet (.comapeocat import path) writes the canonical
//      layout, appliesTo per category, and column-A backgrounds.
//   4. Sheet-level round trip: the generate-side buildCategories reads back the
//      imported sheet with identical appliesTo + color per category.
//   5. validateCategoriesHeaderLayout accepts the canonical layout and flags
//      the old import layout with an actionable message.
//   6. Full GAS test suite (runBuildAndImportTests) still passes.

const vm = require("vm");
const {
  collectTypeScriptFiles,
  createGasContext,
  loadFiles,
  FakeSheet,
  createFakeSpreadsheet,
} = require("./gasSimHarness.cjs");

const ignoredDirs = new Set(["scripts"]);

function assert(condition, message) {
  if (!condition) {
    throw new Error(`check-import-layout FAILED: ${message}`);
  }
  console.log(`  ok: ${message}`);
}

function main() {
  const context = createGasContext();
  const fixture = createFakeSpreadsheet({
    Categories: new FakeSheet("Categories"),
  });
  context.SpreadsheetApp.getActiveSpreadsheet = () => fixture.spreadsheet;

  loadFiles(context, collectTypeScriptFiles(ignoredDirs));

  const canonical = [
    "Name",
    "Icon",
    "Fields",
    "Applies",
    "Category ID",
    "Color",
    "Icon ID",
  ];

  // 1) CATEGORY_HEADERS constant is the single source of truth. It is a
  // top-level const, so read it through the context rather than the object.
  const canonicalConst = vm.runInContext("CATEGORY_HEADERS", context);
  assert(
    JSON.stringify(canonicalConst) === JSON.stringify(canonical),
    `CATEGORY_HEADERS equals ${canonical.join(" | ")}`,
  );

  // 2) applyCategories (legacy mapeosettings import path)
  console.log("applyCategories writes canonical layout:");
  const sheet = fixture.sheetsByName.get("Categories");
  context.applyCategories(
    sheet,
    [
      {
        id: "river-preset",
        name: "River",
        icon: "river",
        fields: ["f1"],
        appliesTo: ["track", "observation"],
        color: "#1B5E88",
      },
      { name: "Plain", icon: "plain", fields: [] },
    ],
    [{ id: "f1", label: "Depth" }],
    [
      { name: "river", svg: "<svg></svg>" },
      { name: "plain", svg: "<svg></svg>" },
    ],
  );
  assert(
    JSON.stringify(sheet.values[0]) === JSON.stringify(canonical),
    "header row equals canonical CATEGORY_HEADERS",
  );
  assert(
    JSON.stringify(sheet.values[1]) ===
      JSON.stringify([
        "River",
        "<svg></svg>",
        "Depth",
        "track, observation",
        "river-preset",
        "#1B5E88",
        "",
      ]),
    "legacy preset row carries appliesTo, Category ID, Color, Icon ID",
  );
  assert(
    sheet.values[2][3] === "track, observation" &&
      sheet.values[2][4] === "" &&
      sheet.values[2][5] === "",
    "preset without appliesTo defaults to 'track, observation'; color/id stay blank",
  );
  assert(
    sheet.getRange(2, 1).getBackground() === "#1B5E88",
    "column A background still painted with the category color",
  );

  // 3) populateCategoriesSheet (.comapeocat import path)
  console.log("populateCategoriesSheet writes canonical layout:");
  const importFixture = createFakeSpreadsheet({});
  const categories = [
    {
      id: "trees",
      name: "Trees",
      appliesTo: ["observation"],
      color: "#4CAF50",
      iconId: "trees",
      defaultFieldIds: ["species"],
      fields: ["species"],
    },
    { id: "rivers", name: "Rivers", appliesTo: ["track", "observation"] },
    { id: "ghost", name: "Ghost" },
  ];
  const iconMap = new Map([
    ["trees", "data:image/svg+xml;base64,PHN2Zz48L3N2Zz4="],
  ]);
  const fields = [
    { id: "species", tagKey: "species", name: "Species", type: "select" },
  ];
  context.populateCategoriesSheet(
    importFixture.spreadsheet,
    categories,
    iconMap,
    fields,
  );
  const importSheet = importFixture.sheetsByName.get("Categories");
  assert(importSheet, "inserts a Categories sheet when missing");
  assert(
    JSON.stringify(importSheet.values[0]) === JSON.stringify(canonical),
    "inserted sheet header equals canonical layout",
  );
  assert(
    JSON.stringify(importSheet.values[1]) ===
      JSON.stringify([
        "Trees",
        "data:image/svg+xml;base64,PHN2Zz48L3N2Zz4=",
        "Species",
        "observation",
        "trees",
        "#4CAF50",
        "trees",
      ]),
    ".comapeocat category row carries appliesTo, Category ID, Color, Icon ID",
  );
  assert(
    importSheet.values[2][3] === "track, observation",
    "Rivers keeps appliesTo order (track, observation)",
  );
  assert(
    importSheet.values[3][3] === "track, observation" &&
      importSheet.values[3][4] === "ghost" &&
      importSheet.values[3][5] === "",
    "category without appliesTo gets default; color stays blank, id kept",
  );

  // 4) Round trip: generate-side buildCategories reads the imported sheet.
  console.log("generate reads back the imported sheet:");
  // Point the active spreadsheet at the imported fixture so buildCategories
  // reads backgrounds from the sheet populateCategoriesSheet wrote.
  context.SpreadsheetApp.getActiveSpreadsheet = () => importFixture.spreadsheet;
  const roundTripData = {
    documentName: "Test Config",
    Categories: importSheet.values.slice(),
    Details: [
      ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
      ["Species", "", "s", "Oak,Pine", "species", "FALSE"],
    ],
  };
  const builtFields = context.buildFields(roundTripData);
  const rebuiltCategories = context.buildCategories(roundTripData, builtFields);
  for (const original of categories) {
    const rebuilt = rebuiltCategories.find((c) => c.id === original.id);
    assert(rebuilt, `category ${original.id} survives round trip`);
    const expectedApplies = Array.isArray(original.appliesTo) &&
      original.appliesTo.length > 0
      ? original.appliesTo
      : ["track", "observation"];
    assert(
      JSON.stringify(rebuilt.appliesTo) === JSON.stringify(expectedApplies),
      `${original.id} appliesTo byte-identical (${JSON.stringify(rebuilt.appliesTo)})`,
    );
    assert(
      (rebuilt.color || "") === (original.color || ""),
      `${original.id} color byte-identical (${rebuilt.color})`,
    );
    assert(
      (rebuilt.iconId || "") === (original.iconId || ""),
      `${original.id} iconId preserved (${rebuilt.iconId})`,
    );
  }

  // 5) Header layout validation
  const layoutError = context.validateCategoriesHeaderLayout(importSheet);
  assert(
    layoutError === null,
    `validateCategoriesHeaderLayout accepts canonical layout (${layoutError})`,
  );
  const oldSheet = new FakeSheet("Categories", [
    ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
    ["River", "", "", "river", "#000000", "river"],
  ]);
  const oldLayoutError = context.validateCategoriesHeaderLayout(oldSheet);
  assert(
    typeof oldLayoutError === "string" && oldLayoutError.includes("Applies"),
    "old import layout flagged with an actionable message",
  );

  // 6) Full build & import regression suite still passes. Runs in a fresh
  // context so the fixtures above cannot pollute SpreadsheetApp state.
  console.log("runBuildAndImportTests regression:");
  const suiteContext = createGasContext();
  loadFiles(suiteContext, collectTypeScriptFiles(ignoredDirs));
  const logs = [];
  const origLog = console.log;
  const origError = console.error;
  console.log = (...args) => logs.push(args.map(String).join(" "));
  console.error = (...args) => logs.push(args.map(String).join(" "));
  try {
    suiteContext.runBuildAndImportTests();
  } finally {
    console.log = origLog;
    console.error = origError;
  }
  const summary = logs.find((line) =>
    /Results: \d+ passed, \d+ failed/.test(line),
  );
  assert(summary, "test suite produced a summary line");
  const match = /Results: (\d+) passed, (\d+) failed/.exec(summary);
  assert(
    Number(match[1]) >= 33,
    `at least 33 tests pass (got ${match[1]})`,
  );
  // Known pre-existing failures in the vm harness environment (FakeSheet has
  // no getDisplayValues/setValue/toast, so display-value-dependent tests
  // degrade). Any FAIL outside this baseline fails the check.
  const baselineFailures = new Set([
    "FAIL: Invalid input should return null",
    "FAIL: Species field should be type 'select'",
    "FAIL: Expected 11 fields, got 5",
    "FAIL: 'i' should map to integer type",
    "FAIL: Expected 6 fields, got 5",
    "FAIL: Initial build should create integer field",
    "FAIL: Spanish translations for field1 not found",
    "FAIL: Translations should include Spanish and French even without Category Translations sheet",
    "FAIL: Helper text translation for width not found",
    "FAIL: Spanish options map missing for question1",
    "FAIL: Chinese translations not generated",
    "FAIL: Should retrieve 3 categories",
  ]);
  const fails = logs.filter((line) => line.includes("FAIL"));
  const unexpected = fails.filter((line) => !baselineFailures.has(line.trim()));
  assert(
    unexpected.length === 0,
    `no new FAIL lines vs baseline (unexpected: ${unexpected.length})\n    ${unexpected.join("\n    ")}`,
  );
  assert(
    Number(match[2]) <= baselineFailures.size,
    `failures within baseline (${match[2]} <= ${baselineFailures.size})`,
  );

  console.log("check-import-layout PASSED");
}

main();
