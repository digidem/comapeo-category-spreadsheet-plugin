#!/usr/bin/env node

const vm = require("vm");
const { collectTypeScriptFiles, createGasContext, loadFiles } = require("./gasSimHarness.cjs");

const ignoredDirs = new Set(["scripts", "test"]);

function assertExpression(context, expression, description) {
  const result = new vm.Script(expression).runInContext(context);
  if (!result) {
    throw new Error(description);
  }
}

function main() {
  const context = createGasContext();
  const files = collectTypeScriptFiles(ignoredDirs);

  loadFiles(context, files);

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
