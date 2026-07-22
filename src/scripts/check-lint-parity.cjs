#!/usr/bin/env node

// Runs the lint parity regression suite (src/test/testLint.ts: runLintParityTests())
// outside the Apps Script editor, so lint/builder drift is caught in CI instead of
// only when someone remembers to run it manually from the script editor.
//
// Unlike check-gas-boot.cjs, this loads src/test/** too (not just src/lint/** and its
// dependencies) — runLintParityTests() and its mock spreadsheet helpers live there.

const { collectTypeScriptFiles, createGasContext, loadFiles } = require("./gasSimHarness.cjs");

const ignoredDirs = new Set(["scripts"]);

function main() {
  const context = createGasContext();
  const files = collectTypeScriptFiles(ignoredDirs);

  loadFiles(context, files);

  if (typeof context.runLintParityTests !== "function") {
    throw new Error("Lint parity check failed: runLintParityTests is not globally available.");
  }

  context.runLintParityTests();
}

main();
