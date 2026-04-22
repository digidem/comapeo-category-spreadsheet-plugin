/**
 * Lint parity regression tests
 *
 * Run these from the Apps Script editor to verify lint and builder parsing stay aligned.
 * Usage:
 *   runLintParityTests();
 *
 * Depends on globals from src/builders/payloadBuilder.ts:
 * normalizeFieldTokens(), splitTranslatedOptions() (loaded earlier by GAS alphabetical order).
 */

function testFieldTokenParity(): boolean {
  console.log("=== testFieldTokenParity ===");

  const cases: Array<{ input: any; expected: string[] }> = [
    { input: "River, Forest, Mountain", expected: ["River", "Forest", "Mountain"] },
    { input: "River;Forest;Mountain", expected: ["River", "Forest", "Mountain"] },
    { input: "River\nForest\nMountain", expected: ["River", "Forest", "Mountain"] },
    { input: "River•Forest•Mountain", expected: ["River", "Forest", "Mountain"] },
    { input: "River·Forest·Mountain", expected: ["River", "Forest", "Mountain"] },
    { input: "River，Forest，Mountain", expected: ["River", "Forest", "Mountain"] },
    { input: "  River  ,  Forest  ", expected: ["River", "Forest"] },
    { input: "", expected: [] },
    { input: undefined, expected: [] },
    { input: null, expected: [] },
    { input: ["River", "Forest"], expected: ["River", "Forest"] },
    { input: "Single", expected: ["Single"] },
  ];

  for (const { input, expected } of cases) {
    const result = normalizeFieldTokens(input);
    const matches =
      result.length === expected.length &&
      result.every((value, index) => value === expected[index]);
    if (!matches) {
      console.error(
        `FAIL: normalizeFieldTokens(${JSON.stringify(input)}) => ` +
          `[${result.map((value) => `"${value}"`).join(", ")}] expected ` +
          `[${expected.map((value) => `"${value}"`).join(", ")}]`,
      );
      return false;
    }
  }

  console.log(`PASS: ${cases.length}/${cases.length} field token parity cases passed`);
  return true;
}

function testCanonicalOptionParity(): boolean {
  console.log("=== testCanonicalOptionParity ===");

  const cases: Array<{
    input: string;
    expected: Array<{ value: string; label: string }>;
  }> = [
    {
      input: "Oak, Pine, Maple",
      expected: [
        { value: "oak", label: "Oak" },
        { value: "pine", label: "Pine" },
        { value: "maple", label: "Maple" },
      ],
    },
    {
      input: "oak:Oak Tree, pine:Pine Tree",
      expected: [
        { value: "oak", label: "Oak Tree" },
        { value: "pine", label: "Pine Tree" },
      ],
    },
    {
      input: "  A  ,  B  ",
      expected: [
        { value: "a", label: "A" },
        { value: "b", label: "B" },
      ],
    },
    {
      input: "single",
      expected: [{ value: "single", label: "single" }],
    },
    {
      input: "",
      expected: [],
    },
    {
      input: "val:Label",
      expected: [{ value: "val", label: "Label" }],
    },
    {
      input: "key:Label:Extra",
      expected: [{ value: "key", label: "Label:Extra" }],
    },
  ];

  for (const { input, expected } of cases) {
    const result = parseCanonicalOptions(input);
    const matches =
      result.length === expected.length &&
      result.every(
        (entry, index) =>
          entry.value === expected[index].value &&
          entry.label === expected[index].label,
      );
    if (!matches) {
      console.error(
        `FAIL: parseCanonicalOptions("${input}") => ` +
          `[${result
            .map((entry) => `{value:"${entry.value}",label:"${entry.label}"}`)
            .join(", ")}] expected ` +
          `[${expected
            .map((entry) => `{value:"${entry.value}",label:"${entry.label}"}`)
            .join(", ")}]`,
      );
      return false;
    }
  }

  console.log(`PASS: ${cases.length}/${cases.length} canonical option parity cases passed`);
  return true;
}

function testTranslationDelimiterParity(): boolean {
  console.log("=== testTranslationDelimiterParity ===");

  const cases: Array<{ input: string; expectedCount: number }> = [
    { input: "School, Hospital, Farm", expectedCount: 3 },
    { input: "School;Hospital;Farm", expectedCount: 3 },
    { input: "学校、病院、農場", expectedCount: 3 },
    { input: "学校，医院，农场", expectedCount: 3 },
    { input: "A,B;C，D、E", expectedCount: 5 },
    { input: "  A  ,  B  ", expectedCount: 2 },
    { input: "", expectedCount: 0 },
    { input: "Single", expectedCount: 1 },
  ];

  for (const { input, expectedCount } of cases) {
    const result = splitTranslatedOptions(input);
    if (result.length !== expectedCount) {
      console.error(
        `FAIL: splitTranslatedOptions("${input}") => ${result.length} parts ` +
          `[${result.map((value) => `"${value}"`).join(", ")}] expected ${expectedCount}`,
      );
      return false;
    }
  }

  console.log(`PASS: ${cases.length}/${cases.length} translation delimiter parity cases passed`);
  return true;
}

type MockLintCall = {
  row: number;
  col: number;
  message: string;
  severity: "error" | "warning" | "advisory";
};

type MockCellState = {
  value: any;
  note: string;
  background: string | null;
  fontColor: string | null;
};

type MockSheetState = {
  cells: MockCellState[][];
};

function getSeverityRank(
  background: string | null,
): 0 | 1 | 2 | 3 {
  switch ((background || "").toUpperCase()) {
    case "#FFFFCC":
      return 1;
    case "#FFF2CC":
      return 2;
    case "#FFC7CE":
    case "#FF0000":
      return 3;
    default:
      return 0;
  }
}

function setMockLintStyle(
  cellState: MockCellState,
  severity: "error" | "warning" | "advisory",
): void {
  switch (severity) {
    case "error":
      cellState.background = "#FFC7CE";
      cellState.fontColor = "red";
      break;
    case "warning":
      cellState.background = "#FFF2CC";
      cellState.fontColor = "orange";
      break;
    case "advisory":
      cellState.background = "#FFFFCC";
      break;
  }
}

function appendMockLintStyle(
  cellState: MockCellState,
  severity: "error" | "warning" | "advisory",
): void {
  const currentRank = getSeverityRank(cellState.background);

  switch (severity) {
    case "error":
      // Preserve #FF0000 critical-mismatch styling — mirrors production
      // appendLintNote() which skips background/font changes for #FF0000 cells.
      if ((cellState.background || "").toUpperCase() !== "#FF0000") {
        setMockLintStyle(cellState, "error");
      }
      break;
    case "warning":
      if (currentRank < 3) {
        setMockLintStyle(cellState, "warning");
      }
      break;
    case "advisory":
      if (currentRank < 2) {
        setMockLintStyle(cellState, "advisory");
      }
      break;
  }
}

function appendMockLintNote(
  cellState: MockCellState,
  message: string,
): void {
  const prefix = "[Lint] ";
  const newMessage = `${prefix}${message}`;
  if (cellState.note) {
    cellState.note = `${cellState.note}\n${newMessage}`;
  } else {
    cellState.note = newMessage;
  }
}

function setMockLintNote(
  cellState: MockCellState,
  message: string,
): void {
  cellState.note = `[Lint] ${message}`;
}

function getMockCellState(target: any): MockCellState {
  if (target && typeof target.__getCellState === "function") {
    return target.__getCellState();
  }
  return target as MockCellState;
}

function runWithMockedLintSpreadsheet(
  sheets: Record<string, any[][]>,
  callback: (lintCalls: MockLintCall[]) => void,
): void {
  const globalScope = globalThis as any;
  const originalSpreadsheetApp = globalScope.SpreadsheetApp;
  const originalSetLintNote = globalScope.setLintNote;
  const originalAppendLintNote = globalScope.appendLintNote;
  const lintCalls: MockLintCall[] = [];
  const sheetStates = new Map<
    string,
    MockSheetState
  >();

  Object.entries(sheets).forEach(([name, values]) => {
    const maxColumns = Math.max(...values.map((currentRow) => currentRow.length), 0);
    const normalizedValues = values.map((currentRow) => {
      const nextRow = currentRow.slice();
      while (nextRow.length < maxColumns) {
        nextRow.push("");
      }
      return nextRow;
    });

    sheetStates.set(name, {
      cells: normalizedValues.map((currentRow) =>
        currentRow.map(
          (value): MockCellState => ({
            value,
            note: "",
            background: null,
            fontColor: null,
          }),
        ),
      ),
    });
  });

  const createRange = (
    sheetState: MockSheetState,
    row: number,
    col: number,
    numRows: number,
    numCols: number,
  ): any => {
    const emptyCell: MockCellState = { value: "", note: "", background: null, fontColor: null };

    const getCell = (rowOffset: number, colOffset: number): MockCellState => {
      // Production GAS getDataRange() returns a 1x1 range with an empty
      // cell even on a sheet with zero rows.  Mirror that behaviour so
      // callers don't get a spurious mock throw on empty sheets.
      if (sheetState.cells.length === 0) {
        return emptyCell;
      }
      if (!sheetState.cells[row - 1 + rowOffset]) {
        throw new Error(`Mock sheet row ${row + rowOffset} is missing`);
      }
      const cell = sheetState.cells[row - 1 + rowOffset][col - 1 + colOffset];
      if (!cell) {
        throw new Error(`Mock sheet cell ${row + rowOffset},${col + colOffset} is missing`);
      }
      return cell;
    };

    const getMatrix = <T>(selector: (cell: MockCellState) => T, fallback: T): T[][] => {
      const values: T[][] = [];
      for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
        const currentRow: T[] = [];
        for (let colOffset = 0; colOffset < numCols; colOffset++) {
          const cell = sheetState.cells[row - 1 + rowOffset]?.[col - 1 + colOffset];
          currentRow.push(cell ? selector(cell) : fallback);
        }
        values.push(currentRow);
      }
      return values;
    };

    const setMatrix = <T>(setter: (cell: MockCellState, value: T) => void, nextValues: T[][]): void => {
      for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
        for (let colOffset = 0; colOffset < numCols; colOffset++) {
          setter(getCell(rowOffset, colOffset), nextValues[rowOffset][colOffset]);
        }
      }
    };

    return {
      row,
      col,
      numRows,
      numCols,
      __getCellState(): MockCellState {
        return getCell(0, 0);
      },
      getNumRows(): number {
        return numRows;
      },
      getNumColumns(): number {
        return numCols;
      },
      getValue(): any {
        return getCell(0, 0).value;
      },
      getValues(): any[][] {
        return getMatrix((cell) => cell.value, "");
      },
      getNote(): string {
        return getCell(0, 0).note || "";
      },
      setNote(note: string): any {
        getCell(0, 0).note = note;
        return this;
      },
      getNotes(): string[][] {
        return getMatrix((cell) => cell.note, "");
      },
      setNotes(notes: string[][]): any {
        setMatrix((cell, note) => {
          cell.note = note;
        }, notes);
        return this;
      },
      getBackground(): string {
        return getCell(0, 0).background || "#FFFFFF";
      },
      setBackground(color: string | null): any {
        getCell(0, 0).background = color;
        return this;
      },
      getBackgrounds(): (string | null)[][] {
        return getMatrix((cell) => cell.background, null);
      },
      setBackgrounds(backgrounds: (string | null)[][]): any {
        setMatrix((cell, color) => {
          cell.background = color;
        }, backgrounds);
        return this;
      },
      setFontColor(color: string | null): any {
        getCell(0, 0).fontColor = color;
        return this;
      },
      getFontColor(): string | null {
        return getCell(0, 0).fontColor;
      },
      getFontColors(): (string | null)[][] {
        return getMatrix((cell) => cell.fontColor, null);
      },
      setFontColors(fontColors: (string | null)[][]): any {
        setMatrix((cell, color) => {
          cell.fontColor = color;
        }, fontColors);
        return this;
      },
    };
  };

  const spreadsheet = {
    getSheetByName(name: string): any {
      const sheetState = sheetStates.get(name);
      if (!sheetState) return null;

      return {
        getLastRow(): number {
          return sheetState.cells.length;
        },
        getLastColumn(): number {
          return Math.max(
            ...sheetState.cells.map((currentRow) => currentRow.length),
            0,
          );
        },
        getDataRange(): any {
          const lastRow = this.getLastRow();
          const lastCol = this.getLastColumn();
          return createRange(
            sheetState,
            1,
            1,
            Math.max(lastRow, 1),
            Math.max(lastCol, 1),
          );
        },
        getRange(row: number, col: number, numRows?: number, numCols?: number): any {
          return createRange(sheetState, row, col, numRows || 1, numCols || 1);
        },
      };
    },
  };

  globalScope.SpreadsheetApp = {
    getActiveSpreadsheet(): any {
      return spreadsheet;
    },
  };
  globalScope.setLintNote = (
    cell: any,
    message: string,
    severity: "error" | "warning" | "advisory",
  ): void => {
    const cellState = getMockCellState(cell);
    setMockLintNote(cellState, message);
    setMockLintStyle(cellState, severity);
    lintCalls.push({ row: cell.row, col: cell.col, message, severity });
  };
  globalScope.appendLintNote = (
    cell: any,
    message: string,
    severity: "error" | "warning" | "advisory",
  ): void => {
    const cellState = getMockCellState(cell);
    appendMockLintNote(cellState, message);
    appendMockLintStyle(cellState, severity);
    lintCalls.push({ row: cell.row, col: cell.col, message, severity });
  };

  try {
    callback(lintCalls);
  } finally {
    globalScope.SpreadsheetApp = originalSpreadsheetApp;
    globalScope.setLintNote = originalSetLintNote;
    globalScope.appendLintNote = originalAppendLintNote;
  }
}

function testAppliesObservationCoverageParity(): boolean {
  console.log("=== testAppliesObservationCoverageParity ===");

  try {
    runWithMockedLintSpreadsheet(
      {
        Categories: [
          ["Name", "Icon", "Fields", "Applies"],
          ["Track One", "", "field-a", "track"],
          ["Track Two", "", "field-b", "track"],
        ],
      },
      (lintCalls) => {
        validateAppliesColumn();

        const observationErrors = lintCalls.filter(
          (call) =>
            call.severity === "error" &&
            call.message.includes('No category includes "observation"'),
        );
        if (observationErrors.length !== 1) {
          throw new Error(
            `Expected one observation coverage error for all-track sheets, got ${observationErrors.length}`,
          );
        }
        if (lintCalls.some((call) => call.message.includes('No category includes "track"'))) {
          throw new Error(
            "Track coverage warning should not be emitted when track is present",
          );
        }
      },
    );

    console.log("PASS: All-track Applies coverage still surfaces the builder's missing-observation failure");
    return true;
  } catch (error) {
    console.error(`FAIL: ${(error as Error).message}`);
    return false;
  }
}

function testAppliesTokenPrefixParity(): boolean {
  console.log("=== testAppliesTokenPrefixParity ===");

  try {
    runWithMockedLintSpreadsheet(
      {
        Categories: [
          ["Name", "Icon", "Fields", "Applies"],
          ["Observation Prefix", "", "field-a", "obs"],
          ["Track Prefix", "", "field-b", "tr"],
          ["Combined Prefixes", "", "field-c", "o, t"],
        ],
      },
      (lintCalls) => {
        validateAppliesColumn();

        if (
          lintCalls.some((call) => call.message.includes("Unrecognized Applies token"))
        ) {
          throw new Error("Valid observation/track prefixes should not be reported as unrecognized");
        }
        if (
          lintCalls.some((call) => call.message.includes('No category includes "observation"'))
        ) {
          throw new Error("Observation coverage should be detected from prefix tokens");
        }
        if (lintCalls.some((call) => call.message.includes('No category includes "track"'))) {
          throw new Error("Track coverage should be detected from prefix tokens");
        }
      },
    );

    console.log("PASS: Applies token prefixes stay aligned with the builder");
    return true;
  } catch (error) {
    console.error(`FAIL: ${(error as Error).message}`);
    return false;
  }
}

function testAppliesHeaderDetectionParity(): boolean {
  console.log("=== testAppliesHeaderDetectionParity ===");

  try {
    runWithMockedLintSpreadsheet(
      {
        Categories: [
          ["Name", "Icon", "Fields", "Application"],
          ["Category", "", "field-a", "track"],
        ],
      },
      (lintCalls) => {
        validateAppliesColumn();

        const missingHeaderWarnings = lintCalls.filter(
          (call) =>
            call.row === 1 &&
            call.col === 1 &&
            call.severity === "warning" &&
            call.message.includes('No "Applies" header found'),
        );
        if (missingHeaderWarnings.length !== 1) {
          throw new Error(
            `Expected one missing Applies header warning, got ${missingHeaderWarnings.length}`,
          );
        }
      },
    );

    console.log("PASS: Applies header detection matches the builder's accepted names");
    return true;
  } catch (error) {
    console.error(`FAIL: ${(error as Error).message}`);
    return false;
  }
}

function testLintAppendAndClearSemantics(): boolean {
  console.log("=== testLintAppendAndClearSemantics ===");

  try {
    runWithMockedLintSpreadsheet(
      {
        Categories: [["Name"]],
      },
      () => {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
          "Categories",
        );
        if (!sheet) {
          throw new Error("Expected mock sheet to exist");
        }

        const cell = sheet.getRange(1, 1);
        setLintNote(cell, "First warning", "warning");
        appendLintNote(cell, "Second advisory", "advisory");
        if (cell.getNote() !== "[Lint] First warning\n[Lint] Second advisory") {
          throw new Error(`Unexpected stacked lint note: ${cell.getNote()}`);
        }
        if (cell.getBackground() !== "#FFF2CC" || cell.getFontColor() !== "orange") {
          throw new Error(
            `Warning styling should be preserved after advisory append, got ${cell.getBackground()} / ${cell.getFontColor()}`,
          );
        }

        appendLintNote(cell, "Third error", "error");
        if (cell.getBackground() !== "#FFC7CE" || cell.getFontColor() !== "red") {
          throw new Error("Error styling should escalate and persist");
        }

        appendLintNote(cell, "Fourth warning", "warning");
        if (cell.getBackground() !== "#FFC7CE" || cell.getFontColor() !== "red") {
          throw new Error("Error styling should not downgrade on warning append");
        }

        const manualCell = sheet.getRange(1, 1);
        manualCell.setNote("Manual note");
        manualCell.setBackground("#ABCDEF");
        manualCell.setFontColor("#123456");
        appendLintNote(manualCell, "Replaced lint note", "advisory");
        if (manualCell.getNote() !== "Manual note\n[Lint] Replaced lint note") {
          throw new Error(`Non-lint notes should be preserved on append, got: ${manualCell.getNote()}`);
        }
        if (manualCell.getBackground() !== "#FFFFCC") {
          throw new Error("Advisory append should apply lint advisory styling");
        }

        const clearCell = sheet.getRange(1, 1);
        clearCell.setNote("[Lint] first line\nmanual line\n[Lint] second line");
        clearCell.setBackground("#FFF2CC");
        clearCell.setFontColor("orange");
        clearLintArtifacts(clearCell);
        if (clearCell.getNote() !== "manual line") {
          throw new Error(`Expected non-lint note lines to be preserved, got "${clearCell.getNote()}"`);
        }
        if (clearCell.getBackground() !== "#FFFFFF") {
          throw new Error("Lint warning background should be cleared");
        }
        if (clearCell.getFontColor() !== null) {
          throw new Error("Lint warning font color should be cleared");
        }
      },
    );

    console.log("PASS: appendLintNote and clearLintArtifacts semantics are preserved in the mock");
    return true;
  } catch (error) {
    console.error(`FAIL: ${(error as Error).message}`);
    return false;
  }
}

function testTranslationSourceOverwriteCleanupPreservesCriticalWhiteText(): boolean {
  console.log("=== testTranslationSourceOverwriteCleanupPreservesCriticalWhiteText ===");

  try {
    runWithMockedLintSpreadsheet(
      {
        "Category Translations": [
          ["English", "Português"],
          ["River", "Rio"],
          ["River", "Ribeira"],
        ],
      },
      () => {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
          "Category Translations",
        );
        if (!sheet) {
          throw new Error("Expected mock translation sheet to exist");
        }

        const criticalCell = sheet.getRange(2, 1);
        criticalCell.setNote(
          '[Lint] Source value "River" produces the same key as row 3. Later values may overwrite earlier ones.',
        );
        criticalCell.setBackground("#FF0000");
        criticalCell.setFontColor("#FFFFFF");

        checkTranslationSourceOverwrites();

        if (criticalCell.getFontColor() !== "#FFFFFF") {
          throw new Error(
            `Expected critical white font color to be preserved, got ${criticalCell.getFontColor()}`,
          );
        }
        if (criticalCell.getBackground() !== "#FF0000") {
          throw new Error(
            `Expected critical red background to be preserved, got ${criticalCell.getBackground()}`,
          );
        }
        if (!criticalCell.getNote().includes('[Lint] Source value "River"')) {
          throw new Error(
            `Expected source-overwrite lint note to be restored after cleanup, got "${criticalCell.getNote()}"`,
          );
        }
      },
    );

    console.log("PASS: Translation source-overwrite cleanup preserves critical white-on-red styling");
    return true;
  } catch (error) {
    console.error(`FAIL: ${(error as Error).message}`);
    return false;
  }
}

function testCaseInsensitiveDuplicateFieldIdParity(): boolean {
  console.log("=== testCaseInsensitiveDuplicateFieldIdParity ===");

  try {
    runWithMockedLintSpreadsheet(
      {
        Details: [
          ["Name", "Helper Text", "Type", "Options", "ID"],
          ["Field One", "", "t", "", "Field-1"],
          ["Field Two", "", "t", "", "field-1"],
        ],
      },
      (lintCalls) => {
        checkDuplicateDetailIds();

        const duplicateErrors = lintCalls.filter(
          (call) =>
            call.severity === "error" &&
            call.message.includes('Duplicate field ID "'),
        );
        if (duplicateErrors.length !== 2) {
          throw new Error(
            `Expected duplicate field ID errors on both rows, got ${duplicateErrors.length}`,
          );
        }
      },
    );

    console.log("PASS: Lint treats duplicate field IDs case-insensitively");
    return true;
  } catch (error) {
    console.error(`FAIL: ${(error as Error).message}`);
    return false;
  }
}

function testPrimaryLanguageBlankA1RequiresMetadataFallbackError(): boolean {
  console.log("=== testPrimaryLanguageBlankA1RequiresMetadataFallbackError ===");

  try {
    runWithMockedLintSpreadsheet(
      {
        Categories: [
          ["", "Icon", "Fields", "Applies"],
          ["Category", "", "field-a", "track"],
        ],
        Metadata: [["Key", "Value"]],
      },
      (lintCalls) => {
        validatePrimaryLanguageInA1();

        const blankA1Errors = lintCalls.filter(
          (call) => call.row === 1 && call.col === 1 && call.severity === "error",
        );
        if (blankA1Errors.length !== 1) {
          throw new Error(
            `Expected one blank-A1 primary-language error when Metadata has no fallback, got ${blankA1Errors.length}`,
          );
        }
      },
    );

    console.log("PASS: Blank Categories!A1 is rejected when Metadata has no primaryLanguage fallback");
    return true;
  } catch (error) {
    console.error(`FAIL: ${(error as Error).message}`);
    return false;
  }
}

function testPrimaryLanguageBlankA1UsesMetadataFallback(): boolean {
  console.log("=== testPrimaryLanguageBlankA1UsesMetadataFallback ===");

  try {
    runWithMockedLintSpreadsheet(
      {
        Categories: [
          ["", "Icon", "Fields", "Applies"],
          ["Category", "", "field-a", "track"],
        ],
        Metadata: [
          ["Key", "Value"],
          ["primaryLanguage", "English"],
        ],
      },
      (lintCalls) => {
        validatePrimaryLanguageInA1();

        if (lintCalls.length !== 0) {
          throw new Error(
            `Expected no primary-language lint when Metadata provides a fallback, got ${lintCalls.length} call(s)`,
          );
        }

        const categoriesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
          "Categories",
        );
        if (!categoriesSheet) {
          throw new Error("Expected mock Categories sheet to exist");
        }

        const a1Cell = categoriesSheet.getRange(1, 1);
        if (a1Cell.getNote() !== "") {
          throw new Error(`Expected Categories!A1 note to stay empty, got "${a1Cell.getNote()}"`);
        }
      },
    );

    console.log("PASS: Metadata primaryLanguage suppresses blank Categories!A1 lint");
    return true;
  } catch (error) {
    console.error(`FAIL: ${(error as Error).message}`);
    return false;
  }
}

function testMetadataPrimaryLanguageDuplicateParity(): boolean {
  console.log("=== testMetadataPrimaryLanguageDuplicateParity ===");

  try {
    runWithMockedLintSpreadsheet(
      {
        Metadata: [
          ["Key", "Value"],
          ["primaryLanguage", ""],
          ["primaryLanguage", "NotALanguage"],
          ["primaryLanguage", "English"],
        ],
      },
      (lintCalls) => {
        lintMetadataSheet();

        const row3DuplicateWarnings = lintCalls.filter(
          (call) =>
            call.row === 3 &&
            call.col === 2 &&
            call.severity === "warning" &&
            call.message.includes('Duplicate metadata key "primaryLanguage"'),
        );
        if (row3DuplicateWarnings.length !== 0) {
          throw new Error(
            `Expected no duplicate warning on row 3 while earlier primaryLanguage rows are still unresolved, got ${row3DuplicateWarnings.length}`,
          );
        }

        const invalidLanguageErrors = lintCalls.filter(
          (call) =>
            call.row === 3 &&
            call.col === 2 &&
            call.severity === "error" &&
            call.message.includes('Metadata primaryLanguage: "NotALanguage" is not a recognized language name or locale code'),
        );
        if (invalidLanguageErrors.length !== 1) {
          throw new Error(
            `Expected row 3 invalid primaryLanguage error for the first non-empty duplicate, got ${invalidLanguageErrors.length}`,
          );
        }

        const row4FallbackWarnings = lintCalls.filter(
          (call) =>
            call.row === 4 &&
            call.col === 2 &&
            call.severity === "warning" &&
            call.message.includes('Duplicate metadata key "primaryLanguage"') &&
            call.message.includes("only used if all earlier primaryLanguage rows are blank or invalid"),
        );
        if (row4FallbackWarnings.length !== 1) {
          throw new Error(
            `Expected row 4 duplicate primaryLanguage warning describing builder fallback semantics, got ${row4FallbackWarnings.length}`,
          );
        }

        const row4Errors = lintCalls.filter(
          (call) => call.row === 4 && call.col === 2 && call.severity === "error",
        );
        if (row4Errors.length !== 0) {
          throw new Error(`Expected no validation errors on fallback duplicate row 4, got ${row4Errors.length}`);
        }
      },
    );

    console.log("PASS: Duplicate metadata primaryLanguage rows mirror first-valid builder semantics");
    return true;
  } catch (error) {
    console.error(`FAIL: ${(error as Error).message}`);
    return false;
  }
}

function testMetadataPrimaryLanguageLocaleCodeParity(): boolean {
  console.log("=== testMetadataPrimaryLanguageLocaleCodeParity ===");

  try {
    runWithMockedLintSpreadsheet(
      {
        Metadata: [
          ["Key", "Value"],
          ["primaryLanguage", "en"],
          ["primaryLanguage", "pt-BR"],
          ["primaryLanguage", "bad_locale"],
        ],
      },
      (lintCalls) => {
        lintMetadataSheet();

        const row2Errors = lintCalls.filter(
          (call) => call.row === 2 && call.col === 2 && call.severity === "error",
        );
        if (row2Errors.length !== 0) {
          throw new Error(`Expected locale code row 2 to be accepted without errors, got ${row2Errors.length}`);
        }

        const row3IgnoredWarnings = lintCalls.filter(
          (call) =>
            call.row === 3 &&
            call.col === 2 &&
            call.severity === "warning" &&
            call.message.includes('Duplicate metadata key "primaryLanguage"') &&
            call.message.includes("this row is ignored"),
        );
        if (row3IgnoredWarnings.length !== 1) {
          throw new Error(
            `Expected row 3 duplicate warning indicating the locale-code duplicate is ignored, got ${row3IgnoredWarnings.length}`,
          );
        }

        const row3Errors = lintCalls.filter(
          (call) => call.row === 3 && call.col === 2 && call.severity === "error",
        );
        if (row3Errors.length !== 0) {
          throw new Error(`Expected no validation errors on ignored locale-code duplicate row 3, got ${row3Errors.length}`);
        }

        const row4IgnoredWarnings = lintCalls.filter(
          (call) =>
            call.row === 4 &&
            call.col === 2 &&
            call.severity === "warning" &&
            call.message.includes('Duplicate metadata key "primaryLanguage"') &&
            call.message.includes("this row is ignored"),
        );
        if (row4IgnoredWarnings.length !== 1) {
          throw new Error(
            `Expected row 4 duplicate warning indicating the invalid locale duplicate is ignored, got ${row4IgnoredWarnings.length}`,
          );
        }

        const row4Errors = lintCalls.filter(
          (call) => call.row === 4 && call.col === 2 && call.severity === "error",
        );
        if (row4Errors.length !== 0) {
          throw new Error(`Expected no validation errors on ignored invalid duplicate row 4, got ${row4Errors.length}`);
        }
      },
    );

    console.log("PASS: Metadata primaryLanguage accepts locale codes and preserves duplicate-row semantics");
    return true;
  } catch (error) {
    console.error(`FAIL: ${(error as Error).message}`);
    return false;
  }
}

function testIconDuplicateTrackingSkipsMalformedDataUris(): boolean {
  console.log("=== testIconDuplicateTrackingSkipsMalformedDataUris ===");

  try {
    runWithMockedLintSpreadsheet(
      {
        Icons: [
          ["ID", "SVG"],
          ["duplicate-icon", "data:image/svg+xml,not-svg"],
          ["duplicate-icon", "data:image/svg+xml;base64,@@@"],
        ],
      },
      (lintCalls) => {
        lintIconsSheet();

        const duplicateIdErrors = lintCalls.filter(
          (call) =>
            call.col === 1 &&
            call.severity === "error" &&
            call.message.includes('Duplicate icon ID "duplicate-icon"'),
        );
        if (duplicateIdErrors.length !== 0) {
          throw new Error(
            `Expected malformed SVG data URIs to skip duplicate tracking, got ${duplicateIdErrors.length} duplicate error(s)` ,
          );
        }

        const sourceErrors = lintCalls.filter(
          (call) =>
            (call.row === 2 || call.row === 3) &&
            call.col === 2 &&
            call.severity === "error" &&
            call.message.includes("SVG data URI is malformed or does not decode to valid SVG content"),
        );
        if (sourceErrors.length !== 2) {
          throw new Error(`Expected two malformed data URI errors, got ${sourceErrors.length}`);
        }
      },
    );

    console.log("PASS: Malformed SVG data URIs do not create duplicate icon ID lint");
    return true;
  } catch (error) {
    console.error(`FAIL: ${(error as Error).message}`);
    return false;
  }
}

function testIconDuplicateTrackingSkipsInvalidDriveSources(): boolean {
  console.log("=== testIconDuplicateTrackingSkipsInvalidDriveSources ===");

  const globalScope = globalThis as any;
  const originalDriveApp = globalScope.DriveApp;

  try {
    globalScope.DriveApp = {
      getFileById(fileId: string): any {
        if (fileId === "missing-file-id-1234567890123") {
          throw new Error("File not found");
        }
        if (fileId === "png-file-id-1234567890123456") {
          return {
            getName(): string {
              return "icon.png";
            },
            getMimeType(): string {
              return "image/png";
            },
            getBlob(): any {
              return {
                getDataAsString(): string {
                  return "PNG";
                },
              };
            },
          };
        }
        throw new Error(`Unexpected Drive file ID ${fileId}`);
      },
    };

    runWithMockedLintSpreadsheet(
      {
        Icons: [
          ["ID", "SVG"],
          ["drive-duplicate", "https://drive.google.com/file/d/missing-file-id-1234567890123/view"],
          ["drive-duplicate", "https://drive.google.com/file/d/png-file-id-1234567890123456/view"],
        ],
      },
      (lintCalls) => {
        lintIconsSheet();

        const duplicateIdErrors = lintCalls.filter(
          (call) =>
            call.col === 1 &&
            call.severity === "error" &&
            call.message.includes('Duplicate icon ID "drive-duplicate"'),
        );
        if (duplicateIdErrors.length !== 0) {
          throw new Error(
            `Expected invalid Drive icon sources to skip duplicate tracking, got ${duplicateIdErrors.length} duplicate error(s)`,
          );
        }

        const inaccessibleErrors = lintCalls.filter(
          (call) =>
            call.row === 2 &&
            call.col === 2 &&
            call.severity === "error" &&
            call.message.includes("Unable to access icon file (Drive ID missing-file-id-1234567890123)"),
        );
        if (inaccessibleErrors.length !== 1) {
          throw new Error(`Expected one inaccessible Drive icon error, got ${inaccessibleErrors.length}`);
        }

        const nonSvgErrors = lintCalls.filter(
          (call) =>
            call.row === 3 &&
            call.col === 2 &&
            call.severity === "error" &&
            call.message.includes("Google Drive icon file is not SVG"),
        );
        if (nonSvgErrors.length !== 1) {
          throw new Error(`Expected one non-SVG Drive icon error, got ${nonSvgErrors.length}`);
        }
      },
    );

    console.log("PASS: Invalid Drive icon sources do not create duplicate icon ID lint");
    return true;
  } catch (error) {
    console.error(`FAIL: ${(error as Error).message}`);
    return false;
  } finally {
    globalScope.DriveApp = originalDriveApp;
  }
}

function testAppliesMissingHeaderPreservesExistingBodyAnnotations(): boolean {
  console.log("=== testAppliesMissingHeaderPreservesExistingBodyAnnotations ===");

  try {
    runWithMockedLintSpreadsheet(
      {
        Categories: [
          ["Name", "Icon", "Fields", "Application"],
          ["Category", "", "field-a", "track"],
        ],
      },
      (lintCalls) => {
        const categoriesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
          "Categories",
        );
        if (!categoriesSheet) {
          throw new Error("Expected mock Categories sheet to exist");
        }

        const bodyCell = categoriesSheet.getRange(2, 3);
        setLintNote(bodyCell, 'Existing non-Applies lint on Fields cell', "warning");
        lintCalls.length = 0;

        validateAppliesColumn();

        const missingHeaderWarnings = lintCalls.filter(
          (call) =>
            call.row === 1 &&
            call.col === 1 &&
            call.severity === "warning" &&
            call.message.includes('No "Applies" header found'),
        );
        if (missingHeaderWarnings.length !== 1) {
          throw new Error(
            `Expected one missing Applies header warning, got ${missingHeaderWarnings.length}`,
          );
        }

        if (bodyCell.getNote() !== '[Lint] Existing non-Applies lint on Fields cell') {
          throw new Error(`Expected existing body-cell lint note to be preserved, got "${bodyCell.getNote()}"`);
        }
        if (bodyCell.getBackground() !== "#FFF2CC") {
          throw new Error(
            `Expected existing body-cell warning background to be preserved, got ${bodyCell.getBackground()}`,
          );
        }
        if (bodyCell.getFontColor() !== "orange") {
          throw new Error(
            `Expected existing body-cell warning font color to be preserved, got ${bodyCell.getFontColor()}`,
          );
        }

        const staleAppliesCell = categoriesSheet.getRange(2, 4);
        staleAppliesCell.setNote('[Lint] Unrecognized Applies token(s): "legacy"');
        staleAppliesCell.setBackground("#FFF2CC");
        staleAppliesCell.setFontColor("orange");

        validateAppliesColumn();

        if (staleAppliesCell.getNote() !== "") {
          throw new Error(`Expected stale Applies note to be cleared, got "${staleAppliesCell.getNote()}"`);
        }
        if (staleAppliesCell.getBackground() !== "#FFFFFF") {
          throw new Error(
            `Expected stale Applies warning background to be cleared, got ${staleAppliesCell.getBackground()}`,
          );
        }
        if (staleAppliesCell.getFontColor() !== null) {
          throw new Error(
            `Expected stale Applies warning font color to be cleared, got ${staleAppliesCell.getFontColor()}`,
          );
        }

        staleAppliesCell.setNote(
          '[Lint] All Applies tokens are unrecognized ("legacy"). The builder will silently default this row to "observation". Use "observation" or "track" explicitly.',
        );
        staleAppliesCell.setBackground("#FFF2CC");
        staleAppliesCell.setFontColor("orange");

        validateAppliesColumn();

        if (staleAppliesCell.getNote() !== "") {
          throw new Error(
            `Expected stale all-invalid Applies note to be cleared, got "${staleAppliesCell.getNote()}"`,
          );
        }
        if (staleAppliesCell.getBackground() !== "#FFFFFF") {
          throw new Error(
            `Expected stale all-invalid Applies warning background to be cleared, got ${staleAppliesCell.getBackground()}`,
          );
        }
        if (staleAppliesCell.getFontColor() !== null) {
          throw new Error(
            `Expected stale all-invalid Applies warning font color to be cleared, got ${staleAppliesCell.getFontColor()}`,
          );
        }
      },
    );

    console.log("PASS: Missing Applies header warning does not wipe unrelated body-cell lint");
    return true;
  } catch (error) {
    console.error(`FAIL: ${(error as Error).message}`);
    return false;
  }
}

function runLintParityTests(): void {
  console.log("=== Lint Parity Regression Tests ===");

  const tests = [
    { name: "Field Token Parity", fn: testFieldTokenParity },
    { name: "Canonical Option Parity", fn: testCanonicalOptionParity },
    { name: "Translation Delimiter Parity", fn: testTranslationDelimiterParity },
    { name: "Lint Append and Clear Semantics", fn: testLintAppendAndClearSemantics },
    {
      name: "Translation Source Overwrite Cleanup Preserves Critical White Text",
      fn: testTranslationSourceOverwriteCleanupPreservesCriticalWhiteText,
    },
    { name: "Applies Observation Coverage Parity", fn: testAppliesObservationCoverageParity },
    { name: "Applies Token Prefix Parity", fn: testAppliesTokenPrefixParity },
    { name: "Applies Header Detection Parity", fn: testAppliesHeaderDetectionParity },
    {
      name: "Applies Missing Header Preserves Existing Body Annotations",
      fn: testAppliesMissingHeaderPreservesExistingBodyAnnotations,
    },
    {
      name: "Primary Language Blank A1 Requires Metadata Fallback Error",
      fn: testPrimaryLanguageBlankA1RequiresMetadataFallbackError,
    },
    {
      name: "Primary Language Blank A1 Uses Metadata Fallback",
      fn: testPrimaryLanguageBlankA1UsesMetadataFallback,
    },
    {
      name: "Metadata Primary Language Duplicate Parity",
      fn: testMetadataPrimaryLanguageDuplicateParity,
    },
    {
      name: "Metadata Primary Language Locale Code Parity",
      fn: testMetadataPrimaryLanguageLocaleCodeParity,
    },
    {
      name: "Icon Duplicate Tracking Skips Malformed Data URIs",
      fn: testIconDuplicateTrackingSkipsMalformedDataUris,
    },
    {
      name: "Icon Duplicate Tracking Skips Invalid Drive Sources",
      fn: testIconDuplicateTrackingSkipsInvalidDriveSources,
    },
    { name: "Case-Insensitive Duplicate Field ID Parity", fn: testCaseInsensitiveDuplicateFieldIdParity },
  ];

  let passed = 0;
  let failed = 0;

  for (const test of tests) {
    try {
      if (test.fn()) {
        passed++;
      } else {
        failed++;
      }
    } catch (error) {
      failed++;
      console.error(`FAIL: ${test.name} threw an exception - ${error}`);
    }
  }

  console.log(`runLintParityTests: ${passed}/${tests.length} passed`);

  if (failed > 0) {
    throw new Error(`${failed} lint parity test(s) failed`);
  }
}
