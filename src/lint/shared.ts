// Import slugify function from utils
/// <reference path="../utils.ts" />

// --- Types -----------------------------------------------------------------

type DriveIconInfo = {
  slug: string | null;
  isSvg: boolean;
  svgContent: string | null;
  errorMessage?: string;
};

type LintSeverity = "error" | "warning" | "advisory";

// --- Constants --------------------------------------------------------------

// LINT_WARNING_BACKGROUND_COLORS is defined in lint-colors.ts (shared with tests)
/**
 * Lint warning font colors.
 *
 * IMPORTANT: See `src/test/testLint.ts` color synchronization table.
 */
// LINT_WARNING_FONT_COLORS is defined in lint-colors.ts (shared with tests)
const LINT_WARNING_FONT_COLORS_WITHOUT_WHITE = LINT_WARNING_FONT_COLORS.filter(
  (color) => color.toUpperCase() !== "#FFFFFF",
);
const LINT_NOTE_PREFIX = "[Lint] ";
const SLUG_COLLISION_LINT_NOTE_PREFIX = `${LINT_NOTE_PREFIX}Slug collision:`;
const SOURCE_OVERWRITE_LINT_NOTE_PREFIX = `${LINT_NOTE_PREFIX}Source value`;

/** Module-level cache for Drive icon info, shared across lint checks to avoid
 *  redundant Drive API calls and keep Drive URL classification consistent. */
const driveIconInfoCache = new Map<string, DriveIconInfo>();

/**
 * Single source of truth for severity → visual style.
 * `background`/`fontColor` of `null` mean "leave that property unchanged".
 * Shared by all four lint-note writers so they can never drift apart.
 */
const LINT_SEVERITY_STYLE: Record<
  LintSeverity,
  { background: string | null; fontColor: string | null }
> = {
  error: { background: LINT_ERROR_BG, fontColor: "red" },
  warning: { background: LINT_WARNING_BG, fontColor: "orange" },
  advisory: { background: LINT_ADVISORY_BG, fontColor: null },
};

// --- Helper functions -------------------------------------------------------

function capitalizeFirstLetter(str: string): string {
  if (!str || typeof str !== "string") return "";
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function validateAndCapitalizeCommaList(value: string): string {
  if (!value || typeof value !== "string") return "";
  return value
    .split(",")
    .map((item) => {
      const trimmed = item.trim();
      // Options may use an explicit "value:label" form. Capitalizing the whole
      // item would corrupt the value prefix (cafe:Café -> Cafe:Café), changing
      // the stored value. Capitalize only the label; keep the value verbatim.
      const colonIndex = trimmed.indexOf(":");
      if (colonIndex > 0) {
        const explicitValue = trimmed.slice(0, colonIndex);
        const label = trimmed.slice(colonIndex + 1).trim();
        return `${explicitValue}:${capitalizeFirstLetter(label)}`;
      }
      return capitalizeFirstLetter(trimmed);
    })
    .filter((item) => item)
    .join(", ");
}

function setInvalidCellBackground(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  col: number,
  color: string,
): void {
  sheet.getRange(row, col).setBackground(color);
}

function clearRangeBackgroundIfMatches(
  range: GoogleAppsScript.Spreadsheet.Range,
  colorsToClear: string[],
): void {
  if (!range) return;
  if (range.getNumRows() === 0 || range.getNumColumns() === 0) return;

  const normalized = colorsToClear.map((color) => color.toUpperCase());
  const backgrounds = range.getBackgrounds();
  let updated = false;

  for (let row = 0; row < backgrounds.length; row++) {
    for (let col = 0; col < backgrounds[row].length; col++) {
      const background = backgrounds[row][col];
      if (background && normalized.includes(background.toUpperCase())) {
        backgrounds[row][col] = null;
        updated = true;
      }
    }
  }

  if (updated) {
    range.setBackgrounds(backgrounds);
  }
}

function clearRangeFontColorIfMatches(
  range: GoogleAppsScript.Spreadsheet.Range,
  colorsToClear: string[],
): void {
  if (!range) return;
  if (range.getNumRows() === 0 || range.getNumColumns() === 0) return;

  const normalized = colorsToClear.map((color) => color.toUpperCase());
  const fontColors = range.getFontColors();
  let updated = false;

  for (let row = 0; row < fontColors.length; row++) {
    for (let col = 0; col < fontColors[row].length; col++) {
      const fontColor = fontColors[row][col];
      if (fontColor && normalized.includes(fontColor.toUpperCase())) {
        fontColors[row][col] = null;
        updated = true;
      }
    }
  }

  if (updated) {
    range.setFontColors(fontColors);
  }
}

function clearRangeLintNotesWithPrefix(
  range: GoogleAppsScript.Spreadsheet.Range,
  prefix: string,
  fontColorsToClear: string[] = LINT_WARNING_FONT_COLORS,
  backgroundColorsToClear: string[] = LINT_WARNING_BACKGROUND_COLORS,
): void {
  if (!prefix) return; // Empty prefix would match all lines — treat as no-op.
  clearRangeLintNotesWithPrefixes(range, [prefix], fontColorsToClear, backgroundColorsToClear);
}

/**
 * Batch version of clearRangeLintNotesWithPrefix.
 * Removes all lines starting with any of the given prefixes from notes in the range.
 * Reads notes/fontColors/backgrounds only once regardless of how many prefixes are provided.
 */
function clearRangeLintNotesWithPrefixes(
  range: GoogleAppsScript.Spreadsheet.Range,
  prefixes: string[],
  fontColorsToClear: string[] = LINT_WARNING_FONT_COLORS,
  backgroundColorsToClear: string[] = LINT_WARNING_BACKGROUND_COLORS,
): void {
  if (!range) return;
  if (range.getNumRows() === 0 || range.getNumColumns() === 0) return;
  if (!prefixes || prefixes.length === 0) return;

  const normalizedWarningColors = fontColorsToClear.map((color) =>
    color.toUpperCase(),
  );
  const normalizedBackgroundColors = backgroundColorsToClear.map((color) =>
    color.toUpperCase(),
  );
  const notes = range.getNotes();
  const fontColors = range.getFontColors();
  const backgrounds = range.getBackgrounds();
  let notesUpdated = false;
  let fontColorsUpdated = false;
  let backgroundsUpdated = false;

  for (let row = 0; row < notes.length; row++) {
    for (let col = 0; col < notes[row].length; col++) {
      const note = notes[row][col];
      if (!note) continue;

      const noteLines = note.split("\n");
      const remainingLines = noteLines.filter(
        (line) => !prefixes.some((prefix) => line.startsWith(prefix)),
      );

      if (remainingLines.length === noteLines.length) continue;

      let start = 0;
      while (start < remainingLines.length && remainingLines[start] === "") {
        start++;
      }

      let end = remainingLines.length;
      while (end > start && remainingLines[end - 1] === "") {
        end--;
      }

      const remainingNote = remainingLines.slice(start, end).join("\n");

      notes[row][col] = remainingNote;
      notesUpdated = true;

      const fontColor = fontColors[row][col];
      if (
        remainingNote === "" &&
        fontColor &&
        normalizedWarningColors.includes(fontColor.toUpperCase())
      ) {
        fontColors[row][col] = null;
        fontColorsUpdated = true;
      }

      const background = backgrounds[row][col];
      if (
        remainingNote === "" &&
        background &&
        normalizedBackgroundColors.includes(background.toUpperCase())
      ) {
        backgrounds[row][col] = null;
        backgroundsUpdated = true;
      }
    }
  }

  if (notesUpdated) {
    range.setNotes(notes);
  }

  if (fontColorsUpdated) {
    range.setFontColors(fontColors);
  }

  if (backgroundsUpdated) {
    range.setBackgrounds(backgrounds);
  }
}

/**
 * Standardized lint note writer. Sets a [Lint]-prefixed note on a cell and applies
 * severity-appropriate background and font colors so that cleanup and UI behavior
 * stay consistent across all lint checks.
 *
 * severity → background / font color mapping:
 *   error    → LINT_ERROR_BG / red
 *   warning  → LINT_WARNING_BG / orange
 *   advisory → LINT_ADVISORY_BG / (default)
 */
function setLintNote(
  cell: GoogleAppsScript.Spreadsheet.Range,
  message: string,
  severity: LintSeverity,
): void {
  cell.setNote(`${LINT_NOTE_PREFIX}${message}`);

  const style = LINT_SEVERITY_STYLE[severity];
  if (style.background !== null) cell.setBackground(style.background);
  if (style.fontColor !== null) cell.setFontColor(style.fontColor);
}

/**
 * Appends a lint note to a cell, preserving any existing note (user-authored or lint).
 * Uses the same severity-based styling as setLintNote but concatenates messages
 * when a note already exists, preventing overwrites from sequential lint passes.
 *
 * severity escalation: if the existing note is a lower severity than the new one,
 * the visual styling is upgraded to match the highest severity present.
 */
function appendLintNote(
  cell: GoogleAppsScript.Spreadsheet.Range,
  message: string,
  severity: LintSeverity,
): void {
  const existingNote = cell.getNote() || "";
  const newMessage = `${LINT_NOTE_PREFIX}${message}`;

  if (existingNote) {
    cell.setNote(`${existingNote}\n${newMessage}`);
  } else {
    cell.setNote(newMessage);
  }

  // Apply severity styling (upgrade if higher severity than current)
  const currentBg = cell.getBackground().toUpperCase();
  // Check for error-level backgrounds. LINT_ERROR_BG is the standard lint error
  // background; LINT_CRITICAL_BG is used by validateSheetConsistency for critical
  // primary-column mismatches and must also be treated as error-level so
  // that appendLintNote("warning") does not overwrite it.
  const isAlreadyError =
    currentBg === LINT_ERROR_BG || currentBg === LINT_CRITICAL_BG;
  const isAlreadyWarning = currentBg === LINT_WARNING_BG;

  switch (severity) {
    case "error":
      // Preserve LINT_CRITICAL_BG critical-mismatch styling; only set LINT_ERROR_BG if not
      // already at a higher visual severity. LINT_CRITICAL_BG cells use white text,
      // so skip fontColor too to keep them readable.
      if (currentBg !== LINT_CRITICAL_BG) {
        cell.setBackground(LINT_SEVERITY_STYLE.error.background!);
        cell.setFontColor(LINT_SEVERITY_STYLE.error.fontColor!);
      }
      break;
    case "warning":
      // Only set warning styling if not already at error level
      if (!isAlreadyError) {
        cell.setBackground(LINT_SEVERITY_STYLE.warning.background!);
        cell.setFontColor(LINT_SEVERITY_STYLE.warning.fontColor!);
      }
      break;
    case "advisory":
      // Only set advisory styling if no higher severity is present
      if (!isAlreadyError && !isAlreadyWarning) {
        cell.setBackground(LINT_SEVERITY_STYLE.advisory.background!);
      }
      break;
  }
}

/**
 * Combined helper that clears all lint-managed visual artifacts from a range:
 * background colors, font colors, and [Lint]-prefixed notes.
 * Range-scoped so category color backgrounds in column A are preserved.
 */
function clearLintArtifacts(
  range: GoogleAppsScript.Spreadsheet.Range,
): void {
  if (!range) return;
  if (range.getNumRows() === 0 || range.getNumColumns() === 0) return;

  clearRangeBackgroundIfMatches(range, LINT_WARNING_BACKGROUND_COLORS);
  clearRangeFontColorIfMatches(range, LINT_WARNING_FONT_COLORS);
  clearRangeLintNotesWithPrefix(range, LINT_NOTE_PREFIX);
}

function clearSourceOverwriteLintArtifacts(
  range: GoogleAppsScript.Spreadsheet.Range,
): void {
  clearRangeLintNotesWithPrefix(
    range,
    SOURCE_OVERWRITE_LINT_NOTE_PREFIX,
    LINT_WARNING_FONT_COLORS_WITHOUT_WHITE,
  );
}

/**
 * Like setLintNote but does NOT change the cell background.
 * Use this for columns where user-set backgrounds must be preserved
 * (e.g., Categories column A category colors).
 */
function setLintNotePreserveBackground(
  cell: GoogleAppsScript.Spreadsheet.Range,
  message: string,
  severity: LintSeverity,
): void {
  cell.setNote(`${LINT_NOTE_PREFIX}${message}`);

  // Font color only — background is intentionally left untouched.
  const { fontColor } = LINT_SEVERITY_STYLE[severity];
  if (fontColor !== null) cell.setFontColor(fontColor);
}

/**
 * Like appendLintNote but does NOT change the cell background.
 * Use this for columns where user-set backgrounds must be preserved
 * (e.g., Categories column A category colors).
 */
function appendLintNotePreserveBackground(
  cell: GoogleAppsScript.Spreadsheet.Range,
  message: string,
  severity: LintSeverity,
): void {
  const existingNote = cell.getNote() || "";
  const newMessage = `${LINT_NOTE_PREFIX}${message}`;

  if (existingNote) {
    cell.setNote(`${existingNote}\n${newMessage}`);
  } else {
    cell.setNote(newMessage);
  }

  // Apply font color only (no background).  Detect existing error state
  // via both background (#FFC7CE from normal setLintNote) AND font color
  // (red from appendLintNotePreserveBackground on a preserved-bg cell).
  const currentBg = cell.getBackground().toUpperCase();
  const currentFont = (cell.getFontColor() || "").toUpperCase();
  const isAlreadyError =
    currentBg === LINT_ERROR_BG ||
    currentFont === "RED" ||
    currentFont === LINT_CRITICAL_FONT;

  switch (severity) {
    case "error":
      cell.setFontColor(LINT_SEVERITY_STYLE.error.fontColor!);
      break;
    case "warning":
      if (!isAlreadyError) {
        cell.setFontColor(LINT_SEVERITY_STYLE.warning.fontColor!);
      }
      break;
    case "advisory":
      break;
  }
}

function isEmptyOrWhitespace(value: any): boolean {
  return (
    value === undefined ||
    value === null ||
    (typeof value === "string" && value.trim() === "")
  );
}

function cleanWhitespaceOnlyCells(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  startRow: number,
  startCol: number,
  numRows: number,
  numCols: number,
): void {
  const range = sheet.getRange(startRow, startCol, numRows, numCols);
  const values = range.getValues();
  let changesMade = false;

  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const value = values[i][j];
      if (typeof value === "string" && value.trim() === "" && value !== "") {
        values[i][j] = "";
        changesMade = true;
      }
    }
  }

  if (changesMade) {
    range.setValues(values);
    getScopedLogger("LintWhitespace").info(`Cleaned whitespace-only cells in ${sheet.getName()}`);
  }
}

function extractDriveFileId(url: string): string | null {
  if (!url) return null;
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function isSupportedSvgDataUri(dataUri: string): boolean {
  if (!dataUri.toLowerCase().startsWith("data:image/svg+xml")) return false;

  try {
    let svgText: string;
    if (dataUri.includes(";base64,")) {
      const base64 = dataUri.split(";base64,")[1];
      if (!base64) return false;
      svgText = Utilities.newBlob(
        Utilities.base64Decode(base64),
      ).getDataAsString();
    } else {
      const commaIndex = dataUri.indexOf(",");
      if (commaIndex === -1) return false;
      svgText = decodeURIComponent(dataUri.substring(commaIndex + 1));
    }
    return svgText.trim().startsWith("<svg");
  } catch (_error) {
    return false;
  }
}

function columnNumberToLetter(columnNumber: number): string {
  let dividend = columnNumber;
  let columnName = "";
  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }
  return columnName;
}

function checkForDuplicates(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  columnIndex: number,
  startRow: number = 2,
  preserveBackground: boolean = false,
): void {
  const lastRow = sheet.getLastRow();
  if (lastRow <= startRow) return;

  const range = sheet.getRange(
    startRow,
    columnIndex,
    lastRow - startRow + 1,
    1,
  );
  if (preserveBackground) {
    clearRangeLintNotesWithPrefix(
      range,
      LINT_NOTE_PREFIX,
      LINT_WARNING_FONT_COLORS,
      [],
    );
    clearRangeFontColorIfMatches(range, LINT_WARNING_FONT_COLORS);
  } else {
    clearRangeLintNotesWithPrefix(range, LINT_NOTE_PREFIX);
    clearRangeBackgroundIfMatches(range, LINT_WARNING_BACKGROUND_COLORS);
    clearRangeFontColorIfMatches(range, LINT_WARNING_FONT_COLORS);
  }
  const values = range
    .getValues()
    .map((row) => row[0].toString().trim().toLowerCase());
  const duplicates = new Map<string, number[]>();

  values.forEach((value, index) => {
    if (value === "") return;

    if (!duplicates.has(value)) {
      duplicates.set(value, [index + startRow]);
    } else {
      duplicates.get(value)?.push(index + startRow);
    }
  });

  // Highlight duplicates
  duplicates.forEach((rows, value) => {
    if (rows.length > 1) {
      getScopedLogger("LintDuplicates").info(
        'Found duplicate value "' + value + '" in rows: ' + rows.join(", "),
      );
      const otherRowsStr = rows.join(", ");
      for (const row of rows) {
        const cell = sheet.getRange(row, columnIndex);
        if (preserveBackground) {
          // Use preserve-background variant to avoid overwriting user-managed
          // colors (e.g., Categories column A category colors read by builder).
          setLintNotePreserveBackground(
            cell,
            `Duplicate value "${value}" found in rows: ${otherRowsStr}`,
            "error",
          );
        } else {
          setLintNote(
            cell,
            `Duplicate value "${value}" found in rows: ${otherRowsStr}`,
            "error",
          );
        }
      }
    }
  });
}

/**
 * Pre-cleanup inspection pass for ID columns (column E, index 4).
 * Reads raw values BEFORE cleanWhitespaceOnlyCells() runs so that
 * whitespace-only IDs (which would be erased by cleanup) are captured.
 *
 * Returns a Map of row number → raw ID value for cells that contain
 * only whitespace characters (not empty, but whitespace-only).
 */
function inspectRawIds(
  sheetName: string,
  columnEIndex: number,
): Map<number, string> {
  const result = new Map<number, string>();
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return result;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return result;

  const range = sheet.getRange(2, columnEIndex, lastRow - 1, 1);
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    const value = values[i][0];
    // Capture cells that are strings with content but only whitespace
    if (
      typeof value === "string" &&
      value.length > 0 &&
      value.trim() === ""
    ) {
      result.set(i + 2, value); // +2 because data starts at row 2
    }
  }

  return result;
}

/**
 * Task 6: Checks manual ID hygiene in column E.
 * - Warns on whitespace-only IDs (from pre-cleanup capture)
 * - Warns on manually entered IDs that are not slug-safe
 * - Skips blank/empty IDs (builder auto-generates these)
 */
function checkManualIdHygiene(
  sheetName: string,
  columnEIndex: number,
  rawIds: Map<number, string>,
): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // NOTE: Do NOT blanket-clear column E here. Duplicate-ID checks (Tasks 1/2)
  // have already set their annotations on column E. We only clear notes/fonts
  // for rows we are about to annotate, to avoid wiping prior findings.
  const logger = getScopedLogger("LintManualIdHygiene");

  // 1. Warn on whitespace-only IDs captured before cleanup
  rawIds.forEach((rawValue, row) => {
    const cell = sheet.getRange(row, columnEIndex);
    appendLintNote(
      cell,
      `Whitespace-only ID "${rawValue}" will be treated as empty and auto-generated.`,
      "warning",
    );
    logger.warn(
      `${sheetName} row ${row}: whitespace-only ID will be treated as empty`,
    );
  });

  // 2. Check all non-empty IDs for slug-safety
  // Skip blank-name rows — builder returns early on those.
  const nameValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const idRange = sheet.getRange(2, columnEIndex, lastRow - 1, 1);
  const idValues = idRange.getValues();
  // Entity IDs must stay ASCII: the .comapeocat packager restricts them to
  // ^[a-zA-Z0-9_-]+$ (package/src/writer.js), so a manually-entered non-ASCII
  // ID would be rejected at build time. Warn on those here.
  const slugSafePattern = /^[a-z0-9]+(-[a-z0-9]+)*$/;

  for (let i = 0; i < idValues.length; i++) {
    const row = i + 2;
    const idValue = String(idValues[i][0] || "").trim();
    const nameValue = String(nameValues[i][0] || "").trim();

    // Skip blank/empty (builder auto-generates)
    if (!idValue) continue;

    // Skip blank-name rows — builder returns early on those.
    if (!nameValue) continue;

    // Skip if this row was already flagged as whitespace-only
    if (rawIds.has(row)) continue;

    // Check slug-safety. Non-ASCII explicit IDs are NOT blocked: the live
    // build accepts them as entered and the schema allows free-string IDs, so
    // users may keep meaningful non-Latin IDs (e.g. Thai). But warn that the
    // standalone comapeocat CLI (package/src/writer.js SAFE_ID_REGEX) rejects
    // them and the import reader cannot read back non-ASCII icon/translation
    // entries — so round-trip and CLI builds are not guaranteed.
    if (!slugSafePattern.test(idValue)) {
      const cell = sheet.getRange(row, columnEIndex);
      // Only blame "non-ASCII" when that's actually true — the pattern also
      // rejects pure-ASCII input (uppercase, spaces, underscores, double
      // hyphens), where the CLI/round-trip caveat below does not apply.
      const isNonAscii = /[^\x00-\x7F]/.test(idValue);
      const message = isNonAscii
        ? `Manual ID "${idValue}" contains non-ASCII characters. The live build accepts it as entered, but the standalone comapeocat CLI rejects such IDs and import cannot read back non-ASCII icon/translation entries. For full toolchain compatibility use lowercase ASCII letters, numbers, and hyphens (e.g., "my-category-id").`
        : `Manual ID "${idValue}" is used as entered by the builder. Recommended format: lowercase letters, numbers, and hyphens (e.g., "my-category-id").`;
      appendLintNote(cell, message, "warning");
      logger.warn(
        `${sheetName} row ${row}: ${isNonAscii ? "non-ASCII" : "non-slug-safe"} ID "${idValue}"`,
      );
    }
  }
}

/**
 * Task 3: Warns when two names in column A produce the same fallback slug,
 * even if their explicit IDs may later differ. This is the earliest visible
 * signal that auto-generated IDs may collide.
 * SEPARATE from checkForDuplicates() which checks exact string matches.
 */
function checkSlugCollisions(
  sheetName: string,
  nameColumnIndex: number,
): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const nameRange = sheet.getRange(2, nameColumnIndex, lastRow - 1, 1);
  // Remove only prior slug-collision warnings so duplicate-name errors from
  // checkForDuplicates() remain intact while stale slug warnings are refreshed.
  // This path preserves user-managed category colors and font choices, so it
  // only strips the note lines and leaves visual formatting untouched.
  clearRangeLintNotesWithPrefix(
    nameRange,
    SLUG_COLLISION_LINT_NOTE_PREFIX,
    [],
    [],
  );

  const names = nameRange.getValues();

  // Build map of slug → rows (original names for display)
  const slugMap = new Map<string, { rows: number[]; originals: string[] }>();
  for (let i = 0; i < names.length; i++) {
    const name = String(names[i][0] || "").trim();
    if (!name) continue;
    const slug = slugify(name);
    if (!slug) continue;

    if (!slugMap.has(slug)) {
      slugMap.set(slug, { rows: [i + 2], originals: [name] });
    } else {
      const entry = slugMap.get(slug)!;
      entry.rows.push(i + 2);
      entry.originals.push(name);
    }
  }

  // Annotate slug collisions with warnings
  slugMap.forEach(({ rows, originals }, slug) => {
    if (rows.length > 1) {
      const logger = getScopedLogger("LintSlugCollisions");
      logger.warn(
        `Slug collision "${slug}" in ${sheetName}: rows ${rows.join(", ")} (names: ${originals.join(", ")})`,
      );
      for (let idx = 0; idx < rows.length; idx++) {
        const row = rows[idx];
        const otherRows = rows.filter((r) => r !== row);
        const otherNames = originals.filter((_, i) => rows[i] !== row);
        const cell = sheet.getRange(row, nameColumnIndex);
        appendLintNotePreserveBackground(
          cell,
          `Slug collision: "${originals[idx]}" → "${slug}" (also produced by "${otherNames.join('", "')}" in row${otherRows.length > 1 ? "s" : ""} ${otherRows.join(", ")})`,
          "warning",
        );
      }
    }
  });
}

/**
 * Task 5: Checks if a sheet is empty (only header row or no rows).
 * Returns true if empty (caller should return early), false otherwise.
 */
function checkEmptySheet(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  sheetDisplayName: string,
): boolean {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    const lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      clearLintArtifacts(sheet.getRange(1, 1, 1, lastCol));
    }
    const cell = sheet.getRange(1, 1);
    clearLintArtifacts(cell);
    setLintNote(
      cell,
      `${sheetDisplayName} sheet is empty. At least one ${sheetDisplayName.toLowerCase()} entry is required for config generation.`,
      "error",
    );
    return true;
  }
  return false;
}

/**
 * Canonical option parser that mirrors the builder's parseOptions() logic.
 * Splits by comma, trims, and for each option uses first-colon value:label format.
 * No-colon entries use slugified label as value.
 */
function parseCanonicalOptions(optionsStr: string): Array<{
  value: string;
  label: string;
  raw: string;
}> {
  if (!optionsStr) return [];

  const opts = optionsStr
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
  if (opts.length === 0) return [];

  return opts.map((opt) => {
    const colonIndex = opt.indexOf(":");
    if (colonIndex > 0) {
      const value = opt.substring(0, colonIndex);
      const label = opt.substring(colonIndex + 1);
      return { value, label, raw: opt };
    }
    // Mirror the builder: fall back to the raw label when the canonical value
    // is empty (emoji-/punctuation-only options) so distinct options never
    // collapse to one empty value. Kept in lockstep with parseOptions().
    return {
      value: canonicalizeOptionValue(opt) || opt,
      label: opt,
      raw: opt,
    };
  });
}

/**
 * Detects ambiguous colon usage in select options that may indicate
 * accidental label splits rather than intentional value:label pairs.
 * Returns an array of warning messages for ambiguous patterns.
 */
function detectAmbiguousColonUsage(
  parsed: Array<{ value: string; label: string; raw: string }>,
): string[] {
  const warnings: string[] = [];

  for (const opt of parsed) {
    // Use the raw option text for colon analysis
    const rawText = opt.raw;
    const colonCount = (rawText.match(/:/g) || []).length;

    // Leading colon with empty value (e.g., ":label")
    if (colonCount > 0 && rawText.startsWith(":")) {
      warnings.push(
        `Option "${rawText}" has an empty value before the colon.`,
      );
      continue;
    }

    // Multiple colons (e.g., "value:label:extra")
    if (colonCount > 1) {
      warnings.push(
        `Option "${rawText}" contains multiple colons. Only the first colon is used to split value:label.`,
      );
      continue;
    }

    // Value part looks like natural language with spaces (e.g., "not applicable:No")
    if (
      opt.value.includes(" ") &&
      opt.label &&
      opt.value.length > 1
    ) {
      warnings.push(
        `Option "${rawText}" has spaces in the value portion ("${opt.value}"), which may be unintentional.`,
      );
    }
  }

  return warnings;
}

function createLintHeaderMap(headers: any[]): Record<string, number> {
  const headerMap: Record<string, number> = {};
  headers.forEach((header, index) => {
    const key = String(header || "")
      .trim()
      .toLowerCase();
    if (key) {
      headerMap[key] = index;
    }
  });
  return headerMap;
}

function getLintColumnIndex(
  headerMap: Record<string, number>,
  ...names: string[]
): number | undefined {
  for (const name of names) {
    const key = name.toLowerCase();
    if (headerMap[key] !== undefined) {
      return headerMap[key];
    }
  }
  return undefined;
}

/**
 * Decodes a data:image/svg+xml URI (plain or base64) to extract the raw SVG
 * content. Returns null if the URI cannot be decoded. Mirrors the decode logic
 * from payloadBuilder's decodeDataSvg() for use in lint checks.
 */
function decodeDataSvgForLint(dataUri: string): string | null {
  try {
    if (dataUri.includes(";base64,")) {
      const base64 = dataUri.split(";base64,")[1];
      const decoded = Utilities.newBlob(
        Utilities.base64Decode(base64),
      ).getDataAsString();
      return decoded;
    }

    const commaIndex = dataUri.indexOf(",");
    if (commaIndex === -1) return null;
    const payload = dataUri.substring(commaIndex + 1);
    return decodeURIComponent(payload);
  } catch (_err) {
    return null;
  }
}

function loadDriveSvgForLint(fileId: string): string | null {
  return getDriveIconInfo(fileId).svgContent;
}

function getDriveIconInfo(fileId: string): DriveIconInfo {
  if (driveIconInfoCache.has(fileId)) return driveIconInfoCache.get(fileId)!;

  let info: DriveIconInfo;
  try {
    const file = DriveApp.getFileById(fileId);
    const fileName = file.getName();
    const mimeType = file.getMimeType();
    const mimeTypeLower = mimeType?.toLowerCase() ?? "";
    const nameWithoutExt = fileName.replace(/\.[^/.]+$/, "");
    const slug = normalizeIconSlug(slugify(nameWithoutExt));
    const isSvgMime = mimeTypeLower.includes("svg");
    let svgContent: string | null = null;

    // Fast path: MIME type already tells us it's SVG, so reading as text is expected.
    if (isSvgMime) {
      svgContent = file.getBlob().getDataAsString().trim();
    } else {
      // Fallback: try reading as text to detect raw SVG content. Binary files
      // (PNG/JPEG) may throw here, so treat those as accessible non-SVG files.
      try {
        const text = file.getBlob().getDataAsString().trim();
        svgContent = text.startsWith("<svg") ? text : null;
      } catch {
        svgContent = null;
      }
    }

    info = { slug: slug || null, isSvg: svgContent !== null, svgContent };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    info = {
      slug: null,
      isSvg: false,
      svgContent: null,
      errorMessage: `Unable to access icon file (Drive ID ${fileId}): ${message}`,
    };
  }

  driveIconInfoCache.set(fileId, info);
  return info;
}

/**
 * Check for whether a string is a recognised icon source that the builder would
 * accept. Data URIs are decoded to verify valid SVG content; Drive URLs are
 * accepted by prefix (full validation is done separately in lintIconsSheet).
 * Used by collision checks that need to know if a row contributes an icon asset.
 */
function hasRecognisedIconSource(iconStr: string): boolean {
  if (!iconStr) return false;
  if (iconStr.startsWith("<svg")) return true;
  if (iconStr.toLowerCase().startsWith("data:image/svg+xml")) {
    // Mirror parseIconSource: must actually decode to valid SVG
    const svg = decodeDataSvgForLint(iconStr);
    return !!svg && svg.trim().startsWith("<svg");
  }
  if (iconStr.startsWith("https://drive.google.com/file/d/")) {
    const fileId = extractDriveFileId(iconStr);
    return fileId ? !!loadDriveSvgForLint(fileId) : false;
  }
  if (/^https?:\/\//i.test(iconStr) && iconStr.toLowerCase().includes(".svg")) return true;
  return false;
}
