/// <reference path="shared.ts" />

/**
 * Task 1: Checks for duplicate effective IDs in Categories column E.
 * The "effective ID" is the explicit ID if present, otherwise the slugified name.
 * Annotates duplicate cells with error-level lint notes.
 */
function checkDuplicateCategoryIds(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Categories");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // Clear previous lint artifacts on column E (ID column, index 5)
  const idRange = sheet.getRange(2, 5, lastRow - 1, 1);
  clearLintArtifacts(idRange);

  // Read name column (A=1) and ID column (E=5)
  const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const ids = idRange.getValues();

  // Build map of effective ID → rows. Mirror builder: explicit ID → slugify(name) →
  // `category-${index+1}` fallback. Skip blank-name rows — builder returns early.
  const effectiveIdMap = new Map<string, number[]>();
  for (let i = 0; i < ids.length; i++) {
    const explicitId = String(ids[i][0] || "").trim();
    const name = String(names[i][0] || "").trim();
    if (!name) continue;
    const effectiveId = explicitId || slugify(name) || `category-${i + 1}`;

    if (!effectiveIdMap.has(effectiveId)) {
      effectiveIdMap.set(effectiveId, [i + 2]);
    } else {
      effectiveIdMap.get(effectiveId)?.push(i + 2);
    }
  }

  // Annotate duplicates
  effectiveIdMap.forEach((rows, effectiveId) => {
    if (rows.length > 1) {
      const logger = getScopedLogger("LintDuplicateCategoryIds");
      logger.warn(
        `Duplicate category ID "${effectiveId}" in rows: ${rows.join(", ")}`,
      );
      for (const row of rows) {
        const cell = sheet.getRange(row, 5);
        setLintNote(
          cell,
          `Duplicate category ID "${effectiveId}" (also in rows ${rows.filter((r) => r !== row).join(", ")})`,
          "error",
        );
      }
    }
  });
}

/**
 * Task 4: Validates the Applies column in the Categories sheet.
 * - Resolves the Applies column by header label, matching the builder
 * - Warns when tokens would be ignored by the builder
 * - Warns if no category includes "track" (non-blocking per strict build validation)
 * - Warns if the column is missing and the builder would need to auto-create it
 */
function validateAppliesColumn(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categoriesSheet = spreadsheet.getSheetByName("Categories");
  if (!categoriesSheet) return;

  const lastRow = categoriesSheet.getLastRow();
  const lastCol = categoriesSheet.getLastColumn();
  if (lastCol > 0) {
    const headerRowRange = categoriesSheet.getRange(1, 1, 1, lastCol);
    clearRangeLintNotesWithPrefixes(
      headerRowRange,
      [
        `${LINT_NOTE_PREFIX}No "Applies" header found.`,
        `${LINT_NOTE_PREFIX}No category includes "observation" in Applies.`,
        `${LINT_NOTE_PREFIX}No category includes "track" in Applies.`,
      ],
    );

    if (lastRow > 1) {
      const bodyRange = categoriesSheet.getRange(2, 1, lastRow - 1, lastCol);
      clearRangeLintNotesWithPrefixes(
        bodyRange,
        [
          `${LINT_NOTE_PREFIX}Unrecognized Applies token(s):`,
          `${LINT_NOTE_PREFIX}All Applies tokens are unrecognized`,
          `${LINT_NOTE_PREFIX}Applies value contains semicolons`,
        ],
      );
    }
  }
  const headerValues =
    lastCol > 0 ? categoriesSheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const normalizedHeaders = headerValues.map((header) =>
    String(header || "")
      .trim()
      .toLowerCase(),
  );
  const appliesColZeroBased = normalizedHeaders.findIndex(
    (header) =>
      header === "applies" ||
      header === "tracks" ||
      header === "applies to" ||
      header === "appliesto",
  );

  if (appliesColZeroBased === -1) {
    // The Applies header was removed or renamed. We've already cleared
    // Applies-specific notes above (within the existing lastCol range).
    // Also clear column D specifically — when lastCol < 4 the header-row
    // clear above doesn't reach it, so a stale warning would accumulate.
    const appliesExpectedCol = 4;
    const headerCell = categoriesSheet.getRange(1, appliesExpectedCol);
    clearRangeLintNotesWithPrefixes(headerCell, [
      `${LINT_NOTE_PREFIX}No "Applies" header found.`,
    ]);
    // Place the warning at the conventional Applies column position (column D)
    // rather than A1 (reserved for the Primary Language annotation) or
    // `lastCol + 1`, which lands in an off-screen empty column on wide sheets.
    // Column D is where Applies normally lives, so the warning stays visible
    // even when the header was renamed or removed.
    appendLintNote(
      headerCell,
      'No "Applies" header found. The builder resolves this column by header name and may auto-create it, seeding the first category with "track, observation".',
      "warning",
    );
    return;
  }

  const appliesColIndex = appliesColZeroBased + 1; // 1-based for Sheet ranges

  // Clear previous lint artifacts on column D (data rows + header)
  if (lastRow > 1) {
    const appliesRange = categoriesSheet.getRange(
      2,
      appliesColIndex,
      lastRow - 1,
      1,
    );
    clearLintArtifacts(appliesRange);
  }
  // Also clear header cell artifacts (Applies-specific notes only, to avoid
  // wiping higher-priority annotations on the same cell, e.g. A1 language error)
  const headerCell = categoriesSheet.getRange(1, appliesColIndex);
  clearRangeLintNotesWithPrefixes(
    headerCell,
    [
      `${LINT_NOTE_PREFIX}No "Applies" header found.`,
      `${LINT_NOTE_PREFIX}No category includes "observation" in Applies.`,
      `${LINT_NOTE_PREFIX}No category includes "track" in Applies.`,
    ],
  );

  if (lastRow <= 1) return;

  // Read both name and Applies so we can skip blank/spacer rows,
  // matching buildCategories() which returns early for rows without a name.
  // Resolve the Name column by header (mirrors buildCategories' header-map
  // lookup) instead of hard-coding column A.
  const appliesValues = categoriesSheet.getRange(
    2,
    appliesColIndex,
    lastRow - 1,
    1,
  ).getValues();
  const nameColZeroBased = normalizedHeaders.findIndex(
    (header) => header === "name",
  );
  const nameColIndex = nameColZeroBased >= 0 ? nameColZeroBased + 1 : 1;
  const nameValues = categoriesSheet.getRange(
    2,
    nameColIndex,
    lastRow - 1,
    1,
  ).getValues();
  let hasObservation = false;
  let hasTrack = false;

  for (let i = 0; i < appliesValues.length; i++) {
    const rawValue = appliesValues[i][0] == null ? "" : String(appliesValues[i][0]).trim();
    const categoryName = String(nameValues[i][0] || "").trim();
    const row = i + 2;

    // Mirror buildCategories(): skip rows without a category name entirely.
    // Blank Applies on a real category row still falls back to observation.
    if (!categoryName) continue;

    if (!rawValue) {
      // Blank Applies cells: mirror builder semantics.
      // The builder uses the physical array index (not a skip-blank counter)
      // for the AUTO_CREATED_APPLIES_COLUMN && index === 0 check.
      const isFirstCategory = i === 0;
      const isAutoCreated =
        typeof AUTO_CREATED_APPLIES_COLUMN !== "undefined" &&
        AUTO_CREATED_APPLIES_COLUMN;
      if (isAutoCreated && isFirstCategory) {
        hasObservation = true;
        hasTrack = true;
      } else {
        hasObservation = true;
      }
      continue;
    }

    // Parse tokens: split by comma, trim, lowercase
    // Mirror builder's parseTokens() which ONLY splits by comma — semicolons,
    // newlines, and other delimiters are NOT recognized in the Applies column
    // (unlike the Fields column which uses normalizeFieldTokens with broader parsing).
    // Keep the original-cased tokens for display in lint messages while matching
    // case-insensitively, so warnings echo exactly what the user typed.
    const rawTokens = rawValue
      .split(",")
      .map((t) => t.trim())
      .filter(Boolean);
    const tokens = rawTokens.map((t) => t.toLowerCase());

    // Warn if the raw value contains non-comma delimiters that the builder ignores.
    // This catches cases like "track; observation" where the semicolon causes the
    // builder to see "track; observation" as a single token → only "track" is kept
    // and "observation" is silently dropped.
    if (/[;；\n•·，、]/.test(rawValue)) {
      const cell = categoriesSheet.getRange(row, appliesColIndex);
      appendLintNote(
        cell,
        'Applies value contains semicolons or other non-comma delimiters. The builder only recognizes commas — use "observation, track" instead of "observation; track". Other delimiters cause tokens to be silently merged and potentially dropped.',
        "warning",
      );
    }

    const normalizedTokens = tokens
      .map((token) => {
        if (token.startsWith("o")) return "observation";
        if (token.startsWith("t")) return "track";
        return "";
      })
      .filter(Boolean);
    // Report invalid tokens using their original casing while matching on the
    // lowercased form, so the warning shows exactly what the user entered.
    const invalidTokens = rawTokens.filter((token) => {
      const lower = token.toLowerCase();
      return !lower.startsWith("o") && !lower.startsWith("t");
    });

    if (invalidTokens.length > 0) {
      const cell = categoriesSheet.getRange(row, appliesColIndex);
      if (normalizedTokens.length === 0 && tokens.length > 0) {
        // All tokens are unrecognized — mirror builder fallback logic
        const isFirstCategory = i === 0;
        const isAutoCreated =
          typeof AUTO_CREATED_APPLIES_COLUMN !== "undefined" &&
          AUTO_CREATED_APPLIES_COLUMN;
        const builderDefault =
          isAutoCreated && isFirstCategory
            ? "track + observation"
            : "observation";
        appendLintNote(
          cell,
          `All Applies tokens are unrecognized ("${invalidTokens.join('", "')}"). The builder will silently default this row to "${builderDefault}". Use "observation" or "track" explicitly.`,
          "warning",
        );
      } else {
        appendLintNote(
          cell,
          `Unrecognized Applies token(s): "${invalidTokens.join('", "')}". The builder only keeps observation/track prefixes and ignores the rest.`,
          "warning",
        );
      }
    }

    // If nothing matched, the builder falls back based on AUTO_CREATED_APPLIES_COLUMN.
    // Mirror builder: uses physical array index, not skip-blank counter.
    if (normalizedTokens.length === 0) {
      const isFirstCategory = i === 0;
      const isAutoCreated =
        typeof AUTO_CREATED_APPLIES_COLUMN !== "undefined" &&
        AUTO_CREATED_APPLIES_COLUMN;
      if (isAutoCreated && isFirstCategory) {
        hasObservation = true;
        hasTrack = true;
      } else {
        hasObservation = true;
      }
    } else {
      if (normalizedTokens.includes("observation")) hasObservation = true;
      if (normalizedTokens.includes("track")) hasTrack = true;
    }
  }

  // The payload builder still requires observation coverage somewhere in the sheet.
  if (!hasObservation) {
    const cell = categoriesSheet.getRange(1, appliesColIndex);
    appendLintNote(
      cell,
      'No category includes "observation" in Applies. Config generation currently fails unless at least one category resolves to observation.',
      "error",
    );
  }

  // When the Applies column was auto-created, the builder seeds the first category
  // with track + observation, so this is only a warning. Otherwise it is a hard
  // error — the builder throws 'At least one category must include "track"'.
  if (!hasTrack) {
    const cell = categoriesSheet.getRange(1, appliesColIndex);
    if (typeof AUTO_CREATED_APPLIES_COLUMN !== "undefined" && AUTO_CREATED_APPLIES_COLUMN) {
      appendLintNote(
        cell,
        'No category includes "track" in Applies. The Applies column was auto-created and the first category will be defaulted to track + observation during generation. Review column D to confirm.',
        "warning",
      );
    } else {
      appendLintNote(
        cell,
        'No category includes "track" in Applies. Config generation will fail — at least one category must include "track".',
        "error",
      );
    }
  }
}

function validateCategoryIcons(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categoriesSheet = spreadsheet.getSheetByName("Categories");

  if (!categoriesSheet) {
    console.log("Categories sheet not found during icon validation");
    return;
  }

  const lastRow = categoriesSheet.getLastRow();
  if (lastRow <= 1) {
    console.log("No category rows available for icon validation");
    return;
  }

  const iconRange = categoriesSheet.getRange(2, 2, lastRow - 1, 1);
  // Clear icon column lint artifacts: backgrounds, font colors, and notes
  clearLintArtifacts(iconRange);

  const iconValues = iconRange.getValues();
  // Also read Icon ID column (column F, 1-based index 6) if present
  const lastCol = categoriesSheet.getLastColumn();
  const hasIconIdColumn = lastCol >= 6;
  const iconIdValues = hasIconIdColumn
    ? categoriesSheet.getRange(2, 6, lastRow - 1, 1).getValues()
    : [];
  // Build a set of known icon IDs from the Icons sheet for validation.
  // Mirror buildIconsFromSheet: only register IDs whose icon source (column B)
  // is present and recognised by parseIconSource.
  const knownIconIds = new Set<string>();
  const iconsSheet = spreadsheet.getSheetByName("Icons");
  if (iconsSheet) {
    const iconsLastRow = iconsSheet.getLastRow();
    if (iconsLastRow > 1) {
      const iconsData = iconsSheet.getRange(2, 1, iconsLastRow - 1, 2).getValues();
      for (const row of iconsData) {
        const id = String(row[0] || "").trim();
        const iconStr = String(row[1] || "").trim();
        if (id && iconStr && hasRecognisedIconSource(iconStr)) {
          knownIconIds.add(id);
        }
      }
    }
  }

  // Also add category-derived icon IDs (mirrors buildIconsFromSheet logic).
  // The builder creates icons from categories with: iconId = iconIdFromSheet || categoryId,
  // where categoryId = idFromSheet || slugify(name) || `category-${index+1}`.
  // These IDs are valid even without an Icons sheet entry.
  const nameValues = categoriesSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const hasCategoryIdColumn = lastCol >= 5;
  const categoryIdValues = hasCategoryIdColumn
    ? categoriesSheet.getRange(2, 5, lastRow - 1, 1).getValues()
    : [];
  for (let i = 0; i < nameValues.length; i++) {
    const name = String(nameValues[i][0] || "").trim();
    const iconRaw = iconValues[i][0];
    const iconStr = typeof iconRaw === "string" ? String(iconRaw).trim() : "";
    // Mirror buildIconsFromSheet(): only rows with name and supported icon source contribute
    if (!name || !iconStr || !hasRecognisedIconSource(iconStr)) continue;

    const idFromSheet = hasCategoryIdColumn
      ? String(categoryIdValues[i][0] || "").trim()
      : "";
    const categoryId = idFromSheet || slugify(name) || `category-${i + 1}`;

    const iconIdFromSheet = hasIconIdColumn
      ? String(iconIdValues[i]?.[0] || "").trim()
      : "";
    const iconId = iconIdFromSheet || categoryId;

    if (iconId) knownIconIds.add(iconId);
  }
  const rowIssues = new Map<number, string[]>();

  const addIssue = (row: number, message: string): void => {
    if (!rowIssues.has(row)) {
      rowIssues.set(row, []);
    }
    const issues = rowIssues.get(row)!;
    if (!issues.includes(message)) {
      issues.push(message);
    }
  };

  iconValues.forEach((row, index) => {
    const rowNumber = index + 2;
    const iconCellValue = row[0];

    if (
      iconCellValue === null ||
      iconCellValue === undefined ||
      iconCellValue === ""
    ) {
      // Check if this row has an Icon ID that the builder can resolve
      const iconId = iconIdValues[index]?.[0];
      const iconIdStr = iconId !== null && iconId !== undefined ? String(iconId).trim() : "";
      if (iconIdStr) {
        if (knownIconIds.has(iconIdStr)) {
          // Builder resolves icon from Icons sheet via Icon ID — skip warning
          return;
        }
        // Icon ID doesn't exist in Icons sheet — error (builder throws for unresolved iconId)
        const cell = categoriesSheet.getRange(rowNumber, 2);
        setLintNote(
          cell,
          `Icon ID "${iconIdStr}" was not found in the Icons sheet. Generate Config will fail because the builder cannot resolve this icon reference. Either add a matching entry to the Icons sheet or clear the Icon ID.`,
          "error",
        );
        return;
      }
      // Missing icon — warn (builder creates category without icon, not a hard error)
      const cell2 = categoriesSheet.getRange(rowNumber, 2);
      setLintNote(
        cell2,
        "Icon is empty — the category will be exported without an icon. Add an icon here or provide an Icon ID that resolves from the Icons sheet.",
        "warning",
      );
      return;
    }

    if (typeof iconCellValue === "string") {
      const iconValue = iconCellValue.trim();
      if (!iconValue) {
        // Check if this row has an Icon ID that the builder can resolve
        const iconId = iconIdValues[index]?.[0];
        const iconIdStr = iconId !== null && iconId !== undefined ? String(iconId).trim() : "";
        if (iconIdStr) {
          if (knownIconIds.has(iconIdStr)) {
            // Builder resolves icon from Icons sheet via Icon ID — skip warning
            return;
          }
          // Icon ID doesn't exist in Icons sheet — error (builder throws for unresolved iconId)
          const cell = categoriesSheet.getRange(rowNumber, 2);
          setLintNote(
            cell,
            `Icon ID "${iconIdStr}" was not found in the Icons sheet. Generate Config will fail because the builder cannot resolve this icon reference. Either add a matching entry to the Icons sheet or clear the Icon ID.`,
            "error",
          );
          return;
        }
        // Whitespace-only — treat as missing
        const cell = categoriesSheet.getRange(rowNumber, 2);
        setLintNote(
          cell,
          "Icon is empty — the category will be exported without an icon. Add an icon here or provide an Icon ID that resolves from the Icons sheet.",
          "warning",
        );
        return;
      }

      if (iconValue.startsWith("<svg")) {
        // Inline SVG markup - passed through
        return;
      } else if (isSupportedSvgDataUri(iconValue)) {
        // Valid SVG data URI - passed through
        return;
      } else if (iconValue.startsWith("data:")) {
        addIssue(
          rowNumber,
          "Only SVG data URIs (data:image/svg+xml) are supported for direct config export; other data URIs are silently dropped during generation.",
        );
      } else if (iconValue.startsWith("https://drive.google.com/file/d/")) {
        // Drive URL in builder-supported form - validate access
        const fileId = extractDriveFileId(iconValue);
        if (fileId) {
          let info = driveIconInfoCache.get(fileId);
          if (!info) {
            info = getDriveIconInfo(fileId);
          }

          if (info.isSvg) {
            // Access is valid and file is SVG; no lint issue.
            return;
          }

          if (info.errorMessage) {
            addIssue(rowNumber, info.errorMessage);
          } else {
            // File is accessible but not SVG (slug may be empty for
            // punctuation-only or non-Latin filenames, but config
            // generation silently drops all non-SVG Drive files).
            addIssue(
              rowNumber,
              "Google Drive icon files must be SVG for direct config export; non-SVG Drive files are silently dropped during generation.",
            );
          }
        } else {
          addIssue(
            rowNumber,
            "Icon URL must contain a valid Google Drive file ID.",
          );
        }
      } else if (iconValue.startsWith("https://drive.google.com/")) {
        addIssue(
          rowNumber,
          "Google Drive icon URLs must use the /file/d/ form (for example, https://drive.google.com/file/d/<FILE_ID>/view) so config generation can package them.",
        );
      } else if (/^http:\/\//i.test(iconValue)) {
        if (iconValue.toLowerCase().includes(".svg")) {
          // HTTP SVG URL - warn about security (should use HTTPS)
          const cell = categoriesSheet.getRange(rowNumber, 2);
          setLintNote(
            cell,
            "Icon URL should use HTTPS instead of HTTP for security",
            "warning",
          );
          return;
        }
        addIssue(
          rowNumber,
          'HTTP(S) icon URLs must point directly to an SVG file (URL must contain ".svg"); non-SVG URLs are silently dropped during config generation.',
        );
      } else if (/^https:\/\//i.test(iconValue)) {
        if (iconValue.toLowerCase().includes(".svg")) {
          // HTTPS SVG URL - passed through
          return;
        }
        addIssue(
          rowNumber,
          'HTTP(S) icon URLs must point directly to an SVG file (URL must contain ".svg"); non-SVG URLs are silently dropped during config generation.',
        );
      } else {
        // Plain text - works with 'Generate Category Icons' search but not direct export
        setLintNote(
          categoriesSheet.getRange(rowNumber, 2),
          `Plain text icon '${iconValue}' works with 'Generate Category Icons' search, but direct config export only packages supported icon sources (inline SVG, SVG data URI, Drive SVG, or direct SVG URL).`,
          "warning",
        );
        return;
      }
    } else if (
      iconCellValue &&
      typeof iconCellValue === "object" &&
      iconCellValue.toString() === "CellImage"
    ) {
      // Cell images ARE supported and will be processed via icon API
      return;
    } else {
      addIssue(
        rowNumber,
        "Unrecognized icon cell value. Expected text, Drive URL, or cell image.",
      );
    }
  });

  rowIssues.forEach((messages, rowNumber) => {
    const cell = categoriesSheet.getRange(rowNumber, 2);
    setLintNote(cell, messages.join("\n"), "error");
    console.warn(
      `Icon issue in Categories row ${rowNumber}: ${messages.join(" | ")}`,
    );
  });

  if (rowIssues.size === 0) {
    console.log("Category icon validation completed with no issues found.");
  }
}

function validatePrimaryLanguageInA1(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categoriesSheet = spreadsheet.getSheetByName("Categories");
  if (!categoriesSheet) return;

  const cell = categoriesSheet.getRange(1, 1);
  clearLintArtifacts(cell);

  // Mirror getPrimaryLanguageName(): if Metadata sheet has a primaryLanguage
  // entry, the builder uses that and ignores A1 entirely. Only validate A1
  // when it is the effective primary-language source.
  const metadataSheet = spreadsheet.getSheetByName("Metadata");
  if (metadataSheet) {
    const metadataValues = metadataSheet.getDataRange().getValues();
    for (let i = 1; i < metadataValues.length; i++) {
      if (String(metadataValues[i][0]).trim() === "primaryLanguage") {
        const lang = String(metadataValues[i][1] || "").trim();
        if (lang) return; // Metadata has primaryLanguage — A1 is not used
      }
    }
  }

  const a1Value = String(categoriesSheet.getRange(1, 1).getValue() || "").trim();
  if (!a1Value) {
    setLintNote(
      cell,
      'Categories A1 is blank and no Metadata primaryLanguage is set. Config generation will fail with an error. Set a valid language name or locale code (e.g. "English", "Português", "en", "pt-BR") in A1 or add a non-empty "primaryLanguage" row in the Metadata sheet.',
      "error",
    );
    return;
  }

  // "Name" is the correct column header in the new format — the primary
  // language is stored in the Metadata sheet instead of A1.  Give a
  // targeted warning pointing the user to Metadata rather than suggesting
  // A1 is wrong.
  const lowerA1 = a1Value.toLowerCase();
  if (lowerA1 === "name") {
    setLintNote(
      cell,
      'Categories A1 is correctly set to "Name". ' +
        'Set the primary language in the Metadata sheet instead: ' +
        'add a row with key "primaryLanguage" and value (e.g. "English", "Português", "en", "pt-BR").',
      "warning",
    );
    return;
  }

  // Other standard header words that are clearly not language names.
  // These WILL cause getPrimaryLanguage() to throw at runtime if Metadata
  // has no primaryLanguage, so we flag them as errors.
  const HEADER_WORDS = ["category", "categories", "label", "type"];
  if (HEADER_WORDS.includes(lowerA1)) {
    setLintNote(
      cell,
      `Categories A1 contains "${a1Value}" which is not a valid primary language value. ` +
        `Set a valid language or locale code (e.g. "English", "Português", "en", "pt-BR") or add a "primaryLanguage" row in the Metadata sheet.`,
      "error",
    );
    return;
  }

  // Validate as a language name or locale code
  const validation = validatePrimaryLanguage(a1Value);
  if (!validation.valid) {
    setLintNote(
      cell,
      `Invalid primary language in A1: ${validation.error || "Unknown error"}. The builder will fall back to Metadata primaryLanguage or default to "en".`,
      "warning",
    );
  }
}

// Specific sheet linting functions
function lintCategoriesSheet(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categoriesSheetRef = spreadsheet.getSheetByName("Categories");
  const detailsSheetRef = spreadsheet.getSheetByName("Details");

  // Task 5: Check if Categories sheet is empty before running lint.
  // Must run BEFORE validatePrimaryLanguageInA1() to avoid the empty-sheet
  // check wiping the A1 language error note (both target cell A1).
  if (categoriesSheetRef && checkEmptySheet(categoriesSheetRef, "Categories")) {
    return;
  }

  // validatePrimaryLanguageInA1() clears A1 artifacts internally, so no
  // separate A1 clear is needed here. Removing the redundant clear avoids
  // accidentally wiping any intermediate state if future phases write to A1.

  // Phase 4: Validate primary language in A1 (only when sheet has data rows)
  validatePrimaryLanguageInA1();

  // Pre-cleanup: capture whitespace-only IDs before cleanWhitespaceOnlyCells() erases them
  const rawCategoryIds = inspectRawIds("Categories", 5);

  if (categoriesSheetRef) {
    const lastRow = categoriesSheetRef.getLastRow();
    if (lastRow > 1) {
      // Clear font colors in icon column (column 2) for lint warnings
      clearRangeFontColorIfMatches(
        categoriesSheetRef.getRange(2, 2, lastRow - 1, 1),
        LINT_WARNING_FONT_COLORS,
      );
      // Clear background colors in fields column (column 3)
      clearRangeBackgroundIfMatches(
        categoriesSheetRef.getRange(2, 3, lastRow - 1, 1),
        [LINT_ERROR_BG],
      );
    }
  }

  // Build a Set of all valid field identifiers: slugified names AND explicit IDs
  let cachedValidFieldIds: Set<string>;
  if (detailsSheetRef) {
    const lastRow = detailsSheetRef.getLastRow();
    if (lastRow > 1) {
      const detailData = detailsSheetRef
        .getRange(2, 1, lastRow - 1, 5)
        .getValues();
      cachedValidFieldIds = new Set<string>();
      for (const row of detailData) {
        const name = String(row[0] || "").trim();
        const explicitId = String(row[4] || "").trim();
        if (name) {
          cachedValidFieldIds.add(slugify(name));
        }
        if (explicitId) {
          cachedValidFieldIds.add(explicitId);
          cachedValidFieldIds.add(explicitId.toLowerCase());
        }
      }
    }
  }
  cachedValidFieldIds = cachedValidFieldIds || new Set<string>();

  const categoriesValidations = [
    // Rule 1: Capitalize the first letter of the category name
    (value: string, row: number, col: number) => {
      if (isEmptyOrWhitespace(value)) return;

      const capitalizedValue = capitalizeFirstLetter(value);
      if (capitalizedValue !== value) {
        try {
          categoriesSheetRef?.getRange(row, col).setValue(capitalizedValue);
        } catch (error) {
          console.error(
            "Error capitalizing value in Categories sheet at row " +
              row +
              ", col " +
              col +
              ":",
            error,
          );
        }
      }
    },
    // Rule 2: Icon validation is handled entirely by validateCategoryIcons()
    // which runs after lintSheet() to avoid clear/write race conditions.
    () => {
      // Icon column — no inline validation; see validateCategoryIcons()
    },
    // Rule 3: Validate field references using normalizeFieldTokens (matches build parsing)
    (value: string, row: number, col: number) => {
      if (isEmptyOrWhitespace(value)) return;

      try {
        // Use normalizeFieldTokens to match build parsing across commas, semicolons, newlines, bullets, fullwidth
        if (cachedValidFieldIds.size > 0) {
          const tokens = normalizeFieldTokens(value);
          const invalidFields: string[] = [];
          for (const token of tokens) {
            const slugified = slugify(token);
            const tokenLower = token.toLowerCase();
            if (
              slugified &&
              !cachedValidFieldIds.has(slugified) &&
              !cachedValidFieldIds.has(token) &&
              !cachedValidFieldIds.has(tokenLower)
            ) {
              invalidFields.push(token);
            }
          }

          if (invalidFields.length > 0) {
            console.log(
              "Invalid fields in row " + row + ": " + invalidFields.join(", "),
            );
            const cell = categoriesSheetRef?.getRange(row, col);
            if (cell) {
              setLintNote(
                cell,
                `Invalid fields: ${invalidFields.join(", ")}. These fields do not exist in the Details sheet`,
                "error",
              );
            }
          }
        }
      } catch (error) {
        console.error(
          "Error validating fields in Categories sheet at row " +
            row +
            ", col " +
            col +
            ":",
          error,
        );
      }
    },
  ];

  // Category name and icon are required
  // Preserve backgrounds in category name column (index 0) because they are user-set category colors
  lintSheet("Categories", categoriesValidations, [0, 1], [0]);
  validateCategoryIcons();

  // Phase 2: Post-lintSheet() edge-case checks
  checkDuplicateCategoryIds(); // Task 1: duplicate effective IDs in col E
  checkSlugCollisions("Categories", 1); // Task 3: slug collisions in col A
  validateAppliesColumn(); // Task 4: Applies column validation
  checkManualIdHygiene("Categories", 5, rawCategoryIds); // Task 6: ID hygiene in col E
}
