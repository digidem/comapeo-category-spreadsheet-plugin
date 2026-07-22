/// <reference path="shared.ts" />

/**
 * Phase 5 Task 1: Validates the Icons sheet.
 * Checks for missing icon IDs, duplicate icon IDs, missing icon sources,
 * and unsupported icon source formats.
 */
function lintIconsSheet(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const iconsSheet = spreadsheet.getSheetByName("Icons");
  if (!iconsSheet) return; // Icons sheet is optional

  const lastRow = iconsSheet.getLastRow();
  if (lastRow <= 1) return; // Header-only or empty

  const logger = getScopedLogger("LintIconsSheet");

  // Clear previous lint artifacts on columns A and B
  const dataRange = iconsSheet.getRange(2, 1, lastRow - 1, 2);
  clearLintArtifacts(dataRange);

  const data = dataRange.getValues();
  const seenIds = new Map<string, number[]>();

  for (let i = 0; i < data.length; i++) {
    const row = i + 2;
    const iconId = String(data[i][0] || "").trim();
    const iconSource = String(data[i][1] || "").trim();

    // Skip fully blank rows — builder ignores them too (buildIconsFromSheet
    // skips rows when either iconId or iconStr is empty).
    if (!iconId && !iconSource) continue;

    // Check for missing icon ID (col A)
    if (!iconId) {
      setLintNote(
        iconsSheet.getRange(row, 1),
        "Icon ID is required.",
        "error",
      );
      logger.warn(`Row ${row}: missing icon ID`);
      continue;
    }

    // Check for missing icon source (col B)
    if (!iconSource) {
      setLintNote(
        iconsSheet.getRange(row, 2),
        "Icon source is required.",
        "error",
      );
      logger.warn(`Row ${row}: missing icon source for ID "${iconId}"`);
      continue;
    }

    // Check for unsupported icon source format.
    // Mirrors parseIconSource() in payloadBuilder: only rows that resolve to a
    // packaged icon asset should contribute to duplicate-ID tracking.
    const isSvg = iconSource.startsWith("<svg");
    const isDataUri = iconSource.toLowerCase().startsWith("data:image/svg+xml");
    const isDriveUrl = iconSource.startsWith("https://drive.google.com/file/d/");
    const isHttpUrl = /^https?:\/\//i.test(iconSource);
    // Builder's isSvgUrl() requires the URL to contain ".svg"
    const isSvgHttpUrl = isHttpUrl && iconSource.toLowerCase().includes(".svg");
    let contributesIconAsset = false;

    if (isSvg) {
      contributesIconAsset = true;
      // Inline SVG — check for basic structural validity
      if (!iconSource.includes("</svg>") && !iconSource.trim().endsWith("/>")) {
        setLintNote(
          iconsSheet.getRange(row, 2),
          'Inline SVG markup appears incomplete (missing closing tag). The builder will include this icon as-is, but malformed SVG may cause rendering issues in CoMapeo.',
          "warning",
        );
      }
    } else if (isDataUri) {
      // Data URI — verify it actually decodes to valid SVG content
      if (!isSupportedSvgDataUri(iconSource)) {
        setLintNote(
          iconsSheet.getRange(row, 2),
          "SVG data URI is malformed or does not decode to valid SVG content. This icon will be silently dropped during config generation.",
          "error",
        );
      } else {
        contributesIconAsset = true;
      }
    } else if (isDriveUrl) {
      // Drive URL — verify the file is accessible and is SVG
      const fileId = extractDriveFileId(iconSource);
      if (fileId) {
        let info = driveIconInfoCache.get(fileId);
        if (!info) {
          info = getDriveIconInfo(fileId);
        }
        if (info.errorMessage) {
          setLintNote(
            iconsSheet.getRange(row, 2),
            info.errorMessage,
            "error",
          );
        } else if (!info.isSvg) {
          setLintNote(
            iconsSheet.getRange(row, 2),
            "Google Drive icon file is not SVG. Non-SVG Drive files are silently dropped during config generation.",
            "error",
          );
        } else {
          contributesIconAsset = true;
        }
      } else {
        setLintNote(
          iconsSheet.getRange(row, 2),
          "Google Drive URL does not contain a valid file ID.",
          "error",
        );
      }
    } else if (!isSvgHttpUrl) {
      if (isHttpUrl) {
        // HTTP URL but not an SVG — will be silently dropped by parseIconSource()
        setLintNote(
          iconsSheet.getRange(row, 2),
          "HTTP(S) icon URLs must point directly to an SVG file (URL must contain \".svg\"). Non-SVG URLs are silently dropped during config generation.",
          "error",
        );
      } else {
        setLintNote(
          iconsSheet.getRange(row, 2),
          "Unsupported icon source format. Expected inline SVG (<svg…>), data:image/svg+xml URI, Google Drive URL (must use /file/d/ format), or HTTP(S) URL ending in .svg.",
          "error",
        );
      }
      logger.warn(
        `Row ${row}: unsupported icon source format for ID "${iconId}"`,
      );
    } else {
      contributesIconAsset = true;
    }

    if (contributesIconAsset) {
      if (!seenIds.has(iconId)) {
        seenIds.set(iconId, [row]);
      } else {
        seenIds.get(iconId)?.push(row);
      }
    }
  }

  // Flag duplicate IDs
  seenIds.forEach((rows, id) => {
    if (rows.length > 1) {
      logger.warn(
        `Duplicate icon ID "${id}" in rows: ${rows.join(", ")}`,
      );
      for (const row of rows) {
        const otherRows = rows.filter((r) => r !== row);
        appendLintNote(
          iconsSheet.getRange(row, 1),
          `Duplicate icon ID "${id}" (also in row${otherRows.length > 1 ? "s" : ""} ${otherRows.join(", ")}).`,
          "error",
        );
      }
    }
  });
}

/**
 * Returns exact icon ID collisions between Icons sheet rows and
 * category-derived icon rows. This must mirror buildIconsFromSheet(), which
 * deduplicates by exact icon.id values rather than sanitized variants.
 */
function findCrossSheetIconIdCollisions(
  iconsEntries: Array<{ id: string; row: number }>,
  categoryEntries: Array<{ id: string; row: number }>,
): Array<{
  iconId: string;
  iconRow: number;
  categoryId: string;
  categoryRow: number;
}> {
  const categoryEntriesById = new Map<string, Array<{ id: string; row: number }>>();
  for (const entry of categoryEntries) {
    if (!entry.id) continue;
    if (!categoryEntriesById.has(entry.id)) {
      categoryEntriesById.set(entry.id, []);
    }
    categoryEntriesById.get(entry.id)?.push(entry);
  }

  const collisions: Array<{
    iconId: string;
    iconRow: number;
    categoryId: string;
    categoryRow: number;
  }> = [];

  for (const iconEntry of iconsEntries) {
    const matchingCategoryEntries = categoryEntriesById.get(iconEntry.id);
    if (!matchingCategoryEntries) continue;
    for (const categoryEntry of matchingCategoryEntries) {
      collisions.push({
        iconId: iconEntry.id,
        iconRow: iconEntry.row,
        categoryId: categoryEntry.id,
        categoryRow: categoryEntry.row,
      });
    }
  }

  return collisions;
}

/**
 * Phase 5 Task 2: Checks for exact icon ID collisions between the Icons sheet
 * and Categories sheet. The builder merges icons from both sources,
 * deduplicating by exact ID, so lint must do the same.
 */
function checkCrossSheetIconCollisions(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const iconsSheet = spreadsheet.getSheetByName("Icons");
  const categoriesSheet = spreadsheet.getSheetByName("Categories");
  if (!iconsSheet || !categoriesSheet) return;

  const iconsLastRow = iconsSheet.getLastRow();
  const categoriesLastRow = categoriesSheet.getLastRow();
  if (iconsLastRow <= 1 || categoriesLastRow <= 1) return;

  const logger = getScopedLogger("LintCrossSheetIconCollisions");

  // Read Icons sheet IDs and sources (mirror buildIconsFromSheet(): both ID and
  // source must be present, and unsupported sources are ignored)
  const iconsData = iconsSheet
    .getRange(2, 1, iconsLastRow - 1, 2)
    .getValues();
  const iconsEntries: Array<{ id: string; row: number }> = [];
  for (let i = 0; i < iconsData.length; i++) {
    const rawIconId = String(iconsData[i][0] || "").trim();
    const iconStr = String(iconsData[i][1] || "").trim();
    if (!rawIconId || !iconStr || !hasRecognisedIconSource(iconStr)) {
      continue;
    }
    const sanitizedIconId = sanitizeIconSlug(rawIconId);
    if (sanitizedIconId !== rawIconId) {
      appendLintNote(
        iconsSheet.getRange(i + 2, 1),
        `Icon ID "${rawIconId}" contains a file extension. The ID is used as-is in the config, but file name generation strips the extension — so using "${rawIconId}" as an ID will produce a file named "${sanitizedIconId}.svg" (or .png), meaning the config references "${rawIconId}" but the actual file is "${sanitizedIconId}.svg". Consider using "${sanitizedIconId}" directly to keep the ID and file name consistent.`,
        "warning",
      );
    }
    iconsEntries.push({ id: rawIconId, row: i + 2 });
  }

  // Read Categories: resolve columns by header name (mirrors buildIconsFromSheet).
  const catLastCol = Math.max(categoriesSheet.getLastColumn(), 6);
  const catHeaders = categoriesSheet.getRange(1, 1, 1, catLastCol).getValues()[0];
  const catHeaderMap = createLintHeaderMap(catHeaders);
  const catNameCol = getLintColumnIndex(catHeaderMap, "name") ?? 0;
  const catIconCol = getLintColumnIndex(catHeaderMap, "icon", "icons") ?? 1;
  const catCategoryIdCol = getLintColumnIndex(catHeaderMap, "category id", "id");
  const catIconIdCol = getLintColumnIndex(catHeaderMap, "icon id", "iconid");

  const catData = categoriesSheet.getRange(2, 1, categoriesLastRow - 1, catLastCol).getValues();
  const categoryEntries: Array<{ id: string; row: number }> = [];
  for (let i = 0; i < catData.length; i++) {
    const name = String(catData[i][catNameCol] || "").trim();
    const iconRaw = catData[i][catIconCol];
    const iconStr = typeof iconRaw === "string" ? iconRaw.trim() : "";
    // Mirror buildIconsFromSheet(): only category rows with a name and a
    // supported icon source contribute packaged icon assets.
    if (!name || !iconStr || !hasRecognisedIconSource(iconStr)) continue;
    // Mirror buildIconsFromSheet(): slugify(name) || `category-${index+1}`
    // The builder falls back to `category-N` when slugify produces an empty
    // string (e.g. non-Latin or punctuation-only names).
    const explicitCategoryId = catCategoryIdCol !== undefined
      ? String(catData[i][catCategoryIdCol] || "").trim()
      : "";
    const categoryId = explicitCategoryId || slugify(name) || `category-${i + 1}`;
    // Builder uses iconIdCol ("icon id") when present, otherwise falls back to categoryId
    const explicitIconId = catIconIdCol !== undefined
      ? String(catData[i][catIconIdCol] || "").trim()
      : "";
    const iconId = explicitIconId || categoryId;
    if (iconId) {
      categoryEntries.push({ id: iconId, row: i + 2 });
    }
  }

  const collisions = findCrossSheetIconIdCollisions(iconsEntries, categoryEntries);
  for (const collision of collisions) {
    logger.warn(
      `Icon ID collision: "${collision.iconId}" (Icons row ${collision.iconRow}) vs "${collision.categoryId}" (Categories row ${collision.categoryRow})`,
    );
    appendLintNote(
      iconsSheet.getRange(collision.iconRow, 1),
      `Icon ID "${collision.iconId}" collides with a category-derived ID in Categories row ${collision.categoryRow}. Entries will be merged — if this Icons sheet row provides SVG data, it may override the category icon.`,
      "warning",
    );
    appendLintNote(
      categoriesSheet.getRange(collision.categoryRow, 2),
      `Category icon ID "${collision.categoryId}" collides with an Icons sheet entry in row ${collision.iconRow}. Entries will be merged — if the Icons sheet entry provides SVG data, it may take priority over the category icon.`,
      "warning",
    );
  }
}
