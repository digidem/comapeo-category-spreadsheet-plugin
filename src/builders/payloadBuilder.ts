/**
 * Shared payload builder helpers used by generateCoMapeoConfig + apiService
 * Keeping these in a single file avoids duplicate Apps Script globals.
 */


/**
 * Creates a BuildRequest payload from spreadsheet data
 *
 * @param data - Spreadsheet data from getSpreadsheetData()
 * @returns BuildRequest payload ready for API
 */
function createBuildPayload(data: SheetData): BuildRequest {
  const fields = buildFields(data);
  // buildCategories now returns categories with both defaultFieldIds and fields populated.
  // Avoid remapping here to prevent accidentally clearing the arrays.
  const categories = buildCategories(data, fields);
  const icons = buildIconsFromSheet(data);

  // Early check: surface a clear error when categories reference icons that
  // failed to resolve, instead of letting the downstream comapeocat validator
  // throw a cryptic "CategoryRefError: Missing icon ref".
  const registeredIconIds = new Set(icons.map((i) => i.id));
  const missingIconCategories = categories.filter(
    (cat) => cat.iconId && !registeredIconIds.has(cat.iconId),
  );
  if (missingIconCategories.length > 0) {
    const names = missingIconCategories
      .map((cat) => `"${cat.name || cat.id}"`)
      .join(", ");
    throw new Error(
      `Could not load icons for categories: ${names}. ` +
        "Check that the Icon column contains valid SVG content, a public SVG URL, " +
        "or a Google Drive link to an SVG file that you have access to.",
    );
  }

  const metadata = buildMetadata(data);
  const locales = buildLocales();
  const translations = buildTranslationsPayload(data, categories, fields);

  getScopedLogger("PayloadBuilder").info("Build payload summary", {
    categories: categories.length,
    fields: fields.length,
    icons: icons.length,
    locales: locales.length,
    translations: Object.keys(translations || {}).length,
  });

  // Set category selection in exact spreadsheet order
  const categoryIds = categories.map((c) => c.id);
  setCategorySelection(categoryIds);

  getScopedLogger("PayloadBuilder").info(
    `Built payload with ${categories.length} categories, ${fields.length} fields, ${icons.length} icons`,
  );

  const payload: BuildRequest = {
    metadata,
    locales,
    categories,
    fields,
    icons: icons.length > 0 ? icons : undefined,
    translations:
      Object.keys(translations).length > 0
        ? (translations as unknown as TranslationsByLocale)
        : undefined,
  };

  const debugPayloadPreview = {
    ...payload,
    icons: payload.icons
      ? `<<${payload.icons.length} icons omitted from preview>>`
      : undefined,
  };
  getScopedLogger("PayloadBuilder").info(
    "[DEBUG] BuildRequest payload preview:",
    JSON.stringify(debugPayloadPreview, null, 2),
  );

  return payload;
}

/**
 * Builds locales array required by API v2
 * Uses Metadata!primaryLanguage if present, otherwise defaults to 'en'
 */
function buildLocales(): string[] {
  const normalizeLocaleInput = (
    raw: string | null | undefined,
  ): string | null => {
    if (!raw) return null;
    const trimmed = raw.trim();
    if (!trimmed) return null;

    // Already an ISO code
    if (/^[a-z]{2,3}(-[a-z]{2,3})?$/i.test(trimmed)) {
      return trimmed.toLowerCase();
    }

    try {
      if (typeof validateLanguageName === "function") {
        const validation = validateLanguageName(trimmed);
        if (validation?.valid && validation.code) {
          return validation.code;
        }
      }
    } catch (error) {
      getScopedLogger("PayloadBuilder").warn(
        "Failed to map primary language name to locale code:",
        error,
      );
    }

    return null;
  };

  // Prefer the Categories!A1 value resolved via getPrimaryLanguage
  try {
    if (typeof getPrimaryLanguage === "function") {
      const primary = getPrimaryLanguage();
      if (primary?.code) {
        return [primary.code];
      }
      const normalized = normalizeLocaleInput(primary?.name);
      if (normalized) {
        return [normalized];
      }
    }
  } catch (error) {
    getScopedLogger("PayloadBuilder").warn(
      "Falling back to Metadata primaryLanguage for locales:",
      error,
    );
  }

  // Fallback to Metadata sheet entry (legacy behavior)
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const metadataSheet = spreadsheet.getSheetByName("Metadata");

  if (!metadataSheet) {
    return ["en"];
  }

  const data = metadataSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === "primaryLanguage") {
      const normalized = normalizeLocaleInput(String(data[i][1] || ""));
      if (normalized) {
        return [normalized];
      }
    }
  }

  return ["en"];
}

/**
 * Builds metadata from spreadsheet
 * Note: data.documentName is a string (spreadsheet name), not an array
 */
function buildMetadata(data: SheetData): Metadata {
  // FIXED: documentName is a string from getSpreadsheetData(), not a nested array
  const documentName =
    typeof data.documentName === "string"
      ? data.documentName
      : String(data.documentName || "Unnamed Config");

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let metadataSheet = spreadsheet.getSheetByName("Metadata");

  if (!metadataSheet) {
    metadataSheet = spreadsheet.insertSheet("Metadata");
    metadataSheet
      .getRange(1, 1, 1, 2)
      .setValues([["Key", "Value"]])
      .setFontWeight("bold");
  }

  const sheetData = metadataSheet.getDataRange().getValues();

  const getValue = (key: string, defaultVal: string): string => {
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][0] === key) {
        if (key === "version") {
          // Always update version with current date
          const newVersion = Utilities.formatDate(
            new Date(),
            "UTC",
            "yy.MM.dd",
          );
          metadataSheet.getRange(i + 1, 2).setValue(newVersion);
          return newVersion;
        }
        return String(sheetData[i][1]);
      }
    }
    // Key not found - append new row
    metadataSheet.appendRow([key, defaultVal]);
    return defaultVal;
  };

  const getBooleanValue = (key: string, defaultVal: boolean): boolean => {
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][0] === key) {
        const val = String(sheetData[i][1]).trim().toUpperCase();
        return val === "TRUE";
      }
    }
    // Key not found - append new row with default value
    metadataSheet.appendRow([key, defaultVal ? "TRUE" : "FALSE"]);
    return defaultVal;
  };

  const name = getValue("name", `config-${slugify(documentName)}`);
  const version = getValue(
    "version",
    Utilities.formatDate(new Date(), "UTC", "yy.MM.dd"),
  );
  const description = getValue("description", "");
  const legacyCompat = getBooleanValue("legacyCompat", false);

  return {
    name,
    version,
    description: description || undefined,
    builderName: PLUGIN_INFO.NAME,
    builderVersion: PLUGIN_INFO.VERSION,
    legacyCompat: legacyCompat || undefined,
  };
}

/**
 * Builds fields array from Details sheet
 */
function buildFields(data: SheetData): Field[] {
  const details = data.Details?.slice(1) || [];

  // Keep a handle to the sheet so we can write back auto-generated IDs for missing values
  let detailsSheet: GoogleAppsScript.Spreadsheet.Sheet | null = null;
  try {
    detailsSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Details");
  } catch (error) {
    getScopedLogger("PayloadBuilder").warn("Unable to fetch Details sheet for ID backfill:", error);
  }

  // Use headers from data directly instead of fetching from sheet again
  const headerRow = data.Details?.[0] || [];
  const headerMap: Record<string, number> = {};
  headerRow.forEach((h, idx) => {
    const key = String(h || "")
      .trim()
      .toLowerCase();
    if (key) headerMap[key] = idx;
  });

  const getCol = (...names: string[]): number | undefined => {
    for (const n of names) {
      const key = n.toLowerCase();
      if (headerMap[key] !== undefined) return headerMap[key];
    }
    return undefined;
  };

  const nameCol = getCol("name", "label") ?? DETAILS_COL.NAME;
  const helperCol =
    getCol("helper text", "helper", "help") ?? DETAILS_COL.HELPER_TEXT;
  const typeCol = getCol("type") ?? DETAILS_COL.TYPE;
  const optionsCol = getCol("options") ?? DETAILS_COL.OPTIONS;
  const idCol = getCol("id") ?? DETAILS_COL.ID;

  // Track rows where we auto-generate IDs so we can write them back to the sheet
  const pendingIdWrites: { row: number; value: string }[] = [];

  return details
    .map((row, index) => {
      const name = String(row[nameCol] || "").trim();
      if (!name) {
        getScopedLogger("PayloadBuilder").info(`Skipping Details row ${index + 2}: empty field name`);
        return null; // Skip blank rows
      }

      const helperText = String(row[helperCol] || "");
      const typeRaw = String(row[typeCol] || "")
        .trim()
        .toLowerCase();
      const optionsStr = String(row[optionsCol] || "");
      const idStr = String(row[idCol] || "").trim();

      let type: FieldType;
      let options: SelectOption[] | undefined;

      const typeKey = typeRaw.charAt(0); // first letter covers translations (t=texto/text, s=select/single, m=multi, n=numero/number)

      switch (typeKey) {
        case "m":
          type = "selectMultiple";
          options = parseOptions(optionsStr);
          break;
        case "n":
          type = "number";
          break;
        case "t":
          type = "text";
          break;
        // default + 's' + blank → single choice
        case "s":
        case "":
        default:
          type = "selectOne";
          options = parseOptions(optionsStr);
      }

      // selectOne/selectMultiple require at least one option; skip invalid rows
      if (
        (type === "selectOne" || type === "selectMultiple") &&
        (!options || options.length === 0)
      ) {
        getScopedLogger("PayloadBuilder").warn(
          `Skipping Details row ${index + 2}: select field without options`,
        );
        return null;
      }

      // Deduplicate option values - this is a final safety net
      // to catch duplicates that would cause API errors
      if (options && options.length > 0) {
        const seenValues = new Set<string>();
        const uniqueOptions: SelectOption[] = [];
        const removedCount = { count: 0 };

        for (const opt of options) {
          if (seenValues.has(opt.value)) {
            removedCount.count++;
          } else {
            seenValues.add(opt.value);
            uniqueOptions.push(opt);
          }
        }

        if (removedCount.count > 0) {
          getScopedLogger("PayloadBuilder").warn(
            `Field "${name}" (row ${index + 2}): Removed ${removedCount.count} duplicate option value(s)`,
          );
          options = uniqueOptions;
        }
      }

      const autoId = idStr || slugify(name) || `field-${index + 1}`;
      if (!idStr && detailsSheet && idCol !== undefined) {
        pendingIdWrites.push({ row: index + 2, value: autoId }); // sheet rows are 1-based, +1 for header, +1 to reach first data row
      }

      return {
        id: autoId, // Use explicit ID if provided, otherwise slugify name
        tagKey: autoId, // API v2 requires tagKey
        name,
        type,
        description: helperText || undefined,
        helperText: helperText || undefined,
        options,
        appliesTo: ["observation", "track"],
        // Note: required property is not set from spreadsheet - universal flag is separate from required
      } as Field;
    })
    .filter((field): field is Field => field !== null)
    .map((field, arrayIndex) => {
      // After we know the sheet exists, write back pending IDs in a single pass to avoid surprises
      if (
        arrayIndex === 0 &&
        pendingIdWrites.length > 0 &&
        detailsSheet &&
        idCol !== undefined
      ) {
        try {
          pendingIdWrites.forEach((write) => {
            detailsSheet!.getRange(write.row, idCol + 1).setValue(write.value);
          });
        } catch (err) {
          getScopedLogger("PayloadBuilder").warn(
            "Failed to write auto-generated field IDs to Details sheet:",
            err,
          );
        }
      }
      return field;
    });
}

/**
 * Parses comma-separated options string into SelectOption array
 * Supports two formats:
 * - "Label" -> {value: slugify(Label), label: Label}
 * - "value:Label" -> {value: value, label: Label}
 */
function parseOptions(optionsStr: string): SelectOption[] | undefined {
  if (!optionsStr) return undefined;

  const opts = optionsStr
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
  if (opts.length === 0) return undefined;

  return opts.map((opt) => {
    // Check for "value:label" format
    const colonIndex = opt.indexOf(":");
    if (colonIndex > 0) {
      const value = opt.substring(0, colonIndex);
      const label = opt.substring(colonIndex + 1);
      return { value, label };
    }
    // Default format: just label
    return {
      value: slugify(opt),
      label: opt,
    };
  });
}

/**
 * Builds categories array from Categories sheet
 * Categories are built in exact spreadsheet order for setCategorySelection
 */
function buildCategories(data: SheetData, fields: Field[]): Category[] {
  const categories = data.Categories?.slice(1) || [];
  if (categories.length === 0) {
    return [];
  }

  // Try to get display values if available (for cleaner token parsing), but handle failure gracefully
  let displayValues: string[][] = [];
  let backgroundColors: string[][] = [];
  let categoriesSheet: GoogleAppsScript.Spreadsheet.Sheet | null = null;
  try {
    categoriesSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Categories");
    if (categoriesSheet) {
      displayValues = categoriesSheet
        .getRange(2, 1, categories.length, categoriesSheet.getLastColumn())
        .getDisplayValues();
      backgroundColors = categoriesSheet
        .getRange(2, 1, categories.length, 1)
        .getBackgrounds();
    }
  } catch (e) {
    getScopedLogger("PayloadBuilder").warn(
      "Could not fetch display values or backgrounds from sheet (might be running in test/mock mode):",
      e,
    );
  }

  // Use headers from data directly
  const headerRow = data.Categories?.[0] || [];
  const headerMap: Record<string, number> = {};
  headerRow.forEach((h, idx) => {
    const key = String(h || "")
      .trim()
      .toLowerCase();
    if (key) headerMap[key] = idx;
  });

  const getCol = (...names: string[]): number | undefined => {
    for (const n of names) {
      const key = n.toLowerCase();
      if (headerMap[key] !== undefined) return headerMap[key];
    }
    return undefined;
  };

  const nameCol = getCol("name") ?? CATEGORY_COL.NAME;
  const iconCol = getCol("icon", "icons") ?? CATEGORY_COL.ICON;
  const categoryIdCol = getCol("category id", "id");
  const iconIdCol = getCol("icon id", "iconid");
  const fieldsCol = getCol("fields", "details") ?? CATEGORY_COL.FIELDS;
  const appliesCol = getCol("applies", "tracks", "applies to", "appliesto");
  const colorCol = getCol("color");

  // Build map from field name to field ID for converting defaultFieldIds
  const fieldNameToId = new Map<string, string>();
  if (fields && fields.length > 0) {
    for (const field of fields) {
      if (field.name && field.id) {
        fieldNameToId.set(field.name, field.id);
        fieldNameToId.set(field.name.toLowerCase(), field.id);
        fieldNameToId.set(field.id, field.id);
        fieldNameToId.set(field.id.toLowerCase(), field.id);
      }
    }
  }
  // Identify universal fields (fields marked as Universal in Details sheet)
  const details = data.Details?.slice(1) || [];
  const universalFieldIds: string[] = [];
  details.forEach((row) => {
    const universalVal = row[DETAILS_COL.UNIVERSAL];
    const isUniversal =
      universalVal === true ||
      universalVal === "TRUE" ||
      universalVal === "true";
    if (isUniversal) {
      const name = String(row[DETAILS_COL.NAME] || "");
      const idStr = String(row[DETAILS_COL.ID] || "").trim();
      const fieldId = idStr || slugify(name);
      if (fieldId) {
        universalFieldIds.push(fieldId);
      }
    }
  });

  let hasTrackCategory = false;
  let hasObservationCategory = false;

  const mappedCategories = categories
    .map((row, index) => {
      const getVal = (col?: number) => (col === undefined ? "" : row[col]);

      const name = String(getVal(nameCol) || "").trim();
      if (!name) {
        getScopedLogger("PayloadBuilder").info(
          `Skipping Categories row ${index + 2}: empty category name`,
        );
        return null; // Skip blank rows
      }

      const iconData = String(getVal(iconCol) || "").trim();
      const displayVal =
        fieldsCol !== undefined
          ? (displayValues[index]?.[fieldsCol] ?? "")
          : "";
      const fieldsVal = getVal(fieldsCol);

      // Prefer what the user sees (display value), fall back to raw value
      let fieldsTokens = normalizeFieldTokens(displayVal);
      if (fieldsTokens.length === 0) {
        fieldsTokens = normalizeFieldTokens(fieldsVal);
      }
      // Applies column: accepts o/observation, t/track, comma separated; defaults to observation
      let appliesTo: string[] = ["observation"];
      if (appliesCol !== undefined) {
        const raw = getVal(appliesCol);
        const rawText = raw === undefined || raw === null ? "" : String(raw);
        const trimmedRaw = rawText.trim();

        const parseTokens = (): Array<"observation" | "track"> =>
          rawText
            .split(",")
            .map((t) => t.trim().toLowerCase())
            .filter(Boolean)
            .map((t) => {
              if (t.startsWith("o")) return "observation";
              if (t.startsWith("t")) return "track";
              return undefined;
            })
            .filter(
              (v): v is "observation" | "track" => v === "observation" || v === "track",
            );

        if (!trimmedRaw) {
          if (AUTO_CREATED_APPLIES_COLUMN && index === 0 && categoriesSheet) {
            getScopedLogger("PayloadBuilder").warn(
              `Auto-created Applies column: defaulting category row ${index + 2} to track + observation.`,
            );
            appliesTo = ["track", "observation"];
            try {
              categoriesSheet
                .getRange(index + 2, appliesCol + 1)
                .setValue("track, observation");
            } catch (seedError) {
              getScopedLogger("PayloadBuilder").warn(
                "Failed to seed Applies cell during auto-create fallback:",
                seedError,
              );
            }
          } else {
            appliesTo = ["observation"];
          }
        } else {
          const tokens = parseTokens();
          appliesTo =
            tokens.length === 0
              ? AUTO_CREATED_APPLIES_COLUMN && index === 0
                ? ["track", "observation"]
                : ["observation"]
              : Array.from(new Set(tokens));
        }
      }

      if (appliesTo.includes("track")) {
        hasTrackCategory = true;
      }
      if (appliesTo.includes("observation")) {
        hasObservationCategory = true;
      }

      // Determine category ID from Column E (Category ID) when available
      const idFromSheet =
        categoryIdCol !== undefined
          ? String(getVal(categoryIdCol) || "").trim()
          : "";
      let categoryId = idFromSheet;
      if (!categoryId) {
        const fallbackId = slugify(name) || `category-${index + 1}`;
        categoryId = fallbackId;
        if (categoryIdCol !== undefined) {
          row[categoryIdCol] = categoryId;
        }
      }

      // Determine color from explicit column first, then fallback to background
      let color: string | undefined;
      const colorFromColumn =
        colorCol !== undefined ? String(getVal(colorCol) || "").trim() : "";
      if (colorFromColumn) {
        color = colorFromColumn;
      } else if (
        backgroundColors[index]?.[CATEGORY_COL.COLOR_BACKGROUND] &&
        backgroundColors[index][CATEGORY_COL.COLOR_BACKGROUND].toLowerCase() !==
          "#ffffff"
      ) {
        color = backgroundColors[index][CATEGORY_COL.COLOR_BACKGROUND];
      }

      // Convert field names to field IDs using actual IDs from Details sheet
      const explicitFieldIds: string[] = [];
      if (fieldsTokens.length > 0) {
        const fieldIds = fieldsTokens
          .map((name) => {
            const id =
              fieldNameToId.get(name) ||
              fieldNameToId.get(name.toLowerCase()) ||
              slugify(name);
            if (!id) {
              getScopedLogger("PayloadBuilder").info(
                `DEBUG: Could not map field token '${name}' to an ID. Available keys sample:`,
                Array.from(fieldNameToId.keys()).slice(0, 5),
              );
            }
            return id;
          })
          .filter(Boolean);
        explicitFieldIds.push(...fieldIds);
      }
      // Merge universal fields with explicit fields (universal fields first, then explicit)
      // Avoid relying on Set in Apps Script; manual de-dup to sidestep any V8 quirks
      const allFieldIds: string[] = [];
      const pushUnique = (id?: string) => {
        // Normalize to string in case Apps Script gives us String objects
        const val = id !== undefined && id !== null ? String(id) : "";
        if (!val) return;
        if (!allFieldIds.includes(val)) allFieldIds.push(val);
      };
      universalFieldIds.forEach(pushUnique);
      explicitFieldIds.forEach(pushUnique);
      // Always set both properties so downstream code (and the API) see the IDs.
      const defaultFieldIds = allFieldIds;

      const explicitIconId =
        iconIdCol !== undefined ? String(getVal(iconIdCol) || "").trim() : "";
      const resolvedIconId =
        explicitIconId || (iconData ? categoryId : undefined);

      return {
        id: categoryId,
        name,
        appliesTo,
        color,
        iconId: resolvedIconId,
        defaultFieldIds,
        fields: defaultFieldIds,
      } as Category;
    })
    .filter((cat): cat is Category => cat !== null);

  if (!hasObservationCategory) {
    const message =
      'At least one category must include "observation" in the Applies column. Please update the Categories sheet and try again.';
    try {
      if (typeof SpreadsheetApp !== "undefined") {
        SpreadsheetApp.getUi().alert(
          "Observation Applies Required",
          `${message}\n\nTip: mark the categories that should appear in the observation viewer with "observation" (you can combine it with "track").`,
          SpreadsheetApp.getUi().ButtonSet.OK,
        );
      }
    } catch (error) {
      getScopedLogger("PayloadBuilder").warn("Unable to show missing observation alert:", error);
    }
    throw new Error(message);
  }

  if (!hasTrackCategory) {
    const appliesColumnMissing = appliesCol === undefined;
    const wasAutoCreated = AUTO_CREATED_APPLIES_COLUMN === true;

    // UX improvement: if the Applies column is missing, auto-create + seed it instead of failing hard
    if (appliesColumnMissing || wasAutoCreated) {
      getScopedLogger("PayloadBuilder").warn(
        "Applies column missing or newly created; seeding first category with track + observation.",
      );

      try {
        if (categoriesSheet && typeof ensureAppliesColumn === "function") {
          ensureAppliesColumn(categoriesSheet);
          AUTO_CREATED_APPLIES_COLUMN = true;
        }

        // Seed first visible row with track + observation so downstream validation passes
        if (categoriesSheet) {
          try {
            const seedCell = categoriesSheet.getRange(2, 4);
            if (!String(seedCell.getValue() || "").trim()) {
              seedCell.setValue("track, observation");
            }
          } catch (seedError) {
            getScopedLogger("PayloadBuilder").warn(
              "Unable to seed Applies column default value:",
              seedError,
            );
          }
        }

        if (mappedCategories.length > 0) {
          mappedCategories[0].appliesTo = ["track", "observation"];
          hasTrackCategory = true;
          hasObservationCategory = true;
        }

        try {
          SpreadsheetApp.getActiveSpreadsheet().toast(
            "Added Applies column and set first category to track + observation. Please review column D.",
            "Applies column added",
            8,
          );
        } catch (toastError) {
          getScopedLogger("PayloadBuilder").warn("Unable to display Applies column toast:", toastError);
        }
      } catch (repairError) {
        getScopedLogger("PayloadBuilder").warn(
          "Auto-create Applies column failed; falling back to validation error.",
          repairError,
        );
      }
    }

    if (!hasTrackCategory) {
      const message =
        'At least one category must include "track" in the Applies column. Please update the Categories sheet and try again.';
      try {
        if (typeof SpreadsheetApp !== "undefined") {
          SpreadsheetApp.getUi().alert(
            "Track Applies Required",
            `${message}

Tip: mark the categories that should appear in the track viewer with "track" (you can combine it with "observation").`,
            SpreadsheetApp.getUi().ButtonSet.OK,
          );
        }
      } catch (error) {
        getScopedLogger("PayloadBuilder").warn("Unable to show missing track alert:", error);
      }
      throw new Error(message);
    }
  }

  return mappedCategories;
}

/**
 * Normalizes field list tokens from the Categories sheet.
 * Supports:
 * - Strings with commas/semicolons/newlines
 * - Multi-select dropdown arrays
 * - Fallback to stringified objects
 */
function normalizeFieldTokens(raw: any): string[] {
  if (raw === undefined || raw === null) return [];

  const tokens: string[] = [];

  const pushTokensFromString = (str: string) => {
    const cleaned = str
      .replace(/\n/g, ",")
      .replace(/;/g, ",")
      .replace(/[•·]/g, ",")
      .replace(/，/g, ",");
    cleaned.split(",").forEach((tok) => {
      const t = tok.trim();
      if (t) tokens.push(t);
    });
  };

  if (Array.isArray(raw)) {
    raw.flat(Infinity).forEach((item) => {
      if (typeof item === "string") {
        pushTokensFromString(item);
      } else if (item !== undefined && item !== null) {
        pushTokensFromString(String(item));
      }
    });
  } else if (typeof raw === "string") {
    pushTokensFromString(raw);
  } else if (raw !== undefined && raw !== null) {
    pushTokensFromString(String(raw));
  }

  return tokens;
}

/**
 * Builds icons array from Icons sheet and Categories sheet column B
 * This preserves all icons from imported configs AND standard workflow icons
 */
function buildIconsFromSheet(data: SheetData): Icon[] {
  // Use headers from data directly
  const headerRow = data.Categories?.[0] || [];
  const headerMap: Record<string, number> = {};
  headerRow.forEach((h, idx) => {
    const key = String(h || "")
      .trim()
      .toLowerCase();
    if (key) headerMap[key] = idx;
  });

  const getCol = (...names: string[]): number | undefined => {
    for (const n of names) {
      const key = n.toLowerCase();
      if (headerMap[key] !== undefined) return headerMap[key];
    }
    return undefined;
  };

  const nameCol = getCol("name") ?? CATEGORY_COL.NAME;
  const iconCol = getCol("icon", "icons") ?? CATEGORY_COL.ICON;
  const categoryIdCol = getCol("category id", "id");
  const iconIdCol = getCol("icon id", "iconid");

  const iconsById = new Map<string, Icon>();
  const addIcon = (icon: Icon) => {
    if (!icon.id) return;
    const existing = iconsById.get(icon.id);
    if (existing) {
      // Prefer keeping svgData; if existing lacks it and new has it, replace
      if (!existing.svgData && icon.svgData) {
        iconsById.set(icon.id, icon);
      } else if (!existing.svgUrl && icon.svgUrl) {
        iconsById.set(icon.id, { ...existing, svgUrl: icon.svgUrl });
      }
      return;
    }
    iconsById.set(icon.id, icon);
  };

  // Icons from Categories sheet (primary source)
  const categories = data.Categories?.slice(1) || [];
  categories.forEach((row, index) => {
    const getVal = (col?: number) => (col === undefined ? "" : row[col]);

    const name = String(getVal(nameCol) || "").trim();
    const iconRaw = getVal(iconCol);
    const iconStr = typeof iconRaw === "string" ? iconRaw.trim() : "";
    if (!name || !iconStr) return;

    const idFromSheet =
      categoryIdCol !== undefined
        ? String(getVal(categoryIdCol) || "").trim()
        : "";
    let categoryId = idFromSheet || slugify(name) || `category-${index + 1}`;
    if (!idFromSheet && categoryIdCol !== undefined) {
      row[categoryIdCol] = categoryId;
    }

    const iconIdFromSheet =
      iconIdCol !== undefined ? String(getVal(iconIdCol) || "").trim() : "";
    const iconId = iconIdFromSheet || categoryId;

    const parsed = parseIconSource(iconStr);
    if (!parsed) return;

    addIcon({ id: iconId, ...parsed });
  });

  // Icons sheet (secondary source to preserve imported/extra icons)
  const iconsData = data["Icons"];
  if (iconsData && iconsData.length > 1) {
    // Skip header (row 0)
    const rows = iconsData.slice(1);
    rows.forEach((row) => {
      const iconId = String(row[0] || "").trim();
      const iconStr = String(row[1] || "").trim();
      if (!iconId || !iconStr) return;
      const parsed = parseIconSource(iconStr);
      if (!parsed) return;
      addIcon({ id: iconId, ...parsed });
    });
  }

  return Array.from(iconsById.values());
}

/**
 * Normalizes supported icon sources into svgData/svgUrl
 * Accepts inline <svg>, data:image/svg+xml (plain or base64), Google Drive links,
 * and public URLs ending in .svg
 */
function parseIconSource(
  iconStr: string,
): { svgData?: string; svgUrl?: string } | null {
  const trimmed = iconStr.trim();
  if (!trimmed) return null;

  // Inline SVG markup
  if (trimmed.startsWith("<svg")) {
    return { svgData: trimmed };
  }

  // data:image/svg+xml (plain or base64)
  if (trimmed.toLowerCase().startsWith("data:image/svg+xml")) {
    const svg = decodeDataSvg(trimmed);
    if (svg && svg.trim().startsWith("<svg")) {
      return { svgData: svg.trim() };
    }
    return null;
  }

  // Google Drive file links → inline SVG (API cannot fetch Drive URLs)
  if (isGoogleDriveLink(trimmed)) {
    const svgData = loadDriveSvg(trimmed);
    return svgData ? { svgData } : null;
  }

  // Direct SVG URLs
  if (isSvgUrl(trimmed)) {
    return { svgUrl: trimmed };
  }

  return null;
}

function decodeDataSvg(dataUri: string): string | null {
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
  } catch (err) {
    getScopedLogger("PayloadBuilder").warn("Failed to decode data URI icon:", err);
    return null;
  }
}

function isGoogleDriveLink(url: string): boolean {
  return url.startsWith("https://drive.google.com/file/d/");
}

function loadDriveSvg(url: string): string | null {
  try {
    const fileId = url.split("/d/")[1]?.split("/")[0];
    if (!fileId) return null;

    const file = DriveApp.getFileById(fileId);
    const mime = file.getMimeType();
    const text = file.getBlob().getDataAsString();

    // Accept SVG mime types or raw SVG content
    const isSvgMime = mime?.toLowerCase().includes("svg");
    const looksLikeSvg = text.trim().startsWith("<svg");
    if (isSvgMime || looksLikeSvg) {
      return text.trim();
    }

    getScopedLogger("PayloadBuilder").warn(`Drive icon ${fileId} rejected: mime=${mime}`);
    return null;
  } catch (err) {
    getScopedLogger("PayloadBuilder").warn(`Failed to inline Drive icon ${url}:`, err);
    return null;
  }
}

function isSvgUrl(url: string): boolean {
  const lower = url.toLowerCase();
  return lower.startsWith("http") && lower.includes(".svg");
}

/**
 * Extracts language columns from translation sheet headers.
 * Handles common language names, ISO/BCP-47 codes, and custom "Name - ISO" formats.
 */
function extractLanguagesFromHeaders(
  headers: any[],
): { lang: string; index: number }[] {
  const allLanguages: LanguageMap = getAllLanguages();
  const resolveHeaderCode = createTranslationHeaderResolver(allLanguages);
  const columns: { lang: string; index: number }[] = [];
  const seenLanguages = new Set<string>();

  const headerB = String(headers[1] || "")
    .trim()
    .toLowerCase();
  const headerC = String(headers[2] || "")
    .trim()
    .toLowerCase();
  const hasMetaColumns = headerB.includes("iso") && headerC.includes("source");
  const languageStartIndex = hasMetaColumns
    ? 3
    : TRANSLATION_COL.FIRST_LANGUAGE;

  for (let i = languageStartIndex; i < headers.length; i++) {
    const header = String(headers[i] || "").trim();
    if (!header) continue;

    const langCode = resolveHeaderCode(header);
    if (!langCode) continue;

    const normalizedCode = langCode.toLowerCase();
    if (seenLanguages.has(normalizedCode)) {
      continue;
    }
    seenLanguages.add(normalizedCode);

    columns.push({ lang: langCode, index: i });
  }

  return columns;
}

type LocaleTranslationFields = Record<string, string>;

interface ApiTranslationLocale {
  category?: Record<string, LocaleTranslationFields>;
  field?: Record<string, LocaleTranslationFields>;
}

type ApiTranslationsByLocale = Record<string, ApiTranslationLocale>;

function splitTranslatedOptions(optionRow: string): string[] {
  if (!optionRow) {
    return [];
  }
  return optionRow
    .split(/[;,，、]/)
    .map((part) => part.trim())
    .filter(Boolean);
}

function buildTranslationsPayload(
  data: SheetData,
  categories: Category[],
  fields: Field[],
): ApiTranslationsByLocale {
  const translations: ApiTranslationsByLocale = {};

  const ensureLocaleEntry = (lang: string): ApiTranslationLocale => {
    if (!translations[lang]) {
      translations[lang] = {};
    }
    return translations[lang];
  };

  const ensureCategoryEntry = (
    lang: string,
    id: string,
  ): LocaleTranslationFields => {
    const localeEntry = ensureLocaleEntry(lang);
    localeEntry.category = localeEntry.category || {};
    localeEntry.category[id] = localeEntry.category[id] || {};
    return localeEntry.category[id]!;
  };

  const ensureFieldEntry = (
    lang: string,
    id: string,
  ): LocaleTranslationFields => {
    const localeEntry = ensureLocaleEntry(lang);
    localeEntry.field = localeEntry.field || {};
    localeEntry.field[id] = localeEntry.field[id] || {};
    return localeEntry.field[id]!;
  };

  const setCategoryNameTranslation = (
    lang: string,
    categoryId: string,
    value: string,
  ) => {
    if (!value) return;
    const entry = ensureCategoryEntry(lang, categoryId);
    entry.name = value;
  };

  const setFieldLabelTranslation = (
    lang: string,
    fieldId: string,
    value: string,
  ) => {
    if (!value) return;
    const entry = ensureFieldEntry(lang, fieldId);
    entry.label = value;
  };

  const setFieldHelperTextTranslation = (
    lang: string,
    fieldId: string,
    value: string,
  ) => {
    if (!value) return;
    const entry = ensureFieldEntry(lang, fieldId);
    entry.helperText = value;
  };

  const setFieldOptionTranslation = (
    lang: string,
    fieldId: string,
    optionIndex: number,
    value: string,
  ) => {
    if (!value) return;
    const entry = ensureFieldEntry(lang, fieldId);
    entry[`options.${optionIndex}`] = value;
  };

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const catTransSheet = spreadsheet.getSheetByName("Category Translations");
  const labelTransSheet = spreadsheet.getSheetByName(
    "Detail Label Translations",
  );
  const helperTransSheet = spreadsheet.getSheetByName(
    "Detail Helper Text Translations",
  );
  const optionTransSheet = spreadsheet.getSheetByName(
    "Detail Option Translations",
  );

  const allLangs = new Set<string>();

  const getLangColumnsFromSheet = (
    sheet: GoogleAppsScript.Spreadsheet.Sheet | null,
  ): { lang: string; index: number }[] => {
    if (!sheet) return [];
    const lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) return [];
    const headerRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const columns = extractLanguagesFromHeaders(headerRow);
    columns.forEach((entry) => allLangs.add(entry.lang));
    return columns;
  };

  const catLangColumns = getLangColumnsFromSheet(catTransSheet);
  const labelLangColumns = getLangColumnsFromSheet(labelTransSheet);
  const helperLangColumns = getLangColumnsFromSheet(helperTransSheet);
  const optionLangColumns = getLangColumnsFromSheet(optionTransSheet);

  // If no translation sheets found with languages, return empty translations
  if (allLangs.size === 0) {
    getScopedLogger("PayloadBuilder").info("No translation sheets found with language columns");
    return translations;
  }

  const langs = Array.from(allLangs);
  getScopedLogger("PayloadBuilder").info(`Found ${langs.length} languages for translations:`, langs);

  // Initialize translations structure
  // Process category translations - match by name to handle blank rows
  // Only process if Category Translations sheet exists
  if (catTransSheet) {
    const catTrans = data["Category Translations"]?.slice(1) || [];
    const catNameToId = new Map(categories.map((c) => [c.name, c.id]));
    for (const row of catTrans) {
      const sourceName = String(row[TRANSLATION_COL.SOURCE_TEXT] || "").trim();
      const catId = catNameToId.get(sourceName);
      if (!catId) continue; // Skip rows that don't match a category

      for (const entry of catLangColumns) {
        const value = String(row[entry.index] || "").trim();
        if (value) {
          setCategoryNameTranslation(entry.lang, catId, value);
        }
      }
    }
  }

  // Process field label translations - match by name to handle blank rows
  const labelTrans = data["Detail Label Translations"]?.slice(1) || [];
  const fieldNameToId = new Map(fields.map((f) => [f.name, f.id]));
  for (const row of labelTrans) {
    const sourceName = String(row[TRANSLATION_COL.SOURCE_TEXT] || "").trim();
    const fieldId = fieldNameToId.get(sourceName);
    if (!fieldId) continue;

    for (const entry of labelLangColumns) {
      const value = String(row[entry.index] || "").trim();
      if (value) {
        setFieldLabelTranslation(entry.lang, fieldId, value);
      }
    }
  }

  // Process field helper text translations
  // Match by helper text value (not by index) to handle blank rows in Details sheet
  // Build array of fields with their corresponding helper text for matching
  const helperTrans = data["Detail Helper Text Translations"]?.slice(1) || [];
  const fieldsWithHelperText: Array<{ field: Field; helperText: string }> = [];
  for (const field of fields) {
    const helperTextValue = field.helperText || field.description;
    if (helperTextValue) {
      fieldsWithHelperText.push({ field, helperText: helperTextValue });
    }
  }

  for (const row of helperTrans) {
    const sourceHelperText = String(
      row[TRANSLATION_COL.SOURCE_TEXT] || "",
    ).trim();

    // Find ALL fields that match this helper text (handles duplicate helper texts)
    const matchingFields = fieldsWithHelperText.filter(
      (item) => item.helperText === sourceHelperText,
    );
    if (matchingFields.length === 0) continue;

    for (const { field } of matchingFields) {
      const fieldId = field.id;
      for (const entry of helperLangColumns) {
        const value = String(row[entry.index] || "").trim();
        if (value) {
          setFieldHelperTextTranslation(entry.lang, fieldId, value);
        }
      }
    }
  }

  // Process field option translations
  // Match by options string (not by index) to handle blank rows in Details sheet
  // Build array of fields with their corresponding options string for matching
  const optionTrans = data["Detail Option Translations"]?.slice(1) || [];
  const fieldsWithOptions: Array<{ field: Field; optionsStr: string }> = [];
  for (const field of fields) {
    if (field.options && field.options.length > 0) {
      // Reconstruct the options string as it appears in Details sheet column D
      const optionsStr = field.options
        .map((o) => {
          const value = o?.value || "";
          const label = o?.label || "";
          if (!label) return "";
          return value === slugify(label) ? label : `${value}:${label}`;
        })
        .filter(Boolean)
        .join(", ");
      if (optionsStr) {
        fieldsWithOptions.push({ field, optionsStr });
      }
    }
  }

  for (const row of optionTrans) {
    const sourceOptionsStr = String(
      row[TRANSLATION_COL.SOURCE_TEXT] || "",
    ).trim();

    // Find ALL fields that match this options string (handles duplicate option lists)
    const matchingFields = fieldsWithOptions.filter(
      (item) => item.optionsStr === sourceOptionsStr,
    );
    if (matchingFields.length === 0) continue;

    for (const { field } of matchingFields) {
      if (!field.options || field.options.length === 0) continue;

      const fieldId = field.id;
      for (const entry of optionLangColumns) {
        const optStr = String(row[entry.index] || "").trim();
        if (!optStr) continue;

        const translatedOpts = splitTranslatedOptions(optStr);

        for (
          let k = 0;
          k < translatedOpts.length && k < field.options.length;
          k++
        ) {
          setFieldOptionTranslation(entry.lang, fieldId, k, translatedOpts[k]);
        }
      }
    }
  }

  // Remove languages without any translation data
  Object.keys(translations).forEach((lang) => {
    const entry = translations[lang];
    const hasCategoryTranslations = Boolean(
      entry.category && Object.keys(entry.category).length > 0,
    );
    const hasFieldTranslations = Boolean(
      entry.field && Object.keys(entry.field).length > 0,
    );

    if (!hasCategoryTranslations && !hasFieldTranslations) {
      delete translations[lang];
    }
  });

  return translations;
}
