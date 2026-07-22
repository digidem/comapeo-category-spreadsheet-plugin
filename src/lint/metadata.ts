/// <reference path="shared.ts" />

/**
 * Mirrors buildLocales() normalization for Metadata!primaryLanguage values.
 * Accepts either a recognized language name or a locale token like en / pt-BR.
 */
function normalizeMetadataPrimaryLanguageValue(value: string): string | null {
  const trimmedValue = value.trim();
  if (!trimmedValue) return null;

  // Use the shared resolver so lint accepts the same BCP-47 tags the builder does
  // (e.g. zh-Hant, es-419) instead of a narrower ad-hoc pattern.
  const resolved = resolvePrimaryLanguageInput(trimmedValue);
  if (resolved?.code) {
    return resolved.code;
  }

  return null;
}

function getMetadataPrimaryLanguageLintMessage(value: string): string {
  return `Metadata primaryLanguage: "${value}" is not a recognized language name or locale code. Use a display name or locale token (e.g. "English", "Português", "en", "pt-BR").`;
}

/**
 * Validates explicit Metadata sheet values for unsafe characters.
 * Only checks rows that are present – does NOT flag missing rows.
 * Keys validated: name, version, primaryLanguage.
 */
function lintMetadataSheet(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const metadataSheet = spreadsheet.getSheetByName("Metadata");
  if (!metadataSheet) return;

  const lastRow = metadataSheet.getLastRow();
  if (lastRow <= 1) return; // Header-only or empty

  // Clear previous lint artifacts on both key (A) and value (B) columns
  clearLintArtifacts(metadataSheet.getRange(2, 1, lastRow - 1, 2));

  // Matches containsUnsafeNameCharacters() in importService: slashes, backslashes, ellipsis
  // These are the only characters that strict build validation actually rejects (errors).
  const STRICT_UNSAFE_PATTERN = /[\\/]|\.\.\./;

  const data = metadataSheet.getRange(2, 1, lastRow - 1, 2).getValues();

  // Track seen keys to detect duplicates. The builder uses exact key
  // matching (sheetData[i][0] === key), so lint must do the same to
  // avoid false-positive duplicate warnings on differently-cased keys
  // (e.g. "PrimaryLanguage" vs "primaryLanguage" are distinct to the builder).
  const seenKeys = new Set<string>();
  let resolvedPrimaryLanguage = false;
  // Track whether a valid primaryLanguage has been seen.
  // The builder's buildLocales() normalizes via normalizeLocaleInput() and skips
  // invalid values, continuing to the next row. Only valid values block later rows.
  let seenValidPrimaryLanguage = false;

  // Keys the builder looks up with exact equality (no trim). If a cell
  // contains " name " the builder will NOT match it and will silently
  // append a default row instead. Lint must use exact key matching for
  // these to avoid false parity.
  const EXACT_MATCH_KEYS = new Set(["name", "version", "description", "legacyCompat"]);

  for (let i = 0; i < data.length; i++) {
    const key = String(data[i][0] ?? "");
    const trimmedKey = key.trim();
    const rawValue = data[i][1];
    // Match builder semantics: no trim, no falsy coercion (builder uses
    // String(sheetData[i][1]) which serializes 0 as "0").
    const value = String(rawValue ?? "");
    // Trimmed variant used only for primaryLanguage validation, where the
    // language name should be trimmed for lookup purposes.
    const trimmedValue = String(rawValue ?? "").trim();
    const row = i + 2;

    // Detect whitespace-padded keys that the builder will not match.
    // e.g. " name " looks like "name" after trimming, but the builder's
    // exact equality check (sheetData[i][0] === "name") will miss it.
    if (
      trimmedKey &&
      key !== trimmedKey &&
      EXACT_MATCH_KEYS.has(trimmedKey)
    ) {
      const cell = metadataSheet.getRange(row, 1);
      appendLintNote(
        cell,
        `Metadata key "${key}" has leading/trailing whitespace. The builder uses exact key matching and will not recognize this row — it will be silently ignored and a default value generated instead. Remove the whitespace so the key is exactly "${trimmedKey}".`,
        "warning",
      );
    }

    // Detect case-only typos in metadata keys (e.g. "Name" instead of "name").
    // The builder uses exact key matching, so a different case is silently ignored.
    // Only flag when the key has no whitespace issues and doesn't match exactly.
    const allRecognizedKeys = new Set([
      ...EXACT_MATCH_KEYS,
      "primaryLanguage",
    ]);
    const lowerTrimmedKey = trimmedKey.toLowerCase();
    if (
      trimmedKey &&
      key === trimmedKey && // no whitespace issues (already checked above)
      !allRecognizedKeys.has(trimmedKey) && // not an exact match
      [...allRecognizedKeys].some((k) => k.toLowerCase() === lowerTrimmedKey) // matches when lowercased
    ) {
      const correctKey = [...allRecognizedKeys].find(
        (k) => k.toLowerCase() === lowerTrimmedKey,
      )!;
      const cell = metadataSheet.getRange(row, 1);
      appendLintNote(
        cell,
        `Metadata key "${key}" differs in casing from the recognized key "${correctKey}". The builder uses exact key matching and will not recognize this row — it will be silently ignored. Use exactly "${correctKey}".`,
        "warning",
      );
    }

    // Flag duplicate keys as a warning — the builder uses only the first row.
    // Use exact key matching (same as builder) for non-primaryLanguage keys,
    // trimmed matching for primaryLanguage (builder resolves it differently).
    const effectiveKey = trimmedKey === "primaryLanguage" ? trimmedKey : key;
    const isDuplicate = effectiveKey && seenKeys.has(effectiveKey);

    if (trimmedKey === "primaryLanguage") {
      const cell = metadataSheet.getRange(row, 2);
      if (trimmedKey) seenKeys.add(trimmedKey);
      if (!trimmedValue) continue;

      // Track whether a VALID primaryLanguage has been seen — the builder's
      // buildLocales() normalizes values via normalizeLocaleInput() and skips
      // invalid ones, continuing to the next row. Only valid values block
      // later rows from being used.
      // (trimmedValue is guaranteed non-empty here due to the continue above.)
      const wasPriorValid = seenValidPrimaryLanguage;

      // Mirror buildLocales(): accept recognized display names and ISO-style locale tokens.
      const normalizedPrimaryLanguage = normalizeMetadataPrimaryLanguageValue(trimmedValue);
      const isValid = normalizedPrimaryLanguage !== null;

      // Only mark as "seen" when the builder would actually use this value.
      if (isValid) seenValidPrimaryLanguage = true;

      if (isDuplicate) {
        if (resolvedPrimaryLanguage || wasPriorValid) {
          // A valid primaryLanguage was already found — the builder uses the
          // first valid occurrence, so this row is definitely ignored.
          appendLintNote(
            cell,
            'Duplicate metadata key "primaryLanguage". The builder uses the first valid occurrence — this row is ignored.',
            "warning",
          );
        } else if (!isValid) {
          // No prior non-empty primaryLanguage AND this one is invalid — still a
          // chance later rows could be valid (builder would skip this invalid value).
          appendLintNote(
            cell,
            getMetadataPrimaryLanguageLintMessage(trimmedValue),
            "error",
          );
        } else {
          // No prior valid primaryLanguage and this one IS valid — it would
          // become the effective value if no earlier valid row claimed it.
          appendLintNote(
            cell,
            'Duplicate metadata key "primaryLanguage". This row is used only if all earlier primaryLanguage rows are blank or have invalid locale codes.',
            "warning",
          );
        }
        if (!isValid || resolvedPrimaryLanguage || wasPriorValid) continue;
      } else {
        if (!isValid) {
          appendLintNote(
            cell,
            getMetadataPrimaryLanguageLintMessage(trimmedValue),
            "error",
          );
          // Don't set resolvedPrimaryLanguage — the builder scans until it finds a valid locale,
          // so subsequent valid primaryLanguage rows should still be checked.
          continue;
        }
      }

      // If this is a valid duplicate with no prior valid resolution, it becomes
      // the effective primaryLanguage — fall through to set resolvedPrimaryLanguage.
      resolvedPrimaryLanguage = true;
      continue;
    }

    if (isDuplicate) {
      const cell = metadataSheet.getRange(row, 2);
      appendLintNote(
        cell,
        `Duplicate metadata key "${key}". The builder only reads the first occurrence — this row is ignored.`,
        "warning",
      );
      continue;
    }
    if (effectiveKey) seenKeys.add(effectiveKey);

    // Use exact key matching for builder-recognized keys (name, version, etc.)
    // to match builder semantics (sheetData[i][0] === key).
    if (key === "name") {
      const cell = metadataSheet.getRange(row, 2);
      if (!trimmedValue) {
        appendLintNote(
          cell,
          'Metadata "name" is present but blank. Config generation keeps the blank value and strict validation will fail.',
          "error",
        );
        continue;
      }
      if (STRICT_UNSAFE_PATTERN.test(value)) {
        appendLintNote(
          cell,
          `Metadata "${key}" contains characters that will fail config generation: slashes, backslashes, and ellipses (…) are not allowed.`,
          "error",
        );
      }
      continue;
    }

    if (key === "version") {
      const cell = metadataSheet.getRange(row, 2);
      if (value) {
        // The builder always overwrites version with today's date (yy.MM.dd),
        // so unsafe-character checks are unnecessary — any slashes/backslashes
        // in the user's value are replaced before config generation.
        // Only advise when the value doesn't match the auto-generated pattern.
        const looksLikeAutoDate = /^\d{2}\.\d{2}\.\d{2}$/.test(value.trim());
        if (!looksLikeAutoDate) {
          appendLintNote(
            cell,
            `Metadata "version" is always overwritten with today's date (yy.MM.dd) when generating config. The current value "${value}" will be ignored.`,
            "advisory",
          );
        }
      }
      continue;
    }

    if (!value) continue; // Skip empty values for remaining keys
  }
}
