/**
 * Comprehensive language map covering common locales
 * Maps ISO 639-1 language codes to their native language names
 */
const ALL_LANGUAGES: Record<string, string> = {
  // Western European
  en: 'English',
  es: 'Español',
  pt: 'Português',
  fr: 'Français',
  de: 'Deutsch',
  it: 'Italiano',
  nl: 'Nederlands',
  sv: 'Svenska',
  no: 'Norsk',
  da: 'Dansk',
  fi: 'Suomi',
  // Eastern European
  pl: 'Polski',
  cs: 'Čeština',
  sk: 'Slovenčina',
  hu: 'Magyar',
  ro: 'Română',
  bg: 'Български',
  hr: 'Hrvatski',
  sr: 'Српски',
  uk: 'Українська',
  ru: 'Русский',
  // Asian
  zh: '中文',
  ja: '日本語',
  ko: '한국어',
  hi: 'हिन्दी',
  th: 'ไทย',
  vi: 'Tiếng Việt',
  id: 'Bahasa Indonesia',
  ms: 'Bahasa Melayu',
  // Middle Eastern & African
  ar: 'العربية',
  he: 'עברית',
  tr: 'Türkçe',
  fa: 'فارسی',
  sw: 'Kiswahili',
  // Other
  el: 'Ελληνικά',
  ca: 'Català',
  eu: 'Euskara',
  gl: 'Galego'
};

/**
 * Converts a string to an ASCII slug format. Used for ENTITY identifiers
 * (field IDs, category/preset IDs, icon IDs, config name) which the
 * `.comapeocat` packager restricts to `^[a-zA-Z0-9_-]+$` (see
 * package/src/writer.js SAFE_ID_REGEX). Non-Latin names fold to "" here and
 * callers fall back to `prefix-N` — that is intentional. Do NOT use this for
 * select option values; use canonicalizeOptionValue() instead, which preserves
 * every script.
 * @param input The input string to be converted.
 * @returns The slugified string.
 */
function slugify(input: string | any): string {
  if (!input) return "";

  const str = typeof input === "string" ? input : String(input);

  // Normalize accents and special characters by removing diacritics
  // This converts "Café" to "Cafe", "naïve" to "naive", etc.
  const normalized = str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");

  return normalized
    .toLocaleLowerCase("en-US")  // Use en-US locale to avoid Turkish 'I' → 'ı' issue
    .trim()
    .replace(/[^\w\s-]/g, "")    // Remove non-word characters (keep ASCII letters, numbers, underscore, whitespace, hyphen)
    .replace(/[\s_-]+/g, "-")    // Replace whitespace/underscore/hyphen sequences with single hyphen
    .replace(/^-+|-+$/g, "");    // Remove leading/trailing hyphens
}

/**
 * Canonicalizes a select-option label into its stored machine value while
 * preserving EVERY script (Thai, Vietnamese, Cyrillic, Greek, CJK, ...).
 *
 * Unlike slugify() — which is ASCII-only and folds diacritics for entity IDs —
 * this keeps Unicode letters and combining marks intact so distinct options
 * never collapse to the same value. Collapsing is what caused the linter to
 * flag and delete all-Thai (and all-Vietnamese/Cyrillic/Greek) option lists as
 * "duplicates". Diacritics are semantic in tone languages (má ≠ mà), so they
 * are preserved rather than stripped.
 *
 * Used in lockstep by the builder (parseOptions) and the linter
 * (parseCanonicalOptions) so generation and validation agree.
 *
 * @param input Raw option label from the spreadsheet.
 * @returns A Unicode-safe canonical value ("" if nothing usable remains —
 *   e.g. emoji- or punctuation-only input). Callers fall back to the raw
 *   label in that case so distinct options never collapse to one empty value.
 */
function canonicalizeOptionValue(input: string | any): string {
  if (!input) return "";

  const str = typeof input === "string" ? input : String(input);

  return (
    str
      // Composed form so a base + combining mark compares equal to its
      // precomposed code point (Thai is NFC-stable; this normalizes the rest).
      .normalize("NFC")
      // Drop variation selectors (emoji presentation hints). Without this,
      // "❤️" and "☕️" both reduce to a lone invisible U+FE0F and collide.
      .replace(/[\u{FE00}-\u{FE0F}]/gu, "")
      .toLocaleLowerCase("en-US") // en-US avoids the Turkish 'I' → 'ı' issue
      .trim()
      // Remove orphan combining marks — a mark run with NO base, i.e. at the
      // start or preceded by a non-letter/number/mark (space, punctuation).
      // The boundary class excludes \p{M} on purpose: a mark preceded by another
      // mark is part of the same cluster (Thai vowel + tone) and must be kept.
      .replace(/(^|[^\p{L}\p{N}\p{M}])(\p{M}+)/gu, "$1")
      // Keep letters, (attached) marks, and numbers from any script; preserve
      // diacritics. Whitespace/underscore/hyphen are kept as separators; every
      // other symbol/punctuation mark is removed.
      .replace(/[^\p{L}\p{M}\p{N}\s_-]/gu, "")
      .replace(/[\s_-]+/g, "-") // collapse separator runs to a single hyphen
      .replace(/^-+|-+$/g, "")
  ); // trim leading/trailing hyphens
}

function sanitizeIconSlug(slug: string | null | undefined): string {
  if (!slug) return "";
  const normalized = String(slug).trim();
  if (!normalized) return "";
  return normalized.replace(/(\.(svg|png))+$/i, "");
}

/**
 * Ensures a deterministic slug for spreadsheet-derived identifiers.
 * Falls back to a prefix + index pattern when the source cannot produce a slug.
 *
 * @param source - Raw value taken from the spreadsheet.
 * @param fallbackPrefix - Prefix to use when the slug is empty.
 * @param index - Zero-based index of the item in its collection, used for fallback uniqueness.
 */
function buildSlugWithFallback(source: string, fallbackPrefix: string, index: number = 0): string {
  const slug = slugify(source);
  if (slug !== "") {
    return slug;
  }

  const sanitizedPrefix = fallbackPrefix && fallbackPrefix.trim() !== ""
    ? slugify(fallbackPrefix)
    : "item";

  return `${sanitizedPrefix || "item"}-${index + 1}`;
}

/**
 * Generates the canonical tag key for CoMapeo fields.
 *
 * @param fieldName - Name column value from the Details sheet.
 * @param index - Zero-based index of the field to guarantee deterministic fallback keys.
 */
function createFieldTagKey(fieldName: string, index?: number): string {
  return buildSlugWithFallback(fieldName, "field", typeof index === "number" ? index : 0);
}

/**
 * Generates the canonical slug for presets/categories.
 *
 * @param presetName - Category name from the Categories sheet.
 * @param index - Zero-based index to ensure unique fallback slugs.
 */
function createPresetSlug(presetName: string, index?: number): string {
  return buildSlugWithFallback(presetName, "category", typeof index === "number" ? index : 0);
}

/**
 * Normalizes icon slugs by removing size suffix variants (e.g., "-100px", "-medium", "-large").
 * Used to match icon filenames with preset slugs during validation.
 *
 * @param slug - The slug to normalize
 * @returns Normalized slug without size suffixes
 */
function normalizeIconSlug(slug: string): string {
  if (!slug) return "";

  const parts = slug.split("-").filter((part) => part !== "");

  // Remove trailing size indicators like "100px", "2x", "small", "medium", "large"
  while (parts.length > 0) {
    const last = parts[parts.length - 1];
    if (/^(?:\d+px|\d+x|small|medium|large)$/.test(last)) {
      parts.pop();
      continue;
    }
    break;
  }

  return parts.join("-");
}

/**
 * Determines the field type based on the type string.
 * @param typeString The type string from the spreadsheet (e.g., "Text", "Number", "Multiple choice", "Select one").
 * @returns The CoMapeo field type.
 */
function getFieldType(typeString: string): "text" | "number" | "selectOne" | "selectMultiple" {
  const firstChar = typeString.charAt(0).toLowerCase();
  if (firstChar === "m") return "selectMultiple";
  if (firstChar === "n") return "number";
  if (firstChar === "t") return "text";
  return "selectOne";
}

/**
 * Parses field options from the options string.
 * @param typeString The type string to determine if options are needed.
 * @param optionsString The comma-separated options string.
 * @param fieldKey The canonical field key used to build deterministic fallback option values.
 * @returns Array of option objects with label and value, or undefined for non-select fields.
 */
function getFieldOptions(
  typeString: string,
  optionsString: string,
  fieldKey?: string,
): Array<{ label: string; value: string }> | undefined {
  const fieldType = getFieldType(typeString);
  if (fieldType === "number" || fieldType === "text") return undefined;
  return optionsString
    .split(",")
    .map((opt) => opt.trim())
    .filter((opt) => opt !== "")
    .map((opt, index) => ({
      label: opt,
      value: createOptionValue(opt, fieldKey, index),
    }));
}

/**
 * Produces the canonical value for select field options with deterministic fallbacks.
 *
 * @param label - Option label taken from the spreadsheet.
 * @param fieldKey - Canonical field key if already computed.
 * @param index - Zero-based option index for fallback uniqueness.
 */
function createOptionValue(label: string, fieldKey: string | undefined, index: number): string {
  return canonicalizeOptionValue(label) || label;
}

// =============================================================================
// Category selection helpers (persisted in Script Properties)
// =============================================================================

const CATEGORY_SELECTION_KEY = "CATEGORY_SELECTION";

/**
 * Persists the current category order for later retrieval.
 */
function setCategorySelection(categoryIds: string[]): void {
  const cleaned = (categoryIds || []).map((id) => String(id).trim()).filter(Boolean);
  PropertiesService.getScriptProperties().setProperty(
    CATEGORY_SELECTION_KEY,
    JSON.stringify(cleaned),
  );
}

/**
 * Retrieves the stored category order. Returns a copy.
 */
function getCategorySelection(): string[] {
  const raw = PropertiesService.getScriptProperties().getProperty(
    CATEGORY_SELECTION_KEY,
  );
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      return [...parsed];
    }
  } catch (e) {
    console.warn("Failed to parse CATEGORY_SELECTION script property", e);
  }
  return [];
}
