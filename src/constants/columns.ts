/** Canonical Categories sheet headers (single source of truth).
 * Shared by generate (payloadBuilder), import (importService, importCategory),
 * migration (apiService), and lint (src/lint/). See issues #61 / #52.
 * Order: Name | Icon | Fields | Applies | Category ID | Color | Icon ID
 */
const CATEGORY_HEADERS = [
  "Name",
  "Icon",
  "Fields",
  "Applies",
  "Category ID",
  "Color",
  "Icon ID",
];

/** Column indices for Categories sheet (0-based), matching CATEGORY_HEADERS */
const CATEGORY_COL = {
  NAME: 0,
  ICON: 1,
  FIELDS: 2,
  APPLIES: 3,
  CATEGORY_ID: 4,
  COLOR: 5,
  ICON_ID: 6,
  COLOR_BACKGROUND: 0, // fallback when no explicit color column
} as const;

/** Column indices for Details sheet (0-based)
 * Expected headers (current implementation): Name/Label | Helper Text | Type | Options | ID | Universal
 */
const DETAILS_COL = {
  NAME: 0,
  HELPER_TEXT: 1,
  TYPE: 2,
  OPTIONS: 3,
  ID: 4,
  UNIVERSAL: 5,
} as const;

/** Column indices for translation sheets (0-based) */
const TRANSLATION_COL = {
  SOURCE_TEXT: 0,
  FIRST_LANGUAGE: 1,
} as const;
