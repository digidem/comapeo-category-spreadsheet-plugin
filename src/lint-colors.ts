/**
 * Shared lint color constants.
 *
 * Used by `src/lint.ts` (production) and `src/test/testLint.ts` (mock framework).
 * Keeping them in one place prevents silent test drift when the palette changes.
 */

const LINT_WARNING_BACKGROUND_COLORS = [
  "#FFC7CE",
  "#FFEB9C",
  "#FFFFCC",
  "#FFF2CC",
  "#FFF3CD",
];

const LINT_WARNING_FONT_COLORS = ["red", "orange", "#FF0000"];
