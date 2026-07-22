/**
 * Shared lint color constants.
 *
 * Used by `src/lint/` modules (production) and `src/test/testLint.ts` (mock framework).
 * Keeping them in one place prevents silent test drift when the palette changes.
 */

/** Background color for error-severity lint notes. */
const LINT_ERROR_BG = "#FFC7CE";

/** Background color for warning-severity lint notes. */
const LINT_WARNING_BG = "#FFF2CC";

/** Background color for advisory-severity lint notes. */
const LINT_ADVISORY_BG = "#FFFFCC";

/** Background color for critical primary-column mismatches (validateSheetConsistency). */
const LINT_CRITICAL_BG = "#FF0000";

/** Font color paired with LINT_CRITICAL_BG for readability. */
const LINT_CRITICAL_FONT = "#FFFFFF";

const LINT_WARNING_BACKGROUND_COLORS = [
  LINT_ERROR_BG,
  "#FFEB9C",
  LINT_ADVISORY_BG,
  LINT_WARNING_BG,
  "#FFF3CD",
];

const LINT_WARNING_FONT_COLORS = ["red", "orange", LINT_CRITICAL_BG];
