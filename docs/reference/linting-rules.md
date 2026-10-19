# Linting rules

This page explains linting color codes and validation rules.

## Color codes

### Bright red background + white text (critical)

- **Meaning:** Translation column A does not match the source sheet.
- **Where:** Translation sheets, column A.
- **Why it matters:** Translation lookup fails and generation will fail.
- **Fix:** Restore the formula in column A (see Translation sheet formulas).

### Yellow background

- **Meaning:** Required data missing or warnings.
- **Examples:** Empty category name, missing icon, unused fields, row count mismatch.

### Red text

- **Meaning:** Invalid reference or URL.
- **Examples:** Icon link not accessible; Fields column references a missing field.

### Light red background

- **Meaning:** Invalid or duplicate values.
- **Examples:** Invalid field type, missing options, duplicate names, invalid translation headers, option count mismatches.

### Orange text

- **Meaning:** HTTP icon URL warning.
- **Fix:** Switch to HTTPS when possible.

## Categories sheet validation

1. **Name (A):**
   - Required; must be unique.
   - Auto-capitalized.

2. **Icon (B):**
   - Required; any non-empty valid icon reference.
   - Red text indicates invalid or inaccessible URLs.

3. **Fields (C):**
   - Each field name must exist in the Details sheet.
   - Red text indicates missing or misspelled fields.
   - Comma-separated lists are auto-formatted.

4. **Applies (D):**
   - Accepts `observation`, `track`, or both (comma-separated).
   - Abbreviations `o`, `t` are accepted.
   - At least one category must include `track`.
   - Invalid values are treated as empty and default to `observation`.

5. **Category ID (E):**
   - Auto-generated.

6. **Color (F) and Icon ID (G):**
   - Optional columns written by config imports.
   - Color falls back to the background color of column A when blank.

7. **Old import layout:**
   - Sheets with `ID` in column D and `Color` in column E (pre-canonical import layout, no `Applies`) get a warning note on the header row.
   - Fix: run **Generate Config** once to migrate, or re-import the config.

## Details sheet validation

1. **Name (A):**
   - Required; must be unique.
   - Yellow highlight if unused by any category.
   - Auto-capitalized.

2. **Helper Text (B):**
   - Optional.
   - Auto-capitalized.

3. **Type (C):**
   - Accepts `t`, `n`, `m`, `s` (case-insensitive).
   - Blank defaults to select-one.
   - Invalid values are highlighted.

4. **Options (D):**
   - Required for select and multiple types.
   - Optional for text and number types.
   - Comma-separated lists are auto-formatted.

5. **Universal (F):**
   - Must be `TRUE`, `FALSE`, or blank.

## Translation sheet validation

1. **Column A (primary column):**
   - Must match the source sheet exactly.
   - Bright red + white text indicates a mismatch.

2. **Headers (row 1):**
   - Must be a valid language name, ISO code, BCP-47 tag, or "Name - ISO" format.

3. **Row counts:**
   - Translation sheets should match the source row count.

4. **Option counts (Detail Option Translations):**
   - Each translation must have the same number of comma-separated options as the source.

5. **General formatting:**
   - Cells are auto-capitalized.
   - Whitespace-only cells are cleared.

See [Translation sheets](translation-sheets.md) for header formats and legacy layouts.

## Common scenarios

### Fresh spreadsheet

- Yellow highlights on empty required fields.
- Action: Fill required data.

### After adding icons

- Red text on invalid icon references.
- Action: Fix or replace invalid icon references.

### After auto-translation

- Possible light red if option counts do not match.
- Action: Ensure option lists have matching counts.

### After import

- Bright red translation mismatches if formulas broke.
- Yellow highlights for unused fields.

## Fixing translation sheet mismatches

1. Open the translation sheet with bright red column A.
2. Clear column A values (rows below the header).
3. In cell A2, re-enter the formula that matches the source sheet:

- Category Translations: `=Categories!A2:A[lastRow]`
- Detail Label Translations: `=Details!A2:A[lastRow]`
- Detail Helper Text Translations: `=Details!B2:B[lastRow]`
- Detail Option Translations: `=Details!D2:D[lastRow]`

Then re-run **Lint Sheets**.
If available, you can also use the Debug menu to fix translation mismatches automatically.

## Best practices

- Lint early and often.
- Fix errors in priority order:
  1. Bright red (critical)
  2. Light red
  3. Yellow
  4. Red text
  5. Orange text
- Keep names consistent and unique.
- Review auto-translations with native speakers.

## When errors are unclear

- Hover over red-text cells for notes.
- Check the browser console for errors (View → Developer → JavaScript Console).
- See [Troubleshooting](../troubleshooting/index.md).
