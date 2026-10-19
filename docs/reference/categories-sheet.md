# Categories sheet

Define the main observation types users will record in the field.

## Columns

### Name (Column A, required)

- Category name in your primary language.
- The **background color** of this cell becomes the category color in CoMapeo.
- Used as the border color in the category list and the dot color on the map.
- Examples: `River`, `Building`, `Animal Sighting`.

### Icon (Column B, required)

Accepted formats:

- Inline SVG (recommended)
- Plain text keyword (auto-lookup)
- Drive URL (`https://drive.google.com/file/d/FILE_ID/view`)
- Cell image
- Data URI
- HTTPS URL

See [Categories sheet tasks](../how-to/categories-sheet-tasks.md#add-icons).

PNG files are converted to SVG during export. PNG sprites are not fully supported.

### Fields (Column C)

Comma-separated list of field names this category uses. Field names must exist in the Details sheet.

Example: `Name, Width, Condition`

This column may be labeled **Details** in older templates.

### Applies (Column D)

Where this category is used:

- `observation`
- `track`
- `observation, track`
- Abbreviations `o`, `t` are accepted

**Requirement:** At least one category must include `track`.

If the column is missing, the plugin auto-creates it with `observation, track` for the first row.

### Category ID (Column E)

Auto-generated unique ID used during import/export. Avoid manual edits.

Changing an ID after data collection can break links to existing observations in CoMapeo.

### Color (Column F, optional)

Explicit hex color per category (for example `#1B5E88`). Written by config imports and read at export time.

If empty, the category color falls back to the background color of column A. The background remains the primary way to set colors by hand.

### Icon ID (Column G, optional)

Icon identifier used to link the category to an entry in the Icons sheet. Written by config imports; leave it blank for hand-built sheets.

## Notes

- Do not place hex colors in **Applies**. Colors are set via the Color column or the background color of column A.
- Canonical column order: `Name, Icon, Fields, Applies, Category ID, Color, Icon ID`. Config imports write exactly this layout.
