# Categories sheet tasks

## Add categories

1. Open the **Categories** sheet.
2. Starting at row 2, enter category names in column A.
3. Repeat for each category.

See [Categories sheet](../reference/categories-sheet.md) for column rules.

## Set Applies values (required)

1. In column D (**Applies**), set where each category is used.
2. Ensure at least one category includes `track`.

Valid values: `observation`, `track`, or `observation, track`.
Abbreviations `o` and `t` are also accepted.

## Link fields to categories

1. In column C (**Fields**), enter a comma-separated list of field names.
2. Field names must match the **Details** sheet exactly.

Example: `Name, Width, Condition`

## Add icons

### Inline SVG (recommended)

1. Go to <https://icons.earthdefenderstoolkit.com>
2. Search for a category name (e.g., `river`, `building`).
3. Choose an icon and color.
4. Click **Copy SVG**.
5. Paste the SVG directly into the **Icon** column (B).

### Alternative: plain text lookup

1. Type a simple keyword in the **Icon** column (e.g., `river`, `forest`).
2. During export, the plugin searches <https://icons.earthdefenderstoolkit.com>.

Faster, but can return less accurate matches.

### Generate icons (menu)

1. Open **CoMapeo Tools → Generate Category Icons**.
2. Choose from previewed icons.
3. Icons are saved to Drive and the **Icon** column is updated with Drive URLs.

Tips:

- Search works best with common English terms (e.g., `river`, `building`).
- Requires an internet connection for the icon API.
- You can edit icons afterward at <https://icons.earthdefenderstoolkit.com>.

### Supported formats

- Inline SVG (recommended)
- Plain text keywords (auto lookup)
- Drive URLs (`https://drive.google.com/file/d/FILE_ID/view`)
- Cell images (paste into the cell)
- Data URIs (`data:image/svg+xml,...`)
- HTTPS URLs
- SVG and PNG files (PNG is converted to SVG)

### Troubleshooting

- If icons fail, run linting and fix any red text in the Icon column.
- If using Drive URLs, confirm the file is shared with “View” access.

## Set category colors

1. Select the category name cells in column A.
2. Use the **Fill color** tool to assign a color per category.

Notes:

- Category colors are based on the background color of column A.
- There is no separate Color column.
- Do not paste hex codes into the **Applies** column.
- Choose distinct, colorblind-friendly colors when possible.
- Bright, high-contrast colors tend to display best.
