# Troubleshooting

## Generation fails

**Common error:** "At least one category must include 'track' in Applies"

- Ensure column D (**Applies**) exists.
- Set at least one category to `track` or `observation, track`.
- Do not put hex color codes in column D.

Other checks:

- Run linting and fix all bright red and red errors.
- Verify icon references are accessible.
- Ensure Fields references exist in **Details**.
- Confirm translation sheets have no bright red cells.
- Check your internet connection (required for the packaging server).

## Icons not working

- Best results: paste inline SVG from <https://icons.earthdefenderstoolkit.com>.
- Plain text keywords (e.g., `river`) are accepted but may be less accurate.
- For Drive URLs, confirm the format and permissions:
  - `https://drive.google.com/file/d/FILE_ID/view`
  - File must be shared with **View** access.
- Run linting to identify invalid icon references.
- Orange text on HTTP URLs is a security warning; icons can still work.

## Translations missing

- Verify languages were selected during generation.
- Confirm translation sheets exist and have data.
- Run linting and fix bright red translation mismatches.
- Re-run **Manage Languages & Translate** to fill blanks.
- Ensure column A formulas are intact.

## Import issues

- Supported files: `.comapeocat` and `.mapeosettings`.
- Large files may take 1-2 minutes.
- Check available Drive storage for icon uploads.
- Open the browser console for errors (View → Developer → JavaScript Console).
- Try refreshing the spreadsheet and importing again.

## Language not recognized

- Use full language names in **Categories!A1** (not codes).
- Examples:
  - `Portuguese` or `Português`
  - `Spanish` or `Español`
  - `French` or `Français`
- Codes like `pt`, `es`, `fr` are not valid in A1.

## Applies column issues

- Valid values: `observation`, `track`, `observation, track`.
- Abbreviations `o`, `t` are accepted.
- At least one category must include `track`.
- If the column is missing, the plugin will auto-create it.

## Slow performance

- Reduce the number of categories or languages.
- Try generating with fewer languages first.
- Close other browser tabs.
- Clear browser cache or reopen the sheet.
- Try again during off-peak hours.

## Linting errors you don't understand

1. Hover over red text cells for detailed notes.
2. Review [Linting rules](../reference/linting-rules.md).
3. Check the browser console for specific error messages.
4. Fix one issue at a time and re-run linting.
