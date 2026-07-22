# Category Generation Process (v2)

**Status**: Current (v2 JSON build flow)

This document describes the current category generation pipeline. For user-facing steps, see `docs/how-to/export-and-share.md`.

## Overview

The v2 generator builds a JSON payload from spreadsheet data and sends it to the API (`/v2`). The API returns a packaged `.comapeocat` ZIP, which is saved to Google Drive.

## API Usage (v2 Build Endpoint)

```
POST http://your-server:3000/v2
Content-Type: application/json
```

### Request Body (Example)

```json
{
  "metadata": {
    "name": "Forest Monitoring",
    "version": "1.0.0",
    "description": "Configuration for forest surveys",
    "builderName": "comapeo-config-spreadsheet-plugin",
    "builderVersion": "2.0.0"
  },
  "locales": ["en"],
  "categories": [
    {
      "id": "trees",
      "name": "Trees",
      "color": "#4CAF50",
      "defaultFieldIds": ["species", "diameter"]
    }
  ],
  "fields": [
    {
      "id": "species",
      "name": "Species",
      "type": "select",
      "options": [
        { "value": "oak", "label": "Oak" },
        { "value": "pine", "label": "Pine" }
      ]
    },
    {
      "id": "diameter",
      "name": "Diameter (cm)",
      "type": "number"
    }
  ],
  "icons": [
    {
      "id": "trees",
      "svgUrl": "https://example.com/tree-icon.svg"
    }
  ],
  "translations": {
    "es": {
      "categories": {
        "trees": { "name": "Arboles" }
      },
      "fields": {
        "species": {
          "name": "Especie",
          "options": { "oak": "Roble", "pine": "Pino" }
        }
      }
    }
  }
}
```

### Response

- **Success (200)**: Binary `.comapeocat` file
  - `Content-Type: application/octet-stream`
  - `Content-Disposition: attachment; filename="name.comapeocat"`
- **Error**: JSON response
  ```json
  {
    "error": "VALIDATION_ERROR",
    "message": "Invalid configuration",
    "details": { "errors": ["Missing required field: name"] }
  }
  ```

## Entry Points

- **Generate CoMapeo Category**: `generateCoMapeoCategory()` (menu)
- **Debug: Export Raw Files**: `generateCoMapeoCategoryDebug()` (menu)
  - Runs the standard generator in v2; raw export is deprecated (see `docs/how-to/export-and-share.md`)
- **Manage Languages & Translate**: `translateCoMapeoCategory()` (separate translation workflow)

## Pipeline (v2)

1. **Pre-checks + Lint**
   - Detect translation mismatches and prompt to fix.
   - Run `lintAllSheets(false)` to validate data.
2. **Read Spreadsheet Data**
   - `getSpreadsheetData()` builds a `SheetData` object from Categories/Details/Translations.
3. **Build JSON Payload**
   - `createBuildPayload()` in `src/builders/payloadBuilder.ts`.
4. **API Build**
   - `sendBuildRequest()` in `src/apiService.ts` sends JSON to `/v2`.
   - API returns `.comapeocat` ZIP as a binary response.
5. **Save & Notify**
   - Save the `.comapeocat` to Drive (also saves a zipped copy for sharing).
   - Show success dialog with a Drive link.

## Category Ordering

Categories are processed in the exact order they appear in the spreadsheet. The plugin calls `setCategorySelection([...])` with an array of category IDs in spreadsheet order to preserve this ordering.

## Payload Shape (Summary)

The JSON payload includes:
- `metadata`
- `categories`
- `fields`
- `icons`
- `translations`

For schema details, see `src/builders/payloadBuilder.ts` and `src/types.ts`.

## Error Handling

- **Translation mismatches**: Detected before linting; users can auto-fix or abort.
- **Lint failures**: Highlighted in sheets; generation stops on critical issues.
- **API errors**: Retried with exponential backoff (see `RETRY_CONFIG` in `src/config.ts`).

## Related Files

- `src/generateCoMapeoConfig.ts` (orchestrator)
- `src/builders/payloadBuilder.ts` (payload construction)
- `src/apiService.ts` (API request/response handling)
- `src/lint/` (spreadsheet validation modules)
- `src/translation.ts` (language management)
- `src/spreadsheetData.ts` (data extraction)
- `context/reference/comapeocat-format.md` (archive structure)
