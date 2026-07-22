# Low Priority Issues

**20+ Low Priority Issues** - Polish, minor improvements, and nice-to-have enhancements.

---

## Code Polish

### LOW-001: Extract Magic Numbers to Constants
**Files**: Multiple | **Effort**: 1h

```typescript
// src/apiService.ts
const MIN_VALID_FILE_SIZE = 10 * 1024; // 10KB
const MAX_RETRIES = 3;
const BASE_RETRY_DELAY_MS = 1000;

// src/driveService.ts
const DRIVE_SYNC_DELAY_MS = 2000;
const PROGRESS_LOG_INTERVAL = 10; // files

// src/importCategory/applyConfiguration.ts
const SHEET_READY_DELAY_MS = 1000;
const PROGRESS_METADATA = 75;
const PROGRESS_CATEGORIES = 80;
```

### LOW-002: Remove Debug Console.log
**File**: `src/importCategory/applyConfiguration.ts:61` | **Effort**: 5min

Remove: `console.log(configData);`

### LOW-003: Remove Debug Console.log from processPresets
**File**: `src/generateConfig/processPresets.ts:9` | **Effort**: 5min

Remove debug console.log statement.

### LOW-004: Extract Color Constants
**File**: `src/lint-colors.ts` | **Effort**: 30min

```typescript
const COLORS = {
  WARNING: '#FFF3CD',
  ERROR: '#F8D7DA',
  INFO: '#D1ECF1',
  SUCCESS: '#D4EDDA'
};
```

---

## Performance Tweaks

### LOW-005: Replace Sleep with Polling
**Files**: `driveService.ts`, `applyConfiguration.ts` | **Effort**: 2h

```typescript
function waitForSheetReady(sheet, maxWaitMs = 5000): boolean {
  const startTime = Date.now();
  while (Date.now() - startTime < maxWaitMs) {
    try {
      sheet.getLastRow();
      return true;
    } catch (error) {
      Utilities.sleep(100);
    }
  }
  return false;
}
```

### LOW-006: Extract Retry Logic to Helper
**File**: `src/driveService.ts` | **Effort**: 1h

Create reusable retry function.

### LOW-007: Add Depth Limit to Drive Recursive Cleanup
**File**: `src/driveService.ts` | **Effort**: 30min

Prevent stack overflow on deep hierarchies (max 50 levels).

### LOW-008: Cache Nested Folder Lookups
**File**: `src/importCategory/parseIconSprite.ts` | **Effort**: 1h

Reduce Drive API calls.

### LOW-009: Add Performance Monitoring for Large Sprites
**File**: `src/importCategory/parseIconSprite.ts` | **Effort**: 1h

Monitor performance with >1000 symbols.

---

## Documentation

### LOW-010: Add JSDoc to Utils
**File**: `src/utils.ts` | **Effort**: 30min

Document getFieldType and getFieldOptions.

### LOW-011: Add JSDoc to Icon Modules
**Files**: `iconApi.ts`, `iconProcessor.ts` | **Effort**: 1h

Document main functions.

### LOW-012: Improve Function Naming
**File**: `src/spreadsheetData.ts` | **Effort**: 1h

Rename generic functions (sheets(), languages()).

### LOW-013: Add JSDoc to Cleanup Functions
**File**: `src/cleanup.ts` | **Effort**: 30min

Document cleanupTempResources.

### LOW-014: Document TranslationLanguage Type
**File**: `src/types.ts` | **Effort**: 15min

Add JSDoc explaining purpose.

### LOW-015: Add Usage Examples to Types
**File**: `src/types.ts` | **Effort**: 1-2h

Add JSDoc examples for complex types.

---

## Validation & Error Handling

### LOW-016: Add Null/Undefined Checks to Utils
**File**: `src/utils.ts` | **Effort**: 30min

Add defensive checks for inputs.

### LOW-017: Add Validation for Options String Format
**File**: `src/utils.ts` | **Effort**: 30min

Validate comma-separated format.

### LOW-018: Add Error Handling to Metadata Processing
**File**: `src/generateConfig/processMetadata.ts` | **Effort**: 1h

Handle sheet creation and date formatting failures.

### LOW-019: Add Format Validation for Metadata
**File**: `src/generateConfig/processMetadata.ts` | **Effort**: 1h

Validate dataset_id and version formats.

### LOW-020: Add Text Length Validation
**File**: `src/translation.ts` | **Effort**: 30min

Warn for translations >5000 chars.

### LOW-021: Add Search Term Length Validation
**File**: `src/generateIcons/iconApi.ts` | **Effort**: 15min

Validate search term length.

### LOW-022: Add PNG Content Validation
**File**: `src/importCategory/extractPngIcons.ts` | **Effort**: 1h

Check PNG magic bytes.

---

## Code Organization

### LOW-023: Split Large Files into Modules
**Files**: `lint.ts`, `parseFiles.ts`, `deconstructSvgSprite.ts` | **Effort**: 4-6h

- lint.ts (1164 lines) → categories, details, translations, icons modules
- parseFiles.ts (440 lines) → smaller focused functions
- deconstructSvgSprite.ts (435 lines) → helper functions

### LOW-024: Extract Common Test Patterns
**File**: `src/test/` | **Effort**: 2h

Create shared test helpers.

### LOW-025: Consider Namespace Organization
**All Files** | **Effort**: 8-12h

Organize global functions into namespaces.

---

## Minor Enhancements

### LOW-026: Add File Content Validation for Extraction
**File**: `src/importCategory/fileExtractor.ts` | **Effort**: 1h

Validate extracted file content beyond JSON parsing.

### LOW-027: Make CoMapeoIcon.svg Optional
**File**: `src/types.ts` | **Effort**: 15min

Allow fallback icons with optional svg property.

### LOW-028: Add Circular Reference Detection
**File**: `src/importCategory/parseFiles.ts` | **Effort**: 1h

Detect circular references in JSON parsing.

### LOW-029: Rename Property for Clarity
**Files**: `parseFiles.ts`, `parseIconSprite.ts` | **Effort**: 30min

Rename `svg` property → `url` or `iconUrl`.

### LOW-030: Add Quota Monitoring
**All Files** | **Effort**: 2-3h

Monitor Drive API calls and warn on approaching limits.

### LOW-031: Add Memory Usage Monitoring
**All Files** | **Effort**: 2-3h

Track memory usage during operations.

---

## Testing

### LOW-032: Add Test for Validation Modules
**File**: `src/test/` | **Effort**: 2h

Test validation, cleanup, utilities.

### LOW-033: Consider Test Isolation Strategies
**File**: `src/test/` | **Effort**: 4-6h

Copy spreadsheet or use test sheets.

### LOW-034: Cache getLastRow() Result
**File**: `src/importCategory/utils.ts` | **Effort**: 15min

Avoid redundant calls.

---

## Summary

**Total Low Priority Issues**: 34 items
**Estimated Total Effort**: 40-60 hours (2-3 weeks)

**Categories**:
- Code Polish: 6 issues
- Performance: 5 issues
- Documentation: 6 issues
- Validation: 7 issues
- Code Organization: 3 issues
- Minor Enhancements: 7 issues

**Implementation Timeline**: Sprint 4 and ongoing maintenance

---

## Prioritization Within LOW Category

**Quick Wins** (< 1 hour each):
- LOW-002: Remove debug logs
- LOW-003: Remove debug console.log
- LOW-004: Extract color constants
- LOW-010: Add JSDoc to utils
- LOW-014: Document TranslationLanguage

**Polish Sprint** (Week 4):
- LOW-001: Extract magic numbers
- LOW-005: Replace sleep with polling
- LOW-011: Add JSDoc to icon modules
- LOW-020: Add text length validation

**Future Improvements**:
- LOW-023: Split large files
- LOW-025: Namespace organization
- LOW-030: Quota monitoring
- LOW-033: Test isolation

---

**See also**:
- [Critical Issues](critical.md)
- [High Issues](high.md)
- [Medium Issues](medium.md)
- [Implementation Sprint 4](../implementation/sprint-04-polish.md)
