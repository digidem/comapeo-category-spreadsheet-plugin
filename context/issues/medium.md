# Medium Priority Issues

**50+ Medium Priority Issues** - Improvements that enhance security, UX, and code quality but are not immediately critical.

---

## Security Enhancements

### MED-001: Add API Authentication
**File**: `src/apiService.ts` | **Effort**: 3-4h | **Depends**: HTTPS (CRITICAL-001)

Add authentication headers for API requests.

### MED-002: Add Content Security Policy
**File**: `src/dialog.ts` | **Effort**: 2h

Implement CSP headers for dialogs.

### MED-003: Sanitize Button Functions
**File**: `src/dialog.ts` | **Effort**: 2h

Use safer callback mechanism instead of direct code injection.

### MED-004: Add SVG Content Validation
**File**: `src/generateIcons/iconProcessor.ts` | **Effort**: 1-2h

Validate SVG content before accepting icons.

---

## UX & Validation

### MED-005: Add A1 Cell Validation
**File**: `src/spreadsheetData.ts` | **Effort**: 2h

- Add data validation dropdown to A1 cell
- Show warning if invalid value detected
- Auto-correct to English with notification

### MED-006: Add ConfigData Schema Validation
**File**: `src/importCategory/applyConfiguration.ts` | **Effort**: 2-3h

Validate configuration structure before import.

### MED-007: Add User Cancellation Handling
**File**: `src/generateCoMapeoConfig.ts` | **Effort**: 2h

Allow users to interrupt long operations.

---

## Performance Optimizations

### MED-008: Add Translation API Batching
**File**: `src/translation.ts` | **Effort**: 3-4h

Batch translation API calls (reduce by 80-90%).

### MED-009: Implement Rate Limiting
**File**: `src/translation.ts`, `iconApi.ts` | **Effort**: 2-3h

Prevent API quota exhaustion.

### MED-010: Optimize validateExtractedFiles
**File**: `src/importCategory/fileExtractor.ts` | **Effort**: 2h

Reduce to single-pass operation.

### MED-011: Add Performance Safeguards for Large Sheets
**File**: `src/lint/` | **Effort**: 2h

Handle sheets with >10K rows gracefully.

### MED-012: Cache Header Row Reads
**File**: `src/translation.ts` | **Effort**: 1h

Avoid repeated header row access.

---

## Code Quality

### MED-013: Extract Magic Array Indices
**Files**: `processFields.ts`, `processPresets.ts` | **Effort**: 1h

Replace `detail[0]`, `category[2]` with named constants.

### MED-014: Add Error Handling to Processing Modules
**Files**: `processFields.ts`, `processPresets.ts` | **Effort**: 2-3h

Add try-catch blocks and validation.

### MED-015: Extract Column Finding Logic
**File**: `src/translation.ts` | **Effort**: 1-2h

Create shared helper function (DRY).

### MED-016: Split Large Functions
**Files**: `lint.ts` (1164 lines), `parseFiles.ts` (440 lines) | **Effort**: 4-6h

Break into smaller, focused modules.

### MED-017: Extract CSS to Separate Function
**File**: `src/dialog.ts` | **Effort**: 2h

Separate styling from HTML generation.

### MED-018: Add Validation for Missing Types
**File**: `src/types.ts` | **Effort**: 2-3h

Add extraction, import, validation, and progress callback types.

---

## Testing & Reliability

### MED-019: Add Local Test Fixtures
**File**: `src/test/` | **Effort**: 2-3h

Eliminate external URL dependency for tests.

### MED-020: Add Negative Test Cases
**File**: `src/test/` | **Effort**: 3-4h

Test malformed files, edge cases, error conditions.

### MED-021: Add Unit Tests
**File**: `src/test/` | **Effort**: 6-8h

Create unit tests for pure functions (utils, processors).

### MED-022: Add Depth Limit for TAR Extraction
**File**: `src/importCategory/fileExtractor.ts` | **Effort**: 1h

Prevent excessive nesting (max 10 levels).

### MED-023: Add JSON Schema Validation
**File**: `src/importCategory/parseFiles.ts` | **Effort**: 2-3h

Validate parsed config structure.

### MED-024: Add Presets Parameter Validation
**File**: `src/importCategory/extractPngIcons.ts` | **Effort**: 30min

Prevent crashes from missing parameters.

---

## Documentation

### MED-025: Fix Spanish Typo
**File**: `src/text/dialog.ts:4` | **Effort**: 5min

Change "lenguages" → "lenguajes".

### MED-026: Add Fallback for Missing Locale
**Files**: `text/dialog.ts`, `text/menu.ts` | **Effort**: 30min

Default to 'en' if locale not found.

### MED-027: Add More Languages
**Files**: `text/dialog.ts`, `text/menu.ts` | **Effort**: 4-8h per language

Add Portuguese, French, German support.

### MED-028: Document Expected Test Results
**File**: `src/test/` | **Effort**: 1-2h

Document what each test should verify.

---

## Configuration & Deployment

### MED-029: Add Timeout Configuration
**File**: `src/preflightValidation.ts` | **Effort**: 1h

Make timeout configurable for API/network checks.

### MED-030: Add Minimum Drive Space Check
**File**: `src/preflightValidation.ts` | **Effort**: 1h

Warn if <100MB available.

### MED-031: Extract API URL to Shared Constant
**Files**: `apiService.ts`, `preflightValidation.ts` | **Effort**: 30min

Share API URL constant between modules.

---

## See Full Details

For comprehensive details on each issue including:
- Problem descriptions
- Code examples
- Testing requirements
- Implementation steps

Refer to:
- [Module Reviews](../reviews/) - Detailed analysis of each module
- [Implementation Plans](../implementation/) - Sprint-by-sprint fixes
- [Regression Strategy](../regression-strategy.md) - Safety protocols

---

## Summary

**Total Medium Priority Issues**: 50+
**Estimated Total Effort**: 60-90 hours (3-4 weeks)

**Categories**:
- Security: 10 issues
- UX/Validation: 8 issues
- Performance: 12 issues
- Code Quality: 15 issues
- Testing: 10 issues
- Documentation: 8 issues

**Implementation Timeline**: Sprints 2-4

---

**See also**:
- [Critical Issues](critical.md)
- [High Issues](high.md)
- [Low Issues](low.md)
