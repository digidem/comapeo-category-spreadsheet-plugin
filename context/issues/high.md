# High Priority Issues

**25+ High Priority Issues** - Code Quality, Reliability, and Security improvements that should be addressed soon.

---

## Overview

These issues significantly impact code maintainability, system reliability, or have moderate security implications. While not as urgent as CRITICAL issues, they should be prioritized in the development roadmap.

---

## Code Quality Issues

### HIGH-001: Extract Hardcoded Language Fallback

**File**: `src/spreadsheetData.ts` (lines 17-142)
**Effort**: 1-2 hours

**Problem**: 125 lines of hardcoded language fallback data embedded in code

**Recommendation**:
- Create external JSON file (`src/data/languages-fallback.json`)
- Load from external source with fallback to local data
- Improves maintainability and reduces code size

**See**: [Module Review - Core Data Flow](../reviews/core-data-flow.md#11-spreadsheet-data)

---

### HIGH-002: Create DRY Helper Functions

**File**: `src/spreadsheetData.ts`
**Effort**: 2-3 hours
**Dependencies**: Caching implementation (CRITICAL-003)

**Problem**: Repeated filter/reduce patterns across 4+ functions

**Recommendation**:
```typescript
function getPrimaryLanguageName(): string { /* ... */ }
function filterLanguagesByPrimary(allLanguages, includePrimary): Record<string, string> { /* ... */ }
```

**See**: [Module Review - Core Data Flow](../reviews/core-data-flow.md#11-spreadsheet-data)

---

### HIGH-003: Remove Duplicate slugify() Function

**File**: `src/utils.ts` and `src/importCategory/utils.ts`
**Effort**: 30 minutes

**Problem**: slugify() duplicated in two files

**Recommendation**:
- Remove duplicate from importCategory/utils.ts
- Use global scope reference (Apps Script pattern)
- Test all import/export functionality

**Status**: Also listed as CRITICAL-006 due to importance

**See**: [Module Review - Utilities](../reviews/cross-cutting.md#utilities--helpers)

---

### HIGH-004: Remove Uncached Spreadsheet Access

**File**: `src/generateConfig/processPresets.ts`
**Effort**: 1 hour

**Problem**: Calls `getActiveSpreadsheet().getSheetByName()` inside function - not cached

**Recommendation**:
- Use cached spreadsheet reference
- Pass sheet as parameter
- Improves performance

**See**: [Module Review - Processing Modules](../reviews/processing-modules.md#22-preset-processing)

---

## Reliability Issues

### HIGH-005: Add Cleanup Logic to Config Generation

**File**: `src/generateCoMapeoConfig.ts`
**Effort**: 2-3 hours

**Problem**: When API fails, Drive folders remain orphaned

**Recommendation**:
```typescript
function generateCoMapeoConfigWithSelectedLanguages(selectedLanguages) {
  let createdFolderId: string | null = null;

  try {
    // ... existing code ...
    createdFolderId = saveConfigToDrive(config).id;
    // ... rest of pipeline ...
  } catch (error) {
    // CLEANUP: Delete created Drive folder
    if (createdFolderId) {
      cleanupDriveFolder(createdFolderId);
    }
    throw error;
  }
}
```

**See**: [Module Review - Core Data Flow](../reviews/core-data-flow.md#12-config-generation)

---

### HIGH-006: Add File Size Validation

**Files**: `src/importCategory/fileExtractor.ts`
**Effort**: 1 hour

**Problem**: No file size limits - could exhaust memory

**Recommendation**:
```typescript
const MAX_FILE_SIZE = 100 * 1024 * 1024; // 100MB

if (fileSize > MAX_FILE_SIZE) {
  throw new Error(`File too large: ${fileSize} bytes (max: ${MAX_FILE_SIZE})`);
}
```

**See**: [Module Review - Import System](../reviews/import-system.md#41-file-extraction)

---

### HIGH-007: Add Path Traversal Validation

**Files**: `src/importCategory/fileExtractor.ts`
**Effort**: 1-2 hours

**Problem**: No path traversal validation in TAR extraction

**Recommendation**:
```typescript
function validateFilePath(path: string): boolean {
  // Check for path traversal attempts
  if (path.includes('../') || path.includes('..\\')) {
    throw new Error(`Invalid file path: ${path}`);
  }
  return true;
}
```

**See**: [Module Review - Import System](../reviews/import-system.md#41-file-extraction)

---

### HIGH-008: Add Validation for Empty Icon Return

**File**: `src/generateIcons/iconProcessor.ts`
**Effort**: 1 hour

**Problem**: Complete failure to generate icon returns empty string (could cause issues)

**Recommendation**:
- Return fallback icon or null
- Add validation before using icon
- Log warnings for missing icons

**See**: [Module Review - Icon Generation](../reviews/icon-generation.md#52-icon-processor)

---

### HIGH-009: Externalize API URL

**File**: `src/apiService.ts`
**Effort**: 1 hour
**Dependencies**: HTTPS implementation (CRITICAL-001)

**Problem**: Hardcoded API URL makes configuration difficult

**Recommendation**:
```typescript
const API_URL = PropertiesService.getScriptProperties().getProperty('API_URL')
  || 'https://api.comapeo.app/process';
```

**See**: [Module Review - Core Data Flow](../reviews/core-data-flow.md#14-api-service)

---

## Security Issues

### HIGH-010: Cache/Embed CoMapeo Logo

**File**: `src/dialog.ts`
**Effort**: 1 hour

**Problem**: CoMapeo logo loaded from GitHub raw - network dependency

**Recommendation**:
- Convert to base64 data URI
- Embed in HTML template
- Eliminates external dependency and improves security

**See**: [Module Review - UI Layer](../reviews/ui-layer.md#81-dialog-management)

---

### HIGH-011: Add Depth Limit to Recursive Operations

**Files**: `src/driveService.ts`, `cleanup.ts`, `fileExtractor.ts`
**Effort**: 1 hour

**Problem**: No depth limit for recursive folder operations

**Recommendation**:
```typescript
const MAX_DEPTH = 50;

function recursiveOperation(folder, depth = 0) {
  if (depth > MAX_DEPTH) {
    throw new Error(`Maximum depth exceeded: ${MAX_DEPTH}`);
  }
  // ... operation ...
  recursiveOperation(subfolder, depth + 1);
}
```

**See**: Multiple module reviews

---

### HIGH-012: Add Input Validation

**Files**: Multiple
**Effort**: 3-4 hours

**Problem**: Missing validation for user inputs and sheet data

**Recommendation**:
- Validate A1 cell value against supported languages
- Validate field types before processing
- Validate category names not empty
- Add schema validation for configData

**See**: Multiple module reviews

---

## Testing & Documentation Issues

### HIGH-013: Add Test Documentation

**File**: `src/test/` directory
**Effort**: 2 hours

**Problem**: No documentation on how to run tests or what to verify

**Recommendation**:
- Create README in test/ folder
- Document run instructions
- Document expected results
- Add troubleshooting guide

**See**: [Module Review - Testing Infrastructure](../reviews/cross-cutting.md#testing-infrastructure)

---

### HIGH-014: Implement Test Cleanup Helpers

**File**: `src/test/` directory
**Effort**: 2-3 hours

**Problem**: Tests modify active spreadsheet without backup/restore

**Recommendation**:
- Create test setup/teardown helpers
- Backup spreadsheet state before tests
- Restore state after tests complete
- Add cleanup for Drive folders

**See**: [Module Review - Testing Infrastructure](../reviews/cross-cutting.md#testing-infrastructure)

---

### HIGH-015: Add Test Summary Report

**File**: `src/test/` directory
**Effort**: 1-2 hours

**Problem**: No pass/fail summary at end of tests

**Recommendation**:
- Collect all test results
- Display summary dialog
- Show pass/fail counts
- Highlight failures for review

**See**: [Module Review - Testing Infrastructure](../reviews/cross-cutting.md#testing-infrastructure)

---

### HIGH-016: Add JSDoc to All Public Functions

**Files**: All modules (priority: utils, processing modules)
**Effort**: 4-6 hours

**Problem**: Inconsistent documentation across codebase

**Standard format**:
```typescript
/**
 * Brief description of function
 *
 * @param paramName - Description
 * @returns Description of return value
 * @throws Error description
 * @example
 * const result = functionName(param);
 */
function functionName(paramName: string): ReturnType {
  // ...
}
```

**Status (2025-10-14)**: JSDoc coverage extended to translation, naming, and preset helpers alongside scoped logging rollout.

**See**: Multiple module reviews

---

## Performance Issues

### HIGH-017: Implement Centralized Logging

**Files**: All modules
**Effort**: 3-4 hours

**Problem**: Inconsistent logging patterns, no log levels

**Recommendation**:
```typescript
class Logger {
  static debug(message: string, ...args: any[]): void { /* ... */ }
  static info(message: string, ...args: any[]): void { /* ... */ }
  static warn(message: string, ...args: any[]): void { /* ... */ }
  static error(message: string, ...args: any[]): void { /* ... */ }
}
```

**See**: [Module Review - Cross-Cutting Concerns](../reviews/cross-cutting.md#123-logging-strategy)

---

### HIGH-018: Optimize validateTranslationSheets

**File**: `src/lint/translations.ts`
**Effort**: 2 hours

**Problem**: Reads entire data range multiple times

**Recommendation**:
- Read data once
- Process in single pass
- Cache frequently accessed ranges

**See**: [Module Review - Validation](../reviews/validation.md#61-linting)

---

### HIGH-019: Add Automatic Cleanup of Old Temp Folders

**File**: `src/cleanup.ts`
**Effort**: 2-3 hours

**Problem**: No expiration policy for temporary files

**Recommendation**:
```typescript
function cleanupOldTempFolders(olderThanHours = 24) {
  const cutoffTime = new Date().getTime() - (olderThanHours * 60 * 60 * 1000);
  // Find and delete folders older than cutoff
}
```

**See**: [Module Review - Validation](../reviews/validation.md#62-cleanup)

---

## Additional High Priority Items

### HIGH-020: Add Retry Logic to Icon API

**File**: `src/generateIcons/iconApi.ts`
**Effort**: 1 hour

**Problem**: No retry logic for API failures

**See**: [Module Review - Icon Generation](../reviews/icon-generation.md#51-icon-api)

---

### HIGH-021: Cache Icon Search Results

**File**: `src/generateIcons/iconProcessor.ts`
**Effort**: 2 hours

**Problem**: Multiple API calls per icon with no caching

**See**: [Module Review - Icon Generation](../reviews/icon-generation.md#52-icon-processor)

---

### HIGH-022: Fix Infinite Loop (See CRITICAL-004)

**File**: `src/generateIcons/iconProcessor.ts:246-249`
**Effort**: 30 minutes

Also listed as CRITICAL issue due to severity.

**See**: [Critical Issues](critical.md#critical-4-infinite-loop-in-icon-generation)

---

### HIGH-023: Create Dependency Documentation

**All Files**
**Effort**: 2-3 hours

**Problem**: Hard to trace dependencies without reading code (Apps Script limitation)

**Recommendation**:
- Create module dependency diagram
- Document which modules call which
- Note data flow patterns
- List external dependencies

**See**: [Module Review - Cross-Cutting Concerns](../reviews/cross-cutting.md#124-code-organization)

---

### HIGH-024: Standardize Naming Conventions

**All Files**
**Effort**: 4-6 hours

**Problem**: Inconsistent naming (camelCase vs snake_case in data structures)

**Recommendation**:
- Document naming conventions
- Create linting rules
- Refactor inconsistent names
- Update style guide

**See**: [Module Review - Cross-Cutting Concerns](../reviews/cross-cutting.md#124-code-organization)

**Status (2025-10-14)**: Naming helpers (`createFieldTagKey`, `createPresetSlug`, `createOptionValue`) centralised in `src/utils.ts`, applied across export/import flows, and documented in `context/process/naming-conventions.md`.

---

### HIGH-025: Add TypeScript Type Annotations

**File**: `src/utils.ts`
**Effort**: 1 hour

**Problem**: Functions use `any` type or missing annotations

**Recommendation**:
```typescript
function slugify(text: string | any): string {
  return String(text)
    .toLowerCase()
    .trim()
    // ... rest of implementation
}

function getFieldType(typeString: string): 'text' | 'number' | 'selectOne' | 'selectMultiple' {
  // ... implementation
}
```

**See**: [Module Review - Utilities](../reviews/cross-cutting.md#91-utils)

---

## Summary Table

| # | Issue | File(s) | Effort | Category |
|---|-------|---------|--------|----------|
| 001 | Extract Language Fallback | spreadsheetData.ts | 1-2h | Code Quality |
| 002 | DRY Helper Functions | spreadsheetData.ts | 2-3h | Code Quality |
| 003 | Remove Duplicate slugify | utils.ts, importCategory/utils.ts | 30min | Code Quality |
| 004 | Remove Uncached Access | processPresets.ts | 1h | Performance |
| 005 | Add Cleanup Logic | generateCoMapeoConfig.ts | 2-3h | Reliability |
| 006 | File Size Validation | import modules | 1h | Security |
| 007 | Path Traversal Validation | import modules | 1-2h | Security |
| 008 | Validate Empty Icon | iconProcessor.ts | 1h | Reliability |
| 009 | Externalize API URL | apiService.ts | 1h | Configuration |
| 010 | Embed CoMapeo Logo | dialog.ts | 1h | Security |
| 011 | Add Depth Limits | multiple | 1h | Security |
| 012 | Input Validation | multiple | 3-4h | Reliability |
| 013 | Test Documentation | test/ | 2h | Documentation |
| 014 | Test Cleanup Helpers | test/ | 2-3h | Testing |
| 015 | Test Summary Report | test/ | 1-2h | Testing |
| 016 | Add JSDoc | all modules | 4-6h | Documentation |
| 017 | Centralized Logging | all modules | 3-4h | Code Quality |
| 018 | Optimize Validation | lint.ts | 2h | Performance |
| 019 | Cleanup Old Temp Folders | cleanup.ts | 2-3h | Maintenance |
| 020 | Icon API Retry Logic | iconApi.ts | 1h | Reliability |
| 021 | Cache Icon Searches | iconProcessor.ts | 2h | Performance |
| 022 | Fix Infinite Loop | iconProcessor.ts | 30min | Reliability |
| 023 | Dependency Docs | all | 2-3h | Documentation |
| 024 | Naming Conventions | all | 4-6h | Code Quality |
| 025 | TypeScript Types | utils.ts | 1h | Code Quality |

**Total Estimated Effort**: 48-65 hours (approximately 2-3 weeks)

---

## Implementation Priority

**Sprint 2 Focus** (Week 2):
- HIGH-002: DRY Helpers
- HIGH-005: Cleanup Logic
- HIGH-009: Externalize API URL

**Sprint 3 Focus** (Week 3):
- HIGH-006: File Size Validation
- HIGH-007: Path Traversal Protection
- HIGH-012: Input Validation

**Sprint 4+ Focus** (Weeks 4+):
- HIGH-016: JSDoc Comments
- HIGH-017: Centralized Logging
- HIGH-023: Dependency Documentation

---

**See also**:
- [Critical Priority Issues](critical.md)
- [Medium Priority Issues](medium.md)
- [Implementation Plans](../implementation/)
