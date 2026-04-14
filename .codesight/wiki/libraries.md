# Libraries

> **Navigation aid.** Library inventory extracted via AST. Read the source files listed here before modifying exported functions.

**17 library files** across 3 modules

## Package (15 files)

- `package/src/lib/errors.js` — isParseError, isInvalidFileError, SchemaError, CategoryRefError, InvalidCategorySelectionError, UnsupportedFileVersionError, …
- `package/src/lib/utils.js` — isNotFoundError, parse, isNonEmptyArray, addRefToMap, typedEntries, unEscapePath, …
- `package/bin/helpers/messages-to-translations.js` — messagesToTranslations, parseMessageId
- `package/src/lib/un-m49.js` — normalizeUnM49ToIso31661Alpha2, unM49
- `package/bin/helpers/generate-category-selection.js` — generateCategorySelection
- `package/bin/helpers/lint.js` — lint
- `package/bin/helpers/migrate-defaults.js` — migrateDefaults
- `package/bin/helpers/migrate-geometry.js` — migrateGeometry
- `package/bin/helpers/read-files.js` — assertSchema
- `package/bin/helpers/validate-category-tags.js` — validateCategoryTags
- `package/src/lib/parse-svg.js` — parseSvg
- `package/src/lib/validate-bcp-47.js` — validateBcp47
- `package/src/lib/validate-references.js` — validateReferences
- `package/src/reader.js` — Reader
- `package/src/writer.js` — Writer

## FormatDetection.ts (1 files)

- `src/formatDetection.ts` — detectConfigFormat, normalizeConfig, NormalizedField, NormalizedPreset, NormalizedMetadata, NormalizedConfig, …

## Version.ts (1 files)

- `src/version.ts` — getVersionInfo, getFullVersionInfo, VERSION, COMMIT, BRANCH

---
_Back to [overview.md](./overview.md)_