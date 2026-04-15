# comapeo-config-spreadsheet-plugin — AI Context Map

> **Stack:** raw-http | none | unknown | typescript

> 0 routes | 0 models | 0 components | 17 lib files | 0 env vars | 6 middleware | 0% test coverage
> **Token savings:** this file is ~1.900 tokens. Without it, AI exploration would cost ~15.100 tokens. **Saves ~13.200 tokens per conversation.**
> **Last scanned:** 2026-04-15 10:25 — re-run after significant changes

---

# Libraries

- `package/bin/helpers/generate-category-selection.js` — function generateCategorySelection: (categoriesMap) => void
- `package/bin/helpers/lint.js` — function lint: (dir) => void
- `package/bin/helpers/messages-to-translations.js` — function messagesToTranslations: (messages) => void, function parseMessageId: (messageId) => void
- `package/bin/helpers/migrate-defaults.js` — function migrateDefaults: (defaults) => void
- `package/bin/helpers/migrate-geometry.js` — function migrateGeometry: (category) => void
- `package/bin/helpers/read-files.js` — function assertSchema: (schema, data, {...}) => void
- `package/bin/helpers/validate-category-tags.js` — function validateCategoryTags: (categoriesMap) => void
- `package/src/lib/errors.js`
  - function isParseError: (err) => void
  - function isInvalidFileError: (err) => void
  - class SchemaError
  - class CategoryRefError
  - class InvalidCategorySelectionError
  - class UnsupportedFileVersionError
  - _...15 more_
- `package/src/lib/parse-svg.js` — function parseSvg: (svg) => void
- `package/src/lib/un-m49.js` — function normalizeUnM49ToIso31661Alpha2: (code) => void, const unM49
- `package/src/lib/utils.js`
  - function isNotFoundError: (err) => void
  - function parse: (schema, data, {...}) => void
  - function isNonEmptyArray: (value) => void
  - function addRefToMap: (map, refId, categoryId) => void
  - function typedEntries: (obj) => void
  - function unEscapePath: (path) => void
  - _...1 more_
- `package/src/lib/validate-bcp-47.js` — function validateBcp47: (tag) => void
- `package/src/lib/validate-references.js` — function validateReferences: ({...}, fieldIds, iconIds, categorySelection, }) => void
- `package/src/reader.js` — class Reader
- `package/src/writer.js` — class Writer
- `src/formatDetection.ts`
  - function detectConfigFormat: (configData) => ConfigFormat
  - function normalizeConfig: (configData) => NormalizedConfig
  - interface NormalizedField
  - interface NormalizedPreset
  - interface NormalizedMetadata
  - interface NormalizedConfig
  - _...1 more_
- `src/version.ts`
  - function getVersionInfo: () => string
  - function getFullVersionInfo: () => string
  - const VERSION
  - const COMMIT
  - const BRANCH

---

# Config

## Config Files

- `tsconfig.json`

---

# Middleware

## custom
- regression-strategy — `context/process/regression-strategy.md`
- generate-category-selection — `package/bin/helpers/generate-category-selection.js`
- migrate-defaults — `package/bin/helpers/migrate-defaults.js`
- migrate-geometry — `package/bin/helpers/migrate-geometry.js`
- generate-fixture — `scripts/generate-fixture.ts`

## logging
- generateCoMapeoConfig — `src/generateCoMapeoConfig.ts`

---

# Dependency Graph

## Most Imported Files (change these carefully)

- `package/src/schema/category.js` — imported by **12** files
- `package/src/schema/categorySelection.js` — imported by **11** files
- `package/src/lib/errors.js` — imported by **10** files
- `package/src/schema/metadata.js` — imported by **10** files
- `package/src/schema/field.js` — imported by **8** files
- `package/src/lib/utils.js` — imported by **7** files
- `package/src/schema/messages.js` — imported by **5** files
- `package/src/schema/translations.js` — imported by **5** files
- `package/bin/helpers/read-files.js` — imported by **3** files
- `package/src/reader.js` — imported by **3** files
- `package/src/lib/validate-references.js` — imported by **3** files
- `package/src/schema/defaults.js` — imported by **3** files
- `package/src/lib/constants.js` — imported by **3** files
- `package/src/lib/validate-bcp-47.js` — imported by **2** files
- `package/src/writer.js` — imported by **2** files
- `package/bin/helpers/lint.js` — imported by **2** files
- `package/bin/helpers/messages-to-translations.js` — imported by **2** files
- `package/bin/helpers/migrate-defaults.js` — imported by **2** files
- `package/bin/helpers/migrate-geometry.js` — imported by **2** files
- `package/src/lib/parse-svg.js` — imported by **2** files

## Import Map (who imports what)

- `package/src/schema/category.js` ← `package/bin/comapeocat-build.mjs`, `package/bin/helpers/generate-category-selection.js`, `package/bin/helpers/lint.js`, `package/bin/helpers/lint.js`, `package/bin/helpers/migrate-geometry.js` +7 more
- `package/src/schema/categorySelection.js` ← `package/bin/comapeocat-build.mjs`, `package/bin/helpers/generate-category-selection.js`, `package/bin/helpers/lint.js`, `package/bin/helpers/migrate-defaults.js`, `package/bin/helpers/read-files.js` +6 more
- `package/src/lib/errors.js` ← `package/bin/comapeocat-build.mjs`, `package/bin/comapeocat-lint.mjs`, `package/bin/comapeocat-validate.mjs`, `package/bin/helpers/lint.js`, `package/bin/helpers/read-files.js` +5 more
- `package/src/schema/metadata.js` ← `package/bin/comapeocat-build.mjs`, `package/bin/comapeocat-build.mjs`, `package/bin/helpers/lint.js`, `package/bin/helpers/read-files.js`, `package/bin/helpers/read-files.js` +5 more
- `package/src/schema/field.js` ← `package/bin/helpers/lint.js`, `package/bin/helpers/read-files.js`, `package/bin/helpers/read-files.js`, `package/src/reader.js`, `package/src/reader.js` +3 more
- `package/src/lib/utils.js` ← `package/bin/helpers/generate-category-selection.js`, `package/bin/helpers/json-files.js`, `package/bin/helpers/lint.js`, `package/bin/helpers/messages-to-translations.js`, `package/bin/helpers/read-files.js` +2 more
- `package/src/schema/messages.js` ← `package/bin/comapeocat-messages.mjs`, `package/bin/helpers/lint.js`, `package/bin/helpers/messages-to-translations.js`, `package/bin/helpers/read-files.js`, `package/bin/helpers/read-files.js`
- `package/src/schema/translations.js` ← `package/bin/helpers/messages-to-translations.js`, `package/src/reader.js`, `package/src/reader.js`, `package/src/writer.js`, `package/src/writer.js`
- `package/bin/helpers/read-files.js` ← `package/bin/comapeocat-build.mjs`, `package/bin/comapeocat-messages.mjs`, `package/bin/helpers/lint.js`
- `package/src/reader.js` ← `package/bin/comapeocat-validate.mjs`, `package/src/index.js`, `package/src/lib/utils.js`

---

# Test Coverage

> **0%** of routes and models are covered by tests
> 23 test files found

---

_Generated by [codesight](https://github.com/Houseofmvps/codesight) — see your codebase clearly_