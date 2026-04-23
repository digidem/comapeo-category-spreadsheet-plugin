# Dependency Graph

## Most Imported Files (change these carefully)

- `package/src/schema/category.js` — imported by **12** files
- `package/src/lib/errors.js` — imported by **11** files
- `package/src/schema/categorySelection.js` — imported by **11** files
- `package/src/schema/metadata.js` — imported by **10** files
- `package/src/schema/field.js` — imported by **8** files
- `package/src/lib/utils.js` — imported by **7** files
- `package/src/lib/validate-bcp-47.js` — imported by **5** files
- `package/src/schema/messages.js` — imported by **5** files
- `package/src/schema/translations.js` — imported by **5** files
- `package/bin/helpers/read-files.js` — imported by **3** files
- `package/src/reader.js` — imported by **3** files
- `package/src/lib/validate-references.js` — imported by **3** files
- `package/src/schema/defaults.js` — imported by **3** files
- `package/src/lib/parse-svg.js` — imported by **3** files
- `package/src/lib/constants.js` — imported by **3** files
- `package/src/writer.js` — imported by **2** files
- `package/bin/helpers/lint.js` — imported by **2** files
- `package/bin/helpers/messages-to-translations.js` — imported by **2** files
- `package/bin/helpers/migrate-defaults.js` — imported by **2** files
- `package/bin/helpers/migrate-geometry.js` — imported by **2** files

## Import Map (who imports what)

- `package/src/schema/category.js` ← `package/bin/comapeocat-build.mjs`, `package/bin/helpers/generate-category-selection.js`, `package/bin/helpers/lint.js`, `package/bin/helpers/lint.js`, `package/bin/helpers/migrate-geometry.js` +7 more
- `package/src/lib/errors.js` ← `package/bin/comapeocat-build.mjs`, `package/bin/comapeocat-lint.mjs`, `package/bin/comapeocat-validate.mjs`, `package/bin/helpers/lint.js`, `package/bin/helpers/read-files.js` +6 more
- `package/src/schema/categorySelection.js` ← `package/bin/comapeocat-build.mjs`, `package/bin/helpers/generate-category-selection.js`, `package/bin/helpers/lint.js`, `package/bin/helpers/migrate-defaults.js`, `package/bin/helpers/read-files.js` +6 more
- `package/src/schema/metadata.js` ← `package/bin/comapeocat-build.mjs`, `package/bin/comapeocat-build.mjs`, `package/bin/helpers/lint.js`, `package/bin/helpers/read-files.js`, `package/bin/helpers/read-files.js` +5 more
- `package/src/schema/field.js` ← `package/bin/helpers/lint.js`, `package/bin/helpers/read-files.js`, `package/bin/helpers/read-files.js`, `package/src/reader.js`, `package/src/reader.js` +3 more
- `package/src/lib/utils.js` ← `package/bin/helpers/generate-category-selection.js`, `package/bin/helpers/json-files.js`, `package/bin/helpers/lint.js`, `package/bin/helpers/messages-to-translations.js`, `package/bin/helpers/read-files.js` +2 more
- `package/src/lib/validate-bcp-47.js` ← `package/bin/comapeocat-build.mjs`, `package/bin/helpers/lint.js`, `package/bin/helpers/read-files.js`, `package/src/reader.js`, `package/src/writer.js`
- `package/src/schema/messages.js` ← `package/bin/comapeocat-messages.mjs`, `package/bin/helpers/lint.js`, `package/bin/helpers/messages-to-translations.js`, `package/bin/helpers/read-files.js`, `package/bin/helpers/read-files.js`
- `package/src/schema/translations.js` ← `package/bin/helpers/messages-to-translations.js`, `package/src/reader.js`, `package/src/reader.js`, `package/src/writer.js`, `package/src/writer.js`
- `package/bin/helpers/read-files.js` ← `package/bin/comapeocat-build.mjs`, `package/bin/comapeocat-messages.mjs`, `package/bin/helpers/lint.js`
