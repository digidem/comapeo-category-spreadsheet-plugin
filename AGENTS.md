# CoMapeo Category Set Spreadsheet Plugin - AI Agent Guidelines

**Mission:** Maintain and enhance the Google Apps Script plugin that generates `.comapeocat` configuration files from Google Spreadsheets.

## ⚠️ Core Directives (CRITICAL)

1.  **Environment Constraints (Google Apps Script):**
    *   **NO ES6 Modules at Runtime:** The runtime environment does not support `import` or `export`. All files share a global scope.
    *   **Syntax:** Use `export` keywords *only* for local TypeScript tooling/type-checking. Do **not** use `import { x } from './y'` in source files. Reference global variables directly.
    *   **Guards:** Use `typeof` checks for optional global references (e.g., `if (typeof VERSION !== 'undefined')`).
    *   **APIs:** Use `SpreadsheetApp`, `DriveApp`, `LanguageApp`, `UrlFetchApp`, `PropertiesService`, `CacheService`.
    *   **Logger load order:** Never call `getScopedLogger()` at module scope. GAS concatenates files alphabetically; `loggingHelpers.ts` is position 45, after all `src/` files. Use inline `getScopedLogger("Scope").method()` at each call site — it caches internally so repeated calls are cheap.

2.  **State Management:**
    *   Global variables do **not** persist between `google.script.run` calls.
    *   Use `PropertiesService` or pass state as arguments for cross-execution persistence.

3.  **Logging & Error Handling:**
    *   **Use:** `getScopedLogger()` (from `src/loggingHelpers.ts`) instead of `Logger.log`.
    *   **Errors:** Wrap major operations in `try-catch` blocks and use `showErrorDialog` or `showIconErrorDialog` for user feedback.

4.  **Icons:**
    *   **Format:** We use **individual PNGs** or **SVGs** for icons (stored in Drive).
    *   **Sprites:** We do **NOT** use sprites for the Apps Script implementation due to parsing limitations (see `context/reference/png-sprite-limitations.md`).

## 🏗️ Architecture & Data Flow

### 1. Data Pipeline
1.  **Extraction:** `src/spreadsheetData.ts` reads Sheets (Categories, Details, Translations) -> `SheetData` object.
2.  **Processing:** `src/generateCoMapeoConfig.ts` orchestrates processing modules in `src/generateConfig/`:
    *   `processFields.ts`: Field definitions -> `CoMapeoField[]`
    *   `processPresets.ts`: Categories -> `CoMapeoPreset[]`
    *   `processMetadata.ts`: Metadata & `package.json`
    *   `processTranslations.ts`: Aggregates all translation sheets
3.  **Export:** `src/driveService.ts` saves JSONs to Drive -> `src/apiService.ts` sends ZIP to external API for packaging -> Returns `.comapeocat`.

### 2. Import System (Reverse Engineering)
*   **Entry:** `src/importCategory.ts`
*   **Flow:** Upload `.comapeocat` -> Extract TAR/ZIP -> Parse JSONs -> Extract Icons -> Populate Sheets.
*   **Critical:** Icons must be saved to a *permanent* Drive folder, not the temp extraction folder.

### 3. Language System (Dual-Name Support)
*   **Feature:** Users can enter "Portuguese" OR "Português".
*   **Modules:** `src/languageLookup.ts`, `src/types.ts` (`LanguageData`), `src/validation.ts`.
*   **Constraint:** Maintain O(1) lookup performance using Map-based indexes.

## 🛠️ Development Workflow

### Commands
| Command | Description |
| :--- | :--- |
| `npm run dev` | Watch mode - auto-push to Apps Script (`clasp push --watch`) |
| `npm run push` | Manual push (`clasp push`). Runs version update first. |
| `bunx biome lint --write --unsafe .` | Format and lint code. **Run before pushing.** |
| `bun run scripts/generate-fixture.ts` | Regenerate fixture data for manual checks. |

### Testing Deployment
When asked to "push for testing" or "deploy for testing", the target is the **"CoMapeo Category Generator"** Google Apps Script project. Use `npm run push` (runs linting + version bump + `clasp push`) to deploy there.

### Code Style
*   **Formatting:** 2-space indent, semicolons, double quotes (per Biome config).
*   **Naming:** `verbNoun` for functions, `PascalCase` for types/interfaces, `camelCase` for variables.
*   **Comments:** Use JSDoc for complex logic. Focus on *why*, not *what*.

## 🧪 Testing Strategy
*   **Automated:** Lightweight tests in `src/test/`. Run via `testRunner.ts`.
*   **Manual (Primary):**
    1.  **Round-Trip:** Generate Config -> Download -> Import Config -> Verify Sheets match.
    2.  **Linting:** Use the "Lint Sheets" menu to verify data integrity.
    3.  **HTML Validation:** Dialogs are validated via `src/generateIcons/svgValidator.ts` and `src/formatDetection.ts` helpers to prevent malformed HTML errors.

## 📂 Documentation Index
*   **`context/reference/architecture.md`**: Full system topology.
*   **`context/reference/cat-generation.md`**: Export pipeline deep dive.
*   **`context/reference/import-cat.md`**: Import process deep dive.
*   **`docs/reference/linting-rules.md`**: Validation rules reference.
