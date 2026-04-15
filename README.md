# CoMapeo Category Set Spreadsheet Plugin

A Google Sheets plugin for generating and importing `.comapeocat` category files for CoMapeo projects.

Current version: **2.0.0**

Quick links: [User Docs](docs/index.md) | [Internal Docs](context/README.md) | [Linting Guide](docs/linting-guide.md) | [API Reference](context/reference/cat-generation.md)

## Table of Contents

- [What's New in v2.0.0](#whats-new-in-v200)
- [Features](#features)
- [Getting Started (Users)](#getting-started-users)
- [Installation (Developers)](#installation-developers)
- [Development](#development)
- [Testing](#testing)
- [Documentation Map](#documentation-map)
- [Contributing & Support](#contributing--support)
- [License](#license)
- [References](#references)

## What's New in v2.0.0

- **JSON build flow**: The plugin sends a JSON build request to the v2 API, which returns a packaged `.comapeocat` (ZIP) file.
- **Import functionality**: Import existing `.comapeocat` files back into the spreadsheet for editing.
- **Category ordering**: Categories are built in exact spreadsheet order using `setCategorySelection`.
- **Simplified packaging**: Local ZIP creation is no longer part of the generation flow; packaging happens on the API side.

## Features

- **Build**: Generate `.comapeocat` files from spreadsheet data.
- **Import**: Load existing `.comapeocat` or `.mapeosettings` files for editing.
- **Auto-translation**: Automatically translate categories and fields using Google Translate.
- **Icon generation**: Generate icons using <https://icons.earthdefenderstoolkit.com> or provide your own SVG/PNG icons.
- **Validation**: Comprehensive linting and validation of spreadsheet data.
- **Dual-name language support**: Set primary language using English OR native names (e.g., "Portuguese" or "Português").

## Getting Started (Users)

The [user docs](docs/index.md) are the source of truth for spreadsheets, menu behavior, and end-to-end workflows. In short:

1. Fill out the spreadsheet tabs (Categories, Details, Translations).
2. Use the menu option **Generate CoMapeo Category** to build a `.comapeocat`.
3. Use **Import Category File** to bring an existing config back into the spreadsheet.

Menu options, spreadsheet structure, and sharing steps are documented in [docs/index.md](docs/index.md).

## Installation (Developers)

1. Install clasp globally:
```bash
npm install -g @google/clasp
```

2. Login to your Google account:
```bash
clasp login
```

3. Enable the Google Apps Script API in your Google Cloud Console.

4. Push your local changes:
```bash
clasp push
```

5. Open in the Apps Script editor:
```bash
clasp open
```

## Development

```bash
# Watch mode - auto-push on changes
npm run dev

# Manual push
npm run push

# Lint code (requires Bun)
npm run lint

# Push to all configured projects
npm run push:all

# Regenerate fixture data for manual checks
npm run fixture
```

## Testing

Run tests from the Apps Script editor console:

```javascript
runAllTests();
```

Regression testing workflows are documented in [context/process/regression-testing-guide.md](context/process/regression-testing-guide.md).

## Documentation Map

- **User docs (source of truth)**: [docs/index.md](docs/index.md)
- **Internal documentation index**: [context/README.md](context/README.md)
- **Linting rules**: [docs/reference/linting-rules.md](docs/reference/linting-rules.md)
- **Architecture & pipelines**: [context/reference/architecture.md](context/reference/architecture.md), [context/reference/cat-generation.md](context/reference/cat-generation.md), [context/reference/import-cat.md](context/reference/import-cat.md)
- **Formats & validation**: [context/reference/comapeocat-format.md](context/reference/comapeocat-format.md), [context/reference/html-validation.md](context/reference/html-validation.md), [context/reference/png-sprite-limitations.md](context/reference/png-sprite-limitations.md)
- **Process & workflows**: [context/process/](context/process/), [AGENTS.md](AGENTS.md)

## Contributing & Support

- Check the [documentation index](context/README.md) and [Linting Guide](docs/linting-guide.md) first.
- Run linting and tests before opening a pull request.
- Use the repository issue tracker for bugs and feature requests.

## License

See LICENSE file for details.

## References

For more information on using clasp, refer to the [@google/clasp documentation](https://github.com/google/clasp).
