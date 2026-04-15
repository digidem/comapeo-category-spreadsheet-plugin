# Translations

## Add languages

1. Open **CoMapeo Tools → Manage Languages & Translate**.
2. In Section 1, check languages to auto-translate.
3. In Section 2, add manual-only language codes (e.g., `gn`, `quz`, `pt-BR`).
4. Click **Translate**.

Creates translation sheets and fills empty cells for auto-translation.

## Update translations

- Edit any translation cell manually.
- Re-run **Manage Languages & Translate** to fill remaining blank cells.
- Existing translations are not overwritten.

## What gets translated

- Category names
- Field labels
- Helper text
- Field options

If a language is not recognized, check spelling or try the native name.

## Language header formats

Translation sheet headers can be:

- Language name: `Spanish`, `Español`, `Portuguese`, `Português`
- ISO or BCP-47 code: `es`, `pt`, `pt-BR`, `zh-Hans`
- Name + code: `Spanish - es`, `Português - pt`

See [Translation sheets](../reference/translation-sheets.md) for the full format rules.

## Best practices

- Review auto-translations with native speakers.
- Use manual-only columns for languages not supported by Google Translate.

## Find language codes

- SIL ISO 639-3: <https://iso639-3.sil.org/>
- Ethnologue: <https://www.ethnologue.com/>
- Glottolog: <https://glottolog.org/>
