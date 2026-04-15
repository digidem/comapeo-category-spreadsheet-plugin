# Getting started

## Overview

The CoMapeo Config Spreadsheet Plugin lets you define categories, fields, icons, colors, and translations in Google Sheets, then export a `.comapeocat` file for CoMapeo.

## Prerequisites

- A Google account with access to Google Sheets
- The template spreadsheet:
  [https://docs.google.com/spreadsheets/d/1bvtbSijac5SPz-pBbeLBhKby6eDefwweLEBmjAOZnlk/edit?usp=drivesdk](https://docs.google.com/spreadsheets/d/1bvtbSijac5SPz-pBbeLBhKby6eDefwweLEBmjAOZnlk/edit?usp=drivesdk)

## Sheet map (quick reference)

| Sheet | Purpose |
| --- | --- |
| Categories | Category names, icons, fields, applies, colors |
| Details | Field definitions and options |
| Translation sheets | Optional translations for names, labels, helper text, options |
| Metadata | Config metadata and versioning |

For exact column rules, see the sheet references:
[Categories](reference/categories-sheet.md), [Details](reference/details-sheet.md),
[Translation sheets](reference/translation-sheets.md), [Metadata](reference/metadata-sheet.md).

## Step 1: Set your primary language

1. Open the **Categories** sheet.
2. Click cell **A1**.
3. Enter your primary language name (English or native name).

See [Translation sheets](reference/translation-sheets.md) for accepted language formats.

## Step 2: Add categories

1. In **Categories**, starting at row 2, enter category names in column A.

For examples and edge cases, see [Categories sheet tasks](how-to/categories-sheet-tasks.md#add-categories).

## Step 3: Add fields

1. In **Details**, starting at row 2, add field names.
2. Set type and options as needed.

For field types and options rules, see [Details sheet tasks](how-to/details-sheet-tasks.md#add-fields) and
[Details sheet](reference/details-sheet.md).

## Step 4: Link fields to categories

1. In **Categories** column C (**Fields**), list field names for each category.
2. Field names must match the **Details** sheet exactly.

See [Categories sheet tasks](how-to/categories-sheet-tasks.md#link-fields-to-categories).

## Step 5: Set Applies values (required)

1. In **Categories** column D (**Applies**), set where each category is used.
2. Ensure at least one category includes `track`.

See [Categories sheet](reference/categories-sheet.md) for valid values.
See [Categories sheet tasks](how-to/categories-sheet-tasks.md#set-applies-values-required) for steps.

## Step 6: Add icons

Paste inline SVG from <https://icons.earthdefenderstoolkit.com> into **Icon** (column B).

See [Categories sheet tasks](how-to/categories-sheet-tasks.md#add-icons) for all supported formats.

## Step 7: Set category colors

Set the background color of each category name cell in column A.

See [Categories sheet tasks](how-to/categories-sheet-tasks.md#set-category-colors).

## Step 8: Add translations (optional)

Use **CoMapeo Tools → Manage Languages & Translate** to add translation columns.

See [Translations](how-to/translations.md#add-languages) and
[Translation sheets](reference/translation-sheets.md).

## Step 9: Run linting

Run **CoMapeo Tools → Lint Sheets** and fix highlighted cells.

See [Export and share](how-to/export-and-share.md#run-linting) and [Linting rules](reference/linting-rules.md).

## Step 10: Generate the config

Open **CoMapeo Tools → Generate CoMapeo Category** and wait for completion.

See [Export and share](how-to/export-and-share.md#generate-a-config).

## Step 11: Share and import in CoMapeo

Share the `.comapeocat` from Drive and import it in CoMapeo.

See [Export and share](how-to/export-and-share.md#share-and-load-a-config).
If you are updating an existing config, see [Import and maintenance](how-to/import-and-maintenance.md#import-and-edit-a-config).

## What to do next

- If you need task-specific help, use the [How-to guides](index.md#how-to-guides).
- If you need exact rules and formats, use [Reference](index.md#reference).
- If something fails, see [Troubleshooting](troubleshooting/index.md).
