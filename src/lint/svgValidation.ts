/// <reference path="shared.ts" />

/**
 * Phase 6 Task 1: Checks inline SVG sizes in Categories column B and Icons column B.
 * - Warning if SVG > 300 KB (307200 bytes)
 * - Error if SVG > 2 MB (2097152 bytes)
 */
function checkInlineSvgSizes(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logger = getScopedLogger("LintInlineSvgSizes");

  const SVG_WARN_BYTES = 307200; // 300 KB
  const SVG_ERROR_BYTES = 2097152; // 2 MB

  const sheetsToCheck = [
    { name: "Categories", col: 2 },
    { name: "Icons", col: 2 },
  ];

  for (const { name, col } of sheetsToCheck) {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;

    // Read both column A (name/ID) and column B (icon source) so we can skip
    // rows that the builder would also skip (no name for Categories, no ID for Icons).
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

    for (let i = 0; i < values.length; i++) {
      const rowKey = String(values[i][0] || "").trim();
      // Mirror buildIconsFromSheet eligibility: skip rows without a name (Categories)
      // or without an ID (Icons), just as the builder does.
      if (!rowKey) continue;

      const value = String(values[i][1] || "").trim();
      // Normalize icon source to extract inline SVG content for size checking.
      // parseIconSource() handles inline <svg>, data:image/svg+xml URIs, and
      // Drive URLs. Only measure sources that resolve to inline svgData.
      let svgContent: string | null = null;
      if (value.startsWith("<svg")) {
        svgContent = value;
      } else if (value.toLowerCase().startsWith("data:image/svg+xml")) {
        // Decode data URI to get the actual SVG content for size measurement
        svgContent = decodeDataSvgForLint(value);
      } else if (value.startsWith("https://drive.google.com/file/d/")) {
        const fileId = extractDriveFileId(value);
        if (fileId) {
          svgContent = loadDriveSvgForLint(fileId);
        }
      }
      // Remote non-Drive URLs are not measured here because they require network
      // access, but Drive SVGs are inlined by the builder and must match export-time
      // size validation.
      if (!svgContent) continue;

      const sizeBytes = Utilities.newBlob(svgContent).getBytes().length;
      const sizeKB = Math.round(sizeBytes / 1024);

      if (sizeBytes > SVG_ERROR_BYTES) {
        const row = i + 2;
        appendLintNote(
          sheet.getRange(row, col),
          `Inline SVG is ${sizeKB}KB (limit: 300KB warning, 2MB error)`,
          "error",
        );
        logger.warn(
          `${name} row ${row}: inline SVG is ${sizeKB}KB (exceeds 2MB error limit)`,
        );
      } else if (sizeBytes > SVG_WARN_BYTES) {
        const row = i + 2;
        appendLintNote(
          sheet.getRange(row, col),
          `Inline SVG is ${sizeKB}KB (limit: 300KB warning, 2MB error)`,
          "warning",
        );
        logger.warn(
          `${name} row ${row}: inline SVG is ${sizeKB}KB (exceeds 300KB warning limit)`,
        );
      }
    }
  }
}
