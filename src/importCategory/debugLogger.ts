/**
 * Debug logging system that writes to a spreadsheet sheet for real-time visibility
 * This solves the problem where Logger.log() and console.log() don't appear in async operations
 */

let debugSheet: GoogleAppsScript.Spreadsheet.Sheet | null = null;
let logBuffer: string[] = [];
let logRowIndex = 2; // Start after header
const BUFFER_SIZE = 20; // Flush every 20 messages
const MAX_LOG_ROWS = 5000; // Maximum rows before clearing old logs

/**
 * Initialize or get the Debug Logs sheet
 */
function getDebugSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  if (debugSheet) {
    return debugSheet;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("Debug Logs");

  if (!sheet) {
    // Create new debug sheet
    sheet = spreadsheet.insertSheet("Debug Logs");
    sheet.setTabColor("#ff0000"); // Red tab for visibility

    // Set up headers
    sheet.getRange(1, 1, 1, 3).setValues([["Timestamp", "Type", "Message"]]);
    sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
    sheet.setFrozenRows(1);

    // Format columns
    sheet.setColumnWidth(1, 150); // Timestamp
    sheet.setColumnWidth(2, 80);  // Type
    sheet.setColumnWidth(3, 800); // Message

    logRowIndex = 2;
  } else {
    // Find the last row with data
    const lastRow = sheet.getLastRow();
    logRowIndex = lastRow > 0 ? lastRow + 1 : 2;

    // Clear old logs if too many rows
    if (logRowIndex > MAX_LOG_ROWS) {
      sheet.getRange(2, 1, sheet.getMaxRows() - 1, 3).clearContent();
      logRowIndex = 2;
    }
  }

  debugSheet = sheet;
  return sheet;
}

/**
 * Flush buffered logs to the sheet
 */
function flushDebugLogs() {
  if (logBuffer.length === 0) {
    return;
  }

  try {
    const sheet = getDebugSheet();
    const timestamp = new Date().toLocaleTimeString();

    // Convert buffer to rows
    const rows = logBuffer.map(msg => {
      // Parse type from message prefix
      let type = "INFO";
      let cleanMsg = msg;

      if (msg.startsWith("❌") || msg.includes("ERROR")) {
        type = "ERROR";
      } else if (msg.startsWith("⚠️") || msg.includes("WARNING")) {
        type = "WARN";
      } else if (msg.startsWith("✓")) {
        type = "SUCCESS";
      } else if (msg.startsWith("===")) {
        type = "SECTION";
      }

      return [timestamp, type, cleanMsg];
    });

    // Write all buffered logs at once
    if (logRowIndex + rows.length > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), rows.length);
    }

    sheet.getRange(logRowIndex, 1, rows.length, 3).setValues(rows);

    // Apply color coding
    rows.forEach((row, i) => {
      const rowNum = logRowIndex + i;
      const type = row[1];

      if (type === "ERROR") {
        sheet.getRange(rowNum, 2).setBackground("#ffcccc");
      } else if (type === "WARN") {
        sheet.getRange(rowNum, 2).setBackground("#fff4cc");
      } else if (type === "SUCCESS") {
        sheet.getRange(rowNum, 2).setBackground("#ccffcc");
      } else if (type === "SECTION") {
        sheet.getRange(rowNum, 1, 1, 3).setFontWeight("bold").setBackground("#e6f2ff");
      }
    });

    logRowIndex += rows.length;
    logBuffer = [];
  } catch (error) {
    // Fallback to console.log if sheet logging fails
    console.error("Failed to write to debug sheet:", error);
    logBuffer.forEach(msg => console.log(msg));
    logBuffer = [];
  }
}

/**
 * Log a debug message to the Debug Logs sheet
 * @param message - The message to log
 * @param forceFlush - Force immediate flush to sheet (default: false)
 */
function debugLog(message: string, forceFlush = false) {
  // Always log to console as backup
  console.log(message);

  // Add to buffer
  logBuffer.push(message);

  // Flush if buffer is full or force flush requested
  if (forceFlush || logBuffer.length >= BUFFER_SIZE) {
    flushDebugLogs();
  }
}

/**
 * Log an error message
 */
function debugError(message: string, error?: any) {
  const errorMsg = error ? `${message}: ${error}` : message;
  debugLog(`❌ ERROR: ${errorMsg}`, true); // Force flush errors

  if (error && error.stack) {
    debugLog(`   Stack: ${error.stack}`);
  }
}

/**
 * Log a warning message
 */
function debugWarn(message: string) {
  debugLog(`⚠️  WARNING: ${message}`);
}

/**
 * Log a success message
 */
function debugSuccess(message: string) {
  debugLog(`✓ ${message}`);
}

/**
 * Log a section header
 */
function debugSection(sectionName: string) {
  debugLog(`\n=== ${sectionName} ===`, true);
}

/**
 * Clear the debug logs sheet
 */
function clearDebugLogs() {
  try {
    const sheet = getDebugSheet();
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, 3).clearContent();
    logRowIndex = 2;
    logBuffer = [];
    debugLog("Debug logs cleared");
  } catch (error) {
    console.error("Failed to clear debug logs:", error);
  }
}

/**
 * Initialize debug logging for an import session
 */
function initDebugLogging() {
  debugSheet = null; // Reset sheet reference
  logBuffer = [];
  logRowIndex = 2;

  // Clear old logs and start fresh
  clearDebugLogs();
  debugSection("IMPORT SESSION STARTED");
  debugLog(`Session started at: ${new Date().toLocaleString()}`);
  flushDebugLogs();
}

/**
 * Finalize debug logging for an import session
 */
function finalizeDebugLogging() {
  debugSection("IMPORT SESSION COMPLETED");
  debugLog(`Session completed at: ${new Date().toLocaleString()}`);
  flushDebugLogs();
}

/**
 * Factory that creates a module-scoped safe debug logger.
 *
 * Each consumer file should call this once at module scope:
 * ```
 * const safeDebugLog = createSafeDebugLogger("PngIcons");
 * ```
 *
 * The returned function tries `debugLog()` first (from this file),
 * falling back to `console.log()` if unavailable.
 *
 * @param moduleName - Prefix string injected into every log message
 */
function createSafeDebugLogger(moduleName: string) {
  return function safeDebugLog(message: string, forceFlush = false) {
    try {
      if (typeof debugLog === "function") {
        debugLog(`[${moduleName}] ${message}`, forceFlush);
        return;
      }
    } catch (_e) {
      // Fall through to console
    }
    console.log(`[${moduleName}] ${message}`);
  };
}
