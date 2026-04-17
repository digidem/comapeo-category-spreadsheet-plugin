/**
 * Shared logging helpers to provide scoped loggers with a safe fallback.
 *
 * Apps Script bundles all TypeScript files into the same global scope, so this
 * helper exposes globally available functions without using import/export.
 */

/// <reference path="./types.ts" />

type ScopedLogger = {
  debug: (message: string, ...args: unknown[]) => void;
  info: (message: string, ...args: unknown[]) => void;
  warn: (message: string, ...args: unknown[]) => void;
  error: (message: string, error?: unknown, ...args: unknown[]) => void;
};

/**
 * Shape of the global AppLogger provided by the logging system when loaded.
 */
type AppLoggerGlobal = {
  scope: (context: string) => ScopedLogger;
  timing?: (operation: string, startTime: number) => void;
};

declare const AppLogger: AppLoggerGlobal | undefined;
const fallbackLoggerCache: Record<string, ScopedLogger> = {};

/**
 * Returns a scoped logger instance. Falls back to console logging when the
 * centralized AppLogger is not yet available (e.g., during Apps Script
 * initialization or unit tests).
 *
 * @param scope - Context name to prefix log messages with.
 * @returns Scoped logger with debug/info/warn/error methods.
 */
function getScopedLogger(scope: string): ScopedLogger {
  if (typeof AppLogger !== "undefined" && AppLogger && typeof AppLogger.scope === "function") {
    return AppLogger.scope(scope);
  }

  if (
    typeof AppAppLogger !== "undefined" &&
    AppAppLogger &&
    typeof AppAppLogger.scope === "function"
  ) {
    return AppAppLogger.scope(scope);
  }

  if (!fallbackLoggerCache[scope]) {
    fallbackLoggerCache[scope] = createFallbackLogger(scope);
  }

  return fallbackLoggerCache[scope];
}

/**
 * Creates a simple console-based scoped logger.
 *
 * @param scope - Context used to annotate log messages.
 */
function createFallbackLogger(scope: string): ScopedLogger {
  const prefix = scope ? `[${scope}]` : "";

  return {
    debug: (message: string, ...args: unknown[]) => {
      console.log(`[DEBUG]${prefix} ${message}`, ...args);
    },
    info: (message: string, ...args: unknown[]) => {
      console.log(`[INFO]${prefix} ${message}`, ...args);
    },
    warn: (message: string, ...args: unknown[]) => {
      console.warn(`[WARN]${prefix} ${message}`, ...args);
    },
    error: (message: string, error?: unknown, ...args: unknown[]) => {
      if (typeof error !== "undefined") {
        console.error(`[ERROR]${prefix} ${message}`, error, ...args);
      } else {
        console.error(`[ERROR]${prefix} ${message}`, ...args);
      }
    },
  };
}
