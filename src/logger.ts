/**
 * Centralized logging system for consistent logging across the codebase
 *
 * Provides structured logging with multiple log levels, context tracking,
 * and configurable output formatting.
 *
 * Usage:
 *   AppLogger.info("Processing started", { count: 10 });
 *   AppLogger.error("Failed to process", error, { userId: "123" });
 *   AppLogger.setLevel(LogLevel.DEBUG);
 */

/**
 * Log levels in order of severity
 */
enum LogLevel {
  DEBUG = 0,
  INFO = 1,
  WARN = 2,
  ERROR = 3,
  NONE = 4,
}

/**
 * Configuration for the AppLogger
 */
interface AppLoggerConfig {
  /** Minimum log level to output */
  level: LogLevel;
  /** Whether to include timestamps in logs */
  includeTimestamp: boolean;
  /** Whether to include file/function context */
  includeContext: boolean;
  /** Maximum string length before truncation */
  maxStringLength: number;
}

/**
 * Centralized logging system
 *
 * Provides consistent, structured logging across the application with
 * configurable log levels, formatting, and context tracking.
 *
 * Note: Renamed from "AppLogger" to "AppAppLogger" to avoid conflict with
 * Google Apps Script's built-in AppLogger object.
 */
class AppAppLogger {
  private static config: AppLoggerConfig = {
    level: LogLevel.INFO,
    includeTimestamp: true,
    includeContext: true,
    maxStringLength: 1000,
  };

  /**
   * Sets the minimum log level
   *
   * @param level - The minimum log level to output
   *
   * @example
   * AppLogger.setLevel(LogLevel.DEBUG);
   */
  static setLevel(level: LogLevel): void {
    AppAppLogger.config.level = level;
  }

  /**
   * Gets the current log level
   *
   * @returns The current log level
   */
  static getLevel(): LogLevel {
    return AppAppLogger.config.level;
  }

  /**
   * Configures the logger
   *
   * @param config - Partial configuration to merge with defaults
   *
   * @example
   * AppLogger.configure({ includeTimestamp: false, maxStringLength: 500 });
   */
  static configure(config: Partial<AppLoggerConfig>): void {
    AppAppLogger.config = { ...AppAppLogger.config, ...config };
  }

  /**
   * Logs a debug message
   *
   * @param message - The log message
   * @param args - Additional arguments to log
   *
   * @example
   * AppLogger.debug("Processing item", { id: 123, name: "test" });
   */
  static debug(message: string, ...args: any[]): void {
    AppAppLogger.log(LogLevel.DEBUG, message, ...args);
  }

  /**
   * Logs an info message
   *
   * @param message - The log message
   * @param args - Additional arguments to log
   *
   * @example
   * AppLogger.info("Config generated successfully", { fileCount: 10 });
   */
  static info(message: string, ...args: any[]): void {
    AppAppLogger.log(LogLevel.INFO, message, ...args);
  }

  /**
   * Logs a warning message
   *
   * @param message - The log message
   * @param args - Additional arguments to log
   *
   * @example
   * AppLogger.warn("Missing optional field", { field: "description" });
   */
  static warn(message: string, ...args: any[]): void {
    AppAppLogger.log(LogLevel.WARN, message, ...args);
  }

  /**
   * Logs an error message
   *
   * @param message - The log message
   * @param error - Optional error object
   * @param args - Additional arguments to log
   *
   * @example
   * AppLogger.error("Failed to process", error, { userId: "123" });
   */
  static error(message: string, error?: Error | any, ...args: any[]): void {
    const allArgs = error ? [error, ...args] : args;
    AppAppLogger.log(LogLevel.ERROR, message, ...allArgs);
  }

  /**
   * Internal logging implementation
   *
   * @param level - The log level
   * @param message - The log message
   * @param args - Additional arguments to log
   */
  private static log(level: LogLevel, message: string, ...args: any[]): void {
    // Skip if below configured level
    if (level < AppAppLogger.config.level) {
      return;
    }

    const levelName = LogLevel[level];
    const parts: string[] = [];

    // Add timestamp
    if (AppAppLogger.config.includeTimestamp) {
      parts.push(`[${AppAppLogger.getTimestamp()}]`);
    }

    // Add level
    parts.push(`[${levelName}]`);

    // Add message
    parts.push(message);

    // Format the log entry
    let logEntry = parts.join(" ");

    // Add additional arguments if present
    if (args.length > 0) {
      const formattedArgs = args
        .map((arg) => AppAppLogger.formatArgument(arg))
        .filter((arg) => arg !== null)
        .join(" ");

      if (formattedArgs) {
        logEntry += ` ${formattedArgs}`;
      }
    }

    // Output to console
    AppAppLogger.output(level, logEntry);
  }

  /**
   * Formats a single argument for logging
   *
   * @param arg - The argument to format
   * @returns Formatted string or null if empty
   */
  private static formatArgument(arg: any): string | null {
    if (arg === null || arg === undefined) {
      return null;
    }

    try {
      // Handle Error objects specially
      if (arg instanceof Error) {
        return `Error: ${arg.message}${arg.stack ? `\nStack: ${arg.stack}` : ""}`;
      }

      // Handle objects and arrays
      if (typeof arg === "object") {
        const json = JSON.stringify(arg, null, 2);
        return AppAppLogger.truncate(json);
      }

      // Handle primitives
      const str = String(arg);
      return AppAppLogger.truncate(str);
    } catch (error) {
      return `[Unserializable: ${typeof arg}]`;
    }
  }

  /**
   * Truncates a string if it exceeds max length
   *
   * @param str - The string to truncate
   * @returns Truncated string
   */
  private static truncate(str: string): string {
    if (str.length <= AppAppLogger.config.maxStringLength) {
      return str;
    }

    const truncated = str.substring(0, AppAppLogger.config.maxStringLength);
    return `${truncated}... [truncated ${str.length - AppAppLogger.config.maxStringLength} chars]`;
  }

  /**
   * Gets formatted timestamp
   *
   * @returns ISO timestamp string
   */
  private static getTimestamp(): string {
    return new Date().toISOString();
  }

  /**
   * Outputs log to console with appropriate method
   *
   * @param level - The log level
   * @param message - The formatted log message
   */
  private static output(level: LogLevel, message: string): void {
    switch (level) {
      case LogLevel.DEBUG:
        console.log(message);
        break;
      case LogLevel.INFO:
        console.log(message);
        break;
      case LogLevel.WARN:
        console.warn(message);
        break;
      case LogLevel.ERROR:
        console.error(message);
        break;
    }
  }

  /**
   * Creates a scoped logger for a specific module/context
   *
   * @param context - The context name (e.g., "IconProcessor", "DriveService")
   * @returns Object with logging methods that include context
   *
   * @example
   * const log = AppLogger.scope("IconProcessor");
   * log.info("Processing icon", { name: "tree" });
   * // Output: [2024-01-01T00:00:00.000Z] [INFO] [IconProcessor] Processing icon {"name":"tree"}
   */
  static scope(context: string): {
    debug: (message: string, ...args: any[]) => void;
    info: (message: string, ...args: any[]) => void;
    warn: (message: string, ...args: any[]) => void;
    error: (message: string, error?: Error | any, ...args: any[]) => void;
  } {
    return {
      debug: (message: string, ...args: any[]) =>
        AppAppLogger.debug(`[${context}] ${message}`, ...args),
      info: (message: string, ...args: any[]) =>
        AppAppLogger.info(`[${context}] ${message}`, ...args),
      warn: (message: string, ...args: any[]) =>
        AppAppLogger.warn(`[${context}] ${message}`, ...args),
      error: (message: string, error?: Error | any, ...args: any[]) =>
        AppAppLogger.error(`[${context}] ${message}`, error, ...args),
    };
  }

  /**
   * Logs performance timing information
   *
   * @param operation - The operation being timed
   * @param startTime - Start time in milliseconds
   *
   * @example
   * const start = Date.now();
   * // ... operation ...
   * AppLogger.timing("processIcons", start);
   * // Output: [INFO] Operation 'processIcons' took 1234ms
   */
  static timing(operation: string, startTime: number): void {
    const duration = Date.now() - startTime;
    AppAppLogger.info(`Operation '${operation}' took ${duration}ms`);
  }

  /**
   * Groups related log messages
   *
   * @param label - The group label
   * @param callback - Function to execute within the group
   *
   * @example
   * AppLogger.group("Processing Categories", () => {
   *   AppLogger.info("Loading data");
   *   AppLogger.info("Validating");
   *   AppLogger.info("Complete");
   * });
   */
  static group(label: string, callback: () => void): void {
    AppAppLogger.info(`╔══ ${label} ══`);
    try {
      callback();
    } finally {
      AppAppLogger.info(`╚══ End ${label} ══`);
    }
  }
}

/**
 * Export LogLevel enum for external use
 */
// Note: In Apps Script global scope, this is accessed as LogLevel
// No explicit export needed, but keep for TypeScript type checking
