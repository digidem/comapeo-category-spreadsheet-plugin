/**
 * Pre-flight validation module
 * Performs validation checks before starting the config generation process
 * to catch potential issues early and provide better user feedback.
 */

interface ValidationResult {
  passed: boolean;
  message?: string;
  details?: string;
}

interface PreflightCheckResults {
  allPassed: boolean;
  checks: {
    driveQuota: ValidationResult;
    apiHealth: ValidationResult;
    networkConnectivity: ValidationResult;
  };
}

/**
 * Check Drive storage quota availability
 * @returns Validation result with quota information
 */
const PREFLIGHT_CACHE_KEY = "preflight_checks";
const PREFLIGHT_CACHE_TTL = 300; // 5 minutes
const PREFLIGHT_FORCE_KEY = "force_refresh";

function checkDriveQuota(): ValidationResult {
  const log = getScopedLogger("PreflightValidation");
  try {
    log.info("[PREFLIGHT] Checking Drive quota...");
    // Note: DriveApp doesn't provide direct quota API in Apps Script
    // We can only check if we can create/access folders
    const testFolderName = "_temp_quota_check_" + new Date().getTime();

    try {
      const testFolder = DriveApp.getRootFolder().createFolder(testFolderName);
      testFolder.setTrashed(true); // Clean up immediately
      log.info("[PREFLIGHT] ✅ Drive quota check passed");
      return {
        passed: true,
        message: "Drive storage is accessible",
      };
    } catch (error) {
      if (error.message.includes("quota") || error.message.includes("storage")) {
        log.error("[PREFLIGHT] ❌ Drive quota exceeded");
        return {
          passed: false,
          message: "Drive storage quota exceeded",
          details: "Please free up space in your Google Drive and try again.",
        };
      }
      throw error; // Re-throw if it's not a quota issue
    }
  } catch (error) {
    log.error("[PREFLIGHT] ⚠️  Drive quota check failed:", error.message);
    return {
      passed: true, // Don't fail the process, just warn
      message: "Could not verify Drive quota (will continue anyway)",
      details: error.message,
    };
  }
}

/**
 * Check API server health and accessibility
 * @returns Validation result with API status
 */
function checkApiHealth(): ValidationResult {
  const log = getScopedLogger("PreflightValidation");
  try {
    log.info("[PREFLIGHT] Checking API health...");
    const apiUrl = "http://137.184.153.36:3000/";

    // Try to connect to the API with a short timeout
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "get",
      muteHttpExceptions: true,
      validateHttpsCertificates: false, // For HTTP endpoint
      // Note: Apps Script has a 60-second default timeout for UrlFetchApp
    };

    try {
      const response = UrlFetchApp.fetch(apiUrl, options);
      const responseCode = response.getResponseCode();

      if (responseCode === 200 || responseCode === 404) {
        // 200 or 404 means server is responding
        log.info("[PREFLIGHT] ✅ API health check passed (status: " + responseCode + ")");
        return {
          passed: true,
          message: "API server is accessible",
        };
      } else {
        log.warn("[PREFLIGHT] ⚠️  API returned unexpected status:", responseCode);
        return {
          passed: true, // Don't fail, the actual API might still work
          message: "API server responded with status " + responseCode,
        };
      }
    } catch (error) {
      const errorMessage = error.message || String(error);

      if (errorMessage.includes("timeout") || errorMessage.includes("timed out")) {
        log.error("[PREFLIGHT] ❌ API health check timeout");
        return {
          passed: false,
          message: "API server is not responding (timeout)",
          details: "The API server may be down or experiencing issues. Please try again later.",
        };
      } else if (errorMessage.includes("DNS") || errorMessage.includes("resolve")) {
        log.error("[PREFLIGHT] ❌ Cannot reach API server");
        return {
          passed: false,
          message: "Cannot reach API server",
          details: "Network connectivity issue. Please check your internet connection.",
        };
      }

      // For other errors, continue anyway (might be a transient issue)
      log.warn("[PREFLIGHT] ⚠️  API health check warning:", errorMessage);
      return {
        passed: true,
        message: "API health check inconclusive (will continue anyway)",
        details: errorMessage,
      };
    }
  } catch (error) {
    log.error("[PREFLIGHT] ⚠️  API health check failed:", error.message);
    return {
      passed: true, // Don't fail the process
      message: "Could not verify API health (will continue anyway)",
      details: error.message,
    };
  }
}

/**
 * Check network connectivity
 * @returns Validation result with connectivity status
 */
function checkNetworkConnectivity(): ValidationResult {
  const log = getScopedLogger("PreflightValidation");
  try {
    log.info("[PREFLIGHT] Checking network connectivity...");

    // Try to fetch a reliable public endpoint
    const testUrl = "https://www.google.com";
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "get",
      muteHttpExceptions: true,
    };

    try {
      const response = UrlFetchApp.fetch(testUrl, options);
      const responseCode = response.getResponseCode();

      if (responseCode >= 200 && responseCode < 400) {
        log.info("[PREFLIGHT] ✅ Network connectivity check passed");
        return {
          passed: true,
          message: "Network connectivity is active",
        };
      } else {
        log.warn("[PREFLIGHT] ⚠️  Unexpected response from connectivity check:", responseCode);
        return {
          passed: true,
          message: "Network connectivity check inconclusive",
        };
      }
    } catch (error) {
      log.error("[PREFLIGHT] ❌ Network connectivity check failed:", error.message);
      return {
        passed: false,
        message: "Network connectivity issue detected",
        details: "Please check your internet connection and try again.",
      };
    }
  } catch (error) {
    log.error("[PREFLIGHT] ⚠️  Network check failed:", error.message);
    return {
      passed: true, // Don't fail the process
      message: "Could not verify network (will continue anyway)",
      details: error.message,
    };
  }
}

/**
 * Run all pre-flight validation checks
 * @returns Results of all validation checks
 */
function runPreflightChecks(): PreflightCheckResults {
  const log = getScopedLogger("PreflightValidation");
  const cache = CacheService.getScriptCache();
  const cachedForceFlag = cache.get(PREFLIGHT_FORCE_KEY);

  if (!cachedForceFlag) {
    const cachedResults = cache.get(PREFLIGHT_CACHE_KEY);
    if (cachedResults) {
      try {
        const parsed = JSON.parse(cachedResults) as PreflightCheckResults;
        log.info("[PREFLIGHT] Using cached results");
        return parsed;
      } catch (error) {
        log.warn("[PREFLIGHT] Failed to parse cached results, running fresh checks");
      }
    }
  } else {
    cache.remove(PREFLIGHT_FORCE_KEY);
  }

  log.info("[PREFLIGHT] ===== Starting Pre-flight Validation =====");
  const startTime = new Date().getTime();

  const checks = {
    driveQuota: checkDriveQuota(),
    apiHealth: checkApiHealth(),
    networkConnectivity: checkNetworkConnectivity(),
  };

  const allPassed = checks.driveQuota.passed &&
    checks.apiHealth.passed &&
    checks.networkConnectivity.passed;

  const totalTime = ((new Date().getTime() - startTime) / 1000).toFixed(1);
  log.info(`[PREFLIGHT] ===== Pre-flight Validation Complete (${totalTime}s) =====`);

  try {
    cache.put(
      PREFLIGHT_CACHE_KEY,
      JSON.stringify({ allPassed, checks }),
      PREFLIGHT_CACHE_TTL,
    );
  } catch (error) {
    log.warn("[PREFLIGHT] Failed to cache preflight results", error);
  }

  if (allPassed) {
    log.info("[PREFLIGHT] ✅ All checks passed");
  } else {
    log.warn("[PREFLIGHT] ⚠️  Some checks failed");
  }

  return {
    allPassed,
    checks,
  };
}

/**
 * Display pre-flight validation results to user if any checks failed
 * @param results - Pre-flight check results
 * @returns true if user wants to continue despite failures, false to abort
 */
function showPreflightResults(results: PreflightCheckResults): boolean {
  if (results.allPassed) {
    return true; // All passed, continue
  }

  const ui = SpreadsheetApp.getUi();
  let message = "Some pre-flight checks did not pass:\n\n";

  if (!results.checks.driveQuota.passed) {
    message += "❌ Drive Storage: " + results.checks.driveQuota.message + "\n";
    if (results.checks.driveQuota.details) {
      message += "   " + results.checks.driveQuota.details + "\n";
    }
    message += "\n";
  }

  if (!results.checks.apiHealth.passed) {
    message += "❌ API Server: " + results.checks.apiHealth.message + "\n";
    if (results.checks.apiHealth.details) {
      message += "   " + results.checks.apiHealth.details + "\n";
    }
    message += "\n";
  }

  if (!results.checks.networkConnectivity.passed) {
    message += "❌ Network: " + results.checks.networkConnectivity.message + "\n";
    if (results.checks.networkConnectivity.details) {
      message += "   " + results.checks.networkConnectivity.details + "\n";
    }
    message += "\n";
  }

  message += "Do you want to continue anyway? (Not recommended)";

  const response = ui.alert(
    "Pre-flight Validation Failed",
    message,
    ui.ButtonSet.YES_NO,
  );

  return response === ui.Button.YES;
}
