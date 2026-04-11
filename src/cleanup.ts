function removeTranslationAndMetadataSheets(): void {
  const log = getScopedLogger('Cleanup');
  const sheetsToRemove = [...sheets(true), "Metadata", "Debug Logs"];

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  sheetsToRemove.forEach((sheetName) => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      spreadsheet.deleteSheet(sheet);
      log.info(`Removed sheet: ${sheetName}`);
    } else {
      log.info(`Sheet not found: ${sheetName}`);
    }
  });

  log.info("Finished removing translation sheets");
}

function deleteIcons(): void {
  const log = getScopedLogger('Cleanup');
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categoriesSheet = spreadsheet.getSheetByName("Categories");

  if (!categoriesSheet) {
    log.info("Categories sheet not found");
    return;
  }

  const lastRow = categoriesSheet.getLastRow();
  if (lastRow <= 1) {
    log.info("Categories sheet is empty or contains only header");
    return;
  }

  // Delete column B content, excluding the header
  const range = categoriesSheet.getRange(2, 2, lastRow - 1, 1);
  range.clearContent();

  log.info(
    "Deleted content from column B in Categories sheet, excluding the header and non-empty cells",
  );
}

function cleanup(): void {
  removeTranslationAndMetadataSheets();
  deleteIcons();
}

/**
 * Automatically cleans up old temporary folders created by the app
 *
 * @param olderThanHours - Delete folders older than this many hours (default: 24)
 * @param dryRun - If true, only logs what would be deleted without actually deleting
 * @returns Summary of cleanup operation
 */
function cleanupOldTempFolders(
  olderThanHours: number = 24,
  dryRun: boolean = false,
): {
  foldersFound: number;
  foldersDeleted: number;
  foldersSkipped: number;
  errors: string[];
} {
  const log = getScopedLogger('Cleanup');

  log.info(`Starting automatic cleanup of temp folders older than ${olderThanHours} hours`, {
    dryRun,
  });

  const results = {
    foldersFound: 0,
    foldersDeleted: 0,
    foldersSkipped: 0,
    errors: [] as string[],
  };

  // Calculate cutoff time
  const cutoffTime = new Date().getTime() - olderThanHours * 60 * 60 * 1000;
  log.debug("Cutoff time:", new Date(cutoffTime).toISOString());

  // Patterns for temp folders created by this app
  const tempFolderPatterns = [
    /^Mapeo_Settings_Import_\d+$/,
    /^Config_Import_\d+$/,
    /^temp_import_\d+$/,
    /^Icons_Test_\d+$/,
    /^Format_Detection_Test_\d+$/,
    /^Translation_Extraction_Test_\d+$/,
    /^Category_Import_Test_\d+$/,
    /^Field_Extraction_Test_\d+$/,
    /^Details_Icons_Test_\d+$/,
  ];

  try {
    // Get all folders in user's Drive root
    const folders = DriveApp.getFolders();

    while (folders.hasNext()) {
      const folder = folders.next();
      const folderName = folder.getName();

      // Check if folder name matches any temp pattern
      const isTemporaryFolder = tempFolderPatterns.some((pattern) =>
        pattern.test(folderName),
      );

      if (!isTemporaryFolder) {
        continue; // Skip non-temp folders
      }

      results.foldersFound++;

      try {
        // Extract timestamp from folder name
        const timestampMatch = folderName.match(/_(\d+)$/);
        let folderAge: number;

        if (timestampMatch) {
          // Use timestamp from folder name
          const folderTimestamp = parseInt(timestampMatch[1], 10);
          folderAge = folderTimestamp;
          log.debug(`Folder ${folderName} timestamp from name: ${new Date(folderTimestamp).toISOString()}`);
        } else {
          // Fallback to modification date
          folderAge = folder.getLastUpdated().getTime();
          log.debug(`Folder ${folderName} using modification date: ${folder.getLastUpdated().toISOString()}`);
        }

        // Check if folder is old enough to delete
        if (folderAge < cutoffTime) {
          const ageHours = Math.round((new Date().getTime() - folderAge) / (60 * 60 * 1000));
          log.info(`Found old temp folder: ${folderName} (${ageHours} hours old)`);

          if (!dryRun) {
            // Delete the folder using our robust cleanup function
            const cleanupResult = cleanupTempResources(folder, {
              maxRetries: 3,
              forceDelete: true,
              deleteChildrenFirst: true,
            });

            if (cleanupResult.success) {
              results.foldersDeleted++;
              log.info(`Successfully deleted: ${folderName}`);
            } else {
              results.foldersSkipped++;
              const errorMsg = `Failed to delete ${folderName}: ${cleanupResult.errors?.join(", ")}`;
              log.warn(errorMsg);
              results.errors.push(errorMsg);
            }
          } else {
            log.info(`[DRY RUN] Would delete: ${folderName}`);
            results.foldersDeleted++; // Count as "would be deleted"
          }
        } else {
          const ageHours = Math.round((new Date().getTime() - folderAge) / (60 * 60 * 1000));
          log.debug(`Skipping recent temp folder: ${folderName} (${ageHours} hours old)`);
          results.foldersSkipped++;
        }
      } catch (folderError) {
        const errorMsg = `Error processing folder ${folderName}: ${folderError}`;
        log.error(errorMsg, folderError);
        results.errors.push(errorMsg);
        results.foldersSkipped++;
      }
    }

    log.info("Cleanup completed", {
      found: results.foldersFound,
      deleted: results.foldersDeleted,
      skipped: results.foldersSkipped,
      errors: results.errors.length,
    });

    return results;
  } catch (error) {
    const errorMsg = `Fatal error during temp folder cleanup: ${error}`;
    log.error(errorMsg, error);
    results.errors.push(errorMsg);
    return results;
  }
}

/**
 * Cleans up temporary resources with robust error handling and retry logic
 * @param tempFolder - Temporary folder to clean up
 * @param options - Cleanup options
 * @param depth - Current recursion depth (for internal use)
 * @returns Success status and any error messages
 */
function cleanupTempResources(
  tempFolder: GoogleAppsScript.Drive.Folder,
  options?: {
    maxRetries?: number;
    retryDelayMs?: number;
    forceDelete?: boolean;
    deleteChildrenFirst?: boolean;
  },
  depth: number = 0,
): { success: boolean; errors?: string[] } {
  // Maximum recursion depth to prevent infinite loops
  const MAX_FOLDER_DEPTH = 50;
  const log = getScopedLogger('Cleanup');

  if (depth > MAX_FOLDER_DEPTH) {
    const errorMsg = `Maximum folder depth (${MAX_FOLDER_DEPTH}) exceeded during cleanup. Possible circular reference detected.`;
    log.error(errorMsg);
    return { success: false, errors: [errorMsg] };
  }

  // Default options
  const maxRetries = options?.maxRetries || 3;
  const retryDelayMs = options?.retryDelayMs || 1000;
  const forceDelete = options?.forceDelete || false;
  const deleteChildrenFirst = options?.deleteChildrenFirst || true;

  const errors: string[] = [];

  try {
    log.info(`Cleaning up temporary folder: ${tempFolder.getName()}`);

    // Check if folder exists and is accessible
    try {
      tempFolder.getName(); // This will throw if folder doesn't exist or is inaccessible
    } catch (error) {
      log.warn("Folder already deleted or inaccessible:", error);
      return { success: true }; // Consider this a success since the folder is gone
    }

    // Delete children first if requested
    if (deleteChildrenFirst) {
      try {
        log.info("Deleting child files first...");
        const files = tempFolder.getFiles();
        while (files.hasNext()) {
          const file = files.next();
          try {
            file.setTrashed(true);
            log.info(`Deleted file: ${file.getName()}`);
          } catch (fileError) {
            const errorMsg = `Failed to delete file ${file.getName()}: ${fileError}`;
            log.warn(errorMsg);
            errors.push(errorMsg);

            // If force delete is enabled, try to delete the file with retries
            if (forceDelete) {
              let retryCount = 0;
              while (retryCount < maxRetries) {
                try {
                  Utilities.sleep(retryDelayMs);
                  file.setTrashed(true);
                  log.info(
                    `Successfully deleted file ${file.getName()} on retry ${retryCount + 1}`,
                  );
                  break;
                } catch (retryError) {
                  retryCount++;
                  log.warn(
                    `Retry ${retryCount} failed for file ${file.getName()}: ${retryError}`,
                  );
                }
              }
            }
          }
        }

        // Delete subfolders
        log.info("Deleting subfolders...");
        const folders = tempFolder.getFolders();
        while (folders.hasNext()) {
          const subfolder = folders.next();
          try {
            // Recursively clean up subfolders with depth tracking
            const subfolderResult = cleanupTempResources(subfolder, options, depth + 1);
            if (!subfolderResult.success && subfolderResult.errors) {
              errors.push(...subfolderResult.errors);
            }
          } catch (folderError) {
            const errorMsg = `Failed to delete subfolder ${subfolder.getName()}: ${folderError}`;
            log.warn(errorMsg);
            errors.push(errorMsg);
          }
        }
      } catch (childrenError) {
        const errorMsg = `Error accessing folder children: ${childrenError}`;
        log.warn(errorMsg);
        errors.push(errorMsg);
      }
    }

    // Delete the folder itself with retries
    let retryCount = 0;
    let folderDeleted = false;

    while (retryCount <= maxRetries && !folderDeleted) {
      try {
        if (retryCount > 0) {
          log.info(`Retrying folder deletion (attempt ${retryCount})...`);
          Utilities.sleep(retryDelayMs * retryCount); // Exponential backoff
        }

        tempFolder.setTrashed(true);
        folderDeleted = true;
        log.info(`Successfully deleted folder: ${tempFolder.getName()}`);
      } catch (error) {
        retryCount++;
        const errorMsg = `Failed to delete folder on attempt ${retryCount}: ${error}`;
        log.warn(errorMsg);

        if (retryCount > maxRetries) {
          errors.push(errorMsg);
        }
      }
    }

    // Return success status
    return {
      success: folderDeleted || errors.length === 0,
      errors: errors.length > 0 ? errors : undefined,
    };
  } catch (error) {
    const errorMsg = `Unexpected error during cleanup: ${error}`;
    log.error(errorMsg);
    errors.push(errorMsg);

    return {
      success: false,
      errors,
    };
  }
}
