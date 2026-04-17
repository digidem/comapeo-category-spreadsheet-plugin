/**
 * Functions for extracting and processing PNG icon files from the icons/ directory
 */

/**
 * Safe debug logger that falls back to console.log if debugLog is not available
 */
function safePngDebugLog(message: string) {
  // Try debug logger first
  try {
    if (typeof debugLog === "function") {
      debugLog(message);
      return;
    }
  } catch (e) {
    // Fall through to console
  }

  // Fall back to console
  console.log(message);
}

/**
 * Extracts PNG icons from the temp folder and copies them to permanent storage
 * @param tempFolder - The temporary folder containing extracted files
 * @param presets - Array of preset objects that reference icon names
 * @param onProgress - Optional progress callback function
 * @returns An array of icon objects with name, URL, and ID
 */
function extractPngIcons(
  tempFolder: GoogleAppsScript.Drive.Folder,
  presets: any[],
  onProgress?: (update: { percent: number; stage: string; detail?: string }) => void,
): { name: string; svg: string; id: string }[] {
  try {
    safePngDebugLog("Extracting PNG icons from temp folder");

    // Get permanent config folder for icon storage
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const parentFolder = DriveApp.getFileById(spreadsheet.getId())
      .getParents()
      .next();
    const configFolderName = slugify(spreadsheet.getName());

    safePngDebugLog(`Looking for config folder: ${configFolderName}`);

    // Find or create config folder
    let configFolder: GoogleAppsScript.Drive.Folder;
    const configFolders = parentFolder.getFoldersByName(configFolderName);
    if (configFolders.hasNext()) {
      configFolder = configFolders.next();
      safePngDebugLog(`Using existing config folder: ${configFolder.getName()}`);
    } else {
      configFolder = parentFolder.createFolder(configFolderName);
      safePngDebugLog(`Created new config folder: ${configFolder.getName()}`);
    }

    // Find or create permanent icons folder
    let permanentIconsFolder: GoogleAppsScript.Drive.Folder;
    const iconsFolders = configFolder.getFoldersByName("icons");
    if (iconsFolders.hasNext()) {
      permanentIconsFolder = iconsFolders.next();
      safePngDebugLog(
        `Using existing icons folder: ${permanentIconsFolder.getName()}`,
      );
    } else {
      permanentIconsFolder = configFolder.createFolder("icons");
      safePngDebugLog(`Created new icons folder: ${permanentIconsFolder.getName()}`);
    }

    // Look for icons/ subdirectory in temp folder
    const tempIconsFolders = tempFolder.getFoldersByName("icons");
    if (!tempIconsFolders.hasNext()) {
      safePngDebugLog("No icons/ directory found in temp folder");
      return [];
    }

    const tempIconsFolder = tempIconsFolders.next();
    safePngDebugLog(`Found icons folder in temp: ${tempIconsFolder.getName()}`);

    // Extract icon names from presets
    const iconNames = new Set<string>();
    presets.forEach((preset) => {
      if (preset.icon) {
        iconNames.add(preset.icon);
      }
    });

    safePngDebugLog(`Looking for ${iconNames.size} icon(s): ${Array.from(iconNames).join(", ")}`);

    // Build a file index for O(1) lookup instead of O(n) getFilesByName() calls
    safePngDebugLog("Indexing icon files for fast lookup...");
    if (onProgress) {
      onProgress({
        percent: 40,
        stage: "Extracting icons",
        detail: `Indexing ${tempIconsFolder.getName()} files for fast lookup...`,
      });
    }

    const availableFiles = new Map<string, GoogleAppsScript.Drive.File>();
    const fileIterator = tempIconsFolder.getFiles();
    let totalFiles = 0;

    while (fileIterator.hasNext()) {
      const file = fileIterator.next();
      availableFiles.set(file.getName(), file);
      totalFiles++;
    }

    safePngDebugLog(`Indexed ${totalFiles} files. Starting icon extraction...`);
    if (onProgress) {
      onProgress({
        percent: 45,
        stage: "Extracting icons",
        detail: `Indexed ${totalFiles} files. Starting extraction...`,
      });
    }

    // Size and resolution priority order (preferred sizes first)
    const sizePriority = ["medium", "small", "large"];
    const resolutionPriority = ["1x", "2x", "3x"];

    const iconObjects: { name: string; svg: string; id: string }[] = [];

    // Process each icon with progress tracking
    let processed = 0;
    const total = iconNames.size;
    const failedIcons: { name: string; error: string }[] = [];

    safePngDebugLog(`\n=== PNG ICON EXTRACTION (${total} icons) ===`);

    iconNames.forEach((iconName) => {
      processed++;

      // Calculate progress percentage (45% to 70% range for icon extraction)
      const iconPercent = Math.round(45 + ((processed / total) * 25));

      // Log progress every 5 icons to show we're making progress
      if (processed % 5 === 0 || processed === total) {
        safePngDebugLog(
          `Extracting icons: ${processed}/${total} (${Math.round((processed / total) * 100)}%)`,
        );
        if (onProgress) {
          onProgress({
            percent: iconPercent,
            stage: "Extracting icons",
            detail: `${processed}/${total} icons extracted`,
          });
        }
      }

      let foundFile: GoogleAppsScript.Drive.File | null = null;
      let foundPattern = "";

      // Try each size/resolution using fast Map lookup
      for (const size of sizePriority) {
        for (const resolution of resolutionPriority) {
          const fileName = `${iconName}-${size}@${resolution}.png`;

          if (availableFiles.has(fileName)) {
            foundFile = availableFiles.get(fileName)!;
            foundPattern = fileName;
            break; // Found, stop searching
          }
        }

        if (foundFile) {
          break; // Found, stop searching sizes
        }
      }

      if (!foundFile) {
        console.warn(`⚠️  No PNG file found for icon: ${iconName}`);
        failedIcons.push({ name: iconName, error: "No matching PNG file in icons/ directory" });
        return; // Skip this icon
      }

      safePngDebugLog(`  ✓ Found PNG for "${iconName}": ${foundPattern}`);

      // Copy to permanent folder
      try {
        const fileName = `${iconName}.png`;
        const existingFiles = permanentIconsFolder.getFilesByName(fileName);
        let permanentFile: GoogleAppsScript.Drive.File;

        if (existingFiles.hasNext()) {
          const existingFile = existingFiles.next();
          const oldSize = existingFile.getSize();
          // Create the replacement BEFORE trashing the old file to avoid
          // data loss if createFile() or getBlob() fails.
          const blob = foundFile.getBlob().setName(fileName);
          permanentFile = permanentIconsFolder.createFile(blob);
          existingFile.setTrashed(true);
          const newSize = permanentFile.getSize();
          safePngDebugLog(`  ↻ Replaced existing "${fileName}": ${oldSize} → ${newSize} bytes`);
        } else {
          const blob = foundFile.getBlob().setName(fileName);
          permanentFile = permanentIconsFolder.createFile(blob);
          const fileSize = permanentFile.getSize();
          safePngDebugLog(`  ✓ Created "${fileName}": ${fileSize} bytes`);
        }

        // Verify file was created/updated successfully
        if (permanentFile.getSize() === 0) {
          console.error(`  ❌ ERROR: "${fileName}" is 0 bytes - file is empty!`);
          failedIcons.push({ name: iconName, error: "Created PNG file is empty (0 bytes)" });
          return;
        }

        const iconUrl = permanentFile.getUrl();
        safePngDebugLog(`  ✓ URL: ${iconUrl.substring(0, 60)}...`);

        iconObjects.push({
          name: iconName,
          svg: iconUrl, // Note: Property named 'svg' but contains PNG URL
          id: iconName,
        });
      } catch (error) {
        console.error(`  ❌ ERROR copying icon ${iconName}:`, error);
        console.error(`     Stack: ${error.stack || "No stack trace"}`);
        failedIcons.push({ name: iconName, error: String(error) });
      }
    });

    // Report extraction results
    safePngDebugLog(`\n=== PNG EXTRACTION RESULTS ===`);
    safePngDebugLog(`✓ Successfully extracted: ${iconObjects.length}/${total} icons`);
    if (failedIcons.length > 0) {
      safePngDebugLog(`❌ Failed to extract: ${failedIcons.length}/${total} icons`);
      safePngDebugLog(`\n=== FAILED PNG ICONS ===`);
      failedIcons.forEach((failed, index) => {
        safePngDebugLog(`  ${index + 1}. "${failed.name}": ${failed.error}`);
      });
      safePngDebugLog(`=== END FAILED PNG ICONS ===`);
    }
    safePngDebugLog(`=== END PNG EXTRACTION ===\n`);

    safePngDebugLog(
      `Successfully extracted ${iconObjects.length} PNG icon(s) to permanent folder`,
    );
    return iconObjects;
  } catch (error) {
    console.error("Error extracting PNG icons:", error);
    return [];
  }
}

/**
 * Helper function to check if icons/ directory exists in temp folder
 * @param tempFolder - The temporary folder to check
 * @returns True if icons/ directory exists, false otherwise
 */
function hasPngIconsDirectory(
  tempFolder: GoogleAppsScript.Drive.Folder,
): boolean {
  try {
    const iconsFolders = tempFolder.getFoldersByName("icons");
    return iconsFolders.hasNext();
  } catch (error) {
    console.error("Error checking for PNG icons directory:", error);
    return false;
  }
}
